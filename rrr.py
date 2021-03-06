"""Rename Resize Rotate

Renames a series of files in a directory, resizing to letter size and
rotating any PDFs found within to portrait orientation.
* Expands all zip files and email attachments found within target directories.
* Converts all MS Word, MS Excel, MS PowerPoint, TIF, JPG, PNG, and TXT files
to PDFs before RRRing.
* Adds tab placeholders down to a specified directory depth for ease of layout
during later printing.
* Copies files that are unmodifiable to the root directory for case-by-case
processing by end user.

* Requires MS Office 2010 or newer installed to convert Office files.
* Requires Adobe Acrobat Pro to convert images *OR* see comments in process_image
for an (untested) ImageMagick implementation.

Usage:
    rrr [<sourcedir> <destdir> <tabdepth>]
    rrr (-h | --help)

Options:
    -h --help               Show this screen
"""
import os,io,zipfile,sys,comtypes.client,email,mimetypes,olefile,shutil,time,StringIO,re,string
import logging,Tkinter,tkFileDialog,tkMessageBox,copy,sys,ttk
from natsort import natsorted
from PyPDF2 import PdfFileReader,PdfFileWriter
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from operator import itemgetter

def main (rootdir,tabdepth,page_setup_settings,email_formatting_status,numbering_status,slipsheets_status,flatten_status,max_pdf_size,max_excel_size):
    global files_to_remove
    files_to_remove = []
    reload(sys)
    sys.setdefaultencoding('utf8')
    logging.basicConfig(filename='rrrlog.txt',level=logging.INFO)
    console = logging.StreamHandler()
    console.setLevel(logging.DEBUG)
    logging.getLogger('').addHandler(console)
    unzip(rootdir,email_formatting_status)
    add_directory_slipsheets(rootdir,tabdepth)
    convert_to_pdf(rootdir,page_setup_settings)
    app.progress_bar["value"] = 50
    app.progress_bar.update()
    rename_resize_rotate(rootdir,numbering_status,slipsheets_status,tabdepth,max_pdf_size,max_excel_size)
    if flatten_status==1:
        notify("Flattening directory structure...")
        flatten(rootdir)
    notify("FINISHED!",alt="")
    app.progress_bar["value"] = 100
    app.progress_bar.update()
    app.progress_text2["text"] = "Ready!"
    app.progress_text2.update()
def notify(text,alt=False,level='i'):
        if level=='i':
            logging.info(text)
        elif level=='w':
            logging.warning(text)
        else:
            logging.debug(text)
        app.progress_text3["text"] = alt or text
        app.progress_text3.update()
def flatten(rootdir):
    for subdir,dirs,files in os.walk(rootdir):
        for file in files:
            if os.path.exists(os.path.join(rootdir,file))==False:
                shutil.copy(os.path.join(subdir,file),rootdir)
                os.remove(os.path.join(subdir,file))
            elif subdir==rootdir:
                pass
            else:
                i=2
                while os.path.exists(os.path.join(rootdir,"Copy {0} of ".format(i)+file)):
                    i+=1
                shutil.copy(os.path.join(subdir,file),os.path.join(rootdir,"Copy {0} of ".format(i)+file))
                os.remove(os.path.join(subdir,file))
    for subdir,dirs,files in os.walk(rootdir):
        for dir in dirs:
            shutil.rmtree(os.path.join(subdir,dir))
        break
def unzip (rootdir,email_formatting_status):
    while (zip_found(rootdir)):
        process_zips(rootdir,email_formatting_status)
def add_directory_slipsheets (rootdir,tabdepth):
    tabdepth+=rootdir.count(os.path.sep)-1
    for subdir,dirs,files in os.walk(rootdir):
        if subdir.count(os.path.sep)<=tabdepth:
            trash,temp = os.path.split(subdir)
            a = os.path.join(subdir,"000 TAB ----- "+temp+".pdf")
            shutil.copy("blankpage.pdf",a)
            add_directory_slipsheet(a,rootdir)
def add_directory_slipsheet(path,rootdir):
    pdf = PdfFileReader(path,strict=False)
    pdf_dest = PdfFileWriter()
    head,file = os.path.split(path)
    head,directory = os.path.split(head)
    add_slipsheet(pdf_dest,directory)
    pdf_write(pdf_dest,path,rootdir)
def rename_resize_rotate(rootdir,numbering_status,slipsheets_status,tabdepth,max_pdf_size,max_excel_size):
    total_size = get_size(rootdir)
    size_so_far = 0
    start_time = time.time()
    n=-1
    if tabdepth==0:
        n+=1
    for filename in customwalk(rootdir):
        previous_percentage=app.progress_bar["value"]
        size_so_far+=os.path.getsize(filename)
        percentage = (float(size_so_far)/float(total_size)*50)+50
        n+=1
        subdir,file = os.path.split(filename)
        if numbering_status==1:
            notify ("RRRing {0}".format(os.path.join(subdir,"{0:05d} ".format(n) + file)),alt='RRRing {0}'.format('{0:05d} '.format(n)+file))
            os.rename(os.path.join(subdir,file),os.path.join(subdir,"{0:05d} ".format(n) + file))
            if file[-4:].upper()==".PDF":
                process_pdf(os.path.join(subdir,"{0:05d} ".format(n) + file),rootdir,slipsheets_status,max_pdf_size,max_excel_size,previous_percentage,percentage,start_time)
        else:
            notify ("RRRing {0}".format(os.path.join(subdir,file)),alt='RRRing {0}'.format(file))
            if file[-4:].upper()==".PDF":
                process_pdf(os.path.join(subdir,file),rootdir,slipsheets_status,max_pdf_size,max_excel_size,previous_percentage,percentage,start_time)
        app.progress_bar["value"]=percentage
        app.progress_bar.update()
        app.progress_text["text"]="{0}%".format(round(percentage,2))
        app.progress_text.update()
        seconds_elapsed = time.time()-start_time
        task_percentage = (percentage-50)*2
        if task_percentage == 0:
            task_percentage = 0.01
        time_remaining = (100-task_percentage)/task_percentage*seconds_elapsed
        app.progress_text2["text"]="Processing PDFs. {0} remaining.".format(time_as_string(int(time_remaining)))
        app.progress_text2.update()
def customwalk(rootdir):
    data = []
    tree = [rootdir]
    tree2 = get_directory_list(rootdir)
    for dir in tree2:
        tree.append(dir)
    for dir in tree:
        files = get_file_list(dir)
        if files!=None:
            for file in files:
                data.append(file)
    return data
def get_directory_list(rootdir):
    data = []
    for set in os.walk(rootdir):
        subdir,dirs,files = set
        break
    if dirs!=[]:
        dirs = customsorted(dirs)
        for dir in dirs:
            data.append(os.path.join(subdir,dir))
            for dir2 in get_directory_list(os.path.join(subdir,dir)):
                data.append(dir2)
            if data[len(data)-1]==[]:
                data.remove([])
    return data
def get_file_list(rootdir):
    data = []
    files = []
    for set in os.walk(rootdir):
        subdir,dirs,files = set
        break
    if files!=[]:
        files = customsorted(files)
    for file in files:
        data.append(os.path.join(rootdir,file))
    return data
def customsorted(files):
    indices = []
    for file in files:
        index = []
        index.append(file)
        split_pre = file.split(" ")[0]
        split = split_pre.split(".")
        for unit in split:
            index.append(CustomSortUnit(unit))
        indices.append(index)
    for index in indices:
        while len(index)<10:
            index.append(None)
    sorted_list = sorted(indices, key=itemgetter(1,2,3,4,5,6,7,8,9))
    return_list = []
    for i in sorted_list:
        return_list.append(i[0])
    return return_list
class CustomSortUnit():
    def __init__(self,unit):
        self.arabic_numeral = None
        self.roman_numeral = None
        self.single_character = None
        self.string = None
        if unit.isdigit():
            self.arabic_numeral = int(unit)
        if is_roman_numeral(unit):
            self.roman_numeral = roman_to_arabic(unit)
        if len(unit)==1:
            self.single_character=unit
        self.string = unit
    def __eq__(self,other):
        if other==None:
            return False
        elif self.arabic_numeral!=None and other.arabic_numeral!=None:
            return self.arabic_numeral==(other.arabic_numeral)
        elif self.roman_numeral!=None and other.roman_numeral!=None:
            return self.roman_numeral==(other.roman_numeral)
        elif self.single_character!=None and other.single_character!=None:
            return self.single_character==(other.single_character)
        elif self.string!=None and other.string!=None:
            return self.string==(other.string)
        elif self.arabic_numeral!=None and other.roman_numeral!=None:
            return self.arabic_numeral==(other.roman_numeral)
        elif self.roman_numeral!=None and other.arabic_numeral!=None:
            return self.roman_numeral==(other.arabic_numeral)
        elif self.arabic_numeral!=None and other.single_character!=None:
            return self.arabic_numeral==(other.single_character)
        elif self.single_character!=None and other.arabic_numeral!=None:
            return self.single_character==(other.arabic_numeral)
        elif self.arabic_numeral!=None and other.string!=None:
            return self.arabic_numeral==(other.string)
        elif self.string!=None and other.arabic_numeral!=None:
            return self.string==(other.arabic_numeral)
        elif (self.roman_numeral!=None and other.single_character!=None) or (self.single_character!=None and other.roman_numeral!=None):
            return self.single_character==(other.single_character)
        elif (self.roman_numeral!=None and other.string!=None) or (self.string!=None and other.roman_numeral!=None):
            return self.string==(other.string)
        elif (self.single_character!=None and other.string!=None) or (self.string!=None and other.single_character!=None):
            return self.string==(other.string)
    def __lt__(self,other):
        if other==None:
            return False
        if self.arabic_numeral!=None and other.arabic_numeral!=None:
            return self.arabic_numeral<(other.arabic_numeral)
        elif self.roman_numeral!=None and other.roman_numeral!=None:
            return self.roman_numeral<(other.roman_numeral)
        elif self.single_character!=None and other.single_character!=None:
            return self.single_character<(other.single_character)
        elif self.string!=None and other.string!=None:
            return self.string<(other.string)
        elif self.arabic_numeral!=None and other.roman_numeral!=None:
            return self.arabic_numeral<(other.roman_numeral)
        elif self.roman_numeral!=None and other.arabic_numeral!=None:
            return self.roman_numeral<(other.arabic_numeral)
        elif self.arabic_numeral!=None and other.single_character!=None:
            return self.arabic_numeral<(other.single_character)
        elif self.single_character!=None and other.arabic_numeral!=None:
            return self.single_character<(other.arabic_numeral)
        elif self.arabic_numeral!=None and other.string!=None:
            return self.arabic_numeral<(other.string)
        elif self.string!=None and other.arabic_numeral!=None:
            return self.string<(other.arabic_numeral)
        elif (self.roman_numeral!=None and other.single_character!=None) or (self.single_character!=None and other.roman_numeral!=None):
            return self.single_character<(other.single_character)
        elif (self.roman_numeral!=None and other.string!=None) or (self.string!=None and other.roman_numeral!=None):
            return self.string<(other.string)
        elif (self.single_character!=None and other.string!=None) or (self.string!=None and other.single_character!=None):
            return self.string<(other.string)
    def __gt__(self,other):
        if other==None:
            return True
        if self.arabic_numeral!=None and other.arabic_numeral!=None:
            return self.arabic_numeral>(other.arabic_numeral)
        elif self.roman_numeral!=None and other.roman_numeral!=None:
            return self.roman_numeral>(other.roman_numeral)
        elif self.single_character!=None and other.single_character!=None:
            return self.single_character>(other.single_character)
        elif self.string!=None and other.string!=None:
            return self.string>(other.string)
        elif self.arabic_numeral!=None and other.roman_numeral!=None:
            return self.arabic_numeral>(other.roman_numeral)
        elif self.roman_numeral!=None and other.arabic_numeral!=None:
            return self.roman_numeral>(other.arabic_numeral)
        elif self.arabic_numeral!=None and other.single_character!=None:
            return self.arabic_numeral>(other.single_character)
        elif self.single_character!=None and other.arabic_numeral!=None:
            return self.single_character>(other.arabic_numeral)
        elif self.arabic_numeral!=None and other.string!=None:
            return self.arabic_numeral>(other.string)
        elif self.string!=None and other.arabic_numeral!=None:
            return self.string>(other.arabic_numeral)
        elif (self.roman_numeral!=None and other.single_character!=None) or (self.single_character!=None and other.roman_numeral!=None):
            return self.single_character>(other.single_character)
        elif (self.roman_numeral!=None and other.string!=None) or (self.string!=None and other.roman_numeral!=None):
            return self.string>(other.string)
        elif (self.single_character!=None and other.string!=None) or (self.string!=None and other.single_character!=None):
            return self.string>(other.string)
    def __le__(self,other):
        if other==None:
            return False
        if self.arabic_numeral!=None and other.arabic_numeral!=None:
            return self.arabic_numeral<=(other.arabic_numeral)
        elif self.roman_numeral!=None and other.roman_numeral!=None:
            return self.roman_numeral<=(other.roman_numeral)
        elif self.single_character!=None and other.single_character!=None:
            return self.single_character<=(other.single_character)
        elif self.string!=None and other.string!=None:
            return self.string<=(other.string)
        elif self.arabic_numeral!=None and other.roman_numeral!=None:
            return self.arabic_numeral<=(other.roman_numeral)
        elif self.roman_numeral!=None and other.arabic_numeral!=None:
            return self.roman_numeral<=(other.arabic_numeral)
        elif self.arabic_numeral!=None and other.single_character!=None:
            return self.arabic_numeral<=(other.single_character)
        elif self.single_character!=None and other.arabic_numeral!=None:
            return self.single_character<=(other.arabic_numeral)
        elif self.arabic_numeral!=None and other.string!=None:
            return self.arabic_numeral<=(other.string)
        elif self.string!=None and other.arabic_numeral!=None:
            return self.string<=(other.arabic_numeral)
        elif (self.roman_numeral!=None and other.single_character!=None) or (self.single_character!=None and other.roman_numeral!=None):
            return self.single_character<=(other.single_character)
        elif (self.roman_numeral!=None and other.string!=None) or (self.string!=None and other.roman_numeral!=None):
            return self.string<=(other.string)
        elif (self.single_character!=None and other.string!=None) or (self.string!=None and other.single_character!=None):
            return self.string<=(other.string)
    def __ge__(self,other):
        if other==None:
            return True
        if self.arabic_numeral!=None and other.arabic_numeral!=None:
            return self.arabic_numeral>=(other.arabic_numeral)
        elif self.roman_numeral!=None and other.roman_numeral!=None:
            return self.roman_numeral>=(other.roman_numeral)
        elif self.single_character!=None and other.single_character!=None:
            return self.single_character>=(other.single_character)
        elif self.string!=None and other.string!=None:
            return self.string>=(other.string)
        elif self.arabic_numeral!=None and other.roman_numeral!=None:
            return self.arabic_numeral>=(other.roman_numeral)
        elif self.roman_numeral!=None and other.arabic_numeral!=None:
            return self.roman_numeral>=(other.arabic_numeral)
        elif self.arabic_numeral!=None and other.single_character!=None:
            return self.arabic_numeral>=(other.single_character)
        elif self.single_character!=None and other.arabic_numeral!=None:
            return self.single_character>=(other.arabic_numeral)
        elif self.arabic_numeral!=None and other.string!=None:
            return self.arabic_numeral>=(other.string)
        elif self.string!=None and other.arabic_numeral!=None:
            return self.string>=(other.arabic_numeral)
        elif (self.roman_numeral!=None and other.single_character!=None) or (self.single_character!=None and other.roman_numeral!=None):
            return self.single_character>=(other.single_character)
        elif (self.roman_numeral!=None and other.string!=None) or (self.string!=None and other.roman_numeral!=None):
            return self.string>=(other.string)
        elif (self.single_character!=None and other.string!=None) or (self.string!=None and other.single_character!=None):
            return self.string>=(other.string)
    def __ne__(self,other):
        if other==None:
            return True
        if self.arabic_numeral!=None and other.arabic_numeral!=None:
            return self.arabic_numeral!=(other.arabic_numeral)
        elif self.roman_numeral!=None and other.roman_numeral!=None:
            return self.roman_numeral!=(other.roman_numeral)
        elif self.single_character!=None and other.single_character!=None:
            return self.single_character!=(other.single_character)
        elif self.string!=None and other.string!=None:
            return self.string!=(other.string)
        elif self.arabic_numeral!=None and other.roman_numeral!=None:
            return self.arabic_numeral!=(other.roman_numeral)
        elif self.roman_numeral!=None and other.arabic_numeral!=None:
            return self.roman_numeral!=(other.arabic_numeral)
        elif self.arabic_numeral!=None and other.single_character!=None:
            return self.arabic_numeral!=(other.single_character)
        elif self.single_character!=None and other.arabic_numeral!=None:
            return self.single_character!=(other.arabic_numeral)
        elif self.arabic_numeral!=None and other.string!=None:
            return self.arabic_numeral!=(other.string)
        elif self.string!=None and other.arabic_numeral!=None:
            return self.string!=(other.arabic_numeral)
        elif (self.roman_numeral!=None and other.single_character!=None) or (self.single_character!=None and other.roman_numeral!=None):
            return self.single_character!=(other.single_character)
        elif (self.roman_numeral!=None and other.string!=None) or (self.string!=None and other.roman_numeral!=None):
            return self.string!=(other.string)
        elif (self.single_character!=None and other.string!=None) or (self.string!=None and other.single_character!=None):
            return self.string!=(other.string)
def is_roman_numeral(input):
    valid_numerals=["M","D","C","L","X","V","I"]
    for letter in input.upper():
        if letter not in valid_numerals:
            return False
    return True
def roman_to_arabic(input):
    total = 0
    input=input.upper()
    for i in range(len(input)):
        char = roman_char_to_arabic(input[i])
        if i!=(len(input)-1):
            next_char = roman_char_to_arabic(input[i+1])
        else:
            next_char = 0
        if next_char>char:
            total-=char
        else:
            total+=char
    return total
def roman_char_to_arabic(char):
    table = [['M',1000],['D',500],['C',100],['L',50],['X',10],['V',5],['I',1]]
    for i in table:
        if char==i[0]:
            return i[1]
def get_size(rootdir):
    total = 0
    for subdir,dirs,files in os.walk(rootdir):
        for file in files:
            total+=os.path.getsize(os.path.join(subdir,file))
    return total
def get_size_without_pdfs(rootdir):
    total = 0
    for subdir,dirs,files in os.walk(rootdir):
        for file in files:
            if file[-4:].upper()!=".PDF":
                total+=os.path.getsize(os.path.join(subdir,file))
    return total
def time_as_string(seconds):
    s = seconds
    d = s/60/60/24
    s-=(d*60*60*24)
    h = s/60/60
    s-=(h*60*60)
    m = s/60
    s-=(m*60)
    value = ""
    if d>0:
        value+="{0}d".format(d)
    if h>0 or d>0:
        value+="{0}h".format(h)
    if m>0 or h>0 or d>0:
        value+="{0:02d}m".format(m)
    value+="{0:02d}s".format(s)
    return value
def convert_to_pdf(rootdir,page_setup_settings):
    total_size = get_size_without_pdfs(rootdir)
    size_so_far=0
    start_time = time.time()
    for i in range(len(files_to_remove)):
        files_to_remove[i] = files_to_remove[i].upper()
    for subdir,dirs,files in os.walk(rootdir):
        for file in files:
            if file[-4:].upper()!=".PDF":
                size_so_far+=os.path.getsize(os.path.join(subdir,file))
                percentage = float(size_so_far)/float(total_size)*50
            try:
                notify("Converting {0}".format(os.path.join(subdir,file)),alt='Converting {0}'.format(file))
            except:
                notify("Converting <NAME CANNOT BE DISPLAYED>.")
            if os.path.join(subdir,file).upper() in files_to_remove:
                pass
            elif file[-4:].upper()==".DOC" or file[-5:].upper()==".DOCX" or file[-4:].upper()==".TXT" or file[-4:].upper()==".RTF" or file[-5:].upper()==".HTML":
                process_doc(rootdir,os.path.join(subdir,file))
                shortened = name_with_first_extension(file)
                if os.path.isdir(os.path.join(subdir,shortened)+'_files'):
                    shutil.rmtree(os.path.join(subdir,shortened)+'_files')
            elif file[-4:].upper()==".XLS" or file[-5:].upper()==".XLSX":
                process_xls(rootdir,os.path.join(subdir,file),page_setup_settings)
            elif file[-4:].upper()==".TIF" or file[-4:].upper()==".JPG" or file[-5:].upper()==".TIFF" or file[-5:].upper()==".JPEG" or file[-4:].upper()==".PNG":
                process_image(os.path.join(subdir,file),rootdir)
            elif file[-4:].upper()==".PPT" or file[-5:].upper()==".PPTX":
                process_ppt(os.path.join(subdir,file),rootdir)
            # elif file[-4:].upper()==".PDF" and pdf_reprocess_status==1:
                # reprocess_pdf(os.path.join(subdir,file))
            elif file[-4:].upper()==".PDF":
                pass
            else:
                process_misc(os.path.join(subdir,file))
            if file[-4:].upper()!=".PDF":
                app.progress_bar["value"]=percentage
                app.progress_bar.update()
                app.progress_text["text"]="{0}%".format(round(percentage,2))
                app.progress_text.update()
                seconds_elapsed = time.time()-start_time
                task_percentage = percentage*2
                if task_percentage == 0:
                    task_percentage = 0.01
                time_remaining = (100-task_percentage)/task_percentage*seconds_elapsed
                app.progress_text2["text"]="Converting files to PDF format. {0} remaining.".format(time_as_string(int(time_remaining)))
                app.progress_text2.update()
    for file in files_to_remove:
        try:
            os.remove(file)
        except:
            pass
def name_with_first_extension(name):
    parts = name.split('.')
    return parts[0]+'.'+parts[1]
# def reprocess_pdf(filename):
    # acrobat = comtypes.client.CreateObject('AcroExch.App')
    # acrobat.Hide()
    # pdf = comtypes.client.CreateObject('AcroExch.AVDoc')
    # pdf.Open(filename,'temp')
    # pdf2 = pdf.GetPDDoc()
    # jso = pdf2.GetJSObject()
    # docs = jso.app.activeDocs
    # for doc in docs:
        # doc.flattenPages()
        # doc.saveAs(filename+".pdf")
    # acrobat.CloseAllDocs()
    # acrobat.Exit()
    # os.remove(filename)
def process_pdf(filename,rootdir,slipsheets_status,max_pdf_size,max_excel_size,previous_percentage,percentage,start_time):
    try:
        pdf = PdfFileReader(filename,strict=False)
        if pdf.isEncrypted:
            notify("ERROR: File encrypted - check PDF",level='w')
            copy_to_root(filename,rootdir)
            return
        pdf_dest = PdfFileWriter()
        if slipsheets_status==1:
            add_slipsheet(pdf_dest,filename)
        if filename[-8:].upper()==".XLS.PDF" or filename[-9:].upper()==".XLSX.PDF":
            max_size=max_excel_size
        else:
            max_size=max_pdf_size
        process_pdf_pages(pdf,pdf_dest,max_size,previous_percentage,percentage,start_time)
        pdf_write(pdf_dest,filename,rootdir)
    except:
        notify("ERROR PROCESSING PDF",level='w')
        copy_to_root(filename,rootdir)
def copy_to_root(filename,rootdir):
    trash,temp = os.path.split(filename)
    if os.path.exists(os.path.join(rootdir,temp))==False:
        shutil.copy(filename,rootdir)
    elif os.path.join(rootdir,temp)==filename:
        pass
    else:
        i=2
        while os.path.exists(os.path.join(rootdir,"Copy {0} of ".format(i)+temp)):
            i+=1
        shutil.copy(filename,os.path.join(rootdir,"Copy {0} of ".format(i)+temp))
    return
def process_misc(filename):
    os.remove(filename)
    pdf_dest = PdfFileWriter()
    add_slipsheet(pdf_dest,"File Unprintable")
    pdf_dest.write(io.open(filename+".pdf",mode='w+b'))
def process_pdf_page(page):
    x,y = get_rotated_page_dimensions(page)
    if x>y:
        page.rotateCounterClockwise(90)
        x,y = get_rotated_page_dimensions(page)
    scale_to_letter(page,scale_factor(x,y))
def determine_scaling_factors(x,y):
    x_scaling = 1
    y_scaling = 1
    if x>(8.5*72):
        x_scaling = float(8.5*72)/float(x)
    if y>(11*72):
        y_scaling = float(11*72)/float(y)
    return (x_scaling,y_scaling)
def scale_factor(x,y):
    scale = 1
    x_scaling,y_scaling = determine_scaling_factors(x,y)
    if x_scaling != 1 or y_scaling != 1:
        if x_scaling < y_scaling:
            scale = x_scaling
        else:
            scale = y_scaling
    return scale
def scale_to_letter(page,scale):
    try:
        page.scaleBy(scale)
    except:
        notify("ERROR: Could not scale - check PDF",level='w')
    a,b,c,d,x,y = get_page_dimensions(page)
    x_target,y_target = get_target_dimensions(page)
    a,b,c,d = adjust_page_dimensions(a,b,c,d,x,y,x_target,y_target)    
    update_page_dimensions(page,a,b,c,d)
def get_page_dimensions(page):
    a,b,c,d = page.mediaBox
    a = float(a)
    b = float(b)
    c = float(c)
    d = float(d)
    x = abs(c-a)
    y = abs(d-b)
    return (a,b,c,d,x,y)
def get_target_dimensions(page):
    is_rotated = False
    rotation = page.get('/Rotate')
    if rotation == 90 or rotation == -90 or rotation == 270 or rotation == -270:
        is_rotated = True
    if is_rotated == False:
        x_target = (8.5*72)
        y_target = (11*72)
    else:
        x_target = (11*72)
        y_target = (8.5*72)
    return (x_target,y_target)
def adjust_page_dimensions(a,b,c,d,x,y,x_target,y_target):
    if x<x_target:
        a-=(x_target-float(x))/2
        c+=(x_target-float(x))/2
    if y<y_target:
        b-=(y_target-float(y))/2
        d+=(y_target-float(y))/2
    return (a,b,c,d)
def update_page_dimensions(page,a,b,c,d):
    page.mediaBox.upperLeft = (a,b)
    page.mediaBox.lowerRight = (c,d)
    page.cropBox.upperLeft = (a,b)
    page.cropBox.lowerRight = (c,d)
def zip_found(rootdir):
    return_value = False
    for subdir,dirs,files in os.walk(rootdir):
        for file in files:
            if file[-4:].upper()==".ZIP" or file[-4:].upper()==".MSG" or file[-4:].upper()==".EML":
                return_value = True
    return return_value
def process_zips(rootdir,email_formatting_status):
    for subdir,dirs,files in os.walk(rootdir):
        for file in files:
            if file[-4:].upper()==".ZIP":
                process_zip(subdir,file)
            elif file[-4:].upper()==".MSG" or file[-4:].upper()==".EML":
                process_msg(subdir,file,email_formatting_status)
def process_zip(subdir,file):
    os.mkdir(os.path.join(subdir,file)+".dir")
    zip = zipfile.ZipFile(os.path.join(subdir,file))
    zip.extractall(os.path.join(subdir,file)+".dir")
    zip.close()
    os.remove(os.path.join(subdir,file))
def process_mime_msg(subdir,file):
    os.mkdir(os.path.join(subdir,file)+".dir")
    fp = open(os.path.join(subdir,file))
    msg = email.message_from_file(fp)
    fp.close()
    counter = 1
    for part in msg.walk():
        counter = process_mime_msg_section(part,subdir,file,counter)
    os.remove(os.path.join(subdir,file))
def process_mime_msg_section(part,subdir,file,counter):
    if part.get_content_maintype() == 'multipart':
        return counter
    filename = part.get_filename()
    if filename==None:
        filename = generate_mime_msg_section_filename(part,counter)
    counter += 1
    fp = open(os.path.join(os.path.join(subdir,file)+".dir",filename),'wb')
    fp.write(part.get_payload(decode=True))
    fp.close()
    return counter
def generate_mime_msg_section_filename(part,counter):
    ext = mimetypes.guess_extension(part.get_content_type())
    if not ext:
        ext = ".bin"
    filename = "part-{0:03d}{1}".format(counter,ext)
    return filename
def process_msg(subdir,file,email_formatting_status):
    if olefile.isOleFile(os.path.join(subdir,file))==False:
        process_mime_msg(subdir,file)
        return
    os.mkdir(os.path.join(subdir,file)+".dir")
    ole = olefile.OleFileIO(os.path.join(subdir,file))
    attach_list = get_msg_attach_list(ole)
    extract_msg_files(attach_list,ole,subdir,file)
    if email_formatting_status==1:
        extract_msg_message(ole,subdir,file)
    else:
        extract_msg_message_plaintext(ole,subdir,file)
    ole.close()
    os.remove(os.path.join(subdir,file))
def get_msg_attach_list(ole):
    attach_list = []
    for i in ole.listdir():
        if i[0][:8]=="__attach":
            if attach_list.count(i[0])==0:
                attach_list.append(i[0])
    return attach_list
def extract_msg_files(attach_list,ole,subdir,file):
    for i in attach_list:
        filename = clean_string(get_msg_attachment_filename(i,ole))
        write_msg_attachment(i,ole,subdir,file,filename)
def get_msg_attachment_filename(index,ole):
    filename = get_msg_attachment_filename_primary(index,ole)
    if filename == None:
        filename = get_msg_attachment_filename_fallback(index,ole)
    if filename == None:
        filename = "ATTACHMENT 1"
    return filename
def get_msg_attachment_filename_primary(index,ole):
    filename = None
    for i in ole.listdir():
        if i[0]==index:
            if i[1][:16]=="__substg1.0_3707":
                name_stream = ole.openstream("{0}/{1}".format(i[0],i[1]))
                filename = name_stream.read()
    return filename
def get_msg_attachment_filename_fallback(index,ole):
    filename = None
    for i in ole.listdir():
        if i[0]==index:
            if i[1][:16]=="__substg1.0_3704":
                name_stream = ole.openstream("{0}/{1}".format(i[0],i[1]))
                filename = name_stream.read()
    return filename
def write_msg_attachment(index,ole,subdir,file,filename):
    for i in ole.listdir():
        if i[0]==index:
            if i[1][:20]=="__substg1.0_37010102":
                file_stream = ole.openstream("{0}/{1}".format(i[0],i[1]))
                file_data = file_stream.read()
                try:
                    fp = io.open(os.path.join(os.path.join(subdir,file)+".dir",filename),"w+b")
                    fp.write(file_data)
                    fp.close()
                except:
                    pass
def extract_msg_message_plaintext(ole,subdir,file):
    msg_from,msg_to,msg_cc,msg_subject,msg_header,msg_body = extract_msg_message_data_plaintext(ole)
    fp = io.open(os.path.join(os.path.join(subdir,file)+".dir","00 {0}.txt".format(file)),"w")
    try:
        fp.write("From: {0}\nTo: {1}\nCC: {2}\nSubject: {3}\n".format(msg_from,msg_to,msg_cc,msg_subject).decode('utf-8'))
    except:
        fp.write("From: {0}\nTo: {1}\nCC: {2}\nSubject: {3}\n".format(msg_from,msg_to,msg_cc,msg_subject).decode('ISO-8859-1'))
    fp.write(unicode("---------------\n\n"))
    try:
        fp.write(msg_body.decode('utf-8'))
    except:
        fp.write(msg_body.decode('ISO-8859-1'))
    fp.close()
def extract_msg_message(ole,subdir,file):
    msg_from,msg_to,msg_cc,msg_subject,msg_header,msg_body = extract_msg_message_data(ole)
    if msg_body=='' or msg_body==None:
        extract_msg_message_plaintext(ole,subdir,file)
        return
    try:
        msg_body = extract_msg_unpack_rtf(msg_body)
    except:
        notify('ERROR: FAILED TO UNPACK RTF MESSAGE FROM {0}. PROCESSING AS PLAINTEXT.'.format(os.path.join(subdir,file)),level='w')
        extract_msg_message_plaintext(ole,subdir,file)
        return
    try:
        msg_body = extract_msg_unpack_html(msg_body)
    except:
        notify('ERROR: FAILED TO UNPACK HTML MESSAGE FROM {0}. PROCESSING AS PLAINTEXT.'.format(os.path.join(subdir,file)),level='w')
        extract_msg_message_plaintext(ole,subdir,file)
        return
    if msg_body=='' or msg_body==None:
        extract_msg_message_plaintext(ole,subdir,file)
        return
    offset = string.find(msg_body,'<body')
    offset = string.find(msg_body,'>',offset)+1
    msg_body = msg_body[:offset]+'<p class=MsoNormal><span style=\'font-size:10.0pt;font-family:"Tahoma","sans-serif"\'><b>From:</b> {0}</span></p>\n'.format(msg_from)+'<p class=MsoNormal><span style=\'font-size:10.0pt;font-family:"Tahoma","sans-serif"\'><b>To:</b> {0}</span></p>\n'.format(msg_to)+'<p class=MsoNormal><span style=\'font-size:10.0pt;font-family:"Tahoma","sans-serif"\'><b>CC:</b> {0}</span></p>\n'.format(msg_cc)+'<p class=MsoNormal><span style=\'font-size:10.0pt;font-family:"Tahoma","sans-serif"\'><b>Subject:</b> {0}</span></p>\n'.format(msg_subject)+msg_body[offset:]
    if '<body' not in msg_body:
        msg_body = '<body>'+msg_body+'</body>'
    if '<html' not in msg_body:
        msg_body = '<html>'+msg_body+'</html>'
    fp = io.open(os.path.join(os.path.join(subdir,file)+".dir","00 {0}.html".format(file)),"wb")
    msg_body = clean_html_attachments(os.path.join(subdir,file)+'.dir',msg_body)
    fp.write(msg_body)
    fp.close()
    for attachment in get_html_attachments(msg_body):
        if attachment[:6]!='http:/':
            files_to_remove.append(os.path.join(subdir,file+".dir",attachment))
def clean_html_attachments(subdir,html):
    i=0
    for attachment in get_html_attachments(html):
        i+=1
        if attachment[:6]!='http:/':
            if not os.path.exists(os.path.join(subdir,attachment)):
                prefix = 'ATT{0:05d}.'.format(i)
                name = ''
                for suffix in ['jpg', 'jpeg', 'png', 'gif', 'tif', 'tiff', 'bmp', 'JPG', 'JPEG', 'PNG', 'GIF', 'TIF', 'TIFF', 'BMP']:
                    if os.path.exists(os.path.join(subdir,prefix+suffix)):
                        name=prefix+suffix
                if name!='':
                    html = html.replace('src="'+attachment+'"','src="'+name+'"',1)
    return html
    
def get_html_attachments(html):
    indices = [m.start() for m in re.finditer('<img',html)]
    values = []
    for index in indices:
        start = string.find(html,'src="',index)+5
        stop = string.find(html,'"',start)
        values.append(html[start:stop])
    return values
def extract_msg_unpack_html(source):
    output = ''
    if '\\fromhtml1' not in source:
        return
    position = 0
    escape_character=False
    group = -1
    font_group = 9999
    htmlrtf = False
    strip_cid_tag = False
    while position<len(source):
        to_write=True
        if escape_character==True:
            to_write=True
            escape_character=False
        elif source[position]=='\r' or source[position]=='\n':
            to_write=False
        elif source[position]=='\\':
            to_write=False
            token,number,position = get_token(source,position)
            if token=='':
                escape_character=True
            if token=='par':
                output+=u'\n'
            if token=='tab':
                output+=u'\t'
            if token=="'":
                if number!=False:
                    if htmlrtf==False:
                        output+='&#'+str(int(number,16))+';'
            if token=='fonttbl' or token=='colortbl':
                font_group = group
            if token=='htmlrtf' and number==None:
                htmlrtf=True
            elif token=='htmlrtf' and number==0:
                htmlrtf=False
        elif htmlrtf==True:
            to_write=False
        elif source[position]=='{':
            to_write=False
            group +=1
        elif source[position]=='}':
            to_write=False
            group -=1
            if font_group>group:
                font_group=9999
        elif font_group<=group:
            to_write=False
        elif source[position]=='s' and source[position:position+9]=='src="cid:':
            output+='src="'
            position+=9
            strip_cid_tag = True
        elif source[position]=='@' and strip_cid_tag==True:
            while source[position]!='"':
                position+=1
            strip_cid_tag=False
        if to_write==True:
            output+=source[position]
        position+=1
    return output
def get_token(source,position):
    value = ''
    numeric_value=False
    position+=1
    while source[position] in 'abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ*\'-':
        value+=source[position]
        position+=1
        if value=="'":
            if source[position] in '1234567890ABCDEFabcdef' and source[position+1] in '1234567890ABCDEFabcdef':
                for i in range(2):
                    if numeric_value==False:
                        numeric_value=source[position]
                    else:
                        numeric_value+=source[position]
                    position+=1
                    if position>=len(source):
                        if source[position]==' ':
                            position+=1
                        return (value,numeric_value,position-1)
            if source[position]==' ':
                position+=1
            return (value,numeric_value,position-1)
        if value=='*' and source[position]=='\\':
            value=''
            position+=1
        if position>=len(source):
            break
    if source[position] in '1234567890':
        while source[position] in '1234567890':
            if numeric_value==False:
                numeric_value=source[position]
            else:
                numeric_value+=source[position]
            position+=1
            if position>=len(source):
                break
    if numeric_value!=False:
        numeric_value = int(numeric_value)
    else:
        numeric_value=None
    if source[position]==' ':
        position+=1
    return (value,numeric_value,position-1)
def extract_msg_unpack_rtf(text):
    output = ''
    control=True
    control_number=0
    controls = [None,None,None,None,None,None,None,None]
    dictionary=list('{\\rtf1\\ansi\\mac\\deff0\\deftab720{\\fonttbl;}{\\f0\\fnil \\froman \\fswiss \\fmodern \\fscript \\fdecor MS Sans SerifSymbolArialTimes New RomanCourier{\\colortbl\\red0\\green0\\blue0\x0d\x0a\\par \\pard\\plain\\f0\\fs20\\b\\i\\u\\tab\\tx')
    dictionary_write_offset=len(dictionary)
    end_offset=len(dictionary)
    for i in range(4096-len(dictionary)):
        dictionary.append('')
    pos=16
    while pos<len(text):
        char = text[pos]
        if control==True:
            control=False
            control_number=-1
            controls = extract_bits(char)
        elif control==False:
            control_number+=1
            if control_number==8:
                control=True
                continue
            else:
                if controls[control_number]==0:
                    output+=char
                    dictionary[dictionary_write_offset]=char
                    dictionary_write_offset+=1
                    if end_offset<4096:
                        end_offset+=1
                    if dictionary_write_offset>=4096:
                        dictionary_write_offset=0
                elif controls[control_number]==1:
                    offset,num_bytes = extract_dictionary_offset(text[pos:pos+2])
                    if offset==dictionary_write_offset:
                        break
                    dictionary_output = extract_dictionary_text(offset,num_bytes,dictionary,end_offset)
                    for char in dictionary_output:
                        dictionary[dictionary_write_offset]=char
                        dictionary_write_offset+=1
                        if end_offset<4096:
                            end_offset+=1
                        if dictionary_write_offset>=4096:
                            dictionary_write_offset=0
                    output+=dictionary_output
                    pos+=1
        pos+=1
    return output
def extract_dictionary_offset(input):
    bits1 = extract_bits(input[0],big_endian=True)
    bits2 = extract_bits(input[1],big_endian=True)
    offset = generate_dictionary_offset(bits1,bits2)
    num_bytes = generate_dictionary_length(bits2)
    return offset,num_bytes
def extract_bits(char,big_endian=False):
    if big_endian==False:
        start_position=7
        end_position=-1
        direction=-1
    else:
        start_position=0
        end_position=8
        direction=1
    bm = [0b10000000,0b01000000,0b00100000,0b00010000,0b00001000,0b00000100,0b00000010,0b00000001]
    bits = []
    for offset in range(start_position,end_position,direction):
        bits.append((bm[offset]&ord(char))>>(7-offset))
    return bits
def generate_dictionary_offset(input1,input2):
    a = 2048
    result=0
    for bit in input1:
        result+=(bit*a)
        a/=2
    for i in range(0,4):
        result+=(input2[i]*a)
        a/=2
    return result
def generate_dictionary_length(input2):
    a = 8
    result=0
    for i in range(4,8):
        result+=(input2[i]*a)
        a/=2
    return result+2
def extract_dictionary_text(offset,num_bytes,dictionary,end_offset):
    output=''
    for i in range(num_bytes):
        output+=dictionary[circle(offset+i,end_offset)]
    return output
def circle(offset,length):
    if offset<length:
        return offset
    while offset>=length:
        offset-=length
    return offset
def extract_msg_message_data(ole):
    msg_from=""
    msg_to=""
    msg_cc=""
    msg_subject=""
    msg_header=""
    msg_body=""
    for i in ole.listdir():
        if i[0][:16]=="__substg1.0_0C1A":
            msg_from = extract_msg_stream_text(i,ole)
        elif i[0][:16]=="__substg1.0_0E04":
            msg_to = extract_msg_stream_text(i,ole)
        elif i[0][:16]=="__substg1.0_0E03":
            msg_cc = extract_msg_stream_text(i,ole)
        elif i[0][:16]=="__substg1.0_0037":
            msg_subject = extract_msg_stream_text(i,ole)
        elif i[0][:16]=="__substg1.0_007D":
            msg_header = extract_msg_stream_text(i,ole)
        elif i[0][:16]=="__substg1.0_1009":
            msg_body = extract_msg_stream_text_noclean(i,ole)
    return (msg_from,msg_to,msg_cc,msg_subject,msg_header,msg_body)
def extract_msg_message_data_plaintext(ole):
    msg_from=""
    msg_to=""
    msg_cc=""
    msg_subject=""
    msg_header=""
    msg_body=""
    for i in ole.listdir():
        if i[0][:16]=="__substg1.0_0C1A":
            msg_from = extract_msg_stream_text(i,ole)
        elif i[0][:16]=="__substg1.0_0E04":
            msg_to = extract_msg_stream_text(i,ole)
        elif i[0][:16]=="__substg1.0_0E03":
            msg_cc = extract_msg_stream_text(i,ole)
        elif i[0][:16]=="__substg1.0_0037":
            msg_subject = extract_msg_stream_text(i,ole)
        elif i[0][:16]=="__substg1.0_007D":
            msg_header = extract_msg_stream_text(i,ole)
        elif i[0][:16]=="__substg1.0_1000":
            msg_body = extract_msg_stream_text(i,ole)
    return (msg_from,msg_to,msg_cc,msg_subject,msg_header,msg_body)
def extract_msg_stream_text(index,ole):
    stream = ole.openstream(index[0])
    text = stream.read()
    text = clean_string(text)
    return text
def extract_msg_stream_text_noclean(index,ole):
    stream = ole.openstream(index[0])
    text = stream.read()
    return text
def clean_string(input):
    output = ""
    save_flag = False
    for letter in input:
        if save_flag == False:
            save_flag = True
        else:
            save_flag = False
        if save_flag == True:
            output = output + letter    
    return output
def process_doc(rootdir,filename):
    try:
        word = comtypes.client.CreateObject('Word.Application')
        word.Visible = False
        doc = word.Documents.Open(filename)
        doc.SaveAs(filename+".pdf",FileFormat=17)
        doc.Close()
        os.remove(filename)
    except:
        notify("FAILED TO PROCESS CORRUPT FILE",level='w')
        copy_to_root(filename,rootdir)
    finally:
        word.Quit()
def process_xls(rootdir,filename,page_setup_settings):
    try:
        excel = comtypes.client.CreateObject('Excel.Application')
        excel.Visible = False
        xls = excel.Workbooks.Open(filename)
        for i in xls.Sheets:
            if i.Visible==-1:
                i.Select(False)
                if page_setup_settings!=None:
                    set_worksheet_page_setup_settings(page_setup_settings,i.PageSetup)
        xls.SaveAs(filename+".pdf",FileFormat=57)
        xls.Close(False)
        os.remove(filename)
    except:
        notify("FAILED TO PROCESS CORRUPT FILE",level='w')
        copy_to_root(filename,rootdir)
    finally:
        excel.Quit()
def process_ppt(filename,rootdir):
    try:
        powerpoint = comtypes.client.CreateObject('PowerPoint.Application')
        ppt = powerpoint.Presentations.Open(filename)
        ppt.ExportAsFixedFormat(filename+".pdf",FixedFormatType=2)
        ppt.Close()
        os.remove(filename)
    except:
        notify("FAILED TO PROCESS CORRUPT FILE",level='w')
        copy_to_root(filename,rootdir)
    finally:
        powerpoint.Quit()
def process_image(filename,rootdir):
    #image = PythonMagick.Image()
    #image.read(filename)
    #image.write(filename+".pdf")
    #ABOVE CODE IS UNTESTED BUT WILL PROBABLY DO THE TRICK INSTEAD OF THE FOLLOWING IF YOU DON'T HAVE ACROBAT PRO
    try:
        acrobat = comtypes.client.CreateObject('AcroExch.App')
        acrobat.Hide()
        image = comtypes.client.CreateObject('AcroExch.AVDoc')
        image.Open(filename,'temp')
        image2 = image.GetPDDoc()
        image2.Save(1,filename+".pdf")
        acrobat.CloseAllDocs()
        os.remove(filename)
    except:
        notify("FAILED TO PROCESS CORRUPT FILE",level='w')
        copy_to_root(filename,rootdir)
    finally:
        try:
            acrobat.Exit()
        except:
            pass
def add_slipsheet(pdf_dest,text):
    slipsheet = pdf_dest.addBlankPage(8.5*72,11*72)
    packet = StringIO.StringIO()
    can = canvas.Canvas(packet, pagesize=letter)
    directory_path,base_filename = os.path.split(text)
    can.setFontSize(16)
    draw_slipsheet_text(can,base_filename)
    can.save()
    packet.seek(0)
    slipsheet_overlay_pdf=PdfFileReader(packet)
    slipsheet_overlay_page=slipsheet_overlay_pdf.getPage(0)
    slipsheet.mergePage(slipsheet_overlay_page)
def draw_slipsheet_text(can,text):
    if len(text)<66:
        can.drawCentredString(8.5*72/2,600,text)
    else:
        position = 620
        for line in split_into_lines(text,65):
            position-=20
            can.drawCentredString(8.5*72/2,position,line)
def split_into_lines(text,length):
    split_text = text.split(" ")
    lines = []
    while split_text!=[]:
        lines.append(grab_line(split_text,length))
    return lines
def grab_line(text,length):
    build_line = ""
    while len(text)>0 and (len(build_line)+len(text[0])+1)<=length:
        build_line+=" "+text[0]
        text.remove(text[0])
    if len(text)>0 and len(text[0])>length:
        diff = length-len(build_line)-1
        build_line+=" "+text[0][0:diff]
        text[0]=text[0][diff:]
    return build_line
def process_pdf_pages(pdf,pdf_dest,max_size,previous_percentage,percentage,start_time):
    numPages = pdf.getNumPages()
    if numPages>max_size:
        numPages = max_size
        truncate = True
    else:
        truncate = False
    for i in range(numPages):
        page = pdf.getPage(i)
        process_pdf_page(page)
        pdf_dest.addPage(page)
        page_percentage = previous_percentage+((percentage-previous_percentage)*i/numPages)
        app.progress_bar["value"]=page_percentage
        app.progress_bar.update()
        app.progress_text["text"]="{0}%".format(round(page_percentage,2))
        app.progress_text.update()
        seconds_elapsed = time.time()-start_time
        task_percentage = (page_percentage-50)*2
        if task_percentage == 0:
            task_percentage = 0.01
        time_remaining = (100-task_percentage)/task_percentage*seconds_elapsed
        app.progress_text2["text"]="Processing PDFs. {0} remaining.".format(time_as_string(int(time_remaining)))
        app.progress_text2.update()
    if truncate == True:
        add_slipsheet(pdf_dest,"Document truncated due to size - see original file.")
def pdf_write(pdf_dest,filename,rootdir):
    try:
        pdf_dest.write(io.open(filename,mode='w+b'))
    except:
        notify ("ERROR: Could not write PDF - check PDF",level='w')
        copy_to_root(filename,rootdir)
def get_rotated_page_dimensions(page):
    a,b,c,d = page.mediaBox
    x = abs(c-a)
    y = abs(d-b)
    previous_rotation = page.get('/Rotate')
    if previous_rotation == 90 or previous_rotation == -90 or previous_rotation == 270 or previous_rotation == -270:
        temp=x
        x=y
        y=temp
    return (x,y)
def launch_main(sourcedir,destdir,tabdepth,page_setup_settings,pdf_reprocess_status,numbering_status,slipsheets_status,flatten_status,max_pdf_size,max_excel_size):
    if sourcedir==None or destdir==None or tabdepth==None:
        print("One or more missing arguments. Exiting.")
        sys.exit()
    sourcedir = os.path.abspath(sourcedir)
    destdir = os.path.abspath(destdir)
    #os.rmdir(destdir)
    print("Copying files to destination directory...")
    trash,temp = os.path.split(sourcedir)
    shutil.copytree(sourcedir,os.path.join(destdir,temp))
    #shutil.copytree(sourcedir,destdir)
    main(destdir,tabdepth,page_setup_settings,pdf_reprocess_status,numbering_status,slipsheets_status,flatten_status,max_pdf_size,max_excel_size)
def launch_gui():
    root = Tkinter.Tk()
    root.title("RRR")
    global app
    app = Application(master=root)
    app.mainloop()
    root.destroy()
class Application(Tkinter.Frame):
    def __init__(self, master = None):
        Tkinter.Frame.__init__(self,master)
        self.pack()
        self.source_directory = ""
        self.dest_directory = ""
        self.page_setup_settings = None
        self.create_widgets()
        self.source_directory = "C:/datasites/1"
        self.dest_directory = "C:/datasites/2"
        self.chosen_source["text"] = self.source_directory
        self.chosen_dest["text"] = self.dest_directory
        
        self.excel = comtypes.client.CreateObject('Excel.Application')
        self.excel.Visible = False
        self.xls = self.excel.Workbooks.Add()
        for i in self.xls.Sheets:
            i.Select(False)
        self.page_setup_settings = self.xls.Sheets(1).PageSetup
        self.page_setup_settings.CenterHorizontally = True
        self.page_setup_settings.FitToPagesTall = False
        self.page_setup_settings.Orientation = 2.0
        self.page_setup_settings.FitToPagesWide = 1
        self.page_setup_settings.Zoom = False
        self.page_setup_settings.PaperSize = 1.0
    def create_widgets(self):
        self.create_exit()
        self.create_start()
        self.create_choose_source()
        self.create_source_text()
        self.create_choose_dest()
        self.create_dest_text()
        self.create_tab_depth_text()
        self.create_tab_depth_picker()
        self.create_page_setup_button()
        self.create_email_formatting_checkbox()
        self.create_numbering_checkbox()
        self.create_slipsheets_checkbox()
        self.create_progress_bar()
        self.create_progress_text()
        self.create_progress_text2()
        self.create_progress_text3()
        self.create_flatten_checkbox()
        self.create_max_size_text()
        self.create_max_size_picker()
        self.create_max_excel_size_text()
        self.create_max_excel_size_picker()
    def create_max_size_text(self):
        self.max_size_text = Tkinter.Label(self)
        self.max_size_text["text"] = "Maximum PDF Pages:"
        self.max_size_text.grid(row=3,column=1,sticky=Tkinter.E)
    def create_max_size_picker(self):
        self.max_size_picker = Tkinter.Spinbox(self,from_=0,to=100000)
        self.max_size_picker["value"] = 100000
        self.max_size_picker.grid(row=3,column=2,sticky=Tkinter.W)
    def create_max_excel_size_text(self):
        self.max_excel_size_text = Tkinter.Label(self)
        self.max_excel_size_text["text"] = "Maximum Excel Pages:"
        self.max_excel_size_text.grid(row=4,column=1,sticky=Tkinter.E)
    def create_max_excel_size_picker(self):
        self.max_excel_size_picker = Tkinter.Spinbox(self,from_=0,to=100000)
        self.max_excel_size_picker["value"] = 100000
        self.max_excel_size_picker.grid(row=4,column=2,sticky=Tkinter.W)
    def create_flatten_checkbox(self):
        self.flatten_checkbox = Tkinter.Checkbutton(self)
        self.flatten_checkbox["text"] = "Flatten directory"
        self.flatten_checkbox_value = Tkinter.IntVar()
        self.flatten_checkbox["variable"] = self.flatten_checkbox_value
        self.flatten_checkbox.grid(row=7,column=2,sticky=Tkinter.W)
    def create_progress_text(self):
        self.progress_text = Tkinter.Label(self)
        self.progress_text["text"] = "0%"
        self.progress_text.grid(row=10,column=0,columnspan=4)
    def create_progress_text2(self):
        self.progress_text2 = Tkinter.Label(self)
        self.progress_text2["text"] = "Ready!"
        self.progress_text2.grid(row=11,column=0,columnspan=4)
    def create_progress_text3(self):
        self.progress_text3 = Tkinter.Label(self)
        self.progress_text3["text"] = ""
        self.progress_text3.grid(row=12,column=0,columnspan=4)
    def create_progress_bar(self):
        self.progress_bar = ttk.Progressbar(self)
        self.progress_bar["length"] = 500
        self.progress_bar["value"] = 0
        self.progress_bar.grid(row=9,column=0,columnspan=4,pady=4)
    def create_email_formatting_checkbox(self):
        self.email_formatting_checkbox = Tkinter.Checkbutton(self)
        self.email_formatting_checkbox["text"] = "Format emails (BETA)"
        self.email_formatting_checkbox_value = Tkinter.IntVar()
        self.email_formatting_checkbox["variable"] = self.email_formatting_checkbox_value
        self.email_formatting_checkbox.grid(row=7,column=1,sticky=Tkinter.E)
    def create_numbering_checkbox(self):
        self.numbering_checkbox = Tkinter.Checkbutton(self)
        self.numbering_checkbox["text"] = "Number files"
        self.numbering_checkbox_value = Tkinter.IntVar()
        self.numbering_checkbox["variable"] = self.numbering_checkbox_value
        self.numbering_checkbox.grid(row=6,column=2,sticky=Tkinter.W)
    def create_slipsheets_checkbox(self):
        self.slipsheets_checkbox = Tkinter.Checkbutton(self)
        self.slipsheets_checkbox["text"] = "Slipsheets"
        self.slipsheets_checkbox_value = Tkinter.IntVar()
        self.slipsheets_checkbox["variable"] = self.slipsheets_checkbox_value
        self.slipsheets_checkbox.grid(row=6,column=1,sticky=Tkinter.E)
    def create_page_setup_button(self):
        self.page_setup = Tkinter.Button(self)
        self.page_setup["text"] = "Excel Page Setup"
        self.page_setup["command"] = self.excel_page_setup
        self.page_setup.grid(row=5,column=1,sticky=Tkinter.E)
    def create_exit(self):
        self.exit = Tkinter.Button(self)
        self.exit["text"] = "Exit"
        self.exit["command"] = self.quit
        self.exit["background"] = "#FF0000"
        self.exit["height"] = 3
        self.exit["width"] = 6
        self.exit.grid(row=8,column=2,sticky=Tkinter.W)
    def create_start(self):
        self.start_button = Tkinter.Button(self)
        self.start_button["text"] = "Start"
        self.start_button["command"] = self.start
        self.start_button["background"]="#00FF00"
        self.start_button["height"] = 3
        self.start_button["width"] = 6
        self.start_button.grid(row=8,column=1,sticky=Tkinter.E)
    def create_choose_source(self):
        self.choose_source = Tkinter.Button(self)
        self.choose_source["text"] = "Source Directory:"
        self.choose_source["command"] = self.source_directory_select
        self.choose_source.grid(row=0,column=1,sticky=Tkinter.E)
    def create_choose_dest(self):
        self.choose_dest = Tkinter.Button(self)
        self.choose_dest["text"] = "Destination Directory:"
        self.choose_dest["command"] = self.dest_directory_select
        self.choose_dest.grid(row=1,column=1,sticky=Tkinter.E)
    def create_source_text(self):
        self.chosen_source = Tkinter.Label(self)
        self.chosen_source["text"] = self.source_directory
        self.chosen_source.grid(row=0,column=2,sticky=Tkinter.W)
    def create_dest_text(self):
        self.chosen_dest = Tkinter.Label(self)
        self.chosen_dest["text"] = self.dest_directory
        self.chosen_dest.grid(row=1,column=2,sticky=Tkinter.W)
    def create_tab_depth_text(self):
        self.tab_depth_text = Tkinter.Label(self)
        self.tab_depth_text["text"] = "Tab Depth:"
        self.tab_depth_text.grid(row=2,column=1,sticky=Tkinter.E)
    def create_tab_depth_picker(self):
        self.tab_depth_picker = Tkinter.Spinbox(self,from_=0,to=10000)
        self.tab_depth_picker.grid(row=2,column=2,sticky=Tkinter.W)
    def source_directory_select(self):
        self.source_directory = tkFileDialog.askdirectory(initialdir = "C:/", title = "Choose Source Directory", mustexist=True)
        self.chosen_source["text"] = self.source_directory
    def dest_directory_select(self):
        self.dest_directory = tkFileDialog.askdirectory(initialdir = "C:/", title = "Choose Destination Directory", mustexist=True)
        self.chosen_dest["text"] = self.dest_directory
    def start(self):
        if self.source_directory==None or self.dest_directory==None or self.tab_depth_picker.get()==None or self.source_directory=="" or self.dest_directory=="" or self.tab_depth_picker.get()=="":
            tkMessageBox.showerror("Error","Missing fields - cannot launch.")
        elif self.source_directory=="C:\\" or self.source_directory=="C:/" or self.source_directory=="C:" or self.dest_directory=="C:\\" or self.dest_directory=="C:/" or self.dest_directory=="C:":
            tkMessageBox.showerror("Error","Will not launch using the root directory.")
        elif os.listdir(self.dest_directory) != []:
            tkMessageBox.showerror("Error","Destination directory must be empty.")
        elif self.source_directory==self.dest_directory:
            tkMessageBox.showerror("Error","Source and destination directories cannot be the same.")
        else:
            launch_main(self.source_directory,self.dest_directory,int(self.tab_depth_picker.get()),self.page_setup_settings,self.email_formatting_checkbox_value.get(),self.numbering_checkbox_value.get(),self.slipsheets_checkbox_value.get(),self.flatten_checkbox_value.get(),int(self.max_size_picker.get()),int(self.max_excel_size_picker.get()))
    def excel_page_setup(self):
        excel = comtypes.client.CreateObject('Excel.Application')
        excel.Visible = False
        xls = excel.Workbooks.Add()
        for i in xls.Sheets:
            i.Select(False)
        excel.CommandBars("File").Controls("Page Set&up...").Execute()
        self.page_setup_settings = xls.Sheets(1).PageSetup
        # xls.Close(False)
        # excel.Quit()
def set_worksheet_page_setup_settings(source,dest):
    dest.AlignMarginsHeaderFooter = source.AlignMarginsHeaderFooter
    dest.BlackAndWhite = source.BlackAndWhite
    dest.BottomMargin = source.BottomMargin
    dest.CenterFooter = source.CenterFooter
    dest.CenterHeader = source.CenterHeader
    dest.CenterHorizontally = source.CenterHorizontally
    dest.CenterVertically = source.CenterVertically
    dest.DifferentFirstPageHeaderFooter = source.DifferentFirstPageHeaderFooter
    dest.Draft = source.Draft
    dest.FirstPageNumber = source.FirstPageNumber
    dest.FitToPagesTall = source.FitToPagesTall
    dest.FitToPagesWide = source.FitToPagesWide
    dest.FooterMargin = source.FooterMargin
    dest.HeaderMargin = source.HeaderMargin
    dest.LeftFooter = source.LeftFooter
    dest.LeftHeader = source.LeftHeader
    dest.LeftMargin = source.LeftMargin
    dest.OddAndEvenPagesHeaderFooter = source.OddAndEvenPagesHeaderFooter
    dest.Order = source.Order
    dest.Orientation = source.Orientation
    dest.PaperSize = source.PaperSize
    dest.PrintArea = source.PrintArea
    dest.PrintComments = source.PrintComments
    dest.PrintErrors = source.PrintErrors
    dest.PrintGridlines = source.PrintGridlines
    dest.PrintHeadings = source.PrintHeadings
    dest.PrintNotes = source.PrintNotes
    dest.PrintTitleColumns = source.PrintTitleColumns
    dest.PrintTitleRows = source.PrintTitleRows
    dest.RightFooter = source.RightFooter
    dest.RightHeader = source.RightHeader
    dest.RightMargin = source.RightMargin
    dest.ScaleWithDocHeaderFooter = source.ScaleWithDocHeaderFooter
    dest.TopMargin = source.TopMargin
    dest.Zoom = source.Zoom
    
if __name__ == "__main__":
    from docopt import docopt
    arguments = docopt(__doc__, version='RRR 0.1')
    sourcedir = arguments["<sourcedir>"]
    if sourcedir == "C:\\" or sourcedir == "C:" or sourcedir=="C:/": #this would be bad
        print("Will not run on root directory. Exiting.")
        sys.exit()
    destdir = arguments["<destdir>"]
    if destdir == "C:\\" or destdir == "C:" or destdir=="C:/": #this would be bad
        print("Will not run on root directory. Exiting.")
        sys.exit()
    if destdir != None and os.listdir(destdir) != []:
        print("Destination directory must be empty. Exiting.")
        sys.exit()
    tabdepth = arguments["<tabdepth>"]
    if tabdepth !=None:
        tabdepth = int(tabdepth)
    if sourcedir==None or destdir==None or tabdepth==None:
        launch_gui()
        sys.exit()
    if sourcedir==destdir:
        print("Source and destination directories are the same. Exiting.")
        sys.exit()
    launch_main(sourcedir,destdir,tabdepth,None)
