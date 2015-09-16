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
import os,io,zipfile,sys,comtypes.client,email,mimetypes,olefile,shutil,time,StringIO
import logging,Tkinter,tkFileDialog,tkMessageBox,copy,sys
from natsort import natsorted
from PyPDF2 import PdfFileReader,PdfFileWriter
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from operator import itemgetter

def main (rootdir,tabdepth,page_setup_settings,pdf_reprocess_status):
    reload(sys)
    sys.setdefaultencoding('utf8')
    logging.basicConfig(filename='rrrlog.txt',level=logging.INFO)
    console = logging.StreamHandler()
    console.setLevel(logging.DEBUG)
    logging.getLogger('').addHandler(console)
    unzip(rootdir)
    add_directory_slipsheets(rootdir,tabdepth)
    convert_to_pdf(rootdir,page_setup_settings,pdf_reprocess_status)
    rename_resize_rotate(rootdir)
    logging.info("FINISHED!")
def unzip (rootdir):
    while (zip_found(rootdir)):
        process_zips(rootdir)
def add_directory_slipsheets (rootdir,tabdepth):
    tabdepth+=rootdir.count(os.path.sep)
    for subdir,dirs,file in os.walk(rootdir):
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
def rename_resize_rotate(rootdir):
    n=0
    for subdir,dirs,files in os.walk(rootdir):
        files = customsorted(files)
        for file in files:
            n+=1
            logging.info ("RRRing {0}".format(os.path.join(subdir,"{0:05d} ".format(n) + file)))
            os.rename(os.path.join(subdir,file),os.path.join(subdir,"{0:05d} ".format(n) + file))
            if file[-4:].upper()==".PDF":
                process_pdf(os.path.join(subdir,"{0:05d} ".format(n) + file),rootdir)
def customsorted(files):
    indices = []
    for file in files:
        index = []
        index.append(file)
        split = file.split(".")
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
def convert_to_pdf(rootdir,page_setup_settings,pdf_reprocess_status):
    for subdir,dirs,files in os.walk(rootdir):
        for file in files:
            try:
                logging.info("Converting {0}".format(os.path.join(subdir,file)))
            except:
                logging.info("Converting <NAME CANNOT BE DISPLAYED>.")
            if file[-4:].upper()==".DOC" or file[-5:].upper()==".DOCX" or file[-4:].upper()==".TXT" or file[-4:].upper()==".RTF":
                process_doc(rootdir,os.path.join(subdir,file))
            elif file[-4:].upper()==".XLS" or file[-5:].upper()==".XLSX":
                process_xls(rootdir,os.path.join(subdir,file),page_setup_settings)
            elif file[-4:].upper()==".TIF" or file[-4:].upper()==".JPG" or file[-5:].upper()==".TIFF" or file[-5:].upper()==".JPEG" or file[-4:].upper()==".PNG":
                process_image(os.path.join(subdir,file))
            elif file[-4:].upper()==".PPT" or file[-5:].upper()==".PPTX":
                process_ppt(os.path.join(subdir,file))
            # elif file[-4:].upper()==".PDF" and pdf_reprocess_status==1:
                # reprocess_pdf(os.path.join(subdir,file))
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
def process_pdf(filename,rootdir):
        pdf = PdfFileReader(filename,strict=False)
        if pdf.isEncrypted:
            logging.warning("ERROR: File encrypted - check PDF")
            shutil.copy(filename,rootdir)
            return
        pdf_dest = PdfFileWriter()
        # add_slipsheet(pdf_dest,filename)
        process_pdf_pages(pdf,pdf_dest)
        pdf_write(pdf_dest,filename,rootdir)
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
        logging.warning("ERROR: Could not scale - check PDF")
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
    for subdirs,dirs,files in os.walk(rootdir):
        for file in files:
            if file[-4:].upper()==".ZIP" or file[-4:].upper()==".MSG":
                return_value = True
    return return_value
def process_zips(rootdir):
    for subdir,dirs,files in os.walk(rootdir):
        for file in files:
            if file[-4:].upper()==".ZIP":
                process_zip(subdir,file)
            elif file[-4:].upper()==".MSG":
                process_msg(subdir,file)
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
        process_mime_msg_section(part,subdir,file)
    os.remove(os.path.join(subdir,file))
def process_mime_msg_section(part,subdir,file):
    if part.get_content_maintype() == 'multipart':
        return
    filename = part.get_filename()
    if filename==None:
        filename = generate_mime_msg_section_filename(part)
    counter += 1
    fp = open(os.path.join(os.path.join(subdir,file)+".dir",filename),'wb')
    fp.write(part.get_payload(decode=True))
    fp.close()
def generate_mime_msg_section_filename(part):
    ext = mimetypes.guess_extension(part.get_content_type())
    if not ext:
        ext = ".bin"
    filename = "part-{0:03d}{1}".format(counter,ext)
    return filename
def process_msg(subdir,file):
    if olefile.isOleFile(os.path.join(subdir,file))==False:
        process_mime_msg(subdir,file)
        return
    os.mkdir(os.path.join(subdir,file)+".dir")
    ole = olefile.OleFileIO(os.path.join(subdir,file))
    attach_list = get_msg_attach_list(ole)
    extract_msg_files(attach_list,ole,subdir,file)
    extract_msg_message(ole,subdir,file)
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
def extract_msg_message(ole,subdir,file):
    msg_from,msg_to,msg_cc,msg_subject,msg_header,msg_body = extract_msg_message_data(ole)    
    fp = io.open(os.path.join(os.path.join(subdir,file)+".dir","000 {0}.txt".format(file)),"w")
    try:
        fp.write("From: {0}\nTo: {1}\nCC: {2}\nSubject: {3}\nHeader: {4}\n".format(msg_from,msg_to,msg_cc,msg_subject,msg_header).decode('utf-8'))
    except:
        fp.write("From: {0}\nTo: {1}\nCC: {2}\nSubject: {3}\nHeader: {4}\n".format(msg_from,msg_to,msg_cc,msg_subject,msg_header).decode('ISO-8859-1'))
    fp.write(unicode("---------------\n\n"))
    try:
        fp.write(msg_body.decode('utf-8'))
    except:
        fp.write(msg_body.decode('ISO-8859-1'))
    fp.close()
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
        elif i[0][:16]=="__substg1.0_1000":
            msg_body = extract_msg_stream_text(i,ole)
    return (msg_from,msg_to,msg_cc,msg_subject,msg_header,msg_body)
def extract_msg_stream_text(index,ole):
    stream = ole.openstream(index[0])
    text = stream.read()
    text = clean_string(text)
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
    word = comtypes.client.CreateObject('Word.Application')
    word.Visible = False
    doc = word.Documents.Open(filename)
    doc.SaveAs(filename+".pdf",FileFormat=17)
    doc.Close()
    word.Quit()
    os.remove(filename)
def process_xls(rootdir,filename,page_setup_settings):
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
    excel.Quit()
    os.remove(filename)
def process_ppt(filename):
    powerpoint = comtypes.client.CreateObject('PowerPoint.Application')
    ppt = powerpoint.Presentations.Open(filename)
    ppt.ExportAsFixedFormat(filename+".pdf",FixedFormatType=2)
    ppt.Close()
    powerpoint.Quit()
    os.remove(filename)
def process_image(filename):
    #image = PythonMagick.Image()
    #image.read(filename)
    #image.write(filename+".pdf")
    #ABOVE CODE IS UNTESTED BUT WILL PROBABLY DO THE TRICK INSTEAD OF THE FOLLOWING IF YOU DON'T HAVE ACROBAT PRO
    acrobat = comtypes.client.CreateObject('AcroExch.App')
    acrobat.Hide()
    image = comtypes.client.CreateObject('AcroExch.AVDoc')
    image.Open(filename,'temp')
    image2 = image.GetPDDoc()
    image2.Save(1,filename+".pdf")
    acrobat.CloseAllDocs()
    acrobat.Exit()
    os.remove(filename)
def add_slipsheet(pdf_dest,text):
    slipsheet = pdf_dest.addBlankPage(8.5*72,11*72)
    packet = StringIO.StringIO()
    can = canvas.Canvas(packet, pagesize=letter)
    directory_path,base_filename = os.path.split(text)
    can.drawCentredString(8.5*72/2,600,base_filename)
    can.save()
    packet.seek(0)
    slipsheet_overlay_pdf=PdfFileReader(packet)
    slipsheet_overlay_page=slipsheet_overlay_pdf.getPage(0)
    slipsheet.mergePage(slipsheet_overlay_page)
def process_pdf_pages(pdf,pdf_dest):
    numPages = pdf.getNumPages()        
    for i in range(numPages):
        page = pdf.getPage(i)
        process_pdf_page(page)
        pdf_dest.addPage(page)
def pdf_write(pdf_dest,filename,rootdir):
    try:
        pdf_dest.write(io.open(filename,mode='w+b'))
    except:
        logging.warning ("ERROR: Could not write PDF - check PDF")
        shutil.copy(filename,rootdir)
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
def launch_main(sourcedir,destdir,tabdepth,page_setup_settings,pdf_reprocess_status):
    if sourcedir==None or destdir==None or tabdepth==None:
        print("One or more missing arguments. Exiting.")
        sys.exit()
    sourcedir = os.path.abspath(sourcedir)
    destdir = os.path.abspath(destdir)
    os.rmdir(destdir)
    shutil.copytree(sourcedir,destdir)
    main(destdir,tabdepth,page_setup_settings,pdf_reprocess_status)
def launch_gui():
    root = Tkinter.Tk()
    root.title("RRR")
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
        # self.create_pdf_reprocess_checkbox()
    def create_pdf_reprocess_checkbox(self):
        self.pdf_reprocess_checkbox = Tkinter.Checkbutton(self)
        self.pdf_reprocess_checkbox["text"] = "Run all PDFs through Adobe before processing"
        self.pdf_reprocess_checkbox_value = Tkinter.IntVar()
        self.pdf_reprocess_checkbox["variable"] = self.pdf_reprocess_checkbox_value
        self.pdf_reprocess_checkbox.grid(row=3,column=1)
    def create_page_setup_button(self):
        self.page_setup = Tkinter.Button(self)
        self.page_setup["text"] = "Excel Page Setup"
        self.page_setup["command"] = self.excel_page_setup
        self.page_setup.grid(row=3,column=0)
    def create_exit(self):
        self.exit = Tkinter.Button(self)
        self.exit["text"] = "Exit"
        self.exit["command"] = self.quit
        self.exit.grid(row=5,column=1)
    def create_start(self):
        self.start_button = Tkinter.Button(self)
        self.start_button["text"] = "Start"
        self.start_button["command"] = self.start
        self.start_button.grid(row=5,column=0)
    def create_choose_source(self):
        self.choose_source = Tkinter.Button(self)
        self.choose_source["text"] = "Source Directory:"
        self.choose_source["command"] = self.source_directory_select
        self.choose_source.grid(row=0,column=0)
    def create_choose_dest(self):
        self.choose_dest = Tkinter.Button(self)
        self.choose_dest["text"] = "Destination Directory:"
        self.choose_dest["command"] = self.dest_directory_select
        self.choose_dest.grid(row=1,column=0)
    def create_source_text(self):
        self.chosen_source = Tkinter.Label(self)
        self.chosen_source["text"] = self.source_directory
        self.chosen_source.grid(row=0,column=1)
    def create_dest_text(self):
        self.chosen_dest = Tkinter.Label(self)
        self.chosen_dest["text"] = self.dest_directory
        self.chosen_dest.grid(row=1,column=1)
    def create_tab_depth_text(self):
        self.tab_depth_text = Tkinter.Label(self)
        self.tab_depth_text["text"] = "Tab Depth:"
        self.tab_depth_text.grid(row=2,column=0)
    def create_tab_depth_picker(self):
        self.tab_depth_picker = Tkinter.Spinbox(self,from_=0,to=10000)
        self.tab_depth_picker.grid(row=2,column=1)
    def source_directory_select(self):
        self.source_directory = tkFileDialog.askdirectory(initialdir = os.getcwd(), title = "Choose Source Directory", mustexist=True)
        self.chosen_source["text"] = self.source_directory
    def dest_directory_select(self):
        self.dest_directory = tkFileDialog.askdirectory(initialdir = os.getcwd(), title = "Choose Destination Directory", mustexist=True)
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
            launch_main(self.source_directory,self.dest_directory,int(self.tab_depth_picker.get()),self.page_setup_settings,self.pdf_reprocess_checkbox_value.get())
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
