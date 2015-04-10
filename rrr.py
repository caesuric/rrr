"""Rename Resize Rotate

Renames a series of files in a directory, resizing to letter size and
rotating any PDFs found within to portrait orientation.
* Expands all zip files and email attachments found within target directories.
* Converts all MS Word, MS Excel, TIF, JPG, PNG, and TXT files to PDFs before
RRRing.
* Adds tab placeholders down to a specified directory depth for ease of layout
during later printing.
* Copies files that are unmodifiable to the root directory for case-by-case
processing by end user.

* Requires MS Word or MS Excel 2010 or newer installed to convert Word and
Excel files.
* Requires Adobe Acrobat Pro to convert images *OR* see comments in process_image
for an (untested) ImageMagick implementation.

Usage:
    rrr [<rootdir> <tabdepth>]
    rrr (-h | --help)

Options:
    -h --help               Show this screen
"""
import os,io,zipfile,sys,comtypes.client,email,mimetypes,olefile,shutil,time,StringIO
import logging
from natsort import natsorted
from PyPDF2 import PdfFileReader,PdfFileWriter
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter

def main (rootdir,tabdepth):
    logging.basicConfig(filename='rrrlog.txt',level=logging.DEBUG)
    console = logging.StreamHandler()
    console.setLevel(logging.DEBUG)
    logging.getLogger('').addHandler(console)
    unzip(rootdir)
    add_directory_slipsheets(rootdir,tabdepth)
    convert_to_pdf(rootdir)
    rename_resize_rotate(rootdir)
def unzip (rootdir):
    while (zip_found(rootdir)):
        process_zips(rootdir)
def add_directory_slipsheets (rootdir,tabdepth):
    tabdepth+=rootdir.count(os.path.sep)
    for subdir,dirs,file in os.walk(rootdir):
        if subdir.count(os.path.sep)<=tabdepth:
            shutil.copy("blankpage.pdf",os.path.join(subdir,"000 ----- INSERT TAB HERE.pdf"))
def rename_resize_rotate(rootdir):
    n=0
    for subdir,dirs,files in os.walk(rootdir):
        files = natsorted(files)
        for file in files:
            n+=1
            logging.info ("RRRing {0}".format(os.path.join(subdir,"{0:05d} ".format(n) + file)))
            os.rename(os.path.join(subdir,file),os.path.join(subdir,"{0:05d} ".format(n) + file))
            if file[-4:].upper()==".PDF" and file[-29:].upper()!="000 ----- INSERT TAB HERE.PDF":
                process_pdf(os.path.join(subdir,"{0:05d} ".format(n) + file),rootdir)
def convert_to_pdf(rootdir):
    for subdir,dirs,files in os.walk(rootdir):
        for file in files:
            logging.info("Converting {0}".format(os.path.join(subdir,file)))
            if file[-4:].upper()==".DOC" or file[-5:].upper()==".DOCX" or file[-4:].upper()==".TXT":
                process_doc(os.path.join(subdir,file))
            elif file[-4:].upper()==".XLS" or file[-5:].upper()==".XLSX":
                process_xls(os.path.join(subdir,file))
            elif file[-4:].upper()==".TIF" or file[-4:].upper()==".JPG" or file[-5:].upper()==".TIFF" or file[-5:].upper()==".JPEG" or file[-4:].upper()==".PNG":
                process_image(os.path.join(subdir,file))
def process_pdf(filename,rootdir):
        pdf = PdfFileReader(filename,strict=False)
        if pdf.isEncrypted:
            logging.warning("ERROR: File encrypted - check PDF")
            shutil.copy(filename,rootdir)
            return
        pdf_dest = PdfFileWriter()
        add_slipsheet(pdf_dest,filename)
        process_pdf_pages(pdf,pdf_dest)
        pdf_write(pdf_dest,filename,rootdir)
def process_pdf_page(page):
    x,y = get_page_dimensions(page)
    if x>y:
        page.rotateCounterClockwise(90)
        x,y = get_page_dimensions(page)
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
        write_msg_attachment(index,ole,subdir,file,filename)
def get_msg_attachment_filename(index,ole):
    filename = get_msg_attachment_filename_primary(index,ole)
    if filename == None:
        filename = get_msg_attachment_filename_fallback(index,ole)
    return filename
def get_msg_attachment_filename_primary(index,ole):
    for i in ole.listdir():
        if i[0]==index:
            if i[1][:16]=="__substg1.0_3707":
                name_stream = ole.openstream("{0}/{1}".format(i[0],i[1]))
                filename = name_stream.read()
    return filename
def get_msg_attachment_filename_fallback(index,ole):
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
                file_stream = ole.openstream("{0}/{1}".format(j[0],j[1]))
                file_data = file_stream.read()
                fp = io.open(os.path.join(os.path.join(subdir,file)+".dir",filename),"w+b")
                fp.write(file_data)
                fp.close()
def extract_msg_message(ole,subdir,file):
    msg_from,msg_to,msg_cc,msg_subject,msg_header,msg_body = extract_msg_message_data(ole)    
    fp = io.open(os.path.join(os.path.join(subdir,file)+".dir","000 message.txt"),"w")
    fp.write(unicode("From: {0}\nTo: {1}\nCC: {2}\nSubject: {3}\nHeader: {4}\n".format(msg_from,msg_to,msg_cc,msg_subject,msg_header)))
    fp.write(unicode("---------------\n\n"))
    fp.write(unicode(msg_body))
    fp.close()
def extract_msg_message_data(ole):
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
def process_doc(filename):
    word = comtypes.client.CreateObject('Word.Application')
    word.Visible = False
    doc = word.Documents.Open(filename)
    doc.SaveAs(filename+".pdf",FileFormat=17)
    doc.Close()
    word.Quit()
    os.remove(filename)
def process_xls(filename):
    excel = comtypes.client.CreateObject('Excel.Application')
    excel.Visible = False
    xls = excel.Workbooks.Open(filename)
    for i in xls.Sheets:
        i.Select(False)
    xls.SaveAs(filename+".pdf",FileFormat=57)
    xls.Close()
    excel.Quit()
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
def add_slipsheet(pdf_dest,filename):
    slipsheet = pdf_dest.addBlankPage(8.5*72,11*72)
    packet = StringIO.StringIO()
    can = canvas.Canvas(packet, pagesize=letter)
    directory_path,base_filename = os.path.split(filename)
    can.drawString(10,600,base_filename)
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
def get_page_dimensions(page):
    a,b,c,d = page.mediaBox
    x = abs(c-a)
    y = abs(d-b)
    previous_rotation = page.get('/Rotate')
    if previous_rotation == 90 or previous_rotation == -90 or previous_rotation == 270 or previous_rotation == -270:
        temp=x
        x=y
        y=temp
    return (x,y)
    
if __name__ == "__main__":
    from docopt import docopt
    arguments = docopt(__doc__, version='RRR 0.1')
    rootdir = arguments["<rootdir>"]
    if rootdir == "C:\\" or rootdir == "C:": #this would be bad
        sys.exit()
    tabdepth = int(arguments["<tabdepth>"])
    main(rootdir,tabdepth)
