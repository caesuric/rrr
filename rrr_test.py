import rrr,unittest,os,shutil,zipfile
from PyPDF2 import PdfFileReader,PdfFileWriter
#all tests use test_ as shorthand for tests_that_it_
class TestMain(unittest.TestCase):
    pass
    #STUB
class TestUnzip(unittest.TestCase):
    def setUp(self):
        self.a = "test_process_zips_test_directory"
        os.mkdir(self.a)
        with zipfile.ZipFile(os.path.join(self.a,"test_zip.zip"),'w') as zip:
            zip.write("rrr.py")
            zip.write("blankpage.pdf")
    def tearDown(self):
        os.rmdir(self.a)
    def test_unzips_a_single_zip_file(self):
        rrr.unzip(self.a,1)
        self.assertTrue(os.path.isfile(os.path.join(self.a,"test_zip.zip.dir","rrr.py")))
        self.assertTrue(os.path.isfile(os.path.join(self.a,"test_zip.zip.dir","blankpage.pdf")))
        os.remove(os.path.join(self.a,"test_zip.zip.dir","rrr.py"))
        os.remove(os.path.join(self.a,"test_zip.zip.dir","blankpage.pdf"))
        os.rmdir(os.path.join(self.a,"test_zip.zip.dir"))
    def test_unzips_multiple_zip_files(self):
        shutil.copy(os.path.join(self.a,"test_zip.zip"),os.path.join(self.a,"test_zip2.zip"))
        rrr.unzip(self.a,1)
        self.assertTrue(os.path.isfile(os.path.join(self.a,"test_zip.zip.dir","rrr.py")))
        self.assertTrue(os.path.isfile(os.path.join(self.a,"test_zip.zip.dir","blankpage.pdf")))
        self.assertTrue(os.path.isfile(os.path.join(self.a,"test_zip2.zip.dir","rrr.py")))
        self.assertTrue(os.path.isfile(os.path.join(self.a,"test_zip2.zip.dir","blankpage.pdf")))
        os.remove(os.path.join(self.a,"test_zip2.zip.dir","rrr.py"))
        os.remove(os.path.join(self.a,"test_zip2.zip.dir","blankpage.pdf"))
        os.rmdir(os.path.join(self.a,"test_zip2.zip.dir"))
        os.remove(os.path.join(self.a,"test_zip.zip.dir","rrr.py"))
        os.remove(os.path.join(self.a,"test_zip.zip.dir","blankpage.pdf"))
        os.rmdir(os.path.join(self.a,"test_zip.zip.dir"))
    def test_unzips_a_nested_zip_file(self):
        with zipfile.ZipFile(os.path.join(self.a,"test_zip2.zip"),'w') as zip:
            zip.write(os.path.join(self.a,"test_zip.zip"))
        os.remove(os.path.join(self.a,"test_zip.zip"))
        rrr.unzip(self.a,1)
        self.assertFalse(os.path.isfile(os.path.join(self.a,"test_zip2.zip.dir","test_zip.zip")))
        self.assertTrue(os.path.isfile(os.path.join(self.a,"test_zip2.zip.dir",self.a,"test_zip.zip.dir","rrr.py")))
        self.assertTrue(os.path.isfile(os.path.join(self.a,"test_zip2.zip.dir",self.a,"test_zip.zip.dir","blankpage.pdf")))
        os.remove(os.path.join(self.a,"test_zip2.zip.dir",self.a,"test_zip.zip.dir","rrr.py"))
        os.remove(os.path.join(self.a,"test_zip2.zip.dir",self.a,"test_zip.zip.dir","blankpage.pdf"))
        os.rmdir(os.path.join(self.a,"test_zip2.zip.dir",self.a,"test_zip.zip.dir"))
        os.rmdir(os.path.join(self.a,"test_zip2.zip.dir",self.a))
        os.rmdir(os.path.join(self.a,"test_zip2.zip.dir"))
class TestAddDirectorySlipsheets(unittest.TestCase):
    def test_creates_a_slipsheet(self):
        a = "test_add_directory_slipsheets_test_directory"
        os.mkdir(a)
        rrr.add_directory_slipsheets(a,1)
        self.assertTrue(os.path.isfile(os.path.join(a,"000 TAB ----- "+a+".pdf")))
        os.remove(os.path.join(a,"000 TAB ----- "+a+".pdf"))
        os.rmdir(a)
    def test_creates_a_slipsheet_at_multiple_levels(self):
        a = "test_add_directory_slipsheets_test_directory"
        os.mkdir(a)
        os.mkdir(os.path.join(a,a))
        os.mkdir(os.path.join(a,a,a))
        rrr.add_directory_slipsheets(a,3)
        self.assertTrue(os.path.isfile(os.path.join(a,"000 TAB ----- "+a+".pdf")))
        self.assertTrue(os.path.isfile(os.path.join(a,a,"000 TAB ----- "+a+".pdf")))
        self.assertTrue(os.path.isfile(os.path.join(a,a,a,"000 TAB ----- "+a+".pdf")))
        os.remove(os.path.join(a,"000 TAB ----- "+a+".pdf"))
        os.remove(os.path.join(a,a,"000 TAB ----- "+a+".pdf"))
        os.remove(os.path.join(a,a,a,"000 TAB ----- "+a+".pdf"))
        os.rmdir(os.path.join(a,a,a))
        os.rmdir(os.path.join(a,a))
        os.rmdir(a)
class TestRenameResizeRotate(unittest.TestCase):
    pass
class TestConvertToPdf(unittest.TestCase):
    pass
    #STUB
class TestProcessPdf(unittest.TestCase):
    pass
class TestProcessPdfPage(unittest.TestCase):
    pass
class TestDetermineScalingFactors(unittest.TestCase):
    def test_leaves_letter_alone(self):
        x_scaling,y_scaling = rrr.determine_scaling_factors(8.5*72,11*72)
        self.assertEqual(x_scaling,1)
        self.assertEqual(y_scaling,1)
    def test_reduces_legal_size_to_letter(self):
        x_scaling,y_scaling = rrr.determine_scaling_factors(8.5*72,14*72)
        self.assertEqual(x_scaling,1)
        self.assertEqual(y_scaling,float(11)/float(14))
    def test_reduces_11_by_16_to_letter(self):
        x_scaling,y_scaling = rrr.determine_scaling_factors(11*72,16*72)
        self.assertEqual(x_scaling,8.5/float(11))
        self.assertEqual(y_scaling,float(11)/float(16))
    def test_ignores_dimensions_smaller_than_letter(self):
        x_scaling,y_scaling = rrr.determine_scaling_factors(5,5)
        self.assertEqual(x_scaling,1)
        self.assertEqual(y_scaling,1)
class TestScaleFactor(unittest.TestCase):
    def test_leaves_letter_alone(self):
        scaling = rrr.scale_factor(8.5*72,11*72)
        self.assertEqual(scaling,1)
    def test_reduces_legal_size_to_letter(self):
        scaling = rrr.scale_factor(8.5*72,14*72)
        self.assertEqual(scaling,float(11)/float(14))
    def test_reduces_11_by_16_to_letter(self):
        scaling = rrr.scale_factor(11*72,16*72)
        self.assertEqual(scaling,float(11)/float(16))
    def test_ignores_dimensions_smaller_than_letter(self):
        scaling = rrr.scale_factor(5,5)
        self.assertEqual(scaling,1)
class TestScaleToLetter(unittest.TestCase):
    pass
class TestGetPageDimensions(unittest.TestCase):
    def test_reads_a_letter_size_pdf_page_properly(self):
        pdf = PdfFileReader("blankpage.pdf",strict=False)
        page = pdf.getPage(0)
        a,b,c,d,x,y = rrr.get_page_dimensions(page)
        self.assertEqual(x,8.5*72)
        self.assertEqual(y,11*72)
class TestGetTargetDimensions(unittest.TestCase):
    def setUp(self):
        pdf = PdfFileReader("blankpage.pdf",strict=False)
        self.page = pdf.getPage(0)
    def test_gets_correct_dimensions_from_a_non_rotated_page(self):
        x_target,y_target = rrr.get_target_dimensions(self.page)
        self.assertEqual(x_target,8.5*72)
        self.assertEqual(y_target,11*72)
    def test_gets_correct_dimensions_from_a_rotated_page(self):
        self.page.rotateClockwise(90)
        x_target,y_target = rrr.get_target_dimensions(self.page)
        self.assertEqual(x_target,11*72)
        self.assertEqual(y_target,8.5*72)
class TestAdjustPageDimensions(unittest.TestCase):
    def setUp(self):
        self.x_target = 100
        self.y_target = 100
    def test_ignores_adjustments_when_moving_to_a_smaller_size(self):
        a = 0
        b = 0
        c = 200
        d = 200
        x = abs(c-a)
        y = abs(d-b)
        a,b,c,d = rrr.adjust_page_dimensions(a,b,c,d,x,y,self.x_target,self.y_target)
        self.assertEqual(a,0)
        self.assertEqual(b,0)
        self.assertEqual(c,200)
        self.assertEqual(d,200)
    def test_adjusts_page_to_larger_size_evenly(self):
        a = 0
        b = 0
        c = 50
        d = 50
        x = abs(c-a)
        y = abs(d-b)
        a,b,c,d = rrr.adjust_page_dimensions(a,b,c,d,x,y,self.x_target,self.y_target)
        self.assertEqual(a,-25)
        self.assertEqual(b,-25)
        self.assertEqual(c,75)
        self.assertEqual(d,75)
class TestUpdatePageDimensions(unittest.TestCase):
    def setUp(self):
        pdf = PdfFileReader("blankpage.pdf",strict=False)
        self.page = pdf.getPage(0)
        self.a = 42
        self.b = 43
        self.c = 80
        self.d = 81
        rrr.update_page_dimensions(self.page,self.a,self.b,self.c,self.d)
    def test_updates_media_box_to_correct_coordinates(self):
        self.assertEqual(self.page.mediaBox.upperLeft,(42,43))
        self.assertEqual(self.page.mediaBox.lowerRight,(80,81))
    def test_updates_crop_box_to_correct_coordinates(self):
        self.assertEqual(self.page.cropBox.upperLeft,(42,43))
        self.assertEqual(self.page.cropBox.lowerRight,(80,81))
class TestZipFound(unittest.TestCase):
    def setUp(self):
        self.a = "test_zip_found_test_directory"
        os.mkdir(self.a)
    def tearDown(self):
        os.rmdir(self.a)
    def test_returns_false_if_no_zip_or_msg_files_are_found(self):
        self.assertFalse(rrr.zip_found(self.a))
    def test_returns_true_if_a_zip_file_is_found(self):
        shutil.copy("blankpage.pdf",os.path.join(self.a,"fakezipfile.zip"))
        self.assertTrue(rrr.zip_found(self.a))
        os.remove(os.path.join(self.a,"fakezipfile.zip"))
    def test_returns_true_if_a_msg_file_is_found(self):
        shutil.copy("blankpage.pdf",os.path.join(self.a,"fakemsgfile.msg"))
        self.assertTrue(rrr.zip_found(self.a))
        os.remove(os.path.join(self.a,"fakemsgfile.msg"))
    def test_returns_true_if_multiple_zip_files_are_found(self):
        shutil.copy("blankpage.pdf",os.path.join(self.a,"fakezipfile.zip"))
        shutil.copy("blankpage.pdf",os.path.join(self.a,"fakezipfile_b.zip"))
        self.assertTrue(rrr.zip_found(self.a))
        os.remove(os.path.join(self.a,"fakezipfile.zip"))
        os.remove(os.path.join(self.a,"fakezipfile_b.zip"))
class TestProcessZips(unittest.TestCase):
    def setUp(self):
        self.a = "test_process_zips_test_directory"
        os.mkdir(self.a)
        with zipfile.ZipFile(os.path.join(self.a,"test_zip.zip"),'w') as zip:
            zip.write("rrr.py")
            zip.write("blankpage.pdf")
    def tearDown(self):
        os.rmdir(self.a)
    def test_unzips_a_single_zip_file(self):
        rrr.process_zips(self.a,1)
        self.assertTrue(os.path.isfile(os.path.join(self.a,"test_zip.zip.dir","rrr.py")))
        self.assertTrue(os.path.isfile(os.path.join(self.a,"test_zip.zip.dir","blankpage.pdf")))
        os.remove(os.path.join(self.a,"test_zip.zip.dir","rrr.py"))
        os.remove(os.path.join(self.a,"test_zip.zip.dir","blankpage.pdf"))
        os.rmdir(os.path.join(self.a,"test_zip.zip.dir"))
    def test_unzips_multiple_zip_files(self):
        shutil.copy(os.path.join(self.a,"test_zip.zip"),os.path.join(self.a,"test_zip2.zip"))
        rrr.process_zips(self.a,1)
        self.assertTrue(os.path.isfile(os.path.join(self.a,"test_zip.zip.dir","rrr.py")))
        self.assertTrue(os.path.isfile(os.path.join(self.a,"test_zip.zip.dir","blankpage.pdf")))
        self.assertTrue(os.path.isfile(os.path.join(self.a,"test_zip2.zip.dir","rrr.py")))
        self.assertTrue(os.path.isfile(os.path.join(self.a,"test_zip2.zip.dir","blankpage.pdf")))
        os.remove(os.path.join(self.a,"test_zip2.zip.dir","rrr.py"))
        os.remove(os.path.join(self.a,"test_zip2.zip.dir","blankpage.pdf"))
        os.rmdir(os.path.join(self.a,"test_zip2.zip.dir"))
        os.remove(os.path.join(self.a,"test_zip.zip.dir","rrr.py"))
        os.remove(os.path.join(self.a,"test_zip.zip.dir","blankpage.pdf"))
        os.rmdir(os.path.join(self.a,"test_zip.zip.dir"))
    #def test_unzips_a_msg_file(self):
    #    pass
    #    STUB
class TestProcessZip(unittest.TestCase):
    def setUp(self):
        self.a = "test_process_zip_test_directory"
        os.mkdir(self.a)
        with zipfile.ZipFile(os.path.join(self.a,"test_zip.zip"),'w') as zip:
            zip.write("rrr.py")
            zip.write("blankpage.pdf")
        rrr.process_zip(self.a,"test_zip.zip")
    def tearDown(self):
        os.remove(os.path.join(self.a,"test_zip.zip.dir","rrr.py"))
        os.remove(os.path.join(self.a,"test_zip.zip.dir","blankpage.pdf"))
        os.rmdir(os.path.join(self.a,"test_zip.zip.dir"))
        os.rmdir(self.a)
    def test_creates_a_directory_to_unzip_files_into(self):
        self.assertTrue(os.path.isdir(os.path.join(self.a,"test_zip.zip.dir")))
    def test_extracts_all_files_from_the_zip(self):
        self.assertTrue(os.path.isfile(os.path.join(self.a,"test_zip.zip.dir","rrr.py")))
        self.assertTrue(os.path.isfile(os.path.join(self.a,"test_zip.zip.dir","blankpage.pdf")))
    def test_deletes_the_zip_file_after_extracting(self):
        self.assertFalse(os.path.isfile(os.path.join(self.a,"test_zip.zip")))
class TestProcessMimeMsg(unittest.TestCase):
    pass
class TestProcessMimeMsgSection(unittest.TestCase):
    pass
class TestGenerateMimeMsgSectionFilename(unittest.TestCase):
    pass
class TestProcessMsg(unittest.TestCase):
    pass
class TestGetMsgAttachList(unittest.TestCase):
    pass
class TestExtractMsgFiles(unittest.TestCase):
    pass
class TestGetMsgAttachmentFilename(unittest.TestCase):
    pass
class TestGetMsgAttachmentFilenamePrimary(unittest.TestCase):
    pass
class TestGetMsgAttachmentFilenameFallback(unittest.TestCase):
    pass
class TestWriteMsgAttachment(unittest.TestCase):
    pass
class TestExtractMsgMessage(unittest.TestCase):
    pass
class TestExtractMsgMessageData(unittest.TestCase):
    pass
class TestExtractMsgStreamText(unittest.TestCase):
    pass
class TestCleanString(unittest.TestCase):
    def test_skips_every_second_character_of_input(self):
        self.assertEqual(rrr.clean_string("A B C D"),"ABCD")
class TestProcessDoc(unittest.TestCase):
    #STUB
    pass
class TestProcessXls(unittest.TestCase):
    #STUB
    pass
class TestProcessImage(unittest.TestCase):
    #STUB
    pass
class TestAddSlipsheet(unittest.TestCase):
    def test_adds_an_additional_page(self):
        pdf = PdfFileReader("blankpage.pdf",strict=False)
        pdf_dest = PdfFileWriter()
        rrr.add_slipsheet(pdf_dest,"blankpage.pdf")
        pdf_dest.addPage(pdf.getPage(0))
        self.assertEqual(pdf_dest.getNumPages(),2)
class TestProcessPdfPages(unittest.TestCase):
    pass
class TestPdfWrite(unittest.TestCase):
    def test_creates_a_pdf(self):
        pdf = PdfFileReader("blankpage.pdf",strict=False)
        pdf_dest = PdfFileWriter()
        pdf_dest.addPage(pdf.getPage(0))
        a = "test_pdf_write_test_pdf.pdf"
        rrr.pdf_write(pdf_dest,a,os.getcwd())
        self.assertTrue(os.path.isfile(a))
        os.remove(a)
class TestGetRotatedPageDimensions(unittest.TestCase):
    def setUp(self):
        pdf = PdfFileReader("blankpage.pdf",strict=False)
        self.page = pdf.getPage(0)
    def test_gets_the_appropriate_dimensions_for_a_letter_page(self):
        x,y = rrr.get_rotated_page_dimensions(self.page)
        self.assertEqual(x,8.5*72)
        self.assertEqual(y,11*72)
    def test_gets_the_appropriate_dimensions_for_a_sideways_letter_page(self):
        self.page.rotateCounterClockwise(90)
        x,y = rrr.get_rotated_page_dimensions(self.page)
        self.assertEqual(x,11*72)
        self.assertEqual(y,8.5*72)
class TestRomanToArabic(unittest.TestCase):
    def test_handles_one_roman_numeral(self):
        self.assertEqual(rrr.roman_to_arabic("I"),1)
        self.assertEqual(rrr.roman_to_arabic("V"),5)
        self.assertEqual(rrr.roman_to_arabic("X"),10)
        self.assertEqual(rrr.roman_to_arabic("L"),50)
        self.assertEqual(rrr.roman_to_arabic("C"),100)
        self.assertEqual(rrr.roman_to_arabic("D"),500)
        self.assertEqual(rrr.roman_to_arabic("M"),1000)
    def test_handles_two_roman_numerals(self):
        self.assertEqual(rrr.roman_to_arabic("IV"),4)
        self.assertEqual(rrr.roman_to_arabic("CC"),200)
        self.assertEqual(rrr.roman_to_arabic("IC"),99)
        self.assertEqual(rrr.roman_to_arabic("MM"),2000)
        self.assertEqual(rrr.roman_to_arabic("XL"),40)
        self.assertEqual(rrr.roman_to_arabic("XX"),20)
    def test_handles_large_roman_numerals(self):
        self.assertEqual(rrr.roman_to_arabic("MMXV"),2015)
    def test_handles_lower_case(self):
        self.assertEqual(rrr.roman_to_arabic("cc"),200)
class TestCustomSorted(unittest.TestCase):
    def test_sorts_simple_numbers(self):
        files = ["1","2","3","4","5","6","10","7","8","9"]
        files = rrr.customsorted(files)
        self.assertEqual(files,["1","2","3","4","5","6","7","8","9","10"])
    def test_sorts_simple_roman_numerals(self):
        files = ["X","IX","VIII","VII","VI","V","IV","III","II","I"]
        files = rrr.customsorted(files)
        self.assertEqual(files,["I","II","III","IV","V","VI","VII","VIII","IX","X"])
    def test_sorts_multiple_levels(self):
        files = ["I.1.a","I.1.b","I.1.c","II.3.b","I.2.a","I.2.b","II.3.a","I.2.c","I.3.a","I.3.b","I.3.c","III","II.1.a","II.1.b","II.1.c","II.2.a","II.2.b","II.2.c","III.3.a","III.3.b","II.3.c","III.1.a","III.1.b","III.3.c","III.1.c","III.2.a","III.2.b","III.2.c"]
        files = rrr.customsorted(files)
        self.assertEqual(files,["I.1.a","I.1.b","I.1.c","I.2.a","I.2.b","I.2.c","I.3.a","I.3.b","I.3.c","II.1.a","II.1.b","II.1.c","II.2.a","II.2.b","II.2.c","II.3.a","II.3.b","II.3.c","III","III.1.a","III.1.b","III.1.c","III.2.a","III.2.b","III.2.c","III.3.a","III.3.b","III.3.c"])
    def test_sorts_i_a_followed_by_roman_numeral_followed_by_text(self):
        files = ["1.a.VIII text","1.a.IX text"]
        files = rrr.customsorted(files)
        self.assertEqual(files,["1.a.VIII text","1.a.IX text"])
    def test_sorts_roman_8_and_9_correctly(self):
        files = ["IX","VIII"]
        files = rrr.customsorted(files)
        self.assertEqual(files,["VIII","IX"])
    def test_sorts_roman_8_and_9_correctly_with_text_afterwards(self):
        files = ["VIII text","IX text"]
        files = rrr.customsorted(files)
        self.assertEqual(files,["VIII text","IX text"])
class TestCustomSortUnit(unittest.TestCase):
    def test_compare_single_character_to_arabic_numeral(self):
        a = rrr.CustomSortUnit("a")
        b = rrr.CustomSortUnit("10")
        self.assertTrue(a>=b)
        self.assertFalse(b>=a)
