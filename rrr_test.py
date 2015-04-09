import rrr,unittest,os
from PyPDF2 import PdfFileReader
#all tests use test_ as shorthand for tests_that_it_
class TestMain(unittest.TestCase):
    pass
class TestUnzip(unittest.TestCase):
    pass
class TestAddDirectorySlipsheets(unittest.TestCase):
    def test_creates_a_slipsheet(self):
        a = "test_creates_a_slipsheet_test_directory"
        os.mkdir(a)
        rrr.add_directory_slipsheets(a,0)
        self.assertTrue(os.path.isfile(os.path.join(a,"000 ----- INSERT TAB HERE.pdf")))
        os.remove(os.path.join(a,"000 ----- INSERT TAB HERE.pdf"))
        os.rmdir(a)
    def test_creates_a_slipsheet_at_multiple_levels(self):
        a = "test_creates_a_slipsheet_test_directory"
        os.mkdir(a)
        os.mkdir(os.path.join(a,a))
        os.mkdir(os.path.join(a,a,a))
        rrr.add_directory_slipsheets(a,2)
        self.assertTrue(os.path.isfile(os.path.join(a,"000 ----- INSERT TAB HERE.pdf")))
        self.assertTrue(os.path.isfile(os.path.join(a,a,"000 ----- INSERT TAB HERE.pdf")))
        self.assertTrue(os.path.isfile(os.path.join(a,a,a,"000 ----- INSERT TAB HERE.pdf")))
        os.remove(os.path.join(a,"000 ----- INSERT TAB HERE.pdf"))
        os.remove(os.path.join(a,a,"000 ----- INSERT TAB HERE.pdf"))
        os.remove(os.path.join(a,a,a,"000 ----- INSERT TAB HERE.pdf"))
        os.rmdir(os.path.join(a,a,a))
        os.rmdir(os.path.join(a,a))
        os.rmdir(a)
class TestRenameResizeRotate(unittest.TestCase):
    pass
class TestConvertToPdf(unittest.TestCase):
    pass
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
    pass
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
    pass
class TestUpdatePageDimensions(unittest.TestCase):
    pass
class TestZipFound(unittest.TestCase):
    pass
class TestProcessZips(unittest.TestCase):
    pass
class TestProcessZip(unittest.TestCase):
    pass
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
    pass
class TestProcessDoc(unittest.TestCase):
    pass
class TestProcessXls(unittest.TestCase):
    pass
class TestProcessImage(unittest.TestCase):
    pass
class TestAddSlipsheet(unittest.TestCase):
    pass
class TestProcessPdfPages(unittest.TestCase):
    pass
class TestPdfWrite(unittest.TestCase):
    pass
class TestGetPageDimensions(unittest.TestCase):
    pass