using System.IO.Compression;
using System.Text;
using static SigSpoox;

namespace sigspooxTests
{
    [TestClass]
    public class SigSpOOXTests
    {

        private bool compareZipContents(string testvector, string test)
        {
            string Tmpdir = Directory.CreateDirectory(Path.GetTempPath() + Path.GetRandomFileName()).ToString();
            string tmpdirTestvector = Path.Combine(Tmpdir, "testvector");
            string tmpdirTest = Path.Combine(Tmpdir, "test");

            Directory.CreateDirectory(tmpdirTestvector);
            Directory.CreateDirectory(tmpdirTest);

            ZipFile.ExtractToDirectory(testvector, tmpdirTestvector);
            ZipFile.ExtractToDirectory(test, tmpdirTest);

            DirectoryInfo testvectorDirInfo = new DirectoryInfo(tmpdirTestvector);
            DirectoryInfo testDirInfo = new DirectoryInfo(tmpdirTest);

            HashSet<string> testvectorFiles = testvectorDirInfo.GetFiles("*.*", SearchOption.AllDirectories).Select(f => f.FullName).ToHashSet();
            HashSet<string> testFiles = testDirInfo.GetFiles("*.*", SearchOption.AllDirectories).Select(f => f.FullName).ToHashSet();
            HashSet<string> testFilesFixedPath = testDirInfo.GetFiles("*.*", SearchOption.AllDirectories).Select(f => f.FullName.Replace("test", "testvector")).ToHashSet();

            // are both lists the same?
            // have to fix path of testFiles to match the path of testvectorFiles
            bool containSameFilesByName = testvectorFiles.SetEquals(testFilesFixedPath);

            if (containSameFilesByName == false)
            {
                Directory.Delete(Tmpdir, true);
                return false;
            }

            // do a bytewise comparison of all files
            foreach (string file in testvectorFiles)
            {
                byte[] testvectorBytes = File.ReadAllBytes(file);
                byte[] testBytes = File.ReadAllBytes(file.Replace("vector", ""));

                if (testvectorBytes.SequenceEqual(testBytes) == false)
                {
                    Directory.Delete(Tmpdir, true);
                    return false;
                }
            }

            Directory.Delete(Tmpdir, true);
            return true;

        }

        [TestMethod]
        public void findBytePattern_Fail()
        {
            byte[] bytes = Encoding.ASCII.GetBytes(@"<?xml version=""1.0"" encoding=""UTF-8""?><Signature");
            byte[] pattern = Encoding.ASCII.GetBytes(@"<?xml version=""1.0"" GARBAGE encoding=""UTF-8""?><Signature");
            int offset = 0;
            int startPos = -1;
            int endPos = -1;
            (int, int) result = findBytePattern(bytes, pattern, offset);
            Assert.AreEqual((startPos, endPos), result);
        }

        [TestMethod]
        public void findBytePattern_Success()
        {
            byte[] bytes = Encoding.ASCII.GetBytes(@"<?xml version=""1.0"" encoding=""UTF-8""?><Signature");
            byte[] pattern = Encoding.ASCII.GetBytes(@"<?xml version=""1.0"" encoding=""UTF-8""?><Signature");
            int startPos = 0;
            int endPos = bytes.Length;
            (int, int) result = findBytePattern(bytes, pattern, startPos);
            Assert.AreEqual((startPos, endPos), result);
        }

        [TestMethod]
        public void ETA_Excel_FunctionalTest()
        {
            EtaOptions options = new()
            {
                SignedFile = @"files\signed\excel_signed.xlsx",
                AttackerFile = @"files\attacker\attacker.docx",
                ResultFile = @"files\results\eta_excel.docx"
            };
            ETA(options);
            Assert.IsTrue(compareZipContents(options.ResultFile, options.ResultFile.Replace("results", "testvectors")));
        }

        [TestMethod]
        public void ETA_PowerPoint_FunctionalTest()
        {
            EtaOptions options = new()
            {
                SignedFile = @"files\signed\powerpoint_signed.pptx",
                AttackerFile = @"files\attacker\attacker.docx",
                ResultFile = @"files\results\eta_powerpoint.docx"
            };
            ETA(options);
            Assert.IsTrue(compareZipContents(options.ResultFile, options.ResultFile.Replace("results", "testvectors")));
        }

        [TestMethod]
        public void DDA_FunctionalTest()
        {
            DdaOptions options = new()
            {
                SignedFile = @"files\signed\word_signed.docx",
                AttackerFile = @"files\attacker\attacker.docx",
                ResultFile = @"files\results\dda.docx"
            };
            DDA(options);
            Assert.IsTrue(compareZipContents(options.ResultFile, options.ResultFile.Replace("results", "testvectors")));
        }

        [TestMethod]
        public void LWA_Excel_FunctionalTest()
        {
            LwaExcelOptions options = new()
            {
                SignedFile = @"files\signed\excel_legacy_signed.xls",
                AttackerFile = @"files\attacker\attacker.docx",
                ResultFile = @"files\results\lwa_excel.docx"
            };
            LWAExcel(options);
            Assert.IsTrue(compareZipContents(options.ResultFile, options.ResultFile.Replace("results", "testvectors")));
        }

        [TestMethod]
        public void LWA_Word_Prepare_FunctionalTest()
        {
            LwaWordPrepOptions options = new()
            {
                SignedFile = @"files\attacker\word_legacy_benign.doc",
                ResultFile = @"files\results\lwa_word_prep.doc"
            };
            LWAWordPrep(options);
            byte[] testvectorBytes = File.ReadAllBytes(options.ResultFile);
            byte[] testBytes = File.ReadAllBytes(options.ResultFile.Replace("results", "testvectors"));
            bool sameBytes = testvectorBytes.SequenceEqual(testBytes);
            Assert.IsTrue(sameBytes);
        }

        [TestMethod]
        public void LWA_Word_Final_FunctionalTest()
        {
            LwaWordFinalOptions options = new()
            {
                SignedFile = @"files\signed\word_legacy_lwa_prepared.doc",
                AttackerFile = @"files\attacker\attacker.docx",
                ResultFile = @"files\results\lwa_word_final.docx"
            };
            LWAWordFinal(options);
            Assert.IsTrue(compareZipContents(options.ResultFile, options.ResultFile.Replace("results", "testvectors")));
        }

        [TestMethod]
        public void USF_AdES_FunctionalTest()
        {
            UsfOptions options = new()
            {
                SignedFile = @"files\signed\odt_ades_signed.odt",
                AttackerFile = @"files\attacker\attacker_selfsigned.docx",
                ResultFile = @"files\results\usf_odt_ades.docx"
            };
            USF(options);
            Assert.IsTrue(compareZipContents(options.ResultFile, options.ResultFile.Replace("results", "testvectors")));
        }

        [TestMethod]
        public void USF_NoAdES_FunctionalTest()
        {
            UsfOptions options = new()
            {
                SignedFile = @"files\signed\odt_no_ades_signed.odt",
                AttackerFile = @"files\attacker\attacker_selfsigned.docx",
                ResultFile = @"files\results\usf_odt_no_ades.docx"
            };
            USF(options);
            Assert.IsTrue(compareZipContents(options.ResultFile, options.ResultFile.Replace("results", "testvectors")));
        }

        [TestMethod]
        public void CIA_FunctionalTest()
        {
            CiaOptions options = new()
            {
                SignedFile = @"files\signed\word_signed.docx",
                AttackerFile = @"files\attacker\attacker_cia.docx",
                ResultFile = @"files\results\cia.docx"
            };
            CIA(options);
            Assert.IsTrue(compareZipContents(options.ResultFile, options.ResultFile.Replace("results", "testvectors")));
        }

        [TestMethod]
        public void FIA_Word_Prepare_FunctionalTest()
        {
            FiaWordPrepOptions options = new()
            {
                AttackerFile = @"files\attacker\attacker_fia.docx",
                ResultFile = @"files\results\fia_word_prep.docx"
            };
            FiaWordPrep(options);
            Assert.IsTrue(compareZipContents(options.ResultFile, options.ResultFile.Replace("results", "testvectors")));
        }

        [TestMethod]
        public void FIA_Word_Final_FunctionalTest()
        {
            FiaWordFinalOptions options = new()
            {
                SignedFile = @"files\signed\fia_signed.docx",
                ResultFile = @"files\results\fia.docx"
            };
            FiaWordFinal(options);
            Assert.IsTrue(compareZipContents(options.ResultFile, options.ResultFile.Replace("results", "testvectors")));
        }

        [TestMethod]
        public void SIA_Word_Prepare_FunctionalTest()
        {
            SiaWordPrepOptions options = new()
            {
                AttackerFile = @"files\attacker\attacker_sia.docx",
                ResultFile = @"files\results\sia_word_prep.docx"
            };
            SiaWordPrep(options);
            Assert.IsTrue(compareZipContents(options.ResultFile, options.ResultFile.Replace("results", "testvectors")));
        }

        [TestMethod]
        public void SIA_Word_Final_Body_FunctionalTest()
        {
            SiaWordFinalOptions options = new()
            {
                SignedFile = @"files\signed\word_sia_prepared_signed.docx",
                AttackerFile = @"files\attacker\attacker.docx",
                ResultFile = @"files\results\sia_body.docx"
            };
            SiaWordFinal(options);
            Assert.IsTrue(compareZipContents(options.ResultFile, options.ResultFile.Replace("results", "testvectors")));
        }

        [TestMethod]
        public void SIA_Word_Final_FunctionalTest()
        {
            SiaWordFinalOptions options = new()
            {
                SignedFile = @"files\signed\word_sia_prepared_signed.docx",
                ResultFile = @"files\results\sia.docx"
            };
            SiaWordFinal(options);
            Assert.IsTrue(compareZipContents(options.ResultFile, options.ResultFile.Replace("results", "testvectors")));
        }
    }
}