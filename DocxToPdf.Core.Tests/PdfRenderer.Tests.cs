using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml;
using System.Xml.Linq;
using Xunit;

namespace DocxToPdf.Core.Tests
{
    public class PdfRendererTests
    {
        [Fact]
        public void CreatePdfShouldCreateValidPdfStructure()
        {
            PdfDocument pdf = new PdfDocument();
            FontObject courierNew = pdf.AddFont("CourierNew");

            pdf.AddInfoObject("pdftest", "Rob", "3Squared");

            var page = pdf.AddPage(new PageExtents(612, 792, 24, 10, 24, 10));
            page.AddFont(courierNew);

            var contentObj = page.AddContentObject();

            contentObj.AddTextObject(0, 0, "BOLLOX", courierNew, 12, "left");
            contentObj.AddTextObject(0, 0, "BOLLOX", courierNew, 12, "left");
            contentObj.AddTextObject(10, 10, "BOLLOX", courierNew, 12, "left");
            contentObj.AddTextObject(20, 20, "BOLLOX", courierNew, 12, "left");
            contentObj.AddTextObject(30, 30, "BOLLOX", courierNew, 12, "left");
            contentObj.AddTextObject(40, 40, "BOLLOX", courierNew, 12, "left");
            contentObj.AddTextObject(50, 50, "BOLLOX", courierNew, 12, "left");
            pdf.Write(@"output\pdfUnit1.pdf");
        }

        [Theory]
        [InlineData("validXml")]
        [InlineData("invalidXml")]

        public void PdfConvertConvertsDocXShouldThrowNoExceptionsWhenDocXIsValid(string folder)
        {
            //process all input folders to the output folder.

            var filter = "*.xml";
            IEnumerable<string> GetFilesFromDir(string dir) =>
                Directory.EnumerateFiles(dir, filter).Concat(
                    Directory.EnumerateDirectories(dir)
                        .SelectMany(subdir => GetFilesFromDir(subdir)));

            var validXml = GetFilesFromDir(folder);

            foreach (var xmlFile in validXml)
            {
                var reader = XmlReader.Create(xmlFile);
                var xdoc = XDocument.Load(reader);
                var outFile = Path.Combine("output", Path.ChangeExtension(Path.GetFileName(xmlFile), "pdf"));
                xdoc.ToPdf().Write(outFile);
            }
        }

        [Theory]
        [InlineData(@"(", @"\(")]
        [InlineData(@")", @"\)")]
        [InlineData(@"\", @"\\")]
        public void PdfSanitizerCheckShouldSanitizeStrings(string input,string sanitized)
        {
            var result = PdfDocument.SanitizePdfCharacters(input);
            Assert.Equal(sanitized,result);
        }
    }
}

