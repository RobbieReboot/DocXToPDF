using System;
using System.IO;
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
            pdf.Write(@"C:\Dumpzone\pdfUnit1.pdf");
        }

        [Theory]
        [InlineData(@"Data\file_3.xml")]
        public void PdfConvertConvertsDocXWhenDocXIsValid(string fileName)
        {
            var reader = XmlReader.Create(fileName);
            var xdoc = XDocument.Load(reader);
            var p = Directory.GetCurrentDirectory();
//            PdfDocument.FromDocX(xdoc).Write(@"C:\Dumpzone\steve\file_1.pdf");
            xdoc.ToPdf().Write($@"C:\Dumpzone\steve\{Path.GetFileName(fileName)}.pdf");
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

