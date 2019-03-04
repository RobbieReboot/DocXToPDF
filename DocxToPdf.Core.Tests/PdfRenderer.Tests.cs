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
        [InlineData(@"(", @"\(")]
        [InlineData(@")", @"\)")]
        [InlineData(@"\", @"\\")]
        public void PdfSanitizerCheckShouldSanitizeStrings(string input,string sanitized)
        {
            var result = PdfDocument.SanitizePdfCharacters(input);
            Assert.Equal(sanitized,result);
        }

        [Theory]
        [InlineData("validDocX")]
        public void ShouldCreateValidPdfFromConsecutiveRunsWithTabs(string folder)
        {
            var filter = "*.docx";
            IEnumerable<string> GetFilesFromDir(string dir) =>
                Directory.EnumerateFiles(dir, filter).Concat(
                    Directory.EnumerateDirectories(dir)
                        .SelectMany(subdir => GetFilesFromDir(subdir)));

            var validDocs = GetFilesFromDir(folder);

            foreach (var docx in validDocs)
            {
                var pdf = PdfDocument.FromDocX(docx);
                var outputName = Path.ChangeExtension(docx, "pdf");
                pdf.Write(outputName);
            }
        }
        //file_3 was DIFFERENT! it contained no p:pPr with tabs , they were only in the styles.
        [Theory]
        [InlineData(@"validDocX\file_3.docx")]
        public void File3ShouldRenderProperly(string file)
        {
            var pdf = PdfDocument.FromDocX(file);
            var outputName = Path.ChangeExtension(file, "pdf");
            pdf.Write(outputName);
        }

        [Theory]
        [InlineData(@"validDocX\file_3.docx")]
        public void File3AsMemoryStream(string fileName)
        {
            using (var memoryStream = new MemoryStream(File.ReadAllBytes(fileName)))
            {
                var pdf = PdfDocument.FromDocX(memoryStream);
                var outputName = Path.ChangeExtension(fileName, "pdf");
                pdf.Write(outputName);
            }
            
        }

        [Theory]
        [InlineData(@"validDocX\file_3.docx")]
        public void File3AsMemoryStream2(string fileName)
        {
            var memoryStream = new MemoryStream();
            var pdf = PdfDocument.FromDocX(fileName);
            var outputName = Path.ChangeExtension(fileName, "pdf");
            pdf.Write(out memoryStream);
            using (FileStream file = new FileStream(outputName, FileMode.OpenOrCreate, FileAccess.ReadWrite))
            {
                memoryStream.WriteTo(file);
                file.Close();
            }
        }

    }
}

