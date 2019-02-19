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
            var converter = new XDocConverter();
            PdfDocument pdf = new PdfDocument();
            FontObject Courier= new FontObject("CourierNew");

            pdf.SetMetadata(new InfoObject("pdftest", "Rob", "3Squared"));

            var page = new PageObject();
            var contentObj = new ContentObject();
            page.AddContent(contentObj);

            pdf.AddPage(page, new PageDescription(612, 792, 10, 10, 10, 10));
            page.AddFont(Courier);

            contentObj.AddObject(new TextObject(0, 0, "BOLLOX", "T1", 12, "left"));
            contentObj.AddObject(new TextObject(10, 10, "BOLLOX", "T1", 12, "left"));
            contentObj.AddObject(new TextObject(20, 20, "BOLLOX", "T1", 12, "left"));
            contentObj.AddObject(new TextObject(30, 30, "BOLLOX", "T1", 12, "left"));
            contentObj.AddObject(new TextObject(40, 40, "BOLLOX", "T1", 12, "left"));
            contentObj.AddObject(new TextObject(50, 50, "BOLLOX", "T1", 12, "left"));

            pdf.Write(@"C:\Dumpzone\pdfUnit1.pdf");
        }

        [Fact]
        public void PdfConvertConvertsDocXWhenDocXIsValid()
        {
            var reader = XmlReader.Create(@"C:\Dumpzone\steve\file_4.xml");
            var xdoc = XDocument.Load(reader);
            var nsm = xdoc.CreateReader().NameTable;

            var converted = new XDocConverter();

            var pdf = converted.ToPdf(xdoc);
        }
    }
}

