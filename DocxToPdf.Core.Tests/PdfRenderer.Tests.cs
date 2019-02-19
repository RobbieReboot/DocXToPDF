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
            FontObject CourierNew = new FontObject("CourierNew");

            pdf.SetMetadata(new InfoObject("pdftest", "Rob", "3Squared"));

            var page = new PageObject();

            pdf.AddPage(page, new PageDescription(612, 792, 10, 10, 10, 10));
            page.AddFont(CourierNew);

            var contentObj = new ContentObject();
            page.AddContent(contentObj);

            contentObj.AddTextObject(0, 0, "BOLLOX", CourierNew, 12, "left");
            contentObj.AddTextObject(0, 0, "BOLLOX", CourierNew, 12, "left");
            contentObj.AddTextObject(10, 10, "BOLLOX", CourierNew, 12, "left");
            contentObj.AddTextObject(20, 20, "BOLLOX", CourierNew, 12, "left");
            contentObj.AddTextObject(30, 30, "BOLLOX", CourierNew, 12, "left");
            contentObj.AddTextObject(40, 40, "BOLLOX", CourierNew, 12, "left");
            contentObj.AddTextObject(50, 50, "BOLLOX", CourierNew, 12, "left");
            pdf.Write(@"C:\Dumpzone\pdfUnit1.pdf");
        }

        [Fact]
        public void PdfConvertConvertsDocXWhenDocXIsValid()
        {
            var reader = XmlReader.Create(@"C:\Dumpzone\steve\file_1.xml");
            var xdoc = XDocument.Load(reader);
            var nsm = xdoc.CreateReader().NameTable;
            
            //PdfDocument pdf = PdfDocument.FromDocX(xdoc);
            //pdf.Write(@"C:\Dumpzone\steve\file_1.pdf");#
            //OR
            PdfDocument.FromDocX(xdoc).Write(@"C:\Dumpzone\steve\file_1.pdf");
            xdoc.ToPdf().Write(@"C:\Dumpzone\steve\file_1a.pdf");
        }
    }
}

