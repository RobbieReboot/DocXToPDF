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

            contentObj.AddTextObject(0, 0, "TEST", courierNew, 12, "left");
            contentObj.AddTextObject(0, 0, "TEST", courierNew, 12, "left");
            contentObj.AddTextObject(10, 10, "TEST", courierNew, 12, "left");
            contentObj.AddTextObject(20, 20, "TEST", courierNew, 12, "left");
            contentObj.AddTextObject(30, 30, "TEST", courierNew, 12, "left");
            contentObj.AddTextObject(40, 40, "TEST", courierNew, 12, "left");
            contentObj.AddTextObject(50, 50, "TEST", courierNew, 12, "left");
            pdf.Write(@"C:\Dumpzone\pdfUnit1.pdf");
        }

        [Fact]
        public void PdfConvertConvertsDocXWhenDocXIsValid()
        {
            var reader = XmlReader.Create(@"C:\Dumpzone\steve\file_4.xml");
            var xdoc = XDocument.Load(reader);
            var nsm = xdoc.CreateReader().NameTable;
            
            PdfDocument.FromDocX(xdoc).Write(@"C:\Dumpzone\steve\file_1.pdf");
            xdoc.ToPdf().Write(@"C:\Dumpzone\steve\file_1a.pdf");
        }
    }
}

