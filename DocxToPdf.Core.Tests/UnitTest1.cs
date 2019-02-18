using System;
using System.IO;
using System.Xml;
using System.Xml.Linq;
using Xunit;

namespace DocxToPdf.Core.Tests
{
    public class UnitTest1
    {
        [Fact]
        public void CreatePdfShouldCreateValidPdfStructure()
        {
            var converter = new XDocConverter();
            FileStream file = new FileStream(@"c:\Dumpzone\pdfgen.pdf", FileMode.Create);
            PdfDocument pdf = new PdfDocument();

            CatalogObject catalogDict = new CatalogObject();
            PageTreeObject pageTreeDict = new PageTreeObject();
            FontObject Courier = new FontObject();
            InfoObject infoDict = new InfoObject();
            Courier.CreateFont("T1", "CourierNew");
            infoDict.SetInfo("pdftest", "Rob", "3Squared");

            int size = 0;
            file.Write(pdf.GetHeader("1.4", out size));
            file.Flush();
            var page = new PageObject();

            var body = new ContentObject();

            page.CreatePage(pageTreeDict.objectNum, new PageDescription(612, 792,10,10,10,10));
            pageTreeDict.AddPage(page);
            page.AddResource(Courier, body.objectNum);

            //AddRow(false, 10, "T1", align, "First Column", "Second Column");
            //textAndtable.AddRow(false, 10, "T1", align, "Second Row", "Second Row");
            body.SetStream(converter.TextObject(0, 0, "BOLLOX", "T1", 12, "left"));
            body.SetStream(converter.TextObject(10, 10, "BOLLOX", "T1", 12, "left"));
            body.SetStream(converter.TextObject(20, 20, "BOLLOX", "T1", 12, "left"));
            body.SetStream(converter.TextObject(30, 30, "BOLLOX", "T1", 12, "left"));
            body.SetStream(converter.TextObject(40, 40, "BOLLOX", "T1", 12, "left"));
            body.SetStream(converter.TextObject(50, 50, "BOLLOX", "T1", 12, "left"));
            body.SetStream(converter.TextObject(60, 60, "BOLLOX", "T1", 12, "left"));

            file.Write(page.GetPageDict(file.Length, out size), 0, size);
            file.Write(body.GetContentDict(file.Length, out size), 0, size);
            file.Write(catalogDict.GetCatalogDict(pageTreeDict.objectNum,file.Length, out size), 0, size);
            file.Write(pageTreeDict.GetPageTree(file.Length, out size), 0, size);
            file.Write(Courier.GetFontDict(file.Length, out size), 0, size);
            file.Write(infoDict.GetInfoDict(file.Length, out size), 0, size);
            file.Write(pdf.CreateXrefTable(file.Length, out size), 0, size);
            file.Write(pdf.GetTrailer(catalogDict.objectNum,infoDict.objectNum, out size), 0, size);
            file.Close();
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

