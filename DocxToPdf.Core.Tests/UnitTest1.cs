using System;
using System.Xml;
using System.Xml.Linq;
using Xunit;

namespace DocxToPdf.Core.Tests
{
    public class UnitTest1
    {
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
