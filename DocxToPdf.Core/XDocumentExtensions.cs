using System;
using System.Collections.Generic;
using System.Text;
using System.Xml.Linq;

namespace DocxToPdf.Core
{
    /// <summary>
    /// Conventience Methods
    /// </summary>
    public static class XDocumentExtensions
    {
        public static PdfDocument ToPdf(this XDocument xdoc)
        {
            return PdfDocument.FromDocXXml(xdoc);
        }
    }
}
