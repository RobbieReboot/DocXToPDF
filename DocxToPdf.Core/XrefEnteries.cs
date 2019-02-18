using System.Collections;

namespace DocxToPdf.Core
{
    /// <summary>
    /// Holds the Byte offsets of the PDF objects
    /// </summary>
    internal class XrefEnteries
    {
        internal static ArrayList offsetArray;

        internal XrefEnteries()
        {
            offsetArray = new ArrayList();
        }
    }
}