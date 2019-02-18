using System.Collections;

namespace DocxToPdf.Core
{
    /// <summary>
    /// Holds the Byte offsets of the PDF objects
    /// </summary>
    internal class XrefEntries
    {
        internal static ArrayList offsetArray;

        internal XrefEntries()
        {
            offsetArray = new ArrayList();
        }
    }
}