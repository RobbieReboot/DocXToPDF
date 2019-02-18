using System;
using System.Collections.Generic;
using System.Text;

namespace DocxToPdf.Core
{
    public class XRefTable
    {
        public XRefTable()
        {
        }

        public byte[] CreateXrefTable(long fileOffset, out int size)
        {
            //Store the Offset of the Xref table for startxRef
            ObjectXRef objList = new ObjectXRef(0, fileOffset);
            XrefEntries.offsetArray.Add(objList);
            XrefEntries.offsetArray.Sort();
            var numTableEntries = (uint)XrefEntries.offsetArray.Count;
            var table = string.Format("xref\r\n{0} {1}\r\n0000000000 65535 f\r\n", 0, numTableEntries);
            for (int entries = 1; entries < numTableEntries; entries++)
            {
                ObjectXRef obj = (ObjectXRef)XrefEntries.offsetArray[entries];
                table += obj.offset.ToString().PadLeft(10, '0');
                table += " 00000 n\r\n";
            }
            return PdfDocument.GetUTF8Bytes(table, out size);
        }
    }


}
