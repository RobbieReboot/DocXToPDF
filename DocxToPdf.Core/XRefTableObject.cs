using System;
using System.Collections;
using System.Collections.Generic;
using System.Text;

namespace DocxToPdf.Core
{
    public class XRefTableObject
    {
        public List<ObjectXRef> ObjectByteOffsets;
        public int XRefCount => ObjectByteOffsets.Count;
        public long FirstByteOffset => ObjectByteOffsets[0].offset;

        public XRefTableObject()
        {
            ObjectByteOffsets = new List<ObjectXRef>();
        }

        public byte[] RenderBytes(long fileOffset, out int size)
        {
            //Store the Offset of the Xref table for startxRef
            ObjectXRef objList = new ObjectXRef(0, fileOffset);
            ObjectByteOffsets.Add(objList);
            ObjectByteOffsets.Sort();
            var table = $"xref\r\n{0} {XRefCount}\r\n0000000000 65535 f\r\n";
            for (int entries = 1; entries < XRefCount; entries++)
            {
                ObjectXRef obj = (ObjectXRef)ObjectByteOffsets[entries];
                table += obj.offset.ToString().PadLeft(10, '0');
                table += " 00000 n\r\n";
            }
            return PdfDocument.GetUTF8Bytes(table, out size);
        }
    }


}
