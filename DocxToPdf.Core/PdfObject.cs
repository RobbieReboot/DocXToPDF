using System;
using System.Text;

namespace DocxToPdf.Core
{
    /// <summary>
    /// Root object class - responsible for assigning sequential object numbers.
    /// </summary>
    public class PdfObject
    {
        protected PdfDocument parentDocument;

        //Incremental object number for EACH OBJECT.

        protected string ObjectRepresenation;

        //the Inherited object number for each derived object type.
        public uint PdfObjectId;

        /// <summary>
        /// Constructor increments the object number to 
        /// reflect the currently used object number
        /// </summary>
        protected PdfObject(PdfDocument parent)
        {
            parentDocument = parent;

            parentDocument.NextObjectNum++;
            PdfObjectId = parentDocument.NextObjectNum;
            ObjectRepresenation = String.Empty;
        }

        ~PdfObject()
        {
            PdfObjectId = 0;
        }


        /// <summary>
        /// return this object
        /// </summary>
        /// <param name="str"></param>
        /// <param name="filePos"></param>
        /// <param name="size"></param>
        /// <returns></returns>
        protected byte[] GetUTF8Bytes(string str, long filePos, out int size)
        {
            ObjectXRef obj = new ObjectXRef(PdfObjectId, filePos);
            byte[] abuf;
            try
            {
                byte[] ubuf = Encoding.Unicode.GetBytes(str);
                Encoding enc = Encoding.GetEncoding("utf-8");
                abuf = Encoding.Convert(Encoding.Unicode, enc, ubuf);
                size = abuf.Length;
                parentDocument.xrefTable.ObjectByteOffsets.Add(obj);
            }
            catch (Exception e)
            {
                string str1 = $"{PdfObjectId},In PdfObjects.GetBytes()";
                Exception error = new Exception(e.Message + str1);
                throw error;
            }

            return abuf;
        }
    }
}