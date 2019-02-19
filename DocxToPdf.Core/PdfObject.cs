using System;
using System.Text;

namespace DocxToPdf.Core
{
    /// <summary>
    /// Root object class - responsible for assigning sequential object numbers.
    /// </summary>
    public class PdfObject
    {
        //Incremental object number for EACH OBJECT.
        protected static uint NextObjectNum { get; set; }

        protected string ObjectRepresenation;

        //the Inherited object number for each derived object type.
        public uint objectNum;

        /// <summary>
        /// Constructor increments the object number to 
        /// reflect the currently used object number
        /// </summary>
        protected PdfObject()
        {
            NextObjectNum++;
            objectNum = NextObjectNum;
            ObjectRepresenation = String.Empty;
        }

        ~PdfObject()
        {
            objectNum = 0;
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
            ObjectXRef obj = new ObjectXRef(objectNum, filePos);
            byte[] abuf;
            try
            {
                byte[] ubuf = Encoding.Unicode.GetBytes(str);
                Encoding enc = Encoding.GetEncoding("utf-8");
                abuf = Encoding.Convert(Encoding.Unicode, enc, ubuf);
                size = abuf.Length;
                PdfDocument.xrefTable.ObjectByteOffsets.Add(obj);
            }
            catch (Exception e)
            {
                string str1 = $"{objectNum},In PdfObjects.GetBytes()";
                Exception error = new Exception(e.Message + str1);
                throw error;
            }

            return abuf;
        }
    }
}