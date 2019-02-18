using System;
using System.Text;

namespace DocxToPdf.Core
{
    public class PdfObject
    {
        //Incremental object number for EACH OBJECT.
        internal static uint NextObjectNum;

        //the Inherited object number for each derived object type.
        public uint objectNum;

        //private UTF8Encoding utf8;
        private XrefEnteries Xref;

        /// <summary>
        /// Constructor increments the object number to 
        /// reflect the currently used object number
        /// </summary>
        protected PdfObject()
        {
            if (NextObjectNum == 0)
                Xref = new XrefEnteries();
            NextObjectNum++;
            objectNum = NextObjectNum;
        }

        ~PdfObject()
        {
            objectNum = 0;
        }

        protected byte[] GetUTF8Bytes(string str, long filePos, out int size)
        {
            ObjectXRef objList = new ObjectXRef(objectNum, filePos);
            byte[] abuf;
            try
            {
                byte[] ubuf = Encoding.Unicode.GetBytes(str);
                Encoding enc = Encoding.GetEncoding("utf-8");
                abuf = Encoding.Convert(Encoding.Unicode, enc, ubuf);
                size = abuf.Length;
                XrefEnteries.offsetArray.Add(objList);
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