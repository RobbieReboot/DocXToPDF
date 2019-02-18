using System;
using System.Collections.Generic;
using System.Text;

namespace DocxToPdf.Core
{
    public class PdfDocument
    {
        public PdfDocument()
        {
            
        }



        public static byte[] GetHeader(string version, out int size)
        {
            //            var headerCode = 0xe2e3cfd3;
            //            string header = $"%PDF-{version}\r%{headerCode}\r\n";
            string header = string.Format("%PDF-{0}\r%{1}\r\n", version, "\x82\x82");
            return GetUTF8Bytes(header, out size);
        }
        /// <summary>
        /// Creates the trailer and return the bytes array
        /// </summary>
        /// <param name="size"></param>
        /// <returns></returns>
        public static byte[] GetTrailer(uint refRoot, uint refInfo, out int size)
        {
            string trailer = null;
            try
            {
                if (refInfo > 0)
                {
                    infoDict = string.Format("/Info {0} 0 R", refInfo);
                }

                //The sorted array will be already sorted to contain the file offset at the zeroth position
                ObjectXRef objList = (ObjectXRef)XrefEntries.offsetArray[0];
                trailer = string.Format("trailer\n<</Size {0}/Root {1} 0 R {2}/ID[<5181383ede94727bcb32ac27ded71c68>" +
                                        "<5181383ede94727bcb32ac27ded71c68>]>>\r\nstartxref\r\n{3}\r\n%%EOF\r\n"
                    , numTableEntries, refRoot, infoDict, objList.offset);

                XrefEntries.offsetArray = null;
                PdfObject.NextObjectNum = 0;
            }
            catch (Exception e)
            {
                Exception error = new Exception(e.Message + " In Utility.GetTrailer()");
                throw error;
            }

            return GetUTF8Bytes(trailer, out size);
        }

        public static byte[] GetUTF8Bytes(string str, out int size)
        {
            try
            {
                byte[] ubuf = Encoding.Unicode.GetBytes(str);
                Encoding enc = Encoding.GetEncoding("utf-8");
                byte[] abuf = Encoding.Convert(Encoding.Unicode, enc, ubuf);
                size = abuf.Length;
                return abuf;
            }
            catch (Exception e)
            {
                Exception error = new Exception(e.Message + " In Utility.GetUTF8Bytes()");
                throw error;
            }
        }
    }
}
