using System;
using System.Collections.Generic;
using System.Text;

namespace DocxToPdf.Core
{
    public struct PageDescription
    {
        public uint xWidth;
        public uint yHeight;
        public uint leftMargin;
        public uint rightMargin;
        public uint topMargin;
        public uint bottomMargin;

        public PageDescription(uint width, uint height, uint left = 0, uint top = 0, uint right = 0, uint bottom = 0)
        {
            xWidth = width;
            yHeight = height;
            leftMargin = left;
            rightMargin = right;
            topMargin = top;
            bottomMargin = bottom;
        }

        public void SetMargins(uint left, uint top, uint right, uint bottom)
        {
            leftMargin = left;
            rightMargin = right;
            topMargin = top;
            bottomMargin = bottom;
        }
    }

    public enum Justification
    {
        Left,
        Center,
        Right
    }


    /// <summary>
    /// This class contains general Utility for the creation of pdf
    /// Creates the Header
    /// Creates XrefTable
    /// Creates the Trailer
    /// </summary>
    public class Utility
    {
        private uint numTableEntries;
        private string table;
        private string infoDict;

        public Utility()
        {
            numTableEntries = 0;
            table = null;
            infoDict = null;
        }

        /// <summary>
        /// Creates the xref table using the byte offsets in the array.
        /// </summary>
        /// <param name="array"></param>
        /// <param name="size"></param>
        /// <returns></returns>
        public byte[] CreateXrefTable(long fileOffset, out int size)
        {
            //Store the Offset of the Xref table for startxRef
            try
            {
                ObjectXRef objList = new ObjectXRef(0, fileOffset);
                XrefEnteries.offsetArray.Add(objList);
                XrefEnteries.offsetArray.Sort();
                numTableEntries = (uint) XrefEnteries.offsetArray.Count;
                table = string.Format("xref\r\n{0} {1}\r\n0000000000 65535 f\r\n", 0, numTableEntries);
                for (int entries = 1; entries < numTableEntries; entries++)
                {
                    ObjectXRef obj = (ObjectXRef) XrefEnteries.offsetArray[entries];
                    table += obj.offset.ToString().PadLeft(10, '0');
                    table += " 00000 n\r\n";
                }
            }
            catch (Exception e)
            {
                Exception error = new Exception(e.Message + " In Utility.CreateXrefTable()");
                throw error;
            }

            return GetUTF8Bytes(table, out size);
        }

        public byte[] GetHeader(string version, out int size)
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
        public byte[] GetTrailer(uint refRoot, uint refInfo, out int size)
        {
            string trailer = null;
            try
            {
                if (refInfo > 0)
                {
                    infoDict = string.Format("/Info {0} 0 R", refInfo);
                }

                //The sorted array will be already sorted to contain the file offset at the zeroth position
                ObjectXRef objList = (ObjectXRef) XrefEnteries.offsetArray[0];
                trailer = string.Format("trailer\n<</Size {0}/Root {1} 0 R {2}/ID[<5181383ede94727bcb32ac27ded71c68>" +
                                        "<5181383ede94727bcb32ac27ded71c68>]>>\r\nstartxref\r\n{3}\r\n%%EOF\r\n"
                    , numTableEntries, refRoot, infoDict, objList.offset);

                XrefEnteries.offsetArray = null;
                PdfObject.NextObjectNum = 0;
            }
            catch (Exception e)
            {
                Exception error = new Exception(e.Message + " In Utility.GetTrailer()");
                throw error;
            }

            return GetUTF8Bytes(trailer, out size);
        }

        private byte[] GetUTF8Bytes(string str, out int size)
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

        private int StrLen(string text, int fontSize)
        {
            char[] cArray = text.ToCharArray();
            int cWidth = 0;
            foreach (char c in cArray)
            {
                cWidth += 500; //(int)(fontSize*1.6)*20;	//Monospaced font width?
            }

            //div by1000??? 100 seems to work better :/
            return (cWidth / 100);
        }

        /// <summary>
        /// start the Page Text element at the X Y position
        /// </summary>
        /// <returns></returns>
        public string AddText(uint X, uint Y, string text, int fontSize, string fontName, Justification alignment)
        {
            long startX = 0;
            switch (alignment)
            {
                case (Justification.Left):
                    startX = X;
                    break;
                case (Justification.Center):
                    startX = X - (StrLen(text, fontSize)) / 2;
                    break;
                case (Justification.Right):
                    startX = X - StrLen(text, fontSize);
                    break;
            }

            ;
            return $"\rBT/{fontName} {fontSize} Tf\r{startX} {(792 - Y)} Td \r({text}) Tj\rET\r";
        }
    }
}