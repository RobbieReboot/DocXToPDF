using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace DocxToPdf.Core
{
    public class PdfDocument
    {
        public static XRefTable xrefTable;

        private CatalogObject catalogObj;
        private PageTreeObject pageTreeObj;
        private InfoObject infoObject;
        private List<PageObject> pageObjects;
        public PdfDocument()
        {
            xrefTable = new XRefTable();
            pageTreeObj = new PageTreeObject();
            catalogObj = new CatalogObject(pageTreeObj);
            pageObjects = new List<PageObject>();
        }

        public byte[] CreateXrefTable(long fileOffset, out int size) => xrefTable.CreateXrefTable(fileOffset, out size);
        public void SetMetadata(InfoObject metadata) => infoObject = metadata;

        public void AddPage(PageObject page,PageDescription pDesc)
        {
            page.CreatePage(pageTreeObj.objectNum, pDesc);
            pageTreeObj.AddPage(page);
            pageObjects.Add(page);
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
            string infoDict = String.Empty;

            if (refInfo > 0)
            {
                infoDict = string.Format("/Info {0} 0 R", refInfo);
            }

            //The sorted array will be already sorted to contain the file offset at the zeroth position
            ObjectXRef objList = xrefTable.ObjectByteOffsets[0];
            trailer = string.Format("trailer\n<</Size {0}/Root {1} 0 R {2}/ID[<5181383ede94727bcb32ac27ded71c68>" +
                                    "<5181383ede94727bcb32ac27ded71c68>]>>\r\nstartxref\r\n{3}\r\n%%EOF\r\n"
                , xrefTable.XRefCount, refRoot, infoDict, objList.offset);

            xrefTable.ObjectByteOffsets = null;
            //PdfObject.NextObjectNum = 0;

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

        //private int StrLen(string text, int fontSize)
        //{
        //    char[] cArray = text.ToCharArray();
        //    int cWidth = 0;
        //    foreach (char c in cArray)
        //    {
        //        cWidth += 500; //(int)(fontSize*1.6)*20;	//Monospaced font width?
        //    }

        //    //div by1000??? 100 seems to work better :/
        //    return (cWidth / 100);
        //}

        ///// <summary>
        ///// start the Page Text element at the X Y position
        ///// </summary>
        ///// <returns></returns>
        //public string AddText(uint X, uint Y, string text, int fontSize, string fontName, Justification alignment)
        //{
        //    long startX = 0;
        //    switch (alignment)
        //    {
        //        case (Justification.Left):
        //            startX = X;
        //            break;
        //        case (Justification.Center):
        //            startX = X - (StrLen(text, fontSize)) / 2;
        //            break;
        //        case (Justification.Right):
        //            startX = X - StrLen(text, fontSize);
        //            break;
        //    }

        //    ;
        //    return $"\rBT/{fontName} {fontSize} Tf\r{startX} {(792 - Y)} Td \r({text}) Tj\rET\r";
        //}


        public void Write(string fileName)
        {
            FileStream file = new FileStream(fileName, FileMode.Create);

            int size = 0;
            file.Write(GetHeader("1.4", out size));

            //Just do first page for now...
            file.Write(pageObjects[0].RenderPageRefs(file.Length, out size), 0, size);
            file.Write(pageObjects[0].RenderPageContent(file.Length, out size), 0, size);

            file.Write(catalogObj.RenderBytes(file.Length, out size), 0, size);
            file.Write(pageTreeObj.RenderBytes(file.Length, out size), 0, size);

            file.Write(pageObjects[0].RenderPageFonts(file.Length,out size),0,size);

            file.Write(infoObject.RenderBytes(file.Length, out size), 0, size);
            file.Write(xrefTable.CreateXrefTable(file.Length, out size), 0, size);
            file.Write(GetTrailer(catalogObj.objectNum, infoObject.objectNum, out size), 0, size);
            file.Close();
        }
    }
}
