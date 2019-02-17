using System;
using System.Collections;
using System.Collections.Generic;
using System.Text;

namespace DocxToPdf.Core
{

    /// <summary>
    /// Specify the page size in 1/72 inches units.
    /// </summary>
    public struct PageSize
    {
        public uint xWidth;
        public uint yHeight;
        public uint leftMargin;
        public uint rightMargin;
        public uint topMargin;
        public uint bottomMargin;

        public PageSize(uint width, uint height)
        {
            xWidth = width;
            yHeight = height;
            leftMargin = 0;
            rightMargin = 0;
            topMargin = 0;
            bottomMargin = 0;
        }
        public void SetMargins(uint L, uint T, uint R, uint B)
        {
            leftMargin = L;
            rightMargin = R;
            topMargin = T;
            bottomMargin = B;
        }
    }

    /// <summary>
    /// Used with the Text inside a Table
    /// </summary>
    public enum Align
    {
        LeftAlign, CenterAlign, RightAlign
    }

    /// <summary>
    /// Holds the Byte offsets of the objects used in the Pdf Document
    /// </summary>
    internal class XrefEnteries
    {
        internal static ArrayList offsetArray;

        internal XrefEnteries()
        {
            offsetArray = new ArrayList();
        }
    }

    /// <summary>
    /// For Adding the Object number and file offset
    /// </summary>
    public class ObjectList : IComparable
    {
        internal long offset;
        internal uint objNum;

        internal ObjectList(uint objectNum, long fileOffset)
        {
            offset = fileOffset;
            objNum = objectNum;
        }
        #region IComparable Members

        public int CompareTo(object obj)
        {

            int result = 0;
            result = (this.objNum.CompareTo(((ObjectList)obj).objNum));
            return result;
        }

        #endregion
    }


    public class PdfObject
    {
        /// <summary>
        /// Stores the Object Number
        /// </summary>
        internal static uint inUseObjectNum;
        public uint objectNum;
        //private UTF8Encoding utf8;
        private XrefEnteries Xref;
        /// <summary>
        /// Constructor increments the object number to 
        /// reflect the currently used object number
        /// </summary>
        protected PdfObject()
        {
            if (inUseObjectNum == 0)
                Xref = new XrefEnteries();
            inUseObjectNum++;
            objectNum = inUseObjectNum;
        }
        ~PdfObject()
        {
            objectNum = 0;
        }
        /// <summary>
        /// Convert the unicode string 16 bits to unicode bytes. 
        /// This is written to the file to create Pdf 
        /// </summary>
        /// <param name="str"></param>
        /// <returns></returns>
        protected byte[] GetUTF8Bytes(string str, long filePos, out int size)
        {
            ObjectList objList = new ObjectList(objectNum, filePos);
            byte[] abuf;
            try
            {
                byte[] ubuf = Encoding.Unicode.GetBytes(str);
                Encoding enc = Encoding.Default;    //GetEncoding(1252);
                abuf = Encoding.Convert(Encoding.Unicode, enc, ubuf);
                size = abuf.Length;
                XrefEnteries.offsetArray.Add(objList);
            }
            catch (Exception e)
            {
                string str1 = string.Format("{0},In PdfObjects.GetBytes()", objectNum);
                Exception error = new Exception(e.Message + str1);
                throw error;
            }
            return abuf;
        }

    }
    /// <summary>
    /// Models the Catalog dictionary within a pdf file. This is the first created object. 
    /// It contains referencesto all other objects within the List of Pdf Objects.
    /// </summary>
    public class CatalogDict : PdfObject
    {
        private string catalog;
        public CatalogDict()
        {

        }
        /// <summary>
        ///Returns the Catalog Dictionary 
        /// </summary>
        /// <param name="refPageTree"></param>
        /// <returns></returns>
        public byte[] GetCatalogDict(uint refPageTree, long filePos, out int size)
        {
            Exception error = new Exception(" In CatalogDict.GetCatalogDict(), PageTree.objectNum Invalid");
            if (refPageTree < 1)
            {
                throw error;
            }
            catalog = string.Format("{0} 0 obj<</Type /Catalog/Lang(EN-US)/Pages {1} 0 R>>\rendobj\r",
                this.objectNum, refPageTree);
            return this.GetUTF8Bytes(catalog, filePos, out size);
        }

    }
    /// <summary>
    /// The PageTree object contains references to all the pages used within the Pdf.
    /// All individual pages are referenced through the Kids arraylist
    /// </summary>
    public class PageTreeDict : PdfObject
    {
        private string pageTree;
        private string kids;
        private static uint MaxPages;

        public PageTreeDict()
        {
            kids = "[ ";
            MaxPages = 0;
        }
        /// <summary>
        /// Add a page to the Page Tree. ObjNum is the object number of the page to be added.
        /// pageNum is the page number of the page.
        /// </summary>
        /// <param name="objNum"></param>
        /// <param name="pageNum"></param>
        public void AddPage(uint objNum)
        {
            Exception error = new Exception("In PageTreeDict.AddPage, PageDict.ObjectNum Invalid");
            if (objNum < 0 || objNum > PdfObject.inUseObjectNum)
                throw error;
            MaxPages++;
            string refPage = objNum + " 0 R ";
            kids = kids + refPage;
        }
        /// <summary>
        /// returns the Page Tree Dictionary
        /// </summary>
        /// <returns></returns>
        public byte[] GetPageTree(long filePos, out int size)
        {
            pageTree = string.Format("{0} 0 obj<</Count {1}/Kids {2}]>>\rendobj\r",
                this.objectNum, MaxPages, kids);
            return this.GetUTF8Bytes(pageTree, filePos, out size);
        }
    }
    /// <summary>
    /// This class represents individual pages within the pdf. 
    /// The contents of the page belong to this class
    /// </summary>
    public class PageDict : PdfObject
    {
        private string page;
        private string pageSize;
        private string fontRef;
        private string resourceDict, contents;
        public PageDict()
        {
            resourceDict = null;
            contents = null;
            pageSize = null;
            fontRef = null;
        }
        /// <summary>
        /// Create The Pdf page
        /// </summary>
        /// <param name="refParent"></param>
        /// <param name="width"></param>
        /// <param name="height"></param>
        public void CreatePage(uint refParent, PageSize pSize)
        {
            Exception error = new Exception("In PageDict.CreatePage(),PageTree.ObjectNum Invalid");
            if (refParent < 1 || refParent > PdfObject.inUseObjectNum)
                throw error;

            pageSize = string.Format("[0 0 {0} {1}]", pSize.xWidth, pSize.yHeight);
            page = string.Format("{0} 0 obj<</Type /Page/Parent {1} 0 R/Rotate 0/MediaBox {2}/CropBox {2}",
                this.objectNum, refParent, pageSize);
        }
        /// <summary>
        /// Add Resource to the pdf page
        /// </summary>
        /// <param name="font"></param>
        public void AddResource(FontDict font, uint contentRef)
        {
            fontRef += string.Format("/{0} {1} 0 R", font.font, font.objectNum);
            if (contentRef > 0)
            {
                contents = string.Format("/Contents {0} 0 R", contentRef);
            }
        }
        /// <summary>
        /// Get the Page Dictionary to be written to the file
        /// </summary>
        /// <param name="size"></param>
        /// <returns></returns>
        public byte[] GetPageDict(long filePos, out int size)
        {
            resourceDict = string.Format("/Resources<</Font<<{0}>>/ProcSet[/PDF/Text]>>", fontRef);
            page += resourceDict + contents + ">>\rendobj\r";
            return this.GetUTF8Bytes(page, filePos, out size);
        }
    }
    /// <summary>
    ///Represents the general content stream in a Pdf Page. 
    ///This is used only by the PageObjec 
    /// </summary>
    public class ContentDict : PdfObject
    {
        private string contentDict;
        private string contentStream;
        public ContentDict()
        {
            contentDict = null;
            contentStream = "%stream\r";
        }
        /// <summary>
        /// Set the Stream of this Content Dict.
        /// Stream is taken from TextAndTable Objects
        /// </summary>
        /// <param name="stream"></param>
        public void SetStream(string stream)
        {
            contentStream += stream;
        }
        /// <summary>
        /// Enter the text inside the table just created.
        /// </summary>
        /// <summary>
        /// Get the Content Dictionary
        /// </summary>
        public byte[] GetContentDict(long filePos, out int size)
        {
            contentDict = string.Format("{0} 0 obj<</Length {1}>>stream\r\n{2}\nendstream\rendobj\r",
                this.objectNum, contentStream.Length, contentStream);

            return GetUTF8Bytes(contentDict, filePos, out size);

        }

    }
    /// <summary>
    ///Represents the font dictionary used in a pdf page
    ///Times-Roman		Helvetica				Courier
    ///Times-Bold		Helvetica-Bold			Courier-Bold
    ///Times-Italic		Helvetica-Oblique		Courier-Oblique
    ///Times-BoldItalic Helvetica-BoldOblique	Courier-BoldOblique
    /// </summary>
    public class FontDict : PdfObject
    {
        private string fontDict;
        public string font;
        public FontDict()
        {
            font = null;
            fontDict = null;
        }
        /// <summary>
        /// Create the font Dictionary
        /// </summary>
        /// <param name="fontName"></param>
        public void CreateFontDict(string fontName, string fontType)
        {
            font = fontName;
            fontDict = string.Format("{0} 0 obj<</Type/Font/Name /{1}/BaseFont/{2}/Subtype/Type1/Encoding /WinAnsiEncoding>>\nendobj\n",
                this.objectNum, fontName, fontType);
        }
        /// <summary>
        /// Get the font Dictionary to be written to the file
        /// </summary>
        /// <param name="size"></param>
        /// <returns></returns>
        public byte[] GetFontDict(long filePos, out int size)
        {
            return this.GetUTF8Bytes(fontDict, filePos, out size);
        }

    }
    /// <summary>
    ///Store information about the document,Title, Author, Company, 
    /// </summary>
    public class InfoDict : PdfObject
    {
        private string info;
        public InfoDict()
        {
            info = null;
        }
        /// <summary>
        /// Fill the Info Dict
        /// </summary>
        /// <param name="title"></param>
        /// <param name="author"></param>
        public void SetInfo(string title, string author, string company)
        {
            info = string.Format("{0} 0 obj<</ModDate({1})/CreationDate({1})/Title({2})/Creator(CRM FactFind)" +
                "/Author({3})/Producer(CRM FactFind)/Company({4})>>\rendobj\r",
                this.objectNum, GetDateTime(), title, author, company);

        }
        /// <summary>
        /// Get the Document Information Dictionary
        /// </summary>
        /// <param name="size"></param>
        /// <returns></returns>
        public byte[] GetInfoDict(long filePos, out int size)
        {
            return GetUTF8Bytes(info, filePos, out size);
        }
        /// <summary>
        /// Get Date as Adobe needs ie similar to ISO/IEC 8824 format
        /// </summary>
        /// <returns></returns>
        private string GetDateTime()
        {
            DateTime universalDate = DateTime.UtcNow;
            DateTime localDate = DateTime.Now;
            string pdfDate = string.Format("D:{0:yyyyMMddhhmmss}", localDate);
            TimeSpan diff = localDate.Subtract(universalDate);
            int uHour = diff.Hours;
            int uMinute = diff.Minutes;
            char sign = '+';
            if (uHour < 0)
                sign = '-';
            uHour = Math.Abs(uHour);
            pdfDate += string.Format("{0}{1}'{2}'", sign, uHour.ToString().PadLeft(2, '0'), uMinute.ToString().PadLeft(2, '0'));
            return pdfDate;
        }

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
                ObjectList objList = new ObjectList(0, fileOffset);
                XrefEnteries.offsetArray.Add(objList);
                XrefEnteries.offsetArray.Sort();
                numTableEntries = (uint)XrefEnteries.offsetArray.Count;
                table = string.Format("xref\r\n{0} {1}\r\n0000000000 65535 f\r\n", 0, numTableEntries);
                for (int entries = 1; entries < numTableEntries; entries++)
                {
                    ObjectList obj = (ObjectList)XrefEnteries.offsetArray[entries];
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
        /// <summary>
        /// Returns the Header
        /// </summary>
        /// <param name="version"></param>
        /// <param name="size"></param>
        /// <returns></returns>
        public byte[] GetHeader(string version, out int size)
        {
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
                ObjectList objList = (ObjectList)XrefEnteries.offsetArray[0];
                trailer = string.Format("trailer\n<</Size {0}/Root {1} 0 R {2}/ID[<5181383ede94727bcb32ac27ded71c68>" +
                    "<5181383ede94727bcb32ac27ded71c68>]>>\r\nstartxref\r\n{3}\r\n%%EOF\r\n"
                    , numTableEntries, refRoot, infoDict, objList.offset);

                XrefEnteries.offsetArray = null;
                PdfObject.inUseObjectNum = 0;
            }
            catch (Exception e)
            {
                Exception error = new Exception(e.Message + " In Utility.GetTrailer()");
                throw error;
            }

            return GetUTF8Bytes(trailer, out size);
        }
        /// <summary>
        /// Converts the string to byte array in utf 8 encoding
        /// </summary>
        /// <param name="str"></param>
        /// <param name="size"></param>
        /// <returns></returns>
        private byte[] GetUTF8Bytes(string str, out int size)
        {
            try
            {
                byte[] ubuf = Encoding.Unicode.GetBytes(str);
                Encoding enc = Encoding.Default;    // Encoding.GetEncoding(1252);
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
                cWidth += 500;  //(int)(fontSize*1.6)*20;	//Monospaced font width?
            }
            //div by1000??? 100 seems to work better :/
            return (cWidth / 100);
        }

        /// <summary>
        /// start the Page Text element at the X Y position
        /// </summary>
        /// <returns></returns>
        public string AddText(uint X, uint Y, string text, int fontSize, string fontName, Align alignment   )
        {
            long startX = 0;
            switch (alignment)
            {
                case (Align.LeftAlign):
                    startX = X;
                    break;
                case (Align.CenterAlign):
                    startX = X - (StrLen(text, fontSize)) / 2;
                    break;
                case (Align.RightAlign):
                    startX = X - StrLen(text, fontSize);
                    break;
            };
            return $"\rBT/{fontName} {fontSize} Tf\r{startX} {(792 - Y)} Td \r({text}) Tj\rET\r";
        }
    }


}
