using System;

namespace DocxToPdf.Core
{
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
            if (objNum < 0 || objNum > PdfObject.NextObjectNum)
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
            pageTree = string.Format("{0} 0 obj <</Count {1}/Kids {2}]>>\rendobj\r",
                this.objectNum, MaxPages, kids);
            return this.GetUTF8Bytes(pageTree, filePos, out size);
        }
    }
}