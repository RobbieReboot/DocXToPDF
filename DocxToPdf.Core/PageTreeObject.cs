using System;

namespace DocxToPdf.Core
{
    /// <summary>
    /// The PageTree object contains references to all the pages used within the Pdf.
    /// All individual pages are referenced through the Kids arraylist
    /// </summary>
    public class PageTreeObject : PdfObject
    {
        private string kids;
        private static uint MaxPages;

        public PageTreeObject()
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
        public void AddPage(PageObject page)
        {
            var objectNum = page.objectNum;

            Exception error = new Exception("In PageTreeDict.AddPage, PageDict.ObjectNum Invalid");
            if (objectNum > PdfObject.NextObjectNum)
                throw error;

            MaxPages++;
            string refPage = objectNum + " 0 R ";
            kids = kids + refPage;
        }
        
        /// <summary>
        /// returns the Page Tree Dictionary
        /// </summary>
        /// <returns></returns>
        public byte[] GetPageTree(long filePos, out int size)
        {
            ObjectRepresenation = $"{this.objectNum} 0 obj <</Count {MaxPages}/Kids {kids}]>>\rendobj\r";
            return this.GetUTF8Bytes(ObjectRepresenation, filePos, out size);
        }
    }
}