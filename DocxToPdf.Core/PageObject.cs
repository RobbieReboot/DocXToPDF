using System;

namespace DocxToPdf.Core
{
    /// <summary>
    /// This class represents individual pages within the pdf. 
    /// The contents of the page belong to this class
    /// </summary>
    public class PageObject : PdfObject
    {
        //private string page;
        private string pageSize;
        private string fontRef;
        private string resourceDict, contents;

        public PageObject()
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
        public void CreatePage(uint refParent, PageDescription pSize)
        {
            pageSize = $"[0 0 {pSize.xWidth} {pSize.yHeight}]";
            ObjectRepresenation = string.Format("{0} 0 obj <</Type /Page/Parent {1} 0 R/Rotate 0/MediaBox {2}/CropBox {2}",
                this.objectNum, refParent, pageSize);
        }

        /// <summary>
        /// Add Resource to the pdf page
        /// </summary>
        /// <param name="font"></param>
        public void AddResource(FontObject font, uint contentRef)
        {
            fontRef += $"/{font.font} {font.objectNum} 0 R";
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
            ObjectRepresenation += resourceDict + contents + ">>\rendobj\r";
            return this.GetUTF8Bytes(ObjectRepresenation, filePos, out size);
        }
    }
}