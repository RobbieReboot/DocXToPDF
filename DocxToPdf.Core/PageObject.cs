using System;
using System.Collections.Generic;
using System.Linq;

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
        private string resourceDict;
        private List<ContentObject> contentObjects;
        private List<FontObject> fontObjects;
        public PageObject()
        {
            resourceDict = null;
            pageSize = null;
            contentObjects = new List<ContentObject>();
            fontObjects = new List<FontObject>();
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

        
        public void AddFont(FontObject font) => fontObjects.Add(font);
        public void AddContent(ContentObject content) => contentObjects.Add(content);

        public byte[] RenderPageRefs(long filePos, out int size)
        {
            //render resource & content refs
            var fonts = fontObjects.Skip(1)
                .Aggregate(fontObjects.First().IndirectRef(), (acc, i) => acc + "\n" + i.IndirectRef());
            resourceDict = $"/Resources<</Font<<{fonts}>>/ProcSet[/PDF/Text]>>";

            var contents = contentObjects.Skip(1)
                .Aggregate(contentObjects.First().IndirectRef(), (acc, i) => acc + "\n" + i.IndirectRef());
            
            ObjectRepresenation += resourceDict + contents + ">>\rendobj\r";
            
            //render the actual page content.

            return this.GetUTF8Bytes(ObjectRepresenation, filePos, out size);
        }
        public byte[] RenderPageContent(long filePos, out int size)
        {
            //var contents = contentObjects.Skip(1)
            //    .Aggregate(contentObjects.First().GetContentDict(), (acc, i) => acc + "\n" + i.IndirectRef());

            //render the actual page content.

            foreach (var c in contentObjects)
            {
                
            }
            return contentObjects[0].RenderBytes(filePos, out size);
        }

        public byte[] RenderPageFonts(long filePos, out int size)
        {
            return fontObjects[0].RenderBytes(filePos, out size);
        }
    }
}