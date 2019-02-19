using System;

namespace DocxToPdf.Core
{
    /// <summary>
    /// Models the Catalog dictionary within a pdf file. This is the first created object. 
    /// It contains references to all other objects within the List of Pdf Objects.
    /// </summary>
    public class CatalogObject : PdfObject, IPdfRenderableObject
    {
        private readonly PageTreeObject _pageTreeObj;

        public CatalogObject(PageTreeObject pageTreeObj, PdfDocument pdfDocument) : base(pdfDocument)
        {
            _pageTreeObj = pageTreeObj;
        }
        
        public string Render()
        {
            return ObjectRepresenation = string.Format("{0} 0 obj <</Type /Catalog /Lang(EN-US) /Pages {1} 0 R>>\rendobj\r",
                this.PdfObjectId, _pageTreeObj.PdfObjectId);
        }

        public byte[] RenderBytes(long filePos, out int size)
        {
            Render();           //update representation
            return this.GetUTF8Bytes(ObjectRepresenation, filePos, out size);
        }
    }
}