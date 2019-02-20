namespace DocxToPdf.Core
{
    /// <summary>
    /// Represents the general content stream in a Pdf Page. 
    /// This is used only by the PageObjec 
    /// </summary>
    public class ContentObject : PdfObject
    {
        //private string contentStream;
        public PageObject ParentPage { get; }

        public ContentObject(PdfDocument pdfDocument,PageObject parent) : base (pdfDocument)
        {
            //contentDict = null;
            ObjectRepresenation = "%stream\r";
            ParentPage = parent;
        }

        /// <summary>
        /// Set the Stream of this Content Dict.
        /// Stream is taken from TextAndTable Objects
        /// </summary>
        /// <param name="stream"></param>
        public void AddObject(IPdfRenderableObject obj)
        {
            ObjectRepresenation += obj.Render();
        }

        public string IndirectRef() => $"/Contents {PdfObjectId} 0 R";

        /// <summary>
        /// Enter the text inside the table just created.
        /// </summary>
        /// <summary>
        /// Get the Content Dictionary
        /// </summary>
        public byte[] RenderBytes(long filePos, out int size)
        {
            ObjectRepresenation = $"{this.PdfObjectId} 0 obj <</Length {ObjectRepresenation.Length}>>stream\r\n{ObjectRepresenation}\nendstream\rendobj\r";
            return GetUTF8Bytes(ObjectRepresenation, filePos, out size);
        }
    }
}