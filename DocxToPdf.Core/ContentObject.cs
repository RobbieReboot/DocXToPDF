namespace DocxToPdf.Core
{
    /// <summary>
    /// Represents the general content stream in a Pdf Page. 
    /// This is used only by the PageObjec 
    /// </summary>
    public class ContentObject : PdfObject
    {
        //private string contentStream;
        private readonly PageObject _parentPage;
        public PageObject ParentPage() => _parentPage;

        public ContentObject()
        {
            //contentDict = null;
            ObjectRepresenation = "%stream\r";
        }
        public ContentObject(PageObject parent) : this()
        {
            _parentPage = parent;
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

        public string IndirectRef() => $"/Contents {objectNum} 0 R";

        /// <summary>
        /// Enter the text inside the table just created.
        /// </summary>
        /// <summary>
        /// Get the Content Dictionary
        /// </summary>
        public byte[] RenderBytes(long filePos, out int size)
        {
            ObjectRepresenation = $"{this.objectNum} 0 obj <</Length {ObjectRepresenation.Length}>>stream\r\n{ObjectRepresenation}\nendstream\rendobj\r";
            return GetUTF8Bytes(ObjectRepresenation, filePos, out size);
        }
    }
}