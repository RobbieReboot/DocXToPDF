namespace DocxToPdf.Core
{
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
            contentDict = string.Format("{0} 0 obj <</Length {1}>>stream\r\n{2}\nendstream\rendobj\r",
                this.objectNum, contentStream.Length, contentStream);

            return GetUTF8Bytes(contentDict, filePos, out size);
        }
    }
}