using System;

namespace DocxToPdf.Core
{
    /// <summary>
    ///Store information about the document,Title, Author, Company, 
    /// </summary>
    public class InfoObject : PdfObject,IPdfRenderableObject
    {
        private readonly string _title;
        private readonly string _author;
        private readonly string _company;

        public InfoObject(string title,string author,string company,PdfDocument pdfDocument) : base(pdfDocument)
        {
            _title = title;
            _author = author;
            _company = company;
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
            pdfDate += string.Format("{0}{1}'{2}'", sign, uHour.ToString().PadLeft(2, '0'),
                uMinute.ToString().PadLeft(2, '0'));
            return pdfDate;
        }

        public string Render()
        {
            return ObjectRepresenation = string.Format("{0} 0 obj <</ModDate({1}) /CreationDate({1}) /Title({2}) /Creator({4}) " +
                                                "/Author({3}) /Producer(3Squared) /Company({4})>>\nendobj\n",
                this.PdfObjectId, GetDateTime(), _title, _author, _company);
        }

        public byte[] RenderBytes(long filePos, out int size)
        {
            Render();
            return GetUTF8Bytes(ObjectRepresenation, filePos, out size);
        }
    }
}