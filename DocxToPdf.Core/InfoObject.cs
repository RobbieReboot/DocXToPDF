using System;

namespace DocxToPdf.Core
{
    /// <summary>
    ///Store information about the document,Title, Author, Company, 
    /// </summary>
    public class InfoObject : PdfObject
    {
        public InfoObject()
        {
        }

        /// <summary>
        /// Fill the Info Dict
        /// </summary>
        /// <param name="title"></param>
        /// <param name="author"></param>
        public void SetInfo(string title, string author, string company)
        {
            ObjectRepresenation = string.Format("{0} 0 obj <</ModDate({1}) /CreationDate({1}) /Title({2}) /Creator(3Squared) " +
                                 "/Author({3}) /Producer(3Squared) /Company({4})>>\nendobj\n",
                this.objectNum, GetDateTime(), title, author, company);
        }

        /// <summary>
        /// Get the Document Information Dictionary
        /// </summary>
        /// <param name="size"></param>
        /// <returns></returns>
        public byte[] GetInfoDict(long filePos, out int size)
        {
            return GetUTF8Bytes(ObjectRepresenation, filePos, out size);
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
    }
}