using System;

namespace DocxToPdf.Core
{
    /// <summary>
    /// Models the Catalog dictionary within a pdf file. This is the first created object. 
    /// It contains references to all other objects within the List of Pdf Objects.
    /// </summary>
    public class CatalogObject : PdfObject
    {
        public CatalogObject()
        {
        }

        /// <summary>
        /// Returns the Catalog Dictionary 
        /// </summary>
        /// <param name="refPageTree"></param>
        /// <returns></returns>
        public byte[] GetCatalogDict(uint refPageTree, long filePos, out int size)
        {
            if (refPageTree < 1)
            {
                throw new Exception("GetCatalogDict - pagetree object number."); ;
            }

            ObjectRepresenation = string.Format("{0} 0 obj <</Type /Catalog /Lang(EN-US) /Pages {1} 0 R>>\rendobj\r",
                this.objectNum, refPageTree);
            return this.GetUTF8Bytes(ObjectRepresenation, filePos, out size);
        }
    }
}