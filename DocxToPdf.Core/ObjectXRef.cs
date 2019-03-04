using System;

namespace DocxToPdf.Core
{
    /// <summary>
    /// Represents the ObjectId and the byte offset in the file - must be sortable because the XRef table must be in OBJECT order!
    /// </summary>
    public class ObjectXRef : IComparable
    {
        internal long offset;
        internal uint objNum;

        internal ObjectXRef(uint objectNum, long fileOffset)
        {
            offset = fileOffset;
            objNum = objectNum;
        }

        #region IComparable Members

        public int CompareTo(object obj)
        {
            int result = 0;
            result = (this.objNum.CompareTo(((ObjectXRef)obj).objNum));
            return result;
        }

        #endregion
    }
}