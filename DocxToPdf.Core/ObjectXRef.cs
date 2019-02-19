using System;

namespace DocxToPdf.Core
{
    /// <summary>
    /// Represents the Object Id and the byte offset in the file.
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
            result = (this.objNum.CompareTo(((ObjectXRef) obj).objNum));
            return result;
        }

        #endregion
    }
}