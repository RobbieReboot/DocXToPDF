namespace DocxToPdf.Core
{
    /// <summary>
    ///Represents the font dictionary used in a pdf page
    ///Times-Roman		Helvetica				Courier
    ///Times-Bold		Helvetica-Bold			Courier-Bold
    ///Times-Italic		Helvetica-Oblique		Courier-Oblique
    ///Times-BoldItalic Helvetica-BoldOblique	Courier-BoldOblique
    /// </summary>
    public class FontDict : PdfObject
    {
        private string fontDict;
        public string font;

        public FontDict()
        {
            font = null;
            fontDict = null;
        }

        /// <summary>
        /// Create the font Dictionary
        /// </summary>
        /// <param name="fontName"></param>
        public void CreateFont(string fontName, string fontType)
        {
            font = fontName;
            fontDict = string.Format(
                "{0} 0 obj <</Type/Font/Name /{1}/BaseFont/{2}/Subtype/Type1/Encoding /WinAnsiEncoding>>\nendobj\n",
                this.objectNum, fontName, fontType);
        }

        /// <summary>
        /// Get the font Dictionary to be written to the file
        /// </summary>
        /// <param name="size"></param>
        /// <returns></returns>
        public byte[] GetFontDict(long filePos, out int size)
        {
            return this.GetUTF8Bytes(fontDict, filePos, out size);
        }
    }
}