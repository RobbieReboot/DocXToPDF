namespace DocxToPdf.Core
{
    ///<summary>
    /// All the Base14 font names from the Acrobat spec in the ISO PDF Standard, ISO 32000-1:2008(E),
    /// the Base 14 Fonts are referred to as the “Standard 14 Fonts”; and are defined in Section 9.6.2.2
    ///
    /// Courier, Courier Bold, Courier Oblique, Courier Bold-Oblique
    /// Helvetica, Helvetica Bold, Helvetica Oblique, Helvetica Bold-Oblique
    /// Times Roman, Times Bold, Times Italic, Times Bold-Italic
    /// Symbol
    /// Zapf Dingbats
    ///</summary>
    public class FontObject : PdfObject, IPdfRenderableObject
    {
        //private string fontObject;
        private string fontRef;                 //the PDF object tag
        private static int fontIndex = 1;   //numerical Object tag..

        public string FontRef() => fontRef;

        private string fontType;            //the font name, eg, bit14 fonts like CourierNew
        public FontObject(string fType)
        {
            fontRef = null;
            fontRef = $"T{fontIndex}";
            fontType = fType;

            fontIndex++;
        }

        
        public string IndirectRef() => $"/{fontRef} {objectNum} 0 R";

        public string Render()
        {
            return ObjectRepresenation = $"{this.objectNum} 0 obj <</Type/Font/Name /{fontRef}/BaseFont/{fontType}/Subtype/Type1/Encoding /WinAnsiEncoding>>\nendobj\n";
        }

        public byte[] RenderBytes(long filePos, out int size)
        {
            Render();
            return this.GetUTF8Bytes(ObjectRepresenation, filePos, out size);
        }
    }
}