namespace DocxToPdf.Core
{
    ///<summary>
    /// Represents the font dictionary used in a pdf page
    /// Times-Roman		Helvetica				CourierNew
    /// Times-Bold		Helvetica-Bold			Courier-Bold
    /// Times-Italic		Helvetica-Oblique		Courier-Oblique
    /// Times-BoldItalic Helvetica-BoldOblique	Courier-BoldOblique
    ///</summary>
    
    public class FontObject : PdfObject, IPdfRenderableObject
    {
        //private string fontObject;
        public string font;                 //the PDF object tag
        private static int fontIndex = 1;   //numerical Object tag..

        private string fontType;            //the font name, eg, bit14 fonts like CourierNew
        public FontObject(string fType)
        {
            font = null;
            font = $"T{fontIndex}";
            fontType = fType;

            fontIndex++;
        }

        
        public string IndirectRef() => $"/{font} {objectNum} 0 R";

        public string Render()
        {
            return $"{this.objectNum} 0 obj <</Type/Font/Name /{font}/BaseFont/{fontType}/Subtype/Type1/Encoding /WinAnsiEncoding>>\nendobj\n";
        }

        public byte[] RenderBytes(long filePos, out int size)
        {
            return this.GetUTF8Bytes(ObjectRepresenation, filePos, out size);
        }
    }
}