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