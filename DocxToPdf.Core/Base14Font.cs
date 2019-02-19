namespace DocxToPdf.Core
{
    ///<summary>
    /// All the Base14 font names from the Acrobat spec in the ISO PDF Standard, ISO 32000-1:2008(E),
    /// the Base 14 Fonts are referred to as the “Standard 14 Fonts”; and are defined in Section 9.6.2.2
    ///
    /// https://en.wikipedia.org/wiki/PDF#Standard_Type_1_Fonts_(Standard_14_Fonts)
    /// 
    /// Courier, Courier Bold, Courier Oblique, Courier Bold-Oblique
    /// Helvetica, Helvetica Bold, Helvetica Oblique, Helvetica Bold-Oblique
    /// Times Roman, Times Bold, Times Italic, Times Bold-Italic
    /// Symbol
    /// Zapf Dingbats
    ///</summary>
    public class Base14Font
    {
        public const string Courier             = "Courier";
        public const string CourierBold         = "Courier Bold";
        public const string CourierOblique      = "Courier Oblique";
        public const string CourierBoldOblique  = "Courier Bold-Oblique";
        public const string Helvetica           = "Helvetica";
        public const string HelveticaBold       = "Helvetica Bold";
        public const string HelveticaOblique    = "Helvetica Oblique";
        public const string TimesRoman          = "Times Roman";
        public const string TimesBold           = "Times Bold";
        public const string TimesItalic         = "Times Italic";
        public const string TimesBoldItalic     = "Times Bold-Italic";
        public const string Symbol              = "Symbol";
        public const string ZapfDingbats        = "Zapf Dingbats";
    }
}