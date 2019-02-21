using System;
using System.Collections.Generic;
using System.Text;

namespace DocxToPdf.Core
{
    public class TextObject : IPdfRenderableObject
    {
        private readonly double _xPos;
        private readonly int _yPos;
        private readonly string _txt;
        private readonly FontObject _font;
        private readonly int _fontSize;
        private readonly string _alignment;
        private readonly PageExtents _extents;


        public TextObject(double xPos, int yPos, string txt, FontObject font, int fontSize, string alignment,
            ContentObject contentObj)
        {
            _xPos = xPos;
            _yPos = yPos;
            _txt = txt;
            _font = font;
            _fontSize = fontSize;
            _alignment = alignment;
            _extents = contentObj.ParentPage.PageDescription;
        }
        public string Render()
        {
            double startX = 0;
            switch (_alignment.ToLower())
            {
                case "clear":
                    startX = _xPos + _extents.leftMargin;
                    break;
                case "left":
                    startX = _xPos + _extents.leftMargin;
                    break;
                case "center":
                    startX = _xPos + _extents.leftMargin - (StrLen(_txt, _fontSize)) / 2;
                    break;
                case "right":
                    startX = _xPos + _extents.leftMargin - StrLen(_txt, _fontSize) + 2;           //+2 gives a little breathing space...
                    break;
            };
            return string.Format("\rBT/{0} {1} Tf\r{2} {3} Td \r({4}) Tj\rET\r",
                _font.FontRef(), _fontSize, startX, (720 - _yPos), _txt);
        }

        public byte[] RenderBytes(long offset, out int size)
        {
            size = 0;
            return new byte[size];
        }

        private int StrLen(string text, int fontSize)
        {
            char[] cArray = text.ToCharArray();
            int cWidth = 0;
            foreach (char c in cArray)
            {
                cWidth += 500;  //(int)(fontSize*1.6)*20;	//Monospaced font width?
            }
            //div by1000??? 100 seems to work better :/
            //$"{text} - {(cWidth / 100)}".Dump("StrLen Em's");
            return (cWidth / 100);
        }
    }
}
