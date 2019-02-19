﻿using System;
using System.Collections.Generic;
using System.Text;

namespace DocxToPdf.Core
{
    /// <summary>
    /// Conventience Methods
    /// </summary>
    public static class ContentObjectExtensions
    {
        public static void AddTextObject(this ContentObject contentObj, double xPos, int yPos, string txt, FontObject font, int fontSize,
            string alignment)
        {
            contentObj.AddObject(new TextObject(xPos, yPos, txt, font, fontSize, alignment));
        }
    }
}
