using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Design;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Media;
using FontFamily = System.Drawing.FontFamily;

//https://docs.microsoft.com/en-us/dotnet/framework/winforms/advanced/how-to-obtain-font-metrics
//https://docs.microsoft.com/en-us/windows/desktop/gdi/character-widths


namespace GlyphTableGenerator
{
    class Program
    {
        static void Main(string[] args)
        {
            //ShowFontDetail();

            //Get Character widths 

            var allFonts = Fonts.GetFontFamilies(@"C:\windows\fonts\Arial.ttf");
            var requiredFonts = new[] {"Arial", "Consolas", "Times New Roman"};

            foreach (var ff in allFonts)
            {
                foreach (var t in ff.GetTypefaces())
                {
                    Console.WriteLine(t.Style);
                    if (t.TryGetGlyphTypeface(out GlyphTypeface gt))
                    {
                        var charToGlyphMap = gt.CharacterToGlyphMap;
                        var c = charToGlyphMap.Count;
                        var asciiGlyphMap = gt.CharacterToGlyphMap.Where(ctgm => ctgm.Key < 255).ToList();
                        foreach (var ctg in gt.CharacterToGlyphMap)
                        {
                            var width = (int)Math.Round(gt.AdvanceWidths[ctg.Value] * 1000);
                            FontFamily fontFamily = new FontFamily("Arial");
                            Font font = new Font(
                                fontFamily,
                                16, FontStyle.Regular,
                                GraphicsUnit.Pixel);

                            var sz = MeasureCharacter((char)ctg.Key, font);
                            Console.WriteLine($"{(char)ctg.Key} ({ctg.Key}) Width = {width}");
                        }
                    }
                }
            }


            Console.ReadKey();
        }

        static Size MeasureCharacter(char character, Font font)
        {
            int height;
            int width;

            using (var graphics = Graphics.FromImage(new Bitmap(1, 1)))
            {
                //Calculate width & height required to fit the text into the image
                //based on the font selected
                var ms = graphics.MeasureString(character.ToString(), font);

                width = (int)ms.Width;
                height = (int)ms.Height;

            }

            return new Size(width, height);
        }

        //private static void ShowFontDetail()
        //{
        //    string infoString = ""; // enough space for one line of output
        //    int ascent; // font family ascent in design units
        //    float ascentPixel; // ascent converted to pixels
        //    int descent; // font family descent in design units
        //    float descentPixel; // descent converted to pixels
        //    int lineSpacing; // font family line spacing in design units
        //    float lineSpacingPixel; // line spacing converted to pixels

        //    FontFamily fontFamily = new FontFamily("Arial");
        //    Font font = new Font(
        //        fontFamily,
        //        16, FontStyle.Regular,
        //        GraphicsUnit.Pixel);
        //    PointF pointF = new PointF(10, 10);
        //    SolidBrush solidBrush = new SolidBrush(Color.Black);

        //    // Display the font size in pixels.
        //    Console.WriteLine($"font.Size returns {font.Size}.");

        //    // Move down one line.
        //    pointF.Y += font.Height;

        //    // Display the font family em height in design units.
        //    Console.WriteLine("fontFamily.GetEmHeight() returns " +
        //                      fontFamily.GetEmHeight(FontStyle.Regular) + ".");

        //    // Move down two lines.
        //    pointF.Y += 2 * font.Height;

        //    // Display the ascent in design units and pixels.
        //    ascent = fontFamily.GetCellAscent(FontStyle.Regular);

        //    // 14.484375 = 16.0 * 1854 / 2048
        //    ascentPixel =
        //        font.Size * ascent / fontFamily.GetEmHeight(FontStyle.Regular);
        //    Console.WriteLine("The ascent is " + ascent + " design units, " + ascentPixel +
        //                      " pixels.");

        //    // Move down one line.
        //    pointF.Y += font.Height;

        //    // Display the descent in design units and pixels.
        //    descent = fontFamily.GetCellDescent(FontStyle.Regular);

        //    // 3.390625 = 16.0 * 434 / 2048
        //    descentPixel =
        //        font.Size * descent / fontFamily.GetEmHeight(FontStyle.Regular);
        //    Console.WriteLine("The descent is " + descent + " design units, " +
        //                      descentPixel + " pixels.");

        //    // Move down one line.
        //    pointF.Y += font.Height;

        //    // Display the line spacing in design units and pixels.
        //    lineSpacing = fontFamily.GetLineSpacing(FontStyle.Regular);

        //    // 18.398438 = 16.0 * 2355 / 2048
        //    lineSpacingPixel =
        //        font.Size * lineSpacing / fontFamily.GetEmHeight(FontStyle.Regular);
        //    Console.WriteLine("The line spacing is " + lineSpacing + " design units, " +
        //                      lineSpacingPixel + " pixels.");
        //}
    }
}
