using Word = Microsoft.Office.Interop.Word;
using Draw = System.Drawing;
using System.Collections.Generic;

namespace XL.Office.Helpers
{
    class Utilities
    {
        public static Word.WdColor RGBwdColor(uint red, uint green, uint blue)
        {
            uint clip = 255;

            red = red > clip ? clip : red;
            green = green > clip ? clip : green;
            blue = blue > clip ? clip : blue;

            return (Word.WdColor)(red + 0x100 * green + 0x10000 * blue);
        }
        public static Word.WdColor RGBwdColor(Draw.Color color)
        {
            return RGBwdColor(color.R, color.G, color.B);
        }

        public static IEnumerator<Draw.Color> Gradient(Draw.Color start, Draw.Color end, int steps)
        {
            int stepA = ((end.A - start.A) / (steps - 1));
            int stepR = ((end.R - start.R) / (steps - 1));
            int stepG = ((end.G - start.G) / (steps - 1));
            int stepB = ((end.B - start.B) / (steps - 1));

            for (int i = 0; i < steps; i++)
            {
                yield return Draw.Color.FromArgb(
                    start.A + (stepA * i),
                    start.R + (stepR * i),
                    start.G + (stepG * i),
                    start.B + (stepB * i)
                );
            }
        }

        public static Draw.Color Contrast(Draw.Color color)
        {
            int d = 0;

            // Counting the perceptive luminance - human eye favors green color... 
            double luminance = (0.299 * color.R + 0.587 * color.G + 0.114 * color.B) / 255;

            if (luminance > 0.5)
                d = 0; // bright colors - black font
            else
                d = 255; // dark colors - white font

            return Draw.Color.FromArgb(d, d, d);
        }

        public static void Notification(string message)
        {
            System.Console.WriteLine(message);
        }

    }
}
