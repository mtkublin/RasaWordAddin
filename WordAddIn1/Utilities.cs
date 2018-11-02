using Word = Microsoft.Office.Interop.Word;

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

        public static void Notification(string message)
        {
            System.Console.WriteLine(message);
        }

    }
}
