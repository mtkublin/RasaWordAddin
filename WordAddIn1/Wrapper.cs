using System;
using System.Windows.Forms;
using XL.Office.Helpers;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1
{
    public partial class ThisAddIn
    {
        private void HighlightContentControl(Word.ContentControl control, ref bool cancel)
        {
            control.Range.Font.Color = Utilities.RGBwdColor(128, 0, 0);
            control.Range.Font.Shading.BackgroundPatternColor = Word.WdColor.wdColorLightYellow;
            control.Range.HighlightColorIndex = Word.WdColorIndex.wdNoHighlight;
        }

        public void WrapContent()
        {
            var activeDocument = Application.ActiveDocument;
            var extendedDocument = Globals.Factory.GetVstoObject(activeDocument);
            var next = DateTime.Now.Ticks.ToString();
            var control = extendedDocument.Controls.AddRichTextContentControl(string.Format("richText{0}", next));
            control.PlaceholderText = "This cannot be empty";
            control.Range.Font.Color = (Word.WdColor)128;
            control.Range.Font.Shading.BackgroundPatternColor = Word.WdColor.wdColorLightYellow;
            control.Range.HighlightColorIndex = Word.WdColorIndex.wdNoHighlight;
        }

        public void UnwrapContent()
        {
            var selection = Application.Selection;
            if (selection.ContentControls.Count > 0)
            {
                foreach (Word.ContentControl control in selection.ContentControls)
                {
                    control.Delete(false);
                }
            }
            if (selection.ParentContentControl != null)
            {
                selection.ParentContentControl.Delete(false);
            }
        }
    }
}
