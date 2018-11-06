using System;
using System.Windows.Forms;
using XL.Office.Helpers;
using Word = Microsoft.Office.Interop.Word;
using System.Drawing; 

namespace WordAddIn1
{
    public partial class ThisAddIn
    {
        private void HighlightControlHierarchy(Word.ContentControl control, ref bool cancel)
        {
            try
            {
                var parent = control.Range.ParentContentControl;
                HighlightContentControl(parent.Tag, parent.Range);
                foreach (Word.ContentControl child in control.Range.ContentControls)
                {
                    HighlightContentControl(child.Tag, child.Range);
                }
            }
            catch (Exception ex)
            {
                Utilities.Notification(ex.ToString());
            }
        }

        private void HighlightContentControl(string tag, Word.Range range)
        {
            //do not wrap if tag is empty or null
            if (string.IsNullOrEmpty(tag)) return;

            range.Font.Color = TagForeColor(tag);
            range.Font.Shading.ForegroundPatternColor = TagBackColor(tag);
            range.Font.Shading.Texture = Word.WdTextureIndex.wdTextureSolid;
            range.HighlightColorIndex = Word.WdColorIndex.wdNoHighlight;
        }

        public void WrapContent()
        {
            //do not wrap if current tag is empty or null
            if (string.IsNullOrEmpty(CurrentTag)) return;

            var selection = Application.Selection;
            //do not wrap if range is collapsed
            if (selection.Start == selection.End) return;

            //do not allow wrapping part of another control
            int start = selection.Start;
            int end = selection.End;
            selection.Collapse(Word.WdCollapseDirection.wdCollapseStart);
            Word.ContentControl startParent = selection.ParentContentControl;
            selection.SetRange(end, end);
            Word.ContentControl endParent = selection.ParentContentControl;
            selection.SetRange(start, end);

            if (startParent != null && endParent == null || startParent == null && endParent != null) return;
            if (startParent != null && endParent != null)
            {
                if (startParent.Range.Start != endParent.Range.Start || startParent.Range.End != endParent.Range.End) return;
            }

            //do not allow the same range is wrapped more than once
            Word.ContentControl parent = selection.ParentContentControl;
            if(parent != null)
            { 
                Word.Range range = parent.Range;
                if (selection.Range.Start == range.Start && selection.Range.End == range.End) return;
            }

            //wrap the content range 
            var activeDocument = Application.ActiveDocument;
            var extendedDocument = Globals.Factory.GetVstoObject(activeDocument);
            var next = DateTime.Now.Ticks.ToString();
            var control = extendedDocument.Controls.AddRichTextContentControl(string.Format("richText{0}", next));
            control.PlaceholderText = CurrentName;
            control.Tag = CurrentTag;
            control.Title = CurrentTag;
            control.Range.Font.Color = CurrentForeColor(); 
            control.Range.Font.Shading.ForegroundPatternColor = CurrentBackColor(); 
            control.Range.Font.Shading.Texture = Word.WdTextureIndex.wdTextureSolid;
            control.Range.HighlightColorIndex = Word.WdColorIndex.wdNoHighlight;
        }

        public void UnwrapContent()
        {
            var selection = Application.Selection;
            Word.ContentControl parent = selection.ParentContentControl;
            if (parent != null)
            {
                parent.Range.Font.Shading.BackgroundPatternColorIndex = Word.WdColorIndex.wdNoHighlight;
                parent.Delete(false);
            }
            if (selection.ContentControls.Count > 0)
            {
                foreach (Word.ContentControl control in selection.ContentControls)
                {
                    control.Range.Font.Shading.BackgroundPatternColorIndex = Word.WdColorIndex.wdNoHighlight;
                    control.Delete(false);
                }
            }
        }
    }
}
