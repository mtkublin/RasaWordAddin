using System;
using System.Windows.Forms;
using XL.Office.Helpers;
using Word = Microsoft.Office.Interop.Word;
using Microsoft.Office.Interop.Word;

namespace WordAddIn1
{
    public partial class ThisAddIn
    {
        private void HighlightControlHierarchy(Word.Range range)
        {
            var parent = range.ParentContentControl;
            if (parent != null)
            {
                HighlightContentControl(parent.Tag, parent.Range);
            }
            foreach (Word.ContentControl child in range.ContentControls)
            {
                HighlightContentControl(child.Tag, child.Range);
            }
        }

        private void HighlightContentControl(string tag, Word.Range range)
        {
            //do not wrap if tag is empty or null
            if (string.IsNullOrEmpty(tag)) return;

            Utilities.Notification(tag);
            try
            {
                range.Font.Color = TagForeColor(tag);
                range.Font.Shading.ForegroundPatternColor = TagBackColor(tag);
                range.Font.Shading.Texture = Word.WdTextureIndex.wdTextureSolid;
                range.HighlightColorIndex = Word.WdColorIndex.wdNoHighlight;
            }
            catch (Exception ex)
            {
                Utilities.Notification(ex.ToString());
            }
        }

        public void WrapContent()
        {
            //do not wrap if current tag is empty or null
            if (string.IsNullOrEmpty(CurrentTag)) return;

            var selection = Application.Selection;
            //do not wrap if range is collapsed
            if (selection.Start == selection.End) return;

            //TODO identify where new content control cannot be created
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

            var activeDocument = Application.ActiveDocument;
            var extendedDocument = Globals.Factory.GetVstoObject(activeDocument);
            var next = DateTime.Now.Ticks.ToString();

            try
            {
                Application.UndoRecord.StartCustomRecord($"Tag Selection ({CurrentTag})");
                Word.ContentControl parent = selection.ParentContentControl;
                //change tag, if entire content control selected
                if (parent != null && selection.Range.Start == parent.Range.Start && selection.Range.End == parent.Range.End)
                {
                    parent.Tag = CurrentTag;
                    parent.Title = CurrentName;
                    HighlightControlHierarchy(parent.Range);
                }
                //wrap the content range 
                else
                { 
                    var control = extendedDocument.Controls.AddRichTextContentControl(string.Format("richText{0}", next));
                    control.PlaceholderText = "...";
                    control.Tag = CurrentTag;
                    control.Title = CurrentName;
                    HighlightControlHierarchy(control.Range);
                }
            }
            catch (Exception ex)
            {
                Utilities.Notification(ex.Message);
            }
            finally
            {
                Application.UndoRecord.EndCustomRecord();
            }
        }

        public void UnwrapContent()
        {
            var selection = Application.Selection;
            Word.Range originalRange = selection.Range;
            Word.ContentControl control = selection.ParentContentControl;
            Word.ContentControl parent = control.ParentContentControl;
            var remainingControls = control.Range.ContentControls;

            Application.UndoRecord.StartCustomRecord("Remove Tag");
            if (control != null)
            {
                //clear content control formatting
                //control.Range.Select();
                //Application.Selection.ClearFormatting();
                //originalRange.Select();

                //remove content control
                control.Delete(false);
            }
            if (parent != null)
            {
                HighlightControlHierarchy(parent.Range);
            }
            if (remainingControls != null && remainingControls.Count > 0)
            {
                foreach (Word.ContentControl survivor in remainingControls)
                {
                    if(survivor != null)
                    {
                        HighlightControlHierarchy(survivor.Range);
                    }
                }
            }
            Application.UndoRecord.EndCustomRecord();

            //UNDONE remove all content controls in entire selection
            //if (selection.ContentControls.Count > 0)
            //{
            //    foreach (Word.ContentControl child in selection.ContentControls)
            //    {                    
            //        var range = child.Range;
            //        range.Font.Color = Word.WdColor.wdColorBlack;
            //        range.Font.Shading.ForegroundPatternColor = Word.WdColor.wdColorWhite;
            //        range.Font.Shading.Texture = Word.WdTextureIndex.wdTextureNone;
            //        range.HighlightColorIndex = Word.WdColorIndex.wdNoHighlight;
            //        child.Delete(false);
            //    }
            //}
        }
    }
}
