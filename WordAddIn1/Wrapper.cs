using System;
using Newtonsoft.Json;
using System.Collections.Generic;
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

        public class Intent
        {
            public string name { get; set; }
            public double confidence { get; set; }
        }

        public class SingleEnt
        {
            public int start { get; set; }
            public int end { get; set; }
            public string value { get; set; }
            public string entity { get; set; }
            public float confidence { get; set; }
            public string extractor { get; set; }
        }

        public class IntentRanking
        {
            public string name { get; set; }
            public double confidence { get; set; }
        }

        public class SentenceObject
        {
            public Intent intent { get; set; }
            public List<SingleEnt> entities { get; set; }
            public List<IntentRanking> intent_ranking { get; set; }
            public string text { get; set; }
        }

        public void WrapContentFromJSON(string JSONresult)
        {
            List<SentenceObject> SentObjList = JsonConvert.DeserializeObject<List<SentenceObject>>(JSONresult);
            SentObjList.Reverse();
            if (string.IsNullOrEmpty(JSONresult)) return;

            var activeDocument = Application.ActiveDocument;
            var extendedDocument = Globals.Factory.GetVstoObject(activeDocument);

            Application.UndoRecord.StartCustomRecord($"Tag Selection ({CurrentTag})");

            int sentEnd = 0;
            foreach (Range sentence in Application.ActiveDocument.Sentences)
            {
                sentEnd += sentence.Text.Length;
            }
            sentEnd -= 1;

            foreach (SentenceObject sent in SentObjList)
            {
                int sentStart = sentEnd - sent.text.Length;
                Range intRange = Application.ActiveDocument.Range(sentStart, sentEnd);

                string intTag = sent.intent.name;
                WrapItem(extendedDocument, intTag, intRange);
                
                List<SingleEnt> EntList = sent.entities;
                EntList.Reverse();
                foreach (SingleEnt ent in EntList)
                {
                    int entStart = sentStart + ent.start + 1;
                    int entEnd = sentStart + ent.end + 1;
                    Range entRange = Application.ActiveDocument.Range(entStart, entEnd);

                    string entTag = ent.entity;
                    WrapItem(extendedDocument, entTag, entRange);
                }
                sentEnd = sentStart - 1;
            }
            Application.UndoRecord.EndCustomRecord();
        }

        public void WrapItem(Microsoft.Office.Tools.Word.Document extendedDocument, string tag, Range range)
        {
            try
            {
                var next = DateTime.Now.Ticks.ToString();
                var control = extendedDocument.Controls.AddRichTextContentControl(range, string.Format("richText{0}", next));
                control.PlaceholderText = "...";
                control.Tag = tag;
                control.Title = tag;
                HighlightControlHierarchy(control.Range);
            }
            catch (Exception ex)
            {
                Utilities.Notification(ex.Message);
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
                control.Range.Select();
                Application.Selection.ClearFormatting();
                originalRange.Select();

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
