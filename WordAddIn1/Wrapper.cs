using System;
using System.Windows.Forms;
using XL.Office.Helpers;
using Word = Microsoft.Office.Interop.Word;
using Microsoft.Office.Interop.Word;
using System.Text.RegularExpressions;

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

        public void UnhighlightControl(Range range)
        {
            try
            {
                range.Font.Color = Utilities.RGBwdColor(System.Drawing.Color.Black);
                range.Font.Shading.ForegroundPatternColor = Utilities.RGBwdColor(System.Drawing.Color.White);
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
            Range range = this.Application.Selection.Range;
            Document activeDocument = Application.ActiveDocument;
            var extendedDocument = Globals.Factory.GetVstoObject(activeDocument);
            Application.UndoRecord.StartCustomRecord($"Tag Selection ({CurrentTag})");

            string currentTagLevelIndicator = CurrentTag.Substring(CurrentTag.Length - 1);
            bool doesRangeOverlapExistingEnt = false;

            foreach (Bookmark existingBM in range.Bookmarks)
            {
                if (existingBM.Name.EndsWith(currentTagLevelIndicator))
                {
                    doesRangeOverlapExistingEnt = true;
                }
            }

            if (range.Start != range.End & doesRangeOverlapExistingEnt == false)
            {
                try
                {
                    int BookmarkNumber = this.Application.ActiveDocument.Bookmarks.Count;

                    string bookmarkName = "_" + BookmarkNumber.ToString();
                    if (CurrentTag.EndsWith("1"))
                    {
                        bookmarkName += "_intent_";
                    }
                    else if (CurrentTag.EndsWith("2"))
                    {
                        bookmarkName += "_entity_";
                    }
                    else
                    {
                        bookmarkName += "_notspecified_";
                    }
                    string NewTag = Regex.Replace(CurrentTag, "-", "_");
                    bookmarkName += NewTag;

                    Microsoft.Office.Tools.Word.Bookmark bookmark = extendedDocument.Controls.AddBookmark(range, bookmarkName);
                    bookmark.Tag = CurrentTag;
                    HighlightContentControl(CurrentTag, bookmark.Range);

                    foreach (Bookmark bm in bookmark.Range.Bookmarks) if (bm != bookmark)
                    {
                        string bmName = bm.Name.ToString();
                        string bmTag = Regex.Replace(bmName, "_[0-9]+_entity_", "");
                        bmTag = Regex.Replace(bmTag, "_[0-9]+_intent_", "");
                        bmTag = Regex.Replace(bmTag, "_[0-9]+_notspecified_", "");
                        bmTag = Regex.Replace(bmTag, "_", "-");
                        HighlightContentControl(bmTag, bm.Range);
                    }

                    bookmark.Selected += new Microsoft.Office.Tools.Word.SelectionEventHandler((sender, e) => bookmark_Selected(sender, e, extendedDocument, bookmark));
                    if (bookmark.Name.EndsWith("1"))
                    {
                        bookmark.SelectionChange += new Microsoft.Office.Tools.Word.SelectionEventHandler((sender2, e2) => bookmark_SelectionChange(sender2, e2, extendedDocument, bookmark));
                    }
                }
                catch (Exception ex)
                {
                    Utilities.Notification(ex.Message);
                }
            }
            Application.UndoRecord.EndCustomRecord();
        }

        public void UnwrapContent()
        {
            foreach (Bookmark bm in Application.Selection.Range.Bookmarks)
            {
                string bmName = bm.Name;
                Range bmRange = bm.Range;
                Microsoft.Office.Tools.Word.Document vstoDoc = Globals.Factory.GetVstoObject(this.Application.ActiveDocument);
                Microsoft.Office.Tools.Word.Bookmark VSTObookmark = vstoDoc.Controls[bmName] as Microsoft.Office.Tools.Word.Bookmark;

                if (currentBookmark == VSTObookmark)
                {
                    currentBookmark = null;
                    Globals.Ribbons.Ribbon1.CurBMtextLabel.Label = "";
                    Globals.Ribbons.Ribbon1.CurBMentLabel.Label = "";
                    Globals.Ribbons.Ribbon1.IntOrEntLabel.Label = "";
                }

                UnhighlightControl(VSTObookmark.Range);
                VSTObookmark.Delete();

                if (bmName.EndsWith("1"))
                {
                    foreach (Bookmark entInInt in bmRange.Bookmarks) if (entInInt.Name.EndsWith("2"))
                    {
                        string entInIntName = entInInt.Name.ToString();
                        string entInIntTag = Regex.Replace(entInIntName, "_[0-9]+_entity_", "");
                        entInIntTag = Regex.Replace(entInIntTag, "_[0-9]+_intent_", "");
                        entInIntTag = Regex.Replace(entInIntTag, "_[0-9]+_notspecified_", "");
                        entInIntTag = Regex.Replace(entInIntTag, "_", "-");
                        HighlightContentControl(entInIntTag, entInInt.Range);
                    }
                }
            }
        }
    }
}
