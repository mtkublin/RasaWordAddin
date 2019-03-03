using Microsoft.Office.Interop.Word;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace WordAddIn1
{
    public partial class ThisAddIn
    {
        Microsoft.Office.Tools.Word.Bookmark currentBookmark;

        public void ThisDocument_SelectionChange(object sender, Microsoft.Office.Tools.Word.SelectionEventArgs e, Microsoft.Office.Tools.Word.Bookmark bookmark, Microsoft.Office.Tools.Word.Document extendedDocument)
        {
            if (currentBookmark != null)
            {
                List<int> RangeStartsList = new List<int>();
                List<int> RangeEndsList = new List<int>();

                string bookmarkLevelIndicator = bookmark.Name.Substring(bookmark.Name.Length - 1);

                int selectionStart = this.Application.Selection.Start;
                RangeStartsList.Add(selectionStart);
                int selectionEnd = this.Application.Selection.End;
                RangeEndsList.Add(selectionEnd);

                int bookmarkStart = bookmark.Range.Start;
                RangeStartsList.Add(bookmarkStart);
                int bookmarkEnd = bookmark.Range.End;
                RangeEndsList.Add(bookmarkEnd);

                int bookmarkLen = bookmarkEnd - bookmarkStart;

                bool isStartActive = Application.Selection.StartIsActive;

                if ((selectionEnd != selectionStart) & ((selectionStart >= bookmarkStart & selectionStart < bookmarkEnd & selectionEnd > bookmarkEnd) || (selectionStart < bookmarkStart & selectionEnd > bookmarkStart & selectionEnd <= bookmarkEnd) || (selectionStart < bookmarkStart & selectionEnd > bookmarkEnd)))
                {
                    Range NewBookmarkRange = this.Application.ActiveDocument.Range(bookmarkStart, bookmarkEnd);
                    bool isNewBokmarkBigger = false;
                    bool continueToReplace = true;

                    if (isStartActive)
                    {
                        if (selectionStart > bookmarkStart & selectionStart < bookmarkEnd & selectionEnd > bookmarkEnd)
                        {
                            NewBookmarkRange = this.Application.ActiveDocument.Range(bookmarkStart, selectionStart);
                        }
                        else if (selectionStart < bookmarkStart & selectionEnd > bookmarkStart & selectionEnd <= bookmarkEnd)
                        {
                            NewBookmarkRange = this.Application.ActiveDocument.Range(selectionStart, bookmarkEnd);
                            isNewBokmarkBigger = true;
                        }
                        else
                        {
                            continueToReplace = false;
                        }
                    }
                    else
                    {
                        if (selectionStart >= bookmarkStart & selectionStart < bookmarkEnd & selectionEnd > bookmarkEnd)
                        {
                            NewBookmarkRange = this.Application.ActiveDocument.Range(bookmarkStart, selectionEnd);
                            isNewBokmarkBigger = true;
                        }
                        else if (selectionStart < bookmarkStart & selectionEnd > bookmarkStart & selectionEnd < bookmarkEnd)
                        {
                            NewBookmarkRange = this.Application.ActiveDocument.Range(selectionEnd, bookmarkEnd);
                        }
                        else
                        {
                            continueToReplace = false;
                        }
                    }

                    if (continueToReplace & NewBookmarkRange != this.Application.ActiveDocument.Range(bookmarkStart, bookmarkEnd) & NewBookmarkRange != null)
                    {
                        string mainBookmarkName = bookmark.Name.ToString();
                        string mainBookmarkTag = Regex.Replace(mainBookmarkName, "_[0-9]+_entity_", "");
                        mainBookmarkTag = Regex.Replace(mainBookmarkTag, "_[0-9]+_intent_", "");
                        mainBookmarkTag = Regex.Replace(mainBookmarkTag, "_[0-9]+_notspecified_", "");
                        mainBookmarkTag = Regex.Replace(mainBookmarkTag, "_", "-");

                        replaceBookmarkWithNewRange(NewBookmarkRange, bookmark, extendedDocument, true);

                        List<Microsoft.Office.Tools.Word.Bookmark> bookmarksToHighlightList = new List<Microsoft.Office.Tools.Word.Bookmark>();

                        if (isNewBokmarkBigger)
                        {
                            List<Microsoft.Office.Tools.Word.Bookmark> bookmarksToReplaceList = new List<Microsoft.Office.Tools.Word.Bookmark>();

                            foreach (Bookmark bm in NewBookmarkRange.Bookmarks)
                            {
                                string bookmarkName = bm.Name;
                                if (bookmarkName.EndsWith(bookmarkLevelIndicator) & bm.Range != NewBookmarkRange)
                                {
                                    Microsoft.Office.Tools.Word.Bookmark VSTObookmark = extendedDocument.Controls[bookmarkName] as Microsoft.Office.Tools.Word.Bookmark;
                                    bookmarksToReplaceList.Add(VSTObookmark);
                                }
                            }

                            foreach (Microsoft.Office.Tools.Word.Bookmark VSTObookmark in bookmarksToReplaceList)
                            {
                                int VSTObookmarkStart = VSTObookmark.Range.Start;
                                RangeStartsList.Add(VSTObookmarkStart);
                                int VSTObookmarkEnd = VSTObookmark.Range.End;
                                RangeEndsList.Add(VSTObookmarkEnd);
                                Range VSTObookmarkRange = this.Application.ActiveDocument.Range(VSTObookmarkStart, VSTObookmarkEnd);

                                bool VSTObookmarkDelete = false;

                                if (selectionStart > VSTObookmarkStart & selectionEnd > VSTObookmarkEnd)
                                {
                                    VSTObookmarkRange = this.Application.ActiveDocument.Range(VSTObookmarkStart, selectionStart);
                                }
                                else if (selectionStart < VSTObookmarkStart & selectionEnd < VSTObookmarkEnd)
                                {
                                    VSTObookmarkRange = this.Application.ActiveDocument.Range(selectionEnd, VSTObookmarkEnd);
                                }
                                else if (selectionStart < VSTObookmarkStart & selectionEnd > VSTObookmarkEnd)
                                {
                                    VSTObookmarkDelete = true;
                                }

                                if (VSTObookmarkDelete)
                                {
                                    VSTObookmark.Delete();
                                }
                                else
                                {
                                    replaceBookmarkWithNewRange(VSTObookmarkRange, VSTObookmark, extendedDocument, false);
                                    bookmarksToHighlightList.Add(VSTObookmark);
                                }
                            }
                        }

                        RangeStartsList.Sort();
                        RangeEndsList.Sort();

                        int HLrangeStart = RangeStartsList[0];
                        int HLrangeEnd = RangeEndsList[RangeEndsList.Count - 1];
                        if (HLrangeStart < HLrangeEnd)
                        {
                            Range HLrange = this.Application.ActiveDocument.Range(HLrangeStart, HLrangeEnd);
                            UnhighlightControl(HLrange);

                            if (bookmark.Name.EndsWith("1"))
                            {
                                foreach (Bookmark BMtoHL in HLrange.Bookmarks) if (BMtoHL.Name.EndsWith("1"))
                                {
                                    string BMtoHLName = BMtoHL.Name.ToString();
                                    string BMtoHLTag = Regex.Replace(BMtoHLName, "_[0-9]+_entity_", "");
                                    BMtoHLTag = Regex.Replace(BMtoHLTag, "_[0-9]+_intent_", "");
                                    BMtoHLTag = Regex.Replace(BMtoHLTag, "_[0-9]+_notspecified_", "");
                                    BMtoHLTag = Regex.Replace(BMtoHLTag, "_", "-");

                                    HighlightContentControl(BMtoHLTag, BMtoHL.Range);
                                }
                                foreach (Bookmark BMtoHL in HLrange.Bookmarks) if (BMtoHL.Name.EndsWith("2"))
                                {
                                    string BMtoHLName = BMtoHL.Name.ToString();
                                    string BMtoHLTag = Regex.Replace(BMtoHLName, "_[0-9]+_entity_", "");
                                    BMtoHLTag = Regex.Replace(BMtoHLTag, "_[0-9]+_intent_", "");
                                    BMtoHLTag = Regex.Replace(BMtoHLTag, "_[0-9]+_notspecified_", "");
                                    BMtoHLTag = Regex.Replace(BMtoHLTag, "_", "-");

                                    HighlightContentControl(BMtoHLTag, BMtoHL.Range);
                                }
                            }
                            else
                            {
                                foreach (Bookmark BMtoHL in HLrange.Bookmarks) if (BMtoHL.Name.EndsWith("2"))
                                {
                                    string BMtoHLName = BMtoHL.Name.ToString();
                                    string BMtoHLTag = Regex.Replace(BMtoHLName, "_[0-9]+_entity_", "");
                                    BMtoHLTag = Regex.Replace(BMtoHLTag, "_[0-9]+_intent_", "");
                                    BMtoHLTag = Regex.Replace(BMtoHLTag, "_[0-9]+_notspecified_", "");
                                    BMtoHLTag = Regex.Replace(BMtoHLTag, "_", "-");

                                    HighlightContentControl(BMtoHLTag, BMtoHL.Range);
                                }
                            }
                        }
                    }
                }
            }
        }

        private void replaceBookmarkWithNewRange(Range NewBookmarkRange, Microsoft.Office.Tools.Word.Bookmark bookmark, Microsoft.Office.Tools.Word.Document extendedDocument, bool isMainEditedBM)
        {
            string bookmarkName = bookmark.Name.ToString();
            string tag = Regex.Replace(bookmarkName, "_[0-9]+_entity_", "");
            tag = Regex.Replace(tag, "_[0-9]+_intent_", "");
            tag = Regex.Replace(tag, "_[0-9]+_notspecified_", "");
            tag = Regex.Replace(tag, "_", "-");

            bookmark.Delete();

            Microsoft.Office.Tools.Word.Bookmark newBookmark = extendedDocument.Controls.AddBookmark(NewBookmarkRange, bookmarkName);
            newBookmark.Tag = tag;
            newBookmark.Selected += new Microsoft.Office.Tools.Word.SelectionEventHandler((sender2, e2) => bookmark_Selected(sender2, e2, extendedDocument, newBookmark));
            if (newBookmark.Name.EndsWith("1"))
            {
                newBookmark.SelectionChange += new Microsoft.Office.Tools.Word.SelectionEventHandler((sender2, e2) => bookmark_SelectionChange(sender2, e2, extendedDocument, newBookmark));
            }

            if (isMainEditedBM)
            {
                currentBookmark = newBookmark;
                if (newBookmark.Text.Length <= 1024)
                {
                    Globals.Ribbons.Ribbon1.CurBMtextLabel.Label = newBookmark.Text;
                }
                else
                {
                    Globals.Ribbons.Ribbon1.CurBMtextLabel.Label = "..." + newBookmark.Text.Substring(newBookmark.Text.Length - 1020);
                }

                if (tag.Length <= 1024)
                {
                    Globals.Ribbons.Ribbon1.CurBMentLabel.Label = tag;
                }
                else
                {
                    Globals.Ribbons.Ribbon1.CurBMentLabel.Label = "..." + tag.Substring(newBookmark.Text.Length - 1020);
                }

                if (bookmarkName.EndsWith("1"))
                {
                    Globals.Ribbons.Ribbon1.IntOrEntLabel.Label = "Intent";
                }
                else if (bookmarkName.EndsWith("2"))
                {
                    Globals.Ribbons.Ribbon1.IntOrEntLabel.Label = "Entity";
                }
            }
        }

        public void bookmark_Selected(object sender, Microsoft.Office.Tools.Word.SelectionEventArgs e, Microsoft.Office.Tools.Word.Document extendedDocument, Microsoft.Office.Tools.Word.Bookmark bookmark)
        {
            string name = bookmark.Name.ToString();
            string NewTag = Regex.Replace(name, "_[0-9]+_entity_", "");
            NewTag = Regex.Replace(NewTag, "_[0-9]+_intent_", "");
            NewTag = Regex.Replace(NewTag, "_[0-9]+_notspecified_", "");
            NewTag = Regex.Replace(NewTag, "_", "-");

            currentBookmark = bookmark;

            if (currentBookmark.Text.Length <= 35)
            {
                Globals.Ribbons.Ribbon1.CurBMtextLabel.Label = currentBookmark.Text;
            }
            else
            {
                Globals.Ribbons.Ribbon1.CurBMtextLabel.Label = currentBookmark.Text.Substring(0, 14) + "... ..." + currentBookmark.Text.Substring(currentBookmark.Text.Length - 14);
            }

            if (NewTag.Length <= 35)
            {
                Globals.Ribbons.Ribbon1.CurBMentLabel.Label = NewTag;
            }
            else
            {
                Globals.Ribbons.Ribbon1.CurBMentLabel.Label = NewTag.Substring(0, 14) + "... ..." + NewTag.Substring(currentBookmark.Text.Length - 14);
            }

            if (NewTag.EndsWith("1"))
            {
                Globals.Ribbons.Ribbon1.IntOrEntLabel.Label = "Intent";
            }
            else if (NewTag.EndsWith("2"))
            {
                Globals.Ribbons.Ribbon1.IntOrEntLabel.Label = "Entity";
            }
        }

        public void bookmark_SelectionChange(object sender, Microsoft.Office.Tools.Word.SelectionEventArgs e, Microsoft.Office.Tools.Word.Document extendedDocument, Microsoft.Office.Tools.Word.Bookmark bookmark)
        {
            if (bookmark.Name.EndsWith("1"))
            {
                Microsoft.Office.Tools.Word.Bookmark VSTObookmark = extendedDocument.Controls[bookmark.Name] as Microsoft.Office.Tools.Word.Bookmark;

                string name = bookmark.Name.ToString();
                string NewTag = Regex.Replace(name, "_[0-9]+_entity_", "");
                NewTag = Regex.Replace(NewTag, "_[0-9]+_intent_", "");
                NewTag = Regex.Replace(NewTag, "_[0-9]+_notspecified_", "");
                NewTag = Regex.Replace(NewTag, "_", "-");

                currentBookmark = VSTObookmark;

                if (currentBookmark.Text.Length <= 35)
                {
                    Globals.Ribbons.Ribbon1.CurBMtextLabel.Label = currentBookmark.Text;
                }
                else
                {
                    Globals.Ribbons.Ribbon1.CurBMtextLabel.Label = currentBookmark.Text.Substring(0, 14) + "... ..." + currentBookmark.Text.Substring(currentBookmark.Text.Length - 14);
                }

                if (NewTag.Length <= 35)
                {
                    Globals.Ribbons.Ribbon1.CurBMentLabel.Label = NewTag;
                }
                else
                {
                    Globals.Ribbons.Ribbon1.CurBMentLabel.Label = NewTag.Substring(0, 14) + "... ..." + NewTag.Substring(currentBookmark.Text.Length - 14);
                }

                Globals.Ribbons.Ribbon1.IntOrEntLabel.Label = "Intent";
            }
        }

        Range FirstSelectedRange;
        Range LastSelectedRange;
        bool IsItFirstBM = true;

        public void HighlightBookmarksInVisibleRange()
        {
            Range WholeDocRange = this.Application.ActiveDocument.Content;
            UnhighlightControl(WholeDocRange);

            System.Windows.Rect rect = System.Windows.Automation.AutomationElement.FocusedElement.Current.BoundingRectangle;

            foreach (Range r in this.Application.ActiveDocument.StoryRanges)
            {
                int left = 0, top = 0, width = 0, height = 0;
                try
                {
                    try
                    {
                        this.Application.ActiveWindow.GetPoint(out left, out top, out width, out height, r);
                    }
                    catch
                    {
                        left = (int)rect.Left;
                        top = (int)rect.Top;
                        width = (int)rect.Width;
                        height = (int)rect.Height;
                    }
                    System.Windows.Rect newRect = new System.Windows.Rect(left, top, width, height);
                    System.Windows.Rect inter;
                    if ((inter = System.Windows.Rect.Intersect(rect, newRect)) != System.Windows.Rect.Empty)
                    {
                        Range r1 = this.Application.ActiveWindow.RangeFromPoint((int)inter.Left, (int)inter.Top);
                        Range r2 = this.Application.ActiveWindow.RangeFromPoint((int)inter.Right, (int)inter.Bottom);
                        r.SetRange(r1.Start, r2.Start);

                        foreach (Bookmark bookmark in r.Bookmarks) if (bookmark.Name.EndsWith("1"))
                        {
                            string name = bookmark.Name.ToString();
                            string NewTag = Regex.Replace(name, "_[0-9]+_entity_", "");
                            NewTag = Regex.Replace(NewTag, "_[0-9]+_intent_", "");
                            NewTag = Regex.Replace(NewTag, "_[0-9]+_notspecified_", "");
                            NewTag = Regex.Replace(NewTag, "_", "-");

                            HighlightContentControl(NewTag, bookmark.Range);
                            LastSelectedRange = bookmark.Range;
                            if (IsItFirstBM)
                            {
                                FirstSelectedRange = bookmark.Range;
                                IsItFirstBM = false;
                            }
                        }
                        foreach (Bookmark bookmark in r.Bookmarks) if (bookmark.Name.EndsWith("2"))
                        {
                            string name = bookmark.Name.ToString();
                            string NewTag = Regex.Replace(name, "_[0-9]+_entity_", "");
                            NewTag = Regex.Replace(NewTag, "_[0-9]+_intent_", "");
                            NewTag = Regex.Replace(NewTag, "_[0-9]+_notspecified_", "");
                            NewTag = Regex.Replace(NewTag, "_", "-");

                            HighlightContentControl(NewTag, bookmark.Range);
                            LastSelectedRange = bookmark.Range;
                            if (IsItFirstBM)
                            {
                                FirstSelectedRange = bookmark.Range;
                                IsItFirstBM = false;
                            }
                        }
                    }
                }
                catch { }
                IsItFirstBM = true;
            }
        }

        public void HighlightBookmarksInNextRange()
        {
            Range WholeDocRange = this.Application.ActiveDocument.Content;
            UnhighlightControl(WholeDocRange);

            if (LastSelectedRange != null)
            {
                int HLstart = LastSelectedRange.End;
                if (HLstart < this.Application.ActiveDocument.Content.End)
                {
                    int HLlength = LastSelectedRange.End - FirstSelectedRange.Start;
                    int HLend = HLstart + HLlength;
                    if (HLend > this.Application.ActiveDocument.Content.End) HLend = this.Application.ActiveDocument.Content.End;
                    Range HLrange = this.Application.ActiveDocument.Range(HLstart, HLend);

                    foreach (Bookmark bookmark in HLrange.Bookmarks) if (bookmark.Name.EndsWith("1"))
                    {
                        string name = bookmark.Name.ToString();
                        string NewTag = Regex.Replace(name, "_[0-9]+_entity_", "");
                        NewTag = Regex.Replace(NewTag, "_[0-9]+_intent_", "");
                        NewTag = Regex.Replace(NewTag, "_[0-9]+_notspecified_", "");
                        NewTag = Regex.Replace(NewTag, "_", "-");

                        HighlightContentControl(NewTag, bookmark.Range);
                        LastSelectedRange = bookmark.Range;
                        if (IsItFirstBM)
                        {
                            FirstSelectedRange = bookmark.Range;
                            IsItFirstBM = false;
                        }
                    }
                    foreach (Bookmark bookmark in HLrange.Bookmarks) if (bookmark.Name.EndsWith("2"))
                        {
                        string name = bookmark.Name.ToString();
                        string NewTag = Regex.Replace(name, "_[0-9]+_entity_", "");
                        NewTag = Regex.Replace(NewTag, "_[0-9]+_intent_", "");
                        NewTag = Regex.Replace(NewTag, "_[0-9]+_notspecified_", "");
                        NewTag = Regex.Replace(NewTag, "_", "-");

                        HighlightContentControl(NewTag, bookmark.Range);
                        LastSelectedRange = bookmark.Range;
                        if (IsItFirstBM)
                        {
                            FirstSelectedRange = bookmark.Range;
                            IsItFirstBM = false;
                        }
                    }

                    IsItFirstBM = true;
                }
            }
        }
    }
}
