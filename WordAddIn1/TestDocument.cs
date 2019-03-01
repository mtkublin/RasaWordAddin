using System;
using System.Collections.Generic;
using System.Windows.Forms;
using System.Diagnostics;
using System.Text.RegularExpressions;
using System.IO;
using XL.Office.Helpers;
using Microsoft.Office.Interop.Word;
using Newtonsoft.Json;
using RestSharp;


namespace WordAddIn1
{
    public partial class ThisAddIn
    {
        public void TestDoc()
        {
            List<String> SentsToExport = new List<String>();
            foreach (Range sent in Globals.ThisAddIn.Application.ActiveDocument.Sentences)
            {
                SentsToExport.Add(sent.Text);
            }
            TextToExportObject textToExportObject = new TextToExportObject(SentsToExport);
            FinalTestDataExportObject finalExportData = new FinalTestDataExportObject(textToExportObject);
            var jsonTestObject = JsonConvert.SerializeObject(finalExportData);

            var client = new RestClient("http://127.0.0.1:6000");
            var request = new RestRequest("api/testdata", Method.POST);
            request.AddParameter("application/json; charset=utf-8", jsonTestObject, ParameterType.RequestBody);
            request.RequestFormat = DataFormat.Json;
            IRestResponse response = client.Execute(request);
            string JSONresultDoc = response.Content.ToString();

            WrapFromJSON(JSONresultDoc);

            Microsoft.Office.Tools.Word.Document vstoDoc = Globals.Factory.GetVstoObject(this.Application.ActiveDocument);
            var handler = new Microsoft.Office.Tools.Word.SelectionEventHandler((sender, e) => ThisDocument_SelectionChange(sender, e, currentBookmark, vstoDoc));
            vstoDoc.SelectionChange -= handler;
            vstoDoc.SelectionChange += handler;
        }

        Microsoft.Office.Tools.Word.Bookmark currentBookmark;

        void ThisDocument_SelectionChange(object sender, Microsoft.Office.Tools.Word.SelectionEventArgs e, Microsoft.Office.Tools.Word.Bookmark bookmark, Microsoft.Office.Tools.Word.Document extendedDocument)
        {
            string bookmarkLevelIndicator = bookmark.Name.Substring(bookmark.Name.Length);

            if (currentBookmark != null)
            {
                int selectionStart = this.Application.Selection.Start;
                int selectionEnd = this.Application.Selection.End;
                int bookmarkStart = bookmark.Range.Start;
                int bookmarkEnd = bookmark.Range.End;
                int bookmarkLen = bookmarkEnd - bookmarkStart;

                bool isStartActive = Application.Selection.StartIsActive;

                if ((selectionEnd != selectionStart) & ((selectionStart >= bookmarkStart & selectionStart < bookmarkEnd & selectionEnd > bookmarkEnd) || (selectionStart < bookmarkStart & selectionEnd > bookmarkStart & selectionEnd <= bookmarkEnd) || (selectionStart < bookmarkStart & selectionEnd > bookmarkEnd)))
                {
                    Range NewBookmarkRange = this.Application.ActiveDocument.Range(bookmarkStart, bookmarkEnd);
                    bool isNewBokmarkBigger = false;
                    bool deleteBookmark = false;
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
                        else if (selectionStart <= bookmarkStart & selectionEnd > bookmarkEnd)
                        {
                            NewBookmarkRange = null;
                            deleteBookmark = true;
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
                        else if (selectionStart < bookmarkStart & selectionEnd > bookmarkEnd)
                        {
                            NewBookmarkRange = this.Application.ActiveDocument.Range(selectionStart, selectionEnd);
                            isNewBokmarkBigger = true;
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

                        UnhighlightControl(bookmark.Range);
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
                                    UnhighlightControl(VSTObookmark.Range);
                                }
                            }

                            foreach (Microsoft.Office.Tools.Word.Bookmark VSTObookmark in bookmarksToReplaceList)
                            {
                                int VSTObookmarkStart = VSTObookmark.Range.Start;
                                int VSTObookmarkEnd = VSTObookmark.Range.End;
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

                        HighlightContentControl(mainBookmarkTag, NewBookmarkRange);
                        foreach (Bookmark bm in NewBookmarkRange.Bookmarks)
                        {
                            if (bm.Range != bookmark.Range & bm.Range != NewBookmarkRange)
                            {
                                string name = bm.Name.ToString();
                                string NewTag = Regex.Replace(name, "_[0-9]+_entity_", "");
                                NewTag = Regex.Replace(NewTag, "_[0-9]+_intent_", "");
                                NewTag = Regex.Replace(NewTag, "_[0-9]+_notspecified_", "");
                                NewTag = Regex.Replace(NewTag, "_", "-");

                                HighlightContentControl(NewTag, bm.Range);
                            }
                        }
                        foreach (Microsoft.Office.Tools.Word.Bookmark bm in bookmarksToHighlightList)
                        {
                            string name = bm.Name.ToString();
                            string NewTag = Regex.Replace(name, "_[0-9]+_entity_", "");
                            NewTag = Regex.Replace(NewTag, "_[0-9]+_intent_", "");
                            NewTag = Regex.Replace(NewTag, "_[0-9]+_notspecified_", "");
                            NewTag = Regex.Replace(NewTag, "_", "-");

                            HighlightContentControl(NewTag, bm.Range);

                            foreach (Bookmark insideBM in bm.Range.Bookmarks)
                            {
                                if (insideBM.Range != bm.Range)
                                {
                                    string insideName = insideBM.Name.ToString();
                                    string insideNewTag = Regex.Replace(insideName, "_[0-9]+_entity_", "");
                                    insideNewTag = Regex.Replace(insideNewTag, "_[0-9]+_intent_", "");
                                    insideNewTag = Regex.Replace(insideNewTag, "_[0-9]+_notspecified_", "");
                                    insideNewTag = Regex.Replace(insideNewTag, "_", "-");

                                    HighlightContentControl(insideNewTag, insideBM.Range);
                                }
                            }
                        }
                    }
                    else if (deleteBookmark)
                    {
                        bookmark.Delete();
                        currentBookmark = null;
                        Globals.Ribbons.Ribbon1.CurBMtextLabel.Label = "";
                        Globals.Ribbons.Ribbon1.CurBMentLabel.Label = "";
                        Globals.Ribbons.Ribbon1.IntOrEntLabel.Label = "";
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
            newBookmark.Deselected += new Microsoft.Office.Tools.Word.SelectionEventHandler((sender2, e2) => bookmark_Deselected(sender2, e2, newBookmark));
            //HighlightContentControl(tag, NewBookmarkRange);

            //foreach (Bookmark bm in NewBookmarkRange.Bookmarks)
            //{
            //    if (bm.Range != bookmark.Range & bm.Range != NewBookmarkRange)
            //    {
            //        string name = bm.Name.ToString();
            //        string NewTag = Regex.Replace(name, "_[0-9]+_entity_", "");
            //        NewTag = Regex.Replace(NewTag, "_[0-9]+_intent_", "");
            //        NewTag = Regex.Replace(NewTag, "_[0-9]+_notspecified_", "");
            //        NewTag = Regex.Replace(NewTag, "_", "-");

            //        HighlightContentControl(NewTag, bm.Range);
            //    }
            //}

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

        private void WrapFromJSON(string JSONresult)
        {
            List<SentenceObject> SentObjList = JsonConvert.DeserializeObject<List<SentenceObject>>(JSONresult);
            SentObjList.Reverse();
            if (string.IsNullOrEmpty(JSONresult)) return;

            bool AreThereEntities = false;
            foreach (SentenceObject sent in SentObjList)
            {
                if (sent.entities.Count != 0)
                {
                    AreThereEntities = true;
                }
                if (sent.intent == null)
                {
                    Globals.Ribbons.Ribbon1.TextMessageOkDialog("Null intents present, empty interpreter");
                    return;
                }
            }

            if (AreThereEntities == false)
            {
                Globals.Ribbons.Ribbon1.TextMessageOkDialog("Interpreter didn't find any entities.");
            }

            Stopwatch stopWatch = new Stopwatch();
            stopWatch.Start();

            Document activeDocument = Application.ActiveDocument;
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
                int sentStart = sentEnd - sent.text.Length + 1;

                Range intRange = Application.ActiveDocument.Range(sentStart, sentEnd);

                int ContentSubstraction = 0;
                string intTag = sent.intent.name;
                if (intTag != "empty-intent-1")
                {
                    WrapItem(extendedDocument, intTag, intRange);
                    ContentSubstraction = 0;
                }

                List<SingleEnt> EntList = sent.entities;
                EntList.Reverse();
                foreach (SingleEnt ent in EntList)
                {
                    int entStart = sentStart + ent.start + ContentSubstraction;
                    int entEnd = sentStart + ent.end + ContentSubstraction;

                    Range entRange = Application.ActiveDocument.Range(entStart, entEnd);

                    string entTag = ent.entity;
                    WrapItem(extendedDocument, entTag, entRange);
                }
                sentEnd = sentStart - 1;
            }
            Application.UndoRecord.EndCustomRecord();

            HighlightBookmarksInVisibleRange();

            stopWatch.Stop();
            TimeSpan ts = stopWatch.Elapsed;

            string elapsedTime = String.Format("{0:00}:{1:00}:{2:00}.{3:00}",
                ts.Hours, ts.Minutes, ts.Seconds,
                ts.Milliseconds / 10);

            using (StreamWriter outputFile = new StreamWriter(Path.Combine(@"C:\Users\Mikołaj\WORD_ADDIN_PROJECT", "TestHighlightTimes.txt"), true))
            {
                outputFile.WriteLine(elapsedTime);
            }
        }

        private void WrapItem(Microsoft.Office.Tools.Word.Document extendedDocument, string tag, Range range)
        {
            try
            {
                int BookmarkNumber = this.Application.ActiveDocument.Bookmarks.Count;

                string bookmarkName = "_" + BookmarkNumber.ToString();
                if (tag.EndsWith("-1"))
                {
                    bookmarkName += "_intent_";
                }
                else if (tag.EndsWith("-2"))
                {
                    bookmarkName += "_entity_";
                }
                else
                {
                    bookmarkName += "_notspecified_";
                }
                string NewTag = Regex.Replace(tag, "-", "_");
                bookmarkName += NewTag;

                Microsoft.Office.Tools.Word.Bookmark bookmark = extendedDocument.Controls.AddBookmark(range, bookmarkName);
                bookmark.Tag = tag;

                if ((tag.EndsWith("-1")))
                {
                    bookmark.Selected += new Microsoft.Office.Tools.Word.SelectionEventHandler((sender, e) => bookmark_Selected(sender, e, extendedDocument, bookmark));
                    bookmark.Deselected += new Microsoft.Office.Tools.Word.SelectionEventHandler((sender, e) => bookmark_Deselected(sender, e, bookmark));
                }
            }
            catch (Exception ex)
            {
                Utilities.Notification(ex.Message);
            }
        }

        void bookmark_Selected(object sender, Microsoft.Office.Tools.Word.SelectionEventArgs e, Microsoft.Office.Tools.Word.Document extendedDocument, Microsoft.Office.Tools.Word.Bookmark bookmark)
        {
            //Range range = bookmark.Range;

            //foreach (Bookmark insideBookmark in range.Bookmarks)
            //{
            //    Range newRange = insideBookmark.Range;

            //    string name = insideBookmark.Name.ToString();
            //    string NewTag = Regex.Replace(name, "_[0-9]+_entity_", "");
            //    NewTag = Regex.Replace(NewTag, "_[0-9]+_intent_", "");
            //    NewTag = Regex.Replace(NewTag, "_[0-9]+_notspecified_", "");
            //    NewTag = Regex.Replace(NewTag, "_", "-");

            //    addContentControlFromBookmark(extendedDocument, newRange, NewTag);
            //}

            string name = bookmark.Name.ToString();
            string NewTag = Regex.Replace(name, "_[0-9]+_entity_", "");
            NewTag = Regex.Replace(NewTag, "_[0-9]+_intent_", "");
            NewTag = Regex.Replace(NewTag, "_[0-9]+_notspecified_", "");
            NewTag = Regex.Replace(NewTag, "_", "-");

            currentBookmark = bookmark;
            Globals.Ribbons.Ribbon1.CurBMtextLabel.Label = currentBookmark.Text;
            Globals.Ribbons.Ribbon1.CurBMentLabel.Label = NewTag;

            if (currentBookmark.Text.Length <= 1024)
            {
                Globals.Ribbons.Ribbon1.CurBMtextLabel.Label = currentBookmark.Text;
            }
            else
            {
                Globals.Ribbons.Ribbon1.CurBMtextLabel.Label = "..." + currentBookmark.Text.Substring(currentBookmark.Text.Length - 1020);
            }

            if (NewTag.Length <= 1024)
            {
                Globals.Ribbons.Ribbon1.CurBMentLabel.Label = NewTag;
            }
            else
            {
                Globals.Ribbons.Ribbon1.CurBMentLabel.Label = "..." + NewTag.Substring(currentBookmark.Text.Length - 1020);
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

        private void bookmark_Deselected(object sender, Microsoft.Office.Tools.Word.SelectionEventArgs e, Microsoft.Office.Tools.Word.Bookmark bookmark)
        {
            Range bookmarkRange = bookmark.Range;

            foreach (ContentControl insideControl in bookmarkRange.ContentControls)
            {
                insideControl.Delete(false);
            }
            if (bookmarkRange.ParentContentControl != null)
            {
                bookmarkRange.ParentContentControl.Delete();
            }
        }

        private void addContentControlFromBookmark(Microsoft.Office.Tools.Word.Document extendedDocument, Range range, string tag)
        {
            range.Select();
            string currentTime = DateTime.Now.Ticks.ToString();
            Microsoft.Office.Tools.Word.RichTextContentControl newContentControl = extendedDocument.Controls.AddRichTextContentControl(currentTime);
            newContentControl.Tag = tag;
            newContentControl.Title = tag;
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

                        foreach (Bookmark bookmark in r.Bookmarks)
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

                    foreach (Bookmark bookmark in HLrange.Bookmarks)
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
