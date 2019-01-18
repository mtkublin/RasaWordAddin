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
            var handler = new Microsoft.Office.Tools.Word.SelectionEventHandler(ThisDocument_SelectionChange);
            vstoDoc.SelectionChange -= handler;
            vstoDoc.SelectionChange += handler;
        }

        Range CurrentIntentRange = null;
        Range PreviousIntentRange = null;

        void ThisDocument_SelectionChange(object sender, Microsoft.Office.Tools.Word.SelectionEventArgs e)
        {
            Selection currentSelection = this.Application.ActiveDocument.ActiveWindow.Selection;

            if (PreviousIntentRange != null)
            {
                if (currentSelection.End < PreviousIntentRange.Start || currentSelection.Start > PreviousIntentRange.End)
                {
                    foreach (ContentControl insideControl in PreviousIntentRange.ContentControls)
                    {
                        insideControl.Delete(false);
                    }
                    if (PreviousIntentRange.ParentContentControl != null)
                    {
                        PreviousIntentRange.ParentContentControl.Delete();
                    }

                    PreviousIntentRange = null;
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
                HighlightContentControl(tag, range);

                int BookmarkNumber = this.Application.ActiveDocument.Bookmarks.Count;

                string bookmarkName = "_" + BookmarkNumber.ToString();
                if (tag.EndsWith("-1"))
                {
                    bookmarkName  += "_intent_";
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
                }
            }
            catch (Exception ex)
            {
                Utilities.Notification(ex.Message);
            }
        }

        void bookmark_Selected(object sender, Microsoft.Office.Tools.Word.SelectionEventArgs e, Microsoft.Office.Tools.Word.Document extendedDocument, Microsoft.Office.Tools.Word.Bookmark bookmark)
        {
            Range range = bookmark.Range;

            foreach (Bookmark insideBookmark in range.Bookmarks)
            {
                Range newRange = insideBookmark.Range;

                string name = insideBookmark.Name.ToString();
                string NewTag = Regex.Replace(name, "_[0-9]+_entity_", "");
                NewTag = Regex.Replace(NewTag, "_[0-9]+_intent_", "");
                NewTag = Regex.Replace(NewTag, "_[0-9]+_notspecified_", "");
                NewTag = Regex.Replace(NewTag, "_", "-");

                addContentControlFromBookmark(extendedDocument, newRange, NewTag);
            }
        }

        private void addContentControlFromBookmark(Microsoft.Office.Tools.Word.Document extendedDocument, Range range, string tag)
        {
            if (tag.EndsWith("-1"))
            {
                PreviousIntentRange = CurrentIntentRange;
                CurrentIntentRange = range;
            }

            range.Select();
            string currentTime = DateTime.Now.Ticks.ToString();
            Microsoft.Office.Tools.Word.RichTextContentControl newContentControl = extendedDocument.Controls.AddRichTextContentControl(currentTime);
            newContentControl.Tag = tag;
            newContentControl.Title = tag;
        }
    }
}
