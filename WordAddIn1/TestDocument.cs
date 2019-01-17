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

                char entLevelIndicator = tag[tag.Length - 1];
                string NewTag = Regex.Replace(tag, "-", "_");

                string bookmarkName;
                if (entLevelIndicator is '1')
                {
                    bookmarkName = "_" + BookmarkNumber.ToString() + "_intent_" + NewTag;
                }
                else if (entLevelIndicator is '2')
                {
                    bookmarkName = "_" + BookmarkNumber.ToString() + "_entity_" + NewTag;
                }
                else
                {
                    bookmarkName = "_" + BookmarkNumber.ToString() + "_notspecified_" + NewTag;
                }

                Microsoft.Office.Tools.Word.Bookmark control = extendedDocument.Controls.AddBookmark(range, bookmarkName);
                control.Tag = tag;
                control.Text = range.Text;

                if (entLevelIndicator is '1')
                {
                    control.Selected += new Microsoft.Office.Tools.Word.SelectionEventHandler((sender, e) => bookmark_Selected(sender, e, extendedDocument, control));
                }
                //control.Selected += new Microsoft.Office.Tools.Word.SelectionEventHandler((sender, e) => bookmark_Selected(sender, e, extendedDocument, control));
            }
            catch (Exception ex)
            {
                Utilities.Notification(ex.Message);
            }
        }

        void bookmark_Selected(object sender, Microsoft.Office.Tools.Word.SelectionEventArgs e, Microsoft.Office.Tools.Word.Document extendedDocument, Microsoft.Office.Tools.Word.Bookmark bookmark)
        {
            //HighlightContentControl(tag, range);
            Range range = bookmark.Range;
            string tag = bookmark.Tag.ToString();
            addContentControlFromToolsBookmark(extendedDocument, range, tag);

            foreach (Microsoft.Office.Interop.Word.Bookmark insideBookmark in bookmark.Bookmarks)
            {
                string name = bookmark.Name.ToString();
                string NewTag;
                if (name.EndsWith("_2") || name.EndsWith("_1"))
                {
                    NewTag = name.Substring(11);
                }
                else
                {
                    NewTag = name.Substring(17);
                }
                string NewTagReplaced = Regex.Replace(NewTag, "_", "-");

                Range newRange = bookmark.Range;
                addContentControlFromToolsBookmark(extendedDocument, newRange, NewTagReplaced);
            }
        }

        void control_Deselected(object sender, Microsoft.Office.Tools.Word.ContentControlExitingEventArgs e, Microsoft.Office.Tools.Word.RichTextContentControl control)
        {
            control.Delete(false);
        }

        private void addContentControlFromToolsBookmark(Microsoft.Office.Tools.Word.Document extendedDocument, Range range, string tag)
        {

            range.Select();
            string currentTime = DateTime.Now.Ticks.ToString();
            Microsoft.Office.Tools.Word.RichTextContentControl newContentControl = extendedDocument.Controls.AddRichTextContentControl(currentTime);
            newContentControl.Tag = tag;
            newContentControl.Title = tag;

            newContentControl.Exiting += new Microsoft.Office.Tools.Word.ContentControlExitingEventHandler((sender, e) => control_Deselected(sender, e, newContentControl));
        }
    }
}
