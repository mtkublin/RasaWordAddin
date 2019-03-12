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
        public Dictionary<string, Range> bmRangesPriorToTestDict;

        public void TestDoc()
        {
            Document activeDocument = Application.ActiveDocument;
            var extendedDocument = Globals.Factory.GetVstoObject(activeDocument);

            bmRangesPriorToTestDict = new Dictionary<string, Range>();

            foreach (Range range in activeDocument.StoryRanges)
            {
                Globals.ThisAddIn.UnhighlightControl(range);
            }

            foreach (Bookmark existingBM in Application.ActiveDocument.Bookmarks)
            {
                Microsoft.Office.Tools.Word.Bookmark VSTOexistingBM = extendedDocument.Controls[existingBM.Name] as Microsoft.Office.Tools.Word.Bookmark;
                bmRangesPriorToTestDict[existingBM.Name] = existingBM.Range;
                VSTOexistingBM.Delete();
            }

            //List<String> SentsToExport = new List<String>();
            //foreach (Range sent in Globals.ThisAddIn.Application.ActiveDocument.Sentences)
            //{
            //    SentsToExport.Add(sent.Text);
            //}
            //TextToExportObject textToExportObject = new TextToExportObject(SentsToExport);
            //FinalTestDataExportObject finalExportData = new FinalTestDataExportObject(textToExportObject);
            //var jsonTestObject = JsonConvert.SerializeObject(finalExportData);

            //var client = new RestClient("http://127.0.0.1:6000");
            //var request = new RestRequest("api/testdata", Method.POST);
            //request.AddParameter("application/json; charset=utf-8", jsonTestObject, ParameterType.RequestBody);
            //request.RequestFormat = DataFormat.Json;
            //IRestResponse response = client.Execute(request);
            //string JSONresultDoc = response.Content.ToString();

            string AppDir = System.AppDomain.CurrentDomain.BaseDirectory.ToString();
            string TestResDir = AppDir.Substring(0, AppDir.Length - 21) + @"Docs\test_result.json";
            string JSONresultDoc;
            using (StreamReader r = new StreamReader(TestResDir))
            {
                JSONresultDoc = r.ReadToEnd();
            }

            WrapFromJSON(JSONresultDoc);

            Globals.Ribbons.Ribbon1.reverseTestBTN.Enabled = true;
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

        public void ReverseTest()
        {
            Microsoft.Office.Interop.Word.Document activeDocument = Globals.ThisAddIn.Application.ActiveDocument;
            var extendedDocument = Globals.Factory.GetVstoObject(activeDocument);

            foreach (Microsoft.Office.Interop.Word.Range range in activeDocument.StoryRanges)
            {
                Globals.ThisAddIn.UnhighlightControl(range);
            }

            foreach (Microsoft.Office.Interop.Word.Bookmark existingBM in activeDocument.Bookmarks)
            {
                Microsoft.Office.Tools.Word.Bookmark VSTOexistingBM = extendedDocument.Controls[existingBM.Name] as Microsoft.Office.Tools.Word.Bookmark;
                VSTOexistingBM.Delete();
            }

            Dictionary<string, Microsoft.Office.Interop.Word.Range> bmsDict = Globals.ThisAddIn.bmRangesPriorToTestDict;
            foreach (string bmName in bmsDict.Keys)
            {
                extendedDocument.Controls.AddBookmark(bmsDict[bmName], bmName);
            }

            Globals.ThisAddIn.HighlightBookmarksInVisibleRange();
            Globals.Ribbons.Ribbon1.reverseTestBTN.Enabled = false;
            Globals.ThisAddIn.currentBookmark = null;
            Globals.Ribbons.Ribbon1.CurBMtextLabel.Label = "";
            Globals.Ribbons.Ribbon1.CurBMentLabel.Label = "";
            Globals.Ribbons.Ribbon1.IntOrEntLabel.Label = "";
        }
    }
}
