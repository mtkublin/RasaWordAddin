using System;
using System.Collections.Generic;
using System.IO;
using System.Timers;
using System.Windows.Forms;
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

            //string req_id_response = response.Content;
            //string reqID = JsonConvert.DeserializeObject<string>(req_id_response);

            //var newRequest = new RestRequest("api/testres/{id}", Method.GET);
            //newRequest.AddParameter("req_id", reqID, ParameterType.UrlSegment);
            //newRequest.AddUrlSegment("id", reqID);

            //IRestResponse newResponse = client.Execute(newRequest);

            //Microsoft.Office.Tools.Word.Document extendedDocument = Globals.Factory.GetVstoObject(this.Application.ActiveDocument);

            //TestReqTimer = new System.Timers.Timer(3000);
            //TestReqTimer.AutoReset = true;
            //TestReqTimer.Elapsed += (sender, e) => TestOnTimedEvent(sender, e, reqID, client, this.Application, extendedDocument);
            //TestReqTimer.Enabled = true;
        }

        //private System.Timers.Timer TestReqTimer;

        //private void TestOnTimedEvent(object source, ElapsedEventArgs e, string reqID, RestClient client, Microsoft.Office.Interop.Word.Application application, Microsoft.Office.Tools.Word.Document extendedDocument)
        //{
        //    var newRequest = new RestRequest("api/testres/{id}", Method.GET);
        //    newRequest.AddParameter("req_id", reqID, ParameterType.UrlSegment);
        //    newRequest.AddUrlSegment("id", reqID);

        //    IRestResponse newResponse = client.Execute(newRequest);
        //    string JSONresultDoc = newResponse.Content.ToString();
        //    List<SentenceObject> SentObjList = JsonConvert.DeserializeObject<List<SentenceObject>>(JSONresultDoc);

        //    if (SentObjList.Count != 0)
        //    {
        //        WrapFromJSON(JSONresultDoc, application, extendedDocument);
        //        //TestReqTimer.Stop();
        //        //TestReqTimer.Dispose();
        //    }
        //}

        private void WrapFromJSON(string JSONresult)
        {
            List<SentenceObject> SentObjList = JsonConvert.DeserializeObject<List<SentenceObject>>(JSONresult);
            SentObjList.Reverse();
            if (string.IsNullOrEmpty(JSONresult)) return;
            
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
            //TestReqTimer.Stop();
            //TestReqTimer.Dispose();
        }

        private void WrapItem(Microsoft.Office.Tools.Word.Document extendedDocument, string tag, Range range)
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

        private class TextToExportObject
        {
            public List<String> SENTS { get; set; }

            public TextToExportObject(List<String> DataToPass)
            {
                SENTS = DataToPass;
            }
        }

        private class FinalTestDataExportObject
        {
            public TextToExportObject DATA { get; set; }

            public FinalTestDataExportObject(TextToExportObject DataToPass)
            {
                DATA = DataToPass;
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

        public class TestDataReqIDobject
        {
            public string req_id { get; set; }
            public string mongo_id { get; set; }
            public string status { get; set; }

        }

    }
}
