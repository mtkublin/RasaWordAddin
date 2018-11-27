using System;
using System.Collections.Generic;
using System.IO;
using System.Windows.Forms;
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
            //string TextToExport = "";
            foreach (Range sent in Globals.ThisAddIn.Application.ActiveDocument.Sentences)
            {
                SentsToExport.Add(sent.Text);
            }
            //TextToExportObject textToExportObject = new TextToExportObject(TextToExport);
            TextToExportObject textToExportObject = new TextToExportObject(SentsToExport);
            FinalTestDataExportObject finalExportData = new FinalTestDataExportObject(textToExportObject);
            var jsonTestObject = JsonConvert.SerializeObject(finalExportData);

            var client = new RestClient("http://127.0.0.1:6000");

            var request = new RestRequest("api/testdata", Method.POST);
            request.AddParameter("application/json; charset=utf-8", jsonTestObject, ParameterType.RequestBody);
            request.RequestFormat = DataFormat.Json;
            IRestResponse response = client.Execute(request);
            string req_id_response = response.Content;
            TestDataReqIDobject ReqIDobject = JsonConvert.DeserializeObject<TestDataReqIDobject>(req_id_response);
            string reqID = ReqIDobject.req_id;


            var newRequest = new RestRequest("api/testres/{id}", Method.GET);
            newRequest.AddParameter("req_id", reqID, ParameterType.UrlSegment);
            newRequest.AddUrlSegment("id", reqID);
            IRestResponse newResponse = client.Execute(newRequest);
            string JSONresultDoc = newResponse.Content.ToString();

            WrapFromJSON(JSONresultDoc);
        }

        public void WrapFromJSON(string JSONresult)
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
                int sentStart = sentEnd - sent.text.Length + 1;
                //if (sentStart < 0)
                //{
                //    sentStart = 0;
                //}

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
