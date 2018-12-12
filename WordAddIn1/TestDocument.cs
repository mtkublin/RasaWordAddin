using System;
using System.Collections.Generic;
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
                    ContentSubstraction = 1;
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
    }
}
