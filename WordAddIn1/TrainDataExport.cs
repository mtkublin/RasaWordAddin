using System;
using System.Collections.Generic;
using System.IO;
using System.Timers;
using Microsoft.Office.Interop.Word;
using Newtonsoft.Json;
using RestSharp;

namespace WordAddIn1
{
    public partial class ThisAddIn
    {
        private RasaNLUdata rasaData { get; set; }

        public void InitiateTraining(RestClient client, string TrainProjectName, string TrainModelName, string ModelPath = null)
        {
            Globals.Ribbons.Ribbon1.ProjectDropDown.Enabled = false;
            Globals.Ribbons.Ribbon1.ProjectAddButton.Enabled = false;
            Globals.Ribbons.Ribbon1.TestModelDropDown.Enabled = false;
            Globals.Ribbons.Ribbon1.TestButton.Enabled = false;
            Globals.Ribbons.Ribbon1.TrainingButton.Enabled = false;
            Globals.Ribbons.Ribbon1.LocalStorageButton.Enabled = false;
            Globals.Ribbons.Ribbon1.AzureStorageButton.Enabled = false;
            Globals.Ribbons.Ribbon1.SetDirButton.Enabled = false;

            var examps = new List<Examp> { };
            if(examps.Count != 0)
            {
                examps.Clear();
            }

            List<string> SentsWithIntent = new List<string>();

            foreach (ContentControl intent in Globals.ThisAddIn.Application.ActiveDocument.ContentControls)
            {
                string intTag = intent.Tag;
                char intLevelIndicator = intTag[intTag.Length - 1];
                
                if (intLevelIndicator is '1')
                {
                    if(intent.Range.Sentences.Count == 0 || intent.Range.Sentences.Count == 1)
                    {
                        SentsWithIntent.Add(intent.Range.Text);
                        GatherEntities(intTag, intent.Range, examps);
                    }
                    else
                    {
                        bool isLastSent = false;
                        foreach(Range subInt in intent.Range.Sentences)
                        {
                            if (subInt.Text == " ")
                            {
                                isLastSent = true;
                                break;
                            }
                            SentsWithIntent.Add(subInt.Text);
                            GatherEntities(intTag, subInt, examps);
                        }

                        if (isLastSent)
                        {
                            GatherEntities(intTag, intent.Range.Sentences.Last, examps);

                        }
                    }
                }
            }

            foreach (Range sent in Globals.ThisAddIn.Application.ActiveDocument.Sentences) if (SentsWithIntent.Contains(sent.Text) == false)
            {
                GatherEntities("empty-intent-1", sent, examps);
            }

            TrainData tData = new TrainData(examps);

            RasaNLUdata rasaData = new RasaNLUdata(tData, ModelPath);

            FinalDataObject DataObjectForApi = new FinalDataObject(rasaData);
            var jsonObject = JsonConvert.SerializeObject(DataObjectForApi);

            var request = new RestRequest("api/traindata/{project}/{model}", Method.POST);
            request.AddParameter("application/json; charset=utf-8", jsonObject, ParameterType.RequestBody);
            request.AddParameter("project", TrainProjectName, ParameterType.UrlSegment);
            request.AddUrlSegment("project", TrainProjectName);
            request.AddParameter("model", TrainModelName, ParameterType.UrlSegment);
            request.AddUrlSegment("model", TrainModelName);
            request.RequestFormat = DataFormat.Json;
            IRestResponse response = client.Execute(request);
            string reqID = JsonConvert.DeserializeObject<string>(response.Content.ToString());

            TrainingStatusCheckTimer = new Timer(3000);
            TrainingStatusCheckTimer.AutoReset = true;
            TrainingStatusCheckTimer.Elapsed += (sender, e) => CheckTrainingStatus(sender, e, client, reqID);
            TrainingStatusCheckTimer.Enabled = true;
        }

        private void GatherEntities(string intTag, Range sent, List<Examp> examps)
        {
            string sentInt = intTag;
            string sentText = sent.Text;
            int intentStart = sent.Start;

            var entities = new List<Ent> { };
            int EntNumber = 0;
            foreach (ContentControl ent in sent.ContentControls)
            {
                string entTag = ent.Tag;
                char entLevelIndicator = entTag[entTag.Length - 1];

                if (entLevelIndicator is '2')
                {
                    int st = ent.Range.Start - intentStart - 1 - EntNumber;
                    int en = ent.Range.End - intentStart - 1 - EntNumber;
                    string val = ent.Range.Text;
                    string tag = entTag;

                    Ent entity = new Ent(st, en, val, tag);
                    entities.Add(entity);

                    EntNumber += 2;
                }
            }
            if (sentText != " " & sentText != "\r")
            {
                Examp examp = new Examp(sentText, sentInt, entities);
                examps.Add(examp);
            }
        }

        private static Timer TrainingStatusCheckTimer;

        private static void CheckTrainingStatus(object source, ElapsedEventArgs e, RestClient client, string reqID)
        {
            var newRequest = new RestRequest("api/traindata/isfinished/{req_id}", Method.GET);
            newRequest.AddParameter("req_id", reqID, ParameterType.UrlSegment);
            newRequest.AddUrlSegment("req_id", reqID);

            IRestResponse newResponse = client.Execute(newRequest);
            string IsFinished = JsonConvert.DeserializeObject<string>(newResponse.Content.ToString());

            if (IsFinished == "True")
            {
                Globals.Ribbons.Ribbon1.ProjectDropDown.Enabled = true;
                Globals.Ribbons.Ribbon1.ProjectAddButton.Enabled = true;
                Globals.Ribbons.Ribbon1.TestModelDropDown.Enabled = true;
                Globals.Ribbons.Ribbon1.TestButton.Enabled = true;
                Globals.Ribbons.Ribbon1.TrainingButton.Enabled = true;
                Globals.Ribbons.Ribbon1.LocalStorageButton.Enabled = true;
                Globals.Ribbons.Ribbon1.AzureStorageButton.Enabled = true;

                if (Globals.Ribbons.Ribbon1.LocalStorageButton.Checked == true)
                {
                    Globals.Ribbons.Ribbon1.SetDirButton.Enabled = true;
                }

                TrainingStatusCheckTimer.Stop();
                TrainingStatusCheckTimer.Dispose();
            }
        }
    }
}
