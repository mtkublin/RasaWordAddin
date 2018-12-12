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
            if (examps.Count != 0)
            {
                examps.Clear();
            }

            List<string> AddedSents = new List<string>();
            foreach (Range sent in Globals.ThisAddIn.Application.ActiveDocument.Sentences)
            {
                if (sent.ParentContentControl is ContentControl)
                {
                    string intTag = sent.ParentContentControl.Tag;
                    GatherEntities(intTag, sent, examps);
                    AddedSents.Add(sent.Text);
                }

                else
                {
                    bool IsNotFirstOrLast = true;
                    foreach (ContentControl control in Globals.ThisAddIn.Application.ActiveDocument.ContentControls) if (control.Tag[control.Tag.Length - 1] == '1')
                    {
                        if (sent.Text == control.Range.Sentences.First.Text || sent.Text == control.Range.Sentences.Last.Text)
                        {
                            string intTag = control.Tag;
                            GatherEntities(intTag, sent, examps);
                            IsNotFirstOrLast = false;
                            AddedSents.Add(sent.Text);
                            break;
                        }
                    }

                    if (IsNotFirstOrLast)
                    {
                        GatherEntities("empty-intent-1", sent, examps);
                    }
                }
            }

            foreach (ContentControl control in Globals.ThisAddIn.Application.ActiveDocument.ContentControls) if (control.Tag[control.Tag.Length - 1] == '1')
            {
                if (AddedSents.Contains(control.Range.Sentences.First.Text) == false)
                {
                    string intTag = control.Tag;
                    GatherEntities(intTag, control.Range.Sentences.First, examps);
                }

                if (AddedSents.Contains(control.Range.Sentences.Last.Text) == false)
                {
                    string intTag = control.Tag;
                    GatherEntities(intTag, control.Range.Sentences.Last, examps);
                }
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

            if (sentText != " " & sentText != "\r")
            {
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
