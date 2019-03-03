using Microsoft.Office.Interop.Word;
using Newtonsoft.Json;
using RestSharp;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using System.Timers;

namespace WordAddIn1
{
    public partial class ThisAddIn
    {
        private RasaNLUdata rasaData { get; set; }

        private void checkIfParentsAreBookmarks(Range objToCheck, Bookmark parentBookmark)
        {
            if (objToCheck.Parent != null)
            {
                if (objToCheck.Parent is Bookmark)
                {
                    parentBookmark = objToCheck.Parent;
                }
                else
                {
                    if (objToCheck.Parent.Range != null)
                    {
                        checkIfParentsAreBookmarks(objToCheck.Parent.Range, parentBookmark);
                    }
                }
            }
        }

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
            foreach (Range sent in Application.ActiveDocument.Sentences)
            {
                //if (sent.Parent is Bookmark)
                //{
                //    string intTag = sent.ParentContentControl.Tag;
                //    GatherEntities(intTag, sent, examps);
                //    AddedSents.Add(sent.Text);
                //}

                Bookmark parentBookmark = null;
                //checkIfParentsAreBookmarks(sent, parentBookmark);

                foreach (Bookmark bm in sent.Bookmarks)
                {
                    if (bm.Name.EndsWith("1"))
                    {
                        parentBookmark = bm;
                        break;
                    }
                }

                if (parentBookmark != null) 
                {
                    if (parentBookmark.Name.EndsWith("1"))
                    {
                        string parentBookmarkName = parentBookmark.Name.ToString();
                        string parentBookmarkTag = Regex.Replace(parentBookmarkName, "_[0-9]+_entity_", "");
                        parentBookmarkTag = Regex.Replace(parentBookmarkTag, "_[0-9]+_intent_", "");
                        parentBookmarkTag = Regex.Replace(parentBookmarkTag, "_[0-9]+_notspecified_", "");
                        parentBookmarkTag = Regex.Replace(parentBookmarkTag, "_", "-");

                        string intTag = parentBookmarkTag;
                        GatherEntities(intTag, sent, examps);
                        AddedSents.Add(sent.Text);
                    }
                }

                else
                {
                    bool IsNotFirstOrLast = true;
                    foreach (Bookmark control in Application.ActiveDocument.Bookmarks) if (control.Name.EndsWith("1"))
                    {
                        if (sent.Text == control.Range.Sentences.First.Text || sent.Text == control.Range.Sentences.Last.Text)
                        {
                            string intName = control.Name.ToString();
                            string intTag = Regex.Replace(intName, "_[0-9]+_entity_", "");
                            intTag = Regex.Replace(intTag, "_[0-9]+_intent_", "");
                            intTag = Regex.Replace(intTag, "_[0-9]+_notspecified_", "");
                            intTag = Regex.Replace(intTag, "_", "-");
                            
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

            foreach (Bookmark control in Application.ActiveDocument.Bookmarks) if (control.Name.EndsWith("1"))
            {
                string intName = control.Name.ToString();
                string intTag = Regex.Replace(intName, "_[0-9]+_entity_", "");
                intTag = Regex.Replace(intTag, "_[0-9]+_intent_", "");
                intTag = Regex.Replace(intTag, "_[0-9]+_notspecified_", "");
                intTag = Regex.Replace(intTag, "_", "-");

                if (AddedSents.Contains(control.Range.Sentences.First.Text) == false)
                {
                    //string intTag = control.Tag;
                    GatherEntities(intTag, control.Range.Sentences.First, examps);
                }

                if (AddedSents.Contains(control.Range.Sentences.Last.Text) == false)
                {
                    //string intTag = control.Tag;
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
                //int EntNumber = 0;
                foreach (Bookmark ent in sent.Bookmarks)
                {
                    string entName = ent.Name.ToString();
                    string entTag = Regex.Replace(entName, "_[0-9]+_entity_", "");
                    entTag = Regex.Replace(entTag, "_[0-9]+_intent_", "");
                    entTag = Regex.Replace(entTag, "_[0-9]+_notspecified_", "");
                    entTag = Regex.Replace(entTag, "_", "-");

                    char entLevelIndicator = entTag[entTag.Length - 1];

                    if (entLevelIndicator is '2')
                    {
                        //int st = ent.Range.Start - intentStart - 1 - EntNumber;
                        //int en = ent.Range.End - intentStart - 1 - EntNumber;
                        int st = ent.Range.Start - intentStart;
                        int en = ent.Range.End - intentStart;
                        string val = ent.Range.Text;
                        string tag = entTag;

                        Ent entity = new Ent(st, en, val, tag);
                        entities.Add(entity);

                        //EntNumber += 2;
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
