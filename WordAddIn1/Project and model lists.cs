using System.Collections.Generic;
using System.IO;
using Newtonsoft.Json;
using RestSharp;
using Microsoft.Office.Tools.Ribbon;

namespace WordAddIn1
{
    public partial class ThisAddIn
    {
        public void ChooseModelDir(System.Windows.Forms.FolderBrowserDialog ModelDirDialog, RibbonLabel ModelDirLabel, RibbonComboBox ProjectComboBox, RibbonDropDown TestProjectDropDown, RibbonComboBox ModelComboBox, RibbonDropDown TestModelDropDown)
        {
            ModelDirDialog.ShowDialog();
            ModelDirLabel.Label = ModelDirDialog.SelectedPath;

            string ModelDir = ModelDirDialog.SelectedPath;

            ProjectComboBox.Items.Clear();
            ProjectComboBox.Text = "";
            TestProjectDropDown.Items.Clear();
            ModelComboBox.Items.Clear();
            ModelComboBox.Text = "";
            TestModelDropDown.Items.Clear();

            GetProjectItemsFromDir(ModelDir, ProjectComboBox, TestProjectDropDown);

            if (TestProjectDropDown.SelectedItem != null)
            {
                string NewModelDir = ModelDir + "\\" + TestProjectDropDown.SelectedItem.Label;
                GetModelItemsFromDirDD(NewModelDir, TestModelDropDown);
            }
        }

        public void ChangeToLocalStorage(System.Windows.Forms.FolderBrowserDialog ModelDirDialog, RibbonLabel ModelDirLabel, RibbonComboBox ProjectComboBox, RibbonDropDown TestProjectDropDown, RibbonComboBox ModelComboBox, RibbonDropDown TestModelDropDown)
        {
            string ModelDir = ModelDirDialog.SelectedPath;

            ProjectComboBox.Items.Clear();
            ProjectComboBox.Text = "";
            TestProjectDropDown.Items.Clear();
            ModelComboBox.Items.Clear();
            ModelComboBox.Text = "";
            TestModelDropDown.Items.Clear();

            GetProjectItemsFromDir(ModelDir, ProjectComboBox, TestProjectDropDown);

            if (TestProjectDropDown.SelectedItem != null)
            {
                string NewModelDir = ModelDir + "\\" + TestProjectDropDown.SelectedItem.Label;
                GetModelItemsFromDirDD(NewModelDir, TestModelDropDown);
            }
        }

        public void GetProjectItemsFromDir(string ModelDir, RibbonComboBox ProjectComboBox, RibbonDropDown TestProjectDropDown)
        {
            string[] ProjFoldList = Directory.GetDirectories(ModelDir);
            foreach (string Pfold in ProjFoldList)
            {
                string ItemLabel = Pfold.Substring(ModelDir.Length + 1, Pfold.Length - ModelDir.Length - 1);

                RibbonDropDownItem folder = Globals.Factory.GetRibbonFactory().CreateRibbonDropDownItem();
                folder.Label = ItemLabel;
                ProjectComboBox.Items.Add(folder);

                RibbonDropDownItem NEWfolder = Globals.Factory.GetRibbonFactory().CreateRibbonDropDownItem();
                NEWfolder.Label = ItemLabel;
                TestProjectDropDown.Items.Add(NEWfolder);
            }
        }

        public void GetModelItemsFromDirDD(string NewModelDir, RibbonDropDown TestModelDropDown)
        {
            if (NewModelDir != null)
            {
                TestModelDropDown.Items.Clear();

                string[] ModelFoldList = Directory.GetDirectories(NewModelDir);

                foreach (string Mfold in ModelFoldList)
                {
                    string ItemLabel = Mfold.Substring(NewModelDir.Length + 1, Mfold.Length - NewModelDir.Length - 1);

                    RibbonDropDownItem NEWfolder = Globals.Factory.GetRibbonFactory().CreateRibbonDropDownItem();
                    NEWfolder.Label = ItemLabel;
                    TestModelDropDown.Items.Add(NEWfolder);
                }
            }      
        }

        public void GetModelItemsFromDirCB(string NewModelDir, RibbonComboBox TestModelDropDown)
        {
            if(NewModelDir != null)
            {
                TestModelDropDown.Items.Clear();

                if(Directory.Exists(NewModelDir))
                {
                    string[] ModelFoldList = Directory.GetDirectories(NewModelDir);

                    foreach (string Mfold in ModelFoldList)
                    {
                        string ItemLabel = Mfold.Substring(NewModelDir.Length + 1, Mfold.Length - NewModelDir.Length - 1);

                        RibbonDropDownItem NEWfolder = Globals.Factory.GetRibbonFactory().CreateRibbonDropDownItem();
                        NEWfolder.Label = ItemLabel;
                        TestModelDropDown.Items.Add(NEWfolder);
                    }
                }
            }
        }

        public void AddItemToProjectComboBox(RibbonComboBox ProjectComboBox, RibbonComboBox ModelComboBox)
        {
            List<string> ProjectsList = new List<string>();
            foreach (RibbonDropDownItem item in ProjectComboBox.Items)
            {
                string ExistingProjName = item.ToString();
                ProjectsList.Add(ExistingProjName);
            }

            string ProjectName = ProjectComboBox.Text;
            if (ProjectsList.Contains(ProjectName) == true)
            {
                ProjectComboBox.Text.Remove(0, ProjectComboBox.Text.Length);
            }
            else
            {
                RibbonDropDownItem NEWitem = Globals.Factory.GetRibbonFactory().CreateRibbonDropDownItem();
                NEWitem.Label = ProjectName;
                ProjectComboBox.Items.Add(NEWitem);

                ModelComboBox.Items.Clear();
                ModelComboBox.Text.Remove(0, ModelComboBox.Text.Length);
            }
        }

        public void AddItemToModelComboBox(RibbonComboBox ModelComboBox)
        {
            List<string> ProjectsList = new List<string>();
            foreach (RibbonDropDownItem item in ModelComboBox.Items)
            {
                string ExistingProjName = item.ToString();
                ProjectsList.Add(ExistingProjName);
            }

            string ProjectName = ModelComboBox.Text;
            if (ProjectsList.Contains(ProjectName) == true)
            {
                ModelComboBox.Text.Remove(0, ModelComboBox.Text.Length);
            }
            else
            {
                RibbonDropDownItem NEWitem = Globals.Factory.GetRibbonFactory().CreateRibbonDropDownItem();
                NEWitem.Label = ProjectName;
                ModelComboBox.Items.Add(NEWitem);
            }
        }

        public void GetProjsList(RestClient client, RibbonComboBox ProjectComboBox)
        {
            var Request = new RestRequest("api/projects", Method.GET);
            IRestResponse Response = client.Execute(Request);
            string JSONresultDoc = Response.Content.ToString();

            List<string> ProjList = JsonConvert.DeserializeObject<List<string>>(JSONresultDoc);

            foreach (string itemName in ProjList)
            {
                RibbonDropDownItem item = Globals.Factory.GetRibbonFactory().CreateRibbonDropDownItem();
                item.Label = itemName;
                ProjectComboBox.Items.Add(item);
            }
        }

        public void TESTGetProjsList(RestClient client, RibbonDropDown TestProjectDropDown, RibbonDropDown TestModelDropDown)
        {
            var Request = new RestRequest("api/projects", Method.GET);
            IRestResponse Response = client.Execute(Request);
            string JSONresultDoc = Response.Content.ToString();

            List<string> ProjList = JsonConvert.DeserializeObject<List<string>>(JSONresultDoc);

            foreach (string itemName in ProjList)
            {
                RibbonDropDownItem item = Globals.Factory.GetRibbonFactory().CreateRibbonDropDownItem();
                item.Label = itemName;
                TestProjectDropDown.Items.Add(item);
            }

            string ProjectName = TestProjectDropDown.SelectedItem.Label;
            Globals.ThisAddIn.TESTGetModelsList(client, ProjectName, TestModelDropDown);
        }

        public void GetModelsList(RestClient client, string ProjectName, RibbonComboBox ModelComboBox)
        {
            var newRequest = new RestRequest("api/models/{project}", Method.GET);
            newRequest.AddParameter("project", ProjectName, ParameterType.UrlSegment);
            newRequest.AddUrlSegment("project", ProjectName);
            IRestResponse newResponse = client.Execute(newRequest);
            string newJSONresultDoc = newResponse.Content.ToString();

            List<string> ModelList = JsonConvert.DeserializeObject<List<string>>(newJSONresultDoc);

            ModelComboBox.Items.Clear();

            foreach (string itemName in ModelList)
            {
                RibbonDropDownItem item = Globals.Factory.GetRibbonFactory().CreateRibbonDropDownItem();
                item.Label = itemName;
                ModelComboBox.Items.Add(item);
            }
        }

        public void TESTGetModelsList(RestClient client, string ProjectName, RibbonDropDown TestModelDropDown)
        {
            var newRequest = new RestRequest("api/models/{project}", Method.GET);
            newRequest.AddParameter("project", ProjectName, ParameterType.UrlSegment);
            newRequest.AddUrlSegment("project", ProjectName);
            IRestResponse newResponse = client.Execute(newRequest);
            string newJSONresultDoc = newResponse.Content.ToString();
            List<string> ModelList = JsonConvert.DeserializeObject<List<string>>(newJSONresultDoc);

            TestModelDropDown.Items.Clear();

            foreach (string itemName in ModelList)
            {
                RibbonDropDownItem item = Globals.Factory.GetRibbonFactory().CreateRibbonDropDownItem();
                item.Label = itemName;
                TestModelDropDown.Items.Add(item);
            }

            string ModelName = TestModelDropDown.SelectedItem.Label;
            Globals.ThisAddIn.UpdateInterpreter(client, ProjectName, ModelName);
        }

        public void UpdateInterpreter(RestClient client, string ProjectName, string ModelName, string model_path = "")
        {
            if(model_path != "")
            {
                ModelPathDataObject DataObjectForApi = new ModelPathDataObject(model_path);
                var jsonObject = JsonConvert.SerializeObject(DataObjectForApi);

                var newRequest = new RestRequest("api/interpreter/local/{project}/{model}", Method.POST);
                newRequest.AddParameter("project", ProjectName, ParameterType.UrlSegment);
                newRequest.AddUrlSegment("project", ProjectName);
                newRequest.AddParameter("model", ModelName, ParameterType.UrlSegment);
                newRequest.AddUrlSegment("model", ModelName);
                newRequest.AddParameter("application/json; charset=utf-8", jsonObject, ParameterType.RequestBody);
                IRestResponse newResponse = client.Execute(newRequest);
            }
            else
            {
                var newRequest = new RestRequest("api/interpreter/azure/{project}/{model}", Method.POST);
                newRequest.AddParameter("project", ProjectName, ParameterType.UrlSegment);
                newRequest.AddUrlSegment("project", ProjectName);
                newRequest.AddParameter("model", ModelName, ParameterType.UrlSegment);
                newRequest.AddUrlSegment("model", ModelName);
                IRestResponse newResponse = client.Execute(newRequest);
            }
        }

        private class ModelPathDataObject
        {
            public string DATA { get; set; }

            public ModelPathDataObject(string DataToPass)
            {
                DATA = DataToPass;
            }
        }

    }
}
