using System.Collections.Generic;
using System.IO;
using Newtonsoft.Json;
using RestSharp;
using Microsoft.Office.Tools.Ribbon;

namespace WordAddIn1
{
    public partial class ThisAddIn
    {
        public void ChooseModelDir(System.Windows.Forms.FolderBrowserDialog ModelDirDialog, RibbonEditBox ModelDirBox, RibbonComboBox ProjectComboBox, RibbonDropDown TestModelDropDown)
        {
            ModelDirDialog.ShowDialog();
            ModelDirBox.Text = ModelDirDialog.SelectedPath;
            string ModelDir = ModelDirDialog.SelectedPath;

            if (Directory.Exists(ModelDir + "\\MODELS") == false)
            {
                Directory.CreateDirectory(ModelDir + "\\MODELS");
            }
            if (Directory.Exists(ModelDir + "\\TRAIN_DATA") == false)
            {
                Directory.CreateDirectory(ModelDir + "\\TRAIN_DATA");
            }

            ChangeToLocalStorage(ModelDir, ProjectComboBox, TestModelDropDown);
        }

        public void ChangeToLocalStorage(string ModelDir, RibbonComboBox ProjectComboBox, RibbonDropDown TestModelDropDown)
        {
            ProjectComboBox.Items.Clear();
            ProjectComboBox.Text = "";
            TestModelDropDown.Items.Clear();

            GetProjectItemsFromDir(ModelDir, ProjectComboBox);
        }

        public void GetProjectItemsFromDir(string ModelDir, RibbonComboBox ProjectComboBox)
        {
            string[] ProjFoldList = Directory.GetDirectories(ModelDir + "\\MODELS");
            foreach (string Pfold in ProjFoldList)
            {
                string ItemLabel = Pfold.Substring(ModelDir.Length + 8, Pfold.Length - ModelDir.Length - 8);

                RibbonDropDownItem folder = Globals.Factory.GetRibbonFactory().CreateRibbonDropDownItem();
                folder.Label = ItemLabel;
                ProjectComboBox.Items.Add(folder);
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

        public void AddItemToProjectComboBox(RibbonButton ExportTXTbtn, RibbonButton WrapFromTestBtn, RestClient client, System.Windows.Forms.FolderBrowserDialog ModelDirDialog, RibbonComboBox ProjectComboBox, RibbonDropDown TestModelDropDown, RibbonToggleButton AzureStorageButton, RibbonToggleButton LocalStorageButton)
        {
            List<string> ProjectsList = new List<string>();
            foreach (RibbonDropDownItem item in ProjectComboBox.Items)
            {
                string ExistingProjName = item.ToString();
                ProjectsList.Add(ExistingProjName);
            }

            string ProjectName = ProjectComboBox.Text;
            if (ProjectsList.Contains(ProjectName) != true)
            {
                RibbonDropDownItem NEWitem = Globals.Factory.GetRibbonFactory().CreateRibbonDropDownItem();
                NEWitem.Label = ProjectName;
                ProjectComboBox.Items.Add(NEWitem);
                TestModelDropDown.Items.Clear();

                TestModelDropDown.Enabled = false;
                WrapFromTestBtn.Enabled = false;
                ExportTXTbtn.Enabled = true;

                Directory.CreateDirectory(ModelDirDialog.SelectedPath + "\\MODELS\\" + ProjectName);
            }
            else
            {
                if (AzureStorageButton.Checked == true)
                {
                    Globals.ThisAddIn.TESTGetModelsList(client, ProjectName, TestModelDropDown);
                }
                else if (LocalStorageButton.Checked == true)
                {
                    Globals.ThisAddIn.GetModelItemsFromDirDD(ModelDirDialog.SelectedPath + "\\MODELS\\" + ProjectName, TestModelDropDown);
                    if (TestModelDropDown.Items.Count != 0)
                    {
                        string ModelName = TestModelDropDown.SelectedItem.Label;
                        string ModelPath = ModelDirDialog.SelectedPath + "\\MODELS\\" + ProjectName + "\\" + ModelName;
                        Globals.ThisAddIn.UpdateInterpreter(client, ProjectName, ModelName, false, ModelPath);
                    }
                }

                if (TestModelDropDown.Items.Count != 0)
                {
                    TestModelDropDown.Enabled = true;
                    WrapFromTestBtn.Enabled = true;
                }
                else
                {
                    TestModelDropDown.Enabled = false;
                    WrapFromTestBtn.Enabled = false;
                }
                ExportTXTbtn.Enabled = true;
            }
        }

        public void GetProjsListAzure(RestClient client, RibbonComboBox ProjectComboBox)
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

        public void UpdateInterpreter(RestClient client, string ProjectName, string ModelName, bool ForceUpdate = false, string model_path = "")
        {
            if(model_path != "")
            {
                ModelPathDataObject DataObjectForApi = new ModelPathDataObject(model_path);
                var jsonObject = JsonConvert.SerializeObject(DataObjectForApi);

                var newRequest = new RestRequest("api/interpreter/local/{project}/{model}/{force}", Method.POST);
                newRequest.AddParameter("project", ProjectName, ParameterType.UrlSegment);
                newRequest.AddUrlSegment("project", ProjectName);
                newRequest.AddParameter("model", ModelName, ParameterType.UrlSegment);
                newRequest.AddUrlSegment("model", ModelName);
                newRequest.AddParameter("force", ForceUpdate.ToString(), ParameterType.UrlSegment);
                newRequest.AddUrlSegment("force", ForceUpdate.ToString());
                newRequest.AddParameter("application/json; charset=utf-8", jsonObject, ParameterType.RequestBody);
                IRestResponse newResponse = client.Execute(newRequest);
            }
            else
            {
                var newRequest = new RestRequest("api/interpreter/azure/{project}/{model}/{force}", Method.POST);
                newRequest.AddParameter("project", ProjectName, ParameterType.UrlSegment);
                newRequest.AddUrlSegment("project", ProjectName);
                newRequest.AddParameter("model", ModelName, ParameterType.UrlSegment);
                newRequest.AddUrlSegment("model", ModelName);
                newRequest.AddParameter("force", ForceUpdate.ToString(), ParameterType.UrlSegment);
                newRequest.AddUrlSegment("force", ForceUpdate.ToString());
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
