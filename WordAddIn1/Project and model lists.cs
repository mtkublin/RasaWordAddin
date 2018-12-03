using System.Collections.Generic;
using System.IO;
using Newtonsoft.Json;
using RestSharp;
using Microsoft.Office.Tools.Ribbon;

namespace WordAddIn1
{
    public partial class ThisAddIn
    {
        public void ChooseModelDir(RestClient client, System.Windows.Forms.FolderBrowserDialog ModelDirDialog, RibbonEditBox ModelDirBox, RibbonDropDown ProjectDropDown, RibbonDropDown TestModelDropDown)
        {
            ModelDirDialog.ShowDialog();
            string ModelDir = ModelDirDialog.SelectedPath;
            ModelDirBox.Text = ModelDir;

            if (Directory.Exists(ModelDir + "\\MODELS") == false)
            {
                Directory.CreateDirectory(ModelDir + "\\MODELS");
            }
            if (Directory.Exists(ModelDir + "\\TRAIN_DATA") == false)
            {
                Directory.CreateDirectory(ModelDir + "\\TRAIN_DATA");
            }

            ChangeToLocalStorage(client, ModelDir, ProjectDropDown, TestModelDropDown);
        }

        public void ChangeToLocalStorage(RestClient client, string ModelDir, RibbonDropDown ProjectDropDown, RibbonDropDown TestModelDropDown)
        {
            ProjectDropDown.Items.Clear();
            ProjectDropDown.SelectedItem = null;
            TestModelDropDown.Items.Clear();

            GetProjectItemsFromDir(ModelDir, ProjectDropDown);

            if (ProjectDropDown.SelectedItem != null)
            {
                GetModelItemsFromDirDD(ModelDir + "\\MODELS\\" + ProjectDropDown.SelectedItem.Label, TestModelDropDown);

                if (TestModelDropDown.Items.Count != 0 )
                {
                    UpdateInterpreter(client, ProjectDropDown.SelectedItem.Label, TestModelDropDown.SelectedItem.Label, false, ModelDir + "\\MODELS\\" + ProjectDropDown.SelectedItem.Label + "\\" + TestModelDropDown.SelectedItem.Label);
                }
            }
        }

        public void GetProjectItemsFromDir(string ModelDir, RibbonDropDown ProjectDropDown)
        {
            string[] ProjFoldList = Directory.GetDirectories(ModelDir + "\\MODELS");
            foreach (string Pfold in ProjFoldList)
            {
                string ItemLabel = Pfold.Substring(ModelDir.Length + 8, Pfold.Length - ModelDir.Length - 8);

                RibbonDropDownItem folder = Globals.Factory.GetRibbonFactory().CreateRibbonDropDownItem();
                folder.Label = ItemLabel;
                ProjectDropDown.Items.Add(folder);
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

        public void AddItemToProjectDropDown(RibbonButton ExportTXTbtn, RibbonButton WrapFromTestBtn, RestClient client, System.Windows.Forms.FolderBrowserDialog ModelDirDialog, RibbonDropDown ProjectDropDown, RibbonDropDown TestModelDropDown, RibbonToggleButton AzureStorageButton, RibbonToggleButton LocalStorageButton)
        {
            List<string> ProjectsList = new List<string>();
            foreach (RibbonDropDownItem item in ProjectDropDown.Items)
            {
                string ExistingProjName = item.ToString();
                ProjectsList.Add(ExistingProjName);
            }

            string ProjectName = ProjectDropDown.SelectedItem.Label;
            if (ProjectsList.Contains(ProjectName) != true)
            {
                RibbonDropDownItem NEWitem = Globals.Factory.GetRibbonFactory().CreateRibbonDropDownItem();
                NEWitem.Label = ProjectName;
                ProjectDropDown.Items.Add(NEWitem);
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

        public void GetProjsListAzure(RestClient client, RibbonDropDown ProjectDropDown)
        {
            var Request = new RestRequest("api/projects", Method.GET);
            IRestResponse Response = client.Execute(Request);
            string JSONresultDoc = Response.Content.ToString();

            List<string> ProjList = JsonConvert.DeserializeObject<List<string>>(JSONresultDoc);

            foreach (string itemName in ProjList)
            {
                RibbonDropDownItem item = Globals.Factory.GetRibbonFactory().CreateRibbonDropDownItem();
                item.Label = itemName;
                ProjectDropDown.Items.Add(item);
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

            if (TestModelDropDown.Items.Count != 0)
            {
                string ModelName = TestModelDropDown.SelectedItem.Label;
                Globals.ThisAddIn.UpdateInterpreter(client, ProjectName, ModelName);
            }

            else
            {
                TestModelDropDown.Enabled = false;
                Globals.Ribbons.Ribbon1.WrapFromTestBtn.Enabled = false;
            }
        }

        public void UpdateInterpreter(RestClient client, string ProjectName, string ModelName, bool ForceUpdate = false, string model_path = "")
        {
            var newRequest = new RestRequest("api/interpreter/{project}/{model}/{force}", Method.POST);
            newRequest.AddParameter("project", ProjectName, ParameterType.UrlSegment);
            newRequest.AddUrlSegment("project", ProjectName);
            newRequest.AddParameter("model", ModelName, ParameterType.UrlSegment);
            newRequest.AddUrlSegment("model", ModelName);
            newRequest.AddParameter("force", ForceUpdate.ToString(), ParameterType.UrlSegment);
            newRequest.AddUrlSegment("force", ForceUpdate.ToString());

            ModelPathDataObject DataObjectForApi = new ModelPathDataObject(model_path);
            var jsonObject = JsonConvert.SerializeObject(DataObjectForApi);
            newRequest.AddParameter("application/json; charset=utf-8", jsonObject, ParameterType.RequestBody);

            IRestResponse newResponse = client.Execute(newRequest);
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
