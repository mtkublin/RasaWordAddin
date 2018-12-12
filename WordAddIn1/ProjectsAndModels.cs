using System.Collections.Generic;
using System.IO;
using System.Timers;
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

            GetItemsLocal(ModelDir + "\\MODELS", ProjectDropDown);

            if (ProjectDropDown.SelectedItem != null)
            {
                GetModelsListLocal(ModelDir + "\\MODELS\\" + ProjectDropDown.SelectedItem.Label, TestModelDropDown);

                if (TestModelDropDown.Items.Count != 0 )
                {
                    UpdateInterpreter(client, ProjectDropDown.SelectedItem.Label, TestModelDropDown.SelectedItem.Label, false, ModelDir + "\\MODELS\\" + ProjectDropDown.SelectedItem.Label + "\\" + TestModelDropDown.SelectedItem.Label);
                }
            }
        }

        public void ChangeCurrentProject(RibbonButton ExportTXTbtn, RibbonButton WrapFromTestBtn, RestClient client, System.Windows.Forms.FolderBrowserDialog ModelDirDialog, RibbonDropDown ProjectDropDown, RibbonDropDown TestModelDropDown, RibbonToggleButton AzureStorageButton, RibbonToggleButton LocalStorageButton)
        {
            string ProjectName = ProjectDropDown.SelectedItem.Label;
            if (AzureStorageButton.Checked == true)
            {
                Globals.ThisAddIn.GetModelsListAzure(client, ProjectName, TestModelDropDown);
            }
            else if (LocalStorageButton.Checked == true)
            {
                Globals.ThisAddIn.GetModelsListLocal(ModelDirDialog.SelectedPath + "\\MODELS\\" + ProjectName, TestModelDropDown);
                if (TestModelDropDown.Items.Count != 0)
                {
                    string ModelName = TestModelDropDown.SelectedItem.Label;
                    string ModelPath = ModelDirDialog.SelectedPath + "\\MODELS\\" + ProjectName + "\\" + ModelName;
                    Globals.ThisAddIn.UpdateInterpreter(client, ProjectName, ModelName, false, ModelPath);
                }
            }
        }

        public void AddNewProject()
        {
            string NewProjName = Globals.Ribbons.Ribbon1.NameInputDialog("New Project name:", "CREATE!");

            if (Globals.Ribbons.Ribbon1.ProjectDropDown.Items.Count != 0)
            {
                List<string> ProjsList = new List<string>();
                foreach (RibbonDropDownItem item in Globals.Ribbons.Ribbon1.ProjectDropDown.Items)
                {
                    string ExistingProjectName = item.ToString();
                    ProjsList.Add(ExistingProjectName);
                }

                while (ProjsList.Contains(NewProjName))
                {
                    NewProjName = Globals.Ribbons.Ribbon1.ProjectNameTakenDialog(NewProjName);
                }
            }

            RibbonDropDownItem newProj = Globals.Factory.GetRibbonFactory().CreateRibbonDropDownItem();
            newProj.Label = NewProjName;
            Globals.Ribbons.Ribbon1.ProjectDropDown.Items.Add(newProj);
            Globals.Ribbons.Ribbon1.ProjectDropDown.SelectedItem = newProj;

            Globals.Ribbons.Ribbon1.TestModelDropDown.Items.Clear();
            Globals.Ribbons.Ribbon1.TestModelDropDown.SelectedItem = null;
            Globals.Ribbons.Ribbon1.TestModelDropDown.Enabled = false;
            Globals.Ribbons.Ribbon1.TestButton.Enabled = false;
        }

        public void UpdateInterpreter(RestClient client, string ProjectName, string ModelName, bool ForceUpdate = false, string model_path = "")
        {
            Globals.Ribbons.Ribbon1.ProjectDropDown.Enabled = false;
            Globals.Ribbons.Ribbon1.ProjectAddButton.Enabled = false;
            Globals.Ribbons.Ribbon1.TestModelDropDown.Enabled = false;
            Globals.Ribbons.Ribbon1.TestButton.Enabled = false;
            Globals.Ribbons.Ribbon1.TrainingButton.Enabled = false;
            Globals.Ribbons.Ribbon1.LocalStorageButton.Enabled = false;
            Globals.Ribbons.Ribbon1.AzureStorageButton.Enabled = false;
            Globals.Ribbons.Ribbon1.SetDirButton.Enabled = false;

            var Request = new RestRequest("api/interpreter/{project}/{model}/{force}", Method.POST);
            Request.AddParameter("project", ProjectName, ParameterType.UrlSegment);
            Request.AddUrlSegment("project", ProjectName);
            Request.AddParameter("model", ModelName, ParameterType.UrlSegment);
            Request.AddUrlSegment("model", ModelName);
            Request.AddParameter("force", ForceUpdate.ToString(), ParameterType.UrlSegment);
            Request.AddUrlSegment("force", ForceUpdate.ToString());

            ModelPathDataObject DataObjectForApi = new ModelPathDataObject(model_path);
            var jsonObject = JsonConvert.SerializeObject(DataObjectForApi);
            Request.AddParameter("application/json; charset=utf-8", jsonObject, ParameterType.RequestBody);

            IRestResponse Response = client.Execute(Request);

            UpdateStatusCheckTimer = new Timer(3000);
            UpdateStatusCheckTimer.AutoReset = true;
            UpdateStatusCheckTimer.Elapsed += (sender, e) => CheckUpdateStatus(sender, e, ProjectName, ModelName, client);
            UpdateStatusCheckTimer.Enabled = true;
        }

        private static Timer UpdateStatusCheckTimer;

        private static void CheckUpdateStatus(object source, ElapsedEventArgs e, string ProjectName, string ModelName, RestClient client)
        {
            var newRequest = new RestRequest("api/interpreter/isloaded/{project}/{model}", Method.GET);
            newRequest.AddParameter("project", ProjectName, ParameterType.UrlSegment);
            newRequest.AddUrlSegment("project", ProjectName);
            newRequest.AddParameter("model", ModelName, ParameterType.UrlSegment);
            newRequest.AddUrlSegment("model", ModelName);

            IRestResponse newResponse = client.Execute(newRequest);
            string IsLoaded = JsonConvert.DeserializeObject<string>(newResponse.Content.ToString());

            if (IsLoaded == "True")
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

                UpdateStatusCheckTimer.Stop();
                UpdateStatusCheckTimer.Dispose();
            }
        }

        public void GetItemsLocal(string ModelDir, RibbonDropDown ProjectDropDown)
        {
            string[] ProjFoldList = Directory.GetDirectories(ModelDir);
            foreach (string Pfold in ProjFoldList)
            {
                string ItemLabel = Pfold.Substring(ModelDir.Length + 1, Pfold.Length - ModelDir.Length - 1);

                RibbonDropDownItem folder = Globals.Factory.GetRibbonFactory().CreateRibbonDropDownItem();
                folder.Label = ItemLabel;
                ProjectDropDown.Items.Add(folder);
            }
        }

        public void GetModelsListLocal(string NewModelDir, RibbonDropDown TestModelDropDown)
        {
            if (NewModelDir != null)
            {
                TestModelDropDown.Items.Clear();

                GetItemsLocal(NewModelDir, TestModelDropDown);
            }
        }

        public void GetProjsListAzure(RestClient client, RibbonDropDown ProjectDropDown)
        {
            var Request = new RestRequest("api/projects", Method.GET);
            IRestResponse Response = client.Execute(Request);
            string JSONresultDoc = Response.Content.ToString();
            List<string> ProjList = JsonConvert.DeserializeObject<List<string>>(JSONresultDoc);

            ProjectDropDown.Items.Clear();

            foreach (string itemName in ProjList)
            {
                RibbonDropDownItem item = Globals.Factory.GetRibbonFactory().CreateRibbonDropDownItem();
                item.Label = itemName;
                ProjectDropDown.Items.Add(item);
            }
        }

        public void GetModelsListAzure(RestClient client, string ProjectName, RibbonDropDown TestModelDropDown)
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
                Globals.Ribbons.Ribbon1.TestButton.Enabled = false;
            }
        }
    }
}
