using Microsoft.Office.Tools.Ribbon;
using System.Collections.Generic;
using RestSharp;

namespace WordAddIn1
{
    public partial class Ribbon1
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
            this.AzureStorageButton.Checked = true;
            this.LocalStorageButton.Checked = false;
            this.SetDirButton.Enabled = false;
            this.ModelDirBox.Enabled = false;

            var client = new RestClient("http://127.0.0.1:6000");
            Globals.ThisAddIn.GetProjsListAzure(client, this.ProjectDropDown);

            if (this.ProjectDropDown.SelectedItem != null)
            {
                Globals.ThisAddIn.GetModelsListAzure(client, this.ProjectDropDown.SelectedItem.Label, this.TestModelDropDown);
            }
        }

        private void ContentControlButton_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.WrapContent();
        }

        private void UnwrapRangeButton_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.UnwrapContent();
        }

        private void ProjectAddButton_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.AddNewProject();
        }

        private void ProjectDropDown_Select(object sender, RibbonControlEventArgs e)
        {
            var client = new RestClient("http://127.0.0.1:6000");
            Globals.ThisAddIn.ChangeCurrentProject(this.TrainingButton, this.TestButton, client, this.ModelDirDialog, this.ProjectDropDown, this.TestModelDropDown, this.AzureStorageButton, this.LocalStorageButton);
        }

        private void ModelDropDown_Select(object sender, RibbonControlEventArgs e)
        {
            string ProjectName = this.ProjectDropDown.SelectedItem.Label;
            string ModelName = this.TestModelDropDown.SelectedItem.Label;

            var client = new RestClient("http://127.0.0.1:6000");

            if (this.AzureStorageButton.Checked == true)
            {
                Globals.ThisAddIn.UpdateInterpreter(client, ProjectName, ModelName);
            }
            else if (this.LocalStorageButton.Checked == true)
            {
                string ModelPath = this.ModelDirDialog.SelectedPath + "\\MODELS\\" + ProjectName + "\\" + ModelName;
                Globals.ThisAddIn.UpdateInterpreter(client, ProjectName, ModelName, false, ModelPath);
            }
        }
        
        private void TestButton_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.TestDoc();
        }

        private void TrainingButton_Click(object sender, RibbonControlEventArgs e)
        {
            if (Globals.ThisAddIn.Application.ActiveDocument.ContentControls.Count == 0)
            {
                Prompt.TextMessageOkDialog("No intent or entities to train");
                return;
            }

            string ModelName = Prompt.NameInputDialog("Model name:", "TRAIN!");

            if (this.TestModelDropDown.Items.Count != 0)
            {
                List<string> ModelsList = new List<string>();
                foreach (RibbonDropDownItem item in this.TestModelDropDown.Items)
                {
                    string ExistingModelName = item.ToString();
                    ModelsList.Add(ExistingModelName);
                }

                bool Overwrite = false;
                while (ModelsList.Contains(ModelName))
                {
                    ModelName = Prompt.ModelNameTakenDialog(ModelName);
                    Overwrite = true;
                }

                //if (ModelName.Length >= 12 & ModelName.Substring(ModelName.Length - 12, 12) == "-ToOverwrite")
                if (Overwrite)
                {
                    ModelName = ModelName.Substring(0, ModelName.Length - 12);
                }
            }

            if (ModelName != "")
            {
                string ProjectName = this.ProjectDropDown.SelectedItem.Label;
                var client = new RestClient("http://127.0.0.1:6000");

                if (this.AzureStorageButton.Checked)
                {
                    Globals.ThisAddIn.InitiateTraining(client, ProjectName, ModelName);
                }
                else
                {
                    string ModelPath = this.ModelDirDialog.SelectedPath;
                    Globals.ThisAddIn.InitiateTraining(client, ProjectName, ModelName, ModelPath);
                }

                RibbonDropDownItem newModel = Globals.Factory.GetRibbonFactory().CreateRibbonDropDownItem();
                newModel.Label = ModelName;
                this.TestModelDropDown.Items.Add(newModel);
                this.TestModelDropDown.SelectedItem = newModel;
            }
        }

        private void AzureStorageButton_Click(object sender, RibbonControlEventArgs e)
        {
            this.AzureStorageButton.Checked = true;
            this.LocalStorageButton.Checked = false;
            this.SetDirButton.Enabled = false;
            this.ModelDirBox.Enabled = false;


            var client = new RestClient("http://127.0.0.1:6000");
            Globals.ThisAddIn.GetProjsListAzure(client, this.ProjectDropDown);

            if (this.ProjectDropDown.SelectedItem != null)
            {
                Globals.ThisAddIn.GetModelsListAzure(client, this.ProjectDropDown.SelectedItem.Label, this.TestModelDropDown);
            }
        }

        private void LocalStorageButton_Click(object sender, RibbonControlEventArgs e)
        {
            this.AzureStorageButton.Checked = false;
            this.LocalStorageButton.Checked = true;
            this.SetDirButton.Enabled = true;

            this.ProjectDropDown.Items.Clear();
            this.TestModelDropDown.Items.Clear();

            var client = new RestClient("http://127.0.0.1:6000");

            if (this.ModelDirBox.Text == "")
            {
                Globals.ThisAddIn.ChooseModelDir(client, this.ModelDirDialog, this.ModelDirBox, this.ProjectDropDown, this.TestModelDropDown);
            }
            else
            {
                Globals.ThisAddIn.ChangeToLocalStorage(client, this.ModelDirDialog.SelectedPath, this.ProjectDropDown, this.TestModelDropDown);
            }
        }

        private void SetDirButton_Click(object sender, RibbonControlEventArgs e)
        {
            var client = new RestClient("http://127.0.0.1:6000");
            Globals.ThisAddIn.ChooseModelDir(client, this.ModelDirDialog, this.ModelDirBox, this.ProjectDropDown, this.TestModelDropDown);
        }
    }
}