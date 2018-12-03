using Microsoft.Office.Tools.Ribbon;
using System.Collections.Generic;
using RestSharp;

namespace WordAddIn1
{
    public partial class Ribbon1
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
            this.TestModelDropDown.Enabled = false;
            this.WrapFromTestBtn.Enabled = false;
            this.AzureStorageButton.Checked = true;
            this.LocalStorageButton.Checked = false;
            this.TestModelDropDown.Enabled = false;
            this.WrapFromTestBtn.Enabled = false;
            this.SetDirButton.Enabled = false;
            this.ModelDirBox.Enabled = false;

            var client = new RestClient("http://127.0.0.1:6000");
            Globals.ThisAddIn.GetProjsListAzure(client, this.ProjectDropDown);

            if (this.ProjectDropDown.SelectedItem != null)
            {
                Globals.ThisAddIn.TESTGetModelsList(client, this.ProjectDropDown.SelectedItem.Label, this.TestModelDropDown);

                if (this.TestModelDropDown.Items.Count != 0)
                {
                    this.WrapFromTestBtn.Enabled = true;
                    this.TestModelDropDown.Enabled = true;
                }
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
            string NewProjName = Prompt.ShowDialog("New Project name:", "CREATE!");

            if (this.ProjectDropDown.Items.Count != 0)
            {
                List<string> ProjsList = new List<string>();
                foreach (RibbonDropDownItem item in this.ProjectDropDown.Items)
                {
                    string ExistingProjectName = item.ToString();
                    ProjsList.Add(ExistingProjectName);
                }

                while (ProjsList.Contains(NewProjName))
                {
                    NewProjName = Prompt.NewProjectShowDialog(NewProjName);
                }
            }

            RibbonDropDownItem newProj = Globals.Factory.GetRibbonFactory().CreateRibbonDropDownItem();
            newProj.Label = NewProjName;
            this.ProjectDropDown.Items.Add(newProj);
            this.ProjectDropDown.SelectedItem = newProj;

            this.TestModelDropDown.Items.Clear();
            this.TestModelDropDown.SelectedItem = null;
            this.TestModelDropDown.Enabled = false;
            this.WrapFromTestBtn.Enabled = false;
        }

        private void ProjectDropDown_Select(object sender, RibbonControlEventArgs e)
        {
            var client = new RestClient("http://127.0.0.1:6000");
            Globals.ThisAddIn.AddItemToProjectDropDown(this.ExportTXTbtn, this.WrapFromTestBtn, client, this.ModelDirDialog, this.ProjectDropDown, this.TestModelDropDown, this.AzureStorageButton, this.LocalStorageButton);
        }

        private void ModelDropDown_Select(object sender, RibbonControlEventArgs e)
        {
            this.WrapFromTestBtn.Enabled = false;
            this.ProjectDropDown.Enabled = false;
            this.ExportTXTbtn.Enabled = false;

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

            this.WrapFromTestBtn.Enabled = true;
            this.ProjectDropDown.Enabled = true;
            this.ExportTXTbtn.Enabled = true;
        }
        
        private void WrapFromTestBtn_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.TestDoc();
        }

        private void ExportTXTbtn_Click(object sender, RibbonControlEventArgs e)
        {
            string ModelName = Prompt.ShowDialog("Model name:", "TRAIN!");
            bool ForceUpdate = false;

            if (this.TestModelDropDown.Items.Count != 0)
            {
                List<string> ModelsList = new List<string>();
                foreach (RibbonDropDownItem item in this.TestModelDropDown.Items)
                {
                    string ExistingModelName = item.ToString();
                    ModelsList.Add(ExistingModelName);
                }

                while (ModelsList.Contains(ModelName))
                {
                    ModelName = Prompt.NewShowDialog(ModelName);
                }

                if (ModelName.Substring(ModelName.Length - 12) == "-ToOverwrite")
                {
                    ModelName = ModelName.Substring(0, ModelName.Length - 12);
                    ForceUpdate = true;
                }
            }

            string ProjectName = this.ProjectDropDown.SelectedItem.Label;
            var client = new RestClient("http://127.0.0.1:6000");

            if (this.AzureStorageButton.Checked)
            {
                Globals.ThisAddIn.ExportTrainData(client, ProjectName, ModelName);
                Globals.ThisAddIn.UpdateInterpreter(client, ProjectName, ModelName, ForceUpdate);
            }
            else
            {
                string ModelPath = this.ModelDirDialog.SelectedPath;
                Globals.ThisAddIn.ExportTrainData(client, ProjectName, ModelName, ModelPath);
                Globals.ThisAddIn.UpdateInterpreter(client, ProjectName, ModelName, ForceUpdate, ModelPath + "\\MODELS\\" + ProjectName + "\\" + ModelName);
            }

            if (ModelName != "")
            {
                RibbonDropDownItem newModel = Globals.Factory.GetRibbonFactory().CreateRibbonDropDownItem();
                newModel.Label = ModelName;
                this.TestModelDropDown.Items.Add(newModel);
                this.TestModelDropDown.SelectedItem = newModel;

                TestModelDropDown.Enabled = true;
                WrapFromTestBtn.Enabled = true;
            }
        }

        private void AzureStorageButton_Click(object sender, RibbonControlEventArgs e)
        {
            this.AzureStorageButton.Checked = true;
            this.LocalStorageButton.Checked = false;
            this.TestModelDropDown.Enabled = false;
            this.WrapFromTestBtn.Enabled = false;
            this.SetDirButton.Enabled = false;
            this.ModelDirBox.Enabled = false;

            this.ProjectDropDown.Items.Clear();
            this.ProjectDropDown.SelectedItem = null;
            this.TestModelDropDown.Items.Clear();
            this.TestModelDropDown.SelectedItem = null;

            var client = new RestClient("http://127.0.0.1:6000");
            Globals.ThisAddIn.GetProjsListAzure(client, this.ProjectDropDown);

            if (this.ProjectDropDown.SelectedItem != null)
            {
                Globals.ThisAddIn.TESTGetModelsList(client, this.ProjectDropDown.SelectedItem.Label, this.TestModelDropDown);

                if (this.TestModelDropDown.Items.Count != 0)
                {
                    this.WrapFromTestBtn.Enabled = true;
                    this.TestModelDropDown.Enabled = true;
                }
            }

            if (this.TestModelDropDown.SelectedItem != null)
            {
                this.TestModelDropDown.Enabled = true;
                this.WrapFromTestBtn.Enabled = true;
            }
        }

        private void LocalStorageButton_Click(object sender, RibbonControlEventArgs e)
        {
            this.AzureStorageButton.Checked = false;
            this.LocalStorageButton.Checked = true;
            this.TestModelDropDown.Enabled = false;
            this.WrapFromTestBtn.Enabled = false;
            this.SetDirButton.Enabled = true;
            this.ModelDirBox.Enabled = true;

            this.ProjectDropDown.Items.Clear();
            this.ProjectDropDown.SelectedItem = null;
            this.TestModelDropDown.Items.Clear();
            this.TestModelDropDown.SelectedItem = null;

            var client = new RestClient("http://127.0.0.1:6000");

            if (this.ModelDirBox.Text == "")
            {
                Globals.ThisAddIn.ChooseModelDir(client, this.ModelDirDialog, this.ModelDirBox, this.ProjectDropDown, this.TestModelDropDown);
            }
            else
            {
                Globals.ThisAddIn.ChangeToLocalStorage(client, this.ModelDirDialog.SelectedPath, this.ProjectDropDown, this.TestModelDropDown);
            }

            if (this.TestModelDropDown.SelectedItem != null)
            {
                this.TestModelDropDown.Enabled = true;
                this.WrapFromTestBtn.Enabled = true;
            }
        }

        private void SetDirButton_Click(object sender, RibbonControlEventArgs e)
        {
            var client = new RestClient("http://127.0.0.1:6000");

            this.WrapFromTestBtn.Enabled = false;
            this.TestModelDropDown.Enabled = false;
            Globals.ThisAddIn.ChooseModelDir(client, this.ModelDirDialog, this.ModelDirBox, this.ProjectDropDown, this.TestModelDropDown);

            if (this.TestModelDropDown.SelectedItem != null)
            {
                this.TestModelDropDown.Enabled = true;
                this.WrapFromTestBtn.Enabled = true;
            }
        }
    }
}