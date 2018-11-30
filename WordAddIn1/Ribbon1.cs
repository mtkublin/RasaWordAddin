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
            this.ExportTXTbtn.Enabled = false;
            this.AzureStorageButton.Checked = true;
            this.LocalStorageButton.Checked = false;
            this.TestModelDropDown.Enabled = false;
            this.ExportTXTbtn.Enabled = false;
            this.WrapFromTestBtn.Enabled = false;
            this.box2.Visible = false;

            var client = new RestClient("http://127.0.0.1:6000");
            Globals.ThisAddIn.GetProjsListAzure(client, this.ProjectComboBox);
        }

        private void ContentControlButton_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.WrapContent();
        }

        private void UnwrapRangeButton_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.UnwrapContent();
        }

        private void ProjectComboBox_TextChanged(object sender, RibbonControlEventArgs e)
        {
            var client = new RestClient("http://127.0.0.1:6000");
            Globals.ThisAddIn.AddItemToProjectComboBox(this.ExportTXTbtn, this.WrapFromTestBtn, client, this.ModelDirDialog, this.ProjectComboBox, this.TestModelDropDown, this.AzureStorageButton, this.LocalStorageButton);
        }

        private void ModelDropDown_Select(object sender, RibbonControlEventArgs e)
        {
            this.WrapFromTestBtn.Enabled = false;
            this.ProjectComboBox.Enabled = false;
            this.ExportTXTbtn.Enabled = false;

            string ProjectName = this.ProjectComboBox.Text;
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
            this.ProjectComboBox.Enabled = true;
            this.ExportTXTbtn.Enabled = true;
        }
        
        private void WrapFromTestBtn_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.TestDoc();
        }

        private void ExportTXTbtn_Click(object sender, RibbonControlEventArgs e)
        {
            string ModelName = Prompt.ShowDialog("Model name:", "");
            bool ForceUpdate = false;

            List<string> ModelsList = new List<string>();
            foreach (RibbonDropDownItem item in TestModelDropDown.Items)
            {
                string ExistingModelName = item.ToString();
                ModelsList.Add(ExistingModelName);
            }

            if (ModelsList.Contains(ModelName))
            {
                while (ModelsList.Contains(ModelName))
                {
                    ModelName = Prompt.NewShowDialog(ModelName);
                }

            }

            if (ModelName.Substring(ModelName.Length - 12) == "-ToOverwrite")
            {
                ModelName = ModelName.Substring(0, ModelName.Length - 12);
                ForceUpdate = true;
            }

            string ProjectName = this.ProjectComboBox.Text;
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

            RibbonDropDownItem newModel = Globals.Factory.GetRibbonFactory().CreateRibbonDropDownItem();
            newModel.Label = ModelName;
            this.TestModelDropDown.Items.Add(newModel);
            this.TestModelDropDown.SelectedItem = newModel;
        }

        private void AzureStorageButton_Click(object sender, RibbonControlEventArgs e)
        {
            this.AzureStorageButton.Checked = true;
            this.LocalStorageButton.Checked = false;
            this.TestModelDropDown.Enabled = false;
            this.ExportTXTbtn.Enabled = false;
            this.WrapFromTestBtn.Enabled = false;
            this.box2.Visible = false;

            this.ProjectComboBox.Items.Clear();
            this.ProjectComboBox.Text = "";
            this.TestModelDropDown.Items.Clear();

            var client = new RestClient("http://127.0.0.1:6000");
            Globals.ThisAddIn.GetProjsListAzure(client, this.ProjectComboBox);
        }

        private void LocalStorageButton_Click(object sender, RibbonControlEventArgs e)
        {
            this.AzureStorageButton.Checked = false;
            this.LocalStorageButton.Checked = true;
            this.TestModelDropDown.Enabled = false;
            this.ExportTXTbtn.Enabled = false;
            this.WrapFromTestBtn.Enabled = false;
            this.box2.Visible = true;
            var client = new RestClient("http://127.0.0.1:6000");

            if (this.ModelDirLabel.Label == "")
            {
                Globals.ThisAddIn.ChooseModelDir(this.ModelDirDialog, this.ModelDirLabel, this.ProjectComboBox, this.TestModelDropDown);
            }
            else
            {
                Globals.ThisAddIn.ChangeToLocalStorage(this.ModelDirDialog.SelectedPath, this.ProjectComboBox, this.TestModelDropDown);
            }
        }

        private void SetDirButton_Click(object sender, RibbonControlEventArgs e)
        {
            this.WrapFromTestBtn.Enabled = false;
            this.TestModelDropDown.Enabled = false;
            Globals.ThisAddIn.ChooseModelDir(this.ModelDirDialog, this.ModelDirLabel, this.ProjectComboBox, this.TestModelDropDown);
        }
    }
}