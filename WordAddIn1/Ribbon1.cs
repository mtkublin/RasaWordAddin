using Microsoft.Office.Tools.Ribbon;
using RestSharp;

namespace WordAddIn1
{
    public partial class Ribbon1
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
            this.TestModelDropDown.Enabled = false;
            this.TestProjectDropDown.Enabled = false;
            this.WrapFromTestBtn.Enabled = false;
            this.ExportTXTbtn.Enabled = false;
            this.AzureStorageButton.Checked = true;
            this.LocalStorageButton.Checked = false;
            this.SetDirButton.Enabled = false;

            var client = new RestClient("http://127.0.0.1:6000");
            Globals.ThisAddIn.TESTGetProjsList(client, this.TestProjectDropDown, this.TestModelDropDown);
            Globals.ThisAddIn.GetProjsList(client, this.ProjectComboBox);

            this.TestModelDropDown.Enabled = true;
            this.TestProjectDropDown.Enabled = true;
            this.WrapFromTestBtn.Enabled = true;
            this.ExportTXTbtn.Enabled = true;
        }

        private void AzureStorageButton_Click(object sender, RibbonControlEventArgs e)
        {
            this.AzureStorageButton.Checked = true;
            this.LocalStorageButton.Checked = false;
            this.SetDirButton.Enabled = false;
            this.ModelDirLabel.Enabled = false;

            this.ProjectComboBox.Items.Clear();
            this.ProjectComboBox.Text = "";
            this.TestProjectDropDown.Items.Clear();
            this.ModelComboBox.Items.Clear();
            this.ModelComboBox.Text = "";
            this.TestModelDropDown.Items.Clear();

            var client = new RestClient("http://127.0.0.1:6000");
            Globals.ThisAddIn.TESTGetProjsList(client, this.TestProjectDropDown, this.TestModelDropDown);
            Globals.ThisAddIn.GetProjsList(client, this.ProjectComboBox);
        }

        private void LocalStorageButton_Click(object sender, RibbonControlEventArgs e)
        {
            this.AzureStorageButton.Checked = false;
            this.LocalStorageButton.Checked = true;
            this.SetDirButton.Enabled = true;
            this.ModelDirLabel.Enabled = true;

            if (this.ModelDirLabel.Label == "")
            {
                Globals.ThisAddIn.ChooseModelDir(this.ModelDirDialog, this.ModelDirLabel, this.ProjectComboBox, this.TestProjectDropDown, this.ModelComboBox, this.TestModelDropDown);

                var client = new RestClient("http://127.0.0.1:6000");
                string ProjectName = this.TestProjectDropDown.SelectedItem.Label;
                string ModelName = this.TestModelDropDown.SelectedItem.Label;
                string ModelPath = this.ModelDirDialog.SelectedPath + "\\" + ProjectName + "\\" + ModelName;
                Globals.ThisAddIn.UpdateInterpreter(client, ProjectName, ModelName, ModelPath);
            }
            else
            {
                Globals.ThisAddIn.ChangeToLocalStorage(this.ModelDirDialog, this.ModelDirLabel, this.ProjectComboBox, this.TestProjectDropDown, this.ModelComboBox, this.TestModelDropDown);
            }
        }

        private void SetDirButton_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.ChooseModelDir(this.ModelDirDialog, this.ModelDirLabel, this.ProjectComboBox, this.TestProjectDropDown, this.ModelComboBox, this.TestModelDropDown);
        }

        private void ProjectComboBox_TextChanged(object sender, RibbonControlEventArgs e)
        {
            this.ModelComboBox.Text = "";
            Globals.ThisAddIn.AddItemToProjectComboBox(this.ProjectComboBox, this.ModelComboBox);
            string ProjectName = ProjectComboBox.Text;

            if (this.AzureStorageButton.Checked == true)
            {
                var client = new RestClient("http://127.0.0.1:6000");
                Globals.ThisAddIn.GetModelsList(client, ProjectName, this.ModelComboBox);
            }
            else if (this.LocalStorageButton.Checked == true)
            {
                Globals.ThisAddIn.GetModelItemsFromDirCB(this.ModelDirDialog.SelectedPath + "\\" + ProjectName, ModelComboBox);
            }
        }

        private void ModelComboBox_TextChanged(object sender, RibbonControlEventArgs e)
        {
            string ModelToCreateName = this.ModelComboBox.Text;
            Globals.ThisAddIn.AddItemToModelComboBox(this.ModelComboBox);
        }

        private void TestProjectDropDown_Select(object sender, RibbonControlEventArgs e)
        {
            this.TestModelDropDown.Enabled = false;
            this.WrapFromTestBtn.Enabled = false;
            this.ExportTXTbtn.Enabled = false;

            string ProjectName = this.TestProjectDropDown.SelectedItem.Label;

            if (this.AzureStorageButton.Checked == true)
            {
                var client = new RestClient("http://127.0.0.1:6000");
                Globals.ThisAddIn.TESTGetModelsList(client, ProjectName, this.TestModelDropDown);
            }
            else if (this.LocalStorageButton.Checked == true)
            {
                Globals.ThisAddIn.GetModelItemsFromDirDD(this.ModelDirDialog.SelectedPath + "\\" + ProjectName, TestModelDropDown);
                var client = new RestClient("http://127.0.0.1:6000");
                string ModelName = this.TestModelDropDown.SelectedItem.Label;
                string ModelPath = this.ModelDirDialog.SelectedPath + "\\" + ProjectName + "\\" + ModelName;
                Globals.ThisAddIn.UpdateInterpreter(client, ProjectName, ModelName, ModelPath);
            }

            this.TestModelDropDown.Enabled = true;
            this.WrapFromTestBtn.Enabled = true;
            this.ExportTXTbtn.Enabled = true;
        }

        private void ModelDropDown_Select(object sender, RibbonControlEventArgs e)
        {
            this.WrapFromTestBtn.Enabled = false;
            this.TestProjectDropDown.Enabled = false;
            this.ExportTXTbtn.Enabled = false;

            string ProjectName = this.TestProjectDropDown.SelectedItem.Label;
            string ModelName = this.TestModelDropDown.SelectedItem.Label;

            var client = new RestClient("http://127.0.0.1:6000");

            if (this.AzureStorageButton.Checked == true)
            {
                Globals.ThisAddIn.UpdateInterpreter(client, ProjectName, ModelName);
            }
            else if (this.LocalStorageButton.Checked == true)
            {
                string ModelPath = this.ModelDirDialog.SelectedPath + "\\" + ProjectName + "\\" + ModelName;
                Globals.ThisAddIn.UpdateInterpreter(client, ProjectName, ModelName, ModelPath);
            }

            this.WrapFromTestBtn.Enabled = true;
            this.TestProjectDropDown.Enabled = true;
            this.ExportTXTbtn.Enabled = true;
        }

        private void ContentControlButton_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.WrapContent();
        }

        private void UnwrapRangeButton_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.UnwrapContent();
        }

        private void ExportTXTbtn_Click(object sender, RibbonControlEventArgs e)
        {
            if(this.AzureStorageButton.Checked)
            {
                Globals.ThisAddIn.ExportTrainData();
            }
            else
            {
                string ModelPath = this.ModelDirDialog.SelectedPath;
                Globals.ThisAddIn.ExportTrainData(ModelPath);
            }
        }

        private void WrapFromTestBtn_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.TestDoc();
        }
    }
}