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

            var client = new RestClient("http://127.0.0.1:6000");
            Globals.ThisAddIn.TESTGetProjsList(client, this.TestProjectDropDown, this.TestModelDropDown);
            Globals.ThisAddIn.GetProjsList(client, this.ProjectDropDown);

            this.TestModelDropDown.Enabled = true;
            this.TestProjectDropDown.Enabled = true;
            this.WrapFromTestBtn.Enabled = true;
            this.ExportTXTbtn.Enabled = true;
        }

        private void ProjectDropDown_Select(object sender, RibbonControlEventArgs e)
        {

        }

        private void ModelBox_TextChanged(object sender, RibbonControlEventArgs e)
        {
            string ModelToCreateName = this.ModelBox.Text;
            string ProjectName = this.TestProjectDropDown.SelectedItem.Label;
            var client = new RestClient("http://127.0.0.1:6000");
            Globals.ThisAddIn.GetModelsList(client, ProjectName, ModelToCreateName);
        }

        private void TestProjectDropDown_Select(object sender, RibbonControlEventArgs e)
        {
            this.TestModelDropDown.Enabled = false;
            this.WrapFromTestBtn.Enabled = false;
            this.ExportTXTbtn.Enabled = false;

            var client = new RestClient("http://127.0.0.1:6000");
            string ProjectName = this.TestProjectDropDown.SelectedItem.Label;
            Globals.ThisAddIn.TESTGetModelsList(client, ProjectName, this.TestModelDropDown);

            this.TestModelDropDown.Enabled = true;
            this.WrapFromTestBtn.Enabled = true;
            this.ExportTXTbtn.Enabled = true;
        }

        private void ModelDropDown_Select(object sender, RibbonControlEventArgs e)
        {
            this.WrapFromTestBtn.Enabled = false;
            this.TestProjectDropDown.Enabled = false;
            this.ExportTXTbtn.Enabled = false;

            var client = new RestClient("http://127.0.0.1:6000");
            string ProjectName = this.TestProjectDropDown.SelectedItem.Label;
            string ModelName = this.TestModelDropDown.SelectedItem.Label;
            Globals.ThisAddIn.UpdateInterpreter(client, ProjectName, ModelName);

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
            Globals.ThisAddIn.ExportTrainData();
        }

        private void WrapFromTestBtn_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.TestDoc();
        }
    }
}