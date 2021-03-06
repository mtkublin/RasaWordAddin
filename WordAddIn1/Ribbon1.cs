﻿using Microsoft.Office.Tools.Ribbon;
using RestSharp;
using System.Collections.Generic;
using System.IO;

namespace WordAddIn1
{
    public partial class Ribbon1
    {
        public StreamWriter myStreamWriter;

        public void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
            this.AzureStorageButton.Checked = false;
            this.LocalStorageButton.Checked = true;
            //this.SetDirButton.Enabled = true;
            this.SetDirButton.Enabled = false;
            this.ModelDirBox.Enabled = false;
            this.ProjectDropDown.Enabled = false;
            this.ProjectAddButton.Enabled = false;
            this.TestModelDropDown.Enabled = false;
            //this.TestButton.Enabled = false;
            this.TestButton.Enabled = true;
            this.TrainingButton.Enabled = false;
            this.reverseTestBTN.Enabled = false;
            this.CurBMtextLabel.Label = "";
            this.CurBMentLabel.Label = "";
            this.IntOrEntLabel.Label = "";
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

        bool Overwrite = false;

        private void TrainingButton_Click(object sender, RibbonControlEventArgs e)
        {
            if (Globals.ThisAddIn.Application.ActiveDocument.Bookmarks.Count == 0)
            {
                this.TextMessageOkDialog("No intent or entities to train");
                return;
            }

            string ModelName = this.NameInputDialog("Model name:", "TRAIN!");

            List<string> ModelsList = new List<string>();
            if (this.TestModelDropDown.Items.Count != 0)
            {
                foreach (RibbonDropDownItem item in this.TestModelDropDown.Items)
                {
                    string ExistingModelName = item.ToString();
                    ModelsList.Add(ExistingModelName);
                }

                while (ModelsList.Contains(ModelName))
                {
                    ModelName = this.ModelNameTakenDialog(ModelName);
                }

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

                if (Overwrite == false)
                {
                    RibbonDropDownItem newModel = Globals.Factory.GetRibbonFactory().CreateRibbonDropDownItem();
                    newModel.Label = ModelName;
                    this.TestModelDropDown.Items.Add(newModel);
                    this.TestModelDropDown.SelectedItem = newModel;
                }
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

        private void HighlightInVisibleBTN_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.HighlightBookmarksInVisibleRange();
        }

        private void HighlightInNextVisibleBTN_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.HighlightBookmarksInNextRange();
        }

        private void UnhighlightBTN_Click(object sender, RibbonControlEventArgs e)
        {
            Microsoft.Office.Interop.Word.Document ActiveDoc = Globals.ThisAddIn.Application.ActiveDocument;

            foreach (Microsoft.Office.Interop.Word.Range range in ActiveDoc.StoryRanges)
            {
                Globals.ThisAddIn.UnhighlightControl(range);
            }
        }

        private void reverseTestBTN_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.ReverseTest();

            //Microsoft.Office.Interop.Word.Document activeDocument = Globals.ThisAddIn.Application.ActiveDocument;
            //var extendedDocument = Globals.Factory.GetVstoObject(activeDocument);

            //foreach (Microsoft.Office.Interop.Word.Range range in activeDocument.StoryRanges)
            //{
            //    Globals.ThisAddIn.UnhighlightControl(range);
            //}

            //foreach (Microsoft.Office.Interop.Word.Bookmark existingBM in activeDocument.Bookmarks)
            //{
            //    Microsoft.Office.Tools.Word.Bookmark VSTOexistingBM = extendedDocument.Controls[existingBM.Name] as Microsoft.Office.Tools.Word.Bookmark;
            //    VSTOexistingBM.Delete();
            //}

            //Dictionary<string, Microsoft.Office.Interop.Word.Range> bmsDict = Globals.ThisAddIn.bmRangesPriorToTestDict;
            //foreach (string bmName in bmsDict.Keys)
            //{
            //    extendedDocument.Controls.AddBookmark(bmsDict[bmName], bmName);
            //}

            //Globals.ThisAddIn.HighlightBookmarksInVisibleRange();
            //this.reverseTestBTN.Enabled = false;
            //Globals.ThisAddIn.currentBookmark = null;
        }
    }
}