﻿using System.Windows.Forms;


namespace WordAddIn1
{
    partial class Ribbon1 : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public Ribbon1()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        /// <summary> 
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        //
        // Text Dialog
        //

        public static class Prompt
        {
            public static string ShowDialog(string text, string caption)
            {
                Form prompt = new Form()
                {
                    Width = 300,
                    Height = 140,
                    FormBorderStyle = FormBorderStyle.FixedDialog,
                    Text = caption,
                    StartPosition = FormStartPosition.CenterScreen
                };
                Label textLabel = new Label() { Left = 25, Top = 15, Text = text };
                TextBox textBox = new TextBox() { Left = 25, Top = 40, Width = 225 };
                Button confirmation = new Button() { Text = "TRAIN!", Left = 175, Width = 75, Top = 70, DialogResult = DialogResult.OK };
                confirmation.Click += (sender, e) => { prompt.Close(); };
                prompt.Controls.Add(textBox);
                prompt.Controls.Add(confirmation);
                prompt.Controls.Add(textLabel);
                prompt.AcceptButton = confirmation;

                return prompt.ShowDialog() == DialogResult.OK ? textBox.Text : "";
            }

            public static string NewShowDialog(string TakenModelName)
            {
                Form prompt = new Form()
                {
                    Width = 235,
                    Height = 140,
                    FormBorderStyle = FormBorderStyle.FixedDialog,
                    Text = "",
                    StartPosition = FormStartPosition.CenterScreen
                };
                Label textLabel = new Label() { Left = 25, Top = 15, Width = 200, Text = "This model name is already taken." };
                Label textLabel2 = new Label() { Left = 25, Top = 40, Width = 200, Text = "Do you want to overwrite it?" };
                Button NOTconfirmation = new Button() { Text = "NO", Left = 25, Width = 75, Top = 65, DialogResult = DialogResult.OK };
                Button confirmation = new Button() { Text = "YES", Left = 110, Width = 75, Top = 65, DialogResult = DialogResult.OK };
                confirmation.Click += (sender, e) => {TakenModelName += "-ToOverwrite"; prompt.Close(); };
                NOTconfirmation.Click += (sender, e) => { TakenModelName = Prompt.ShowDialog("Model name:", ""); prompt.Close(); };
                prompt.Controls.Add(confirmation);
                prompt.Controls.Add(NOTconfirmation);
                prompt.Controls.Add(textLabel);
                prompt.Controls.Add(textLabel2);
                prompt.AcceptButton = confirmation;

                return prompt.ShowDialog() == DialogResult.OK ? TakenModelName : "";
            }
        }

        #region Component Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Ribbon1));
            this.ModelDirDialog = new System.Windows.Forms.FolderBrowserDialog();
            this.tab1 = this.Factory.CreateRibbonTab();
            this.group2 = this.Factory.CreateRibbonGroup();
            this.ContentControlButton = this.Factory.CreateRibbonButton();
            this.UnwrapRangeButton = this.Factory.CreateRibbonButton();
            this.group3 = this.Factory.CreateRibbonGroup();
            this.box1 = this.Factory.CreateRibbonBox();
            this.button1 = this.Factory.CreateRibbonButton();
            this.WrapFromTestBtn = this.Factory.CreateRibbonButton();
            this.ExportTXTbtn = this.Factory.CreateRibbonButton();
            this.box4 = this.Factory.CreateRibbonBox();
            this.ProjectComboBox = this.Factory.CreateRibbonComboBox();
            this.TestModelDropDown = this.Factory.CreateRibbonDropDown();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.box3 = this.Factory.CreateRibbonBox();
            this.box5 = this.Factory.CreateRibbonBox();
            this.LocalStorageButton = this.Factory.CreateRibbonToggleButton();
            this.AzureStorageButton = this.Factory.CreateRibbonToggleButton();
            this.box6 = this.Factory.CreateRibbonBox();
            this.SetDirButton = this.Factory.CreateRibbonButton();
            this.ModelDirBox = this.Factory.CreateRibbonEditBox();
            this.ProjectDropDown = this.Factory.CreateRibbonDropDown();
            this.tab1.SuspendLayout();
            this.group2.SuspendLayout();
            this.group3.SuspendLayout();
            this.box1.SuspendLayout();
            this.box4.SuspendLayout();
            this.group1.SuspendLayout();
            this.box3.SuspendLayout();
            this.box5.SuspendLayout();
            this.box6.SuspendLayout();
            this.SuspendLayout();
            // 
            // ModelDirDialog
            // 
            this.ModelDirDialog.SelectedPath = "C:\\Users\\Mikołaj";
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.group2);
            this.tab1.Groups.Add(this.group3);
            this.tab1.Groups.Add(this.group1);
            this.tab1.Label = "RasaNLU addin";
            this.tab1.Name = "tab1";
            // 
            // group2
            // 
            this.group2.Items.Add(this.ContentControlButton);
            this.group2.Items.Add(this.UnwrapRangeButton);
            this.group2.Label = "Wrapper";
            this.group2.Name = "group2";
            // 
            // ContentControlButton
            // 
            this.ContentControlButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.ContentControlButton.Image = ((System.Drawing.Image)(resources.GetObject("ContentControlButton.Image")));
            this.ContentControlButton.KeyTip = "W";
            this.ContentControlButton.Label = "Wrap Content";
            this.ContentControlButton.Name = "ContentControlButton";
            this.ContentControlButton.ShowImage = true;
            this.ContentControlButton.Tag = "controlTag";
            this.ContentControlButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ContentControlButton_Click);
            // 
            // UnwrapRangeButton
            // 
            this.UnwrapRangeButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.UnwrapRangeButton.Image = ((System.Drawing.Image)(resources.GetObject("UnwrapRangeButton.Image")));
            this.UnwrapRangeButton.Label = "Unwrap Range";
            this.UnwrapRangeButton.Name = "UnwrapRangeButton";
            this.UnwrapRangeButton.ShowImage = true;
            this.UnwrapRangeButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.UnwrapRangeButton_Click);
            // 
            // group3
            // 
            this.group3.Items.Add(this.box1);
            this.group3.Items.Add(this.box4);
            this.group3.Label = "Test/Train";
            this.group3.Name = "group3";
            // 
            // box1
            // 
            this.box1.BoxStyle = Microsoft.Office.Tools.Ribbon.RibbonBoxStyle.Vertical;
            this.box1.Items.Add(this.button1);
            this.box1.Items.Add(this.WrapFromTestBtn);
            this.box1.Items.Add(this.ExportTXTbtn);
            this.box1.Name = "box1";
            // 
            // button1
            // 
            this.button1.Image = ((System.Drawing.Image)(resources.GetObject("button1.Image")));
            this.button1.Label = "Project (Add)";
            this.button1.Name = "button1";
            this.button1.ShowImage = true;
            // 
            // WrapFromTestBtn
            // 
            this.WrapFromTestBtn.Image = ((System.Drawing.Image)(resources.GetObject("WrapFromTestBtn.Image")));
            this.WrapFromTestBtn.Label = "Test with";
            this.WrapFromTestBtn.Name = "WrapFromTestBtn";
            this.WrapFromTestBtn.ShowImage = true;
            this.WrapFromTestBtn.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.WrapFromTestBtn_Click);
            // 
            // ExportTXTbtn
            // 
            this.ExportTXTbtn.Image = ((System.Drawing.Image)(resources.GetObject("ExportTXTbtn.Image")));
            this.ExportTXTbtn.Label = "Train";
            this.ExportTXTbtn.Name = "ExportTXTbtn";
            this.ExportTXTbtn.ShowImage = true;
            this.ExportTXTbtn.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ExportTXTbtn_Click);
            // 
            // box4
            // 
            this.box4.BoxStyle = Microsoft.Office.Tools.Ribbon.RibbonBoxStyle.Vertical;
            this.box4.Items.Add(this.ProjectComboBox);
            this.box4.Items.Add(this.TestModelDropDown);
            this.box4.Items.Add(this.ProjectDropDown);
            this.box4.Name = "box4";
            // 
            // ProjectComboBox
            // 
            this.ProjectComboBox.Label = " ";
            this.ProjectComboBox.MaxLength = 21;
            this.ProjectComboBox.Name = "ProjectComboBox";
            this.ProjectComboBox.SizeString = "model_20181017-154908aa";
            this.ProjectComboBox.Text = null;
            this.ProjectComboBox.TextChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ProjectComboBox_TextChanged);
            // 
            // TestModelDropDown
            // 
            this.TestModelDropDown.Label = " ";
            this.TestModelDropDown.Name = "TestModelDropDown";
            this.TestModelDropDown.ShowItemImage = false;
            this.TestModelDropDown.SizeString = "model_20181017-154908aa";
            this.TestModelDropDown.SelectionChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ModelDropDown_Select);
            // 
            // group1
            // 
            this.group1.Items.Add(this.box3);
            this.group1.Label = "Storage";
            this.group1.Name = "group1";
            // 
            // box3
            // 
            this.box3.BoxStyle = Microsoft.Office.Tools.Ribbon.RibbonBoxStyle.Vertical;
            this.box3.Items.Add(this.box5);
            this.box3.Items.Add(this.box6);
            this.box3.Items.Add(this.ModelDirBox);
            this.box3.Name = "box3";
            // 
            // box5
            // 
            this.box5.Items.Add(this.LocalStorageButton);
            this.box5.Items.Add(this.AzureStorageButton);
            this.box5.Name = "box5";
            // 
            // LocalStorageButton
            // 
            this.LocalStorageButton.Image = ((System.Drawing.Image)(resources.GetObject("LocalStorageButton.Image")));
            this.LocalStorageButton.Label = "Local";
            this.LocalStorageButton.Name = "LocalStorageButton";
            this.LocalStorageButton.ShowImage = true;
            this.LocalStorageButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.LocalStorageButton_Click);
            // 
            // AzureStorageButton
            // 
            this.AzureStorageButton.Image = ((System.Drawing.Image)(resources.GetObject("AzureStorageButton.Image")));
            this.AzureStorageButton.Label = "Azure";
            this.AzureStorageButton.Name = "AzureStorageButton";
            this.AzureStorageButton.ShowImage = true;
            this.AzureStorageButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.AzureStorageButton_Click);
            // 
            // box6
            // 
            this.box6.Items.Add(this.SetDirButton);
            this.box6.Name = "box6";
            // 
            // SetDirButton
            // 
            this.SetDirButton.Image = ((System.Drawing.Image)(resources.GetObject("SetDirButton.Image")));
            this.SetDirButton.Label = "Directory (click to change):";
            this.SetDirButton.Name = "SetDirButton";
            this.SetDirButton.ShowImage = true;
            this.SetDirButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.SetDirButton_Click);
            // 
            // ModelDirBox
            // 
            this.ModelDirBox.Enabled = false;
            this.ModelDirBox.Label = " ";
            this.ModelDirBox.Name = "ModelDirBox";
            this.ModelDirBox.SizeString = "C:\\Users\\Mikołaj\\0.NEW_RASA_DATA_FOLD";
            this.ModelDirBox.SuperTip = this.ModelDirBox.Text;
            this.ModelDirBox.Text = null;
            // 
            // ProjectDropDown
            // 
            this.ProjectDropDown.Label = "dropDown1";
            this.ProjectDropDown.Name = "ProjectDropDown";
            // 
            // Ribbon1
            // 
            this.Name = "Ribbon1";
            this.RibbonType = "Microsoft.Word.Document";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon1_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.group2.ResumeLayout(false);
            this.group2.PerformLayout();
            this.group3.ResumeLayout(false);
            this.group3.PerformLayout();
            this.box1.ResumeLayout(false);
            this.box1.PerformLayout();
            this.box4.ResumeLayout(false);
            this.box4.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.box3.ResumeLayout(false);
            this.box3.PerformLayout();
            this.box5.ResumeLayout(false);
            this.box5.PerformLayout();
            this.box6.ResumeLayout(false);
            this.box6.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton ContentControlButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton UnwrapRangeButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton ExportTXTbtn;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton WrapFromTestBtn;
        internal Microsoft.Office.Tools.Ribbon.RibbonDropDown TestModelDropDown;
        internal Microsoft.Office.Tools.Ribbon.RibbonComboBox ProjectComboBox;
        private System.Windows.Forms.FolderBrowserDialog ModelDirDialog;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group2;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group3;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonToggleButton LocalStorageButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonToggleButton AzureStorageButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton SetDirButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button1;
        internal Microsoft.Office.Tools.Ribbon.RibbonBox box4;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox ModelDirBox;
        internal Microsoft.Office.Tools.Ribbon.RibbonBox box1;
        internal Microsoft.Office.Tools.Ribbon.RibbonBox box3;
        internal Microsoft.Office.Tools.Ribbon.RibbonBox box5;
        internal Microsoft.Office.Tools.Ribbon.RibbonBox box6;
        internal Microsoft.Office.Tools.Ribbon.RibbonDropDown ProjectDropDown;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon1 Ribbon1
        {
            get { return this.GetRibbon<Ribbon1>(); }
        }
    }
}
