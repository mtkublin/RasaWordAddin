﻿namespace WordAddIn1
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
            this.WrapperGroup = this.Factory.CreateRibbonGroup();
            this.ContentControlButton = this.Factory.CreateRibbonButton();
            this.UnwrapRangeButton = this.Factory.CreateRibbonButton();
            this.TestTrainGroup = this.Factory.CreateRibbonGroup();
            this.box1 = this.Factory.CreateRibbonBox();
            this.ProjectAddButton = this.Factory.CreateRibbonButton();
            this.TestButton = this.Factory.CreateRibbonButton();
            this.TrainingButton = this.Factory.CreateRibbonButton();
            this.box4 = this.Factory.CreateRibbonBox();
            this.ProjectDropDown = this.Factory.CreateRibbonDropDown();
            this.TestModelDropDown = this.Factory.CreateRibbonDropDown();
            this.StorageGroup = this.Factory.CreateRibbonGroup();
            this.box3 = this.Factory.CreateRibbonBox();
            this.box5 = this.Factory.CreateRibbonBox();
            this.LocalStorageButton = this.Factory.CreateRibbonToggleButton();
            this.AzureStorageButton = this.Factory.CreateRibbonToggleButton();
            this.SetDirButton = this.Factory.CreateRibbonButton();
            this.ModelDirBox = this.Factory.CreateRibbonEditBox();
            this.tab1.SuspendLayout();
            this.WrapperGroup.SuspendLayout();
            this.TestTrainGroup.SuspendLayout();
            this.box1.SuspendLayout();
            this.box4.SuspendLayout();
            this.StorageGroup.SuspendLayout();
            this.box3.SuspendLayout();
            this.box5.SuspendLayout();
            this.SuspendLayout();
            // 
            // ModelDirDialog
            // 
            this.ModelDirDialog.SelectedPath = "C:\\Users\\Mikołaj";
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.WrapperGroup);
            this.tab1.Groups.Add(this.TestTrainGroup);
            this.tab1.Groups.Add(this.StorageGroup);
            this.tab1.Label = "RasaNLU addin";
            this.tab1.Name = "tab1";
            // 
            // WrapperGroup
            // 
            this.WrapperGroup.Items.Add(this.ContentControlButton);
            this.WrapperGroup.Items.Add(this.UnwrapRangeButton);
            this.WrapperGroup.Label = "Wrapper";
            this.WrapperGroup.Name = "WrapperGroup";
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
            // TestTrainGroup
            // 
            this.TestTrainGroup.Items.Add(this.box1);
            this.TestTrainGroup.Items.Add(this.box4);
            this.TestTrainGroup.Label = "Test/Train";
            this.TestTrainGroup.Name = "TestTrainGroup";
            // 
            // box1
            // 
            this.box1.BoxStyle = Microsoft.Office.Tools.Ribbon.RibbonBoxStyle.Vertical;
            this.box1.Items.Add(this.ProjectAddButton);
            this.box1.Items.Add(this.TestButton);
            this.box1.Items.Add(this.TrainingButton);
            this.box1.Name = "box1";
            // 
            // ProjectAddButton
            // 
            this.ProjectAddButton.Image = ((System.Drawing.Image)(resources.GetObject("ProjectAddButton.Image")));
            this.ProjectAddButton.Label = "Project (Add)";
            this.ProjectAddButton.Name = "ProjectAddButton";
            this.ProjectAddButton.ShowImage = true;
            this.ProjectAddButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ProjectAddButton_Click);
            // 
            // TestButton
            // 
            this.TestButton.Image = ((System.Drawing.Image)(resources.GetObject("TestButton.Image")));
            this.TestButton.Label = "Test with";
            this.TestButton.Name = "TestButton";
            this.TestButton.ShowImage = true;
            this.TestButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.TestButton_Click);
            // 
            // TrainingButton
            // 
            this.TrainingButton.Image = ((System.Drawing.Image)(resources.GetObject("TrainingButton.Image")));
            this.TrainingButton.Label = "Train";
            this.TrainingButton.Name = "TrainingButton";
            this.TrainingButton.ShowImage = true;
            this.TrainingButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.TrainingButton_Click);
            // 
            // box4
            // 
            this.box4.BoxStyle = Microsoft.Office.Tools.Ribbon.RibbonBoxStyle.Vertical;
            this.box4.Items.Add(this.ProjectDropDown);
            this.box4.Items.Add(this.TestModelDropDown);
            this.box4.Name = "box4";
            // 
            // ProjectDropDown
            // 
            this.ProjectDropDown.Label = " ";
            this.ProjectDropDown.Name = "ProjectDropDown";
            this.ProjectDropDown.SizeString = "model_20181017-154908aa";
            this.ProjectDropDown.SelectionChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ProjectDropDown_Select);
            // 
            // TestModelDropDown
            // 
            this.TestModelDropDown.Label = " ";
            this.TestModelDropDown.Name = "TestModelDropDown";
            this.TestModelDropDown.ShowItemImage = false;
            this.TestModelDropDown.SizeString = "model_20181017-154908aa";
            this.TestModelDropDown.SelectionChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ModelDropDown_Select);
            // 
            // StorageGroup
            // 
            this.StorageGroup.Items.Add(this.box3);
            this.StorageGroup.Label = "Storage";
            this.StorageGroup.Name = "StorageGroup";
            // 
            // box3
            // 
            this.box3.BoxStyle = Microsoft.Office.Tools.Ribbon.RibbonBoxStyle.Vertical;
            this.box3.Items.Add(this.box5);
            this.box3.Items.Add(this.SetDirButton);
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
            // Ribbon1
            // 
            this.Name = "Ribbon1";
            this.RibbonType = "Microsoft.Word.Document";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon1_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.WrapperGroup.ResumeLayout(false);
            this.WrapperGroup.PerformLayout();
            this.TestTrainGroup.ResumeLayout(false);
            this.TestTrainGroup.PerformLayout();
            this.box1.ResumeLayout(false);
            this.box1.PerformLayout();
            this.box4.ResumeLayout(false);
            this.box4.PerformLayout();
            this.StorageGroup.ResumeLayout(false);
            this.StorageGroup.PerformLayout();
            this.box3.ResumeLayout(false);
            this.box3.PerformLayout();
            this.box5.ResumeLayout(false);
            this.box5.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton ContentControlButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton UnwrapRangeButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton TrainingButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton TestButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonDropDown TestModelDropDown;
        private System.Windows.Forms.FolderBrowserDialog ModelDirDialog;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup WrapperGroup;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup TestTrainGroup;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup StorageGroup;
        internal Microsoft.Office.Tools.Ribbon.RibbonToggleButton LocalStorageButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonToggleButton AzureStorageButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton SetDirButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton ProjectAddButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonBox box4;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox ModelDirBox;
        internal Microsoft.Office.Tools.Ribbon.RibbonBox box1;
        internal Microsoft.Office.Tools.Ribbon.RibbonBox box3;
        internal Microsoft.Office.Tools.Ribbon.RibbonBox box5;
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
