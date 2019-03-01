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
            this.HighlightGroup = this.Factory.CreateRibbonGroup();
            this.HighlightInVisibleBTN = this.Factory.CreateRibbonButton();
            this.HighlightInNextVisibleBTN = this.Factory.CreateRibbonButton();
            this.CurBMgroup = this.Factory.CreateRibbonGroup();
            this.box2 = this.Factory.CreateRibbonBox();
            this.label1 = this.Factory.CreateRibbonLabel();
            this.CurBMtextLabel = this.Factory.CreateRibbonLabel();
            this.box6 = this.Factory.CreateRibbonBox();
            this.label2 = this.Factory.CreateRibbonLabel();
            this.CurBMentLabel = this.Factory.CreateRibbonLabel();
            this.box7 = this.Factory.CreateRibbonBox();
            this.label3 = this.Factory.CreateRibbonLabel();
            this.IntOrEntLabel = this.Factory.CreateRibbonLabel();
            this.tab1.SuspendLayout();
            this.WrapperGroup.SuspendLayout();
            this.TestTrainGroup.SuspendLayout();
            this.box1.SuspendLayout();
            this.box4.SuspendLayout();
            this.StorageGroup.SuspendLayout();
            this.box3.SuspendLayout();
            this.box5.SuspendLayout();
            this.HighlightGroup.SuspendLayout();
            this.CurBMgroup.SuspendLayout();
            this.box2.SuspendLayout();
            this.box6.SuspendLayout();
            this.box7.SuspendLayout();
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
            this.tab1.Groups.Add(this.HighlightGroup);
            this.tab1.Groups.Add(this.CurBMgroup);
            this.tab1.Label = "RasaNLU addin";
            this.tab1.Name = "tab1";
            // 
            // WrapperGroup
            // 
            this.WrapperGroup.Items.Add(this.ContentControlButton);
            this.WrapperGroup.Items.Add(this.UnwrapRangeButton);
            this.WrapperGroup.Label = "Wrapper";
            this.WrapperGroup.Name = "WrapperGroup";
            this.WrapperGroup.Visible = false;
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
            this.box5.Visible = false;
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
            // HighlightGroup
            // 
            this.HighlightGroup.Items.Add(this.HighlightInVisibleBTN);
            this.HighlightGroup.Items.Add(this.HighlightInNextVisibleBTN);
            this.HighlightGroup.Label = "Highlight";
            this.HighlightGroup.Name = "HighlightGroup";
            // 
            // HighlightInVisibleBTN
            // 
            this.HighlightInVisibleBTN.Image = ((System.Drawing.Image)(resources.GetObject("HighlightInVisibleBTN.Image")));
            this.HighlightInVisibleBTN.Label = "Highlight visible";
            this.HighlightInVisibleBTN.Name = "HighlightInVisibleBTN";
            this.HighlightInVisibleBTN.ShowImage = true;
            this.HighlightInVisibleBTN.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.HighlightInVisibleBTN_Click);
            // 
            // HighlightInNextVisibleBTN
            // 
            this.HighlightInNextVisibleBTN.Image = ((System.Drawing.Image)(resources.GetObject("HighlightInNextVisibleBTN.Image")));
            this.HighlightInNextVisibleBTN.Label = "Highlight next";
            this.HighlightInNextVisibleBTN.Name = "HighlightInNextVisibleBTN";
            this.HighlightInNextVisibleBTN.ShowImage = true;
            this.HighlightInNextVisibleBTN.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.HighlightInNextVisibleBTN_Click);
            // 
            // CurBMgroup
            // 
            this.CurBMgroup.Items.Add(this.box2);
            this.CurBMgroup.Items.Add(this.box6);
            this.CurBMgroup.Items.Add(this.box7);
            this.CurBMgroup.Label = "Current Bookmark";
            this.CurBMgroup.Name = "CurBMgroup";
            // 
            // box2
            // 
            this.box2.Items.Add(this.label1);
            this.box2.Items.Add(this.CurBMtextLabel);
            this.box2.Name = "box2";
            // 
            // label1
            // 
            this.label1.Label = "Text:";
            this.label1.Name = "label1";
            // 
            // CurBMtextLabel
            // 
            this.CurBMtextLabel.Label = "label1";
            this.CurBMtextLabel.Name = "CurBMtextLabel";
            // 
            // box6
            // 
            this.box6.Items.Add(this.label2);
            this.box6.Items.Add(this.CurBMentLabel);
            this.box6.Name = "box6";
            // 
            // label2
            // 
            this.label2.Label = "Tag:";
            this.label2.Name = "label2";
            // 
            // CurBMentLabel
            // 
            this.CurBMentLabel.Label = "label2";
            this.CurBMentLabel.Name = "CurBMentLabel";
            // 
            // box7
            // 
            this.box7.Items.Add(this.label3);
            this.box7.Items.Add(this.IntOrEntLabel);
            this.box7.Name = "box7";
            // 
            // label3
            // 
            this.label3.Label = "Type:";
            this.label3.Name = "label3";
            // 
            // IntOrEntLabel
            // 
            this.IntOrEntLabel.Label = "Intent";
            this.IntOrEntLabel.Name = "IntOrEntLabel";
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
            this.HighlightGroup.ResumeLayout(false);
            this.HighlightGroup.PerformLayout();
            this.CurBMgroup.ResumeLayout(false);
            this.CurBMgroup.PerformLayout();
            this.box2.ResumeLayout(false);
            this.box2.PerformLayout();
            this.box6.ResumeLayout(false);
            this.box6.PerformLayout();
            this.box7.ResumeLayout(false);
            this.box7.PerformLayout();
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
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup HighlightGroup;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton HighlightInVisibleBTN;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton HighlightInNextVisibleBTN;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup CurBMgroup;
        internal Microsoft.Office.Tools.Ribbon.RibbonLabel CurBMtextLabel;
        internal Microsoft.Office.Tools.Ribbon.RibbonLabel CurBMentLabel;
        internal Microsoft.Office.Tools.Ribbon.RibbonBox box2;
        internal Microsoft.Office.Tools.Ribbon.RibbonLabel label1;
        internal Microsoft.Office.Tools.Ribbon.RibbonBox box6;
        internal Microsoft.Office.Tools.Ribbon.RibbonLabel label2;
        internal Microsoft.Office.Tools.Ribbon.RibbonBox box7;
        internal Microsoft.Office.Tools.Ribbon.RibbonLabel label3;
        internal Microsoft.Office.Tools.Ribbon.RibbonLabel IntOrEntLabel;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon1 Ribbon1
        {
            get { return this.GetRibbon<Ribbon1>(); }
        }
    }
}
