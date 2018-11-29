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
            this.ModelDirDialog = new System.Windows.Forms.FolderBrowserDialog();
            this.tab1 = this.Factory.CreateRibbonTab();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.ContentControlButton = this.Factory.CreateRibbonButton();
            this.UnwrapRangeButton = this.Factory.CreateRibbonButton();
            this.separator1 = this.Factory.CreateRibbonSeparator();
            this.ProjectComboBox = this.Factory.CreateRibbonComboBox();
            this.ModelComboBox = this.Factory.CreateRibbonComboBox();
            this.ExportTXTbtn = this.Factory.CreateRibbonButton();
            this.separator2 = this.Factory.CreateRibbonSeparator();
            this.TestProjectDropDown = this.Factory.CreateRibbonDropDown();
            this.TestModelDropDown = this.Factory.CreateRibbonDropDown();
            this.WrapFromTestBtn = this.Factory.CreateRibbonButton();
            this.separator3 = this.Factory.CreateRibbonSeparator();
            this.label1 = this.Factory.CreateRibbonLabel();
            this.buttonGroup2 = this.Factory.CreateRibbonButtonGroup();
            this.LocalStorageButton = this.Factory.CreateRibbonToggleButton();
            this.AzureStorageButton = this.Factory.CreateRibbonToggleButton();
            this.box2 = this.Factory.CreateRibbonBox();
            this.SetDirButton = this.Factory.CreateRibbonButton();
            this.ModelDirLabel = this.Factory.CreateRibbonLabel();
            this.tab1.SuspendLayout();
            this.group1.SuspendLayout();
            this.buttonGroup2.SuspendLayout();
            this.box2.SuspendLayout();
            this.SuspendLayout();
            // 
            // ModelDirDialog
            // 
            this.ModelDirDialog.SelectedPath = "C:\\Users\\Mikołaj";
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.group1);
            this.tab1.Label = "TabAddIns";
            this.tab1.Name = "tab1";
            // 
            // group1
            // 
            this.group1.Items.Add(this.ContentControlButton);
            this.group1.Items.Add(this.UnwrapRangeButton);
            this.group1.Items.Add(this.separator1);
            this.group1.Items.Add(this.ProjectComboBox);
            this.group1.Items.Add(this.ModelComboBox);
            this.group1.Items.Add(this.ExportTXTbtn);
            this.group1.Items.Add(this.separator2);
            this.group1.Items.Add(this.TestProjectDropDown);
            this.group1.Items.Add(this.TestModelDropDown);
            this.group1.Items.Add(this.WrapFromTestBtn);
            this.group1.Items.Add(this.separator3);
            this.group1.Items.Add(this.label1);
            this.group1.Items.Add(this.buttonGroup2);
            this.group1.Items.Add(this.box2);
            this.group1.Name = "group1";
            // 
            // ContentControlButton
            // 
            this.ContentControlButton.KeyTip = "W";
            this.ContentControlButton.Label = "Wrap Content";
            this.ContentControlButton.Name = "ContentControlButton";
            this.ContentControlButton.Tag = "controlTag";
            this.ContentControlButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ContentControlButton_Click);
            // 
            // UnwrapRangeButton
            // 
            this.UnwrapRangeButton.Label = "Unwrap Range";
            this.UnwrapRangeButton.Name = "UnwrapRangeButton";
            this.UnwrapRangeButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.UnwrapRangeButton_Click);
            // 
            // separator1
            // 
            this.separator1.Name = "separator1";
            // 
            // ProjectComboBox
            // 
            this.ProjectComboBox.Label = "Project";
            this.ProjectComboBox.MaxLength = 20;
            this.ProjectComboBox.Name = "ProjectComboBox";
            this.ProjectComboBox.TextChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ProjectComboBox_TextChanged);
            // 
            // ModelComboBox
            // 
            this.ModelComboBox.Label = "Model";
            this.ModelComboBox.MaxLength = 20;
            this.ModelComboBox.Name = "ModelComboBox";
            this.ModelComboBox.Text = null;
            this.ModelComboBox.TextChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ModelComboBox_TextChanged);
            // 
            // ExportTXTbtn
            // 
            this.ExportTXTbtn.Label = "EXPORT AND TRAIN";
            this.ExportTXTbtn.Name = "ExportTXTbtn";
            this.ExportTXTbtn.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ExportTXTbtn_Click);
            // 
            // separator2
            // 
            this.separator2.Name = "separator2";
            // 
            // TestProjectDropDown
            // 
            this.TestProjectDropDown.Label = "Test Project";
            this.TestProjectDropDown.Name = "TestProjectDropDown";
            this.TestProjectDropDown.SelectionChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.TestProjectDropDown_Select);
            // 
            // TestModelDropDown
            // 
            this.TestModelDropDown.Label = "Test Model";
            this.TestModelDropDown.Name = "TestModelDropDown";
            this.TestModelDropDown.SelectionChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ModelDropDown_Select);
            // 
            // WrapFromTestBtn
            // 
            this.WrapFromTestBtn.Label = "TEST THIS DOC";
            this.WrapFromTestBtn.Name = "WrapFromTestBtn";
            this.WrapFromTestBtn.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.WrapFromTestBtn_Click);
            // 
            // separator3
            // 
            this.separator3.Name = "separator3";
            // 
            // label1
            // 
            this.label1.Label = "Where to store model?";
            this.label1.Name = "label1";
            // 
            // buttonGroup2
            // 
            this.buttonGroup2.Items.Add(this.LocalStorageButton);
            this.buttonGroup2.Items.Add(this.AzureStorageButton);
            this.buttonGroup2.Name = "buttonGroup2";
            // 
            // LocalStorageButton
            // 
            this.LocalStorageButton.Label = "Locally";
            this.LocalStorageButton.Name = "LocalStorageButton";
            this.LocalStorageButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.LocalStorageButton_Click);
            // 
            // AzureStorageButton
            // 
            this.AzureStorageButton.Label = "Azure";
            this.AzureStorageButton.Name = "AzureStorageButton";
            this.AzureStorageButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.AzureStorageButton_Click);
            // 
            // box2
            // 
            this.box2.Items.Add(this.SetDirButton);
            this.box2.Items.Add(this.ModelDirLabel);
            this.box2.Name = "box2";
            // 
            // SetDirButton
            // 
            this.SetDirButton.Label = "Directory:";
            this.SetDirButton.Name = "SetDirButton";
            this.SetDirButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.SetDirButton_Click);
            // 
            // ModelDirLabel
            // 
            this.ModelDirLabel.Label = "";
            this.ModelDirLabel.Name = "ModelDirLabel";
            // 
            // Ribbon1
            // 
            this.Name = "Ribbon1";
            this.RibbonType = "Microsoft.Word.Document";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon1_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.buttonGroup2.ResumeLayout(false);
            this.buttonGroup2.PerformLayout();
            this.box2.ResumeLayout(false);
            this.box2.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton ContentControlButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton UnwrapRangeButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton ExportTXTbtn;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton WrapFromTestBtn;
        internal Microsoft.Office.Tools.Ribbon.RibbonDropDown TestModelDropDown;
        internal Microsoft.Office.Tools.Ribbon.RibbonDropDown TestProjectDropDown;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator1;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator2;
        internal Microsoft.Office.Tools.Ribbon.RibbonComboBox ProjectComboBox;
        internal Microsoft.Office.Tools.Ribbon.RibbonComboBox ModelComboBox;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator3;
        internal Microsoft.Office.Tools.Ribbon.RibbonLabel label1;
        internal Microsoft.Office.Tools.Ribbon.RibbonToggleButton LocalStorageButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonToggleButton AzureStorageButton;
        private System.Windows.Forms.FolderBrowserDialog ModelDirDialog;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton SetDirButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonLabel ModelDirLabel;
        internal Microsoft.Office.Tools.Ribbon.RibbonButtonGroup buttonGroup2;
        internal Microsoft.Office.Tools.Ribbon.RibbonBox box2;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon1 Ribbon1
        {
            get { return this.GetRibbon<Ribbon1>(); }
        }
    }
}
