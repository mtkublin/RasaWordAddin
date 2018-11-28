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
            this.tab1 = this.Factory.CreateRibbonTab();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.ContentControlButton = this.Factory.CreateRibbonButton();
            this.UnwrapRangeButton = this.Factory.CreateRibbonButton();
            this.ExportTXTbtn = this.Factory.CreateRibbonButton();
            this.TestProjectDropDown = this.Factory.CreateRibbonDropDown();
            this.TestModelDropDown = this.Factory.CreateRibbonDropDown();
            this.WrapFromTestBtn = this.Factory.CreateRibbonButton();
            this.ProjectDropDown = this.Factory.CreateRibbonDropDown();
            this.ModelBox = this.Factory.CreateRibbonEditBox();
            this.separator1 = this.Factory.CreateRibbonSeparator();
            this.separator2 = this.Factory.CreateRibbonSeparator();
            this.tab1.SuspendLayout();
            this.group1.SuspendLayout();
            this.SuspendLayout();
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
            this.group1.Items.Add(this.ProjectDropDown);
            this.group1.Items.Add(this.ModelBox);
            this.group1.Items.Add(this.ExportTXTbtn);
            this.group1.Items.Add(this.separator2);
            this.group1.Items.Add(this.TestProjectDropDown);
            this.group1.Items.Add(this.TestModelDropDown);
            this.group1.Items.Add(this.WrapFromTestBtn);
            this.group1.Label = "group1";
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
            // ProjectDropDown
            // 
            this.ProjectDropDown.Label = "Project";
            this.ProjectDropDown.Name = "ProjectDropDown";
            this.ProjectDropDown.SelectionChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ProjectDropDown_Select);
            // 
            // ModelBox
            // 
            this.ModelBox.Label = "Model";
            this.ModelBox.Name = "ModelBox";
            this.ModelBox.TextChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ModelBox_TextChanged);
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
        internal Microsoft.Office.Tools.Ribbon.RibbonDropDown ProjectDropDown;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox ModelBox;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator2;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon1 Ribbon1
        {
            get { return this.GetRibbon<Ribbon1>(); }
        }
    }
}
