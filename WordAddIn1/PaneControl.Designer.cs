namespace WordAddIn1
{
    partial class PaneControl
    {
        /// <summary> 
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

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
            System.Windows.Forms.TreeNode treeNode8 = new System.Windows.Forms.TreeNode("Node5");
            System.Windows.Forms.TreeNode treeNode9 = new System.Windows.Forms.TreeNode("Node6");
            System.Windows.Forms.TreeNode treeNode10 = new System.Windows.Forms.TreeNode("Intent A", new System.Windows.Forms.TreeNode[] {
            treeNode8,
            treeNode9});
            System.Windows.Forms.TreeNode treeNode11 = new System.Windows.Forms.TreeNode("Intent B");
            System.Windows.Forms.TreeNode treeNode12 = new System.Windows.Forms.TreeNode("Project I", new System.Windows.Forms.TreeNode[] {
            treeNode10,
            treeNode11});
            System.Windows.Forms.TreeNode treeNode13 = new System.Windows.Forms.TreeNode("Node4");
            System.Windows.Forms.TreeNode treeNode14 = new System.Windows.Forms.TreeNode("Project II", new System.Windows.Forms.TreeNode[] {
            treeNode13});
            this.treeView1 = new System.Windows.Forms.TreeView();
            this.menuStrip1 = new System.Windows.Forms.MenuStrip();
            this.toolStripLabel = new System.Windows.Forms.ToolStripTextBox();
            this.toolStripNewTag = new System.Windows.Forms.ToolStripTextBox();
            this.menuStrip1.SuspendLayout();
            this.SuspendLayout();
            // 
            // treeView1
            // 
            this.treeView1.CheckBoxes = true;
            this.treeView1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.treeView1.FullRowSelect = true;
            this.treeView1.ItemHeight = 24;
            this.treeView1.LabelEdit = true;
            this.treeView1.Location = new System.Drawing.Point(0, 27);
            this.treeView1.Name = "treeView1";
            treeNode8.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(192)))), ((int)(((byte)(192)))), ((int)(((byte)(0)))));
            treeNode8.Name = "Node5";
            treeNode8.Text = "Node5";
            treeNode9.BackColor = System.Drawing.Color.Olive;
            treeNode9.Name = "Node6";
            treeNode9.Text = "Node6";
            treeNode10.BackColor = System.Drawing.Color.Yellow;
            treeNode10.Name = "IntentA";
            treeNode10.Text = "Intent A";
            treeNode11.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(224)))), ((int)(((byte)(192)))));
            treeNode11.Name = "IntentB";
            treeNode11.Text = "Intent B";
            treeNode12.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(255)))), ((int)(((byte)(192)))));
            treeNode12.Name = "ProjectI";
            treeNode12.Text = "Project I";
            treeNode13.BackColor = System.Drawing.Color.Lime;
            treeNode13.Name = "Node4";
            treeNode13.Text = "Node4";
            treeNode14.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(192)))), ((int)(((byte)(255)))), ((int)(((byte)(192)))));
            treeNode14.Name = "ProjectII";
            treeNode14.Text = "Project II";
            this.treeView1.Nodes.AddRange(new System.Windows.Forms.TreeNode[] {
            treeNode12,
            treeNode14});
            this.treeView1.Size = new System.Drawing.Size(217, 270);
            this.treeView1.TabIndex = 0;
            this.treeView1.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.treeView1_KeyPress);
            // 
            // menuStrip1
            // 
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.toolStripNewTag,
            this.toolStripLabel});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Size = new System.Drawing.Size(217, 27);
            this.menuStrip1.TabIndex = 1;
            this.menuStrip1.Text = "menuStrip1";
            // 
            // toolStripLabel
            // 
            this.toolStripLabel.Alignment = System.Windows.Forms.ToolStripItemAlignment.Right;
            this.toolStripLabel.AutoSize = false;
            this.toolStripLabel.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.toolStripLabel.Margin = new System.Windows.Forms.Padding(1, 0, 10, 0);
            this.toolStripLabel.Name = "toolStripLabel";
            this.toolStripLabel.ReadOnly = true;
            this.toolStripLabel.Size = new System.Drawing.Size(60, 16);
            this.toolStripLabel.Text = "Add Tag:";
            this.toolStripLabel.TextBoxTextAlign = System.Windows.Forms.HorizontalAlignment.Right;
            // 
            // toolStripNewTag
            // 
            this.toolStripNewTag.AcceptsReturn = true;
            this.toolStripNewTag.Alignment = System.Windows.Forms.ToolStripItemAlignment.Right;
            this.toolStripNewTag.AutoSize = false;
            this.toolStripNewTag.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.toolStripNewTag.Name = "toolStripNewTag";
            this.toolStripNewTag.Size = new System.Drawing.Size(120, 23);
            this.toolStripNewTag.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.toolStripNewTag_KeyPress);
            // 
            // PaneControl
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.treeView1);
            this.Controls.Add(this.menuStrip1);
            this.Name = "PaneControl";
            this.Size = new System.Drawing.Size(217, 297);
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TreeView treeView1;
        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.ToolStripTextBox toolStripLabel;
        private System.Windows.Forms.ToolStripTextBox toolStripNewTag;
    }
}
