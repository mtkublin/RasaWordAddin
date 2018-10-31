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
            System.Windows.Forms.TreeNode treeNode1 = new System.Windows.Forms.TreeNode("Node5");
            System.Windows.Forms.TreeNode treeNode2 = new System.Windows.Forms.TreeNode("Node6");
            System.Windows.Forms.TreeNode treeNode3 = new System.Windows.Forms.TreeNode("Intent A", new System.Windows.Forms.TreeNode[] {
            treeNode1,
            treeNode2});
            System.Windows.Forms.TreeNode treeNode4 = new System.Windows.Forms.TreeNode("Intent B");
            System.Windows.Forms.TreeNode treeNode5 = new System.Windows.Forms.TreeNode("Project I", new System.Windows.Forms.TreeNode[] {
            treeNode3,
            treeNode4});
            System.Windows.Forms.TreeNode treeNode6 = new System.Windows.Forms.TreeNode("Node4");
            System.Windows.Forms.TreeNode treeNode7 = new System.Windows.Forms.TreeNode("Project II", new System.Windows.Forms.TreeNode[] {
            treeNode6});
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
            treeNode1.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(192)))), ((int)(((byte)(192)))), ((int)(((byte)(0)))));
            treeNode1.Name = "Node5";
            treeNode1.Text = "Node5";
            treeNode2.BackColor = System.Drawing.Color.Olive;
            treeNode2.Name = "Node6";
            treeNode2.Text = "Node6";
            treeNode3.BackColor = System.Drawing.Color.Yellow;
            treeNode3.Name = "IntentA";
            treeNode3.Text = "Intent A";
            treeNode4.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(224)))), ((int)(((byte)(192)))));
            treeNode4.Name = "IntentB";
            treeNode4.Text = "Intent B";
            treeNode5.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(255)))), ((int)(((byte)(192)))));
            treeNode5.Name = "ProjectI";
            treeNode5.Text = "Project I";
            treeNode6.BackColor = System.Drawing.Color.Lime;
            treeNode6.Name = "Node4";
            treeNode6.Text = "Node4";
            treeNode7.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(192)))), ((int)(((byte)(255)))), ((int)(((byte)(192)))));
            treeNode7.Name = "ProjectII";
            treeNode7.Text = "Project II";
            this.treeView1.Nodes.AddRange(new System.Windows.Forms.TreeNode[] {
            treeNode5,
            treeNode7});
            this.treeView1.Size = new System.Drawing.Size(217, 270);
            this.treeView1.TabIndex = 0;
            this.treeView1.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.treeView1_KeyPress);
            // 
            // menuStrip1
            // 
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.toolStripLabel,
            this.toolStripNewTag});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Size = new System.Drawing.Size(217, 27);
            this.menuStrip1.TabIndex = 1;
            this.menuStrip1.Text = "menuStrip1";
            // 
            // toolStripLabel
            // 
            this.toolStripLabel.Name = "toolStripLabel";
            this.toolStripLabel.ReadOnly = true;
            this.toolStripLabel.Size = new System.Drawing.Size(100, 23);
            this.toolStripLabel.Text = "Add Tag:";
            // 
            // toolStripNewTag
            // 
            this.toolStripNewTag.AcceptsReturn = true;
            this.toolStripNewTag.Name = "toolStripNewTag";
            this.toolStripNewTag.Size = new System.Drawing.Size(100, 23);
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
