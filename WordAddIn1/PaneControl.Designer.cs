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
            this.treeView1 = new System.Windows.Forms.TreeView();
            this.menuStrip1 = new System.Windows.Forms.MenuStrip();
            this.toolStripNewTag = new System.Windows.Forms.ToolStripTextBox();
            this.toolStripLabel = new System.Windows.Forms.ToolStripTextBox();
            this.menuStrip1.SuspendLayout();
            this.SuspendLayout();
            // 
            // treeView1
            // 
            this.treeView1.Cursor = System.Windows.Forms.Cursors.Hand;
            this.treeView1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.treeView1.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.treeView1.FullRowSelect = true;
            this.treeView1.HideSelection = false;
            this.treeView1.HotTracking = true;
            this.treeView1.Indent = 20;
            this.treeView1.ItemHeight = 28;
            this.treeView1.LabelEdit = true;
            this.treeView1.Location = new System.Drawing.Point(0, 27);
            this.treeView1.Name = "treeView1";
            this.treeView1.Size = new System.Drawing.Size(217, 513);
            this.treeView1.TabIndex = 0;
            this.treeView1.KeyDown += new System.Windows.Forms.KeyEventHandler(this.treeView1_KeyDown);
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
            // PaneControl
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.treeView1);
            this.Controls.Add(this.menuStrip1);
            this.Name = "PaneControl";
            this.Size = new System.Drawing.Size(217, 540);
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.ToolStripTextBox toolStripLabel;
        private System.Windows.Forms.ToolStripTextBox toolStripNewTag;
        public System.Windows.Forms.TreeView treeView1;
    }
}
