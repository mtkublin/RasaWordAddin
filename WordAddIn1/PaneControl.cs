using System.Windows.Forms;

namespace WordAddIn1
{
    public partial class PaneControl : UserControl
    {
        public PaneControl()
        {
            InitializeComponent();
        }

        private void treeView1_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == '+')
            {
                TreeView tree = (TreeView)sender;
                tree.Nodes.Add("<new>");
            }
        }

        private void toolStripNewTag_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == 13)
            { 
                treeView1.Nodes.Add(toolStripNewTag.Text);
                toolStripNewTag.Text = string.Empty;
            }
        }
    }
}
