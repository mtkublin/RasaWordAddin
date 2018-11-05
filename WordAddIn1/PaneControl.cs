using System.Windows.Forms;
using XL.Office.Helpers;

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

        private void treeView1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Control)
            {
                switch(e.KeyCode)
                {
                    case Keys.Right:
                        Utilities.Notification("Right");
                        break;
                    case Keys.Left:
                        Utilities.Notification("Left");
                        break;
                    case Keys.Up:
                        Utilities.Notification("Up");
                        break;
                    case Keys.Down:
                        Utilities.Notification("Down");
                        break;
                    default:
                        Utilities.Notification("Other");
                        break;
                }
            }
        }
    }
}
