using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
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
    }
}
