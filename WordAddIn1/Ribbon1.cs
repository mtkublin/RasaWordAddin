using System.IO;
using Microsoft.Office.Tools.Ribbon;

namespace WordAddIn1
{
    public partial class Ribbon1
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void ContentControlButton_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.WrapContent();
        }

        private void UnwrapRangeButton_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.UnwrapContent();
        }

        private void ExportTXTbtn_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.ExportTrainData();
        }

        private void WrapFromTestBtn_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.TestDoc();
        }
    }
}