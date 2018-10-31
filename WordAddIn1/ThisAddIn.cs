using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Word = Microsoft.Office.Interop.Word;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools;
using XL.Office.Helpers;
using System.Windows.Forms;

namespace WordAddIn1
{
    public partial class ThisAddIn
    {
        private PaneControl TaskPaneControl;
        private CustomTaskPane AppTaskPane;
        private Dictionary<KeyState, KeyHandlerDelegate> KeyHandlers;
        
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {        
            var ctrlW = new KeyState(Keys.W, ctrl: true);
            var ctrlshiftW = new KeyState(Keys.W, ctrl: true, shift: true);
            var ctrlshiftT = new KeyState(Keys.T, ctrl: true, shift: true);
            const int STOP = 1;

            KeyHandlers = new Dictionary<KeyState, KeyHandlerDelegate>
            {
                {
                    ctrlW, (bool repeated) =>
                    {
                        if (!repeated) WrapContent();
                        return STOP;
                    }
                },
                {
                    ctrlshiftW, (bool repeated) =>
                    {
                        if (!repeated) UnwrapContent();
                        return STOP;
                    }
                },
                {
                    ctrlshiftT, (bool repeated) =>
                    {
                        if(!repeated)
                        {
                            AppTaskPane.Visible = !AppTaskPane.Visible;
                        }
                        return STOP;
                    }
                }
            };
            InterceptKeys.SetHooks(KeyHandlers);

            Addin_Setup(null);
        }

        private void Addin_Setup(Word.Document Doc)
        {
            TaskPaneControl = new PaneControl();
            AppTaskPane = CustomTaskPanes.Add(TaskPaneControl, "Annotation Task Pane");
            AppTaskPane.Visible = true;

            Application.ActiveDocument.ContentControlOnExit += HighlightContentControl;

            if (Doc == null)
            {
                Application.DocumentOpen += Addin_Setup;
            }
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            InterceptKeys.ReleaseHook();
        }

        private void HighlightContentControl(Word.ContentControl control, ref bool cancel)
        {
            control.Range.Font.Color = Utilities.RGBwdColor(128, 0, 0);
            control.Range.Font.Shading.BackgroundPatternColor = Word.WdColor.wdColorLightYellow;
            control.Range.HighlightColorIndex = Word.WdColorIndex.wdNoHighlight;
        }

        public void WrapContent()
        {
            var activeDocument = Application.ActiveDocument;
            var extendedDocument = Globals.Factory.GetVstoObject(activeDocument);
            var next = DateTime.Now.Ticks.ToString();
            var control = extendedDocument.Controls.AddRichTextContentControl(string.Format("richText{0}", next));
            control.PlaceholderText = "This cannot be empty";
            control.Range.Font.Color = (Word.WdColor)128;
            control.Range.Font.Shading.BackgroundPatternColor = Word.WdColor.wdColorLightYellow;
            control.Range.HighlightColorIndex = Word.WdColorIndex.wdNoHighlight;
        }

        public void UnwrapContent()
        {
            var selection = Application.Selection;
            if (selection.ContentControls.Count > 0)
            {
                foreach (Word.ContentControl control in selection.ContentControls)
                {
                    control.Delete(false);
                }
            }
            if (selection.ParentContentControl != null)
            {
                selection.ParentContentControl.Delete(false);
            }
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
