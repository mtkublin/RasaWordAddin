using Microsoft.Office.Tools;
using System;
using System.Collections.Generic;
using System.Windows.Forms;
using XL.Office.Helpers;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1
{
    public partial class ThisAddIn
    {
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
                            var activeWindow = Application.ActiveWindow;
                            var taskPane = WindowTaskPanes[activeWindow];
                            taskPane.Visible = !taskPane.Visible;
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
            var taskPane = CustomTaskPanes.Add(new PaneControl(), "Annotation Task Pane (setup)");
            taskPane.Visible = true;

            WindowTaskPanes = new Dictionary<Word.Window, CustomTaskPane>();
            var activeWindow = Application.ActiveWindow;
            WindowTaskPanes.Add(activeWindow, taskPane);

            Application.ActiveDocument.ContentControlOnExit -= HighlightContentControl;
            Application.ActiveDocument.ContentControlOnExit += HighlightContentControl;

            if (Doc == null)
            {
                Application.WindowActivate += Application_WindowActivate;
            }
        }

        private Dictionary<Word.Window, CustomTaskPane> WindowTaskPanes;

        private void Application_WindowActivate(Word.Document Doc, Word.Window Wn)
        {
            Application.ActiveDocument.ContentControlOnExit -= HighlightContentControl;
            Application.ActiveDocument.ContentControlOnExit += HighlightContentControl;

            var activeWindow = Application.ActiveWindow;
            if (!WindowTaskPanes.ContainsKey(activeWindow))
            {
                var taskPane = CustomTaskPanes.Add(new PaneControl(), "Annotation Task Pane (activate)");
                taskPane.Visible = true;
                WindowTaskPanes.Add(activeWindow, taskPane);
            }

            Dictionary<Word.Window, CustomTaskPane> tempTaskPains = new Dictionary<Word.Window, CustomTaskPane>();
            foreach (Word.Window window in Application.Windows)
            {
                tempTaskPains.Add(window, WindowTaskPanes[window]);
                WindowTaskPanes.Remove(window);
            }
            foreach (CustomTaskPane pane in WindowTaskPanes.Values)
            {
                CustomTaskPanes.Remove(pane);
            }
            WindowTaskPanes = tempTaskPains;
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
