using Microsoft.Office.Interop.Word;
using Microsoft.Office.Tools;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using System.Xml;
using System.Xml.Linq;
using XL.Office.Helpers;
using System.Text.RegularExpressions;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1
{
    public partial class ThisAddIn
    {
        private XElement XmlDocument;
        private Dictionary<string, Color> TagColors;
        private Dictionary<KeyState, KeyHandlerDelegate> KeyHandlers;
        private Dictionary<Word.Window, CustomTaskPane> WindowTaskPanes;

        private string CurrentProject;
        public string CurrentTag;
        public string CurrentName;
        private TreeNode CurrentNode;
        private string CurrentPath;

        private Word.WdColor TagBackColor(string tag)
        {
            Color color = TagColors[tag];
            return Utilities.RGBwdColor(color);
        }

        private Word.WdColor TagForeColor(string tag)
        {
            Color color = TagColors[tag];
            return Utilities.RGBwdColor(Utilities.Contrast(color));
        }

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            WindowTaskPanes = new Dictionary<Word.Window, CustomTaskPane>();
            XmlDocument = XElement.Load(new XmlNodeReader(Properties.Settings.Default.Projects));
            CurrentProject = Properties.Settings.Default.RecentProject;
            CurrentTag = Properties.Settings.Default.RecentTag;
            TagColors = new Dictionary<string, Color>();

            Application.ActiveDocument.Bookmarks.ShowHidden = true;

            KeyboardShortcuts();
            Application.WindowActivate += ActivateDocumentWindow;
            Application.WindowDeactivate += DeactivateDocumentWindow;

            Microsoft.Office.Tools.Word.Document vstoDoc = Globals.Factory.GetVstoObject(Application.ActiveDocument);
            var handler = new Microsoft.Office.Tools.Word.SelectionEventHandler((sender2, e2) => ThisDocument_SelectionChange(sender2, e2, currentBookmark, vstoDoc));
            vstoDoc.SelectionChange -= handler;
            vstoDoc.SelectionChange += handler;

            this.Application.DocumentOpen += new ApplicationEvents4_DocumentOpenEventHandler(DocOpen); 
            ((ApplicationEvents4_Event)this.Application).NewDocument += new ApplicationEvents4_NewDocumentEventHandler(NewDoc);

            string AppDir = System.AppDomain.CurrentDomain.BaseDirectory.ToString();
            string TestDocDir = AppDir.Substring(0, AppDir.Length - 21) + @"Docs\test_data.txt";
            string TestText;
            using (StreamReader r = new StreamReader(TestDocDir))
            {
                TestText = r.ReadToEnd();
            }
            Application.ActiveDocument.Content.Text = TestText;
        }

        private void DocOpen(Document OpenedDoc)
        {
            Microsoft.Office.Tools.Word.Document vstoDoc = Globals.Factory.GetVstoObject(OpenedDoc);
            var handler = new Microsoft.Office.Tools.Word.SelectionEventHandler((sender2, e2) => ThisDocument_SelectionChange(sender2, e2, currentBookmark, vstoDoc));
            vstoDoc.SelectionChange -= handler;
            vstoDoc.SelectionChange += handler;

            foreach (Bookmark bookmark in OpenedDoc.Bookmarks)
            { 
                Microsoft.Office.Tools.Word.Bookmark VSTObookmark;
                if (vstoDoc.Controls.Contains(bookmark.Name))
                {
                    VSTObookmark = vstoDoc.Controls[bookmark.Name] as Microsoft.Office.Tools.Word.Bookmark;
                }
                else
                {
                    VSTObookmark = vstoDoc.Controls.AddBookmark(bookmark, bookmark.Name);
                }
                    
                var VSTOselectedHandler = new Microsoft.Office.Tools.Word.SelectionEventHandler((sender2, e2) => bookmark_Selected(sender2, e2, vstoDoc, VSTObookmark));
                VSTObookmark.Selected -= VSTOselectedHandler;
                VSTObookmark.Selected += VSTOselectedHandler;
                if (VSTObookmark.Name.EndsWith("1"))
                {
                    VSTObookmark.SelectionChange += new Microsoft.Office.Tools.Word.SelectionEventHandler((sender2, e2) => bookmark_SelectionChange(sender2, e2, vstoDoc, VSTObookmark));
                }
            }
            
        }

        private void NewDoc(Document NewDoc)
        {
            Microsoft.Office.Tools.Word.Document vstoDoc = Globals.Factory.GetVstoObject(NewDoc);
            var handler = new Microsoft.Office.Tools.Word.SelectionEventHandler((sender2, e2) => ThisDocument_SelectionChange(sender2, e2, currentBookmark, vstoDoc));
            vstoDoc.SelectionChange -= handler;
            vstoDoc.SelectionChange += handler;
        }

        private void KeyboardShortcuts()
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
        }

        private TreeNode[] GetTreeNodes()
        {
            IEnumerable<XElement> projects = from item in XmlDocument.Descendants("Project") select item;

            int intentCount = (from item in XmlDocument.Descendants("Intention") select item).Count();
            int entityCount = (from item in XmlDocument.Descendants("Entity") select item).Count();
            Color projectColor = Color.Blue;
            Color intentColor = Color.Yellow;
            Color entityColor = Color.YellowGreen;
            IEnumerator<Color> intentColors = Utilities.Gradient(intentColor, entityColor, intentCount);
            IEnumerator<Color> entColors = Utilities.Gradient(entityColor, projectColor, entityCount);

            TreeNode[] projectNodes = new TreeNode[projects.Count()];
            int index = 0;

            foreach (XElement project in projects)
            {
                string projectName = (string)project.Attribute("Name");
                TreeNode projectNode = new TreeNode(projectName);
                IEnumerable<XElement> intentions = from item in project.Descendants("Intention") select item;
                foreach (XElement intention in intentions)
                {
                    string intentName = (string)intention.Attribute("Name");
                    string intentTag = (string)intention.Attribute("Tag");
                    if (intentColors.MoveNext())
                    {
                        intentColor = intentColors.Current;
                    }
                    if (!TagColors.ContainsKey(intentTag))
                    {
                        TagColors.Add(intentTag, intentColor);
                    }
                    TreeNode intentionNode = new TreeNode(intentName)
                    {
                        ForeColor = Utilities.Contrast(intentColor),
                        BackColor = intentColor,
                        Tag = intentTag
                    };

                    IEnumerable<XElement> entities = from item in intention.Descendants("Entity") select item;
                    foreach (XElement entity in entities)
                    {
                        string entityName = (string)entity.Attribute("Name");
                        string entityTag = (string)entity.Attribute("Tag");
                        if (entColors.MoveNext())
                        {
                            entityColor = entColors.Current;
                        }
                        if (!TagColors.ContainsKey(entityTag))
                        {
                            TagColors.Add(entityTag, entityColor);
                        }
                        TreeNode entityNode = new TreeNode(entityName)
                        {
                            ForeColor = Utilities.Contrast(entityColor),
                            BackColor = entityColor,
                            Tag = entityTag
                        };
                        intentionNode.Nodes.Add(entityNode);
                    }

                    projectNode.Nodes.Add(intentionNode);
                }

                if (projectName == CurrentProject)
                {
                    projectNode.ExpandAll();
                    projectNode.ForeColor = Utilities.Contrast(projectColor);
                    projectNode.BackColor = projectColor;
                }
                else
                {
                    projectNode.Collapse();
                    projectNode.ForeColor = Color.Gray;
                    projectNode.BackColor = Color.LightGray;
                }

                projectNodes[index++] = projectNode;
            }

            return projectNodes;
        }

        private void ActivateDocumentWindow(Word.Document Doc, Word.Window activeWindow)
        {
            PaneControl paneControl;
            if (WindowTaskPanes.ContainsKey(activeWindow))
            {
                paneControl = WindowTaskPanes[activeWindow].Control as PaneControl;
            }
            else
            {
                paneControl = new PaneControl();
                paneControl.treeView1.Nodes.AddRange(GetTreeNodes());
                paneControl.treeView1.AfterSelect += TreeView1_AfterSelect;
                var taskPane = CustomTaskPanes.Add(paneControl, "Annotation Task Pane (activate)");
                taskPane.Visible = true;

                WindowTaskPanes.Add(activeWindow, taskPane);
            }

            //select current tree node
            if (!string.IsNullOrEmpty(CurrentPath))
            {
                var path = CurrentPath.Split('\\').ToList();
                TreeNodeCollection nodeCollection = paneControl.treeView1.Nodes;
                CurrentNode = searchPath(nodeCollection, path, 0);
                SetTreeNodeColors();
            }

            //refresh list of active task panes
            Dictionary<Word.Window, CustomTaskPane> tempTaskPains = new Dictionary<Word.Window, CustomTaskPane>();
            foreach (Word.Window window in Application.Windows)
            {
                tempTaskPains.Add(window, WindowTaskPanes[window]);
                WindowTaskPanes.Remove(window);
            }
            //clear orphan task panes
            foreach (CustomTaskPane pane in WindowTaskPanes.Values)
            {
                CustomTaskPanes.Remove(pane);
            }
            WindowTaskPanes = tempTaskPains;
        }

        private static TreeNode searchPath(TreeNodeCollection nodes, List<string> path, int depth)
        {
            string key = path[depth++];
            foreach (TreeNode node in nodes)
            {
                if (node.Text == key)
                {
                    if (depth >= path.Count)
                    {
                        return node;
                    }
                    else
                        return searchPath(node.Nodes, path, depth);

                }
            }
            return null;
        }

        private void DeactivateDocumentWindow(Word.Document Doc, Word.Window Wn)
        {
            RestoreTreeNodeColors();
            CurrentNode = null;
        }

        private void RestoreTreeNodeColors()
        {
            if (CurrentNode == null || CurrentTag == null) return;

            CurrentNode.BackColor = TagColors[CurrentTag];
            CurrentNode.ForeColor = Utilities.Contrast(TagColors[CurrentTag]);
        }

        private void SetTreeNodeColors()
        {
            if (CurrentNode == null) return;

            CurrentNode.BackColor = Color.DarkRed;
            CurrentNode.ForeColor = Color.White;
        }

        private void TreeView1_AfterSelect(object sender, TreeViewEventArgs e)
        {
            //don't handle automatic selection 
            if (e.Node == null || e.Node.Tag == null || e.Action == TreeViewAction.Unknown) return;

            RestoreTreeNodeColors();
            CurrentNode = e.Node;
            CurrentTag = e.Node.Tag as string;
            CurrentName = e.Node.Text;
            CurrentPath = e.Node.FullPath;
            WrapContent();

            //enable repeated use of the node for many subsequent ranges
            (sender as TreeView).SelectedNode = null;
            SetTreeNodeColors();
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            InterceptKeys.ReleaseHook();
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
