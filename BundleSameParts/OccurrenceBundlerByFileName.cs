using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Inventor;

namespace Hjalte.OccurrenceBundler
{
    public class OccurrenceBundlerByFileName : iOccurrenceBundler
    {
        private ButtonDefinition bundleButtonDefinition;        

        public OccurrenceBundlerByFileName()
        {
            string name = "Bundle by file name";
            bundleButtonDefinition = Globals.inventor.CommandManager.ControlDefinitions.AddButtonDefinition(
                    name, "Hjalte.OccurrenceBundler.handler",
                    CommandTypesEnum.kSchemaChangeCmdType, Globals.GUID,
                    name, name);
            bundleButtonDefinition.OnExecute += onClickBundle;

            Globals.inventor.CommandManager.UserInputEvents.OnContextMenu += OnContextMenu;
        }
        public void OnContextMenu(SelectionDeviceEnum SelectionDevice, NameValueMap AdditionalInfo, CommandBar CommandBar)
        {
            if (SelectionDevice != SelectionDeviceEnum.kBrowserSelection) return;
            if (Globals.Settings.bundleByFileName == false) return;
            Document oDoc = Globals.inventor.ActiveDocument;
            if (!(oDoc is AssemblyDocument)) return;
            AssemblyDocument doc = (AssemblyDocument)oDoc;

            if (doc.BrowserPanes[Globals.MODELBROWSERNAME].TopNode.Selected == true) return;
            BrowserNode selectedNode = BrowserManipulator.getSelectedNode(doc.BrowserPanes[Globals.MODELBROWSERNAME].TopNode);

            if (selectedNode == null) return;

            try
            {
                if (selectedNode.NativeObject is ComponentOccurrence)
                {
                    CommandBar.Controls.AddButton(bundleButtonDefinition, 6);
                }
                else if (selectedNode.NativeObject is BrowserFolder)
                {
                    CommandBar.Controls.AddButton(bundleButtonDefinition, 5);
                }
            }
            catch (Exception) { }

        }
        private void onClickBundle(NameValueMap Context)
        {
            Document oDoc = Globals.inventor.ActiveDocument;
            if (!(oDoc is AssemblyDocument)) return;

            AssemblyDocument doc = (AssemblyDocument)oDoc;
            BrowserPane browserPane = doc.BrowserPanes[Globals.MODELBROWSERNAME];
            BrowserNode topNode = browserPane.TopNode;

            BrowserNode selectedNode = BrowserManipulator.getSelectedNode(topNode);
            if (selectedNode == null) return;

            string fileToBundle = "";
            if (selectedNode.NativeObject is ComponentOccurrence)
            {
                ComponentOccurrence occ = (ComponentOccurrence)selectedNode.NativeObject;
                fileToBundle = occ.ReferencedFileDescriptor.FullFileName;
                                
                BrowserFolder folder = BrowserManipulator.getFolderForFile(browserPane, fileToBundle);

                List<BrowserNode> nodesWithFileName = BrowserManipulator.getNodesWithFileName(fileToBundle, topNode);
                foreach (BrowserNode foundNode in nodesWithFileName)
                {
                    folder.Add(foundNode);
                }
                folder.Name = string.Format(
                        Globals.Settings.folderNameTemplate,
                        nodesWithFileName.Count,
                        System.IO.Path.GetFileNameWithoutExtension(fileToBundle));
            }
            else if (selectedNode.NativeObject is BrowserFolder)
            {
                BrowserFolder folder = (BrowserFolder)selectedNode.NativeObject;
                
                if (folder.BrowserNode.BrowserNodes[1].NativeObject is ComponentOccurrence)
                {
                    ComponentOccurrence occ = (ComponentOccurrence)folder.BrowserNode.BrowserNodes[1].NativeObject;
                    fileToBundle = occ.ReferencedFileDescriptor.FullFileName;

                    folder.Delete();
                    folder = BrowserManipulator.getFolderForFile(browserPane, fileToBundle);

                    List<BrowserNode> nodesWithFileName = BrowserManipulator.getNodesWithFileName(fileToBundle, topNode);
                    foreach (BrowserNode foundNode in nodesWithFileName)
                    {
                        folder.Add(foundNode);
                    }
                    folder.Name = string.Format(
                        Globals.Settings.folderNameTemplate,
                        nodesWithFileName.Count,
                        System.IO.Path.GetFileNameWithoutExtension(fileToBundle));
                }
            }            
        }
    }
}
