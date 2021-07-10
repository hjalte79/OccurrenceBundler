using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Inventor;

namespace Hjalte.OccurrenceBundler
{
    public class OccurrenceBundlerByFileName
    {
        private ButtonDefinition bundleButtonDefinition;
        private Application inventor;
        private const string MODELBROWSERNAME = "AmBrowserArrangement";
        

        public OccurrenceBundlerByFileName(Application inventor, string GUID)
        {
            this.inventor = inventor;
            
            bundleButtonDefinition = inventor.CommandManager.ControlDefinitions.AddButtonDefinition(
                    "Bundle same files in folder.", "Hjalte.OccurrenceBundler.handler",
                    CommandTypesEnum.kSchemaChangeCmdType, GUID,
                    "Bundle same files in folder.", "Bundle same files in folder.");
            bundleButtonDefinition.OnExecute += onClickBundle;
        }
        public void OnContextMenu(SelectionDeviceEnum SelectionDevice, NameValueMap AdditionalInfo, CommandBar CommandBar)
        {

            Document oDoc = inventor.ActiveDocument;
            if (!(oDoc is AssemblyDocument)) return;
            AssemblyDocument doc = (AssemblyDocument)oDoc;


            BrowserNode selectedNode = BrowserManipulator.getSelectedNode(doc.BrowserPanes[MODELBROWSERNAME].TopNode);

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
            Document oDoc = inventor.ActiveDocument;
            if (!(oDoc is AssemblyDocument)) return;

            AssemblyDocument doc = (AssemblyDocument)oDoc;
            BrowserPane browserPane = doc.BrowserPanes[MODELBROWSERNAME];
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
                folder.Name = nodesWithFileName.Count + "x " + System.IO.Path.GetFileNameWithoutExtension(fileToBundle);
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
                    folder.Name = nodesWithFileName.Count + "x " + System.IO.Path.GetFileNameWithoutExtension(fileToBundle);
                }
            }            
        }
    }
}
