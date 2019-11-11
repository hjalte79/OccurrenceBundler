using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Inventor;

namespace Hjalte.OccurrenceBundler
{
    public class OccurrenceBundler
    {
        private ButtonDefinition bundleButtonDefinition;
        private Application inventor;
        private const string MODELBROWSERNAME = "AmBrowserArrangement";

        public OccurrenceBundler(Application inventor, string GUID)
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


            BrowserNode selectedNode = getSelectedNode(doc.BrowserPanes[MODELBROWSERNAME].TopNode);

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

            BrowserNode selectedNode = getSelectedNode(topNode);
            if (selectedNode == null) return;

            string fileToBundle = "";
            if (selectedNode.NativeObject is ComponentOccurrence)
            {
                ComponentOccurrence occ = (ComponentOccurrence)selectedNode.NativeObject;
                fileToBundle = occ.ReferencedFileDescriptor.FullFileName;
                                
                BrowserFolder folder = getFolderForFile(browserPane, fileToBundle);

                List<BrowserNode> nodesWithFileName = getNodesWithFileName(fileToBundle, topNode);
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
                    folder = getFolderForFile(browserPane, fileToBundle);

                    List<BrowserNode> nodesWithFileName = getNodesWithFileName(fileToBundle, topNode);
                    foreach (BrowserNode foundNode in nodesWithFileName)
                    {
                        folder.Add(foundNode);
                    }
                    folder.Name = nodesWithFileName.Count + "x " + System.IO.Path.GetFileNameWithoutExtension(fileToBundle);
                }
            }
            

        }

        private BrowserFolder getFolderForFile(BrowserPane browserPane, string searchedFileName)
        {
            BrowserFolder folder = null;

            foreach (BrowserFolder curentFolder in browserPane.TopNode.BrowserFolders)
            {
                foreach (BrowserNode node in curentFolder.BrowserNode.BrowserNodes)
                {
                    if (node.NativeObject is ComponentOccurrence)
                    {
                        ComponentOccurrence occ = (ComponentOccurrence)node.NativeObject;
                        string currentFileName = occ.ReferencedFileDescriptor.FullFileName;
                        if (!searchedFileName.Equals(currentFileName, StringComparison.InvariantCultureIgnoreCase))
                        {
                            continue;
                        }
                    }
                    else
                    {
                        continue;
                    }
                    folder = curentFolder;
                    break;
                }
            }

            if (folder == null )
            {
                ObjectCollection colNodes = inventor.TransientObjects.CreateObjectCollection();
                folder = browserPane.AddBrowserFolder("temp", colNodes);
            }
            
            return folder;
        }

        private List<BrowserNode> getNodesWithFileName(string fileName, BrowserNode topNode)
        {
            List<BrowserNode> nodes = new List<BrowserNode>();
            foreach (BrowserNode node in topNode.BrowserNodes)
            {
                if (node.NativeObject is ComponentOccurrence)
                {
                    ComponentOccurrence occ = (ComponentOccurrence)node.NativeObject;
                    string foundOccfilename = occ.ReferencedFileDescriptor.FullFileName;
                    if (fileName.Equals(foundOccfilename))
                    {
                        nodes.Add(node);
                    }
                }
                if (node.NativeObject is BrowserFolder)
                {
                    BrowserFolder folder = (BrowserFolder)node.NativeObject;
                    nodes.AddRange(getNodesWithFileName(fileName, folder.BrowserNode));
                }
            }
            return nodes;
        }

        private BrowserNode getSelectedNode(BrowserNode topNode)
        {
            BrowserNode foundNode = null;
            foreach (BrowserNode node in topNode.BrowserNodes)
            {
                if (node.Selected) return node;
                foundNode = getSelectedNode(node);
                if (foundNode != null)
                {
                    return foundNode;
                }
            }
            return foundNode;
        }


    }
}
