using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Inventor;

namespace Hjalte.OccurrenceBundler
{
    public static class BrowserManipulator
    {
        private const string MODELBROWSERNAME = "AmBrowserArrangement";
        

        public static BrowserFolder getFolderForFile(BrowserPane browserPane, string searchedFileName)
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

            if (folder == null)
            {
                ObjectCollection colNodes = Globals.inventor.TransientObjects.CreateObjectCollection();
                folder = browserPane.AddBrowserFolder("temp", colNodes);
            }

            return folder;
        }

        public static  List<BrowserNode> getNodesWithFileName(string fileName, BrowserNode topNode)
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

        public static BrowserNode getSelectedNode(BrowserNode topNode)
        {
            BrowserNode foundNode = null;
            foreach (BrowserNode node in topNode.BrowserNodes)
            {
                // For some unknow reason interface checking doesn't work here...
                // there for the function nodeNativeObjectIs() is used! 
                if (node.NativeObject == null) continue;
                //if (node.NativeObject is ComponentOccurrences )
                if (nodeNativeObjectIs(node, ObjectTypeEnum.kComponentOccurrenceObject))
                {
                    if (node.Selected) return node;
                }
                //f (node.NativeObject is BrowserFolder)
                if (nodeNativeObjectIs(node, ObjectTypeEnum.kBrowserFolderObject))
                {
                    foundNode = getSelectedNode(node);
                    if (foundNode != null)
                    {
                        return foundNode;
                    }
                }
            }
            return foundNode;
        }

        public static bool nodeNativeObjectIs(BrowserNode node , ObjectTypeEnum inventorType)
        {
            try
            {
                dynamic nativeObject = node.NativeObject;
                ObjectTypeEnum inventorTypeOrg = (ObjectTypeEnum)nativeObject.Type;
                if (inventorType == inventorTypeOrg) return true;
            }
            catch (Exception ex) { }

            return false;
        }
    }
}
