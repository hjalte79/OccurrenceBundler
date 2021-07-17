using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Inventor;

namespace Hjalte.OccurrenceBundler
{
    public class OccurrenceBundlerByProperty : iOccurrenceBundler
    {
        private ButtonDefinition bundleButtonDefinition;
        private ButtonDefinition bundleButtonDefinitionSingle;
        public OccurrenceBundlerByProperty()
        {
            string buttonName = $"Bundle by iProperty '{Globals.Settings.iPropertyName}'";
            bundleButtonDefinition = Globals.inventor.CommandManager.ControlDefinitions.AddButtonDefinition(
                    buttonName, "Hjalte.OccurrenceBundlerByFileName.handler",
                    CommandTypesEnum.kSchemaChangeCmdType, Globals.GUID,
                    buttonName, buttonName);
            bundleButtonDefinition.OnExecute += onClickBundle;


            string buttonNameSingle = $"Bundle by iProperty '{Globals.Settings.iPropertyName}'";
            bundleButtonDefinitionSingle = Globals.inventor.CommandManager.ControlDefinitions.AddButtonDefinition(
                    buttonNameSingle, "Hjalte.OccurrenceBundlerByFileName.handler.single",
                    CommandTypesEnum.kSchemaChangeCmdType, Globals.GUID,
                    buttonNameSingle, buttonNameSingle);
            bundleButtonDefinitionSingle.OnExecute += onClickBundleSingle;

            Globals.inventor.CommandManager.UserInputEvents.OnContextMenu += OnContextMenu;
        }



        private void onClickBundle(NameValueMap Context)
        {
            Document oDoc = Globals.inventor.ActiveDocument;
            if (!(oDoc is AssemblyDocument)) return;

            AssemblyDocument doc = (AssemblyDocument)oDoc;
            BrowserPane browserPane = doc.BrowserPanes[Globals.MODELBROWSERNAME];
            BrowserNode topNode = browserPane.TopNode;
            List<BrowserFolder> changedFolders = new List<BrowserFolder>();

            foreach (BrowserNode node in topNode.BrowserNodes)
            {
                if (!(node.NativeObject is ComponentOccurrence)) continue;
                try
                {
                    string propertyValue = getPropertyValue(node);
                    if (string.IsNullOrEmpty(propertyValue)) continue;

                    BrowserFolder folder = getFolder(browserPane, propertyValue);

                    folder.Add(node);
                    changedFolders.Add(folder);
                    
                }
                catch (Exception) { }              
                
            }

            foreach (BrowserFolder folder in changedFolders)
            {
                string propertyValue = getPropertyValue(folder.BrowserNode.BrowserNodes[1]);

                int n = folder.BrowserNode.BrowserNodes.Count;
                
                folder.Name = string.Format(Globals.Settings.folderNameTemplate, n, propertyValue);
            }
        }

        private BrowserFolder getFolder(BrowserPane browserPane, string name)
        {
            BrowserFolder folder = null;
            BrowserNode topNode = browserPane.TopNode;

            IEnumerable<BrowserFolder> folders = topNode.BrowserFolders.Cast<BrowserFolder>().
                        Where(f => f.Name.Contains(name));
            
            string patternOrg = Globals.Settings.folderNameTemplate
                .Replace(@"(", @"\(")
                .Replace(@")", @"\)")
                .Replace(@"*", @"\*")
                .Replace(@".", @"\.)");

            string pattern = string.Format(
                patternOrg,
                @"(?<number>\d*)",
                @"(?<name>.*)");
            Regex rg = new Regex(pattern);

            
            foreach (BrowserFolder cFolder in folders)
            {
                Match match = rg.Match(cFolder.Name);
                if (match.Success)
                {
                    string matchName = match.Groups["name"].Value;
                    if (name.Equals(matchName))
                    {
                        folder = cFolder;
                    }                    
                }
            }

            if (folder == null) 
            {
                ObjectCollection colNodes = Globals.inventor.TransientObjects.CreateObjectCollection();
                folder = browserPane.AddBrowserFolder(name, colNodes);
                folder.Name = string.Format(Globals.Settings.folderNameTemplate, "1", name);
            }
            return folder;
        }

        private string getPropertyValue(BrowserNode node)
        {
            if (node.NativeObject is BrowserFolder && node.BrowserNodes.Count > 0)
            {
                return getPropertyValue(node.BrowserNodes[1]);
            }

            if (!(node.NativeObject is ComponentOccurrence)) return null;

            try
            {
                ComponentOccurrence occ = (ComponentOccurrence)node.NativeObject;
                Document refDoc = (Document)occ.ReferencedDocumentDescriptor.ReferencedDocument;

                PropertySet set = refDoc.PropertySets[Globals.Settings.iPropertySetName];
                Property prop = set[Globals.Settings.iPropertyName];
                return (string)prop.Value;
            }
            catch (Exception)
            {
                return null;
            }


                       
        }



        public void OnContextMenu(SelectionDeviceEnum SelectionDevice, NameValueMap AdditionalInfo, CommandBar CommandBar)
        {
            if (SelectionDevice != SelectionDeviceEnum.kBrowserSelection) return;
            if (Globals.Settings.bundleByProperty == false) return;
            Document oDoc = Globals.inventor.ActiveDocument;
            if (!(oDoc is AssemblyDocument)) return;
            AssemblyDocument doc = (AssemblyDocument)oDoc;

            try
            {
                if (doc.BrowserPanes[Globals.MODELBROWSERNAME].TopNode.Selected)
                {
                    CommandBarControl controll = CommandBar.Controls.AddButton(bundleButtonDefinition, CommandBar.Controls.Count);
                    controll.GroupBegins = true;
                }
                else
                {
                    BrowserNode selectedNode = BrowserManipulator.getSelectedNode(doc.BrowserPanes[Globals.MODELBROWSERNAME].TopNode);
                    if (selectedNode == null) return;

                    CommandBarControl control = null;
                    string propertyValue = getPropertyValue(selectedNode);
                    if (selectedNode.NativeObject is ComponentOccurrence)
                    {
                        control = CommandBar.Controls.AddButton(bundleButtonDefinitionSingle, 6);                        
                    }
                    else if (selectedNode.NativeObject is BrowserFolder)
                    {
                        control = CommandBar.Controls.AddButton(bundleButtonDefinitionSingle, 5);
                    }
                    else
                    {
                        return;
                    }
                    control.ControlDefinition.Enabled = true;                    
                    if (string.IsNullOrEmpty(propertyValue))
                    {
                        control.ControlDefinition.Enabled = false;
                    }
                    
                }

            }
            catch (Exception) { }
        }

        private void onClickBundleSingle(NameValueMap Context)
        {
            Document oDoc = Globals.inventor.ActiveDocument;
            if (!(oDoc is AssemblyDocument)) return;
            AssemblyDocument doc = (AssemblyDocument)oDoc;
            BrowserPane browserPane = doc.BrowserPanes[Globals.MODELBROWSERNAME];
            BrowserNode selectedNode = BrowserManipulator.getSelectedNode(browserPane.TopNode);

            string propertyValue = getPropertyValue(selectedNode);
            if (string.IsNullOrEmpty(propertyValue)) return;


            BrowserFolder folder = getFolder(browserPane, propertyValue);
            foreach (BrowserNode node in browserPane.TopNode.BrowserNodes)
            {
                try
                {
                    if (node.NativeObject is ComponentOccurrence)
                    {
                    
                            string propertyValueSletcedNode = getPropertyValue(node);
                            if (!propertyValue.Equals(propertyValueSletcedNode)) continue;

                            folder.Add(node);
                    
                    }
                    if (node.NativeObject is BrowserFolder)
                    {
                        foreach (BrowserNode childNode in node.BrowserNodes)
                        {
                            if (childNode.NativeObject is ComponentOccurrence)
                            {
                                string propertyValueSletcedNode = getPropertyValue(childNode);
                                if (!propertyValue.Equals(propertyValueSletcedNode)) continue;

                                folder.Add(childNode);
                            }
                        }
                    }
                }
                catch (Exception) { }
            }

            int n = folder.BrowserNode.BrowserNodes.Count;
            folder.Name = string.Format(Globals.Settings.folderNameTemplate, n, propertyValue);
            
        }

    }
}
