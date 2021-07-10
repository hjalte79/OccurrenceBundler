using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Inventor;

namespace Hjalte.OccurrenceBundler
{
    public class OccurrenceBundlerByProperty
    {
        private ButtonDefinition bundleButtonDefinition;
        private Application inventor;
        private const string MODELBROWSERNAME = "AmBrowserArrangement";
        private const string CUSTOMPROPERTIESSETNAME = "Inventor User Defined Properties";
        private const string IPROPERTYNAME = "Assembly_Folder";

        public OccurrenceBundlerByProperty(Application inventor, string GUID)
        {
            this.inventor = inventor;
            string buttonName = "Bundle with same value for iProperty 'Assembly_Folder'";
            bundleButtonDefinition = inventor.CommandManager.ControlDefinitions.AddButtonDefinition(
                    buttonName, "Hjalte.OccurrenceBundlerByFileName.handler",
                    CommandTypesEnum.kSchemaChangeCmdType, GUID,
                    buttonName, buttonName);
            bundleButtonDefinition.OnExecute += onClickBundle;
        }

        private void onClickBundle(NameValueMap Context)
        {
            Document oDoc = inventor.ActiveDocument;
            if (!(oDoc is AssemblyDocument)) return;

            AssemblyDocument doc = (AssemblyDocument)oDoc;
            BrowserPane browserPane = doc.BrowserPanes[MODELBROWSERNAME];
            BrowserNode topNode = browserPane.TopNode;

            foreach (BrowserNode node in topNode.BrowserNodes)
            {
                if (node.NativeObject is ComponentOccurrence) continue;
                try
                {
                    string propertyValue = getPropertyValue(node);
                    if (string.IsNullOrEmpty(propertyValue)) continue;

                    BrowserFolder folder = topNode.BrowserFolders.Cast<BrowserFolder>().
                        Where(f => f.Name.Equals(propertyValue, StringComparison.InvariantCultureIgnoreCase)).FirstOrDefault();
                        
                    if (folder == null)
                    {
                        ObjectCollection colNodes = inventor.TransientObjects.CreateObjectCollection();
                        folder = browserPane.AddBrowserFolder(propertyValue, colNodes);
                    }

                    folder.Add(node);
                }
                catch (Exception) { }              
                
            }
        }

        private string getPropertyValue(BrowserNode node)
        {
            if (node.NativeObject is ComponentOccurrence) return null;

            try
            {
                ComponentOccurrence occ = (ComponentOccurrence)node.NativeObject;
                Document refDoc = (Document)occ.ReferencedDocumentDescriptor.ReferencedDocument;

                PropertySet set = refDoc.PropertySets[CUSTOMPROPERTIESSETNAME];
                Property prop = set[IPROPERTYNAME];
                return (string)prop.Value;
            }
            catch (Exception)
            {
                return null;
            }            
        }



        public void OnContextMenu(SelectionDeviceEnum SelectionDevice, NameValueMap AdditionalInfo, CommandBar CommandBar)
        {
            Document oDoc = inventor.ActiveDocument;
            if (!(oDoc is AssemblyDocument)) return;
            AssemblyDocument doc = (AssemblyDocument)oDoc;

            if (doc.BrowserPanes[MODELBROWSERNAME].TopNode.Selected)
            {
                CommandBar.Controls.AddButton(bundleButtonDefinition);
            }


        }
    }
}
