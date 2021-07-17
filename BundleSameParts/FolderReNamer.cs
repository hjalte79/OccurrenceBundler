using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Inventor;

namespace Hjalte.OccurrenceBundler
{
    public class FolderReNamer
    {
        public ApplicationEvents appEvents;
        public FolderReNamer()
        {
            if (Globals.Settings.updateFolderNamesAutomaticaly == false) return;
            appEvents = Globals.inventor.ApplicationEvents;
            appEvents.OnDocumentChange += AppEvents_OnDocumentChange;
        }
        private bool isUpdating = false;
        private void AppEvents_OnDocumentChange(_Document DocumentObject, EventTimingEnum BeforeOrAfter, CommandTypesEnum ReasonsForChange, NameValueMap Context, out HandlingCodeEnum HandlingCode)
        {
            HandlingCode = HandlingCodeEnum.kEventHandled;
            if (isUpdating) return;
            if (BeforeOrAfter != EventTimingEnum.kAfter) return;
            if (ReasonsForChange != CommandTypesEnum.kShapeEditCmdType) return;
            if (DocumentObject.DocumentType != DocumentTypeEnum.kAssemblyDocumentObject) return;
            if (Globals.Settings.folderNameTemplate.Contains("{0}") == false) return;
            try
            {
                string command = (string)Context.Item[1];
                if (
                    (command.Equals("AssemblyRestructure") == false) &&
                    (command.Equals("CompositeChange") == false) &&
                    (command.Equals("Reorder") == false) &&
                    (command.Equals("Restructure Assembly") == false) &&
                    (command.Equals("Delete Selections") == false) &&
                    (command.Equals("Drag Over Folder") == false))
                    return;

                isUpdating = true;
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

                AssemblyDocument doc = (AssemblyDocument)DocumentObject;
                BrowserPane browserPane = doc.BrowserPanes[Globals.MODELBROWSERNAME];

                foreach (BrowserNode node in browserPane.TopNode.BrowserNodes)
                {
                    if (!(node.NativeObject is BrowserFolder)) continue;
                    BrowserFolder folder = (BrowserFolder)node.NativeObject;
                    Match match = rg.Match(folder.Name);
                    if (match.Success)
                    {
                        // string number = match.Groups["number"].Value;
                        string name = match.Groups["name"].Value;

                        int n = folder.BrowserNode.BrowserNodes.Count;
                        folder.Name = string.Format(Globals.Settings.folderNameTemplate, n, name);
                    }

                }

            }
            catch (Exception) { }
            isUpdating = false;
        }
    }
}
