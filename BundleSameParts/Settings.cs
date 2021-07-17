using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Inventor;

namespace Hjalte.OccurrenceBundler
{
    public class Settings
    {

        public bool updateFolderNamesAutomaticaly { get; set; } = true;
        public bool bundleByFileName { get; set; } = true;
        public bool bundleByProperty { get; set; } = true;
        public string folderNameTemplate { get; set; } = "{0}x {1}";
        public string iPropertySetName { get; set; } = PropertySetName.CUSTOM;
        public string iPropertyName { get; set; } = "Assembly_Folder";

        public void loadBundlers()
        {
            

            
        }

    }

}
