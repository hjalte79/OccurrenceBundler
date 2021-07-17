using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace Hjalte.OccurrenceBundler
{
    public static class Globals
    {
        public static string MODELBROWSERNAME = "AmBrowserArrangement";
        public static Inventor.Application inventor;
        public static Settings Settings;
        public static OccurrenceBundlerByFileName bundlerByFileName = null;
        public static OccurrenceBundlerByProperty bundlerByProperty = null;
        
        public  static string GUID
        {
            get
            {
                GuidAttribute addInCLSID = (GuidAttribute)GuidAttribute.GetCustomAttribute(
                typeof(StandardAddInServer),
                typeof(GuidAttribute));
                return "{" + addInCLSID.Value + "}";
            }
            
        }
    }

    public static class PropertySetName
    {
        public const string DOCUMENTSUMMARY = "Inventor Document Summary Information";
        public const string SUMMARY = "Inventor Summary Information";
        public const string DESIGNTRACKING = "Design Tracking Properties";
        public const string CUSTOM = "Inventor User Defined Properties";
    }
}
