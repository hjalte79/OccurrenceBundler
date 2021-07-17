using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;

namespace Hjalte.OccurrenceBundler
{
    public static class SettingSerializer
    {
        public static void save(Settings buttonSettingsXml)
        {
            using (var writer = new System.IO.StreamWriter(getSettingsFilePath()))
            {
                var serializer = new XmlSerializer(typeof(Settings));
                serializer.Serialize(writer, buttonSettingsXml);
                writer.Flush();
            }
        }

        public static Settings load()
        {
            using (var stream = System.IO.File.OpenRead(getSettingsFilePath()))
            {
                var serializer = new XmlSerializer(typeof(Settings));
                return serializer.Deserialize(stream) as Settings;
            }
        }

        private static string getSettingsFilePath()
        {
            string codeBase = System.Reflection.Assembly.GetExecutingAssembly().CodeBase;
            UriBuilder uri = new UriBuilder(codeBase);
            string path = Uri.UnescapeDataString(uri.Path);
            path = System.IO.Path.GetDirectoryName(path);

            return System.IO.Path.Combine(path, "Hjalte.OccurrenceBundler.conf.Xml");
        }
    }
}
