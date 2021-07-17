using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;

namespace Hjalte.OccurrenceBundler
{
    /// <summary>
    /// Interaction logic for SettingsForm.xaml
    /// </summary>
    public partial class SettingsForm : Window
    {
        private Settings settings;
        public SettingsForm()
        {
            InitializeComponent();

            settings = Globals.Settings;

            TbTemplateText.Text = settings.folderNameTemplate;
            CbBundleByFileName.IsChecked = settings.bundleByFileName;
            CbBundleByProperty.IsChecked = settings.bundleByProperty;
            CbUpdateFileName.IsChecked = settings.updateFolderNamesAutomaticaly;
            setCbPropertyiesValues();
        }

        private void save()
        {
            settings.folderNameTemplate = TbTemplateText.Text;
            settings.bundleByFileName = (bool)CbBundleByFileName.IsChecked;
            settings.bundleByProperty = (bool)CbBundleByProperty.IsChecked;
            settings.updateFolderNamesAutomaticaly = (bool)CbUpdateFileName.IsChecked;

            PossibleProperty newProp = (PossibleProperty)CbPropertyiesValues.SelectedItem;
            settings.iPropertySetName = newProp.setName;
            settings.iPropertyName = newProp.propName;
            SettingSerializer.save(settings);
        }

        private void setCbPropertyiesValues()
        {
            setCbPropertyiesValue(
                Globals.Settings.iPropertySetName,
                Globals.Settings.iPropertyName);
            CbPropertyiesValues.SelectedIndex = 0;

            string summary = PropertySetName.SUMMARY;
            setCbPropertyiesValue(summary, "Author");
            setCbPropertyiesValue(summary, "Comments");
            setCbPropertyiesValue(summary, "Keywords");
            setCbPropertyiesValue(summary, "Last Saved By");
            setCbPropertyiesValue(summary, "Revision Number");
            setCbPropertyiesValue(summary, "Subject");
            setCbPropertyiesValue(summary, "Title");

            string docSummary = PropertySetName.DOCUMENTSUMMARY;
            setCbPropertyiesValue(docSummary, "Category");
            setCbPropertyiesValue(docSummary, "Company");
            setCbPropertyiesValue(docSummary, "Manager");

            string tracking = PropertySetName.DESIGNTRACKING;
            setCbPropertyiesValue(tracking, "Authority");
            setCbPropertyiesValue(tracking, "Categories");
            setCbPropertyiesValue(tracking, "Checked By");
            setCbPropertyiesValue(tracking, "Cost Center");
            setCbPropertyiesValue(tracking, "Description");
            setCbPropertyiesValue(tracking, "Description");
            setCbPropertyiesValue(tracking, "Designer");
            setCbPropertyiesValue(tracking, "Document SubType");
            setCbPropertyiesValue(tracking, "Document SubType Name");
            setCbPropertyiesValue(tracking, "Engineer");
            setCbPropertyiesValue(tracking, "Engr Approved By");
            setCbPropertyiesValue(tracking, "External Property Revision Id");
            setCbPropertyiesValue(tracking, "Manufacturer");
            setCbPropertyiesValue(tracking, "Material");
            setCbPropertyiesValue(tracking, "Mfg Approved By");
            setCbPropertyiesValue(tracking, "Part Number");
            setCbPropertyiesValue(tracking, "Part Property Revision Id");
            setCbPropertyiesValue(tracking, "Project");
            setCbPropertyiesValue(tracking, "Size Designation");
            setCbPropertyiesValue(tracking, "Standard");
            setCbPropertyiesValue(tracking, "Standard Revision");
            setCbPropertyiesValue(tracking, "Standards Organization");
            setCbPropertyiesValue(tracking, "Stock Number");
            setCbPropertyiesValue(tracking, "Template Row");
            setCbPropertyiesValue(tracking, "User Status");
            setCbPropertyiesValue(tracking, "Vendor");
            setCbPropertyiesValue(tracking, "Weld Material");

            string custom = PropertySetName.CUSTOM;
            List<string> customNames = new List<string>();
            foreach (Inventor.Document refDoc in Globals.inventor.Documents)
            {
                foreach (Inventor.Property prop in refDoc.PropertySets[custom])
                {
                    if (!(prop.Value is string)) continue;
                    string propName = prop.Name;
                    if (customNames.Contains(propName) == false)
                    {
                        customNames.Add(propName);
                        setCbPropertyiesValue(custom, propName);
                    }
                }
            }
        }
        private void setCbPropertyiesValue(string setName, string propName)
        {
            CbPropertyiesValues.Items.Add(new PossibleProperty(setName, propName));
        }

        private void button1_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        private void button_Click(object sender, RoutedEventArgs e)
        {
            save();
            this.Close();
        }

        private void TbTemplateText_KeyUp(object sender, KeyEventArgs e)
        {
            BtnOk.IsEnabled = true;
            if (TbTemplateText.Text.Contains("{1}") == false)
            {
                BtnOk.IsEnabled = false;
            }
        }


    }

    public struct PossibleProperty
    {
        public PossibleProperty(string setName, string propName)
        {
            this.setName = setName;
            this.propName = propName;
        }
        public string setName { get; set; }
        public string propName { get; set; }

        public override string ToString()
        {
            return $"{propName} ({setName})";
        }
    }
}
