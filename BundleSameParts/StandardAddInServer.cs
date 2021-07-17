////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// Copyright (c) Hjälte. All rights reserved 
// Written by Jelte de Jong - 2019
//
// Permission to use, copy, modify, and distribute this software in
// object code form for any purpose and without fee is hereby granted, 
// provided that the above copyright notice appears in all copies and 
// that both that copyright notice and the limited warranty and
// restricted rights notice below appear in all supporting 
// documentation.
//
// HJALTE PROVIDES THIS PROGRAM "AS IS" AND WITH ALL FAULTS. 
// HJALTE SPECIFICALLY DISCLAIMS ANY IMPLIED WARRANTY OF
// MERCHANTABILITY OR FITNESS FOR A PARTICULAR USE.  
// DOES NOT WARRANT THAT THE OPERATION OF THE PROGRAM WILL BE
// UNINTERRUPTED OR ERROR FREE.
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.InteropServices;
using Inventor;

namespace Hjalte.OccurrenceBundler
{
    [GuidAttribute("1cecf0ae-ce16-4473-9b8b-320842f927f3"), ComVisible(true)]
    public class StandardAddInServer : ApplicationAddInServer
    {
        //https://forums.autodesk.com/t5/inventor-ideas/bundle-multiple-instances-of-same-part-in-assembly-tree/idi-p/5643016

        private ButtonDefinition settingsButton;
        private FolderReNamer folderReNamer;

        public void Activate(ApplicationAddInSite addInSiteObject, bool firstTime)
        {
            try
            {
                Globals.inventor = addInSiteObject.Application;
                Globals.Settings = SettingSerializer.load();

                Globals.bundlerByFileName = new OccurrenceBundlerByFileName();
                Globals.bundlerByProperty = new OccurrenceBundlerByProperty();

                createSettingsButton();

                folderReNamer = new FolderReNamer();
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show($"Exception was thrown while loading addon. Message: {ex.Message}");
            }
        }

        private void createSettingsButton()
        {
            ControlDefinitions conDefs = Globals.inventor.CommandManager.ControlDefinitions;
            settingsButton = conDefs.AddButtonDefinition(
                "Bundler Configuration", 
                "Hjalte.Bundler.Settings", 
                CommandTypesEnum.kEditMaskCmdType, 
                Guid.NewGuid().ToString(),
                "Bundler Configuration",
                "Bundler Configuration");
            settingsButton.OnExecute += SettingsButton_OnExecute;

            Ribbon ribbon = Globals.inventor.UserInterfaceManager.Ribbons["ZeroDoc"];
            RibbonTab tab = ribbon.RibbonTabs["id_TabTools"];
            RibbonPanel panel = tab.RibbonPanels["id_PanelP_ToolsOptions"];
            panel.SlideoutControls.AddButton(settingsButton);

            ribbon = Globals.inventor.UserInterfaceManager.Ribbons["Part"];
            tab = ribbon.RibbonTabs["id_TabTools"];
            panel = tab.RibbonPanels["id_PanelP_ToolsOptions"];
            panel.SlideoutControls.AddButton(settingsButton);

            ribbon = Globals.inventor.UserInterfaceManager.Ribbons["Assembly"];
            tab = ribbon.RibbonTabs["id_TabTools"];
            panel = tab.RibbonPanels["id_PanelP_ToolsOptions"];
            panel.SlideoutControls.AddButton(settingsButton);

            ribbon = Globals.inventor.UserInterfaceManager.Ribbons["Drawing"];
            tab = ribbon.RibbonTabs["id_TabTools"];
            panel = tab.RibbonPanels["id_PanelP_ToolsOptions"];
            panel.SlideoutControls.AddButton(settingsButton);



        }
        private void SettingsButton_OnExecute(NameValueMap Context)
        {
            SettingsForm settingsForm = new SettingsForm();
            settingsForm.ShowDialog();
        }

        public void Deactivate()
        {
            // Release objects.
            Marshal.ReleaseComObject(Globals.inventor);
            Globals.inventor = null;

            GC.WaitForPendingFinalizers();
            GC.Collect();
        }
        public void ExecuteCommand(int commandID)
        {
            // Note:this method is now obsolete, you should use the 
            // ControlDefinition functionality for implementing commands.
        }
        public object Automation
        {
            get
            {
                // TODO: Add ApplicationAddInServer.Automation getter implementation
                return null;
            }
        }
    }
}
