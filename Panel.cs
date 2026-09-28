using System;
using System.Runtime.InteropServices;
using Inventor;

namespace InvAddIn
{
    class Panel
    {
        
        #region DataMembers
            private static Inventor.Application m_inventorApplication;
            public static Inventor.CommandControl m_commandControl;
            private Inventor.CommandControls m_commandControls;
            private readonly UserInterfaceManager uim = I.app.UserInterfaceManager;
        #endregion
        #region "Properties"
        public Inventor.Application InventorApplication
            {
                set
                {
                    m_inventorApplication = value;
                }
                get
                {
                    return m_inventorApplication;
                }
            }
            public Inventor.CommandControl CmdControl 
            {
                get
                {
                    return m_commandControl;
                }
            }
        #endregion
        #region "Methods"
            public Panel( Inventor.Application invApp, ButtonDefinition btnDef ,string ribbonName, string ribbonTabName, string displayName, string intName, string m_guid)
            {
                Ribbons ribbons;
                ribbons = uim.Ribbons;
                Ribbon ribbon;
                ribbon = ribbons[ribbonName];
                RibbonTabs ribbonTabs;
                ribbonTabs = ribbon.RibbonTabs;
                RibbonTab ribbonTab;
                ribbonTab = ribbonTabs[ribbonTabName];
                RibbonPanels ribbonPanels;
                ribbonPanels = ribbonTab.RibbonPanels;
                RibbonPanel ribbonPanel;
                ribbonPanel = ribbonPanels.Add(displayName,intName, m_guid,"",false);
                //CommandControls commandControls;
                m_commandControls = ribbonPanel.CommandControls;
                m_commandControl = m_commandControls.AddButton(btnDef);
            }
        public Panel(string ribbonName, string ribbonTabName, string displayName, string intName, string m_guid)
        {
            RibbonPanel rp = uim.Ribbons[ribbonName].RibbonTabs[ribbonTabName].RibbonPanels.Add(displayName, intName, m_guid);
            m_commandControls = rp.CommandControls;
        }
            public Panel(string ribbonName, string ribbonTabName, string ribbonPanelName)
            {
                m_commandControls = uim.Ribbons[ribbonName].RibbonTabs[ribbonTabName].RibbonPanels[ribbonPanelName].CommandControls;
            }
        public void addControl(ButtonDefinition btnDef)
        {
            m_commandControl = m_commandControls.AddButton(btnDef);
        }
        public void addControl(ComboBoxDefinition cbDef, string name = "")
        {
            if (name == "")
                m_commandControl = m_commandControls.AddComboBox(cbDef);
            else m_commandControl = m_commandControls.AddComboBox(cbDef, name);
        }
        #endregion
    }
}
