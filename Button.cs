using System;
using System.Runtime.InteropServices;
using Inventor;
using System.Text;
using System.Drawing;
using Microsoft.VisualBasic.Compatibility.VB6;
using System.Windows.Forms;

namespace InvAddIn
{                        
    internal abstract class Button
    {
#region Data Members
        private static Inventor.Application m_inventorApplication;
        private ButtonDefinition m_buttonDefinition;
        private ButtonDefinitionSink_OnExecuteEventHandler ButtonDefinition_OnExecuteEventDelegate;
#endregion
#region "Properties"
        public static Inventor.Application InventorApplication
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
        public Inventor.ButtonDefinition _ButtonDefinition
        {
            get
            {                                       
                return m_buttonDefinition;
            }
        }
#endregion
#region "Methods"
        public Button(string DisplayName, string internalName, CommandTypesEnum commandType, string clientId, string description, string tooltip, Icon standarticon, Icon largeIcon, ButtonDisplayEnum buttonDisplayType)
        {
            try
            {
                stdole.IPictureDisp standartIconIPictureDisp;
                standartIconIPictureDisp = (stdole.IPictureDisp)Support.IconToIPicture(standarticon);
                stdole.IPictureDisp largeIconIPictureDisp;
                largeIconIPictureDisp = (stdole.IPictureDisp)Support.IconToIPicture(largeIcon);
                m_buttonDefinition = m_inventorApplication.CommandManager.ControlDefinitions.AddButtonDefinition(DisplayName,
                    internalName, commandType, clientId, description, tooltip, standartIconIPictureDisp, largeIconIPictureDisp, buttonDisplayType);
                m_buttonDefinition.Enabled = true; 
                ButtonDefinition_OnExecuteEventDelegate = new ButtonDefinitionSink_OnExecuteEventHandler(ButtonDefinition_OnExecute);
                m_buttonDefinition.OnExecute += ButtonDefinition_OnExecuteEventDelegate;
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        public Button(string displayName, string internalName, CommandTypesEnum commandType, string clientId, string description, string tooltip, ButtonDisplayEnum buttonDisplayType)
        {
            try
            {
                m_buttonDefinition = m_inventorApplication.CommandManager.ControlDefinitions.AddButtonDefinition(displayName,internalName,commandType,clientId,description,tooltip,Type.Missing,Type.Missing,buttonDisplayType);
                m_buttonDefinition.Enabled = true;
                ButtonDefinition_OnExecuteEventDelegate = new ButtonDefinitionSink_OnExecuteEventHandler(ButtonDefinition_OnExecute);
                m_buttonDefinition.OnExecute += ButtonDefinition_OnExecuteEventDelegate;
            }
            catch (System.Exception ex)
            {
	            MessageBox.Show(ex.ToString());
            }
                    }

        abstract protected void ButtonDefinition_OnExecute(NameValueMap context);

#endregion
    }

    internal abstract class InvComboBox
    {
        private static Inventor.Application m_inventorApplication;
        private ComboBoxDefinition m_ComboBoxDef;
        private ComboBoxDefinitionSink_OnSelectEventHandler m_cbEhDelegate;
        public static Inventor.Application InventorApplication
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
        public Inventor.ComboBoxDefinition _ComboBoxDef
        {
            get
            {
                return m_ComboBoxDef;
            }
        }
        public InvComboBox(string DisplayName, string InternalName, CommandTypesEnum Classification, int DropDownWidth,
            object ClientId, object StandardIcon,object LargeIcon, string DescriptionText = "", string ToolTipText = "", 
             ButtonDisplayEnum ButtonDisplay = ButtonDisplayEnum.kDisplayTextInLearningMode)
        {
            m_ComboBoxDef = m_inventorApplication.CommandManager.ControlDefinitions.AddComboBoxDefinition(DisplayName, InternalName, Classification, DropDownWidth, ClientId,
                DescriptionText, ToolTipText, StandardIcon, LargeIcon);
            m_ComboBoxDef.Enabled = true;
            m_cbEhDelegate = new ComboBoxDefinitionSink_OnSelectEventHandler(ComboBoxDefinition_OnSelect);
            m_ComboBoxDef.OnSelect += m_cbEhDelegate;
        }

        abstract protected void ComboBoxDefinition_OnSelect(NameValueMap nvm);
    }
}



