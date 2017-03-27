using System;
using System.Windows.Forms;
using System.Drawing;
using Inventor;
using System.Linq;

namespace InvAddIn
{
    internal class Spec : Button
    {
        public static Table m_Tbl;
        public Inventor.Document pDoc { get; set; }
        public static System.Collections.Generic.List<string> vars { get; set; }
        public static Table getTable
        {
            get
            {
                return m_Tbl;
            }
        }

#region "Methods"
        public Spec(string displayName, string internalName, CommandTypesEnum commandType, string clientId, string description, string tooltip, Icon standardIcon, Icon largeIcon, ButtonDisplayEnum buttonDisplayType)
            : base(displayName, internalName, commandType, clientId, description, tooltip, standardIcon, largeIcon, buttonDisplayType)
        {

        }
        public Spec(string displayName, string internalName, CommandTypesEnum commandType, string clientId, string description, string tooltip, ButtonDisplayEnum buttonDisplayType)
			: base(displayName, internalName, commandType, clientId, description, tooltip, buttonDisplayType)
		{
			
		}

        protected override void ButtonDefinition_OnExecute(NameValueMap context)
        {
            Macros.StandardAddInServer.forms.Add(m_Tbl);
            if (Macros.StandardAddInServer.activeteForm()) System.Windows.Forms.Application.Run(m_Tbl = new Table(InventorApplication.ActiveDocument, InventorApplication));
        }
        
#endregion
    }                        
}
