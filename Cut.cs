using System;
using System.Windows.Forms;
using System.Drawing;
using Inventor;
using System.Linq;
using u = InvDoc.u;
using InvDoc;
using System.Collections.Generic;

namespace InvAddIn
{
    internal class CutButton: Button
    {
        public static CutOp m_Cut;
        public Inventor.Document pDoc { get; set; }
        public static CutOp getBreak
        {
            get
            {
                return m_Cut;
            }
        }

        #region "Methods"
        public CutButton(string displayName, string internalName, CommandTypesEnum commandType, string clientId, string description, string tooltip, Icon standardIcon, Icon largeIcon, ButtonDisplayEnum buttonDisplayType)
            : base(displayName, internalName, commandType, clientId, description, tooltip, standardIcon, largeIcon, buttonDisplayType)
        {

        }
        public CutButton(string displayName, string internalName, CommandTypesEnum commandType, string clientId, string description, string tooltip, ButtonDisplayEnum buttonDisplayType)
            : base(displayName, internalName, commandType, clientId, description, tooltip, buttonDisplayType)
        {

        }

        protected override void ButtonDefinition_OnExecute(NameValueMap context)
        {
            //System.Windows.Forms.Application.Run(m_Break = new BreakOp(Inventor.Document oDoc));
            m_Cut = new CutOp(InventorApplication.ActiveDocument);
        }

        #endregion
    }

    public class CutOp 
    {
        private PartDocument m_PartDoc;
        private SheetMetalComponentDefinition smcd;
        private SheetMetalFeatures smf;

        public CutOp(Inventor.Document newDoc)
        {
            if (newDoc.DocumentType == DocumentTypeEnum.kAssemblyDocumentObject) m_PartDoc = (PartDocument)newDoc.ActivatedObject;
                if (newDoc.DocumentType == DocumentTypeEnum.kPartDocumentObject) m_PartDoc = (PartDocument)newDoc;
                if (m_PartDoc.SubType == "{9C464203-9BAE-11D3-8BAD-0060B0CE6BB4}") smcd = (SheetMetalComponentDefinition)m_PartDoc.ComponentDefinition;
                smf = (SheetMetalFeatures)smcd.Features;
                CutDefinition cd = smf.CutFeatures.CreateCutDefinition(selectFromSketch());
            try
                {cd.SetCutAcrossBendsExtent("Толщина");}
            catch (Exception)
            {
                cd.SetCutAcrossBendsExtent("Thickness");
            }
                smf.CutFeatures.Add(cd);
            
        }

        private Profile selectFromSketch()
        {
            CommandManager cmdMgr = ((Inventor.Application)m_PartDoc.Parent).CommandManager;
            PlanarSketch ps = (PlanarSketch)cmdMgr.Pick(SelectionFilterEnum.kAllPlanarEntities, "Выберите эскиз:");
            return ps.Profiles.AddForSolid(false);
        }
    }

    public interface unsetStatic
    {
        void unset();
    }
}
                                              