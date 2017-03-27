using System;
using System.Windows.Forms;
using System.IO;
using System.Drawing;
using System.Linq;
using Inventor;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using InvDoc;

namespace InvAddIn
{
    internal class RecoverBtn : Button
    {
        public static RecoverOp m_Recov;
        public Inventor.Document pDoc { get; set; }
        public static ContentCenter cc;
        public static RecoverOp getRecover
        {
            get
            {
                return m_Recov;
            }
        }
        public static Inventor.ContentCenter getCC
        {
            set
            {
                cc = value;
            }
            get
            {
                return cc;
            }
        }

        #region "Methods"
        public RecoverBtn(string displayName, string internalName, CommandTypesEnum commandType, string clientId, string description, string tooltip, Icon standardIcon, Icon largeIcon, ButtonDisplayEnum buttonDisplayType)
            : base(displayName, internalName, commandType, clientId, description, tooltip, standardIcon, largeIcon, buttonDisplayType)
        {

        }
        public RecoverBtn(string displayName, string internalName, CommandTypesEnum commandType, string clientId, string description, string tooltip, ButtonDisplayEnum buttonDisplayType)
            : base(displayName, internalName, commandType, clientId, description, tooltip, buttonDisplayType)
        {

        }

        protected override void ButtonDefinition_OnExecute(NameValueMap context)
        {
            //System.Windows.Forms.Application.Run(m_Break = new BreakOp(Inventor.Document oDoc));
            m_Recov = new RecoverOp(InventorApplication.ActiveDocument);
        }

        #endregion
    }

    public class RecoverOp
    {
        private bool ins = false;
        private Inventor.Application invApp;
        private AssemblyDocument m_AsmDoc;
        private DrawingDocument m_DrwDoc;
        Point2d ptPoz;
        Sheet sh;
        private ComponentOccurrence co, co1, co2;
        private AssemblyComponentDefinition compDef;
        private InsertConstraint insConstr;
        private Inventor.EdgeProxy e1, e2;
        private ObjectsEnumerator objs;
        private Inventor.Point pt;
        private SelectionFilterEnum[] sel = new SelectionFilterEnum[] {SelectionFilterEnum.kAllCircularEntities};
        //private ContentCenter cc;
        private Transaction tr;
        string pathFile;
        bool ao;
        //OpenFileDialog ofd;
        double max, min,offsetDouble, delta = 0.005;
        //InvDoc.XML set, folders;
        Char separator;

        public RecoverOp(Inventor.Document newDoc)
        {
            try
            {
                invApp = (Inventor.Application)newDoc.Parent;
                pathFile = newDoc.FullFileName.ToString();
                pathFile = pathFile.Substring(0, pathFile.LastIndexOf('\\'));
                foreach (Document doc in invApp.Documents.VisibleDocuments)
                {
                    XMLDoc xdoc = new XMLDoc(@"C:\ProgramData\Autodesk\Inventor Addins\rename.xml", "head");
                    recoverReference(doc, xdoc); 
                }
                if (newDoc.DocumentType == DocumentTypeEnum.kAssemblyDocumentObject)
                {
                    m_AsmDoc = (AssemblyDocument)newDoc;
                    recoverAssembly();
                    tr.End();
                }
                if (newDoc.DocumentType == DocumentTypeEnum.kDrawingDocumentObject)
                {
                    m_DrwDoc = newDoc as DrawingDocument;
                    recoverDrawing();
                }
            }
            catch 
            {
                tr.End();
            }
        }

        private void recoverDrawing()
        {
            sh = m_DrwDoc.ActiveSheet;
            foreach (Balloon bal in sh.Balloons)
            {
                if (bal.Attached == false)
                {
                    if (bal.Leader.AllNodes.Count != 2) continue;
                    LeaderNode node = bal.Leader.AllNodes[2];
                    ptPoz = node.Position;
                    GeometryIntent gi = attach(ptPoz);
                    if (gi != null) node.AttachedEntity = gi;
                }
            }
            foreach (LinearGeneralDimension dim in sh.DrawingDimensions.GeneralDimensions)
            {
                if (dim.Attached == false)
                {
                    if (dim.IntentOne == null)
                    {

                    }
                }
            }
            foreach (HoleThreadNote htn in sh.DrawingNotes.HoleThreadNotes)
            {
                if (htn.Attached == false)
                {
                    Point2d poz = ((LineSegment2d)htn.DimensionLine).StartPoint;
                    GeometryIntent gi = attach(poz);
                    if (gi != null) htn.Intent = gi;
                }
            }
        }

        private GeometryIntent attach(Point2d ptPoz)
        {
            ObjectsEnumerator obj = sh.FindUsingPoint(ptPoz);
            if (obj.Count != 0)
            {
                if (obj[1] is DrawingCurveSegment)
                {
                    DrawingCurveSegment dcs = obj[1] as DrawingCurveSegment;
                    GeometryIntent gi = sh.CreateGeometryIntent(dcs.Parent, ptPoz);
                    return gi;
                }
            }
            return null;
        }

        private void recoverReference(Document doc)
        {
            Regex regex = new Regex(@"(\w\.\d*\.\d*)"), regexFirst = new Regex(@"(\w\.\d*\.\d*-\d*)");
            List<string> files;
            foreach (DocumentDescriptor d in doc.ReferencedDocumentDescriptors)
            {
                if (d.ReferenceMissing)
                {
                    string fdn = d.FullDocumentName;
                    string ext = System.IO.Path.GetExtension(fdn), path = System.IO.Path.GetDirectoryName(fdn);
                    IEnumerable<string> ie = Directory.EnumerateFiles(path, "*" + ext, SearchOption.TopDirectoryOnly);
                    files = ie.ToList();
                    Match mf = regexFirst.Match(fdn), m = regex.Match(fdn);
                    string f = null;
                    if (mf.Groups[0].Value != "")
                        f = mf.Groups[0].Value;
                    else if (m.Groups[0].Value != "")
                        f = m.Groups[0].Value;
                    if (f != null)
                    {
                        string ffn = files.FirstOrDefault(el => el.IndexOf(f) != -1);
                        if (ffn != null) d.ReferencedFileDescriptor.ReplaceReference(ffn);
                    }
                }
            }
            foreach (DocumentDescriptor d in doc.ReferencedDocumentDescriptors)
            {
                if (d.ReferenceMissing)
                {
                    string fdn = d.FullDocumentName;
                }
            }
        }

        private void recoverReference(Document doc, XMLDoc xdoc)
        {
            string p = System.IO.Path.GetDirectoryName(doc.FullDocumentName);
            foreach (DocumentDescriptor d in doc.ReferencedDocumentDescriptors)
            {
                string fdn = d.FullDocumentName;
                XElement el = xdoc.find("old", fdn);
                if (el == null)
                {
                    el = xdoc.find("old", System.IO.Path.Combine(p, System.IO.Path.GetFileName(fdn)));
                }
                if (d.FullDocumentName.EndsWith("iam"))
                {
                    Document rd = d.ReferencedDocument as Document;
                    if (rd == null && el != null)
                    {
                        string n = System.IO.Path.Combine(p, System.IO.Path.GetFileName(el.Attribute("new").Value));
                        if (System.IO.File.Exists(n))
                        rd = I.open(n);
                    }
                    if (rd != null) recoverReference(rd, xdoc);
                }
                if (d.ReferenceMissing)
                {
                    if (el != null && System.IO.File.Exists(el.Attribute("new").Value))
                    {
                        d.ReferencedFileDescriptor.ReplaceReference(el.Attribute("new").Value);
                    }
                }
            }
        }

        private void recoverAssembly()
        {
            compDef = (AssemblyComponentDefinition)m_AsmDoc.ComponentDefinition;

            for (int i = 1; i < compDef.Constraints.Count; i++)
            {
                if (compDef.Constraints[i].Type != ObjectTypeEnum.kInsertConstraintObject) continue;
                insConstr = (InsertConstraint)compDef.Constraints[i];
                if (insConstr.HealthStatus == HealthStatusEnum.kDriverLostHealth)
                {
                    ao = insConstr.AxesOpposed;
                    offsetDouble = (double)insConstr.Distance.Value;
                    e1 = (EdgeProxy)insConstr.EntityOne;
                    pt = e1.PointOnEdge;
                    min = 0.03;
                    max = 0.15;
                    do
                    {
                        objs = compDef.FindUsingPoint(pt, ref sel, min);
                        min = objs.Count > 2 ? min - delta : min + delta;
                    }
                    while (objs.Count != 2 || objs.Count < max);

                    if (objs.Count == 2)
                    {
                        e2 = (EdgeProxy)objs[1];
                        insConstr.Delete();
                        insConstr = e1.ContainingOccurrence.Name == e2.ContainingOccurrence.Name ? compDef.Constraints.AddInsertConstraint(e1, (EdgeProxy)objs[2], ao, offsetDouble)
                            : compDef.Constraints.AddInsertConstraint(e1, (EdgeProxy)objs[1], ao, offsetDouble);
                    }
                }
            }
        }
    }
}
