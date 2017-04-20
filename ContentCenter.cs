using System;
using System.Windows.Forms;
using System.Drawing;
using Inventor;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using ExtensionMethods;
using u = InvDoc.u;
using InvDoc;

namespace InvAddIn
{
    internal class ContentBtn : Button
    {
        public static ContentOp m_Cont;
        public Inventor.Document pDoc { get; set; }
        public static ContentCenter cc;
        public static ContentOp getBreak
        {
            get
            {
                return m_Cont;
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
        public ContentBtn(string displayName, string internalName, string clientId, string description, string tooltip, Icon standardIcon, Icon largeIcon)
            : base(displayName, internalName, clientId, description, tooltip, standardIcon, largeIcon)
        {

        }
        public ContentBtn(string displayName, string internalName, CommandTypesEnum commandType, string clientId, string description, string tooltip, ButtonDisplayEnum buttonDisplayType)
            : base(displayName, internalName, commandType, clientId, description, tooltip, buttonDisplayType)
        {

        }

        protected override void ButtonDefinition_OnExecute(NameValueMap context)
        {
            //System.Windows.Forms.Application.Run(m_Break = new BreakOp(Inventor.Document oDoc));
            m_Cont = new ContentOp(InventorApplication.ActiveDocument);
        }

        #endregion
    }

    public class ContentOp
    {
        private Inventor.Application invApp;
        private AssemblyDocument m_AsmDoc;
        private ComponentOccurrence co;
        private AssemblyComponentDefinition compDef;
        //private ContentCenter cc;
        private ContentTreeViewNode treeNode;
        private ContentFamily fam;
        private List<ContentFamily> fams;
        private Matrix mtx;
        private iMateResult iMateRes;
        private InsertConstraint constr;
        private List<Inventor.Edge> edge, edgeflip, edgeCmp, edgeCmp1;
        private MemberManagerErrorsEnum err;
        private Transaction tr;
        private InteractionEvents intEvts;
        private KeyboardEvents key;
        private SelectEvents sel;
        private Inventor.FaceCollection fCol;
        string memberfilename;
        string failuremessage;
        string filename;
        string pathFile;
        bool flag;
        OpenFileDialog ofd;
        double offsetDouble, hd;
        InvDoc.XML set, folders;
        Char separator;
        System.Collections.Generic.List<string> attr = new System.Collections.Generic.List<string>();
        System.Collections.Generic.List<string> val = new System.Collections.Generic.List<string>();
        ComponentOccurrence coOld = null, coNew = null;

        public ContentOp() { }

        public ContentOp(Inventor.Document newDoc)
        {
            if (newDoc.DocumentType == DocumentTypeEnum.kAssemblyDocumentObject) m_AsmDoc = (AssemblyDocument)newDoc;
            invApp = (Inventor.Application)newDoc.Parent;
            pathFile = newDoc.FullFileName.ToString();
            pathFile = pathFile.Substring(0, pathFile.LastIndexOf('\\'));
            compDef = (AssemblyComponentDefinition)m_AsmDoc.ComponentDefinition;
            fCol = invApp.TransientObjects.CreateFaceCollection();
            edge = new List<Edge>(); edgeflip = new List<Edge>(); edgeCmp = new List<Edge>(); edgeCmp1 = new List<Edge>();
            selOp(ref edgeCmp, ref edgeCmp1);
            ofd = new OpenFileDialog();
            ofd.Filter = "XML Files|*.xml";
            ofd.Title = "Выберите файл описания крепежа";
            ofd.InitialDirectory = @"C:\ProgramData\Autodesk\Inventor Addins\";
            ofd.ShowDialog();
            filename = ofd.FileName;
            set = new InvDoc.XML(filename/*@"C:\ProgramData\Autodesk\Inventor Addins\ContentCenter.xml"*/);
            set.ReadXML("Fasteners", ref val, ref attr);
            val.Add("Fasteners"); attr.Add("End");
            separator = System.Globalization.CultureInfo.CurrentCulture.NumberFormat.CurrencyDecimalSeparator[0];

            addFastener();
        }

         ~ContentOp()
        {
            fams = null;
            //cc = null;
        }

        private List<EdgeProxy> findMirrorEdge(ComponentOccurrence occ ,string name, int ind)
         {
             FaceProxy f = occ.SurfaceBodies[1].Faces.OfType<FaceProxy>().FirstOrDefault(face => face.CreatedByFeature is MirrorFeatureProxy && face.CreatedByFeature.Name == name);
            if (f == null) return null;
            MirrorFeatureProxy mfp = f.CreatedByFeature as MirrorFeatureProxy;
            List<EdgeProxy> edges = new List<EdgeProxy>();
            foreach (FaceProxy item in mfp.Faces)
            {
                EdgeProxy ed = item.Edges[ind] as EdgeProxy;
                if (check(ed.ContainingOccurrence, ed)) continue;
                edges.Add(ed);
            }
            return edges;
         }

        private void isRectPat(AssemblyComponentDefinition compDef, ref List<Edge> edges)
         {
             foreach (OccurrencePattern pat in compDef.OccurrencePatterns)
             {
                 try
                 {
                     if (pat is FeatureBasedOccurrencePattern)
                     {
                         object fpp = ((FeatureBasedOccurrencePattern)pat).FeaturePattern;
                         RectangularPatternFeature rpf = ((RectangularPatternFeatureProxy)fpp).NativeObject;
                         if (rpf.ParentFeatures[1] is HoleFeature)
                         {
                             for (int i = 2; i < rpf.PatternElements.Count; i++)
                             {
                                 edges.Add(rpf.PatternElements[i].Faces[1].Edges[1]);
                             }
                         }
                     }
                 }
                 catch
                 {

                 }
             }
         }
        private RectangularPatternFeature isRectPat(AssemblyComponentDefinition compDef, InsertiMateDefinitionProxy ins)
        {
            object ent = ins.Entity;
            Edge ed = null;
            if (ent is EdgeProxy)
            {
                ed = ((EdgeProxy)ent).NativeObject;
                Face face = ed.Faces.OfType<Face>().Where(f => f.SurfaceType == SurfaceTypeEnum.kCylinderSurface).FirstOrDefault();
                if (face == null) return null;
                PartFeature feat = face.CreatedByFeature;
                foreach (RectangularPatternFeature item in (feat.Parent as PartComponentDefinition).Features.RectangularPatternFeatures)
                {
                    PartFeature pFeat = item.ParentFeatures.OfType<PartFeature>().FirstOrDefault(f => f.Equals(feat));
                    if (pFeat != null) 
                        return item;  
                }
            }
            return null;
        }

        private void addRectPat(AssemblyComponentDefinition compDef, ComponentOccurrence occ, ObjectCollection occs, RectangularPatternFeature rpf)
        {
            object result = null; RectangularPatternFeature rpfe = null;
            occ.CreateGeometryProxy(rpf, out result);
            FeatureBasedOccurrencePattern bfop = compDef.OccurrencePatterns.AddFeatureBasedPattern(occs, result as PartFeature);
            foreach (RectangularPatternFeature item in (rpf.Parent as PartComponentDefinition).Features.RectangularPatternFeatures)
            {
                PartFeature pFeat = item.ParentFeatures.OfType<PartFeature>().FirstOrDefault(f => f.Equals(rpf));
                if (pFeat != null)
                    rpfe = item;
            }
            if (rpfe != null)
            {
                ObjectCollection col = I.objs.CreateObjectCollection();
                col.Add(bfop);
                occ.CreateGeometryProxy(rpfe, out result);
                compDef.OccurrencePatterns.AddFeatureBasedPattern(col, result as PartFeature);
            }
        }

        private bool check(ComponentOccurrence occ, object ent)
        {
            bool con = false;
            foreach (AssemblyConstraint constr in occ.Constraints)
            {
                try
                {
                    if (constr.ResultOfiMate == true) continue;
                    if (constr.EntityTwo != null && constr.EntityTwo.Equals(ent))
                    {
                        con = true;
                        break;
                    }
                }
                catch {}
            }
            return con;
        }
        private void insertComp(AssemblyDocument doc, XElement elem, ref ObjectCollection occs, double dist = 0, Edge ed = null, ComponentOccurrence coc = null)
        {
            bool first = true;
            foreach (var attr in elem.Elements())
            {
                if (!attr.HasAttributes && coc != null)
                {
                    coOld = coc; continue;
                }
                bool boolDir = false;
                string dir = "";
                XAttribute att = attr.Attributes().FirstOrDefault(e => e.Name == "Direction");
                if (att != null) dir = att.Value;
                if (dir == "Встречно") boolDir = true;
                if (dir == "Сонаправлено") boolDir = false;
                att = attr.Attributes().FirstOrDefault(at => at.Name == "Offset");
                string offset = "";
                if (att != null) offset = att.Value;
                if (coNew != null && coc == null) coOld = coNew;
                if (attr.Value.StartsWith("v3#"))
                {
                    coNew = place(attr.Value, doc.ComponentDefinition);
                    occs.Add(coNew);
                }
                else if (attr.Value.IndexOf('\\') != -1)
                {
                    Matrix mtx = I.tg.CreateMatrix();
                    coNew = ((AssemblyComponentDefinition)doc.ComponentDefinition).Occurrences.Add(attr.Value, mtx);
                }
                if (offset != "")
                    offsetDouble = double.Parse(offset.Replace('.', ','));
                if (dist != 0 && first)
                {
                    offsetDouble = dist * 10;
                    if (offsetDouble < 0)
                    {
                        boolDir = !boolDir;
                        offsetDouble = -offsetDouble;
                    }
                    first = false;
                }
                if (coOld != null && coNew != coOld)
                {
                    if (offset != "" && boolDir == true)
                        constr = insertMate(coNew, coOld, offsetDouble, true, "Вставка", "IDS_CONNECT", "iInsert", doc.ComponentDefinition);
                    else if (offset != "" && boolDir == false)
                        constr = insertMate(coNew, coOld, offsetDouble, false, "Вставка", "IDS_CONNECT", "iInsert", doc.ComponentDefinition);
                    else
                        iMateRes = insertMate(coNew, coOld, "Вставка", "IDS_CONNECT", "iInsert");
                }
                if (coOld == null && coNew != null)
                {
                    if (offset != "")
                        insertMate(coNew, ed, offsetDouble / 10, boolDir, "Вставка", "IDS_CONNECT", "iInsert", doc.ComponentDefinition);
                    else if (offset == "" && boolDir == true)
                        insertMate(coNew, ed, 0, true, "Вставка", "IDS_CONNECT", "iInsert", doc.ComponentDefinition);
                    else if (offset == "" && boolDir == false)
                        insertMate(coNew, ed, 0, true, "Вставка", "IDS_CONNECT", "iInsert", doc.ComponentDefinition);
                }
            }
        }

        public bool filter(string name, ref string iname)
        {
            int ind = iname.IndexOf("$");
            if (ind != -1)
            {
                var spl = iname.Split('$');
                ind = int.Parse(spl[1]);
                iname = spl[0];
                var occspl = name.Split(':');
                int indooc = int.Parse(occspl[1]);
                if (ind != indooc)
                    return true;
            }
            return false;
        }

        public void programmAdd(AssemblyDocument doc)
         {
            XMLDoc xmldoc = new XMLDoc(@"C:\ProgramData\Autodesk\Inventor Addins\AutoFastener.xml", "Fasteners");
            invApp =  Macros.StandardAddInServer.m_inventorApplication;
            TransientGeometry tg = invApp.TransientGeometry;
            TransientObjects to = invApp.TransientObjects;
            EdgeCollection col = to.CreateEdgeCollection();
            AssemblyComponentDefinition compDef = doc.ComponentDefinition;
            List<Edge> edges = new List<Edge>();
            //List<ComponentOccurrence> occs = new List<ComponentOccurrence>();
            ObjectCollection occs = I.objs.CreateObjectCollection();
            addIMateCompositeResults(compDef, xmldoc.Doc.Root);
            try
            {
                //isRectPat(compDef, ref edges);
            foreach (ComponentOccurrence occ in compDef.Occurrences)
            {
                var imate = occ.iMateDefinitions.OfType<InsertiMateDefinitionProxy>().Where(e => !(e.Name.StartsWith("iIns")));
                if (!imate.Any()) continue;
                foreach (var insImate in imate)
                {
                    string iName = insImate.Name;
                    if (filter(occ.Name, ref iName))
                        continue;
                    //var constrain = doc.ComponentDefinition.Constraints.OfType<InsertConstraint>().Where(con =>/* con.EntityOne.Equals(insImate.Entity) ||*/ con.EntityTwo.Equals(insImate.Entity));
                    //IEnumerator<InsertConstraint> ienum = constrain.GetEnumerator();
                    //if (ienum.Current != null) 
                    //    continue;
                    if (insImate.Suppressed == true) continue;
                    XElement elem = xmldoc.Doc.Root.Elements("Fastener").FirstOrDefault(ele => ele.FirstAttribute.Value == iName);
                     if (elem != null)
                     {
                         //IEnumerable<AssemblyConstraint> ii = occ.Constraints.OfType<AssemblyConstraint>().Where(ent => ent.EntityTwo != null);
                         //AssemblyConstraint con = null;
                         //if (!ii.IsEmpty()) con = ii.FirstOrDefault(co => co.EntityTwo.Equals(insImate.Entity));
                         //else continue;
                         if (check(occ, insImate.Entity)) continue;
                         insertComp(doc, elem, ref occs, (double)insImate.Distance.Value, (Edge)insImate.Entity);

                         RectangularPatternFeature rpf = isRectPat(compDef, insImate);
                         if (rpf != null)
                         {
                             addRectPat(compDef, occ, occs, rpf);
                             VariableData vd = new VariableData(doc as Document);
                             Inventor.Attribute att = vd.getAttrib(rpf, "Mirror");
                             if (att != null)
                             {
                                string nameRP = att.Value as string;
                                List<EdgeProxy> eds = findMirrorEdge(occ, nameRP, getInd(insImate));
                                foreach (EdgeProxy item in eds)
                                {
                                    coOld = null; coNew = null;
                                    insertComp(doc, elem,ref occs, (double)insImate.Distance.Value, (Edge)item); 
                                }
                             }
                         }
                         coOld = null; coNew = null; occs.Clear();
                     }
                 }
             }
            addToFolder(ref doc);
            }
            catch {}
         }

        private int getInd(InsertiMateDefinitionProxy insImate)
        {
            int i = 1;
            EdgeProxy ed = insImate.Entity as EdgeProxy;
            FaceProxy f = ed.Faces.OfType<FaceProxy>().FirstOrDefault(face => face.CreatedByFeature is HoleFeatureProxy);
            if (f == null) return i;
            for (int j = 1; j < f.Edges.Count; j++)
            {
                if (f.Edges[j].Equals(ed)) return j;
            }
            return i;
        }

        private void addIMateCompositeResults(AssemblyComponentDefinition compDef, XElement el)
        {
            string name = ""; CompositeiMateDefinition imateDef = null; ComponentOccurrence coNew = null;
            foreach (ComponentOccurrence occ in compDef.Occurrences)                                         
            {
                //IEnumerable<iMateDefinition> imate = u.gets<iMateDefinition>(occ.iMateDefinitions, f => f is CompositeiMateDefinitionProxy);
                var imate = occ.iMateDefinitions.OfType<CompositeiMateDefinitionProxy>();
                if (!imate.Any()) continue;
                foreach (CompositeiMateDefinitionProxy insImate in imate)
                {
                    string iName = insImate.Name;
                    if (filter(occ.Name, ref iName))
                        continue;
                    if (insImate.IsConsumed == true) continue;
                    XElement elem = el.Elements("Composite").FirstOrDefault(ele => ele.FirstAttribute.Value == iName);
                    if (elem != null)
                    {
                        imateDef = (CompositeiMateDefinition)insImate;
                        name = iName;
                        foreach (var item in elem.Elements())
                        {
                            if (item.HasAttributes && item.Attribute("name") != null) 
                            {
                                name = item.FirstAttribute.Value; imateDef = coNew.iMateDefinitions.OfType<CompositeiMateDefinition>().FirstOrDefault(e => e.Name == name);
                            }
                            coNew = null;
                            if (item.Value.StartsWith("v3#"))
                            {
                                coNew = place(item.Value, compDef);
                            }
                            else if (elem.Value.IndexOf('\\') != -1)
                            {
                                Matrix mtx = I.tg.CreateMatrix();
                                coNew = (compDef).Occurrences.Add(item.Value, mtx);
                            }
                            if (coNew != null)
                            {
                                IEnumerable<CompositeiMateDefinition> ie = coNew.iMateDefinitions.OfType<CompositeiMateDefinition>();
                                CompositeiMateDefinition iComp = ie.FirstOrDefault(e => e.Name == name);
                                if (iComp == null && ie.Count() == 1)
                                {
                                    iComp = ie.First();
                                    if (elem.Elements().Count() > 1)
                                    {
                                        ObjectCollection occs = I.objs.CreateObjectCollection();
                                        insertComp(compDef.Parent as AssemblyDocument, elem, ref occs, 0,
                                            coNew.iMateDefinitions.OfType<InsertiMateDefinition>().ElementAt(1).Entity as Edge, coNew);
                                    }
                                }
                                if (iComp != null)
                                {
                                    compDef.iMateResults.AddByTwoiMates((iMateDefinition)iComp, (iMateDefinition)imateDef);
                                }
                                else coNew.Delete();
                            }
                        }
                    }
                }
            }
        }

        private void addFastener()
         {
             try
             {
                 int j = 0; string holediameter = "";
                 List<string> attr1 = new List<string>();
                 List<string> val1 = new List<string>();
                 fams = new List<ContentFamily>();
                 tr = ((Inventor.Application)m_AsmDoc.Parent).TransactionManager.StartTransaction((_Document)m_AsmDoc, "Автовставка");

                 while (j <= attr.Count - 2)
                 {
                     /*if (j == 0)*/
                     holediameter = set.substring(attr[j], "HoleDiameter=").Replace('.', separator);
                     //else holediameter = substring(attr[j-1], "HoleDiameter=").Replace('.', separator);
                     edge.Clear(); edgeflip.Clear(); attr1.Clear(); val1.Clear(); fCol.Clear();
                     if (holediameter != "")
                         hd = Convert.ToDouble(holediameter);
                     if (edgeCmp.Count != 0)
                     {
                         edge = edgeCmp;
                         edgeflip = edgeCmp1;
                     }
                     else
                     { findHoles(ref edge, ref edgeflip, hd); }

                     coNew = null;

                     for (int x = j + 1; x < val.Count; x++)
                     {
                         if (val[x] == "Fasteners")
                         {
                             j = x;
                             break;
                         }
                         attr1.Add(attr[x]);
                         val1.Add(val[x]);
                     }

                     for (int k = 0; k < edge.Count; k++)
                     {

                         for (int i = 0; i < val1.Count; i++)
                         {
                             if (val1[i] != "Fasteners" && edge.Count != 0)
                             {
                                 if (val1[i].IndexOf('\\') == -1)
                                 {
                                     if (fams.Count != 0)
                                     {
                                         fam = fams.Find(delegate(ContentFamily cf)
                                         {
                                             return /*cf.DisplayName.IndexOf(val1[i]) != -1;*/ val1[i].IndexOf(cf.DisplayName) != -1;
                                         });
                                         if (fam == null)
                                         {
                                             fam = find(val1[i]); fams.Add(fam);
                                         }
                                     }
                                     else
                                     {
                                         fam = find(val1[i]); fams.Add(fam);
                                     }
                                 }
                                 string d = set.substring(attr1[i], "Diameter=").Replace('.', separator);
                                 string dname = set.substring(attr1[i], "DiameterName=");
                                 string len = set.substring(attr1[i], "Lenght=").Replace('.', separator);
                                 string lname = set.substring(attr1[i], "DiameterLen=");
                                 string thick = set.substring(attr1[i], "Thickness=").Replace('.', separator);
                                 string thickname = set.substring(attr1[i], "ThicknessName=");
                                 string offset = (set.substring(attr1[i], "Offset=").Replace('.', separator));
                                 if (offset != "")
                                     offsetDouble = Convert.ToDouble(offset);
                                 string dir = set.substring(attr1[i], "Direction=");
                                 bool boolDir = false;
                                 if (dir == "Встречно") boolDir = true;
                                 if (dir == "Сонаправлено") boolDir = false;

                                 if (coNew != null) coOld = coNew;
                                 if (val1[i].StartsWith("v3#")) coNew = place(val1[i],compDef);
                                     else if (val1[i].IndexOf('\\') != -1)
                                 {
                                     coNew = compDef.Occurrences.Add(val1[i], I.tg.CreateMatrix());
                                 }
                                 else
                                 {
                                     if (d != "" && len != "")
                                         coNew = place(fam, Convert.ToDouble(d), dname, Convert.ToDouble(len), lname);
                                     if (d != "" && len == "" && thick == "")
                                         coNew = place(fam, Convert.ToDouble(d), dname);
                                     if (d != "" && len == "" && thick != "")
                                         coNew = place(fam, Convert.ToDouble(d), dname, Convert.ToDouble(thick), thickname);
                                 }
                                 if (coOld != null && coNew != coOld)
                                 {
                                     if (offset != "" && boolDir == true)
                                         constr = insertMate(coNew, coOld, offsetDouble, true, "Вставка", "IDS_CONNECT", "iInsert",compDef);
                                     else if (offset != "" && boolDir == false)
                                         constr = insertMate(coNew, coOld, offsetDouble, false, "Вставка", "IDS_CONNECT", "iInsert",compDef);
                                     else
                                         iMateRes = insertMate(coNew, coOld, "Вставка", "IDS_CONNECT", "iInsert");
                                 }
                                 if (coOld == null && coNew != null)
                                 {
                                     if (offset != "")
                                         insertMate(coNew, edge[k], offsetDouble / 10, boolDir, "Вставка", "IDS_CONNECT", "iInsert", compDef);
                                     else if (offset == "" && boolDir == true)
                                         insertMate(coNew, edge[k], 0, true, "Вставка", "IDS_CONNECT", "iInsert", compDef);
                                     else if (offset == "" && boolDir == false)
                                         insertMate(coNew, edgeflip[k], 0, true, "Вставка", "IDS_CONNECT", "iInsert", compDef);
                                 }
                                 //else
                                 //{
                                 //    insertMate(coNew, edge[k], 0, true, "Вставка", "IDS_CONNECT", "iInsert");
                                 //}
                             }
                             offsetDouble = 0;

                         }
                         coNew = null; coOld = null;
                     }
                 }
                 fams = null;
                 addToFolder(ref m_AsmDoc);
                 tr.End();
             }
             catch (Exception ex)
             {
                 MessageBox.Show(ex.ToString());
             }
         }

        private void insertComponent(string strAttr, ref ComponentOccurrence coNew, ref ComponentOccurrence coOld)
        {
            string d = set.substring(strAttr, "Diameter=").Replace('.', separator);
            string dname = set.substring(strAttr, "DiameterName=");
            string len = set.substring(strAttr, "Lenght=").Replace('.', separator);
            string lname = set.substring(strAttr, "DiameterLen=");
            string thick = set.substring(strAttr, "Thickness=").Replace('.', separator);
            string thickname = set.substring(strAttr, "ThicknessName=");

            if (coNew != null) coOld = coNew;
            if (d != "" && len != "")
                coNew = place(fam, Convert.ToDouble(d), dname, Convert.ToDouble(len), lname);
            if (d != "" && len == "" && thick == "")
                coNew = place(fam, Convert.ToDouble(d), dname);
            if (d != "" && len == "" && thick != "")
                coNew = place(fam, Convert.ToDouble(d), dname, Convert.ToDouble(thick), thickname);
            if (coOld != null && coNew != coOld)
            {
                iMateRes = insertMate(coNew, coOld, "Вставка", "IDS_CONNECT", "iInsert");
            }
            //else
            //{
            //    insertMate(coNew, edge[k], 0, true, "Вставка", "IDS_CONNECT", "iInsert");
            //}
        }

        private void findHoles(ref List<Inventor.Edge> lst, ref List<Inventor.Edge> lstflip, double d)
        {
            
            foreach (ComponentOccurrence co in compDef.Occurrences)
            {
                findEdges(co,ref lst,ref lstflip,d,co);
                if (co.Suppressed == false && co.SubOccurrences.Count != 0)
                {
                    foreach (ComponentOccurrence co1 in co.SubOccurrences)
                    {
                        ComponentOccurrence co3 = co1;
                        findEdges(co1, ref lst, ref lstflip, d,co3);
                   }
                }
            }    
        }

        private void findEdges(ComponentOccurrence co, ref List<Inventor.Edge> lst, ref List<Inventor.Edge> lstflip, double d, ComponentOccurrence co1)
        {
            bool flag = false;
            if (co.DefinitionDocumentType == DocumentTypeEnum.kPartDocumentObject)
                {
                    Inventor.Document oDoc = (Document)co.Definition.Document;
                    //fCol = faceFromScketch((PartDocument)oDoc);
                    if (oDoc.SubType == "{9C464203-9BAE-11D3-8BAD-0060B0CE6BB4}")
                    {
                        foreach (PartFeature pf in co.SurfaceBodies[1].AffectedByFeatures)
                        {
                            if (pf.Type == ObjectTypeEnum.kHoleFeatureProxyObject)
                            {
                                HoleFeatureProxy hf = (HoleFeatureProxy)pf;
                                foreach (Inventor.Face face in hf.Faces)
                                {
                                    if (face.Edges[1].GeometryType == CurveTypeEnum.kCircleCurve)
                                    {
                                        Circle geom = (Circle)face.Edges[1].Geometry;
                                        if ((int)(geom.Radius * 1000) == (int)(d * 50))
                                        {
                                            flag = findInAsm(face.Edges[1], co1);
                                            if (!flag) flag = findInAsm(face.Edges[2], co1);
                                            if (!flag)
                                            {
                                                lst.Add(face.Edges[1]);
                                                lstflip.Add(face.Edges[2]);
                                            }
                                        }
                                    }
                                }
                            }
                            if (pf.Type == ObjectTypeEnum.kRectangularPatternFeatureProxyObject)
                            {
                                RectangularPatternFeatureProxy ro = (RectangularPatternFeatureProxy)pf;
                                foreach (Inventor.Face face in ro.Faces)
                                {
                                    if (face.Edges[1].GeometryType == CurveTypeEnum.kCircleCurve)
                                    {
                                        Circle geom = (Circle)face.Edges[1].Geometry;
                                        if ((int)(geom.Radius * 1000) == (int)(d * 50))
                                        {
                                            flag = findInAsm(face.Edges[1], co1);
                                            if (!flag) flag = findInAsm(face.Edges[2], co1);
                                            if (!flag)
                                            {
                                                lst.Add(face.Edges[1]);
                                                lstflip.Add(face.Edges[2]);
                                            }
                                        }
                                    }
                                }
                            }
                            //if (pf.Type == ObjectTypeEnum.kiFeatureObject)
                            //{
                            //    iFeature iF = (iFeature)pf;
                            //    foreach (Inventor.Face face in iF.Faces)
                            //    {
                            //        if (face.Edges[1].GeometryType == CurveTypeEnum.kCircleCurve)
                            //        {
                            //            Circle geom = (Circle)face.Edges[1].Geometry;
                            //            if ((int)(geom.Radius * 1000) == (int)(d * 50))
                            //            {
                            //                flag = findInAsm(face.Edges[1], co1);
                            //                if (!flag) flag = findInAsm(face.Edges[2], co1);
                            //                if (!flag)
                            //                {
                            //                    lst.Add(face.Edges[1]);
                            //                    lstflip.Add(face.Edges[2]);
                            //                }
                            //            }
                            //        }
                            //    }
                            //}
                        }
                    }
                }                                                      
        }

        private bool cmpD(Inventor.Edge edge, double d)
        {
            if (edge.GeometryType == CurveTypeEnum.kCircleCurve)
            {
                Circle geom = (Circle)edge.Geometry;
                if ((int)(geom.Radius * 1000) == (int)(d * 50))
                    return true;
            }
            return false;
        }

        private void insertMate (ComponentOccurrence co, Edge edge, double offset, bool axes, string name, string nameEng, string nameEng2,AssemblyComponentDefinition compDef)
        {
           foreach (iMateDefinition idef1 in co.iMateDefinitions)
           {
               if (idef1.Name.IndexOf(name) != -1 || idef1.Name.IndexOf(nameEng) != -1 || idef1.Name.IndexOf(nameEng2) != -1)
               {
                   Edge ed1 = (Edge)((InsertiMateDefinitionProxy)idef1).Entity;
                   compDef.Constraints.AddInsertConstraint(ed1, edge, axes, offset);
                   break;
               }
           }
        }

        private iMateResult insertMate(ComponentOccurrence co1, ComponentOccurrence co2, string name, string nameEng, string nameEng2)
        {
            bool flag = false;
            foreach (iMateDefinition idef1 in co1.iMateDefinitions)
            {
                if (idef1.Name.IndexOf(name) != -1 || idef1.Name.IndexOf(nameEng) != -1 || idef1.Name.IndexOf(nameEng2) != -1)
                {
                    flag = ((InsertiMateDefinitionProxy)idef1).IsConsumed;
                    if (flag) continue;
                    flag = false;
                    foreach (iMateDefinition idef2 in co2.iMateDefinitions)
                    {
                        if (idef2.Name.IndexOf(name) != -1 || idef2.Name.IndexOf("IDS_CONNECT") != -1 || idef2.Name.IndexOf(nameEng2) != -1)
                        {
                            flag = ((InsertiMateDefinitionProxy)idef2).IsConsumed;
                            if (flag) continue;
                            flag = false;
                            return compDef.iMateResults.AddByTwoiMates(idef1, idef2);
                        }
                    }
                }
            }
            return null;    
        }

        private InsertConstraint insertMate(ComponentOccurrence co1, ComponentOccurrence co2, double offset ,bool axes , string name, string nameEng, string nameEng2, AssemblyComponentDefinition compDef,
             int start1 = 0, int start2 = 0)
        {
            bool flag = false; int count1 = 0, count2 = 0;
            foreach (iMateDefinition idef1 in co1.iMateDefinitions)
            {
                if (idef1.Name.IndexOf(name) != -1 || idef1.Name.IndexOf(nameEng) != -1 || idef1.Name.IndexOf(nameEng2) != -1)
                {
                    flag = ((InsertiMateDefinitionProxy)idef1).IsConsumed;
                    if (flag) continue;
                    flag = false;
                    count1++;
                    if (count1 < start1) continue;
                    foreach (iMateDefinition idef2 in co2.iMateDefinitions)
                    {
                        if (idef2.Name.IndexOf(name) != -1 || idef2.Name.IndexOf("IDS_CONNECT") != -1 || idef2.Name.IndexOf(nameEng2) != -1)
                        {
                            flag = ((InsertiMateDefinitionProxy)idef2).IsConsumed;
                            if (flag) continue;
                            flag = false;
                            count2++;
                            if (count2 < start2) continue;
                            return compDef.Constraints.AddInsertConstraint(((InsertiMateDefinitionProxy)(idef1)).Entity, ((InsertiMateDefinitionProxy)(idef2)).Entity,
                                axes, offset/10);
                        }
                    }
                }
            }
            return null;
        }

        private bool findInAsm(iMateDefinition idef, AssemblyComponentDefinition acd)
        {
            foreach (iMateResult ires in acd.iMateResults)
                    {
                        if (idef.ReferencedEntity == ires.Constraints[1].EntityOne || idef.ReferencedEntity == ires.Constraints[1].EntityTwo)
                            return true;
                    }
            return false;
        }

        private bool findInAsm(Edge edge, ComponentOccurrence co)
        {
            Circle cir1;
            cir1 = (Circle)edge.Geometry;
            Inventor.Point pt1 = cir1.Center; 
            //pt1.TransformBy(co.Transformation);
            List<InsertConstraint> insconstr = new List<InsertConstraint>();
            foreach (ComponentOccurrence co1 in co.OccurrencePath)
            {
                try
                {
                    foreach (InsertConstraint ires in co1.Constraints)
                    {
                        if (ires.Suppressed == false)
                        insconstr.Add(ires);
                    }
                }
                catch (Exception)
                {
                    continue;
                }
            }

            foreach (InsertConstraint ires in insconstr)
            {

                Circle e1 = (Circle)ires.GeometryOne; Circle e2 = (Circle)ires.GeometryTwo;
                Inventor.Point pt2 = e1.Center; Inventor.Point pt3 = e2.Center;
                //pt2.TransformBy(co.Transformation); pt3.TransformBy(co.Transformation);
                //if (edge.Equals(ires.EntityOne) || edge.Equals(ires.EntityTwo))
                //if (cir1.Equals(ires.GeometryOne) || cir1.Equals(ires.GeometryTwo))
                if (eq(pt1,pt2,pt3))
                    return true;
            }
            return false;
        }

        private bool eq(Inventor.Point pt1, Inventor.Point pt2, Inventor.Point pt3)
        {
            if ((int)pt1.X * 10000 == (int)pt2.X * 10000 && (int)pt1.Y * 10000 == (int)pt2.Y * 10000 && (int)pt1.Z * 10000 == (int)pt2.Z * 10000 ||
                (int)pt1.X * 10000 == (int)pt3.X * 10000 && (int)pt1.Y * 10000 == (int)pt3.Y * 10000 && (int)pt1.Z * 10000 == (int)pt3.Z * 10000)
            {
                return true;
            }
            return false;
        }

        private ComponentOccurrence place(ContentFamily cf, double d,string dname)
        {
            mtx = invApp.TransientGeometry.CreateMatrix();
            //foreach (ContentTableColumn col in cf.TableColumns)
            //{
            //    memberfilename = col.DisplayHeading;
            //    object exp = col.Expression;
            //    bool val = col.HasPropertyMap;

            //}
            foreach (ContentTableRow row in cf.TableRows)
            {
                if (Convert.ToDouble(row.GetCellValue(dname)) == d)
                {
                    if (Macros.StandardAddInServer.StandartPart == "1")
                        memberfilename = cf.CreateMember(row, out err, out failuremessage);
                    else
                    {
                        string name = pathFile + "\\" + row.GetCellValue("FILENAME").Replace("/", "-") + ".ipt";
                        if (System.IO.File.Exists(name))
                        {
                            memberfilename = name;
                        }
                        else
                        {
                            memberfilename = cf.CreateMember(row, out err, out failuremessage, Custom: true, FileName: name);
                        }
                    }
                    co = compDef.Occurrences.Add(memberfilename, mtx);
                    return co;
                    //min = co.RangeBox.MinPoint.X; max = co.RangeBox.MaxPoint.Y;
                    //offset += (max - min) * 1.1;
                    //mtx.Cell[2, 4] = offset;
                    //break;
                }
            }
            return null;
        }

        private void placeAll(string path)
        {
            ContentFamily cf = find(path);
            foreach (var row in cf.TableRows)
            {
                memberfilename = cf.CreateMember(row, out err, out failuremessage);
            }
        }

        private ComponentOccurrence place(ContentFamily cf, double d, string dname, double l, string lname)
        {
            mtx = invApp.TransientGeometry.CreateMatrix();
            foreach (ContentTableRow row in cf.TableRows)
            {
                if (Convert.ToDouble(row.GetCellValue(dname)) == d &&
                    Convert.ToDouble(row.GetCellValue(lname)) == l)
                {
                    if (Macros.StandardAddInServer.StandartPart == "1")
                        memberfilename = cf.CreateMember(row, out err, out failuremessage);
                    else
                        {
                            string name = pathFile + "\\" + row.GetCellValue("FILENAME").Replace("/", "-") + ".ipt";
                            if (System.IO.File.Exists(name))
                            {
                                memberfilename = name;
                            }
                            else
                            {
                                memberfilename = cf.CreateMember(row, out err, out failuremessage, Custom: true, FileName: name);
                            }
                        }
                    co = compDef.Occurrences.Add(memberfilename, mtx);
                    return co;
                }
            }
            return null;
        }

        private ComponentOccurrence place(string id, AssemblyComponentDefinition compDef)
        {
            mtx = invApp.TransientGeometry.CreateMatrix();
            ContentCenter c = invApp.ContentCenter;
            ContentFamily cf = (ContentFamily)c.GetContentObject(id.Substring(0,id.LastIndexOf('#')+1));
            ContentTableRow row = (ContentTableRow)c.GetContentObject(id);
                    if (Macros.StandardAddInServer.StandartPart == "1")
                        memberfilename = cf.CreateMember(row, out err, out failuremessage);
                    else
                    {
                        string name = pathFile + "\\" + row.GetCellValue("FILENAME").Replace("/", "-") + ".ipt";
                        if (System.IO.File.Exists(name))
                        {
                            memberfilename = name;
                        }
                        else
                        {
                            memberfilename = cf.CreateMember(row, out err, out failuremessage, Custom: true, FileName: name);
                        }
                    }
                    co = compDef.Occurrences.Add(memberfilename, mtx);
                    return co;
            //return null;
        }

        static public string memberForPlace(string id)
        {
            string memberfilename;
            string failuremessage;
            MemberManagerErrorsEnum err;
            Matrix mtx = I.tg.CreateMatrix();
            ContentCenter c = Macros.StandardAddInServer.m_inventorApplication.ContentCenter;
            ContentFamily cf = (ContentFamily)c.GetContentObject(id.Substring(0, id.LastIndexOf('#') + 1));
            ContentTableRow row = (ContentTableRow)c.GetContentObject(id);
            return memberfilename = cf.CreateMember(row, out err, out failuremessage);
        }

        //private void printToFile(string fileName)
        //{
        //    System.IO.StreamWriter textFile = new System.IO.StreamWriter(@"C:\path\textfile.txt");
        //    textFile.WriteLine("Hello World!");
        //    textFile.WriteLine("And goodbye");
        //    textFile.Close();
        //}

        private ContentFamily find(string path)
        {
            if (path.StartsWith("v3#"))
            {
                string findStr = path.Substring(0,path.LastIndexOf('#')+1);
                return (ContentFamily)invApp.ContentCenter.GetContentObject(findStr);
            }
            string[] spl = path.Split(':');
            //ContentCenter cc;
            ContentTreeViewNode ctvn;
            invApp = ((Inventor.Application)m_AsmDoc.Parent);
            //if (InvAddIn.ContentBtn.cc == null)
            //ContentCenterAccessOptionEnum opt;
            //string loc;
            //invApp.ContentCenterOptions.GetAccessOption(out opt, out loc);
            //invApp.ContentCenterOptions.CustomFamilyAsStandard = true;
            ContentCenter ccc = invApp.ContentCenter;
            ctvn = InvAddIn.ContentBtn.cc.TreeViewTopNode;

            for (int i = 0; i < spl.Length - 1; i++)
            {
                try 
                {
                    ctvn = ctvn.ChildNodes[spl[i]];
                }
                catch (Exception)
                {
                    ctvn = ctvn.ChildNodes[spl[i]];
                }
            }

            //ctvn = findNode(ctvn, 0, spl);

                                       
            ContentFamily cf;
            foreach (ContentFamily fam in ctvn.Families)
            {
                if (fam.DisplayName == spl[spl.Length - 1]) { cf = fam; return cf ; }
            }
            return null;
        }

        private ContentTreeViewNode findNode(ContentTreeViewNode ctvn,int i, string[] path)
        {
            for (int j=1; j <= ctvn.ChildNodes.Count; j++)
            {
                if (ctvn.ChildNodes[j].DisplayName == path[i])
                {
                    i++;
                    ContentTreeViewNode node = ctvn.ChildNodes[j];
                    return findNode(node, i, path);
                }
            }
            treeNode = ctvn;
            return ctvn;
        }

        public bool containtsNode(BrowserNode node, BrowserFolder folder)
        {
            string[] spl = node.FullPath.Split(':');
            string name = spl[spl.Length - 2] + ":" + spl[spl.Length-1];
            BrowserNode tmp = folder.BrowserNode.BrowserNodes.OfType<BrowserNode>().FirstOrDefault(e => e.FullPath.EndsWith(name));
            return tmp == null ? false : true;
        }

        public void addToFolder(ref AssemblyDocument asmDoc)
        {
            AssemblyComponentDefinition compDef = asmDoc.ComponentDefinition;
            BrowserPane pane = asmDoc.BrowserPanes["Модель"];
            ObjectCollection occNodes = invApp.TransientObjects.CreateObjectCollection();
            BrowserNode parentNode = pane.TopNode;
            string[] tmp;
            System.Collections.Generic.List<string> match = new System.Collections.Generic.List<string>();
            System.Collections.Generic.List<string> result = new System.Collections.Generic.List<string>();
            folders = new InvDoc.XML(@"C:\ProgramData\Autodesk\Inventor Addins\Folders.xml");
            folders.ReadXML("Folders", ref match, ref result);

            foreach (ComponentOccurrence occ in compDef.Occurrences)
            {
                BrowserNode node = pane.GetBrowserNodeFromObject(occ);
                string name = node.FullPath;
                tmp = name.Split(':');
                if (tmp.Length > 3) continue;
                if (occ.IsPatternElement == false)
                {
                    for (int i = 0; i < result.Count; i++)
                    {
                        tmp = match[i+1].Split('$');
                        foreach (string str in tmp)
                        {
                            if (name.IndexOf(str) != -1)
                            {
                                string ss = result[i].Substring(5);
                                ss = ss.Substring(0, ss.Length - 1);
                                BrowserFolder folder = addFolder(ss, ref asmDoc);
                                if (!containtsNode(node, folder))
                                folder.Add(node);
                            }
                        }
                    }

                    //if (name.IndexOf("Гайка-заклепка") != -1 || name.IndexOf("Закл.") != -1)
                    //{
                    //    BrowserFolder folder = addFolder("Гайки-заклепки", ref asmDoc);
                    //    folder.Add(node);
                    //}
                    //else if (name.IndexOf("Заклепка тяговая") != -1)
                    //{
                    //    BrowserFolder folder = addFolder("Заклепки тяговые", ref asmDoc);
                    //    folder.Add(node);
                    //}
                    //else if (name.IndexOf("Винт") != -1 || name.IndexOf("DIN 7380") != -1)
                    //{
                    //    BrowserFolder folder = addFolder("Винты", ref asmDoc);
                    //    folder.Add(node);
                    //}
                    //else if (name.IndexOf("Шайба") != -1)
                    //{
                    //    BrowserFolder folder = addFolder("Шайбы", ref asmDoc);
                    //    folder.Add(node);
                    //}
                    //else if (name.IndexOf("Болт") != -1)
                    //{
                    //    BrowserFolder folder = addFolder("Болты", ref asmDoc);
                    //    folder.Add(node);
                    //}
                }
            }
        }

        public BrowserFolder addFolder(string name, ref AssemblyDocument asmDoc)
        {
            AssemblyComponentDefinition compDef = asmDoc.ComponentDefinition;
            BrowserPane pane = asmDoc.BrowserPanes["Модель"];
            try
            {
                return pane.TopNode.BrowserFolders[name];
            }
            catch (Exception)
            {
                return pane.AddBrowserFolder(name);
            }
        }
        public void selOp(ref List<Edge> edgeCmp, ref List<Edge> edgeCmp1)
        {
            CommandManager cmdMgr = Macros.StandardAddInServer.m_inventorApplication.CommandManager;
                    //input = cmdMgr.UserInputEvents;
            try
            {
                intEvts = cmdMgr.CreateInteractionEvents();
                intEvts.InteractionDisabled = false;
                sel = intEvts.SelectEvents;
                sel.WindowSelectEnabled = true;
                sel.AddSelectionFilter(SelectionFilterEnum.kPartEdgeCircularFilter);
                sel.OnSelect += new SelectEventsSink_OnSelectEventHandler(select);
                key = intEvts.KeyboardEvents;
                key.OnKeyPress += new KeyboardEventsSink_OnKeyPressEventHandler(keyOp);
                intEvts.Start();
                intEvts.StatusBarText = "Выберите отверстия:";
                //input.OnSelect += new UserInputEventsSink_OnSelectEventHandler(select);
                //face = (Face)cmdMgr.Pick(SelectionFilterEnum.kPartFaceFilter, "Выберите поверхность:");

                flag = true;
                while (flag)
                {
                    Macros.StandardAddInServer.m_inventorApplication.UserInterfaceManager.DoEvents();
                }

                for (int i = 0; i < sel.SelectedEntities.Count; i++)
                {
                    try
                    {
                        Edge ed = (Edge)sel.SelectedEntities[i + 1];
                        Inventor.Face f = (ed.Faces[1].SurfaceType == SurfaceTypeEnum.kCylinderSurface) ? ed.Faces[1] : ed.Faces[2];
                        edgeCmp.Add(ed);
                        ed = (f.Edges[1].Equals(ed)) ? f.Edges[2] : f.Edges[1];
                        edgeCmp1.Add(ed);
                    }
                    catch (Exception) { }
                }

                intEvts.Stop();

                sel.OnSelect -= new SelectEventsSink_OnSelectEventHandler(select);
                key.OnKeyPress -= new KeyboardEventsSink_OnKeyPressEventHandler(keyOp);
                sel = null;
                key = null;
                intEvts = null;
            }
            catch (Exception)
            {
                sel = null;
                key = null;
                intEvts = null;
            }
        }

        private void select(ObjectsEnumerator JustSelectedEntities, SelectionDeviceEnum SelectionDevice, Inventor.Point ModelPosition, Point2d ViewPosition, Inventor.View View)
        {
        }
        private void keyOp(int KeyASCII)
        {
            if (KeyASCII == 32)
            {
                Edge ed = (Edge)sel.SelectedEntities[sel.SelectedEntities.Count];
                ComponentOccurrence co = ((EdgeProxy)ed).ContainingOccurrence;
                double r = ((Circle)ed.Geometry).Radius;
                Inventor.Face f = (ed.Faces[1].SurfaceType == SurfaceTypeEnum.kPlaneSurface) ? ed.Faces[1] : ed.Faces[2];
                foreach (Edge e in f.Edges)
                {
                    if (e.GeometryType == CurveTypeEnum.kCircleCurve && ((Circle)e.Geometry).Radius == r)
                    {
                        if (!findInAsm(e,co))
                        sel.AddToSelectedEntities(e);
                    }
                }
            }
            else if (KeyASCII == 13)
            {
                flag = false;
            }
        }
        public FaceCollection faceFromScketch(PartDocument doc)
        {
            FaceCollection fc = invApp.TransientObjects.CreateFaceCollection();
            foreach (PlanarSketch s in doc.ComponentDefinition.Sketches)
            {
                try
                { fc.Add((Face)s.PlanarEntity); }
                catch (Exception) { }
             }
            return fc;
        }

    }
}

namespace ExtensionMethods
{
    public static class MyIEnumerable
    {
        public static bool IsEmpty<T>(this IEnumerable<T> items)
        {
            var enumerator = items.GetEnumerator();
            return (enumerator.Current == null) ? true : false;
        }
    }
}
