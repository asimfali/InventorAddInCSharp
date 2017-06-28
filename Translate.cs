using System;
using System.Windows.Forms;
using System.Drawing;
using Inventor;
using System.Linq;
using InvDoc;
using ADDIN = Macros.StandardAddInServer;
using System.Collections;
using ut = InvDoc.u;
using System.Collections.Generic;

namespace InvAddIn
{
    internal class PDFButton : Button
    {
        public static PDFOp m_PDF;
        public Inventor.Document pDoc { get; set; }
        public static PDFOp getPDF
        {
            get
            {
                return m_PDF;
            }
        }

        #region "Methods"
        public PDFButton(string displayName, string internalName, string clientId, string description, string tooltip, Icon standardIcon, Icon largeIcon)
            : base(displayName, internalName, clientId, description, tooltip, standardIcon, largeIcon)
        {

        }

        protected override void ButtonDefinition_OnExecute(NameValueMap context)
        {
            //System.Windows.Forms.Application.Run(m_Break = new BreakOp(Inventor.Document oDoc));
            try
            {
                if (InventorApplication.ActiveDocument.DocumentType == DocumentTypeEnum.kAssemblyDocumentObject)
                    m_PDF = new PDFOp((AssemblyDocument)InventorApplication.ActiveDocument);
                else
                m_PDF = new PDFOp(InventorApplication.ActiveDocument);
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.ToString());
            }

        }

        #endregion
    }

    internal class Print : Button
    {
        public static PrintOp m_Print;
        public Inventor.Document pDoc { get; set; }
        public static PrintOp getPrint
        {
            get
            {
                return m_Print;
            }
        }

        #region "Methods"
        public Print(string displayName, string internalName, string clientId, string description, string tooltip, Icon standardIcon, Icon largeIcon)
            : base(displayName, internalName, clientId, description, tooltip, standardIcon, largeIcon)
        {

        }

        protected override void ButtonDefinition_OnExecute(NameValueMap context)
        {
            //System.Windows.Forms.Application.Run(m_Break = new BreakOp(Inventor.Document oDoc));
            m_Print = new PrintOp(InventorApplication.ActiveDocument, InventorApplication);
        }

        #endregion
    }

    internal class FP : Button
    {
        public static FPOp m_FP;
        public Inventor.Document pDoc { get; set; }
        public static FPOp getFP
        {
            get
            {
                return m_FP;
            }
        }

        #region "Methods"
        public FP(string displayName, string internalName, string clientId, string description, string tooltip, Icon standardIcon, Icon largeIcon)
            : base(displayName, internalName, clientId, description, tooltip, standardIcon, largeIcon)
        {

        }

        protected override void ButtonDefinition_OnExecute(NameValueMap context)
        {
            //System.Windows.Forms.Application.Run(m_Break = new BreakOp(Inventor.Document oDoc));
            m_FP = new FPOp(InventorApplication.ActiveDocument, InventorApplication);
            //Params par = new Params((PartDocument)InventorApplication.ActiveDocument);
        }

        #endregion
    }

    //public class Params
    //{
    //    public Params(PartDocument doc)
    //    {
    //        PartComponentDefinition compDef = doc.ComponentDefinition;
    //        if (compDef.Parameters.ParameterTables.Count != 0)
    //        {
    //            foreach (ParameterTable table in compDef.Parameters.ParameterTables)
    //            {
    //                foreach (TableParameter param in table.TableParameters)
    //                {
    //                    compDef.Parameters.UserParameters.AddByValue(param.Name, param.Value, param.Units);
    //                }
    //            }
    //        }
    //    }
    //}

    public class PDFOp
    {
        private DrawingDocument m_Drw;
        private System.Collections.Generic.List<Document> drws;
        InvDocument<AssemblyDocument> invDoc;
        private PartDocument m_Prt;
        private TranslatorAddIn pdfAddIn;
        //private Inventor.Application invApp;
        private TranslationContext context;
        private NameValueMap nvm;
        private DataMedium dm;
        private double l;
        private string path, name, dat, Nizv, projectPath, docPath;
        EdgeLoop el; PlanarSketch ps;
        System.Collections.Generic.HashSet<Edge> edges = new System.Collections.Generic.HashSet<Edge>();
        System.Collections.Generic.HashSet<PunchToolFeature> punch = new System.Collections.Generic.HashSet<PunchToolFeature>();
        System.Collections.Generic.HashSet<EdgeLoop> loops = new System.Collections.Generic.HashSet<EdgeLoop>();
        private System.Collections.Generic.HashSet<string> pathes = new System.Collections.Generic.HashSet<string>();
        System.Collections.Generic.IEnumerable<string> en;
        private Document refDoc;
        private SheetMetalComponentDefinition compDef;
        private FlatPattern fp;
        private System.Collections.Generic.Dictionary<char,string> dic = new System.Collections.Generic.Dictionary<char,string>();
        private System.Collections.Generic.Dictionary<char, string> dic1 = new System.Collections.Generic.Dictionary<char, string>();
        private InvDocument<Document> m_Doc;

        public PDFOp()
        {
            iniDic();
        }

        public PDFOp(System.Collections.Generic.List<Document> docs)
        {
            iniPDF();
            foreach (Document doc in docs)
            {
                if (doc == null) continue;
                string pdfDir = doc.FullDocumentName, dxfDif = doc.FullDocumentName;
                createPath(ref pdfDir, "PDF"); createPath(ref dxfDif, "DXF");
                if (doc.DocumentType == DocumentTypeEnum.kDrawingDocumentObject)
                addPDF(doc, pdfDir, false);
                else if (doc.SubType == "{9C464203-9BAE-11D3-8BAD-0060B0CE6BB4}")
                    addDXF(doc, false, dxfDif, true);
            }
        }

        public PDFOp(Inventor.Document newDoc)                               
        {
            projectPath = ut.pathDoc(newDoc);                 
            createPath(ref projectPath, "PDF");
            iniPDF();

            foreach (Document doc in I.app.Documents.VisibleDocuments)
            {
                if (doc.DocumentType == DocumentTypeEnum.kDrawingDocumentObject)
                {
                    addPDF(doc, ut.pathDoc(doc, @"\PDF\"), false);
                }
                else if (doc.DocumentType == DocumentTypeEnum.kPartDocumentObject)
                {
                    addDXF(doc, false, ut.pathDoc(doc, @"\DXF\"));
                }
            }
        }

        private void iniPDF()
        {
            pdfAddIn = (TranslatorAddIn)Macros.StandardAddInServer.m_inventorApplication.ApplicationAddIns.ItemById["{0AC6FD96-2F4D-42CE-8BE0-8AEA580399E4}"];
            context = ADDIN.m_inventorApplication.TransientObjects.CreateTranslationContext();
            context.Type = IOMechanismEnum.kFileBrowseIOMechanism;
            nvm = ADDIN.m_inventorApplication.TransientObjects.CreateNameValueMap();
            dm = ADDIN.m_inventorApplication.TransientObjects.CreateDataMedium();
            iniDic();
        }

        public PDFOp(AssemblyDocument asmDoc)
        {
            I.path = ut.pathDoc(asmDoc as Document);
            iniDic();
            string p = file.p(asmDoc.FullFileName);
            IEnumerable<Document> docs = I.getDocs(asmDoc as Document, null);
            ut.action<Document>(docs, a => addDXF(a), f => file.p(f.FullDocumentName) == p); I.path = null;
        }

        private void copyFile(string name, string findName)
        {
            if (en == null)
            {
                foreach (var item in pathes)
                {
                    if (en == null) en = System.IO.Directory.EnumerateFiles(item);
                    else en = en.Concat(System.IO.Directory.EnumerateFiles(item));
                }
                en = en.OrderByDescending(s => System.IO.File.GetCreationTime(s));
            }
            string f = null;/* = en.FirstOrDefault(str => str.IndexOf(findName) != -1)*/;
            foreach (var s in en)
            {
                if (s.ToUpper().IndexOf(findName.ToUpper()) != -1)
                {
                    f = s; break;
                }
            }
            if (f != null && f != name)
            {
                System.IO.File.Copy(f, name, true);
            }
        }

        public string addPDF(Document doc, string path, bool copy)
        {
            if (doc.DocumentType == DocumentTypeEnum.kDrawingDocumentObject)
            {
                string lit, dec = "";
                refDoc = ut.referendedDoc(doc);//[doc.ReferencedDocuments.Count];
                try
                {
                    dec = refDoc.PropertySets[3][2].Value.ToString();
                    //if (dec.Length > 3 && dec[dec.Length - 3] == '-') 
                    //    dec = dec.Remove(dec.LastIndexOf('-'));
                }
                catch
                {
                    dec = "";
                }
                if (dec == "")
                    dec = refDoc.PropertySets[3][14].Value.ToString();
                lit = litera(refDoc);
                dat = refDoc.PropertySets[3][1].Value.ToString().Substring(0, refDoc.PropertySets[3][1].Value.ToString().LastIndexOf(' '));
                Nizv = ut.getPropValue(refDoc, "Изв");
                if (Nizv != "") 
                {
                    Nizv += "_";
                    if (Nizv.StartsWith("ТПМШ")) Nizv = Nizv.Substring(3);
                    dat = u.getPropValue(refDoc, "ИзвД");
                    Nizv = "_IZV_" + Nizv + u.getDate(dat);
                }
                if (refDoc.DocumentType == DocumentTypeEnum.kPartDocumentObject)
                {
                    name = lit + "EKD_" + dec + "_" + dat;
                    if (pdfAddIn.get_HasSaveCopyAsOptions(m_Drw, context, nvm))
                    {
                        nvm.set_Value("ScaleMode", PrintScaleModeEnum.kPrintFullScale);
                        nvm.set_Value("All_Color_AS_Black", 0);
                        nvm.set_Value("Remove_Line_Weights", 0);
                        nvm.set_Value("Vector_Resolution", 600);
                        nvm.set_Value("Sheet_Range", PrintRangeEnum.kPrintAllSheets);
                    }
                }
                else
                {
                    name = lit + "EKD_" + dec + "_SB_" + dat;
                    if (pdfAddIn.get_HasSaveCopyAsOptions(m_Drw, context, nvm))
                    {
                        nvm.set_Value("ScaleMode", PrintScaleModeEnum.kPrintFullScale);
                        Property p = InvDoc.u.getProp(doc, "thick");
                        if (p == null)
                            nvm.set_Value("Remove_Line_Weights", 0);
                        else nvm.set_Value("Remove_Line_Weights", 1);
                        nvm.set_Value("All_Color_AS_Black", 0);
                        nvm.set_Value("Vector_Resolution", 600);
                        nvm.set_Value("Sheet_Range", PrintRangeEnum.kPrintAllSheets);
                    }
                }
                string prop = u.getPropValue(refDoc, "Catalog web link");

                if (dec != "")
                    name = translit(name);
                Nizv = forDxf(Nizv);
                if (prop == "")
                {
                    dm.FileName = path + name + Nizv + ".pdf";
                }
                else
                {
                    name = u.nameUtil(doc);
                    dm.FileName = path + name + ".pdf";
                }
                string n = dm.FileName;
                System.TimeSpan ts = System.DateTime.Now - System.IO.File.GetCreationTime(n);
                if (System.IO.File.Exists(n) && ts.Minutes < 2) return n;
                if (copy)
                {
                    //if (!System.IO.Directory.Exists(docPath + "\\PDF\\")) System.IO.Directory.CreateDirectory(docPath + "\\PDF\\");
                    string p = doc.FullDocumentName.Remove(doc.FullDocumentName.LastIndexOf('\\'));
                    if (invDoc.path != p + "\\"/* && System.IO.Directory.Exists(p + "\\PDF\\")*/)
                    {
//                         if (System.IO.File.Exists(p + "\\PDF\\" + name + ".pdf"))
//                             System.IO.File.Copy(p + "\\PDF\\" + name + ".pdf", path + name + ".pdf", true);
//                         else pdfAddIn.SaveCopyAs(doc, context, nvm, dm);
                        copyFile(dm.FileName, translit(dec));
                    }
                    else pdfAddIn.SaveCopyAs(doc, context, nvm, dm);
                }
                else
                    pdfAddIn.SaveCopyAs(doc, context, nvm, dm);
                refDoc.Dirty = false;
                refDoc.ReleaseReference();
                I.silent(true);
                doc.Save();
                I.silent(false);
                return dm.FileName;
                //doc.Dirty = false;
            }
            return null;
        }

        public static void createPath(ref string path, string val)
        {
            if (System.IO.Path.GetExtension(path) != "") path = System.IO.Path.GetDirectoryName(path);
            path += "\\" + val + "\\";
            if (System.IO.Directory.Exists(path)) return;
            else System.IO.Directory.CreateDirectory(path);
        }

        public string addDXF(Document doc, bool copy = false, string path = "", bool my = true)
        {
            if (path == "")
            {
                path = ut.createDir(doc, @"\DXF\");
            }
            SheetMetalComponentDefinition smcd = I.getSMCD(doc);
            if (smcd == null || !smcd.HasFlatPattern) return null;
            fp = smcd.FlatPattern; I.silent(true);
            m_Prt = doc as PartDocument;
                    string reviz = "";
                    string name = MatName(smcd, doc, out reviz);
                    string rev = doc.PropertySets[1][7].Value.ToString();
                    if (rev.EndsWith(" "))
                    {
                        name = rev.Substring(0, rev.Length - 1);
                    }
                    else if (doc.PropertySets[1][7].Value.ToString() != reviz)
                    {
                        doc.PropertySets[1][7].Value = reviz;
                        doc.Save();
                    }

                        string n = path + name;
                        System.TimeSpan ts = System.DateTime.Now - System.IO.File.GetCreationTime(n);
                        if (System.IO.File.Exists(n) && ts.Minutes < 2) return n;
                            if (copy)
                            {
                                string p = doc.FullDocumentName.Remove(doc.FullDocumentName.LastIndexOf('\\'));
                                if (invDoc.path != p + "\\"/* && System.IO.Directory.Exists(p + "\\DXF\\"*/)
                                {
                                    copyFile(path + name, name.Substring(name.IndexOf('_')));
                                }
                                else if(my) FPOp.createDXF(fp, path + name);
                                else
                                {
                                    DataIO Io = (doc as PartDocument).ComponentDefinition.DataIO;
                                    string sOut = "FLAT PATTERN DXF?AcadVersion=2000&OuterProfileLayer=0&InteriorProfilesLayer=0" +
                                    "&FeatureProfileLayer=1" + "&UnconsumedSketchesLayer=0" +
                                    "&SimplifySplines=True&SplineTolerance=0,1&RebaseGeometry=True&MergeProfilesIntoPolyline=False" +
                                    "&InvisibleLayers=IV_TANGENT;IV_BEND;IV_BEND_DOWN;IV_TOOL_CENTER;IV_TOOL_CENTER_DOWN;IV_ARC_CENTERS;IV_FEATURE_PROFILES_DOWN";
                                    Io.WriteDataToFile(sOut, path + name);}
                            }
                            else if (my) FPOp.createDXF(fp, path + name);
                            else {
                                DataIO Io = (doc as PartDocument).ComponentDefinition.DataIO;
                                string sOut = "FLAT PATTERN DXF?AcadVersion=2000&OuterProfileLayer=0&InteriorProfilesLayer=0" +
                                "&FeatureProfileLayer=1" + "&UnconsumedSketchesLayer=0" +
                                "&SimplifySplines=True&SplineTolerance=0,1&RebaseGeometry=True&MergeProfilesIntoPolyline=False" +
                                "&InvisibleLayers=IV_TANGENT;IV_BEND;IV_BEND_DOWN;IV_TOOL_CENTER;IV_TOOL_CENTER_DOWN;IV_ARC_CENTERS;IV_FEATURE_PROFILES_DOWN"; 
                                Io.WriteDataToFile(sOut, path + name);
                            }
                    m_Prt.Dirty = false;
                    fp = null;
                    I.silent(false);
                    return path + name;
        }

        public void undo()
        {
            foreach (PunchToolFeature item in punch)
            {
                item.Suppressed = false;
            }
            fp.Edit();
            fp.Sketches["fordxf"].Delete();
            fp.ExitEdit();
        }

        public bool findEdges(FlatPattern fp, double l)
        {
            bool flag = false;
            foreach (Edge ed in fp.TopFace.Edges)
            {
                if (u.getLenght(ed) == l / 10)
                {
                    flag = true;
                    if (edges.FirstOrDefault(e => e.Equals(ed)) != null) continue;
                    foreach (EdgeUse item in ed.EdgeUses)
                    {
                        if (u.eq(fp.TopFace,item.EdgeLoop.Edges[1].StartVertex.Point) && u.eq(fp.TopFace,item.EdgeLoop.Edges[3].StartVertex.Point))
                        {
                            el = item.EdgeLoop;
                            foreach (Edge edg in el.Edges)
                            {
                                edges.Add(edg);
                            }
                        }
                    }
                    if (el != null) loops.Add(el);
                }
            }
            return flag;
        }

        public void cmpFP(Document doc, double l)
        {
            if (doc.SubType == "{9C464203-9BAE-11D3-8BAD-0060B0CE6BB4}")
            {
                SheetMetalFeatures smf = compDef.Features as SheetMetalFeatures;
//                 foreach (PunchToolFeature item in smf.PunchToolFeatures)
//                 {
//                     if (item.iFeatureTemplateDescriptor.LastKnownSourceFileName.IndexOf("ПодПровод") != -1)
//                     {
//                         punch.Add(item);
//                     }
//                 }
//                 if (punch.Count == 0) return;
                compDef = (doc as PartDocument).ComponentDefinition as SheetMetalComponentDefinition;

                fp.Edit();
                
                fp.ExitEdit();
                foreach (PunchToolFeature item in punch)
                {
                    item.Suppressed = true;
                }
                doc.Update2();
            }
        }

        public void addToSketch()
        {
            Macros.StandardAddInServer.m_inventorApplication.ScreenUpdating = false;
            fp.Edit();
            ps = ps ?? fp.Sketches.Add(fp.TopFace);
            PlanarSketch psForExtr = fp.Sketches.Add(fp.TopFace);
            if (Macros.StandardAddInServer.m_inventorApplication.ActiveEditObject is PlanarSketch) { }
            else
                ps.Edit();
            foreach (EdgeLoop el in loops)
            {
//                 int maxcount = el.Edges.Count - 1, mincount = el.Edges.Count / 2;
//                 if (el.Edges.Count == 4) {mincount++; maxcount++;}
//                 for (int i = mincount; i < maxcount; i++)
//                 {
//                     SketchEntity se = ps.AddByProjectingEntity(el.Edges[i]);
//                     //se.Reference = false;
//                 }
                foreach (Edge item in getEdges(el))
                {
                    SketchEntity se = ps.AddByProjectingEntity(item);
                }
            }
            ps.ExitEdit();
            psForExtr.Edit();
            foreach (EdgeLoop el in loops)
            {
                foreach (Edge e in el.Edges)
                {
                    psForExtr.AddByProjectingEntity(e); 
                }
            }
            Profile pr = psForExtr.Profiles.AddForSolid();
            ExtrudeDefinition edef = fp.Features.ExtrudeFeatures.CreateExtrudeDefinition(pr, PartFeatureOperationEnum.kJoinOperation);
            edef.SetDistanceExtent(compDef.Thickness, PartFeatureExtentDirectionEnum.kNegativeExtentDirection);
            fp.Features.ExtrudeFeatures.Add(edef);
            if (ps.Name != "fordxf") ps.Name = "fordxf";
            psForExtr.ExitEdit();
            fp.ExitEdit();
            Macros.StandardAddInServer.m_inventorApplication.ScreenUpdating = true;
        }

        public System.Collections.Generic.HashSet<Edge> getEdges(EdgeLoop el)
        {
            System.Collections.Generic.HashSet<Edge> edges = new System.Collections.Generic.HashSet<Edge>();
            int min = -1, max = -1;
            for (int i = 1; i <= el.Edges.Count; i++)
            {
                if (u.eq(u.getLenght(el.Edges[i]), l/10))
                {
                    if (min == -1) min = i+1;
                    else max = i;
                }
            }
            for (int i = min; i < max; i++)
            {
                edges.Add(el.Edges[i]);
            }
            return edges;
        }

        public void iniDic()
        {
            dic.Add('а', "a"); dic.Add('б', "b"); dic.Add('в', "v"); dic.Add('г', "g"); dic.Add('д', "d"); dic.Add('е', "e"); dic.Add('ё', "jo"); dic.Add('ж', "zh");
            dic.Add('з', "z"); dic.Add('и', "i"); dic.Add('й', "jj"); dic.Add('к', "k"); dic.Add('л', "l"); dic.Add('м', "m"); dic.Add('н', "n"); dic.Add('о', "o"); dic.Add('п', "p");
            dic.Add('р', "r"); dic.Add('с', "s"); dic.Add('т', "t"); dic.Add('у', "u"); dic.Add('ф', "f"); dic.Add('х', "h"); dic.Add('ц', "c"); dic.Add('ч', "ch");
            dic.Add('ш', "sh"); dic.Add('щ', "zch"); dic.Add('ъ', "''"); dic.Add('ы', "'y"); dic.Add('ь', "'"); dic.Add('э', "e"); dic.Add('ю', "ju"); dic.Add('я', "ja");
            dic.Add('А', "a"); dic.Add('Б', "b"); dic.Add('В', "v"); dic.Add('Г', "g"); dic.Add('Д', "d"); dic.Add('Е', "e"); dic.Add('Ё', "jo"); dic.Add('Ж', "zh");
            dic.Add('З', "z"); dic.Add('И', "i"); dic.Add('Й', "jj"); dic.Add('К', "k"); dic.Add('Л', "l"); dic.Add('М', "m"); dic.Add('Н', "n"); dic.Add('О', "o"); dic.Add('П', "p");
            dic.Add('Р', "r"); dic.Add('С', "s"); dic.Add('Т', "t"); dic.Add('У', "u"); dic.Add('Ф', "f"); dic.Add('Х', "h"); dic.Add('Ц', "c"); dic.Add('Ч', "ch");
            dic.Add('Ш', "sh"); dic.Add('Щ', "zch"); dic.Add('Ъ', "''"); dic.Add('Ы', "'y"); dic.Add('Ь', "'"); dic.Add('Э', "e"); dic.Add('Ю', "ju"); dic.Add('Я', "ja");
            dic.Add(' ', "_"); dic.Add('.', "_"); dic.Add('-', "_");
            dic1.Add(' ', "_"); dic1.Add('.', "_"); dic1.Add('-', "_");
        }

        public string translit (string str)
        {
            string ret = "";
            
            foreach (char ch in str)
            {
                if (dic.ContainsKey(ch)) ret += dic[ch];
                else ret += ch;
            }
            return ret.ToUpper();
        }

        public string forDxf(string str)
        {
            string ret = "";
            //if (dic1)
            foreach (char ch in str)
            {
                if (dic1.ContainsKey(ch)) ret += dic1[ch];
                else ret += ch;
//                 try
//                 {
//                     ret += dic1[ch];
//                 }
//                 catch 
//                 {
//                     ret += ch;
//                 }
            }
            return ret;
        }

        public string litera(Document doc)
        {
            string l = ut.getPropValue(doc, "Литера1") + ut.getPropValue(doc, "Литера2") + "_";
            if (l == "_") return "";
            return l;
        }

        public string izve(Document doc)
        {
            string l = "_" + ut.getPropValue(doc, "Изв") + "_" + ut.getPropValue(doc, "ИзвД");
            if (l == "__") return "";
            return l;
        }

        public string MatName(SheetMetalComponentDefinition m_CompDef, Document doc, out string rev)
        {
            string MatLine = "", MatUpLine = "", MatDownLine = "", MatCenter = "", DXF = "", name1 = ""; int ind = 0; int count = 0;
            string lit = "", izv = "", path = ut.pathDoc(doc), fn = "", _t;
            if (System.IO.File.Exists(path + "\\Material.xml"))
            {
                fn = path + "\\Material.xml"; 
            } 
            InvDoc.XML XMLMaterial = new InvDoc.XML(fn);
            double t = (double)m_CompDef.Thickness.Value * 10;
            XMLMaterial.ReadXML(m_CompDef.Material.Name, ref MatLine, ref MatUpLine, ref MatDownLine, ref MatCenter, ref DXF);
            string name = doc.PropertySets[3][2].Value.ToString();
            string isp = "";
            if (name[name.Length - 3] == '-')
            {
                isp = name.Substring(name.Length - 2, 2);
            }
            if (name.IndexOf('.') != -1)
            {
                name = name.Substring(0, name.IndexOf('.'));
                if (name.IndexOf("КЭВ") == -1) name = name.Replace('-', '_');
                ind = name.IndexOf('-');
                if (ind != -1)
                {
                    name = name.Substring(ind+1, name.Length - ind-1);
                }
            }
            lit = litera(doc);
            izv = izve(doc);
            name1 = lit + name + "_" + doc.PropertySets[3][14].Value.ToString();
            if (isp != "")
            {
                name1 = name1 + "_" + isp;
            }
            _t = t.ToString();
            string p = ut.getPropValue(doc, "dxf");
            if (p != "")
            {
                _t += "_" + p;
            }

//             if (DXF[1] < 256)
//             {
//                 //DXF = "Galv.Steel";
//                 _t = (t / 25.4).ToString("0.000") + "_" + t + "_mm";
//             }
            name1 = forDxf(name1);
            name = name1 + "_" + DXF  + "_" + _t;
            rev = name1 + "_" + DXF + "_" + _t + izv + ".dxf";

            for (int i = 0; i < ut.convToInt(ut.getPropValue(doc, "CountDXF")); i++)
            {
                rev = rev + "    " + forDxf(name1) + "_" + (i + 1).ToString("00") + " " + DXF + "_" + _t + izv + ".dxf"; 
            }

            name = forDxf(name);
            return name + izv + ".dxf";
        }
    }

    public struct rect
    {
        public double x;
        public string fn;

        public rect(string str, double p1)
        {
            x = p1;
            fn = str;
        }
    }

    public class PrintOp
    {
        public DrawingPrintManager prntMgr;             
        public Sheet oSheet, oldSheet;
        private double w, h;
        string path;
        NameValueMap nvm;
        System.Collections.Generic.List<rect> ss = new System.Collections.Generic.List<rect>();
        public PrintOp(Document doc, Inventor.Application invApp)
        {
            if (Macros.StandardAddInServer.Printflag == "1")
            {
                //Inventor.ApprenticeServerComponent appServer;
                //Type appType = Type.GetTypeFromProgID("Inventor.ApprenticeServer");
                //appServer = (Inventor.ApprenticeServerComponent)Activator.CreateInstance(appType);
                OpenFileDialog ofd = new OpenFileDialog();
                //ofd.Filter = "Чертежи|*.idw";
                ofd.Title = "Выберите файлы";
                ofd.Multiselect = true;

                path = doc.FullFileName.ToString();
                path = path.Substring(0, path.LastIndexOf('\\'));
                path += "\\";

                ofd.InitialDirectory = path;
                ofd.ShowDialog();
                invApp.SilentOperation = true;
                nvm = invApp.TransientObjects.CreateNameValueMap();
                nvm.Add("SkipAllUnresolvedFiles", true);
                nvm.Add("DeferUpdates", true);
                foreach (string fn in ofd.SafeFileNames)
                {
                    //InvDocument<Inventor.Document> Docs = new InvDocument<Document>(doc);
                    //DrawingDocument drwDoc = (DrawingDocument)Docs.addDoc(DocumentTypeEnum.kDrawingDocumentObject, TypeDocs.DrawingDoc);
                    //AssemblyDocument asmDoc = (AssemblyDocument)Docs.addDoc(DocumentTypeEnum.kAssemblyDocumentObject, TypeDocs.AssemblyDoc);
                    //PartDocument prtDoc = (PartDocument)Docs.derivedDoc("Экран.ipt");

                    Inventor.Document oDoc = invApp.Documents.OpenWithOptions(path + fn, nvm, false);
                    string fiilename = path + fn;
                    //ApprenticeServerDrawingDocument oDoc = invApp.Open(fiilename);
                    if (oDoc.DocumentType != DocumentTypeEnum.kDrawingDocumentObject) continue;
                    foreach (Inventor.Sheet s in ((DrawingDocument)oDoc).Sheets)
                    {
                        ss.Add(new rect(fn, s.Width * 10));
                    }
                    //oDoc.Close();
                }
                ss.Sort(delegate(rect r1, rect r2) { return r1.x <= r2.x ? 1 : -1; });

                System.IO.StreamWriter fs = new System.IO.StreamWriter(path + "Размеры чертежей.txt", false);
                foreach (rect str in ss)    
                {
                    fs.WriteLine(str.x.ToString()+ "\t" + str.fn);
                }
                fs.Close();
                invApp.SilentOperation = false;
            }
            else
            {
                foreach (Inventor.Document actdoc in invApp.Documents.VisibleDocuments)
                {
                    if (actdoc.DocumentType == DocumentTypeEnum.kDrawingDocumentObject)
                    {
                        prntMgr = (DrawingPrintManager)actdoc.PrintManager;
                        prntMgr.Printer = "\\\\kb4\\HP LaserJet 700 M712 UPD PCL 6";
                        actdoc.Activate();
                        oldSheet = ((DrawingDocument)actdoc).ActiveSheet;
                        foreach (Sheet sheet in ((DrawingDocument)actdoc).Sheets)
                        {
                            oSheet = sheet;
                            print();
                        }
                        oldSheet.Activate();
                        actdoc.Dirty = false;
                    }
                }
            }
        }

        public void print(int num = 1)
        {
            oSheet.Activate();
            w = oSheet.Width;
            h = oSheet.Height;
            prntMgr.ColorMode = PrintColorModeEnum.kPrintColorPalette;
            prntMgr.NumberOfCopies = num;
            prntMgr.ScaleMode = PrintScaleModeEnum.kPrintFullScale;
            if (w == 21)
            {
                prntMgr.PaperSize = PaperSizeEnum.kPaperSizeA4;
                prntMgr.Orientation = PrintOrientationEnum.kPortraitOrientation;
                prntMgr.SubmitPrint();
            }
            else if (w == 42)
            {
                prntMgr.PaperSize = PaperSizeEnum.kPaperSizeA3;
                prntMgr.Orientation = PrintOrientationEnum.kLandscapeOrientation;
                prntMgr.SubmitPrint();
            }
            else if ((int)(h*1) == 297 && w > 42 && w <= 601)
            {
                prntMgr.Printer = "HP Designjet 510 24in Printer HPGL2_1";
                prntMgr.PaperSize = PaperSizeEnum.kPaperSizeCustom;
                prntMgr.PaperHeight = 29.7;
                prntMgr.PaperWidth = w;
                prntMgr.Orientation = PrintOrientationEnum.kLandscapeOrientation;
            }
        }
    }
    public class FPOp
    {
        SheetMetalComponentDefinition sheetMetal;
        FlatPattern fp; int dir = 0; double l, w; UnitVector v; PartFeature pf; DialogResult dr;
        object obj; AlignmentTypeEnum al; bool re;
        int dir2 = 0;
        PlanarSketch sketch;
        SketchEntity se; SketchPoint sp;
        public FPOp(Document doc, Inventor.Application invApp)
        {
            if (doc.SubType == "{9C464203-9BAE-11D3-8BAD-0060B0CE6BB4}")
            {
                sheetMetal = (SheetMetalComponentDefinition)((PartDocument)doc).ComponentDefinition;
                if (sheetMetal.HasFlatPattern == false)
                {
                    Face f = (Face)invApp.CommandManager.Pick(SelectionFilterEnum.kPartFacePlanarFilter, "Выберите переднюю грань");
                    sheetMetal.Unfold2(f);
                    sheetMetal = (SheetMetalComponentDefinition)((PartDocument)doc).ComponentDefinition;
                    SheetMetalFeatures smf = (SheetMetalFeatures)sheetMetal.Features;
                    double  w = 0;
                    //bool flag = false;

                    pf = sheetMetal.SurfaceBodies[1].CreatedByFeature;
                    if (pf.Type == ObjectTypeEnum.kReferenceFeatureObject)
                    {
                        pf = ((SurfaceBody)((ReferenceFeature)pf).ReferencedEntity).CreatedByFeature;
                    }

                    if (pf.Type == ObjectTypeEnum.kFaceFeatureObject)
                    {
                        FaceFeature ff = (FaceFeature)pf;
                        w = ff.RangeBox.MaxPoint.X - ff.RangeBox.MinPoint.X;
                        //h = ff.RangeBox.MaxPoint.Y - ff.RangeBox.MinPoint.Y;
                    }
                    else if (pf.Type == ObjectTypeEnum.kContourFlangeFeatureObject)
                    {
                        ContourFlangeFeature cff = (ContourFlangeFeature)pf;
                        v = ((PlanarSketch)((SketchEntity)cff.Definition.Path[1].SketchEntity).Parent).PlanarEntityGeometry.Normal;
                        w = (double)((DistanceExtent)cff.Definition.DefaultWidthExtent).Distance.Value;
                    }
                    
                    fp = sheetMetal.FlatPattern;
                    fp.GetAlignment(out al, out obj, out re);
                    if (Math.Abs((fp.Length - w) / fp.Length) > 0.1)
                    {
                        //object obj; AlignmentTypeEnum al; bool re; 
                        fp.GetAlignment(out al, out obj, out re);
                        if (obj == null)
                        {
                            Edge ed = fp.TopFace.Edges.OfType<Edge>().FirstOrDefault(e => e.GeometryType == CurveTypeEnum.kLineSegmentCurve && ((LineSegment)e.Geometry).Direction.Y == -1);
                            if (ed != null)
                            {
                                obj = ed;     
                            }
                            al = AlignmentTypeEnum.kHorizontalAlignment;
                            if (v != null)
                            {
                                if (v.Z != 0) al = AlignmentTypeEnum.kVerticalAlignment;
                            }
                            fp.SetAlignment(al, obj, false);
                            fp.GetAlignment(out al, out obj, out re);
                            invApp.CommandManager.ControlDefinitions["AppZoomAllCmd"].Execute();
                            dr = MessageBox.Show("Поверуть на 180 градусов - Да, на 90 - Нет, оставить - Отмена", "Развертка", MessageBoxButtons.YesNoCancel);
                            if (dr == DialogResult.Yes)
                            {
                                fp.GetAlignment(out al, out obj, out re);
                                fp.SetAlignment(al, obj, !re);
                            }
                            else if (dr == DialogResult.No)
                            {

                                al = (al == AlignmentTypeEnum.kHorizontalAlignment)?AlignmentTypeEnum.kVerticalAlignment: AlignmentTypeEnum.kHorizontalAlignment;
                                fp.SetAlignment(al, obj, re);
                            }
                            fp.ExitEdit();
                            return;
                        }
                        if (al == AlignmentTypeEnum.kHorizontalAlignment)
                        {
                            fp.SetAlignment(AlignmentTypeEnum.kVerticalAlignment, obj, false);
                        }
                        else fp.SetAlignment(AlignmentTypeEnum.kHorizontalAlignment, obj, false);
                    }
                    else if (obj != null && ((Edge)obj).GeometryType != CurveTypeEnum.kLineSegmentCurve)
                    {
                        Edge ed = fp.TopFace.Edges.OfType<Edge>().FirstOrDefault(e => e.GeometryType == CurveTypeEnum.kLineSegmentCurve && ((LineSegment)e.Geometry).Direction.Y == -1);
                        if (ed != null)
                        {
                            obj = ed;
                        }
                        al = AlignmentTypeEnum.kVerticalAlignment;
                        fp.SetAlignment(al, obj, false);
                    }
                    invApp.CommandManager.ControlDefinitions["AppZoomAllCmd"].Execute();
                    dr = MessageBox.Show("Поверуть на 180 градусов - Да, на 90 - Нет, оставить - Отмена", "Развертка", MessageBoxButtons.YesNoCancel);
                    if (dr == DialogResult.Yes)
                    {
                        fp.GetAlignment(out al, out obj, out re);
                        fp.SetAlignment(al, obj, !re);
                    }
                    else if (dr == DialogResult.No)
                    {
                        fp.GetAlignment(out al, out obj, out re);
                        al = (al == AlignmentTypeEnum.kHorizontalAlignment) ? AlignmentTypeEnum.kVerticalAlignment : AlignmentTypeEnum.kHorizontalAlignment;
                        fp.SetAlignment(al, obj, re);
                    }
                    fp.ExitEdit();
                    return;
                }
                else
                {
                    flatCut(sheetMetal.FlatPattern,sheetMetal);
                    return;
                }
            }
        }

        public void flatCut(FlatPattern fp, SheetMetalComponentDefinition compDef)
        {
            sketch = fp.Sketches.Add(fp.TopFace);
            Inventor.Application invApp = Macros.StandardAddInServer.m_inventorApplication;
            string path = invApp.iFeatureOptions.RootPath;
            ObjectCollection col = invApp.TransientObjects.CreateObjectCollection();
            foreach (FlatBendResult fbr in fp.FlatBendResults)
            {
                if (fbr.IsOnBottomFace == false && fbr.InnerRadius < 0.2)
                {
                    if (fbr.Edge != null)
                    {
                        sp = (SketchPoint)sketch.AddByProjectingEntity(fbr.Edge.StartVertex);
                        col.Add(sp);
                        sp = (SketchPoint)sketch.AddByProjectingEntity(fbr.Edge.StopVertex);
                        col.Add(sp);
                    }
                }
            }
            iFeatureDefinition ifd = fp.Features.PunchToolFeatures.CreateiFeatureDefinition(path + "ПодСгиб.ide");
            PunchToolFeature ptf = fp.Features.PunchToolFeatures.Add(col, ifd, 0);
        }



        static public object flatCut(SheetMetalFeatures smf, PlanarSketch ps, string name, string a = "5",  string b = "8")
        {
            Inventor.Application invApp = Macros.StandardAddInServer.m_inventorApplication;
            //SketchPoint sp = null;
            string path = invApp.iFeatureOptions.RootPath;
            ObjectCollection col = invApp.TransientObjects.CreateObjectCollection();
            foreach (SketchPoint item in ps.SketchPoints)
            {
                if (item.HoleCenter == true)
                col.Add(item);
            }
            string iFN = path + "Punches\\" + name + ".ide";
            if (System.IO.File.Exists(iFN))
            {
                iFeatureDefinition ifd = smf.PunchToolFeatures.CreateiFeatureDefinition(iFN);
                iFeatureParameterInput inp = ifd.iFeatureInputs[1] as iFeatureParameterInput;
                inp.Expression = a;
                inp = ifd.iFeatureInputs[2] as iFeatureParameterInput;
                inp.Expression = b;
                PunchToolFeature ptf = smf.PunchToolFeatures.Add(col, ifd, 0);
                return ptf;
            }
            return null;    
        }
        static public void createDXF(FlatPattern fp, string p)
        {
            double m = 10;
            //Inventor.Point origin = fp.RangeBox.MinPoint;
            //dxf.vector v = new dxf.vector(origin.X*m, origin.Y*m);
            dxf.DXF dXf = new dxf.DXF();
            dxf.DXF.tol = 5;
            ut.action<EdgeLoop>(fp.TopFace.EdgeLoops, e => createDXF(e, dXf, m));
            ut.action<PlanarSketch>(fp.Sketches, e => createDXF(e, dXf, m), f => !f.Consumed);
            dXf.bounds();
            dxf.vector v = new dxf.vector(-dXf.gab.l, -dXf.gab.b);
            dXf.move(v);
            dXf.addFileName(p);
            dXf.safe("");
        }
        static public void createDXF(EdgeLoop el, dxf.DXF dxf, double m = 10)
        {
            if (el.Edges.Count == 2)
            {
                BSplineCurve spl1 = el.Edges[1].Geometry as BSplineCurve, spl2 = el.Edges[2].Geometry as BSplineCurve;
                bool flag = false;
                if (spl1 != null && spl2 != null)
                {
                    flag = addCircle(dxf, spl1, spl2);
                }
                if (flag) return;
            }
            IEnumerable<Edge> col = ut.gets<Edge>(el.Edges, f => f.GeometryType == CurveTypeEnum.kLineSegmentCurve && ut.getLenght(f) == 0.01);
            if (col.Count() == 2)
            {
                Edge e1 = col.ElementAt(0), e2 = col.ElementAt(1);
                bool flag = false;
                col = ut.gets<Edge>(el.Edges, f =>
                {
                    if (f.Equals(e1) || f.Equals(e2)) flag = !flag;
                    return flag;
                });
                ut.action<Edge>(col, ed => createDXF(ed, dxf), f => !(f.Equals(e1) || f.Equals(e2)));
                return;
            }
            foreach (Edge e in el.Edges)
            {
                createDXF(e, dxf);
            }
        }
        static public void createDXF(PlanarSketch s, dxf.DXF dxf, double m = 10)
        {
            foreach (SketchEntity ent in s.SketchEntities)
            {
                if (ent.Construction) continue;
                SketchLine l = ent as SketchLine;
                Inventor.Point sp, ep, cen;
                if (l != null)
                {
                    sp = l.StartSketchPoint.Geometry3d; ep = l.EndSketchPoint.Geometry3d;    
                    dxf.addLine(new dxf.coord(sp.X * m, sp.Y * m),
                        new dxf.coord(ep.X * m, ep.Y * m));
                }
                SketchCircle c = ent as SketchCircle;
                if (c != null)
                {
                    cen = c.CenterSketchPoint.Geometry3d;
                    dxf.addCircle(new dxf.coord(cen.X * m, cen.Y * m), c.Radius * m);
                }
                SketchArc a = ent as SketchArc;
                if (a != null)
                {
                    sp = a.StartSketchPoint.Geometry3d; ep = a.EndSketchPoint.Geometry3d; cen = a.CenterSketchPoint.Geometry3d;
                    if (a.Geometry3d.Normal.Z < 0)
                    {
                        sp = ep; ep = a.StartSketchPoint.Geometry3d;
                    }

                    addArc(dxf, sp, ep, cen, a.Radius);
                    //double [] param = ut.getParamExtents(a.Geometry.Evaluator);
                    //double [] pts = ut.getPointAtParam(a.Geometry.Evaluator, new []{(param[1]+param[0])/2});
                    //addArc(dxf, sp, ep, cen, a.Radius, pts);
                }
            }
        }
        static public void addArc(dxf.DXF dxf,Inventor.Point sp, Inventor.Point ep, Inventor.Point cen, double R, double [] pts, double m = 10)
        {
            double sa = ut.angFromPoint(ep, cen, m),
                   ea = ut.angFromPoint(sp, cen, m);
            double an =   (ea + sa) / 2;
            if (an > 180) an -= 180;
            double[] evalPts = pt(R, an, I.CP2d(cen.X,cen.Y));
            double dist = I.CP2d(pts[0], pts[1]).DistanceTo(I.CP2d(evalPts[0], evalPts[1]));
            //if (!(ut.eq(pts[0], evalPts[0]) && ut.eq(pts[1], evalPts[1])))
            if (sa > ea && dist < R/2)
                dxf.addArc(new dxf.coord(cen.X * m, cen.Y * m), R * m, sa, ea);
            else
                dxf.addArc(new dxf.coord(cen.X * m, cen.Y * m), R * m, ea, sa);
        }
        static public void addArc(dxf.DXF dxf, Inventor.Point sp, Inventor.Point ep, Inventor.Point cen, double R, double m = 10)
        {
            dxf.addArc(new dxf.coord(sp.X * m, sp.Y * m), new dxf.coord(ep.X * m, ep.Y * m), new dxf.coord(cen.X * m, cen.Y * m), R*m);
        }
        static public bool addCircle(dxf.DXF dxf, BSplineCurve spl1, BSplineCurve spl2, double m = 10)
        {
            List<Inventor.Point> pts1 = ut.getPointAtParam(spl1.Evaluator), pts2 = ut.getPointAtParam(spl2.Evaluator);
            if (ut.eq(pts1[0], pts2[1]) || ut.eq(pts1[1], pts2[0]))
            {
                Inventor.Point mPt = ut.midPt(pts1[2], pts2[2]);
                double R = ut.round(pts1[2].VectorTo(pts2[2]).Length,2)/2;
                dxf.addCircle(new dxf.coord(mPt.X * m, mPt.Y * m), R * m);
                return true;
            }
            return false;
        }
        static public void createDXF(Edge e, dxf.DXF dxf, double m = 10)
        {
            Inventor.Point sp, ep, cen; 
            Circle c = e.Geometry as Circle;
            if (c != null)
            {
               dxf.addCircle(new dxf.coord(c.Center.X*m,c.Center.Y*m), c.Radius*m); 
            }
            LineSegment l = e.Geometry as LineSegment;
            if (l != null)
            {
                dxf.addLine(new dxf.coord(l.StartPoint.X*m, l.StartPoint.Y*m), new dxf.coord(l.EndPoint.X*m, l.EndPoint.Y*m));
            }
            Arc3d a = e.Geometry as Arc3d;
            if (a != null)
            {
                sp = a.StartPoint; ep = a.EndPoint; cen = a.Center;
                if (a.Normal.Z < 0)
                {
                    sp = ep; ep = a.StartPoint;
                }

                addArc(dxf, sp, ep, cen, a.Radius);
            }
            BSplineCurve b = e.Geometry as BSplineCurve;
            if (b != null)
            {
                double min, max;
                double[] poles = null, knots = null, w = null, pv = new double[3];
                int ord, numP, numK;
                bool isR, isP, isC, isPl;
                b.GetBSplineInfo(out ord, out numP, out numK, out isR, out isP, out isC, out isPl, ref pv);
                //poles = new double[numP]; knots = new double[numK]; w = new double[numK];
                //b.GetBSplineData(ref poles, ref knots, ref w);
                b.Evaluator.GetParamExtents(out min, out max);
                double [] pts = new double [9]; double [] param = new double [] {min,max,(max+min)/2};
                b.Evaluator.GetPointAtParam(ref param, ref pts);
                double tol = 0.5;
                if (check(pts))
                    dxf.addLine(new dxf.coord(pts[0]*m, pts[1]*m), new dxf.coord(pts[3]*m, pts[4]*m));
                else if (ut.eq(pts[0],pts[3]) && ut.eq(pts[1],pts[4]))
                {
                    
                    double step = (max - min) / 4;
                    //double[] spl = new double[] { min, step, step*2, step*3};
                    //double[] splPts = ut.getPointAtParam(b.Evaluator, spl);
                    List<Inventor.Point> ptsInv = ut.getPoleAtIndex(b,ord, numP);
                    double l1 = ptsInv[0].DistanceTo(ptsInv[2]), l2 = ptsInv[1].DistanceTo(ptsInv[3]);

                    if (Math.Abs(l2 - l1) < tol)
                    {
                        Inventor.Point mPt = ut.midPt(ptsInv[0], ptsInv[2]);
                        dxf.addCircle(new dxf.coord(mPt.X * m, mPt.Y * m), ut.round(l2, 2)/2 * m); 
                    }
                }
                else if (!isC)
                {
                    List<Inventor.Point> ptsI = ut.getPointAtArray(pts);
                    Inventor.Point mPt = ut.midPt(ptsI[0], ptsI[1]);
                    double l1 = ptsI[0].DistanceTo(ptsI[1]), l2 = mPt.DistanceTo(ptsI[2]);
                    if (Math.Abs(l1 / 2 - l2) < tol)
                    {
                        Vector v = ut.CrossProduct(ptsI[0], ptsI[2], mPt);
                        if (ut.isNullVector(v))
                        {
                            ptsI = ut.getPointAtParam(b.Evaluator, 2);
                            v = ut.CrossProduct(ptsI[0], ptsI[2], mPt);
                        }
                        sp = ptsI[0]; ep = ptsI[1]; cen = mPt;
                        if (v.Z < 0)
                        {
                            sp = ep; ep = ptsI[0];
                        }
                        addArc(dxf, sp, ep, cen, ut.round(ut.max(l1,l2*2), 2)/2);
                    }
                }
            }
        }
        static public bool check(double [] pts)
        {
            Inventor.Point min = I.CP(pts[0], pts[1], pts[2]), max = I.CP(pts[3], pts[4], pts[5]), mid = I.CP(pts[6], pts[7], pts[8]);
            Inventor.Point m = ut.midPt(min, max);
            return ut.eq(m, mid);
        }
        static public double[] pt(double r, double a, Point2d cen)
        {
            double[] pts = new double[2];
            pts[0] = r * Math.Cos(ut.degToRad(a,6)) + cen.X; pts[1] = r * Math.Sin(ut.degToRad(a,6)) + cen.Y;
            return pts;
        }
    }
}
