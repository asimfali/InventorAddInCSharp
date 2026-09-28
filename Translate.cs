//#define INV14

using System;
using System.Windows.Forms;
using System.Drawing;
using Inventor;
using System.Linq;
using InvDoc;
using ADDIN = Macros.StandardAddInServer;
using System.Text.RegularExpressions;
using System.Collections;
using ut = InvDoc.u;
using System.Collections.Generic;
using iTextSharp.text.pdf;
using System.IO;
using iTextSharp.text.xml.xmp;
using iTextSharp.xmp;
using System.Xml.Linq;
using System.Xml;
using printLib;

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
        private System.Collections.Generic.Dictionary<char, string> dic = new System.Collections.Generic.Dictionary<char, string>();
        private System.Collections.Generic.Dictionary<char, string> dic1 = new System.Collections.Generic.Dictionary<char, string>();
        private InvDocument<Document> m_Doc;

        public PDFOp()
        {
            iniPDF();
        }

        public PDFOp(System.Collections.Generic.List<Document> docs)
        {
            start(docs);
        }

        public PDFOp(Inventor.Document newDoc)
        {
            start(I.app.Documents.VisibleDocuments.OfType<Document>().ToList());
        }

        private void start(List<Document> docs)
        {
            iniPDF();
            List<CmpData> lst = new List<CmpData>();
            foreach (Document doc in docs)
            {
                if (doc == null) continue;
                string pdfDir = doc.FullDocumentName, dxfDir = doc.FullDocumentName;
                createPath(ref pdfDir, "PDF"); createPath(ref dxfDir, "DXF");
                if (doc.DocumentType == DocumentTypeEnum.kDrawingDocumentObject)
                    lst.Add(addPDF(doc, pdfDir, false, dxfDir));
                else if (doc.SubType == "{9C464203-9BAE-11D3-8BAD-0060B0CE6BB4}")
                    addDXF(doc, false, dxfDir, true);
            }
            foreach (var item in lst)
            {
                checkTime(item);
            }
        }

        private void checkTime(CmpData d)
        {
            foreach (var item in d.dxfNames)
            {
                string dxf = item;
                if (file.check(d.pdf) && file.check(dxf))
                {
                    var pdfTime = System.IO.File.GetLastAccessTimeUtc(d.pdf);
                    var dxfTime = System.IO.File.GetLastAccessTimeUtc(dxf);
                    var ticks = Math.Abs(pdfTime.Ticks - dxfTime.Ticks);
                    System.TimeSpan ts = pdfTime - dxfTime;
                    if (ts.Minutes < 1)
                    {
                        System.IO.File.SetLastAccessTimeUtc(dxf, pdfTime);
                    }
                }
            }
        }

        public void add(Document doc, string path)
        {
            if (doc.DocumentType == DocumentTypeEnum.kDrawingDocumentObject)
                addPDF(doc, path, false);
            else if (doc.SubType == "{9C464203-9BAE-11D3-8BAD-0060B0CE6BB4}")
                addDXF(doc, false, path, true);
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
        public void addPDF(DrawingDocument drw, string path)
        {
            if (pdfAddIn.HasSaveCopyAsOptions[drw, context, nvm])
            {
                nvm.set_Value("ScaleMode", PrintScaleModeEnum.kPrintFullScale);
                nvm.set_Value("All_Color_AS_Black", 0);
                nvm.set_Value("Remove_Line_Weights", 0);
                nvm.set_Value("Vector_Resolution", 600);
                nvm.set_Value("Sheet_Range", PrintRangeEnum.kPrintAllSheets);
            }
            var name = file.name(drw.FullDocumentName);
            dm.FileName = path + name + ".pdf";
            pdfAddIn.SaveCopyAs(drw, context, nvm, dm);
        }

        public class CmpData
        {
            public string pdf, dxfPath;
            public List<string> dxfNames = new List<string>();
            public CmpData(string pdf, string dxfPath, string dxf)
            {
                Regex regex = new Regex(@"([\w\.\,]+)");
                foreach (Match item in regex.Matches(dxf))
                {

                    dxfNames.Add(System.IO.Path.Combine(dxfPath, item.Value));
                }
                this.pdf = pdf; this.dxfPath = dxfPath;
                //foreach (var item in dxf.Split(' '))
                //{
                //    if (item == "") continue;

                //}
            }
        }

        public CmpData addPDF(Document doc, string path, bool copy, string dxfPath = "")
        {
            if (doc.DocumentType == DocumentTypeEnum.kDrawingDocumentObject)
            {
                string lit, dec = "";
                m_Drw = doc as DrawingDocument;
                refDoc = m_Drw.Sheets[1].DrawingViews[1].ReferencedDocumentDescriptor.ReferencedDocument as Document;
                //ut.referendedDoc(doc);//[doc.ReferencedDocuments.Count];
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
                    if (pdfAddIn.HasSaveCopyAsOptions[m_Drw, context, nvm])
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
                    if (pdfAddIn.HasSaveCopyAsOptions[m_Drw, context, nvm])
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
                string dxfName = refDoc.PropertySets[1][7].Value.ToString();

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
                //System.TimeSpan ts = System.DateTime.Now - System.IO.File.GetCreationTime(n);
                //if (System.IO.File.Exists(n) && ts.Minutes < 2) return new CmpData(dm.FileName, dxfPath, dxfName);
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
                //readProp(getFN(dm.FileName));
                //addProp(dm.FileName, getFN(dm.FileName), "текст");

                return new CmpData(dm.FileName, dxfPath, dxfName);

                //doc.Dirty = false;
            }
            return null;
        }

        public bool setNameIzv(Document refDoc, string dat)
        {
            string fn;
            string name;
            var lit = litera(refDoc);
            var dec = u.getPropValue(refDoc, "Part Number");
            string Nizv = ut.getPropValue(refDoc, "Изв");
            //string prop = u.getPropValue(refDoc, "Vendor");
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
            }
            else
                name = lit + "EKD_" + dec + "_SB_" + dat;
            //if (prop == "")
            //{
            name = translit(name);
            fn = name + Nizv + ".pdf";
            //}
            //else
            //{
            //name = u.nameUtil(refDoc);
            //fn = name + ".pdf";
            u.addProp(refDoc, "Vendor", fn);
            //}
            return true;
        }

        public string getFN(Document refDoc)
        {
            string fn = "";
            string name = "";
            var dec = u.getPropValue(refDoc, "Part Number");
            if (dec == "") return "";
            if (dec[dec.Length - 3] == '-')
                dec = dec.Substring(0, dec.Length - 3);
            var lit = litera(refDoc);
            var dat = refDoc.PropertySets[3][1].Value.ToString().Substring(0, refDoc.PropertySets[3][1].Value.ToString().LastIndexOf(' '));
            var Nizv = ut.getPropValue(refDoc, "Изв");
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
            }
            else
            {
                name = lit + "EKD_" + dec + "_SB_" + dat;
            }
            string prop = u.getPropValue(refDoc, "Vendor");

            if (dec != "")
                name = translit(name);
            Nizv = forDxf(Nizv);
            if (prop == "")
            {
                fn = path + name + Nizv + ".pdf";
            }
            else
                fn = prop;
            return fn;
        }

        public void readProp(string fn)
        {
            if (!file.check(fn)) return;
            using (PdfReader reader = new PdfReader(fn))
                if (reader.Metadata != null)
                {
                    //XmpReader xmp = new XmpReader(reader.Metadata);
                    //    xmp.
                    //string md = System.Text.Encoding.UTF8.GetString(reader.Metadata);
                    var xml = XDocument.Load(new MemoryStream(reader.Metadata));
                    XNamespace dc = "http://purl.org/dc/elements/1.1/";
                    XNamespace rdf = "http://www.w3.org/1999/02/22-rdf-syntax-ns#";
                    XNamespace xmp = "http://ns.adobe.com/xap/1.0/";
                    XNamespace pdf = "http://ns.adobe.com/pdf/1.3/";
                    XElement desc = xml.Descendants(rdf + "Description").First();
                    var prod = desc.Attribute(pdf + "Producer");
                    foreach (var item in xml.Descendants(dc + "title"))
                    {
                        XElement el = item as XElement;
                        if (el == null) continue;
                        var val = el.Descendants(rdf + "li").First().Value;
                    }

                    //XmpCore.IXmpMeta meta = XmpCore.XmpMetaFactory.ParseFromBuffer(reader.Metadata);
                    //var r = meta.GetPropertyString(XmpConst.NS_DC, "title");
                }
        }

        public void addProp(string fn, string of, string dxf)
        {
            //using(PdfReader reader = new PdfReader(fn))
            //using(PdfWriter writer = new PdfWriter(reader, ))
            //{

            //}
            FileInfo fi = new FileInfo(fn);
            using (PdfReader reader = new PdfReader(fn))
            using (PdfStamper stamper = new PdfStamper(reader, new FileStream(of, FileMode.Create)))
            {
                Dictionary<string, string> info = reader.Info;
                //info["Author"] = "";
                //info["Creator"] = "";
                //info["Producer"] = "";
                info.Add("Title", dxf);
                MemoryStream baos = new MemoryStream();
                XmpWriter xmp = new XmpWriter(baos, info);
                if (reader.Metadata != null)
                {
                    var xml = XDocument.Load(new MemoryStream(reader.Metadata));
                }
                xmp.Close();
                stamper.XmpMetadata = baos.ToArray();
            }
        }

        public static void createPath(ref string path, string val)
        {
            if (System.IO.Path.GetExtension(path) != "") path = System.IO.Path.GetDirectoryName(path);
            path += "\\" + val + "\\";
            if (System.IO.Directory.Exists(path)) return;
            else System.IO.Directory.CreateDirectory(path);
        }

        public bool addDXF(SheetMetalComponentDefinition smcd)
        {
            var factory = smcd.iPartFactory;
            if (factory == null) return false;
            int i = 0;
            string p = factory.MemberCacheDir;
            string reviz = "";
            Document bdoc = smcd.Document as Document;
            string name = MatName(smcd, bdoc, out reviz);
            if (reviz != "") setReviz(bdoc, reviz);
            string path = ut.createDir(bdoc, @"\DXF\");
            foreach (iPartTableRow tr in smcd.iPartFactory.TableRows)
            {
                if (i > 0)
                {
                    name = MatName(smcd, smcd.Document as Document, out reviz, i);
                }
                var fn = p + "\\" + tr.PartName;
                var doc = I.open(fn);
                if (doc.RequiresUpdate) doc.Update2();
                addDXF(doc, false, path, true, name, reviz);
                doc.Close();
                i++;
            }
            return true;
        }

        public void setReviz(Document doc, string reviz)
        {
            doc.PropertySets[1][7].Value = reviz;
            doc.Save();
        }

        public string addDXF(Document doc, bool copy = false, string path = "", bool my = true,
            string name = "", string reviz = "")
        {
            if (path == "")
            {
                path = ut.createDir(doc, @"\DXF\");
            }
            SheetMetalComponentDefinition smcd = I.getSMCD(doc);
            if (smcd == null || !smcd.HasFlatPattern) return null;
            if (addDXF(smcd)) return null;
            fp = smcd.FlatPattern; I.silent(true);
            m_Prt = doc as PartDocument;
            string rev = doc.PropertySets[1][7].Value.ToString();
            if (rev.EndsWith(" "))
            {
                name = rev.Substring(0, rev.Length - 1);
            }
            if (name == "") { name = MatName(smcd, doc, out reviz); }
            if (reviz != "") rev = reviz;
            if (rev.EndsWith(" "))
            {
                name = rev.Substring(0, rev.Length - 1);
            }
            else if (doc.PropertySets[1][7].Value.ToString() != reviz)
            {
                setReviz(doc, reviz);
            }

            string n = path + name;
            //System.TimeSpan ts = System.DateTime.Now - System.IO.File.GetCreationTime(n);
            //if (System.IO.File.Exists(n) && ts.TotalSeconds < 2) return n;
            if (copy)
            {
                string p = doc.FullDocumentName.Remove(doc.FullDocumentName.LastIndexOf('\\'));
                if (invDoc.path != p + "\\"/* && System.IO.Directory.Exists(p + "\\DXF\\"*/)
                {
                    copyFile(path + name, name.Substring(name.IndexOf('_')));
                }
                else if (my) FPOp.createDXF(fp, path + name);
                else
                {
                    DataIO Io = (doc as PartDocument).ComponentDefinition.DataIO;
                    string sOut = "FLAT PATTERN DXF?AcadVersion=2000&OuterProfileLayer=0&InteriorProfilesLayer=0" +
                    "&FeatureProfileLayer=1" + "&UnconsumedSketchesLayer=0" +
                    "&SimplifySplines=True&SplineTolerance=0,1&RebaseGeometry=True&MergeProfilesIntoPolyline=False" +
                    "&InvisibleLayers=IV_TANGENT;IV_BEND;IV_BEND_DOWN;IV_TOOL_CENTER;IV_TOOL_CENTER_DOWN;IV_ARC_CENTERS;IV_FEATURE_PROFILES_DOWN";
                    Io.WriteDataToFile(sOut, path + name);
                }
            }
            else if (my) FPOp.createDXF(fp, path + name);
            else
            {
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
                        if (u.eq(fp.TopFace, item.EdgeLoop.Edges[1].StartVertex.Point) && u.eq(fp.TopFace, item.EdgeLoop.Edges[3].StartVertex.Point))
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
                if (u.eq(u.getLenght(el.Edges[i]), l / 10))
                {
                    if (min == -1) min = i + 1;
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

        public string translit(string str)
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

        public static string litera(Document doc)
        {
            string l = ut.getPropValue(doc, "Литера1") + ut.getPropValue(doc, "Литера2") + "_";
            if (l == "_") return "";
            return l;
        }

        public static string izve(Document doc)
        {
            string l = "_" + ut.getPropValue(doc, "Изв") + "_" + ut.getPropValue(doc, "ИзвД");
            if (l == "__") return "";
            return l;
        }

        public string MatName(SheetMetalComponentDefinition m_CompDef, Document doc, out string rev, int add = 0)
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
                    name = name.Substring(ind + 1, name.Length - ind - 1);
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
                _t = $"{_t}_{p}";
            }

            //             if (DXF[1] < 256)
            //             {
            //                 //DXF = "Galv.Steel";
            //                 _t = (t / 25.4).ToString("0.000") + "_" + t + "_mm";
            //             }
            name1 = forDxf(name1);
            name = name1 + "_" + DXF + "_" + _t;
            rev = name1 + "_" + DXF + "_" + _t + izv + ".dxf";
            if (add > 0)
            {
                name = name1 + "_" + add.ToString("00") + "_" + DXF + "_" + _t;
                rev = name1 + "_" + add.ToString("00") + "_" + DXF + "_" + _t + izv + ".dxf";
            }

            for (int i = 0; i < ut.convToInt(ut.getPropValue(doc, "CountDXF")); i++)
            {
                rev = rev + "    " + forDxf(name1) + "_" + (i + 1).ToString("00") + "_" + DXF + "_" + _t + izv + ".dxf";
            }

            name = forDxf(name);
            return name + izv + ".dxf";
        }
    }

    public class IzvData
    {
        public string path;
        public string drwTempl;
        public string fn;
        public string date;
        public string izv;
        public string num;
        public IzvData(string p, string fn, string time, string izv)
        {
            path = p; this.fn = fn;
            Regex r = new Regex(@"(\d{3})");
            var m = r.Match(izv);
            if (m.Groups.Count == 2) num = m.Groups[1].Value;
            this.izv = izv;
            string d = time;
            date = d.Remove(d.Length - 4, 2);
        }
        public void setTempl(string fn)
        {
            drwTempl = fn;
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
                ss.Sort(delegate (rect r1, rect r2) { return r1.x <= r2.x ? 1 : -1; });

                System.IO.StreamWriter fs = new System.IO.StreamWriter(path + "Размеры чертежей.txt", false);
                foreach (rect str in ss)
                {
                    fs.WriteLine(str.x.ToString() + "\t" + str.fn);
                }
                fs.Close();
                invApp.SilentOperation = false;
            }
            else
            {
                string namePrinter = "";
                foreach (var item in System.Drawing.Printing.PrinterSettings.InstalledPrinters)
                {
                    //if (item.ToString() == "Microsoft Print to PDF") { namePrinter = item.ToString(); break; }
                    if (item.ToString() == "Microsoft XPS Document Writer") { namePrinter = item.ToString(); break; }

                }
                foreach (Inventor.Document actdoc in invApp.Documents.VisibleDocuments)
                {
                    var printDll = new PrintDll();
                    if (actdoc.DocumentType == DocumentTypeEnum.kDrawingDocumentObject)
                    {
                        actdoc.Activate();
                        actdoc.Update();
                        prntMgr = (DrawingPrintManager)actdoc.PrintManager;
                        prntMgr.Printer = namePrinter;
                        oldSheet = ((DrawingDocument)actdoc).ActiveSheet;
                        foreach (Sheet sheet in ((DrawingDocument)actdoc).Sheets)
                        {
                            oSheet = sheet;
                            oSheet.Activate();
                            string path = file.p(actdoc.FullDocumentName) + "PDF\\";
                            file.dir(path);
                            string name = file.name(actdoc.FullDocumentName);
                            string xps = path + name + ".xps";

                            print(xps);
                            printDll.XPSToPDF(xps);
                        }
                        printDll.run();
                        //oldSheet.Activate();
                        actdoc.Dirty = false;
                    }
                }
            }
        }

        public void print(string fn, int num = 1)
        {
            //oSheet.Activate();
            w = oSheet.Width;
            h = oSheet.Height;
            //prntMgr.SetCurrentView(oSheet.ClientViews[1]);
            prntMgr.ColorMode = PrintColorModeEnum.kPrintColorPalette;
            prntMgr.NumberOfCopies = num;
            prntMgr.ScaleMode = PrintScaleModeEnum.kPrintFullScale;
            if (w == 21)
            {
                prntMgr.PaperSize = PaperSizeEnum.kPaperSizeA4;
                prntMgr.Orientation = PrintOrientationEnum.kPortraitOrientation;
                prntMgr.PrintToFile(fn);
            }
            else if (w == 42)
            {
                prntMgr.PaperSize = PaperSizeEnum.kPaperSizeA3;
                prntMgr.Orientation = PrintOrientationEnum.kLandscapeOrientation;
                prntMgr.PrintToFile(fn);
            }
            else
            {
                prntMgr.PaperSize = PaperSizeEnum.kPaperSizeCustom;
                prntMgr.PaperHeight = 29.7;
                prntMgr.PaperWidth = w;
                prntMgr.Orientation = PrintOrientationEnum.kPortraitOrientation;
                prntMgr.PrintToFile(fn);
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
                    var st = getStyle(f);
                    if (st != null) sheetMetal.ActiveSheetMetalStyle.Name = st.Name;
                    sheetMetal.Unfold2(f);
                    sheetMetal = (SheetMetalComponentDefinition)((PartDocument)doc).ComponentDefinition;
                    Camera cam = doc.Views[1].Camera;
                    SheetMetalFeatures smf = (SheetMetalFeatures)sheetMetal.Features;
                    double w = 0;
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
                            //CreateComponent.setView(fp.BaseFace);
                            //invApp.CommandManager.ControlDefinitions["AppZoomAllCmd"].Execute();
                            dr = MessageBox.Show("Поверуть на 180 градусов - Да, на 90 - Нет, оставить - Отмена", "Развертка", MessageBoxButtons.YesNoCancel);
                            if (dr == DialogResult.Yes)
                            {
                                fp.GetAlignment(out al, out obj, out re);
                                fp.SetAlignment(al, obj, !re);
                            }
                            else if (dr == DialogResult.No)
                            {

                                al = (al == AlignmentTypeEnum.kHorizontalAlignment) ? AlignmentTypeEnum.kVerticalAlignment : AlignmentTypeEnum.kHorizontalAlignment;
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
                        //fp.SetAlignment(al, obj, false);
                    }

                    //CreateComponent.setView(fp.BaseFace, cam);
                    //doc.Update();
                    //invApp.CommandManager.ControlDefinitions["AppZoomAllCmd"].Execute();
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
                    Cut(sheetMetal.FlatPattern);
                    //flatCut(sheetMetal.FlatPattern,sheetMetal);
                    return;
                }
            }
        }

        public static SheetMetalStyle findStyle(SheetMetalComponentDefinition smcd, string name)
        {
            foreach (SheetMetalStyle item in smcd.SheetMetalStyles)
            {
                if (item.Name == name) return item;
            }
            return null;
        }

        public static SheetMetalStyle getStyle(Face f)
        {
            var sb = f.Parent;
            SheetMetalStyle sms = null;
            var feat = sb.CreatedByFeature;
            if (feat is ReferenceFeature)
            {
                var sbpar = ((ReferenceFeature)feat).ReferencedEntity as SurfaceBody;
                var smcd1 = sbpar.ComponentDefinition as SheetMetalComponentDefinition;
#if INV14
                sms = smcd1.ActiveSheetMetalStyle;
#else
                sms = smcd1.GetBodySheetMetalStyle(sbpar);
#endif

            }
            return sms;
        }

        public void flatCut(FlatPattern fp, SheetMetalComponentDefinition compDef)
        {
            sketch = fp.Sketches.Add(fp.TopFace);
            constraints.ps = sketch;
            Inventor.Application invApp = Macros.StandardAddInServer.m_inventorApplication;
            string path = invApp.iFeatureOptions.RootPath;
            ObjectCollection col = invApp.TransientObjects.CreateObjectCollection();
            foreach (FlatBendResult fbr in fp.FlatBendResults)
            {
                if (fbr.IsOnBottomFace == false && fbr.InnerRadius < 0.2)
                {
                    if (fbr.Edge != null)
                    {
                        var v = fbr.Edge.StartVertex.Point.VectorTo(fbr.Edge.StopVertex.Point);
                        Vector bvv = I.CV(0, 1, 0), bvh = I.CV(1, 0, 0);
                        if (v.IsParallelTo(bvv, 0.001) || v.IsParallelTo(bvh, 0.001))
                        {
                            addCuts(sketch, fbr.Edge);
                            sp = (SketchPoint)sketch.AddByProjectingEntity(fbr.Edge.StartVertex);
                            col.Add(sp);
                            sp = (SketchPoint)sketch.AddByProjectingEntity(fbr.Edge.StopVertex);
                            col.Add(sp);
                        }
                    }
                }
            }
            iFeatureDefinition ifd = fp.Features.PunchToolFeatures.CreateiFeatureDefinition(path + "ПодСгиб.ide");
            PunchToolFeature ptf = fp.Features.PunchToolFeatures.Add(col, ifd, 0);
        }
        public static bool check_collinear(Edge e1, Edge e2)
        {
            Vector v1 = e1.StartVertex.Point.VectorTo(e1.StopVertex.Point);
            Vector v2 = e1.StartVertex.Point.VectorTo(e2.StopVertex.Point);
            if (u.isNullVector(v2))
                v2 = e1.StartVertex.Point.VectorTo(e2.StartVertex.Point);
            //return v1.IsParallelTo(v2);
            var v3 = v1.CrossProduct(v2);
            return u.isNullVector(v3);
        }
        public static bool check_collinear(Inventor.Point origin, Inventor.Point pt, Vector v)
        {
            Vector v1 = origin.VectorTo(pt);
            if (u.isNullVector(v1))
                return false;
            var v3 = v.CrossProduct(v1);
            return u.isNullVector(v3);
        }
        public static bool check_dic(Dictionary<Edge, List<Edge>> eds, Edge ed)
        {
            foreach (var item in eds)
            {
                if (item.Value.Contains(ed)) return true;
            }
            return false;
        }
        public static void set_bend_num(FlatBendResult fbr, int num)
        {
            fbr.SetBendOrder(num, true);
        }
        public static Dictionary<Edge, List<Edge>> get_edge_dic(IEnumerable<FlatBendResult> res)
        {
            Dictionary<Edge, List<Edge>> dic = new Dictionary<Edge, List<Edge>>();
            foreach (var item in res)
            {
                var ed1 = item.Edge;
                if (check_dic(dic, ed1)) continue;
                foreach (var e in res)
                {
                    var ed2 = e.Edge;
                    if (ed1.Equals(ed2)) continue;

                    if (check_collinear(ed1, ed2))
                    {
                        if (!dic.ContainsKey(ed1))
                        {
                            var lst = new List<Edge>() { ed1 };
                            dic[ed1] = lst;
                        }
                        var eds = dic[ed1];
                        if (!eds.Contains(ed2)) eds.Add(ed2);
                    }
                }
            }
            dic = dic.OrderByDescending(x => x.Value.Count).ToDictionary(x => x.Key, x => x.Value);
            return dic;
        }
        public static List<Edge> sort_list(List<Edge> eds)
        {
            if (eds.Count == 1) return eds;
            return eds.OrderBy(e => e.StartVertex.Point.X).ThenBy(e => e.StartVertex.Point.Y).ThenBy(e => e.StartVertex.Point.Z).ToList();
        }
        public static void renameBendOrder(FlatPattern fp)
        {
            int num = 1;
            Dictionary<Edge, List<Edge>> dic = new Dictionary<Edge, List<Edge>>();
            var res = u.gets<FlatBendResult>(fp.FlatBendResults, f => f.IsOnBottomFace == false);
            dic = get_edge_dic(res);
            foreach (var item in dic)
            {
                //var fbr = u.get<FlatBendResult>(res, f => f.Edge.Equals(item.Key));
                //set_bend_num(fbr, num);
                if (item.Value.Count > 1)
                {
                    foreach (var el in item.Value)
                    {
                        var fbr = u.get<FlatBendResult>(res, f => f.Edge.Equals(el));
                        set_bend_num(fbr, num);
                    }
                }
                num++;
            }
        }
        public void Cut(FlatPattern fp)
        {
            sketch = fp.Sketches.Add(fp.TopFace);
            constraints.ps = sketch;
            foreach (FlatBendResult fbr in fp.FlatBendResults)
            {
                if (fbr.IsOnBottomFace == false && fbr.InnerRadius < 0.2 && fbr.Edge != null)
                {
                    addCuts(sketch, fbr.Edge);
                }
            }
            var pr = sketch.Profiles.AddForSolid(false);
            var cd = fp.Features.CutFeatures.CreateCutDefinition(pr);
            fp.Features.CutFeatures.Add(cd);
        }
        public void addCuts(PlanarSketch ps, Edge e)
        {
            SketchLine sl = (SketchLine)ps.AddByProjectingEntity(e); sl.Construction = true;
            Point2d pt1 = sl.StartSketchPoint.Geometry, pt2 = sl.EndSketchPoint.Geometry;
            addCutTriangle(ps, pt1, pt2, sl, sl.StartSketchPoint);
            addCutTriangle(ps, pt2, pt1, sl, sl.EndSketchPoint);
        }
        public void addCutTriangle(PlanarSketch ps, Point2d pt1, Point2d pt2, SketchLine sl, SketchPoint pt)
        {
            double d = 0.1;
            var v = pt1.VectorTo(pt2); u.setDist(v, -d / 2);
            var sp4 = I.CP2d(pt1, v); u.setDist(v, -d);
            var sl4 = ps.SketchLines.AddByTwoPoints(pt, sp4);

            var n = u.normal(pt1, pt2); u.setDist(n, d);
            var sp1 = I.CP2d(sp4, n); u.setDist(n, -d);
            var sp2 = I.CP2d(sp4, n);
            var sp3 = I.CP2d(sp4, v);
            var sl1 = ps.SketchLines.AddByTwoPoints(sp1, sp2);
            constraints.addMidPoint(pt, sl4);
            constraints.addMidPoint(sl4.EndSketchPoint, sl1);
            constraints.addPerp((SketchEntity)sl, (SketchEntity)sl1);
            constraints.addCollinear(sl4, sl);
            var sl2 = ps.SketchLines.AddByTwoPoints(sl1.EndSketchPoint, sp3);
            var sl3 = ps.SketchLines.AddByTwoPoints(sl2.EndSketchPoint, sl1.StartSketchPoint);
            //constraints.addConsid((SketchEntity)sl1.StartSketchPoint, (SketchEntity)sl);
            constraints.addEq(sl2, sl3); constraints.addPerp((SketchEntity)sl2, (SketchEntity)sl3);
            string val = "1.5";
            if (ps.DimensionConstraints.Count > 0)
            {
                var se = ps.DimensionConstraints[1] as TwoPointDistanceDimConstraint;
                var asl = se.PointOne.AttachedEntities[1] as SketchLine;
                constraints.addEq(asl, sl1);
                var se2 = ps.DimensionConstraints[2] as TwoPointDistanceDimConstraint;
                var as2 = se2.PointOne.AttachedEntities[2] as SketchLine;
                constraints.addEq(as2, sl4);
            }
            else
            {
                var dc = constraints.addTwoPointDist(sl1, null, 1, 1, DimensionOrientationEnum.kAlignedDim);
                dc.Parameter.Expression = val;
                var dc1 = constraints.addTwoPointDist(sl4, null, 1, 1, DimensionOrientationEnum.kAlignedDim);
                dc1.Parameter.Expression = "0.2";
            }
        }

        static public object flatCut(SheetMetalFeatures smf, PlanarSketch ps, string name, string a = "5", string b = "8")
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
        static public void createDXF(Face face, string p)
        {
            double m = 10;
            dxf.DXF dXf = new dxf.DXF();
            dxf.DXF.tol = 5;
            dXf.initPre();
            Plane pl = face.Geometry as Plane;
            var ov = pl.RootPoint;
            var n = pl.Normal;
            var e = u.get<Edge>(face.Edges, el => el.GeometryType == CurveTypeEnum.kLineSegmentCurve);
            Vector x;
            if (e == null)
            {
                e = face.Edges[1];
                var ev = e.Evaluator;
                double[] tang = new double[] { }, par = new double[] { 0 };
                ev.GetTangent(ref par, ref tang);

                x = I.CV(tang[0], tang[1], tang[2]);
            }
            else x = u.getVector(e); 
            x.Normalize();
            var y = x.CrossProduct(n.AsVector());
            var mtx = I.getMatrix();
            mtx.SetToAlignCoordinateSystems(ov, x, y, n.AsVector(), I.CP(), I.CV(1, 0, 0), I.CV(0, 1, 0), I.CV(0, 0, 1));


            ut.action<EdgeLoop>(face.EdgeLoops, ed => createDXF(ed, dXf, m, mtx));

            dXf.bounds();
            dxf.vector v = new dxf.vector(-dXf.gab.l, -dXf.gab.b);
            dXf.move(v);
            dXf.initPost();

            //name = "VPORT"; type = dxf.entityType.Viewport;
            //vals = dXf.init(type);
            //dXf.convert(vals);
            //t = dXf.addTable(name, 8);
            //tr = t.addRec(name, type, vals); tr.conf = "*ACTIVE";

            dXf.addFileName(p);
            dXf.safe("");
        }
        static public void createDXF(FlatPattern fp, string p)
        {
            double m = 10;
            //Inventor.Point origin = fp.RangeBox.MinPoint;
            //dxf.vector v = new dxf.vector(origin.X*m, origin.Y*m);
            dxf.DXF dXf = new dxf.DXF();
            dxf.DXF.tol = 5;
            dXf.initPre();
            //var d = dXf.addDict(null, null);
            //dXf.addDict("ACAD_GROUP", d);
            //var b = dXf.addBlock("*Model_Space");
            //b = dXf.addBlock("*Paper_Space");

            //List<dxf.Data> vals = null;

            //string name = "LTYPE";dxf.entityType type = dxf.entityType.Linetype;
            //var t = dXf.addTable(name, 5); t.conf = name;
            //name = "ByBlock";
            //vals = dXf.init(type);
            //var tr = t.addRec(name, type, vals); tr.conf = name;
            //name = "ByLayer";
            //vals = dXf.init(type);
            //tr = t.addRec(name, type, vals); tr.conf = name;
            //name = "Continuous";
            //vals = dXf.init(type, name);
            //tr = t.addRec(name, type, vals); tr.conf = name;

            //name = "LAYER"; type = dxf.entityType.Layer;
            //vals = dXf.init(type);
            //t = dXf.addTable(name, 2); t.conf = name;
            //tr = t.addRec(name, type, vals); tr.conf = "0";

            //name = "STYLE"; type = dxf.entityType.TextStyle;
            //vals = dXf.init(type); 
            //t = dXf.addTable(name, 3); t.conf = name;
            //tr = t.addRec(name, type, vals); tr.conf = "Standard";

            //name = "VIEW"; type = dxf.entityType.View;
            //vals = dXf.init(type);
            //t = dXf.addTable(name, 6); t.conf = name;

            //name = "UCS"; type = dxf.entityType.Usc;
            //vals = dXf.init(type, "Ucs");
            //t = dXf.addTable(name, 7); t.conf = name;

            //name = "APPID"; type = dxf.entityType.RegApp;
            //vals = dXf.init(type);
            //t = dXf.addTable(name, 9); t.conf = name; t.conf = "ACAD";
            //tr = t.addRec(name, type, vals); tr.conf = "ACAD";

            //name = "DIMSTYLE"; type = dxf.entityType.DimStyle;
            //vals = dXf.init(type);
            //t = dXf.addTable(name, 10); t.conf = name;
            //tr = t.addRec(name, type, vals); tr.conf = "Standard";

            //name = "BLOCK_RECORD"; type = dxf.entityType.Block;
            //vals = dXf.init(type);
            //t = dXf.addTable(name, 1); t.conf = name;
            //name = "*Model_Space";
            //tr = t.addRec(name, type, vals); tr.conf = name;

            //name = "*Paper_Space";
            //tr = t.addRec(name, type, vals); tr.conf = name;
#if INV14
            ut.action<EdgeLoop>(fp.TopFace.EdgeLoops, e => createDXF(e, dXf, m));
#else
            foreach (Face item in fp.TopFaces)
                        {
                            //if (item.Edges.Count < 4) continue;
                            ut.action<EdgeLoop>(item.EdgeLoops, e => createDXF(e, dXf, m));
                        }
#endif

            ut.action<PlanarSketch>(fp.Sketches, e => createDXF(e, dXf, m), f => !f.Consumed);
            dXf.bounds();
            dxf.vector v = new dxf.vector(-dXf.gab.l, -dXf.gab.b);
            dXf.move(v);
            dXf.initPost();

            //name = "VPORT"; type = dxf.entityType.Viewport;
            //vals = dXf.init(type);
            //dXf.convert(vals);
            //t = dXf.addTable(name, 8);
            //tr = t.addRec(name, type, vals); tr.conf = "*ACTIVE";

            dXf.addFileName(p);
            dXf.safe("");
        }
        static public void createDXF(EdgeLoop el, dxf.DXF dxf, double m = 10, Matrix mtx = null)
        {
            //if (el.Edges.Count == 2)
            //{
            //    BSplineCurve spl1 = el.Edges[1].Geometry as BSplineCurve, spl2 = el.Edges[2].Geometry as BSplineCurve;
            //    bool flag = false;
            //    if (spl1 != null && spl2 != null)
            //    {
            //        flag = addCircle(dxf, spl1, spl2);
            //    }
            //    if (flag) return;
            //}
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
                ut.action<Edge>(col, ed => createDXF(ed, dxf, spl: true, mtx: mtx), f => !(f.Equals(e1) || f.Equals(e2)));
                return;
            }
            foreach (Edge e in el.Edges)
            {
                createDXF(e, dxf, 10, true, mtx);
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
                    try
                    {
                        if (a.Geometry3d.Normal.Z < 0)
                        {
                            sp = ep; ep = a.StartSketchPoint.Geometry3d;
                        }
                    }
                    catch (System.Exception ex)
                    {

                    }


                    addArc(dxf, sp, ep, cen, a.Radius);
                    //double [] param = ut.getParamExtents(a.Geometry.Evaluator);
                    //double [] pts = ut.getPointAtParam(a.Geometry.Evaluator, new []{(param[1]+param[0])/2});
                    //addArc(dxf, sp, ep, cen, a.Radius, pts);
                }
            }
        }
        static public void addArc(dxf.DXF dxf, Inventor.Point sp, Inventor.Point ep, Inventor.Point cen, double R, double[] pts, double m = 10)
        {
            double sa = ut.angFromPoint(ep, cen, m),
                   ea = ut.angFromPoint(sp, cen, m);
            double an = (ea + sa) / 2;
            if (an > 180) an -= 180;
            double[] evalPts = pt(R, an, I.CP2d(cen.X, cen.Y));
            double dist = I.CP2d(pts[0], pts[1]).DistanceTo(I.CP2d(evalPts[0], evalPts[1]));
            //if (!(ut.eq(pts[0], evalPts[0]) && ut.eq(pts[1], evalPts[1])))
            if (sa > ea && dist < R / 2)
                dxf.addArc(new dxf.coord(cen.X * m, cen.Y * m), R * m, sa, ea);
            else
                dxf.addArc(new dxf.coord(cen.X * m, cen.Y * m), R * m, ea, sa);
        }
        static public void addArc(dxf.DXF dxf, Inventor.Point sp, Inventor.Point ep, Inventor.Point cen, double R, double m = 10)
        {
            dxf.addArc(new dxf.coord(sp.X * m, sp.Y * m), new dxf.coord(ep.X * m, ep.Y * m), new dxf.coord(cen.X * m, cen.Y * m), R * m);
        }
        static public bool addCircle(dxf.DXF dxf, BSplineCurve spl1, BSplineCurve spl2, double m = 10)
        {
            List<Inventor.Point> pts1 = ut.getPointAtParam(spl1.Evaluator), pts2 = ut.getPointAtParam(spl2.Evaluator);
            if (ut.eq(pts1[0], pts2[1]) || ut.eq(pts1[1], pts2[0]))
            {
                Inventor.Point mPt = ut.midPt(pts1[2], pts2[2]);
                double R = ut.round(pts1[2].VectorTo(pts2[2]).Length, 2) / 2;
                dxf.addCircle(new dxf.coord(mPt.X * m, mPt.Y * m), R * m);
                return true;
            }
            return false;
        }
        static public void createDXF(Edge e, dxf.DXF dxf, double m = 10, bool spl = false, Matrix mtx = null)
        {
            Inventor.Point sp, ep, cen;
            Circle c = e.Geometry as Circle;
            if (c != null)
            {
                cen = c.Center;
                if (mtx != null) cen.TransformBy(mtx);
                dxf.addCircle(new dxf.coord(cen.X * m, cen.Y * m), c.Radius * m);
                return;
            }
            LineSegment l = e.Geometry as LineSegment;
            if (l != null)
            {
                sp = l.StartPoint; ep = l.EndPoint;
                if (mtx != null) { sp.TransformBy(mtx); ep.TransformBy(mtx); }
                dxf.addLine(new dxf.coord(sp.X * m, sp.Y * m), new dxf.coord(ep.X * m, ep.Y * m));
                return;
            }
            Arc3d a = e.Geometry as Arc3d;
            if (a != null)
            {
                sp = a.StartPoint; ep = a.EndPoint; cen = a.Center;
                if (mtx != null) { sp.TransformBy(mtx); ep.TransformBy(mtx); cen.TransformBy(mtx); }
                if (a.Normal.Z < 0)
                {
                    sp = ep; ep = a.StartPoint;
                }

                addArc(dxf, sp, ep, cen, a.Radius);
                return;
            }
            EllipseFull el = e.Geometry as EllipseFull;
            if (el != null)
            {
                cen = el.Center;
                if (mtx != null) cen.TransformBy(mtx);
                if (u.eq(el.MinorMajorRatio, 1))
                {
                    dxf.addCircle(new dxf.coord(cen.X * m, cen.Y * m), u.round(el.MajorAxisVector.Length, 3) * m);
                }
            }
            BSplineCurve b = e.Geometry as BSplineCurve;
            if (b != null)
            {
                double min, max;
                double[] poles = new double[] { }, knots = new double[] { }, w = new double[] { }, pv = new double[] { };
                int ord, numP, numK;
                bool isR, isP, isC, isPl;
                b.GetBSplineInfo(out ord, out numP, out numK, out isR, out isP, out isC, out isPl, ref pv);
                //poles = new double[numP]; knots = new double[numK]; w = new double[numK];
                //b.GetBSplineData(ref poles, ref knots, ref w);
                b.Evaluator.GetParamExtents(out min, out max);
                double[] pts = new double[9]; double[] param = new double[] { min, max, (max + min) / 2 };
                b.Evaluator.GetPointAtParam(ref param, ref pts);
                double tol = 0.5;
                if (Macros.StandardAddInServer.DXFSpl == "0" && check(pts))
                {
                    dxf.addLine(new dxf.coord(pts[0] * m, pts[1] * m), new dxf.coord(pts[3] * m, pts[4] * m));
                    return;
                }
                else if (spl)
                {
                    b.GetBSplineData(ref poles, ref knots, ref w);
                    var info = new List<int>() { };
                    List<dxf.myDouble> myKnots = new List<dxf.myDouble>();
                    List<dxf.coord> coord = new List<dxf.coord>();
                    var par = new double[] { min, max };
                    var tang = new double[] { };
                    b.Evaluator.GetTangent(ref par, ref tang);
                    dxf.coord st = new dxf.coord(tang[0], tang[1], 12), et = new dxf.coord(tang[3], tang[4], 13);
                    foreach (var item in knots)
                    {
                        myKnots.Add(new dxf.myDouble(item));
                    }
                    for (int i = 0; i < poles.Length; i += 3)
                    {
                        cen = I.CP(poles[i], poles[i + 1], poles[i + 2]);
                        if (mtx != null) cen.TransformBy(mtx);
                        coord.Add(new dxf.coord(cen.X * m, cen.Y * m, 10));
                    }
                    int flag = 0;
                    if (isPl) flag += 8;
                    if (isC) flag += 1;
                    if (isP) flag += 2;
                    if (isR) flag += 4;
                    if (flag == 0) flag = 8;
                    dxf.addSpline(st, et, coord, myKnots, new List<int>() { flag, ord - 1, numK, numP });
                    return;
                }
                if (ut.eq(pts[0], pts[3]) && ut.eq(pts[1], pts[4]))
                {

                    double step = (max - min) / 4;
                    //double[] spl = new double[] { min, step, step*2, step*3};
                    //double[] splPts = ut.getPointAtParam(b.Evaluator, spl);
                    List<Inventor.Point> ptsInv = ut.getPoleAtIndex(b, ord, numP);
                    double l1 = ptsInv[0].DistanceTo(ptsInv[2]), l2 = ptsInv[1].DistanceTo(ptsInv[3]);

                    if (Math.Abs(l2 - l1) < tol)
                    {
                        Inventor.Point mPt = ut.midPt(ptsInv[0], ptsInv[2]);
                        dxf.addCircle(new dxf.coord(mPt.X * m, mPt.Y * m), ut.round(l2, 2) / 2 * m);
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
                        addArc(dxf, sp, ep, cen, ut.round(ut.max(l1, l2 * 2), 2) / 2);
                    }
                }
            }
        }
        static public bool check(double[] pts)
        {
            Inventor.Point min = I.CP(pts[0], pts[1], pts[2]), max = I.CP(pts[3], pts[4], pts[5]), mid = I.CP(pts[6], pts[7], pts[8]);
            Inventor.Point m = ut.midPt(min, max);
            if (min.DistanceTo(max) > 10) return false;
            double d = max.DistanceTo(min);
            return ut.eq(m, mid, d * 0.2);
            //return ut.eq(m, mid);
        }
        static public double[] pt(double r, double a, Point2d cen)
        {
            double[] pts = new double[2];
            pts[0] = r * Math.Cos(ut.degToRad(a, 6)) + cen.X; pts[1] = r * Math.Sin(ut.degToRad(a, 6)) + cen.Y;
            return pts;
        }
    }

    public class toOBJ
    {
        TranslatorAddIn objTrans = null;
        TranslationContext context = null;
        NameValueMap nvm = null;
        DataMedium dm = null;
        public toOBJ()
        {
            InvDoc.u.action<Document>(I.app.Documents.VisibleDocuments, e => create(e));
        }
        public void create(Document doc)
        {
            if (doc.DocumentType != DocumentTypeEnum.kAssemblyDocumentObject) return;
            objTrans = (TranslatorAddIn)I.app.ApplicationAddIns.ItemById["{f539fb09-fc01-4260-a429-1818b14d6bac}"];
            context = I.app.TransientObjects.CreateTranslationContext();
            context.Type = IOMechanismEnum.kFileBrowseIOMechanism;
            nvm = I.app.TransientObjects.CreateNameValueMap();
            dm = I.app.TransientObjects.CreateDataMedium();
            string path = file.p(doc.FullDocumentName) + "obj";
            file.createPath(path);
            dm.FileName = path + "\\" + file.name(doc.FullDocumentName) + ".obj";
            objTrans.SaveCopyAs(doc, context, nvm, dm);
        }
    }
}
