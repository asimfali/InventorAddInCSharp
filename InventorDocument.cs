#define INV14

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Reflection;
using Inventor;
using System.Xml.Linq;
using InvAddIn;
using System.Text.RegularExpressions;

namespace InvDoc
{
    public enum TypeDocs : int { DrawingDoc = 0, AssemblyDoc = 2, PartDoc = 1}
    public enum TypeViews : int { BaseView = 0, DetailView = 1, SectionView = 2, ProjectedView = 3, OverlayView = 4, AuxiliaryView = 5}
    public class InvDocument <T>
    {
        protected T oDoc;
        protected BOM m_BOM;
        public Document doc {get; set;}
        DerivedPartUniformScaleDef derDef;
        DerivedAssemblyDefinition asmDef;
        public NameValueMap nvmOptions;

        public static Inventor.Application invApp;
        public Transaction tr;
        public List<Document> files;
        Sheet sh;
        //private InvDoc.XML name;
        public string path;
        public List<string> exept = new List<string>{"OldVersions"};
        public static List<string> pathTemplate = new List<string>();
        public static List<string> pathLibrary = new List<string>();
        public static List<string> library = new List<string>();
        public static string templateName;
        public delegate T add <T>();
        public delegate void newDoc(string str);
        public delegate string str(string s);

        #region property
        public List<string> libraryPath
        {
            get
            {
                return pathLibrary;
            }
            set
            {
                pathLibrary.Add(path);
                foreach (ProjectPath p in invApp.DesignProjectManager.ActiveDesignProject.LibraryPaths)
                {
                     pathLibrary.Add(p.Path + "\\");
                }
                List<string> tmp = new List<string>(10);
                foreach (string p in pathLibrary)
                {
                    IEnumerable<string> dir = System.IO.Directory.EnumerateDirectories(p, "*", System.IO.SearchOption.AllDirectories);
                    foreach (string str in dir)          
                    {
                        string old = str.Substring(str.LastIndexOf("\\"), str.Length - str.LastIndexOf("\\"));
                        if (!exept.Exists(e => "\\" + e == old)) tmp.Add(str);
                    }
                }
                foreach (var item in tmp)
                    if (item.EndsWith("\\"))
                        pathLibrary.Add(item);
                    else pathLibrary.Add(item + "\\");
            }
        }

        public List<string> pathes
        {
            get
            {
                return library;
            }
            set
            {
                library.Add(path);
                foreach (ProjectPath p in invApp.DesignProjectManager.ActiveDesignProject.LibraryPaths)
                {
                    library.Add(p.Path + "\\");
                }
            }
        }

        public string pPath { get { str l = s => s.Substring(0, s.LastIndexOf("\\") + 1); return l(invApp.DesignProjectManager.ActiveDesignProject.FullFileName); } }

        public PartComponentDefinition getPartCompDef{get { return ((PartDocument)oDoc).ComponentDefinition; }}

        public AssemblyComponentDefinition getAsmCompDef { get { return ((AssemblyDocument)oDoc).ComponentDefinition; } }

        public SheetMetalComponentDefinition getSheetMetalCompDef { get { return (((Document)oDoc).SubType == "{9C464203-9BAE-11D3-8BAD-0060B0CE6BB4}") ? (SheetMetalComponentDefinition)getPartCompDef : null; } }

        public Sheets getSheets { get { return ((DrawingDocument)oDoc).Sheets; } }

        public T GetDoc { get { return oDoc; } set { oDoc = value; } }

        public string getType { get { return ((Document)oDoc).PropertySets[3][2].Value.ToString().Substring(0, ((Document)oDoc).PropertySets[3][2].Value.ToString().IndexOf('.')); } }

        #endregion

        public InvDocument(T doc)
        {
            oDoc = doc;  
            invApp = (Inventor.Application)((Inventor._Document)doc).Parent;
            str l = s => s.Substring(0, s.LastIndexOf("\\") + 1);
            path = l(((Document)doc).FullDocumentName);
            nvmOptions = invApp.TransientObjects.CreateNameValueMap();
            nvmOptions.Add("SkipAllUnresolvedFiles", true);

            str template = s =>
            {
                InvDoc.XML n = new InvDoc.XML(I.p() + @"\TemplatePath.xml");
                List<string> strs = new List<string>();
                strs = n.ReadXML("Template", s);
                return templateName = invApp.DesignProjectManager.ActiveDesignProject.TemplatesPath + strs[0];
            };
            if (pathTemplate.Count != 3)
            {
                pathTemplate.Add(template("DrawingTemplate"));
                pathTemplate.Add(template("PartTemplate"));
                pathTemplate.Add(template("AssemblyTemplate"));
            }
        }

        public DrawingDocument addDrwDoc (bool visible =  false) {return (DrawingDocument)invApp.Documents.Add(DocumentTypeEnum.kDrawingDocumentObject, pathTemplate[0], CreateVisible: visible);}
        public PartDocument addPrtDoc(bool visible = false) { return (PartDocument)invApp.Documents.Add(DocumentTypeEnum.kPartDocumentObject, pathTemplate[1], CreateVisible: visible); }
        public AssemblyDocument addAsmDoc(bool visible = false) { return (AssemblyDocument)invApp.Documents.Add(DocumentTypeEnum.kAssemblyDocumentObject, pathTemplate[2], CreateVisible: visible); }

        public DrawingDocument openDrwDoc(string name, bool visible = false) { return (DrawingDocument)invApp.Documents.OpenWithOptions(name, nvmOptions, visible); }
        public PartDocument openPrtDoc(string name, bool visible = false) { return (PartDocument)invApp.Documents.OpenWithOptions(name, nvmOptions, visible); }
        public AssemblyDocument openAsmDoc(string name, bool visible = false) { return (AssemblyDocument)invApp.Documents.OpenWithOptions(name, nvmOptions, visible); }
        public Document openDoc(string name, bool visible = false) 
        {
//             if (System.IO.Path.GetExtension(name) == ".ipt")
//                 return (Document)invApp.Documents.Open(name, visible);
//             else 
            return (Document)invApp.Documents.OpenWithOptions(name, nvmOptions, visible);
        }

        public PartDocument derivedDoc(string name, string nameFind = "", PartDocument doc = null, string parNames = "", DerivedPartMirrorPlaneEnum pln = DerivedPartMirrorPlaneEnum.kDerivedPartNoMirrorPlane)
        {
            bool open = true;
            if (doc == null) {doc = addPrtDoc(); open = false;}
            SheetMetalComponentDefinition partDef = null;
            SheetMetalStyle sms = null;
            if (name.EndsWith(".ipt"))
            {
                if (doc == null)
                {
                    partDef = (SheetMetalComponentDefinition)((PartDocument)invApp.Documents.Open(name, false)).ComponentDefinition;
                    sms = partDef.ActiveSheetMetalStyle;
                }
                if (!open)
                    derDef = doc.ComponentDefinition.ReferenceComponents.DerivedPartComponents.CreateUniformScaleDef(name);
                else derDef = doc.ComponentDefinition.ReferenceComponents.DerivedPartComponents[1].Definition as DerivedPartUniformScaleDef;
                //derDef.IncludeAllSketches = DerivedComponentOptionEnum.kDerivedExcludeAll;
                derDef.IncludeAlliMateDefinitions = DerivedComponentOptionEnum.kDerivedIncludeAll;
                derDef.UseColorOverridesFromSource = true;
                if (derDef.DeriveStyle != DerivedComponentStyleEnum.kDeriveAsSingleBodyWithSeams) {
                    derDef.DeriveStyle = DerivedComponentStyleEnum.kDeriveAsSingleBodyWithSeams;
                }
                                
                if (pln != DerivedPartMirrorPlaneEnum.kDerivedPartNoMirrorPlane)
                    derDef.MirrorPlane = pln;
                if (parNames != "")
                {
                    derDef.IncludeAllParameters = false;
                    string[] tmp;
                    List<string> lst = new List<string>();
                    if (parNames.IndexOf(';') != -1)
                    {
                        if (parNames.EndsWith(";")) parNames = parNames.Remove(parNames.Length - 1);
                        tmp = parNames.Split(';');
                        lst = tmp.ToList();
                    }
                    else 
                        lst.Add(parNames);
                    foreach (DerivedPartEntity item in derDef.Parameters)
                    {
                        if (lst.Exists(e => e == (item.ReferencedEntity as Parameter).Name))
                            item.IncludeEntity = true;
                    }
                }
                foreach (DerivedPartEntity item in derDef.WorkFeatures)
                {
                    if (item.IncludeEntity == true) item.IncludeEntity = false;
                }
                foreach (DerivedPartEntity item in derDef.Sketches)
                {
                    item.IncludeEntity = false;
                }
                if (nameFind != "")

                    if (nameFind.IndexOf(';') != -1)
                    {
                       if (nameFind.EndsWith(";")) nameFind = nameFind.Remove(nameFind.Length - 1);
                       string [] tmp = nameFind.Split(';');
                       foreach (DerivedPartEntity item in derDef.Sketches)
                       {
                           if (tmp.ToList().Exists(e => ((Sketch)item.ReferencedEntity).Name.IndexOf(e) != -1))
                           {
                               item.IncludeEntity = true;
                           }
                           else item.IncludeEntity = false; 
                       }
                    }

                    else foreach (DerivedPartEntity item in derDef.Sketches)
                    {
                        if (((Sketch)item.ReferencedEntity).Name.IndexOf(nameFind) != -1)
                        {
                            item.IncludeEntity = true;
                        }
                        else item.IncludeEntity = false;
                    }
                }
            else 
            {
                asmDef = doc.ComponentDefinition.ReferenceComponents.DerivedAssemblyComponents.CreateDefinition(name);
                asmDef.DeriveStyle = DerivedComponentStyleEnum.kDeriveAsMultipleBodies;
                asmDef.ReducedMemoryMode = true;
                asmDef.SetHolePatchingOptions(DerivedHolePatchEnum.kDerivedPatchNone);
            }

            if (name.EndsWith(".ipt")) 
            {
                DerivedPartComponent derComp = null;
                if (!open)
                    derComp = doc.ComponentDefinition.ReferenceComponents.DerivedPartComponents.Add((DerivedPartDefinition)derDef);
                else doc.ComponentDefinition.ReferenceComponents.DerivedPartComponents[1].Definition = (Inventor.DerivedPartDefinition)derDef;
            }
                
            else doc.ComponentDefinition.ReferenceComponents.DerivedAssemblyComponents.Add(asmDef);
            if (doc.SubType == "{9C464203-9BAE-11D3-8BAD-0060B0CE6BB4}")
            {
                partDef = (SheetMetalComponentDefinition)((PartDocument)doc).ComponentDefinition;
                if (partDef.ReferenceComponents.DerivedPartComponents.Count == 1)
                {
                    try
                    {
                        foreach (DerivedParameter param in partDef.ReferenceComponents.DerivedPartComponents[1].Parameters)
                        {
                            param.ReferencedEntity.ExposedAsProperty = true;
                        }
                    foreach (SurfaceBody item in partDef.ReferenceComponents.DerivedPartComponents[1].SurfaceBodies)
                    {
                        item.Exported = true;  
                    }
//                     foreach (WorkAxis item in partDef.ReferenceComponents.DerivedPartComponents[1].WorkFeatures)
//                     {
//                         item.ReferencedEntity.Exported = true;
//                     }
//                     foreach (WorkPlane item in partDef.ReferenceComponents.DerivedPartComponents[1].WorkFeatures)
//                     {
//                         item.ReferencedEntity.Exported = true;
//                     }
//                     foreach (WorkPoint item in partDef.ReferenceComponents.DerivedPartComponents[1].WorkFeatures)
//                     {
//                         item.ReferencedEntity.Exported = true;
//                     }
                    }
                    catch (Exception)
                    {
                    }
                }
                if (doc == null)
                {
                    var st = partDef.SheetMetalStyles.OfType<SheetMetalStyle>().FirstOrDefault(e => e.Name == sms.Name);
                    if (st != null) partDef.SheetMetalStyles[sms.Name].Activate();
                }
            }
            return doc;
        }

        public AssemblyDocument copyAsm(string name,string baseName)
        {
            System.IO.File.Copy(baseName, name);
            return openAsmDoc(name);
        }

        static public DrawingView addView(Sheet sh, Document doc, XElement e, string list, double s, List<double> scales, Camera cam = null)
        {
            if (!System.IO.File.Exists(I.p() + @"\sheet.xml")) return null;
            //AutomatedCenterlineSettings acls = ((DrawingDocument)sh.Parent).DrawingSettings.AutomatedCenterlineSettings;
            //acls.ApplyToCylinders = false; acls.ApplyToHoles = true; acls.ApplyToRectangularPatterns = true; acls.ApplyToCircularPatterns = true;
            //acls.ApplyToRevolutions = true; acls.ProjectionParallelAxis = true;
            NameValueMap nvm = Macros.StandardAddInServer.m_inventorApplication.TransientObjects.CreateNameValueMap();
            if (list == "view2")
                nvm.Add("SheetMetalFoldedModel", false);
            DrawingView view = null; int i = 0;

            foreach (var el in e.Elements(list))
            {
                bool offsetV = false;
                Point2d pt = PositionView(el, new string[] { "offsetX", "offsetY" }, view, ref offsetV);
                ViewOrientationTypeEnum viewEnum = ViewOrientationTypeEnum.kDefaultViewOrientation;
                //s = 1;
                bool flag = false;
                if (s == 0) { s = 1; flag = true; }
                if (el.Attribute("scale") != null)
                s = double.Parse(el.Attribute("scale").Value.Replace('.',','));
                if (i == 0)
                {
                    viewEnum = (ViewOrientationTypeEnum)(int.Parse(el.Attribute("orient").Value));
                    view = addDV(doc, pt, sh, TypeViews.BaseView, s, viewEnum, nvm, cam: cam);
                    if (flag)
                    {
                        view.Scale = InvDocument<string>.autoScale(view, scales);
                    }
                    if (el.Attribute("rotate") != null)
                        view.Rotation = double.Parse(el.Attribute("rotate").Value);
                }
                else
                {
                    bool align = false;
                    if (el.Attribute("align") != null )
                    {
                        if (el.Attribute("align").Value == "false")
                            align = false;
                        else align = true;
                    }
                    addDV(doc, pt, sh, TypeViews.ProjectedView, view.Scale, viewEnum, nvm, view, align, offsetV);
                }
                i++;
            }

            return view;
        }

        static public Point2d PositionView(XElement el, string [] nameAttr, DrawingView dv, ref bool offsetV)
        {
            Point2d pt = Macros.StandardAddInServer.m_inventorApplication.TransientGeometry.CreatePoint2d();
            XAttribute attr = null; double[] arr = new double[2];
            int i = 0;
            foreach (var item in nameAttr)
            {
                if ((attr = el.Attribute(item)) != null)
                {
                    arr[i] = double.Parse(attr.Value.Replace('.', ','));
                }
                i++;
            }
            bool ret = false;
            if (dv != null)
               ret = offsetView(ref pt, dv, arr);
            offsetV = ret;
            return !ret ?
                Macros.StandardAddInServer.m_inventorApplication.TransientGeometry.CreatePoint2d(double.Parse(el.Attribute("basePointX").Value.Replace('.',',')) * 0.1,
                    double.Parse(el.Attribute("basePointY").Value.Replace('.', ',')) * 0.1) : pt;
        }

        static public double autoScale(DrawingView dv, List<double> scales)
        {
            Sheet sh = dv.Parent; double cur = 0;
            double h = dv.Height, w = dv.Width, hSh = sh.Height - 7;
            if (h / w > 2 || h / w < 0.5)
            {
                cur = w;
            }
            else cur = Math.Max(h, w);
            double curScale = hSh / cur;
            for (int i = 1; i < scales.Count; i++)
			{
			   if(curScale < scales[i])
               {
                   if (i > 2)
                   return scales[i-2];
               }
			}
            return 1;
        }

        static private bool offsetView(ref Point2d pt, DrawingView dv, double [] xy)
        {
            bool ret = false;
            pt = dv.Position;
            if (xy[0] < 0)
            { pt.X = pt.X - dv.Width / 2 + xy[0] * 0.1; ret = true; }
            else if (xy[0] > 0)
            { pt.X = pt.X + dv.Width / 2 + xy[0] * 0.1; ret = true; }
            if (xy[1] < 0)
            { pt.Y = pt.Y - dv.Height / 2 + xy[1] * 0.1; ret = true; }
            else if (xy[1] > 0)
            { pt.Y = pt.Y + dv.Height / 2 + xy[1] * 0.1; ret = true; }
            return ret;
        }
        static public DesignViewRepresentation getDVR(Document doc, string n)
        {
            RepresentationsManager mgr = null;
            if (doc.DocumentType == DocumentTypeEnum.kPartDocumentObject)
                mgr = ((PartDocument)doc).ComponentDefinition.RepresentationsManager;
            else if (doc.DocumentType == DocumentTypeEnum.kAssemblyDocumentObject)
                mgr = ((AssemblyDocument)doc).ComponentDefinition.RepresentationsManager;
            foreach (DesignViewRepresentation item in mgr.DesignViewRepresentations)
            {
                if (item.Name.ToLower().StartsWith(n))
                {
                    return item;
                }
            }
            return mgr.ActiveDesignViewRepresentation;
        }

        static public DrawingView addDV(Document doc, Point2d pt, Sheet sh, TypeViews type, double scale, 
            ViewOrientationTypeEnum viewEnum, NameValueMap nvm, DrawingView parent = null, bool align = false, bool offseV = false, Camera cam = null)
        {
            DrawingView dv = null;
            switch (type)
            {
                case TypeViews.BaseView:
                    
                    dv = sh.DrawingViews.AddBaseView((_Document)doc, pt, scale, viewEnum,
                        DrawingViewStyleEnum.kHiddenLineRemovedDrawingViewStyle, "", cam,
                        AdditionalOptions: nvm);
                    dv.IsRasterView = false;
                    //doc.Update();
                    if (nvm.Count == 0)
                    {
                        if (viewEnum != ViewOrientationTypeEnum.kCurrentViewOrientation &&
                            viewEnum != ViewOrientationTypeEnum.kArbitraryViewOrientation)
                        {
                            dv.Camera.ViewOrientationType = viewEnum;
                            dv.Camera.Apply();
                        }
                    }
                    break;
                case TypeViews.DetailView:
                    break;
                case TypeViews.SectionView:
                    break;
                case TypeViews.ProjectedView:
                    dv = sh.DrawingViews.AddProjectedView(parent, pt, DrawingViewStyleEnum.kHiddenLineRemovedDrawingViewStyle);
                    if (align == false)
                        dv.Aligned = false;
                    if (offseV)
                    {
                        Vector2d vec = parent.Position.VectorTo(pt);
                        vec.Normalize();
                        if (pt.Y.Equals(parent.Position.Y)) vec.ScaleBy(dv.Width/2);
                        else vec.ScaleBy(dv.Height/2);
                        pt = dv.Position;
                        pt.TranslateBy(vec);
                        dv.Position = pt;
                    }
                    break;
                case TypeViews.OverlayView:
                    break;
                case TypeViews.AuxiliaryView:
                    break;
                default:
                    break;
            }
            dv.DisplayTangentEdges = true;
            return dv;
        }

        public void copySketchSymbolDefinition(DrawingDocument from, DrawingDocument to ,string [] names)
        {
            foreach (var name in names)
            {
                from.TitleBlockDefinitions[name].CopyTo((_DrawingDocument)to, true);
            }
        }

        static public void copySketchSymbolDefinition(DrawingDocument from, DrawingDocument to, string name)
        {
            from.SketchedSymbolDefinitions[name].CopyTo((_DrawingDocument)to, true);
        }

        public Sheet addSheet(double w = 420, double h = 297)
        {
            tr = invApp.TransactionManager.StartTransaction((_Document)doc, "Добавить лист");
            sh = ((DrawingDocument)doc).Sheets.Add(DrawingSheetSizeEnum.kCustomDrawingSheetSize, Width: w / 10, Height: h / 10);
            sh = sheet(sh, new List<string> { "ГОСТ - Доп. графы 1", "ГОСТ - Доп. графы 2", "ГОСТ - Доп. графы 3", "АР", "Развертка 1:N АР" }, w, h);
            tr.End();
            return sh;
        }

        public Sheet changeSheet(Sheet sh ,double w = 420, double h = 297)
        {
            tr = invApp.TransactionManager.StartTransaction((_Document)doc, "Добавить лист");
            sh = sheet(sh, new List<string> { "ГОСТ - Доп. графы 1", "ГОСТ - Доп. графы 2", "ГОСТ - Доп. графы 3", "АР", "Развертка 1:N АР" }, w, h);
            tr.End();
            return sh;
        }

        public Sheet sheet(Sheet sh, List<string> except, double w = 420, double h = 297, string tbName = "ГОСТ - Форма 2a", string[] prompt = null)
        {
            invApp.ScreenUpdating = false;
            DrawingDocument drw = ((DrawingDocument)doc);
            DrawingDocument temp = null;

            double oldW = sh.Width; double oldH = sh.Height;

            double offsetX = ((int)(w - oldW * 10)) * 0.1;
            double offsetY = ((int)(h - oldH * 10)) * 0.1;   
            BorderDefinition bordDef = null; TitleBlockDefinition tbd = null;
            string[] borders = new string[] { "ГОСТ - A4", "ГОСТ - A3", "ГОСТ - A2" };
            bool flag = true;
            //foreach (BorderDefinition item in drw.BorderDefinitions)
            //{
            //    if (borders.FirstOrDefault(e => e == item.Name) != null)
            //    { flag = false; break; }
            //}
            //if (flag)
            //{
            //    temp = openDrwDoc(pathTemplate[0]);
            //    copySketchSymbolDefinition(temp, drw, borders);
            //}
            if (h == 297 && w == 210) 
            {
                bordDef = drw.BorderDefinitions[border(drw, w, h, "А4")]; 
                sh.Size = DrawingSheetSizeEnum.kA4DrawingSheetSize; sh.Orientation = PageOrientationTypeEnum.kPortraitPageOrientation;
            }
            else if (h == 297 && w == 420) 
            {
                bordDef = drw.BorderDefinitions[border(drw, w, h, "А3")];
                sh.Size = DrawingSheetSizeEnum.kA3DrawingSheetSize; sh.Orientation = PageOrientationTypeEnum.kLandscapePageOrientation;
            }
            else if (h == 420 && w == 297) 
            {
                bordDef = drw.BorderDefinitions[border(drw, w, h, "А3")];
                sh.Size = DrawingSheetSizeEnum.kA3DrawingSheetSize; sh.Orientation = PageOrientationTypeEnum.kPortraitPageOrientation;
            }
            else if (h == 420 && w == 594) 
            {
                bordDef = drw.BorderDefinitions[border(drw, w, h, "А2")];
                sh.Size = DrawingSheetSizeEnum.kA2DrawingSheetSize; sh.Orientation = PageOrientationTypeEnum.kLandscapePageOrientation;
            }
            else if (h == 594 && w == 420) 
            {
                bordDef = drw.BorderDefinitions[border(drw, w, h, "А2")];
                sh.Size = DrawingSheetSizeEnum.kA2DrawingSheetSize; sh.Orientation = PageOrientationTypeEnum.kPortraitPageOrientation;
            }
            else if (h == 594 && w == 840) 
            {
                bordDef = drw.BorderDefinitions[border(drw, w, h, "А1")];
                sh.Size = DrawingSheetSizeEnum.kA2DrawingSheetSize; sh.Orientation = PageOrientationTypeEnum.kLandscapePageOrientation;
            }
            else if (h == 840 && w == 594) 
            {
                bordDef = drw.BorderDefinitions[border(drw, w, h, "А1")];
                sh.Size = DrawingSheetSizeEnum.kA2DrawingSheetSize; sh.Orientation = PageOrientationTypeEnum.kPortraitPageOrientation;
            }
            else
            {
                bordDef = drw.BorderDefinitions[border(drw, w, h)];
                sh.Size = DrawingSheetSizeEnum.kCustomDrawingSheetSize; sh.Orientation = PageOrientationTypeEnum.kLandscapePageOrientation;
                sh.Width = w * 0.1; sh.Height = h * 0.1;
            }
            if (sh.Border != null) sh.Border.Delete();

            sh.AddBorder(bordDef);
            try
            {
            if (sh.Name == "Лист:1") tbd = drw.TitleBlockDefinitions["ГОСТ - Форма 1"];
            else tbd = drw.TitleBlockDefinitions[tbName];
            }
            catch (Exception)
            {
                temp = temp ?? openDrwDoc(pathTemplate[0]);
                copySketchSymbolDefinition(temp, drw, new string [] { "ГОСТ - Форма 1", "ГОСТ - Форма 2a", tbName });
                //temp.TitleBlockDefinitions["ГОСТ - Форма 1"].CopyTo((_DrawingDocument)drw, true);
                //temp.TitleBlockDefinitions["ГОСТ - Форма 2a"].CopyTo((_DrawingDocument)drw, true);
                if (sh.Name == "Лист:1") tbd = drw.TitleBlockDefinitions["ГОСТ - Форма 1"];
                else tbd = drw.TitleBlockDefinitions[tbName];
                temp.Close();
            }

            if (sh.TitleBlock == null)
            {
                var tbs = u.gets<Inventor.TextBox>(tbd.Sketch.TextBoxes, e => e.FormattedText != null && e.FormattedText.Contains("Prompt"));
                if (tbs.Count() > 0)
                {
                    prompt = Enumerable.Repeat<string>("", tbs.Count()).ToArray();
                }
                //sh.TitleBlock.Delete();
                sh.AddTitleBlock(tbd, PromptStrings: prompt);
            }

            List<Point2d> poz = new List<Point2d> {invApp.TransientGeometry.CreatePoint2d(2,0.5), invApp.TransientGeometry.CreatePoint2d(2,h*0.1 - 12.5), 
                invApp.TransientGeometry.CreatePoint2d(2,h*0.1 - 0.5)}; 

            for (int i = 0; i < except.Count; i++)
			{
                SketchedSymbol ss = sh.SketchedSymbols.OfType<SketchedSymbol>().FirstOrDefault(e => e.Name == except[i]);
                if (ss == null && i <3)
                    try
                    {
                        SketchedSymbolDefinition def = drw.SketchedSymbolDefinitions[except[i]];
                        sh.SketchedSymbols.Add(def, poz[i]);
                    }
                    catch (Exception)
                    {
                    }     
            }

            //if (sh.SketchedSymbols.Count < 3)
            //{
            //    sh.SketchedSymbols.Add(drw.SketchedSymbolDefinitions["ГОСТ - Доп. графы 1"],invApp.TransientGeometry.CreatePoint2d(2,0.5));
            //    sh.SketchedSymbols.Add(drw.SketchedSymbolDefinitions["ГОСТ - Доп. графы 2"],invApp.TransientGeometry.CreatePoint2d(2,h*0.1 - 12.5));
            //    sh.SketchedSymbols.Add(drw.SketchedSymbolDefinitions["ГОСТ - Доп. графы 3"],invApp.TransientGeometry.CreatePoint2d(2,h*0.1 - 0.5));
            //}

            if (offsetX != 0 || offsetY != 0)
            {
                foreach (SketchedSymbol item in sh.SketchedSymbols)
                {
                    if (!except.Exists(e => e == item.Name))
                    {
                        Point2d pt = invApp.TransientGeometry.CreatePoint2d(item.Position.X + offsetX, item.Position.Y + offsetY);
                        item.Position = pt;
                    }
                    if (item.Name.EndsWith("графы 3"))
                    {
                        var pt = I.CP2d(2, sh.Height - 0.5);
                        item.Position = pt;
                    }
                }
                foreach (DrawingSketch item in sh.Sketches)
                {
                    if (item.Name == "Технические требования")
                    {
                        Macros.StandardAddInServer.m_inventorApplication.SilentOperation = !Macros.StandardAddInServer.m_inventorApplication.SilentOperation;
                        Macros.StandardAddInServer.m_inventorApplication.ScreenUpdating = !Macros.StandardAddInServer.m_inventorApplication.ScreenUpdating;
                        ObjectCollection col = Macros.StandardAddInServer.m_inventorApplication.TransientObjects.CreateObjectCollection();
                        Vector2d vec = Macros.StandardAddInServer.m_inventorApplication.TransientGeometry.CreateVector2d(offsetX, offsetY);
                        item.Edit();
                        foreach (TextBox tb in item.TextBoxes)
                        {
                            Point2d pt = invApp.TransientGeometry.CreatePoint2d(tb.Origin.X + offsetX, tb.Origin.Y);
                            tb.Origin = pt;
                            //col.Add(tb);
                        }
                        //item.MoveSketchObjects(col, vec);
                        item.ExitEdit();
                        Macros.StandardAddInServer.m_inventorApplication.SilentOperation = !Macros.StandardAddInServer.m_inventorApplication.SilentOperation;
                        Macros.StandardAddInServer.m_inventorApplication.ScreenUpdating = !Macros.StandardAddInServer.m_inventorApplication.ScreenUpdating;
                    }
                }
            }

            invApp.ScreenUpdating = true;
            return sh;
        }

        public string border(DrawingDocument drw ,double w = 420, double h = 297, string name = "")
        {
            Sheet sh = drw.ActiveSheet; BorderDefinition bordDef = null; DrawingSketch ds;
            if (name == "")
            {
                name = w + "x" + h;
                if (w % 210 == 0 && h == 297)
                {
                    name = "A4x" + w / 210;
                }
                else if (w % 297 == 0 && h == 420)
                {
                    name = "A3x" + w / 297;
                }
                else if (w % 420 == 0 && h == 594)
                {
                    name = "A2x" + w / 420;
                }
            }
            try
            {
                bordDef = drw.BorderDefinitions.Add("ГОСТ - " + name);
            }
            catch (Exception)
            {
                return "ГОСТ - " + name;
            }
            bordDef.Edit(out ds);
            SketchEntitiesEnumerator en = ds.SketchLines.AddAsTwoPointRectangle(invApp.TransientGeometry.CreatePoint2d(), invApp.TransientGeometry.CreatePoint2d(w / 10, h / 10));
            Layer lay = drw.StylesManager.Layers.OfType<Layer>().FirstOrDefault(la => la.LineWeight == 0.025);
            foreach (SketchEntity item in en)
            {
                item.Layer = lay;
            }
            en = ds.SketchLines.AddAsTwoPointRectangle(invApp.TransientGeometry.CreatePoint2d(2,0.5), invApp.TransientGeometry.CreatePoint2d(w / 10 - 0.5, h / 10 - 0.5));
            lay = drw.StylesManager.Layers.OfType<Layer>().FirstOrDefault(la => la.LineWeight == 0.05);
            foreach (SketchEntity item in en)
            {
                item.Layer = lay;
            }
            TextBox tb = ds.TextBoxes.AddFitted(invApp.TransientGeometry.CreatePoint2d(w/10 - 0.5, 0.02), "Формат " + name);
            tb.HorizontalJustification = HorizontalTextAlignmentEnum.kAlignTextRight;
            bordDef.ExitEdit();
            return "ГОСТ - " + name;
        }

        static public SketchedSymbol addSketchedSymbol(Sheet sh ,string name, string[] vals, Point2d pt, Vector2d vec = null, bool s = false)
        {
            DrawingDocument drw = (DrawingDocument)sh.Parent;
            try
            {
            SketchedSymbolDefinition ssd = drw.SketchedSymbolDefinitions[name];
            SketchedSymbol ss = sh.SketchedSymbols.Add(ssd, pt, PromptStrings: vals); ss.Static = s;   
            if (vec != null) ss.Position = u.translate(ss.Position, vec);
            return ss;
            }
            catch (Exception)
            {
                return null;
            }
        }

        public void replaceFile(FileDescriptor fd ,string fullName)
        {
            fd.ReplaceReference(fullName);
        }

        public BOMView getBOMView(bool firstLevelOnly = true, bool firstView = false, string t = "Структурированный", bool part = false, bool struc = true)
        {
            m_BOM = getAsmCompDef.BOM;
            if (firstView) return m_BOM.BOMViews[1];
            m_BOM.StructuredViewEnabled = struc;
            m_BOM.PartsOnlyViewEnabled = part;
            if (m_BOM.StructuredViewFirstLevelOnly != firstLevelOnly)
                m_BOM.StructuredViewFirstLevelOnly = firstLevelOnly;
            return m_BOM.BOMViews[t];
        }

        public List<string> openFiles(string filter, bool path1 = false)
        {
            List<string> names = new List<string>();
//             if (filter == "*.idw")
//                 nvmOptions.Add("DeferUpdates", true);
            pathes = new List<string>();
            if (pathes.Count() != 0 && !path1)
            {
                foreach (var p in pathes)
                {
                   names.AddRange(System.IO.Directory.GetFiles(p, filter, System.IO.SearchOption.AllDirectories));
                }
            }
            else names.AddRange(System.IO.Directory.GetFiles(path, filter, System.IO.SearchOption.AllDirectories));
            names = names.Where(n => n.IndexOf("Применяемые детали") == -1).ToList();
            return names.Where(n => n.IndexOf("OldVersions") == -1).ToList();
        }

        public bool exception(string name, List<string> excpt)
        {
            return excpt.Exists(e => name.ToLower().IndexOf(e) != -1) ? true : false;
        }

        public void findFiles(BOMRowsEnumerator rows, ref System.Collections.Generic.HashSet<Document> lstDoc, ref List<string> names, ref HashSet<string> pathes, string mainPath,List<string> except)
        {
            foreach (BOMRow row in rows)
            {
                if (row.ChildRows != null)
                {
                    findFiles(row.ChildRows, ref lstDoc, ref names, ref pathes, mainPath, except);
                }
                doc = (Document)row.ComponentDefinitions[1].Document;
                string path = doc.FullDocumentName.Remove(doc.FullDocumentName.LastIndexOf('\\'));
                if (System.IO.Directory.Exists(path + "\\PDF\\") && path != mainPath) pathes.Add(path + "\\PDF\\");
                if (System.IO.Directory.Exists(path + "\\DXF\\") && path != mainPath) pathes.Add(path + "\\DXF\\");
                addDocToList(doc,ref lstDoc,ref names, except);
            }
        }

        public void addDocToList(Document doc, ref System.Collections.Generic.HashSet<Document> lstDoc, ref List<string> names, List<string> except)
        {
            string pn = u.getProp(doc,"Part Number").Value.ToString();
            if (pn != "")
            {
                lstDoc.Add(doc);
                string name = names.FirstOrDefault(e => e.ToUpper().IndexOf(pn.ToUpper()) != -1);
                if (name != null)
                {
                    if (exception(name, except)) return;
                    lstDoc.Add(openDrwDoc(name, false) as Document);
                    //lstDoc.Add(openDoc(doc.FullFileName, false));
                }
            }
        }

        public static object addProp(Inventor.Document doc, string name, string val = "")
        {
            Property p;
            try{p = doc.PropertySets[3][name];}
            catch
            {
                try{ p = doc.PropertySets[4][name];}
                catch{p = doc.PropertySets[4].Add(val, name);}
            }
            try
            {
                if (val != p.Value.ToString())
                    p.Value = val;
            }
            catch { };
            //}
            return p.Value;
        }

        public Property getProp(string name)
        {
            Property p = null; 
            try { p = doc.PropertySets[2][name]; }
            catch
            {
                try { p = doc.PropertySets[3][name]; }
                catch
                {
                    try { p = doc.PropertySets[4][name]; }
                    catch { return null; }
                }
            }
            return p;
        }

        public void addProp(string name,string val) 
        { 
            Property p = getProp(name);
            if (p != null && p.Value.ToString() != val) p.Value = val;
        }

        static public void WithElem<T>(List<T> col, Action<T> act)
        {
            for (int i = 0; i < col.Count; i++)
            {
                act(col[i]);
            }
        }

        static public void withEnumAct<T>(IEnumerable<T> source, Action<T> act)
        {
            foreach (T element in source)
                act(element);
        }

        static public string getDir(string path, string name)
        {
            return (System.IO.Directory.Exists(path + name + "\\")) ? path + name + "\\" : System.IO.Directory.CreateDirectory(path + name + "\\").FullName;
        }

        static public string getPath(Document doc) { str l = s => s.Substring(0, s.LastIndexOf("\\") + 1);
        return l(doc.FullDocumentName);
        }
        static public LinearGeneralDimension addDim(DrawingView dv, DrawingCurve ent1, DrawingCurve ent2, Vector2d v, Point2d rpt)
        {
            Sheet sh = dv.Parent;
            GeometryIntent i1 = sh.CreateGeometryIntent(ent1, ent1.CenterPoint),
                i2 = sh.CreateGeometryIntent(ent2);
            Point2d pt = Drawings.midPoint(rpt, ent2.StartPoint, v, 0, 0);
            if (!InvDoc.u.eq(ent1.CenterPoint.X, rpt.X))
                return sh.DrawingDimensions.GeneralDimensions.AddLinear(pt, i1, i2, DimensionTypeEnum.kVerticalDimensionType);
            else if (!InvDoc.u.eq(ent1.CenterPoint.Y, rpt.Y))
                return sh.DrawingDimensions.GeneralDimensions.AddLinear(pt, i1, i2, DimensionTypeEnum.kHorizontalDimensionType);
            else
                return sh.DrawingDimensions.GeneralDimensions.AddLinear(pt, i1, i2, DimensionTypeEnum.kAlignedDimensionType);
        }
    }
    public static partial class u
    {
        public static Transaction tr = null;
        public static HighlightSet hs = null, hs2 = null;
        public static Parameter p;
        public static Document sdoc;
        public static Property getProp(Document doc, string name)
        {
            Property p = null;
            //p = ;
            p = get<Property>(doc.PropertySets[3], f => f.Name.ToLower() == name.ToLower()) ?? get<Property>(doc.PropertySets[4], f => f.Name.ToLower() == name.ToLower())
                ?? get<Property>(doc.PropertySets[1], f => f.Name.ToLower() == name.ToLower()) ?? get<Property>(doc.PropertySets[2], f => f.Name.ToLower() == name.ToLower());
            //             try { p = doc.PropertySets[2][name]; }
            //             catch
            //             {
            //                 try { p = doc.PropertySets[3][name]; }
            //                 catch
            //                 {
            //                     try { p = doc.PropertySets[4][name]; }
            //                     catch
            //                     {
            //                         try { p = doc.PropertySets[1][name]; }
            //                         catch { return null; }
            //                     }
            //                 }
            //             }
            return p;
        }
        public static void clear()
        {
            p = null;
            if (!isNull(sdoc)) { if (checkDoc(sdoc)) { sdoc.Close(); } sdoc = null; }
        }
        public static bool checkDoc(Document doc)
        {
            Document d = get<Document>(I.app.Documents.VisibleDocuments, f => f.Equals(doc));
            return isNull(d);
        }
        public static string getPropValue(Property p)
        {
            return p == null ? "" : p.Value.ToString();
        }

        public static string getPropValue(Document doc, string name)
        {
            Property p = getProp(doc, name);
            return p == null ? "" : p.Value.ToString();
        }

        public static Dictionary<string, string> getProps(Document doc, Dictionary<string, string> names)
        {
            Dictionary<string, string> dics = new Dictionary<string, string>();
            foreach (var item in names)
            {
                string tmp = getPropValue(doc, item.Value);
                if (tmp != "") dics[item.Key] = tmp;
            }
            return dics;
        }

        public static List<Property> getProps(Document doc, string[] names)
        {
            List<Property> col = new List<Property>();
            foreach (var item in names)
            {
                col.Add(getProp(doc, item));
            }
            return col;
        }

        public static Document referendedDoc(Document doc)
        {
            if (doc.ReferencedDocuments.Count == 1) return doc.ReferencedDocuments[1];
            Regex regex = new Regex(@"-\d*\w*\d*\w*\.\d*\.\d*(-\d\d)");
            foreach (Document refDoc in doc.ReferencedDocuments)
            {
                string n = System.IO.Path.GetFileNameWithoutExtension(refDoc.FullDocumentName);
                Match m = regex.Match(n);
                if (m != null && m.Groups[1].Value == "")
                    return refDoc;
            }
            return doc.ReferencedDocuments[1];
        }
        public static Document referendedDoc(DrawingDocument doc)
        {
            DrawingView dv = doc.Sheets[1].DrawingViews[1];
            if (dv == null) return null;
            return dv.ReferencedDocumentDescriptor.ReferencedDocument as Document;
        }
        public static Document referendedDoc(Document doc, string name)
        {
            Document tmp = null;
            if (doc.ReferencedDocuments.Count == 1) return doc.ReferencedDocuments[1];
            foreach (Document item in doc.ReferencedDocuments)
            {
                Property p = u.getProp(item, name);
                if (p != null) return item;
            }
            return doc.ReferencedDocuments[1];
        } 
        public static List<Document> referendedDocs(Document doc, string name)
        {
            List<Document> docs = new List<Document>();
            foreach (Document item in doc.ReferencedDocuments)
            {
                Property p = u.getProp(item, name);
                if (p != null) docs.Add(item);
            }
            return docs;
        }

        public static DocumentDescriptor referendedDocDesc(Document doc)
        {
            if (doc.ReferencedDocuments.Count == 1) return doc.ReferencedDocumentDescriptors[1];
            Regex regex = new Regex(@"-\d*\w*\d*\w\\.\d*\.\d*(-/d/d)");
            foreach (DocumentDescriptor refDoc in doc.ReferencedDocumentDescriptors)
            {
                Match m = regex.Match(System.IO.Path.GetFileNameWithoutExtension(refDoc.FullDocumentName));
                if (m != null && m.Groups[1].Value != "") return refDoc;
            }
            return doc.ReferencedDocumentDescriptors[1];
        }

        public static string getFormat(string v)
        {
            var spl = v.Split(':');
            spl = spl[0].Split('-');
            return spl[1].Trim();
        }

        public static bool changeProp(Document doc, string name, string val)
        {
            Property p = getProp(doc, name);
            if (p != null && p.Value.ToString() != val) { p.Value = val; return true; }
            return false;
        }

        public static bool changeProps(Document doc, List<string> names, List<string> vals, int end, int start = 0)
        {
            bool flag = false;
            if (names.Count != vals.Count) return flag;
            for (int i = start; i < end; i++)
            {
                if (changeProp(doc, names[i], vals[i])) flag = true;
            }
            return flag;
        }

        public static void addProp(Document doc, string name, object val, bool lib = true)
        {
            Property p = getProp(doc, name);
            try
            {
                if (p != null && val.ToString() != "" && p.Value.ToString() != val.ToString()) p.Value = val;
                else if (p == null)
                {
                    string fn = doc.FullFileName;
                    if (!file.check(fn)) return;
                    if (lib && fn.IndexOf("Content") != -1) return;

                    doc.PropertySets["Inventor User Defined Properties"].Add(val, name);
                }
            }
            catch (Exception)
            {
            }
        }

        public static void addProps(Document doc, string vals)
        {
            var tmp = vals.Split('$');
            foreach (var v in tmp)
            {
                var spl = v.Split('=');
                if (spl.Count() == 1) continue;
                addProp(doc, spl[0], spl[1]);
            }
        }

        static public object findAtPoint(SketchedSymbolDefinition ssd, Point2d pt)
        {
            foreach (Inventor.TextBox tb in ssd.Sketch.TextBoxes)
            {
                if (tb.RangeBox.Contains(pt)) return tb;
            }
            return null;
        }
        static public void mirrorPoint(Point pt, Plane pl)
        {
            double d1 = u.distToPlane(pl, pt, true);
            Vector v1 = pl.Normal.AsVector();
            v1.ScaleBy(-d1 * 2); 
            pt.TranslateBy(v1);
        }
        static public Box mirrorBox(Box b, Plane pl)
        {
            Point p1 = b.MinPoint, p2 = b.MaxPoint;
            mirrorPoint(p1, pl); mirrorPoint(p2, pl);
            var m = I.Box(p1, p2);
            return m;
        }
        public static List<Face> findFace(List<Point> pts, SurfaceBody sb)
        {
            List<Face> fs = new List<Face>();
            foreach (var pt in pts)
            {
                foreach (Face item in sb.Faces)
                {
                    if (item.SurfaceType != SurfaceTypeEnum.kPlaneSurface) continue;
                    var pl = item.Geometry as Plane;
                    var d = u.distToPlane(pl, pt);
                    if (u.eq(d, 0))
                    {
                        var b = item.Evaluator.RangeBox; b.Expand(0.1);
                        if (b.Contains(pt))
                            fs.Add(item);
                    }
                }
            }
            return fs;
        }
        static public List<Face> findFace(HoleFeature hf, SurfaceBody sb, Plane pl = null)
        {
            List<Face> fs = new List<Face>();
            List<Point> pts = new List<Point>();
            Face f = hf.Faces[1];
            foreach (Edge e in f.Edges)
            {
                if (e.CurveType != CurveTypeEnum.kCircleCurve && e.CurveType != CurveTypeEnum.kCircularArcCurve)
                    continue;
                //var c = e.Geometry as Circle;
                pts.Add(e.PointOnEdge);
            }
            if (pl != null)
            {
                foreach (var item in pts)
                {
                    mirrorPoint(item, pl);
                }
            }
            fs = findFace(pts, sb);
            return fs;
        }
        static public ObjectCollection get_bodies(HashSet<SurfaceBody> b)
        {
            var col = I.COC();
            foreach (var item in b)
            {
                col.Add(item);
            }
            return col;
        }
        static public ObjectCollection findBodies(PartComponentDefinition def, object f)
        {
            HashSet<SurfaceBody> bodies = new HashSet<SurfaceBody>();
            List<HoleFeature> hfs = new List<HoleFeature>();
            List<Box> boxes = new List<Box>();
            MirrorFeatureDefinition mir = f as MirrorFeatureDefinition;
            Plane plane = null;
            if (mir != null)
            {
                var pl = mir.MirrorPlaneEntity as WorkPlane;
                plane = pl.Plane;
                foreach (var item in mir.ParentFeatures)
                {
                    var hf = item as HoleFeature;
                    if (hf == null) continue;
                    hfs.Add(hf);
                } 
            }
            HoleFeature hf1 = f as HoleFeature;
            if (hf1 != null)
            {
                hfs.Add(hf1);
            }
            foreach (var item in hfs)
            {
                foreach (SurfaceBody body in def.SurfaceBodies)
                {
                    if (!body.Visible) continue;
                    var fs = findFace(item, body, plane);
                    if (fs.Count > 1)
                    {
                        bodies.Add(body);
                        break;
                    }
                }
            }
            return get_bodies(bodies);
        }
        static public void filter_line(PlanarSketch ps)
        {
            Dictionary<SketchLine, SketchLine> lines = new Dictionary<SketchLine, SketchLine>();
            for (int i = 0; i < ps.SketchLines.Count; i++)
            {
                var sl = ps.SketchLines[i + 1];
                var ie = u.gets<SketchLine>(ps.SketchLines, f => u.eq(sl, f));
                if (ie.Count() == 2)
                {
                    lines[ie.ElementAt(0)] = ie.ElementAt(1);
                }
            }
            foreach (var item in lines)
            {
                item.Key.Delete();
                item.Value.Delete();
            }
        }
        static public void merge_points(List<SketchPoint> pts)
        {
            List<IEnumerable<SketchPoint>> lst = new List<IEnumerable<SketchPoint>>();
            HashSet<SketchPoint> filter = new HashSet<SketchPoint>();
            for (int i = 1; i < pts.Count; i++)
            {
                if (filter.Contains(pts[i])) continue;
                var ie = u.gets<SketchPoint>(pts, f => u.eq(pts[0].Geometry, f.Geometry));
                if (ie.Count() > 1)
                {
                    lst.Add(ie);
                    foreach (var item in ie)
                    {
                        filter.Add(item);
                    }
                }
            }
            foreach (var item in lst)
            {
                var l = item.ToList();
                int c = l.Count;
                for (int i = 1; i < c; i++)
                {
                    if (l.Count < i) break;
                    l[i-1].Merge(l[i]);
                }
            }
        }
        static public void constr_line(PlanarSketch ps)
        {
            Dictionary<SketchPoint, SketchPoint> pts = new Dictionary<SketchPoint, SketchPoint>();
            List<HashSet<SketchPoint>> l_pts = new List<HashSet<SketchPoint>>();
            for (int i = 0; i < ps.SketchPoints.Count; i++)
            {
                var sp1 = ps.SketchPoints[i + 1]; 
                var ie = u.gets<SketchPoint>(ps.SketchPoints, f => u.eq(sp1.Geometry, f.Geometry, 3));
                if (ie.Count() == 2)
                {
                    pts[ie.ElementAt(0)] = ie.ElementAt(1);
                    ie.ElementAt(0).Merge(ie.ElementAt(1));
                }
                else if (ie.Count() > 2)
                {
                    HashSet<SketchPoint> hss = new HashSet<SketchPoint>();
                    foreach (var item in ie)
                    {
                        hss.Add(item);
                    }
                    l_pts.Add(hss);
                }
            }
            foreach (var item in l_pts)
            {
                int k = 0;
                SketchPoint sp = null;
                foreach (SketchPoint hs in item)
                {
                    if (k == 0) sp = hs;
                    else
                    {
                        sp.Merge(hs);
                    }
                    k++;
                }
            }
            foreach (var item in pts)
            {
                item.Key.Merge(item.Value);
            }

        }
        static public bool eq(Point pt1, Vector n1, Point pt2, Vector n2)
        {
            var par = n1.IsParallelTo(n2,0.01);
            if (!par) return false;
            var vec = pt1.VectorTo(pt2);
            var col = vec.DotProduct(n1);
            return eq(col, 0);
        }
        static public bool isNullVector(UnitVector vec)
        {
            return (Math.Round(vec.X, 3) == 0 && Math.Round(vec.Y, 3) == 0 && Math.Round(vec.Z, 3) == 0) ? true : false;
        }

        static public bool isNullVector(Vector vec)
        {
            return (Math.Round(vec.X, 3) == 0 && Math.Round(vec.Y, 3) == 0 && Math.Round(vec.Z, 3) == 0) ? true : false;
        }

        static public bool isNullVector(Vector2d vec)
        {
            return (Math.Round(vec.X, 3) == 0 && Math.Round(vec.Y, 3) == 0) ? true : false;
        }

        static public bool eq(Point pt1, UnitVector v1, Point pt2, UnitVector v2)
        {
            set(v1, ref pt1); set(v2, ref pt2);
            return eq(pt1, pt2);
        }

        static public bool eq(Box2d b1, Box2d b2, int tol = 3)
        {
            return eq(b1.MinPoint, b2.MinPoint, tol) && eq(b1.MaxPoint, b2.MaxPoint, tol);
        }
        static public bool eq(Box b1, Box b2, int tol = 3)
        {
            return eq(b1.MinPoint, b2.MinPoint, tol) && eq(b1.MaxPoint, b2.MaxPoint, tol);
        }

        static public bool eq(Edge e, double d)
        {
            return eq(getLenght(e), d);
        }

        static public double crossLenght(Point pt1, UnitVector v1, Point pt2)
        {
            Vector v = pt1.VectorTo(pt2);
            Vector cross = v.CrossProduct(v1.AsVector());
            return cross.Length;
        }

        static public bool eqScalar(Point pt1, UnitVector v1, Point pt2, double l = 1.5)
        {
            Vector v = pt1.VectorTo(pt2);
            if (!v.IsParallelTo(v1.AsVector())) return false;
            double d = Math.Abs(scalar(v, v1.AsVector()));
            return (d > l) ? false: true;
        }

        static public bool eq(Point pt1, UnitVector v1, Point pt2, double l = 1.5)
        {
            Vector v = pt1.VectorTo(pt2);
            double dp = v.DotProduct(v1.AsVector());
            //if (dp != 0) return false;
            //double d = Math.Abs(scalar(v, v1.AsVector()));
            if (Math.Abs(dp) > l) return false;
            if (isNullVector(v)) return false;
            //return true;
            return eq(v.AsUnitVector(), v1, true);
        }

        static public void set(UnitVector v, ref Point pt)
        {
            double [] coords = new double [3], vcoords = new double [3];
            pt.GetPointData(ref coords);
            v.GetUnitVectorData(ref vcoords);
            for (int i = 0; i < coords.Length; i++)
            {
                if(eq(vcoords[i],0)) continue;
                coords[i] = 0;
            }
            pt.PutPointData(coords);
        }

        static public void setFlangeOffset(Document doc)
        {
            var ss = doc.SelectSet;
            double offset = 20;
            if (ss.Count == 0) return;
            foreach (var item in ss)
            {
                var fl = item as FlangeFeature;
                if (fl == null) return;
                var def = fl.Definition;
                fl.SetEndOfPart(true);
                foreach (var el in def.Edges)
                {
                    var e = el as Edge;
                    def.SetOffsetWidthExtent(e, e.StartVertex, offset / 10, e.StopVertex, offset / 10);
                }
                fl.SetEndOfPart(false);
            }
        }


        static public void higlight(Document doc, object ent, object ent2, Inventor.Color col, Inventor.Color col2)
        {
            col = col ?? I.app.TransientObjects.CreateColor(255, 0, 0);
            col2 = col2 ?? I.app.TransientObjects.CreateColor(0, 255, 0);
            if (ent != null)
            {
                if (hs == null)
                    hs = doc.CreateHighlightSet();
                hs.Color = col;
                hs.AddItem(ent);
            }

            if (ent2 != null)
            {
                if (hs2 == null)
                    hs2 = doc.CreateHighlightSet();
                hs2.Color = col2;
                hs2.AddItem(ent2);
            }
        }

        static public DrawingView findDV(Sheet sh, string name)
        {
            foreach (DrawingView item in sh.DrawingViews)
            {
                if (item.Name == name) return item; 
            }
            return null;
        }

        static public DrawingView findDV(Sheet sh, Point2d pt)
        {
            foreach (DrawingView item in sh.DrawingViews)
            {
                if (u.inRange(pt, item.Center, item.Width, item.Height)) return item;
            }
            return null;
        }


        static public bool intent(Sheet sh, DrawingCurve dc)
        {
            foreach (Balloon item in sh.Balloons)
            {
                GeometryIntent i = item.Leader.AllNodes[2].AttachedEntity;
                if (i == null) continue;
                if (i.Geometry.Equals(dc)) return true;
            }
            return false;
        }

        static public Balloon addBallon(Sheet sh, List<Point2d> pts, string vName, CurveTypeEnum type)
        {
            ObjectCollection col = I.COC();
            DrawingView dv = findDV(sh, vName);
            if (dv == null) return null;
            Point2d pt;
            //foreach (var item in pts)
            //{
                pt = I.CP2d(dv.Center.X + pts[0].X, dv.Center.Y + pts[0].Y);
                col.Add(pt);
            //}
                pt = I.CP2d(dv.Center.X + pts[1].X, dv.Center.Y + pts[1].Y);
               //col.Add(pt);
//             DrawingSketch ds = sh.Sketches.Add();
//             ds.Edit();
//             ds.SketchPoints.Add(pt);
//             ds.ExitEdit();
                
            DrawingCurve dc = u.findAtPoint(sh, pt, e => e.CurveType == type, 0);
            if (dc == null) return null;
            GeometryIntent i = null;
            //double p = u.getParamAtPoint(dc.Evaluator2D, pt);
            i = sh.CreateGeometryIntent(dc,pt);
            col.Add(i);
            if (intent(sh, dc)) return null;
//             try
//             {
                return sh.Balloons.Add(col);
//             }
//             catch (System.Exception ex)
//             {
//             	
//             }
            return null;
        }

        static public List<Point2d> getPts(Balloon b)
        {
            List<Point2d> pts = new List<Point2d>();
            DrawingView dv = b.ParentView;
            foreach (LeaderNode item in b.Leader.AllNodes)
            {
                Point2d pt = I.CP2d(item.Position.X - dv.Center.X, item.Position.Y - dv.Center.Y);
                pts.Add(pt);
            }
            return pts;
        }

        static public void clearHiglight()
        {
            if (hs != null)
                hs.Clear();
            if (hs2 != null)
                hs2.Clear();
        }

        static public object findAtPoint(PartComponentDefinition compDef, Point pt, SelectionFilterEnum[] filter, double tol = 15, bool axis = false)
        {
            ObjectsEnumerator en = compDef.FindUsingPoint(pt, ref filter, tol);
            MeasureTools meas = I.app.MeasureTools;
            if (en.Count == 0) return null;
            object ax = null;
            ax = en.OfType<Edge>().OrderBy(e => meas.GetMinimumDistance(pt, e)).First();
            if (axis)
                return axisFromEdge(compDef, ax);
            return ax;
        }
        static public object findAtPoint(PartComponentDefinition def, Point pt, SurfaceBody sb, SelectionFilterEnum[] filter, double tol = 15)
        {
            ObjectsEnumerator en = def.FindUsingPoint(pt, filter, tol, true);
            if (en.Count == 0) return null;
            return en.OfType<Edge>().Where(el => el.Parent.Equals(sb) && el.TangentiallyConnectedEdges.Count == 0).FirstOrDefault();
        }

        static public ObjectsEnumerator findAtPoint(AssemblyComponentDefinition compDef, Point pt, SelectionFilterEnum[] filter, double tol = 15)
        {
            var en = compDef.FindUsingPoint(pt, ref filter, tol);
            if (en.Count == 0) return null;
            return en;
        }

        static public bool findAtPoint(AssemblyComponentDefinition compDef, Point pt, double tol = 1.5)
        {
            SelectionFilterEnum[] f = new SelectionFilterEnum[] { SelectionFilterEnum.kAssemblyOccurrenceFilter };
            var en = findAtPoint(compDef, pt, f, tol);
            if (!u.isNull(en) && en.Count > 1)
            {
                foreach (ComponentOccurrence item in en)
                {
                    if (item.ReferencedDocumentDescriptor.FullDocumentName.IndexOf("Content Center Files") != -1) return false;
                }
                return true;
            }
            else return true;
            
        }

        static public T findAtRay<T>(PartComponentDefinition compDef, Point pt, UnitVector dir, SelectionFilterEnum[] filter, double tol = 15, int ind = 0)
        {
            object locPts = null;
            ObjectsEnumerator en = compDef.FindUsingVector(pt, dir, ref filter, true, tol, true, out locPts);
            MeasureTools meas = I.app.MeasureTools;
            if (en.Count == 0) return default(T);
            T ax = en.OfType<T>().OrderBy(e => meas.GetMinimumDistance(pt, e)).ElementAt(ind);
            return ax;
        }
        static public IEnumerable<DrawingCurve> findAtVector(DrawingView dv, Point2d pt, Vector2d v, 
            double r = 0)
        {
            var m = I.getMatrix2d();
            var n = v.Copy(); normal(n, null);
            m.SetToAlignCoordinateSystems(pt, v, n, I.CP2d(), I.CV2d(1, 0), I.CV2d(0, 1));
            var dcs = gets<DrawingCurve>(dv.DrawingCurves, e => e.CurveType == CurveTypeEnum.kCircleCurve ||
            e.CurveType == CurveTypeEnum.kCircularArcCurve);
            foreach (DrawingCurve dc in dcs)
            {
                var cen = dc.CenterPoint;
                cen.TransformBy(m);
                if (cen.X == 0 && cen.Y == 0) continue;
                if (r != 0 && !u.eq(getRadius(dc), r)) continue;
                if (eq(cen.Y, 0)) yield return dc;
            }
        }
        static public double getRadius(DrawingCurve dc)
        {
            dynamic e = dc.ModelGeometry;
            return e.Geometry.Radius;
        }
        static public bool filterByVector(DrawingCurve dc, Vector2d v)
        {
            LineSegment2d l = dc.Segments[1] as LineSegment2d;
            if (l == null) return false;
            var v2 = l.StartPoint.VectorTo(l.EndPoint);
            if (v.IsParallelTo(v2)) return false;
            return true;
        }
        static public Vector2d findMinDist(DrawingView dv, Point2d pt, Vector2d v)
        {
            Point2d min = getPoint(dv, true), max = getPoint(dv, false);
            Vector2d n1, n2; //normal(n, null);
            u.normal(v, pt.VectorTo(min), out n1);
            u.normal(v, pt.VectorTo(max), out n2);
            var l = n1.Length + n2.Length;
            return n1.Length <= n2.Length ? n1: n2;
        }
        static public Point2d getPoint(DrawingView dv, bool min)
        {
            Point2d pt = dv.Center.Copy();
            var vec = min ? I.CV2d(-dv.Width / 2, -dv.Height / 2): I.CV2d(dv.Width / 2, dv.Height / 2);
            pt.TranslateBy(vec);
            return pt;
        }
        static public void findUsingRay(PartComponentDefinition def, Point pt, UnitVector v, 
             Action<object, object> a, double tol = 0.01, bool vis = false)
        {
            ObjectsEnumerator o, l;
            def.FindUsingRay(pt, v, tol, out o, out l);
            for (int i = 1; i < o.Count+1; i++)
            {
                if (vis)
                {
                    Face fa = o as Face;
                    if (fa != null)
                    {
                        if (!fa.SurfaceBody.Visible) continue;
                    }
                }
                a(o[i], l[i]);
            }
        }
        public static object getValue(FeatureDimensions dims, FeatureDimensionTypeEnum t)
        {
            foreach (FeatureDimension item in dims)
            {
                if (item.DimensionType == t) return item.Parameter.ModelValue;
            }
            return null;
        }
        public static void booleanCut(Document doc, SurfaceBody b, SurfaceBody tool, PlanarSketch ps)
        {
            I.screenSilent(true);
            var def = I.getPCD(doc);
            var tb = I.app.TransientBRep;
            var transBase = tb.Copy(b);
            var transTool = tb.Copy(tool);
            var mtx = I.getMatrix();
            var lastPt = I.CP();
            var pts = addArray(ps, 10, 10, 20, 20);
            foreach (Point pt in pts)
            {
                mtx.Cell[1, 4] = pt.X - lastPt.X;
                mtx.Cell[2, 4] = pt.Y - lastPt.Y;
                mtx.Cell[3, 4] = pt.Z - lastPt.Z;
                tb.Transform(transTool, mtx);
                tb.DoBoolean(transBase, transTool, BooleanTypeEnum.kBooleanTypeDifference);
                lastPt = pt;
            }
            var npfs = def.Features.NonParametricBaseFeatures;
            var npdef = npfs.CreateDefinition();
            var col = I.COC();
            col.Add(transBase);
            npdef.BRepEntities = col;
            npdef.OutputType = BaseFeatureOutputTypeEnum.kSolidOutputType;
            npfs.AddByDefinition(npdef);
            tool.Visible = false;
            b.Visible = false;
            I.screenSilent(false);
            I.app.ActiveView.Update();
        }
        public static List<Point> addArray(PlanarSketch ps, double dx, double dy, int cx, int cy)
        {
            var last = u.get<SketchPoint>(ps.SketchPoints, f => f.HoleCenter).Geometry;
            List<Point> pts = new List<Point>();
            for (int i = 0; i < cx; i++)
            {
                for (int j = 0; j < cy; j++)
                {
                    pts.Add(ps.SketchToModelSpace(I.CP2d(last.X + i * dx, last.Y + j * dy)));
                    //ps.SketchPoints.Add(I.CP2d(last.X + i * dx, last.Y + j * dy));
                }
            }
            return pts;
        }
        public static void changePlane(ref WorkPlane wp, PartComponentDefinition def)
        {
            if (wp == null) wp = def.WorkPlanes[1];
            else
            {
                bool n = false;
                var wps = def.WorkPlanes;
                for (int i = 1; i <= wps.Count + 1; i++)
                {
                    if (i == wps.Count + 1)
                    {
                        wp = null; u.clearHiglight(); return;
                    }
                    if (wps[i].Equals(wp)) { n = true; continue; }
                    if (n) { wp = wps[i]; break; }
                }
            }
            setWPlane(wp, def.Document as Document);
        }
        static public void setWPlane(WorkPlane wp, Document doc)
        {
            u.clearHiglight();
            u.higlight(doc, wp, null, null, null);
        }
        static public MirrorFeature addMirror(Document doc, WorkPlane pl, List<PartFeature> pfs)
        {
            SheetMetalComponentDefinition smcd = I.getSMCD(doc);
            ObjectCollection col = I.COC();
            foreach (PartFeature item in pfs)
            {
                col.Add(item);
            }
            SheetMetalFeatures smf = smcd.Features as SheetMetalFeatures;
#if INV14
return smf.MirrorFeatures.Add(col, pl, false, PatternComputeTypeEnum.kIdenticalCompute);
#else
            var def = smf.MirrorFeatures.CreateDefinition(col, pl, PatternComputeTypeEnum.kIdenticalCompute);
            return smf.MirrorFeatures.AddByDefinition(def);
#endif
            //smf.MirrorFeatures.Add(col, pl, false, PatternComputeTypeEnum.kIdenticalCompute);
        }
        static public HoleFeature addHole(SheetMetalComponentDefinition smcd, PartFeature pf)
        {
            HoleFeature hf = pf as HoleFeature;
            if (hf != null) return addHole(smcd, hf);
            PunchToolFeature punch = pf as PunchToolFeature;
            if (punch != null) return addHole(smcd, punch);
            return null;
        }
        static public void changeDiam(PartFeature pf, ref double d)
        {
            var spl = pf.Name.Split('$');
            if (spl.Count() == 2)
            {
                d = u.convToDouble(spl[1])*0.1;
            }
        }
        static public void changeDir(HoleFeature f, PartFeatureExtentDirectionEnum dir)
        {
            if (f.HealthStatus == HealthStatusEnum.kDriverLostHealth)
            {
                dir = changeDir(dir);
#if INV14
                f.SetDistanceExtent(f.Parameters[2], dir);
#else
                f.SetDistanceExtent(f.Depth, dir);
#endif
            }
        }
        static public HoleFeature addHole(SheetMetalComponentDefinition smcd, PunchToolFeature punch)
        {
            PartFeatureExtentDirectionEnum dir = PartFeatureExtentDirectionEnum.kNegativeExtentDirection;
            ObjectCollection col = I.COC(), colBodies = I.COC();
            HashSet<SurfaceBody> bodies = new HashSet<SurfaceBody>();
            Vector v = I.CV();
            col = findUsingVector(smcd, punch, dir, ref bodies, ref v);
            if (col.Count == 0)
            {
                dir = changeDir(dir);
                col = findUsingVector(smcd, punch, dir, ref bodies, ref v);
            }
            if (col.Count == 0) return null;
            SheetMetalFeatures smf = smcd.Features as SheetMetalFeatures;
            var def = smf.HoleFeatures.CreateSketchPlacementDefinition(col);
            var d = 0.63;
            changeDiam(punch as PartFeature, ref d);
            var f = smf.HoleFeatures.AddDrilledByDistanceExtent(def, d, v.Length * 1.2, dir);
            //var ext = f.Extent as DistanceExtent;
            //ext.Direction = PartFeatureExtentDirectionEnum.kPositiveExtentDirection;
            foreach (var item in bodies) { colBodies.Add(item); }
            f.SetAffectedBodies(colBodies);
            changeDir(f, dir);
            return f;
        }
        static public HoleFeature addHole(SheetMetalComponentDefinition smcd, HoleFeature hf)
        {
            PartFeatureExtentDirectionEnum dir = PartFeatureExtentDirectionEnum.kNegativeExtentDirection;
            ObjectCollection col = I.COC(), colBodies = I.COC();
            HashSet<SurfaceBody> bodies = new HashSet<SurfaceBody>();
            Vector v = I.CV();
            col = findUsingVector(smcd, hf, dir, ref bodies, ref v);
            if (col.Count == 0)
            {
                dir = changeDir(dir);
                col = findUsingVector(smcd, hf, dir, ref bodies, ref v);
            }
            if (col.Count == 0) return null;
            SheetMetalFeatures smf = smcd.Features as SheetMetalFeatures;
            var def = smf.HoleFeatures.CreateSketchPlacementDefinition(col);
            var d = hf.HoleDiameter.ModelValue;
            changeDiam(hf as PartFeature, ref d);
            var f = smf.HoleFeatures.AddDrilledByDistanceExtent(def, d, v.Length*1.2, dir);
            //var ext = f.Extent as DistanceExtent;
            //ext.Direction = PartFeatureExtentDirectionEnum.kPositiveExtentDirection;
            foreach (var item in bodies) { colBodies.Add(item); }
            f.SetAffectedBodies(colBodies);
            changeDir(f, dir);
            return f;
        }
        static public void changePFName(PartFeature from, PartFeature to)
        {
            var spl = from.Name.Split('$');
            if (spl.Count() == 2) to.Name = to.Name + "$" + spl[1];
        }
        static public PartFeatureExtentDirectionEnum changeDir(PartFeatureExtentDirectionEnum dir)
        {
            return dir == PartFeatureExtentDirectionEnum.kNegativeExtentDirection ?
                PartFeatureExtentDirectionEnum.kPositiveExtentDirection :
                PartFeatureExtentDirectionEnum.kNegativeExtentDirection;
        }
        static public ObjectCollection findUsingVector(SheetMetalComponentDefinition smcd, HoleFeature hf, 
            PartFeatureExtentDirectionEnum dir, ref HashSet<SurfaceBody> bodies, ref Vector v)
        {
            ObjectCollection col = I.COC();
            int n = 0;
            foreach (Face f in hf.Faces)
            {
                n++;
                Cylinder cyl = f.Geometry as Cylinder;
                var pt = cyl.BasePoint;
                var vec = cyl.AxisVector;
                if (dir == PartFeatureExtentDirectionEnum.kPositiveExtentDirection) vec = reverse(vec);
                if (findUsingVector(smcd, pt, vec, ref bodies, ref v)){ col.Add(hf.HoleCenterPoints[n]); }
            }
            return col;
        }
        static public ObjectCollection findUsingVector(SheetMetalComponentDefinition smcd, PunchToolFeature punch,
            PartFeatureExtentDirectionEnum dir, ref HashSet<SurfaceBody> bodies, ref Vector v)
        {
            ObjectCollection col = I.COC();
            int n = 0;
            var vec = getPlane(punch.PunchCenterPoints[1] as SketchPoint).Normal;
            foreach (SketchPoint sp in punch.PunchCenterPoints)
            {
                n++;
                var pt = sp.Geometry3d;
                if (dir == PartFeatureExtentDirectionEnum.kPositiveExtentDirection) vec = reverse(vec);
                if (findUsingVector(smcd, pt, vec, ref bodies, ref v)) 
                { 
                    col.Add(punch.PunchCenterPoints[n]); 
                }
            }
            return col;
        }
        static public Plane getPlane(SketchPoint sp)
        {
            PlanarSketch ps = sp.Parent as PlanarSketch;
            return ps.PlanarEntityGeometry;
        }
        static public bool findUsingVector(SheetMetalComponentDefinition smcd, Point pt, UnitVector uv,
            ref HashSet<SurfaceBody> bodies, ref Vector v, double l = 5)
        {
            SelectionFilterEnum[] filter = new SelectionFilterEnum[] { SelectionFilterEnum.kPartFacePlanarFilter };
            object obj;
            var el = smcd.FindUsingVector(pt, uv, ref filter, true, 0.01, true, out obj);
            ObjectsEnumerator loc = obj as ObjectsEnumerator;
            int n = 0;
            foreach (Point item in loc)
            {
                n++;
                Face f = el[n] as Face;
                Plane pl = f.Geometry as Plane; 
                if (n % 2 != 0) continue;
                var vec = pt.VectorTo(item);
                if (vec.Length > l) continue;
                if (vec.Length > v.Length) v = vec;
                bodies.Add(f.SurfaceBody);
                return true;
                //clientTxt(smcd as PartComponentDefinition, item, n.ToString());
            }
            return false;
        }
        static public void findUsingRay(AssemblyComponentDefinition def, Point pt, UnitVector v,
             Action<object, object> a, double tol = 0.01)
        {
            ObjectsEnumerator o, l;
            def.FindUsingRay(pt, v, tol, out o, out l);
            for (int i = 1; i < o.Count + 1; i++)
            {
                a(o[i], l[i]);
            }
        }
        static public void regex(ref string f, string reg, string repl)
        {
            Regex r = new Regex(reg, RegexOptions.IgnoreCase);
            if (r.IsMatch(f))
            {
                f = Regex.Replace(f, reg, repl);
            }
        }

        static public bool regex(string f, string reg)
        {
            Regex r = new Regex(reg, RegexOptions.IgnoreCase);
            return r.IsMatch(f);
        }

        static public string getDate(string s)
        {
            if (s == "") return "";
            return s.Substring(s.Length - 2);
        }

        static public Path createPath(PartDocument doc, object ent)
        {
            return doc.ComponentDefinition.Features.CreatePath(ent);
        }

        static public R createProxy<R>(AssemblyDocument doc, ComponentOccurrence oc, object p)
        {
            object res;
            oc.CreateGeometryProxy(p, out res);
            return (R)res;
        }

        static public string getRefkey<T>(Document doc, T ent, int ind = 0)
        {
            byte[] oRefKey = new byte[] { };
            oRefKey = Reflect.runMethod<T, byte[]>(ent, "GetReferenceKey", new object[] { oRefKey, 0 }, ind);
            return doc.ReferenceKeyManager.KeyToString(oRefKey);
        }

        static public object bindRefkey(Document doc, string str)
        {
            byte[] oRefKey = new byte[] { };
            doc.ReferenceKeyManager.StringToKey(str, ref oRefKey);
            object obRef, obContext;
            if (doc.ReferenceKeyManager.CanBindKeyToObject(oRefKey, 0, out obRef, out obContext))
                return obRef;
            return null;
        }

        static public object axisFromEdge(PartComponentDefinition compDef, object ax)
        {
            UnitVector vec = null;
            if (ax is WorkAxis) return ax;
            if (ax is Edge)
            {
                Edge ed = ax as Edge;
                vec = ed.StopVertex.Point.VectorTo(ed.StartVertex.Point).AsUnitVector();
            }
            if (ax is UnitVector) vec = ax as UnitVector;
            foreach (WorkAxis item in compDef.WorkAxes)
            {
                UnitVector vec1 = item.Line.Direction;
                if (vec1.IsParallelTo(vec))
                {
                    return item;
                }
            }
            foreach (Edge item in compDef.SurfaceBodies[1].Edges)
            {
                if (item.GeometryType != CurveTypeEnum.kLineSegmentCurve) continue;
                UnitVector vec1 = (item.Geometry as LineSegment).Direction;
                if (vec.IsParallelTo(vec1, 0.01))
                    return item;
            }
            return null;
        }

        static public object mirrorDir(PartComponentDefinition compDef, UnitVector vec, WorkPlane pl)
        {
            vec = mirrorVector(vec, pl.Plane);
            return axisFromEdge(compDef, vec);
        }

        static public UnitVector mirrorVector(UnitVector vec, Plane pl)
        {
            Matrix mtx = I.tg.CreateMatrix();
            UnitVector norm = pl.Normal;
            for (int i = 1; i < 5; i++)
            {
                mtx.set_Cell(i, i, 1);
            }
            if (!eq(norm.X, 0)) mtx.set_Cell(1, 1, -norm.X);
            if (!eq(norm.Y, 0)) mtx.set_Cell(2, 2, -norm.Y);
            if (!eq(norm.Z, 0)) mtx.set_Cell(3, 3, -norm.Z);
            vec.TransformBy(mtx);
            return vec;
        }

        static public WorkPlane planeFromAxis(PartComponentDefinition compDef, UnitVector dir)
        {
            foreach (WorkPlane item in compDef.WorkPlanes)
            {
                if (item.Plane.Normal.IsParallelTo(dir)) return item;
            }
            return null;
        }

        static public bool translit(Document doc, string name, Dictionary<string, string> dic)
        {
            Property p = getProp(doc, name);
            if (p != null && dic.ContainsKey(p.Value.ToString()))
            {
                p.Value = dic[p.Value.ToString()];
                return true;
            }
            return false;
        }

        static public void translit(DocumentsEnumerator docs, string name, Dictionary<string, string> dic)
        {
            foreach (Document doc in docs)
            {
                if (doc.ReferencedDocuments.Count > 1)
                {
                    translit(doc.ReferencedDocuments, name, dic);
                }
                translit(doc, name, dic);
            }
        }


        static public void addNameToFeature<T>(T feature, string name) where T : PartFeature
        {

            if (name != "" && feature.Name != name) feature.Name = name;
        }

        static public object findAtPoint(Sheet sh, Point2d pt, ref Point2d retPt, ref double dist, double tol = 2, bool x = true)
        {
            ObjectsEnumerator en = sh.FindUsingPoint(pt, tol); IEnumerable<DrawingCurveSegment> ie = null;
            MeasureTools meas = I.app.MeasureTools;
            if (en.Count == 0) return null;
            if (x)
                ie = en.OfType<DrawingCurveSegment>().Where(s => s.GeometryType == Curve2dTypeEnum.kLineSegmentCurve2d && isVertical(s));
            else ie = en.OfType<DrawingCurveSegment>().Where(s => s.GeometryType == Curve2dTypeEnum.kLineSegmentCurve2d && !isVertical(s));
            object ax = null; Point2d ret = null;
            if (ie.Count() == 0) return null;
            ax = ie.OrderBy(e => getDist(pt, e.Parent, 1000, ref ret)).First();
            dist = getDist(pt, ((DrawingCurveSegment)ax).Parent, 1000, ref ret);
            if ((ret.X - pt.X + ret.Y - pt.Y) < 0)
                dist = -dist;
            retPt = ret;
            return ax;
        }

        static public DrawingCurve findAtPoint(Sheet sh, Point2d pt, Func<DrawingCurve, bool> f, double tol = 2)
        {
            ObjectsEnumerator en = sh.FindUsingPoint(pt, tol);
            foreach (var item in en)
            {
                DrawingCurve dc = (item as DrawingCurveSegment).Parent;
                if (f(dc)) return dc;
            }
            return null;
        }

        static public DrawingCurve findAtPoint(Sheet sh, Point2d pt, double tol, CurveTypeEnum type)
        {
            if (pt == null) return null;
            DrawingView dv = u.findDV(sh, pt);
            DrawingCurve dc = null;
            foreach (DrawingCurve item in dv.DrawingCurves)
            {
                if (item.CurveType != type) continue;
                if (item.StartPoint != null && item.EndPoint != null) dc = findAtPoint(dv, pt, tol, f => f.CurveType == item.CurveType);
                if (dc != null) return dc;
                if (item.CenterPoint != null && u.inRange(pt, item.CenterPoint, tol, tol)) return item;
            }
            return null;
        }

        static public Point2d getNearPoint(DrawingCurve dc, Point2d pt)
        {
            double sdist, edist;
            if (dc.ProjectedCurveType == Curve2dTypeEnum.kLineSegmentCurve2d)
            {
                sdist = pt.DistanceTo(dc.StartPoint);
                edist = pt.DistanceTo(dc.EndPoint);
                return sdist < edist ? dc.StartPoint : dc.EndPoint;
            }
            else if (dc.ProjectedCurveType == Curve2dTypeEnum.kCircularArcCurve2d || dc.ProjectedCurveType == Curve2dTypeEnum.kCircleCurve2d)
            {
                return dc.CenterPoint;
            }
            return null;
        }


        static public Vector2d getNormal(DrawingCurve dc)
        {
            if (dc.ProjectedCurveType != Curve2dTypeEnum.kLineSegmentCurve2d) return null;
            Vector2d v = dc.StartPoint.VectorTo(dc.EndPoint);
            v.Normalize();
            return v;
        }

        static public Point2d spiralPosition(DrawingView dv, Sheet sh, double x, double y, Point2d mpt, ref List<Point2d> ptsExc)
        {
            Vector2d vec = I.tg.CreateVector2d(x, y);
            if (mpt.X > dv.Position.X || mpt.Y > dv.Position.Y)
            {
                vec = I.tg.CreateVector2d(-x, -y);
            }
            Point2d pt = mpt.Copy();
            //invApp.TransientGeometry.CreatePoint2d(midPt.X + x, midPt.Y + y);
            ObjectsEnumerator ob; double tol = 0.5;
            pt.TranslateBy(vec);
            ob = sh.FindUsingPoint(pt, tol);
            int i = 0; double ofset = 0.2;

            while (ob.Count != 0 || ptsExc.Exists(e => InvDoc.u.eq(e, pt)))
            {
                i++;
                vec = rotate(vec, mpt, Math.PI / 2);
                pt = mpt.Copy();
                pt.TranslateBy(vec);
                //ptsExc.Add(pt);
                ob = sh.FindUsingPoint(pt, tol);
                if (i % 4 == 0)
                {
                    vec = I.tg.CreateVector2d(x + ofset, y + ofset);
                    ofset += 0.2;
                }
            }
            ptsExc.Add(pt);
            return pt;
        }

        static public UnitVector reverse(UnitVector v)
        {
            return I.CV(0-v.X,0- v.Y, 0-v.Z).AsUnitVector();
        }

        static public double reverse(double d)
        {
            return 0 - d;
        }

        static public UnitVector2d vecProjection(Point2d origin, Point2d pt, bool x = true)
        {
            Vector2d vec = origin.VectorTo(pt);
            if (x)
            {
                vec.X = 0;
            }
            else vec.Y = 0;
            return vec.AsUnitVector();
        }
        static public Point ProjectPointOntoVector(Point pt, Vector v)
        {
            // Calculate the dot product of the point and the vector
            double dotProduct = pt.X * v.X + pt.Y * v.Y + pt.Z * v.Z;

            // Calculate the magnitude of the vector
            double vectorMagnitude = Math.Sqrt(v.X * v.X + v.Y * v.Y + v.Z * v.Z);

            // Calculate the projection of the point onto the vector
            double projectionX = (dotProduct / (vectorMagnitude * vectorMagnitude)) * v.X;
            double projectionY = (dotProduct / (vectorMagnitude * vectorMagnitude)) * v.Y;
            double projectionZ = (dotProduct / (vectorMagnitude * vectorMagnitude)) * v.Z;

            return I.CP(projectionX, projectionY, projectionZ);
        }

        static public double getParam(Point pt, Point sp, Point ep)
        {
            Vector v = ep.VectorTo(sp), v1 = pt.VectorTo(sp);
            double dot = v.DotProduct(v), dot1 = v1.DotProduct(v);
            return dot1 / dot;
        }
        static public double distToEdge(Point pt, Point sp, Point ep, double t)
        {
            Vector v = sp.VectorTo(ep);
            v.ScaleBy(t); sp.TranslateBy(v);
            Vector v1 = pt.VectorTo(sp);
            return v1.Length;
        }

        static public bool isPointOnEdge(Point pt, Point sp, Point ep, double tol = 0.01)
        {
            double t = getParam(pt, sp, ep);
            if (t >= 0 && t <= 1)
            {
                double d = distToEdge(pt, sp, ep, t);
                return d < tol;
            }
            return false;
        }

        static public bool isPointOnEdge(Point pt, Edge ed, double tol = 0.01)
        {
            return isPointOnEdge(pt, ed.StartVertex.Point, ed.StopVertex.Point, tol);
        }

        static public bool directionIs(Vector2d v1, Vector2d v2, bool x = true)
        {
            if (x)
                return (Math.Sign(v1.X) == Math.Sign(v2.X));
            else
                return (Math.Sign(v1.Y) == Math.Sign(v2.Y));
        }

        static public bool isCollinear(UnitVector2d vec1, UnitVector2d vec2)
        {
            return u.eq(vec1, vec2);
        }
        static public bool isCollinear(UnitVector v1, UnitVector v2)
        {
            return u.eq(v1, v2, 0.01);
        }

        public static void faceToDXF()
        {
            var doc = I.aDoc();
            var ss = doc.SelectSet;
            if (ss.Count != 1) return;
            Face f = ss[1] as Face;
            if (f == null) return;
            if (f.SurfaceType != SurfaceTypeEnum.kPlaneSurface) return;
            Plane pl = f.Geometry as Plane;
            var norm = pl.Normal;
            List<Face> fs = new List<Face>();
            var def = I.getPCD(doc);
            foreach (Face fa in def.SurfaceBodies[1].Faces)
            {
                if (fa.SurfaceType == SurfaceTypeEnum.kPlaneSurface && u.eq(norm, ((Plane)fa.Geometry).Normal, 0.01))
                    fs.Add(fa);
            }
            string fn = doc.FullFileName, path = file.p(fn) + "DXF\\";
            file.createPath(path);
            string n = file.name(fn);

            var cm = I.app.CommandManager;
            var sp = f.PointOnFace;
            fs = fs.OrderBy(el => u.distToPlane(((Plane)el.Geometry), sp)).ToList();
            for (int i = 0; i < fs.Count; i++)
            {
                //ss.Clear();
                //cm.DoSelect(fs[i]);
                ////doc.SelectSet.Select(fs[i]);
                var name = $"{path}{n}_{i+1}.dxf";
                //cm.PostPrivateEvent(PrivateEventTypeEnum.kFileNameEvent, name);

                //var el = cm.ControlDefinitions["GeomToDXFCommand"];

                //el.Execute2(true);
                //cm.ClearPrivateEvents();

                //doc.Update2();
                FPOp.createDXF(fs[i], name);
            }

        }

        static public void splineInfo(BSplineCurve spl)
        {
            int order, nPoles, nKnots;
            bool rational, periodic, planar, closed;
            double[] vec = new double[] { }, poles = new double[] { }, knots = new double[] { }, weight = new double[] { };
            spl.GetBSplineInfo(out order, out nPoles, out nKnots, out rational, out periodic, out closed, out planar, ref vec);
            spl.GetBSplineData(ref poles, ref knots, ref weight);
            closed = true;
        }


        static public IEnumerable<DrawingCurveSegment> filterCurve(IEnumerable<DrawingCurveSegment> ie, Point2d origin, Point2d pt, bool x = true)
        {
            UnitVector2d vecOrigin = vecProjection(origin, pt, x);
            ie = ie.Where(s => isCollinear(vecProjection(origin, s.StartPoint, x), vecOrigin));
            return ie;
        }

        static public IEnumerable<DrawingCurveSegment> filterCurve(IEnumerable<DrawingCurveSegment> ie, DrawingView dv, Point2d pt, bool x = true)
        {
            Point2d origin = dv.Center;
            Vector2d vecOrigin = origin.VectorTo(pt);
            ie = ie.Where(s => directionIs(vecOrigin, pt.VectorTo(s.StartPoint), x));
            return ie;
        }

        static public object findAtVector(DrawingView view, Point2d pt, ref Point2d retPt, ref double dist, ref Vector2d v, double tol = 2, bool x = true)
        {
            IEnumerable<DrawingCurveSegment> ie = view.DrawingCurves.OfType<DrawingCurve>().Where(c => c.EdgeType != DrawingEdgeTypeEnum.kBendDownEdge
                && c.EdgeType != DrawingEdgeTypeEnum.kBendUpEdge)
                .SelectMany(d => d.Segments.OfType<DrawingCurveSegment>()).Where(s => s.GeometryType == Curve2dTypeEnum.kLineSegmentCurve2d);
            Point2d origin = view.Center;
            MeasureTools meas = I.app.MeasureTools;
            ie = filterCurve(ie, view, pt, x); //filterCurve(ie, origin, pt, x);
            if (ie.Count() == 0) return null;
            if (x)
                ie = ie.Where(s => isVertical(s));
            else ie = ie.Where(s => !isVertical(s));
            object ax = null; Point2d ret = null;
            if (ie.Count() == 0) return null;
            ax = ie.OrderByDescending(e => getDist(pt, ((DrawingCurveSegment)e).Parent)).First();
            dist = getDist(pt, ((DrawingCurveSegment)ax).Parent); //Vector2d v;
            if (x)
            {
                v = origin.VectorTo(pt); v.X = 0; v.Normalize(); v.ScaleBy(dist / 10);
                pt.TranslateBy(v); ret = pt;
            }
            else
            {
                v = origin.VectorTo(pt); v.Y = 0; v.Normalize(); v.ScaleBy(dist / 10);
                pt.TranslateBy(v); ret = pt;
            }
            //dist = getDist(pt, ((DrawingCurveSegment)ax).Parent, 1000, ref ret);
            //if ((ret.X - pt.X + ret.Y - pt.Y) < 0)
            //    dist = -dist;
            retPt = ret;
            return ax;
        }

        static public T findInCol<T>(System.Collections.IEnumerable col, Func<T, bool> f) where T : class
        {
            foreach (T item in col)
            {
                if (f(item)) return item;
            }
            return null;
        }

        static public double sumDist(double baseDist, double dist)
        {
            if (baseDist > 0)
            {
                return (dist > 0) ? dist + baseDist : dist - baseDist;
            }
            else
                return (dist > 0) ? baseDist - dist : baseDist + dist;
        }

        static public bool isVertical(DrawingCurveSegment ls)
        {
            Vector2d vec = ls.StartPoint.VectorTo(ls.EndPoint);
            return eq(vec.X, 0) ? true : false;
        }

        static public bool isVertical(SketchLine sl)
        {
            Vector2d vec = sl.StartSketchPoint.Geometry.VectorTo(sl.EndSketchPoint.Geometry);
            return eq(vec.X, 0) ? true : false;
        }

        static public bool isHorizontal(DrawingCurveSegment ls)
        {
            Vector2d vec = ls.StartPoint.VectorTo(ls.EndPoint);
            return eq(0, vec.Y) ? true : false;
        }

        static public bool isHorizontal(SketchLine sl)
        {
            Vector2d vec = sl.StartSketchPoint.Geometry.VectorTo(sl.EndSketchPoint.Geometry);
            return eq(0, vec.Y) ? true : false;
        }

        static public string convert(string s, double k, int tol)
        {
            return (Math.Round((convToDouble(s) / k), tol)).ToString();
        }

        static public bool eq(Face face, Point pt1)
        {
            Plane pl1 = face.Geometry as Plane;
            return (eq(pl1.DistanceTo(pt1), 0)) ? true : false;
        }

        static public bool eq(double tol, double d1, double d2)
        {
            return Math.Abs(d1 - d2) < tol ? true :
                false;
        }

        static public bool eq(double d1, double d2, int tol = 3)
        {
            return (Math.Round(d1, tol) == Math.Round(d2, tol));
        }

        static public bool eq(object d1, double d2)
        {
            return (Math.Round(double.Parse(d1.ToString()), 3) == Math.Round(d2, 3));
        }

        static public bool eq(Point2d pt1, Point2d pt2)
        {
            return (eq(pt1.X, pt2.X) && eq(pt1.Y, pt2.Y));
        }

        static public bool eq(Point2d pt1, Point2d pt2, int tol = 3)
        {
            return (eq(pt1.X, pt2.X, tol) && eq(pt1.Y, pt2.Y, tol));
        }

        static public bool eq(Point pt1, Point pt2, int tol = 3)
        {
            return (eq(pt1.X, pt2.X, tol) && eq(pt1.Y, pt2.Y, tol) && eq(pt1.Z, pt2.Z, tol));
        }

        static public bool eq(Point pt1, Point pt2, double tol)
        {
            double dist = pt1.DistanceTo(pt2);
            return dist < tol;
        }

        static public bool eq(Vertex pt1, Vertex pt2)
        {
            return (eq(pt1.Point.X, pt2.Point.X) && eq(pt1.Point.Y, pt2.Point.Y) && eq(pt1.Point.Z, pt2.Point.Z));
        }

        static public bool eq(UnitVector vec1, UnitVector vec2, bool ab = false, double tol = 0.01)
        {
            int s = 0;
            var d1 = getVData(vec1); var d2 = getVData(vec2);
            for (int i = 0; i < d1.Length; i++)
            {
                if (Math.Abs(d1[i] - d2[i]) < tol) s++;
            }
            return s > 1;
            //return (eq(vec1.X, vec2.X) && eq(vec1.Y, vec2.Y) && eq(vec1.Z, vec2.Z));
        }
        static public bool eq(UnitVector2d v1, UnitVector2d v2, double tol = 0.01)
        {
            var d1 = getCoord(v1); var d2 = getCoord(v2);
            int s = 0;
            for (int i = 0; i < d1.Length; i++)
            {
                if (!eq(d1[i], d2[i])) return false;
            }
            return true;
        }
        static public bool eq(UnitVector v1, UnitVector v2, double tol)
        {
            var d1 = getCoord(v1); var d2 = getCoord(v2);
            int s = 0;
            for (int i = 0; i < d1.Length; i++)
            {
                if (!eq(d1[i], d2[i])) return false;
            }
            return true;
        }

        static public double[] getVData(UnitVector v)
        {
            double[] d = new double[] { };
            v.GetUnitVectorData(ref d);
            return d;
        }

        static public bool eq(Vector vec1, Vector vec2)
        {
            return (eq(vec1.X, vec2.X) && eq(vec1.Y, vec2.Y) && eq(vec1.Z, vec2.Z));
        }

        static public bool eq(Vector2d vec1, Vector2d vec2)
        {
            if (vec1 == null || vec2 == null) return false;
            return (eq(vec1.X, vec2.X) && eq(vec1.Y, vec2.Y));
        }

        static public bool eq(SketchLine sl1, SketchLine sl2)
        {
            if (eq(sl1.Geometry.StartPoint, sl2.Geometry.StartPoint) && eq(sl1.Geometry.EndPoint, sl2.Geometry.EndPoint) ||
                eq(sl1.Geometry.EndPoint, sl2.Geometry.StartPoint) && eq(sl1.Geometry.StartPoint, sl2.Geometry.EndPoint)) return true;
            return false;
        }
        static public bool eq(SketchLine sl1, List<object> lst)
        {
            if (eq(sl1.Geometry3d.StartPoint, lst[0] as Point) && eq(sl1.Geometry3d.EndPoint, lst[1] as Point) ||
                eq(sl1.Geometry3d.EndPoint, lst[0] as Point) && eq(sl1.Geometry3d.StartPoint, lst[1] as Point)) return true;
            return false;
        }

        static public bool eq(SketchArc sl1, SketchArc sl2)
        {
            if (eq(sl1.StartSketchPoint.Geometry, sl2.StartSketchPoint.Geometry) && eq(sl1.EndSketchPoint.Geometry, sl2.EndSketchPoint.Geometry) ||
                eq(sl1.EndSketchPoint.Geometry, sl2.StartSketchPoint.Geometry) && eq(sl1.StartSketchPoint.Geometry, sl2.EndSketchPoint.Geometry)) return true;
            return false;
        }

        static public void abs(ref UnitVector vec, int tol = 4)
        {
            vec.X = Math.Abs(Math.Round(vec.X, tol)); vec.Y = Math.Abs(Math.Round(vec.Y, tol)); vec.Z = Math.Abs(Math.Round(vec.Z, tol));
        }

        static public void createDrawing(Document doc, string path)
        {
            path += "DWG\\";
            var drw = I.addDrw(path + file.name(doc.FullDocumentName) + ".idw");
            XMLDoc xmlDoc = new XMLDoc(I.p() + @"\sheet.xml", "head");
            string v = "Текущий";
            var sh = drw.ActiveSheet;
            var el = xmlDoc.El.Elements("view").FirstOrDefault(elem => elem.Attribute("name").Value.ToLower() == v.ToLower());
            var v1 = InvDocument<PartDocument>.addView(sh, doc, el, "view1", 0, new List<double> { 0 , 1});
            Point2d pt = I.CP2d(v1.Position, I.CV2d(v1.Width*2, 0));
            var v2 = InvDocument<PartDocument>.addDV((Document)drw, pt, sh, TypeViews.ProjectedView, 1, ViewOrientationTypeEnum.kArbitraryViewOrientation, I.nvm, v1, true, true);
            pt = I.CP2d(v1.Position, I.CV2d(0, v1.Height*2));
            var v3 = InvDocument<PartDocument>.addDV((Document)drw, pt, sh, TypeViews.ProjectedView, 1, ViewOrientationTypeEnum.kArbitraryViewOrientation, I.nvm, v1, true, true);
            pt = I.CP2d(v1.Position, I.CV2d(0, -v1.Height*2));
            v3 = InvDocument<PartDocument>.addDV((Document)drw, pt, sh, TypeViews.ProjectedView, 1, ViewOrientationTypeEnum.kArbitraryViewOrientation, I.nvm, v1, true, true);
            pt = I.CP2d(v2.Position, I.CV2d(v2.Width * 2, 0));
            v3 = InvDocument<PartDocument>.addDV((Document)drw, pt, sh, TypeViews.ProjectedView, 1, ViewOrientationTypeEnum.kArbitraryViewOrientation, I.nvm, v2, true, true);
            drw.Save(); 
        }

        static public void saveToDWG(DrawingDocument doc, string path)
        {
            ApplicationAddIn DWGAddin = null;

            foreach (ApplicationAddIn appAddin in I.app.ApplicationAddIns)
            {
                if (appAddin.ClassIdString == "{C24E3AC2-122E-11D5-8E91-0010B541CD80}")
                {
                    DWGAddin = appAddin;
                    break;
                }
            }
            TranslatorAddIn transl = (TranslatorAddIn)DWGAddin;
            var cont = I.objs.CreateTranslationContext();
            cont.Type = IOMechanismEnum.kFileBrowseIOMechanism;
            var opts = I.objs.CreateNameValueMap();
            var dm = I.objs.CreateDataMedium();
            string iniFile = path  + "DWG.ini";
            opts.Add("Export_Acad_IniFile", iniFile);
            string dwgPath = path + "DWG\\";
            file.createPath(dwgPath);
            dm.FileName = dwgPath + file.name(doc.FullDocumentName) + ".dwg";
            transl.SaveCopyAs(doc, cont, opts, dm);
        }

        static public void getTitleBlockToXML(DrawingDocument doc, string name)
        {
            InterfaceDll.MyXML xml = new InterfaceDll.MyXML(name, "root");
            var f2 = doc.TitleBlockDefinitions[name];
            var ds = f2.Sketch;
            foreach (Inventor.TextBox item in ds.TextBoxes)
            {
                XElement el = new XElement("tb");
                el.Add(new XAttribute("name", item.Text), new XAttribute("posX", item.RangeBox.MinPoint.X),
                    new XAttribute("posY", item.RangeBox.MinPoint.Y));
                xml.addElem(el);
            }
            xml.save();
        }

        static public void abs(ref Vector2d v, int tol = 4)
        {
            v.X = Math.Abs(Math.Round(v.X, tol)); v.Y = Math.Abs(Math.Round(v.Y));
        }
        static public T min<T>(System.Collections.IEnumerable ie, Func<T, double> f) where T : class
        {
            bool first = true;
            T m = null;
            foreach (T item in ie)
            {
                if (first)
                {
                    m = item;
                    first = false;
                    continue;
                }
                if (f(item) < f(m)) m = item;
            }
            return m;
        }

        static public T min<T>(T[] dc, Func<T, double> f)
        {
            T m = dc[0];
            foreach (var item in dc)
            {
                if (f(item) < f(m)) m = item;
            }
            return m;
        }

        static public T max<T>(T[] dc, Func<T, double> f)
        {
            T m = dc[0];
            foreach (var item in dc)
            {
                if (f(item) > f(m)) m = item;
            }
            return m;
        }

        static public double max(double v1, double v2)
        {
            return v1 > v2 ? v1 : v2;
        }
        static public double minAbs(double v1, double v2)
        {
            return Math.Abs(v1) < Math.Abs(v2) ? v1 : v2;
        }

        static public bool filter(double l, double min, double max)
        {
            return l <= max && l >= min;
        }

        static public GeometryIntent addGIDC(DrawingView dv, DrawingCurve dc, Point2d pt)
        {
            GeometryIntent gi;
            if (pt != null)
            {
                var min = u.min<Point2d>(new Point2d[] { dc.StartPoint, dc.EndPoint },
                    f => f.DistanceTo(pt));
                gi = dv.Parent.CreateGeometryIntent(dc, min);
            }
            else
                gi = dv.Parent.CreateGeometryIntent(dc);
            return gi;
        }

        static public double getLenght(Edge e, int tol = 3)
        {
            double minParam, maxParam, length;
            CurveEvaluator eval = e.Evaluator;
            eval.GetParamExtents(out minParam, out maxParam);
            eval.GetLengthAtParam(minParam, maxParam, out length);
            length = Math.Round(length, tol);
            return length;
        }

        static public double getLength(SketchArc e, int tol = 3)
        {
            double minParam, maxParam, length;
            var eval = e.Geometry3d.Evaluator;
            eval.GetParamExtents(out minParam, out maxParam);
            eval.GetLengthAtParam(minParam, maxParam, out length);
            length = Math.Round(length, tol);
            return length;
        }

        static public double getParam(Edge e, int tol = 3)
        {
            double l = getR(e);
            if (l == default(double)) return getLenght(e, tol);
            return l;
        }

        static public double getR(Edge ted)
        {
            double l = default(double);
            Circle c = ted.Geometry as Circle;
            if (!isNull(c)) return c.Radius;
            Arc2d a = ted.Geometry as Arc2d;
            if (!isNull(a)) return a.Radius;
            return l;
        }

        static public double[] getParamExtents(CurveEvaluator e)
        {
            double[] m = new double[2];
            e.GetParamExtents(out m[0], out m[1]);
            return m;
        }

        static public double[] getPointAtParam(CurveEvaluator e, double[] p)
        {
            double[] pts = new double[p.Length * 3];
            e.GetPointAtParam(ref p, ref pts);
            return pts;
        }

        static public List<Point> getPointAtParam(CurveEvaluator e, double div = 2)
        {
            double[] m = getParamExtents(e);
            double[] pts = getPointAtParam(e, new double[] { m[0], m[1], (m[0] + m[1]) / div });
            return getPointAtArray(pts);
        }

        static public double[] getParamExtents(Curve2dEvaluator e)
        {
            double[] m = new double[2];
            e.GetParamExtents(out m[0], out m[1]);
            return m;
        }

        static public double getParamAtPoint(Curve2dEvaluator e, Point2d pt)
        {
            double [] pts = new double []{};
            pt.GetPointData(ref pts);
            double [] guess = new double []{};
            double [] maxDeviation = new double []{};
            double [] p = new double []{};
            SolutionNatureEnum[] sol = new SolutionNatureEnum[] { };
            e.GetParamAtPoint(ref pts, ref guess, ref maxDeviation, ref p, ref sol);
            return p[0];
        }

        static public double getParamAtPoint(CurveEvaluator e, Point pt)
        {
            double[] pts = new double[] { };
            pt.GetPointData(ref pts);
            double[] guess = new double[] { };
            double[] maxDeviation = new double[] { };
            double[] p = new double[] { };
            SolutionNatureEnum[] sol = new SolutionNatureEnum[] { };
            e.GetParamAtPoint(ref pts, ref guess, ref maxDeviation, ref p, ref sol);
            return p[0];
        }

        static public double[] getPointAtParam(Curve2dEvaluator e, double[] p)
        {
            double[] pts = new double[p.Length * 2];
            e.GetPointAtParam(ref p, ref pts);
            return pts;
        }

        static public Point2d getPointAtParam(Curve2dEvaluator e, double p)
        {
            double [] pts = getPointAtParam(e, new double[] {p});
            return u.createPoint2d(pts[0], pts[1]);
        }

        static public List<Point> getPointAtArray(double[] arr)
        {
            List<Point> pts = new List<Point>();
            for (int i = 0; i < arr.Length; i += 3)
            {
                pts.Add(I.CP(arr[i], arr[i + 1], arr[i + 2]));
            }
            return pts;
        }

        static public Vector CrossProduct(Point sp, Point ep, Point cen)
        {
            Vector v1 = cen.VectorTo(sp), v2 = cen.VectorTo(ep);
            return v1.CrossProduct(v2);
        }

        static public List<Point> getPoleAtIndex(BSplineCurve curve, int ind, int numP)
        {
            List<Point> pts = new List<Point>();
            int step = numP / ind;
            for (int i = 0; i < numP; i += step)
            {
                pts.Add(curve.get_PoleAtIndex(i + 1));
            }
            return pts;
        }

        static public double getLenght(Box rb, int tol = 3)
        {
            return Math.Round(rb.MaxPoint.DistanceTo(rb.MinPoint), tol);
        }

        static public double getSquareLenght(Box rb, bool max = false, int tol = 3)
        {
            Vector v = rb.MaxPoint.VectorTo(rb.MinPoint);
            if (max)
            {
                double[] data = new double[3]; v.GetVectorData(ref data);
                return Math.Round(data.Max(), tol);
            }
            else
                return Math.Round(scalar(v, v), tol);
        }

        static public bool checkSquareLenght(Box rb, double val)
        {
            double v = val * val;
            return getSquareLenght(rb, true) < val;
        }

        static public double getLenght(DrawingCurve e, int tol = 3)
        {
            double minParam, maxParam, length;
            Curve2dEvaluator eval = e.Evaluator2D;
            eval.GetParamExtents(out minParam, out maxParam);
            eval.GetLengthAtParam(minParam, maxParam, out length);
            length = Math.Round(length, tol);
            return length;
        }
        static public double getLengthModel(DrawingCurve dc, int tol = 3)
        {
            var l = dc.ModelGeometry as Edge;
            if (l == null) return 0;
            return getLenght(l);
        }
        static public double[] getParam(Curve2dEvaluator ev, Point2d pt)
        {
            double[] pts = { pt.X, pt.Y };
            double[] guess = { };
            double[] maxDev = { }, param = { };
            SolutionNatureEnum[] e = { };
            ev.GetParamAtPoint(ref pts, ref guess, ref maxDev, ref param, e);
            return param;
        }

        static public double[] getTangent(Curve2dEvaluator ev, double inc = 0.1, Point2d pt = null)
        {
            double min, max;
            ev.GetParamExtents(out min, out max);
            double[] p = { min + (max - min) * inc };
            if (pt != null) p = getParam(ev, pt);
            double[] t = { };
            ev.GetTangent(ref p, ref t);
            return t;
        }

        static public double getMax(Box b)
        {
            double[] min = {};
            double[] max = {};
            b.GetBoxData(ref min, ref max);
            return getMax(min, max);
        }

        static public double getMax(double[] min, double [] max)
        {
            List<double> v = new List<double>();
            for (int i = 0; i < min.Length; i++)
			{
			    v.Add(max[i]-min[i]);
			}
            return v.Max();
        }

        public static Tuple<Point2d, Vector2d, Point2d, Vector2d> getTangentMinMax(DrawingCurve dc)
        {
            //double min, max;
            Curve2dEvaluator ev = dc.Evaluator2D;
            var p1 = getParam(ev, dc.StartPoint);
            var p2 = getParam(ev, dc.EndPoint);
            double[] p = new double[] { p1[0], p2[0] };
            double[] t = {};
            double[] pt = {};
            ev.GetTangent(ref p, ref t);
            ev.GetPointAtParam(ref p, ref pt);
            Point2d m = dc.MidPoint;
            Point2d pt1 = I.CP2d(pt[0], pt[1]);
            Point2d pt2 = I.CP2d(pt[2], pt[3]);
            Vector2d v1 = I.CV2d(t[0], t[1]);
            Vector2d v2 = I.CV2d(t[2], t[3]);
            //if (direct(m, pt1, v1)) 
                v1.ScaleBy(-1);
            //if (!direct(m, pt2, v2)) 
                //v2.ScaleBy(-1);
            Tuple<Point2d, Vector2d, Point2d, Vector2d> tup = new Tuple<Point2d, Vector2d, Point2d, Vector2d>(pt1, v1, pt2, v2);
            return tup;
        }

        public static bool direct(Point2d mid, Point2d pt, Vector2d v)
        {
            Vector2d tmpv = I.CV2d(v); tmpv.Normalize();
            double d1 = pt.DistanceTo(mid);
            Point2d tmppt = I.CP2d(pt);
            tmppt.TranslateBy(tmpv);
            double d2 = tmppt.DistanceTo(mid);
            return d1 < d2;
        }

        static public Vector2d getTangentVec(Curve2dEvaluator ev, double inc = 0.1, Point2d pt = null)
        {
            double[] v = getTangent(ev, inc);
            if (pt != null) v = getTangent(ev, inc, pt);
            return I.tg.CreateVector2d(v[0], v[1]);
        }

        static public List<T> getCurves<T>(DrawingView dv, Func<DrawingCurve, bool> f) where T : class
        {
            List<T> cur = new List<T>();
            foreach (DrawingCurve dc in dv.DrawingCurves)
            {
                if (f(dc)) cur.Add(dc as T);
            }
            return cur;
        }

        static public double getPar(DrawingView dv, Func<SheetMetalComponentDefinition, object> f)
        {
            SheetMetalComponentDefinition smcd = docCompDef<SheetMetalComponentDefinition>(dv.ReferencedDocumentDescriptor.ReferencedDocument);
            return (smcd != null) ? convToDouble(f(smcd).ToString()) : 0;
        }

        static public T getStyle<T>(DrawingView dv, string name) where T : class
        {
            return get<T>((dv.Parent.Parent as DrawingDocument).StylesManager.DimensionStyles, e => Reflect.getProp(e, "Name") == name) as T;
        }

        static public T get<T>(System.Collections.IEnumerable ie, Func<T, bool> f) where T : class
        {
            foreach (var el in ie)
            {
                T item = el as T;
                if (item == null) continue;
                if (f(item)) return item as T;
            }
            return null;
        }

        static public object get(System.Collections.IEnumerable ie, int i)
        {
            foreach (var item in ie)
            {
                i--;
                if (i < 0) break;
                if (i == 0) return item;
            }
            return null;
        }

        static public IEnumerable<T> sort<T>(IEnumerable<T> ie, string s)
        {
            bool rev = false; int num = 0;
            if (s.StartsWith("-"))
            {
                rev = true; s = s.TrimStart(new char[] { '-' });
            }
            switch (s.ToLower())
            {
                case "x":
                    num = 0;
                    break;
                case "y":
                    num = 1;
                    break;
                case "z":
                    num = 2;
                    break;
                default:
                    break;
            }
            if (typeof(T) == typeof(SketchPoint))
                return sort(ie as IEnumerable<SketchPoint>, rev, num) as IEnumerable<T>;
            else if (typeof(T) == typeof(Face))
                return sort(ie as IEnumerable<Face>, rev, num) as IEnumerable<T>;
            else if (typeof(T) == typeof(Edge))
                return sort(ie as IEnumerable<Edge>, rev, num) as IEnumerable<T>;
            return null;
        }

        static public IEnumerable<SketchPoint> sort(IEnumerable<SketchPoint> ie, bool r, int n)
        {
            var s = ie.OrderBy(f => getCoord(f, n));
            if (r) s.Reverse();
            return s;
        }

        static public IEnumerable<Face> sort(IEnumerable<Face> ie, bool r, int n)
        {
            return r ? ie.OrderBy<Face, double>(f => getCenterFace(f, n)) : ie.OrderByDescending<Face, double>(f => getCenterFace(f, n));
        }

        static public IEnumerable<Edge> sort(IEnumerable<Edge> ie, bool r, int n)
        {
            return r ? ie.OrderBy<Edge, double>(f => getCenterEdge(f, n)) : ie.OrderByDescending<Edge, double>(f => getCenterEdge(f, n));
        }

        static public double getCenterFace(Face f, int n)
        {
            var p = f;
            var pt = midPt(p.Evaluator.RangeBox.MinPoint, p.Evaluator.RangeBox.MaxPoint, 0.5);
            return getCoord(pt, n);
        }

        static public Face getFace(object f)
        {
            FaceProxy fp = f as FaceProxy;
            Face face = f as Face;
            if (fp != null)
            {
                face = fp.NativeObject as Face;
            }
            return face;
        }

        static public SketchPoint getSketchPoint(object f)
        {
            SketchPoint sp = f as SketchPoint;
            if (sp == null) return null;
            SketchPointProxy spp = f as SketchPointProxy;
            if (spp != null)
            {
                sp = spp.NativeObject as SketchPoint;
            }
            return sp;
        }

        static public void addSketchPoints(object sketch, PlanarSketch par)
        {
            PlanarSketchProxy psp = sketch as PlanarSketchProxy;
            PlanarSketch ps = sketch as PlanarSketch;
            if (ps == null && psp == null) return;
            if (psp != null)
            {
                ps = psp.NativeObject as PlanarSketch;
            }
            foreach (SketchPoint item in u.gets<SketchPoint>(ps.SketchPoints, f => f.HoleCenter))     
            {
                par.AddByProjectingEntity(item);
            }
        }
        static public T getGeomEdge<T>(Edge e) where T: class
        {
            return e.Geometry as T;
        }
        static public Point getCenPoint(Edge e)
        {
            var l = getGeomEdge<LineSegment>(e);
            if (l != null) return l.MidPoint;
            var c = getGeomEdge<Circle>(e);
            if (c != null) return c.Center;
            var a = getGeomEdge<Arc3d>(e);
            if (a != null) return a.Center;
            return null;
        }
        static public Point2d getCenPoint(DrawingCurve dc)
        {
            return dc.MidPoint ?? dc.CenterPoint;
        }
        static public Vector getDir(Edge e)
        {
            var l = getGeomEdge<LineSegment>(e);
            return l?.Direction.AsVector();
        }
        static public double getCenterEdge(Edge e, int n)
        {
            var p = e.Evaluator.RangeBox;
            var pt = midPt(p.MinPoint, p.MaxPoint);
            return getCoord(pt, n);
        }

        static public Point getCenter(Edge e)
        {
            dynamic tmp = e.Geometry;
            Point pt = tmp.Center;
            return pt;
        }

        static public double getCoord<T>(T sp, int n) where T : SketchPoint
        {
            return getCoord(sp.Geometry3d, n);
        }

        static public double getCoord(Point pt, int n)
        {
            double[] d = new double[] { };
            pt.GetPointData(ref d);
            return d[n];
        }

        static public object getLast(System.Collections.IEnumerable ie)
        {
            var e = ie.GetEnumerator();
            object cur = null;
            while (e.MoveNext()) cur = e.Current;
            return cur;
        }

        static public int getCount(System.Collections.IEnumerable ie)
        {
            var e = ie.GetEnumerator();
            int cur = 0;
            while (e.MoveNext()) cur++;
            return cur;
        }

        static public object get<T>(System.Collections.IEnumerable ie, string val, string name = "Name")
        {
            foreach (T item in ie)
            {
                if (Reflect.getProp<T>(item, name).ToString() == val) return item;
            }
            return null;
        }

        static public IEnumerable<T> gets<T>(System.Collections.IEnumerable ie, Func<T, bool> filter) where T: class
        {
            foreach (var item in ie)
            {
                T el = item as T;
                if (el == null) continue;
                if (filter(el)) yield return el;
            }
        }

        static public IEnumerable<T> gets<T>(System.Collections.IEnumerable ie, string filter, char sep) where T : class
        {
            string[] spl = getSpl(filter, sep);
            if (spl == null) yield break;
            int i = 0;
            List<string> num = spl.ToList();
            int max = getCount(ie);
            List<int> n = num.Select(f => lMinus(getIntElem(f, 0), max)).ToList();
            foreach (var item in ie)
            {
                i++;
                T el = item as T;
                if (el == null) continue;
                if (n.Contains(i)) yield return el;
            }
        }

        static public void clientLine(PartComponentDefinition pcd, Edge e, string t)
        {
            var cl = pcd.ClientGraphicsCollection;
            var ds = ((Document)pcd.Document).GraphicsDataSetsCollection.Add(t);
            var col = ds.CreateColorSet(1);
            ClientGraphics gds = null;
            try
            {
                gds = cl[t];
            }
            catch (System.Exception ex)
            {
                gds = cl.Add(t);
            }
            GraphicsNode n = gds.AddNode(1);
            var tg = n.AddLineGraphics();
            var cs = ds.CreateCoordinateSet(1);
            cs.Add(1, e.StartVertex.Point); cs.Add(2, e.StopVertex.Point);
            tg.CoordinateSet = cs;
            col.Add(1, 255, 0, 0);
            tg.ColorSet = col;
        }
        static public Point2d getPt(double[] v)
        {
            Point2d pt = I.CP2d();
            pt.PutPointData(ref v);
            return pt;
        }
        static public void clientTxt(IEnumerable<Face> fs, PartComponentDefinition pcd)
        {
            int i = 0;
            foreach (var item in fs)
            {
                i++;
                var rb = item.Evaluator.RangeBox;
                var pt = u.midPt(rb.MinPoint, rb.MaxPoint, 0.5);
                clientTxt(pcd, pt, i.ToString());
            }
        }

        static public void clientTxt(IEnumerable<FaceProxy> fs, AssemblyComponentDefinition pcd)
        {
            int i = 0;
            foreach (var item in fs)
            {
                i++;
                var rb = item.Evaluator.RangeBox;
                var pt = u.midPt(rb.MinPoint, rb.MaxPoint, 0.5);
                clientTxt(pcd, pt, i.ToString());
            }
        }

        static public void clientTxt(IEnumerable<EdgeProxy> fs, AssemblyComponentDefinition pcd)
        {
            int i = 0;
            foreach (var item in fs)
            {
                i++;
                var rb = item.Evaluator.RangeBox;
                var pt = u.midPt(rb.MinPoint, rb.MaxPoint, 0.5);
                clientTxt(pcd, pt, i.ToString());
            }
        }

        static public void clientTxt(PartComponentDefinition pcd, Point pt, string t)
        {
            var cl = pcd.ClientGraphicsCollection;
            ClientGraphics gds = null;
            try
            {
                gds = cl[t];
            }
            catch (System.Exception ex)
            {
                gds = cl.Add(t);
            }
            GraphicsNode n = gds.AddNode(1);
            var tg = n.AddTextGraphics();
            tg.Anchor = pt;
            tg.Text = t;
        }

        static public void clientTxt(AssemblyComponentDefinition pcd, Point pt, string t)
        {
            var cl = pcd.ClientGraphicsCollection;
            ClientGraphics gds = null;
            try
            {
                gds = cl[t];
            }
            catch (System.Exception ex)
            {
                gds = cl.Add(t);
            }
            GraphicsNode n = gds.AddNode(1);
            var tg = n.AddTextGraphics();
            tg.Anchor = pt;
            tg.Text = t;
        }
        static public void clientArrow(PartComponentDefinition pcd, string name, Point p1, Point p2)
        {
            var cl = pcd.ClientGraphicsCollection;
            ClientGraphics gds = null;
            try
            {
                gds = cl[name];
                gds.Delete();
                I.app.ActiveView.Update();
            }
            catch (Exception)
            {
                gds = pcd.ClientGraphicsCollection.Add(name);
                var node = gds.AddNode(1);
                var tbr = I.app.TransientBRep;
                var v = p1.VectorTo(p2);

                SurfaceBody cyl = tbr.CreateSolidCylinderCone(p2, p1, v.Length / 10, v.Length / 10, 0, null);
                var sgs = node.AddSurfaceGraphics(cyl);
                I.app.ActiveView.Update();
            }
        }

        static public Point[] getRect(Point pt, double w, double h)
        {
            return new Point[]
            {
                I.CP(pt.X, pt.Y, 0),
                I.CP(pt, I.CV(w,0,0),1),
                I.CP(pt, I.CV(w,h,0),1),
                I.CP(pt, I.CV(0,h,0),1),
                //I.CP(pt, I.CV(0,h,0),1),
                //I.CP(pt, I.CV(-w,0,0),1),
                I.CP(pt.X, pt.Y, 0),
            };
        }

        static public Point[] getPoints(Point2d [] pts)
        {
            Point[] p = new Point[pts.Length];
            for (int i = 0; i < pts.Length; i++)
            {
                p[i] = I.CP(pts[i].X, pts[i].Y, 0);
            }
            return p;
        }

        static public Point[] getRect(Point2d pt, Vector2d v, Vector2d n)
        {
            Point2d p2 = pt.Copy();
            p2.TranslateBy(v);
            Point2d p3 = p2.Copy();
            p3.TranslateBy(n);
            Point2d p4 = pt.Copy();
            p4.TranslateBy(n);
            return new Point[]
            {
                I.CP(pt.X, pt.Y, 0), I.CP(p2.X, p2.Y, 0),I.CP(p3.X, p3.Y, 0),I.CP(p4.X, p4.Y, 0),
                I.CP(pt.X, pt.Y, 0)
            };
        }

        static public void clientLine(PartComponentDefinition def, string name, string colName, Point[] pts)
        {
            var cl = def.ClientGraphicsCollection;
            PartDocument doc = def.Document as PartDocument;
            ClientGraphics gds = null;
            GraphicsDataSets col = null;
            try
            {
                col = doc.GraphicsDataSetsCollection[colName];
                gds = cl[name];
                gds.Delete();
                col.Delete();
                I.app.ActiveView.Update();
            }
            catch (Exception)
            {
                col = doc.GraphicsDataSetsCollection.Add(colName);
                gds = def.ClientGraphicsCollection.Add(name);
                var node = gds.AddNode(1);
                var tbr = I.app.TransientBRep;
                var coord = col.CreateCoordinateSet(1);
                for (int i = 0; i < pts.Length; i++)
                {
                    coord.Add(i + 1, pts[i]);
                }
                var color = col.CreateColorSet(2);
                color.Add(1, 255, 0, 0);
                var lg = node.AddLineGraphics();
                lg.CoordinateSet = coord;
                lg.ColorSet = color;
                viewUpdate(true);
            }
        }

        static public T getNode<T>(object o, int num) where T : class
        {
            T r = null;
            dynamic d = o;
            try
            {
                r = d.ItemById[num] as T;
            }
            catch (Exception)
            {
                return null;
            }
            return r;
        }

        static public GraphicsColorSet addColor(GraphicsDataSets col, int num, int r, int g, int b)
        {
            GraphicsColorSet c = getNode<GraphicsColorSet>(col, num) ?? col.CreateColorSet(num);
            c.Add(num, (byte)r, (byte)g, (byte)b);
            return c;
        }

        static public LineGraphics clientLine(ClientGraphics g, GraphicsDataSets col, string name, 
            string colName, Point[] pts, int num = 1, bool upd = true, bool cl = false)
        {
            //var cl = doc.ActiveSheet.ClientGraphicsCollection;
            //ClientGraphics g = null;
            //GraphicsDataSets col = null;
            //clearClientLine(doc, name, colName);
            //getClient(doc, name, colName, ref col, ref g);
            GraphicsNode node;
            GraphicsCoordinateSet coord;
            node = getNode<GraphicsNode>(g, num) ?? g.AddNode(num);
            //node = g.Count < num ? g.AddNode(num) : g.ItemById[num];
            //var tbr = I.app.TransientBRep;
            coord = getNode<GraphicsCoordinateSet>(col, num) ?? col.CreateCoordinateSet(num);
            //try
            //{
            //    coord = col.ItemById[num] as GraphicsCoordinateSet;
            //}
            //catch (Exception)
            //{
            //    coord = col.CreateCoordinateSet(num);
            //}
            //coord = col.ItemById[num] as GraphicsCoordinateSet;
            //if (coord == null) coord = col.CreateCoordinateSet(num);
            //coord = /*col.Count < num ?*/ col.CreateCoordinateSet(num); 
            //col.ItemById[num] as GraphicsCoordinateSet;
            if (cl)
            {
                coord.Delete();
                coord = col.CreateCoordinateSet(num); 
            }
            double[] c = new double[pts.Length * 3];
            int ind = pts.Length*2;
            for (int i = 0; i < pts.Length; i++)
            {
                //ind++;
                coord.Add(ind, pts[i]);
                if (i > 0)
                {
                    ind++;
                    coord.Add(ind, pts[i]);
                }
            }
            //var color = col.CreateColorSet(100+num);
            //color.Add(1, 255, 0, 0);
            LineGraphics lg;
#if INV14
            lg = node.Count < num ? node.AddLineGraphics() : node[num] as LineGraphics;
#else
        lg = node.Count < num ? node.AddLineGraphics(): node.ItemById[num] as LineGraphics;
#endif

            lg.CoordinateSet = coord;
            //lg.ColorSet = color;
            //Point p;
            //DisplayTransformBehaviorEnum dt;
            //double px;
            //lg.SetTransformBehavior(I.CP(), DisplayTransformBehaviorEnum.kFrontFacingAndPixelScaling);
            //lg.GetTransformBehavior(out p, out dt, out px);


            //viewUpdate(upd);
            return lg;
        }


        static public LineGraphics addClientLine(LineGraphics lg, Point pt, bool upd)
        {
            var coord = lg.CoordinateSet;
            coord.Add(coord.Count + 5, pt);
            lg.CoordinateSet = coord;
            //viewUpdate(upd);
            return lg;
        }

        static public void viewUpdate(bool upd)
        {
            if (upd) I.app.ActiveView.Update();
        }

        static public void viewUpdate(bool upd, View v, InteractionGraphics ig)
        {
            if (upd) ig.UpdateOverlayGraphics(v);
        }

        static public void getClient(DrawingDocument doc, string name, string colName,
            ref GraphicsDataSets col, ref ClientGraphics g)
        {
            var cl = doc.ActiveSheet.ClientGraphicsCollection;
            if (doc.GraphicsDataSetsCollection.Count > 0)
                col = u.get<GraphicsDataSets>(doc.GraphicsDataSetsCollection, f => f.ClientId == colName);
            if (cl.Count > 0)
                g = u.get<ClientGraphics>(cl, f => f.ClientId == name);
            if (col == null) col = doc.GraphicsDataSetsCollection.Add(colName);
            if (g == null) g = cl.Add(name);
        }

        static public void clearClientLine(DrawingDocument doc, string name, string colName)
        {
            var cl = doc.ActiveSheet.ClientGraphicsCollection;
            ClientGraphics g = null;
            GraphicsDataSets col = null;
            try
            {
                col = doc.GraphicsDataSetsCollection[colName];
                g = cl[name];
                g.Delete();
                col.Delete();
                I.app.ActiveView.Update();
            }
            catch (Exception)
            {
            }
        }

        static public IEnumerable<int> lMinus(IEnumerable<int> vals, int max)
        {
            return vals.Select(e => lMinus(e, max));
        }

        static public int lMinus(int v, int max)
        {
            return v < 0 ? max + 1 + v : v;
        }

        static public int getNum(string str, int max)
        {
            if (isNull(str)) return default(int);
            return lMinus(getIntElem(str, 0), max);
        }

        static public void getDerSolid(PartComponentDefinition pcd)
        {
            if (pcd.ReferenceComponents.DerivedPartComponents.Count == 0) return;
            DerivedPartComponent d = pcd.ReferenceComponents.DerivedPartComponents[1];
            if (isNull(d)) return;
            DerivedPartDefinition def = d.Definition; bool ch = false;
            foreach (DerivedPartEntity item in def.Solids)
            {
                if (!item.IncludeEntity) { item.IncludeEntity = true; ch = true; }
            }
            if (ch) d.Definition = def;
        }

        static public string[] getSpl(string val, char sep)
        {
            //if (val.IndexOf(sep) == -1) return null;
            if (isNull(val)) return null;
            var spl = val.Split(sep);
            if (spl == null) return null;
            return spl;
        }

        static public IEnumerable<T> add<T>(T ob)
        {
            yield return ob;
        }

        static public IEnumerable<T> add<T>(IEnumerable<T> ie, T ob)
        {
            return ie != null ? ie.Concat<T>(add<T>(ob)) : add<T>(ob);
        }

        static public IEnumerable<T> add<T>(IEnumerable<T> ie1, IEnumerable<T> ie2)
        {
            return (ie1 != null && ie2 != null) ? ie1.Concat<T>(ie2): ie2;
        }

        static public void action<T>(IEnumerable<T> ie, Action<T> a, Func<T, bool> f = null)
        {
            foreach (T item in ie)
            {
                if (f == null) a(item);
                else if (f(item)) a(item);
            }
        }

        static public void action<T>(System.Collections.IEnumerable ie, Action<T> a, Func<T, bool> f = null)
        {
            try
            {
                foreach (T item in ie)
                {
                    if (f == null) a(item);
                    else if (f(item)) a(item);
                }
            }
            catch (System.Exception ex)
            {

            }
        }

        static public void actionFor<T>(T[] arr, Action<T> a, Func<T, bool> f = null)
        {
            for (int i = 0; i < arr.Length; i++)
            {
                if (f == null) a(arr[i]);
                else if (f(arr[i])) a(arr[i]);
            }
        }

        static public void actionFor<T>(T[] arr, Action<T, T> a, Func<T, bool> f = null)
        {
            for (int i = 1; i < arr.Length; i++)
            {
                if (f == null) a(arr[i - 1], arr[i]);
            }
        }

        static public void actionFor<T>(List<T> arr, Action<T> a, Func<T, bool> f = null)
        {
            for (int i = 0; i < arr.Count; i++)
            {
                if (f == null) a(arr[i]);
                else if (f(arr[i])) a(arr[i]);
            }
        }

        static public List<T> actionFor<T>(List<T> arr, Action<T, T> a, Func<T, bool> r = null, double center = 1)
        {
            List<T> tmp = new List<T>();
            if (r != null)
                for (int i = 0; i < arr.Count; i++)
                {
                    if (r(arr[i])) tmp.Add(arr[i]);
                }
            else tmp = arr;
            if (tmp.Count < 2) return null;
            for (int i = 1; i < tmp.Count * center; i++)
            {
                a(tmp[i - 1], tmp[i]);
            }
            return tmp;
        }

        static public List<Arc3d> slotArcs(EdgeLoop l)
        {
            var cir = u.gets<Edge>(l.Edges, el => el.CurveType == CurveTypeEnum.kCircleCurve).Select(e => e.Geometry as Arc3d).ToList();
            if (cir.Count != 2) return null;
            if (!u.eq(cir[0].Radius, cir[1].Radius)) return null;
            return cir;
        }
        static public double [] slotData(EdgeLoop l)
        {
            var cir = u.slotArcs(l);
            if (cir == null) return null;
            var r = new double[] { cir[0].Radius * 2, cir[0].Radius * 2 + cir[0].Center.DistanceTo(cir[1].Center) };
            r[0] = u.round(r[0]); r[1] = u.round(r[1]);
            return r;
        }
        static public double? slotGroup(EdgeLoop l)
        {
            var r = u.slotData(l);
            if (r == null) return null;
            return r[0] + r[1];
        }
        static public Point slotCenter(EdgeLoop l)
        {
            var cir = u.slotArcs(l);
            if (cir == null) return null;
            var pt = u.midPt(cir[0].Center, cir[1].Center);
            return pt;
        }
        static public Vector slotDir(EdgeLoop l)
        {
            var cir = u.slotArcs(l);
            if (cir == null) return null;
            return cir[0].Center.VectorTo(cir[1].Center);
        }

        static public string getName(string name, char sep, int ind)
        {
            var spl = name.Split(sep);
            return spl[ind];
        }

        static public string shortName(string name)
        {
            var spl = name.Split(' ');
            var r = "";
            var word = "";
            List<char> fon = new List<char>() { 'а', 'е', 'у', 'о', 'и', 'ё', 'ы', 'э', 'я', 'ю' };
            try
            {
                var l = 4;
                foreach (var item in spl)
                {
                    if (item.Length < l) l = item.Length;
                    for (int i = 0; i < l; i++)
                    {
                        if ((i == 2 || i == 3) && fon.Contains(item[i])) break;
                        word += item[i];
                    }
                    word = word[0].ToString().ToUpper() + word.Substring(1);
                    r += word;
                    word = "";
                }
            }
            catch (Exception)
            {
                return "unknow";
            }

            return r;
        }

        static public string shortName(Document doc)
        {
            var name =  getPropValue(doc, "Description").ToLower();
            if (name == "") return "";
            return shortName(name);
        }

        static public double radToDeg(double rad, int tol = 3)
        {
            return round(rad * 180 / Math.PI, tol);
        }

        static public double degToRad(double deg, int tol = 3)
        {
            return round(deg * Math.PI / 180, tol);
        }

        static public double angFromPoint(Point pt, Point cen, double m)
        {
            Point p = I.CP(pt.X * m, pt.Y * m, pt.Z * m), c = I.CP(cen.X * m, cen.Y * m, cen.Z * m);
            Vector v = c.VectorTo(p);
            return angFromVector(v);
        }
        static public double angFromPoint(Point2d pt, Point2d cen, double m)
        {
            Point p = I.CP(pt.X * m, pt.Y * m, 0), c = I.CP(cen.X * m, cen.Y * m, 0);
            Vector v = c.VectorTo(p);
            return angFromVector(v);
        }
        static public double angFromVector(Vector v)
        {
            double a = v.AngleTo(I.CV(1, 0, 0));
            double x = round(v.X, 10), y = round(v.Y, 10);
            //if (x < 0 && y > 0) a = Math.PI - a;
            //if (y < 0 && x > 0) a = Math.PI * 2 - a;
            if (y < 0) a = Math.PI * 2 - a;
            //if (x == 0 && y > 0) a = Math.PI/2;
            //if (x == 0 && y < 0) a = 3*Math.PI / 2;
            return radToDeg(a, 10);
        }

        static public double round(double val, int tol = 3)
        {
            return Math.Round(val, tol);
        }

        static public UnitVector round(UnitVector v)
        {
            return I.CUV(round(v.X), round(v.Y), round(v.Z));
        }

        static public double scalar(DrawingCurve sl, DrawingCurve prev)
        {
            return scalar(sl.StartPoint.VectorTo(sl.EndPoint).AsUnitVector(), prev.StartPoint.VectorTo(prev.EndPoint).AsUnitVector());
        }

        static public double scalar(SketchLine sl, SketchLine prev)
        {
            return scalar(sl.Geometry.Direction, prev.Geometry.Direction);
        }

        static public double scalar(UnitVector2d v1, UnitVector2d v2)
        {
            return v1.X * v2.X + v1.Y * v2.Y;
        }

        static public double scalar(Vector v1, Vector v2)
        {
            return v1.X * v2.X + v1.Y * v2.Y + v1.Z * v2.Z;
        }

        static public double scalar(Point pt1, UnitVector dir)
        {
            return pt1.X * Math.Abs(dir.X) + pt1.Y * Math.Abs(dir.Y) + pt1.Z * Math.Abs(dir.Z);
        }

        static public double scalar(Point2d pt, UnitVector2d dir)
        {
            return pt.X * Math.Abs(dir.X) + pt.Y * Math.Abs(dir.Y);
        }
        static public Point2d ptIntent(GeometryIntent gi)
        {
            var cl = gi.Geometry as Centerline;
            if (cl != null) return cl.StartPoint;
            var cm = gi.Geometry as Centermark;
            if (cm != null) return cm.Position;
            return null;
        }
        static public Point2d midPt(Box2d b)
        {
            return midPt(b.MinPoint, b.MaxPoint);
        }
        static public Point midPt(Box b)
        {
            return midPt(b.MinPoint, b.MaxPoint);
        }
        static public Point2d modelToView(Point pt, DrawingView dv)
        {
            var c = dv.Camera;
            return c.ModelToViewSpace(pt);
        }
        static public Point viewToModel(Point2d pt, DrawingView dv)
        {
            var c = dv.Camera;
            return c.ViewToModelSpace(pt);
        }
        static public Vector viewToModel(Vector2d v, DrawingView dv)
        {
            var p1 = viewToModel(I.CP2d(), dv); var p2 = viewToModel(I.CP2d(I.CP2d(), v), dv);
            return p1.VectorTo(p2);
        }
        static public void addText(DrawingView dv, string txt, Point2d pt)
        {
            dv.Parent.DrawingNotes.GeneralNotes.AddFitted(pt, txt);
        }
        static public UnitVector2d vector3dTo2d(object v, UnitVector n)
        {
            Vector bv = v as Vector;
            UnitVector uv = v as UnitVector;
            double[] coord = getCoord(v);
            double[] ncoord = getCoord(n);  
            double[] ucoord = new double[2];
            int j = 0;
            for (int i = 0; i < coord.Length; i++)
            {
                if (eq(Math.Abs(ncoord[i]), 1)) continue;
                ucoord[j] = coord[i];
                j++;
            }
            var vec = I.CV2d(ucoord[0], ucoord[1]);
            return vec.AsUnitVector();
        }
        static public T vector2dTo3d<T>(object vec) where T : class
        {
            double[] coord = getCoord(vec);
            return (vec is UnitVector2d) ? I.CUV(coord[0], coord[1], 0) as T : I.CV(coord[0], coord[1], 0) as T;
        }
        static public double distToPlane(Plane pl, Point pt, bool dir = false)
        {
            var n = pl.Normal.AsVector(); var v = pl.RootPoint.VectorTo(pt);

            return dir ? dotProduct(n, v) : Math.Abs(dotProduct(n, v));
        }
        static public double[] getCoord(object vec)
        {
            UnitVector2d uv2d = vec as UnitVector2d;
            UnitVector uv = vec as UnitVector;
            Vector v = vec as Vector;
            Vector2d v2d = vec as Vector2d;
            double[] coord = { };
            if (uv2d != null) uv2d.GetUnitVectorData(ref coord);
            if (uv != null) uv.GetUnitVectorData(ref coord);
            if (v != null) v.GetVectorData(ref coord);
            if (v2d != null) v2d.GetVectorData(ref coord);
            return coord;
        }
        static public void normal(object vec, double [] tr, object bv = null)
        {
            double[] t = tr ?? new double[] { -1, 1 };
            UnitVector2d uv2d = vec as UnitVector2d;
            UnitVector uv = vec as UnitVector;
            Vector v = vec as Vector;
            Vector2d v2d = vec as Vector2d;
            double[] coord = getCoord(vec);
            if (coord.Length == 3)
            {
                if (uv != null)
                {
                    var tmp =  uv.CrossProduct((UnitVector)bv);
                    coord = getCoord(tmp);
                    uv.PutUnitVectorData(coord);
                }
                if (v != null)
                {
                    var tmp = v.CrossProduct((Vector)bv);
                    coord = getCoord(tmp);
                    v.PutVectorData(coord);
                }
            } 
            else if (coord.Length == 2)
            {
                double[] r = { t[0]*coord[1], t[1]*coord[0] };
                if (uv2d != null) uv2d.PutUnitVectorData(r);
                else if (v2d != null) v2d.PutVectorData(r);
            }
        }
        static public Vector normal(UnitVector bv, Vector vec)
        {
            var v = I.CV(vec);
            normal(v, null, bv.AsVector());
            return v;
        }
        static public bool isRight(Edge e, Edge be, UnitVector n)
        {
            UnitVector2d n1 = getNorm(e, n), n2 = getNorm(be, n);
            if (eq(n1, n2)) return true;
            if (n1.IsParallelTo(n2, 0.001)) return false;
            UnitVector v1 = vector2dTo3d<UnitVector>(n1), v2 = vector2dTo3d<UnitVector>(n2);
            var s = v1.CrossProduct(v2);
            double d = s.X + s.Y + s.Z;
            return d <= 0;
        }
        static public UnitVector normal(Edge e, UnitVector n)
        {
            UnitVector v2 = e.StartVertex.Point.VectorTo(e.StopVertex.Point).AsUnitVector();
            var v = v2.CrossProduct(n);
            return v;
        }
        static public Vector normal(Edge e1, Point pt, UnitVector bn)
        {
            var v = getDir(e1);
            var vec = e1.StartVertex.Point.VectorTo(pt);
            var n = normal(bn, v);
            var d = vec.DotProduct(n);
            n.ScaleBy(d);
            return n;
        }
        static public Vector normal(Circle c, Plane pl)
        {
            var vec = c.Center.VectorTo(pl.RootPoint);
            var n = pl.Normal.AsVector();
            var d = vec.DotProduct(n);
            n.ScaleBy(d);
            return n;
        }
        static public Vector2d normal(LineSegment2d ls, Point2d pt)
        {
            return normal(ls.StartPoint, ls.EndPoint, pt);
        }
        static public Vector2d normal(Point2d pt1, Point2d pt2, Point2d pt3)
        {
            var v = pt1.VectorTo(pt2);
            var vec = pt1.VectorTo(pt3);
            var n = I.CV2d(v);
            normal(v, vec, out n);
            return n;
        }
        static public void normal(Vector2d v, Vector2d vec, out Vector2d n)
        {
            n = I.CV2d(v);
            normal(n, null);
            n.Normalize();
            var d = vec.DotProduct(n);
            n.ScaleBy(d);
        }
        static public double dotProduct(Vector v, Point pt)
        {
            return dotProduct(v, I.CV(pt.X, pt.Y, pt.Z));
        }
        static public double dotProduct(Vector2d v, Point2d pt, Point2d st, bool abs = false)
        {
            var vec1 = I.CV2d(v); vec1.Normalize();
            var vec = st.VectorTo(pt);
            var d = dotProduct(vec, vec1);
            if (abs) d = Math.Abs(d);
            return d;
        }
        static public double dotProduct(Vector v1, Vector v2)
        {
            var c1 = getCoord(v1); var c2 = getCoord(v2);
            double s = 0;
            for (int i = 0; i < c1.Length; i++)
            {
                if (eq(c1[i], 0) || eq(c2[i], 0)) continue;
                s += c1[i] * c2[i];
            }
            return s;
        }
        static public double dotProduct(Vector2d v, Point2d pt)
        {
            return dotProduct(v, I.CV2d(pt.X, pt.Y));
        }
        static public double dotProduct(Vector2d v1, Vector2d v2)
        {
            var c1 = getCoord(v1); var c2 = getCoord(v2);
            double s = 0;
            for (int i = 0; i < c1.Length; i++)
            {
                if (eq(c1[i], 0) || eq(c2[i], 0)) continue;
                s += c1[i] * c2[i];
            }
            return s;
        }
        static public bool outherTest(Edge e)
        {
            foreach (EdgeUse item in e.EdgeUses)
            {
                if (!item.IsOpposedToEdge) return item.EdgeLoop.IsOuterEdgeLoop;
            }
            return true;
        }
        static public int octant(Vector2d v)
        {
            double x = v.X, y = v.Y;
            int o = 0;
            if (eq(x,0))
            {
                o = y > 0 ? 3 : 7;
            }
            else if (eq(y, 0))
            {
                o = x > 0 ? 1 : 4;
            }
            else if (x > 0)
            {
                if (y > 0) o = x >= y ? 1 : 2;
                else if (y < 0) o = x >= Math.Abs(y) ? 8 : 7;
            }
            else if (x < 0)
            {
                if (y > 0) o = Math.Abs(x) >= y ? 4 : 3;
                else if (y < 0) o = Math.Abs(x) >= Math.Abs(y) ? 5 : 6;
            }
            return o;
        }
        static public double getDiam(Edge e)
        {
            Circle c = e.Geometry as Circle;
            Arc3d a = e.Geometry as Arc3d;
            return c != null ? c.Radius * 2: a.Radius * 2;
        }
        static public List<WorkPlaneProxy> getWPs(ComponentOccurrence occ)
        {
            var adef = occ.Definition as AssemblyComponentDefinition;
            var pdef = occ.Definition as PartComponentDefinition;
            List<WorkPlaneProxy> lst = new List<WorkPlaneProxy>();
            var wpls = adef != null ? adef.WorkPlanes : pdef.WorkPlanes;
            for (int i = 0; i < 3; i++)
            {
                object pr; occ.CreateGeometryProxy(wpls[i + 1], out pr);
                lst.Add(pr as WorkPlaneProxy);
            }
            return lst;
        }
        static public List<WorkPlane> getWPs(AssemblyComponentDefinition def, List<Vector> vecs)
        {
            List<WorkPlane> lst = new List<WorkPlane>();
            foreach (var item in vecs)
            {
                lst.Add(getWP(def, item));
            }
            return lst;
        }
        static public WorkPlane getWP(AssemblyComponentDefinition def, Vector v)
        {
            for (int i = 0; i < 3; i++)
            {
                var wp = def.WorkPlanes[i + 1];
                if (wp.Plane.Normal.IsParallelTo(v.AsUnitVector())) return wp;
            }
            return null;
        }
        static public void getPoints(Point2d c, double w, double h, out Point2d [] pts)
        {
            pts = new Point2d[4];
            pts[0] = I.CP2d(c, I.CV2d(-w * 0.5, -h * 0.5));
            pts[1] = I.CP2d(c, I.CV2d(w * 0.5, h * 0.5));
            pts[2] = I.CP2d(c, I.CV2d(-w * 0.5, h * 0.5));
            pts[3] = I.CP2d(c, I.CV2d(w * 0.5, -h * 0.5));
        }
        static public void getPoints(Point2d p1, Point2d p2, Vector2d v, out Vector2d n, out Point2d p3, out Point2d p4 )
        {
            var vec = p1.VectorTo(p2);
            n = I.CV2d(v);
            normal(v,vec,out n);
            if (eq(v.X, 0) || eq(v.Y, 0))
            {
                p3 = I.CP2d(p1.X, p2.Y);
                p4 = I.CP2d(p1.Y, p2.X);
            }
            else
            {
                n.Normalize();
                p3 = I.CP2d(p1.X + dotProduct(vec, v), p1.Y + dotProduct(vec, n));
                p4 = I.CP2d(p1.Y + dotProduct(vec, n), p1.X + dotProduct(vec, v));
            }
        }
        static public Matrix2d mirrorMtx(Line2d l)
        {
            var mtx = I.tg.CreateMatrix2d();
            var n = I.CUV2d(l.Direction);
            normal(l.Direction, null);
            mtx.SetCoordinateSystem(l.RootPoint, l.Direction.AsVector(), n.AsVector());
            mtx.Cell[1, 1] = -mtx.Cell[1,1];
            double[] r = { };
            mtx.GetMatrixData(ref r);
            return mtx;
        }
        static public void mirror(Matrix2d mtx, Point2d pt)
        {
            pt.TransformBy(mtx);
        }
        static public bool isParallel(Edge e1, Edge e2, UnitVector n, double tol = 0.01)
        {
            var v1 = e1.StartVertex.Point.VectorTo(e1.StopVertex.Point).AsUnitVector();
            var v2 = e2.StartVertex.Point.VectorTo(e2.StopVertex.Point).AsUnitVector();
            normal(v2, null, n);
            double d = v1.DotProduct(v2);
            return Math.Abs(d) <= tol;
        }
        static public bool isParallel(Edge e1, Edge e2)
        {
            LineSegment l1 = e1.Geometry as LineSegment, l2 = e2.Geometry as LineSegment;
            return l1.Direction.IsParallelTo(l2.Direction, 0.01);
        }
        static public bool testLine(Edge e1, Edge e2, Point cen, UnitVector n)
        {
            Vector v = e1.StartVertex.Point.VectorTo(e1.StopVertex.Point),
                v1 = cen.VectorTo(e1.StartVertex.Point), v2 = cen.VectorTo(e2.StartVertex.Point);
            Vector v3 = v1.CrossProduct(v), v4 = v2.CrossProduct(v);
            v3.Normalize(); v4.Normalize();
            return eq(v3, v4);
        }
        static public bool testBox(Box b, Edge e)
        {
            var b2 = e.Evaluator.RangeBox;
            bool inters = !b.IsDisjoint(b2), cont1 = b.Contains(b2.MinPoint), cont2 = b.Contains(b2.MaxPoint);
            return inters;
        }
        static public Box createBox(Edge e, Vector vnorm, double dist)
        {
            Vector rev = I.CV(vnorm, -1);
            Point ptb1 = I.CP(e.StopVertex.Point, vnorm, dist);
            Point ptb2 = I.CP(e.StartVertex.Point, rev, dist);
            double d = e.StartVertex.Point.DistanceTo(e.StopVertex.Point);
            Box b = I.Box(ptb1, ptb2);
            b.Expand(d*0.1);
            return b;
        }
        static public UnitVector2d getNorm(Edge e, UnitVector un)
        {
            var uv = vector3dTo2d(e.StartVertex.Point.VectorTo(e.StopVertex.Point), un);
            normal(uv, null);
            return uv;
        }
        static public bool isInner(Edge e)
        {
            var r = u.get<EdgeUse>(e.EdgeUses, fi => !fi.EdgeLoop.IsOuterEdgeLoop);
            return r != null ? true : false;
        }
        static public bool isFirstSheet(DrawingView dv, int num)
        {
            return isTDoc<DrawingDocument>(dv.Parent.Parent).Sheets[num].Equals(dv.Parent);
        }
        static public T isTDoc<T>(object doc) where T : class
        {
            return doc as T;
        }

        static public T docCompDef<T>(object doc) where T : class
        {
            PartDocument pdoc = isTDoc<PartDocument>(doc);
            return (pdoc != null && pdoc.SubType == "{9C464203-9BAE-11D3-8BAD-0060B0CE6BB4}") ? pdoc.ComponentDefinition as T : null;
        }

        static public T findAtPoint<T>(DrawingView dv, Func<DrawingCurve, bool> f) where T : class
        {
            foreach (DrawingCurve dc in dv.DrawingCurves)
            {
                if (f(dc)) return dc as T;
            }
            return null;
        }
        static public T findAtPoint<T>(PartComponentDefinition def, Point pt, SelectionFilterEnum[] sf, double tol
            , Func<T, bool> f) where T : class
        {
            ObjectsEnumerator objs = def.FindUsingPoint(pt, sf, tol, true);
            foreach (T item in objs)
            {
                if (f(item)) return item as T;
            }
            return null;
        }

        static public T findAtPoint<T>(List<DrawingCurve> dcs, Func<DrawingCurve, bool> f) where T : class
        {
            foreach (DrawingCurve dc in dcs)
            {
                if (f(dc)) return dc as T;
            }
            return null;
        }

        static public DrawingCurve findAtPoint(DrawingView dv, Point2d pt, double R, Func<DrawingCurve, bool> f)
        {
            List<DrawingCurve> cs = getCurves<DrawingCurve>(dv, f);
            return min<DrawingCurve>(cs.ToArray(), e => min<Point2d>(new Point2d[] { e.StartPoint, e.EndPoint }, p => p.VectorTo(pt).Length).VectorTo(pt).Length);
        }

        static public bool containts(Point2d pt, Point2d cen, double r)
        {
            return (pt.VectorTo(cen).Length < r) ? true : false;
        }

        static public Vector normal(Face face, ref Point pt)
        {
            double[] norm = new double[3], coord = new double[3], par = new double[2] { 0.5, 0.5 };
            SurfaceEvaluator eval = face.Evaluator;
            eval.GetPointAtParam(ref par, ref coord);
            eval.GetNormalAtPoint(ref coord, ref norm);
            pt = I.tg.CreatePoint(coord[0], coord[1], coord[2]);
            return I.tg.CreateVector(norm[0], norm[1], norm[2]);
        }    
        static public Matrix transformAsmToPart(ComponentOccurrence to)
        {
            var mtxto = to.Transformation;
            mtxto.Invert();
            return mtxto;
        }
        //static public KeyValuePair<Point,object> findNearestFace(ComponentDefinition cd, Face f, Point pt)
        //{
        //    if (f.SurfaceType != SurfaceTypeEnum.kPlaneSurface) return new KeyValuePair<Point, object>();
        //    Plane pl = (Plane)f.Geometry;
        //    UnitVector n = pl.Normal, r = u.reverse(n);
        //    Point ptn = I.CP(pt, n.AsVector(), 0.01), ptr = I.CP(pt, r.AsVector(), 0.01);
        //    var pn = findAtRay(cd, ptn, n, fi => fi is Face || fi is FaceProxy);
        //    var pr = findAtRay(cd, ptr, r, fi => fi is Face || fi is FaceProxy);
        //    if (pn.Key != null && pr.Key == null) return pn;
        //    else if (pn.Key == null && pr.Key != null) return pr;
        //    else if (pn.Key != null && pr.Key != null)
        //    {
        //        double d1 = pt.DistanceTo(pn.Key), d2 = pt.DistanceTo(pr.Key);
        //        return u.eq(d1, 0) ? pr : eq(d2, 0) ? pn :
        //        (d1 <= d2) ? pn : pr;
        //    }
        //    return new KeyValuePair<Point, object>();
        //}
        static public FaceData findNearestFace(PartComponentDefinition cd, FaceProxy f, Point pt)
        {
            Plane pl = (Plane)f.NativeObject.Geometry;
            List<FaceData> tmp = new List<FaceData>();
            UnitVector v = pl.Normal; v = reverse(v);
            findUsingRay(cd, pt, v, (o, l) =>
            {
                FaceData fd = new FaceData(o, (Point)l);
                if (fd.f != null)
                {
                    fd.setSP(pt);
                    tmp.Add(fd);
                }
            }, 0.1, true);
            //tmp = tmp.Where(el => el.f != null).ToList();
            if (tmp.Count >= 2)
            {
                int i = 1;
                if (!tmp[i - 1].check(tmp[i])) i = 2;
                tmp[i].setSP(pt);
                return tmp[i];
            }
            else if (tmp.Count == 1) return tmp[0];
            return new FaceData(null, null);
        }
        static public bool pointOnEdge(Edge e, Point pt)
        {
            Vector v1 = e.StartVertex.Point.VectorTo(pt),
                v2 = e.StartVertex.Point.VectorTo(e.StopVertex.Point);
            var c = isCollinear(v1.AsUnitVector(), v2.AsUnitVector());
            return (v1.Length <= v2.Length && c);
        }
        static public double distToEdge(Edge e, Point pt)
        {
            pt = pt ?? I.CP();
            var f = get<Face>(e.Faces, fi => fi.SurfaceType == SurfaceTypeEnum.kPlaneSurface);
            Plane pl = (Plane)f.Geometry;
            return distToPlane(pl, pt);
        }
        static public PartFeatureExtentDirectionEnum holeDir(Plane pl, PlanarSketch ps)
        {
            var n = pl.Normal;
            var v = ps.PlanarEntityGeometry.Normal;
            return (u.dotProduct(v.AsVector(), n.AsVector()) > 0) ? PartFeatureExtentDirectionEnum.kNegativeExtentDirection :
                PartFeatureExtentDirectionEnum.kPositiveExtentDirection;
        }
        static public KeyValuePair<Point, object> findAtRay(ComponentDefinition cd, Point pt, UnitVector v, Func<object, bool> f, double r = 0.1)
        {
            ObjectsEnumerator objs = null, loc = null;
            if (cd is PartComponentDefinition)
            {
                ((PartComponentDefinition)cd).FindUsingRay(pt, v, r, out objs, out loc);
            }
            else if (cd is AssemblyComponentDefinition)
            {
                ((AssemblyComponentDefinition)cd).FindUsingRay(pt, v, r, out objs, out loc);
            }
            int i = 0;
            foreach (Point item in loc)
            {
                i++;
                Point p = item;
                if (f(objs[i])) return new KeyValuePair<Point, object>( p, objs[i] );
            }
            return new KeyValuePair<Point, object>();
        }
        static public Point getPoint(Face face, Edge ed, double offset)
        {
            Vector vec = null; Point mPt = ed.StartVertex.Point;
            Edge ed2 = null;
            foreach (Edge item in face.Edges)
            {
                if (!item.Equals(ed) && getLenght(item) == getLenght(ed))
                {
                    ed2 = item;
                    vec = midPt(ed.StartVertex.Point, ed.StopVertex.Point).VectorTo(midPt(ed2.StartVertex.Point, ed2.StopVertex.Point));
                    if (vec.Length < offset) { vec = null; continue; }
                    vec.Normalize();
                    vec.ScaleBy(offset);
                    break;
                }
                //                 if (eq(item.StartVertex, ed.StopVertex))
                //                 {
                //                     vec = item.StartVertex.Point.VectorTo(item.StopVertex.Point);
                //                     if (vec.Length < offset) continue;
                //                     vec.Normalize();             
                //                     vec.ScaleBy(offset);
                //                     break;
                //                 }
                //                 if (eq(item.StopVertex, ed.StartVertex))
                //                 {
                //                     vec = item.StartVertex.Point.VectorTo(item.StopVertex.Point);
                //                     if (vec.Length < offset) continue;
                //                     vec.Normalize();
                //                     vec.ScaleBy(-offset);
                //                     break;
                //                 }
                //                 if (eq(item.StartVertex, ed.StartVertex))
                //                 {
                //                     vec = item.StartVertex.Point.VectorTo(item.StopVertex.Point);
                //                     if (vec.Length < offset) continue;
                //                     vec.Normalize();
                //                     vec.ScaleBy(offset);
                //                     break;
                //                 }
            }
            if (vec != null)
            {
                mPt.TranslateBy(vec);
                return mPt;
            }
            return null;
        }

        static public bool inRange(Point2d pt, Point2d cen, double w, double h)
        {
            if ((pt.X <= cen.X + w / 2 && pt.X >= cen.X - w / 2) && (pt.Y <= cen.Y + h / 2 && pt.Y >= cen.Y - h / 2)) return true;
            return false;
        }

        public enum RetCode { In, Out, Err }

        static public RetCode inBody(ComponentDefinition compDef, Edge ed, SelectionFilterEnum[] filter, double tol, double offset)
        {
            Face f1 = ed.Faces[1], f2 = ed.Faces[2];

            Point p1 = getPoint(f1, ed, offset), p2 = getPoint(f2, ed, offset);
            if (p1 == null || p2 == null) return RetCode.Err;
            p1 = midPt(p1, p2);
            ObjectsEnumerator en = compDef.FindUsingPoint(p1, ref filter, tol);
            return (en.Count == 0) ? RetCode.In : RetCode.Out;
        }

        static public bool inBody(SheetMetalComponentDefinition compDef, Edge ed, double offset)
        {
            HashSet<Edge> eds = new HashSet<Edge>();
            SelectionFilterEnum[] filter = new SelectionFilterEnum[] { SelectionFilterEnum.kPartEdgeLinearFilter };
            ObjectsEnumerator en = compDef.FindUsingPoint(ed.StartVertex.Point, filter, 0.01);
            UnitVector vec = ed.StartVertex.Point.VectorTo(ed.StopVertex.Point).AsUnitVector();
            HashSet<Edge> tmp = new HashSet<Edge>();
            //if (getLenght(ed) < offset) return false;
            if (addEdge(en, vec, ref eds))
            {
                foreach (Edge item in eds)
                {
                    if (getLenght(item) < offset) return false;
                }
                foreach (EdgeUse item in eds.ElementAt(0).EdgeUses)
                {
                    if (!item.EdgeLoop.IsOuterEdgeLoop) return false;
                }
            }
            return true;
        }

        static public RetCode inBody(ComponentDefinition compDed, Edge ed, double offset)
        {
            HashSet<Edge> eds = new HashSet<Edge>();
            SelectionFilterEnum[] filter = new SelectionFilterEnum[] { SelectionFilterEnum.kPartEdgeLinearFilter };
            ObjectsEnumerator en = compDed.FindUsingPoint(ed.StartVertex.Point, filter, 0.01);
            UnitVector vec = ed.StartVertex.Point.VectorTo(ed.StopVertex.Point).AsUnitVector();
            HashSet<Edge> tmp = new HashSet<Edge>();
            if (addEdge(en, vec, ref eds))
            {
                foreach (Edge item in eds)
                {
                    if (getLenght(item) < offset) return RetCode.Err;
                    //                     if (eq(ed.StartVertex, item.StopVertex))
                    //                     {
                    //                         en = compDed.FindUsingPoint(item.StartVertex.Point, filter, 0.01);     
                    //                     }
                    //                     else en = compDed.FindUsingPoint(item.StopVertex.Point, filter, 0.01);
                    //                     addEdge(en, vec, ref tmp);
                }
                foreach (EdgeUse item in eds.ElementAt(0).EdgeUses)
                {
                    if (!item.EdgeLoop.IsOuterEdgeLoop) return RetCode.Err;
                }



                //foreach (Edge item in tmp)
                //{
                //    eds.Add(item); 
                //}
                //                  if (eds.Count == 4)              
                //                  {
                // List<Edge> e = eds.ToList();
                //if (leftOrRight(eds.ElementAt(0), eds.ElementAt(1), vec) == leftOrRight(eds.ElementAt(0), eds.ElementAt(2), vec)
                //    && leftOrRight(eds.ElementAt(0), eds.ElementAt(1), vec))
                //    return RetCode.Err;


                //PartDocument doc = I.app.ActiveDocument as PartDocument;
                //Vector vect = eds.ElementAt(0).StartVertex.Point.VectorTo(eds.ElementAt(0).StopVertex.Point);
                //vect.ScaleBy(2);
                //Point pt = eds.ElementAt(0).StartVertex.Point.Copy();
                //pt.TranslateBy(vect);
                //addClientLine(doc, "t", doc.Assets[2], eds.ElementAt(0).StartVertex.Point, pt);

                //vect = eds.ElementAt(1).StartVertex.Point.VectorTo(eds.ElementAt(1).StopVertex.Point);
                //vect.ScaleBy(2);
                //pt = eds.ElementAt(1).StartVertex.Point.Copy();
                //pt.TranslateBy(vect);
                //addClientLine(doc, "n", doc.Assets[6], eds.ElementAt(1).StartVertex.Point, pt);

                //if (isDir(e[1], e[2]) && isDir(e[0],e[3]))

                //}

                //                     if (leftOrRight(eds.ElementAt(0), eds.ElementAt(1), vec))
                //                         return RetCode.In;
                return RetCode.Out;
            }
            return RetCode.Out;
        }

        static public void drawDirection(FlatPattern compDef, Asset asst, Edge ed, string name)
        {
            EdgeLoop el = null;
            foreach (EdgeUse item in ed.EdgeUses)
            {
                if (item.EdgeLoop.Edges.Count > 4) el = item.EdgeLoop;
            }
            if (el == null) return;
            foreach (Edge item in el.Edges)
            {
                Vector vect = item.StartVertex.Point.VectorTo(item.StopVertex.Point);
                vect.ScaleBy(1.2);
                Point pt = item.StartVertex.Point.Copy();
                pt.TranslateBy(vect);
                addClientLine(compDef, name, asst, item.StartVertex.Point, pt);
            }
        }

        static public void drawDirection(PartComponentDefinition compDef, Asset asst, string name)
        {
            //EdgeLoop el = null;
            //if (el == null) return;
            Color col = I.objs.CreateColor(0, 255, 0);
            hs = I.app.ActiveDocument.CreateHighlightSet();
            hs.Color = col;
            foreach (Edge item in compDef.SurfaceBodies[1].ConvexEdges)
            {
                hs.AddItem(item);
                //                 Vector vect = item.StartVertex.Point.VectorTo(item.StopVertex.Point);
                //                 vect.ScaleBy(1.2);
                //                 Point pt = item.StartVertex.Point.Copy();
                //                 pt.TranslateBy(vect);
                //                 addClientLine(compDef.Document as PartDocument, name, asst, item.StartVertex.Point, pt);
            }
        }

        static public bool isDir(Edge ed1, Edge ed2)
        {
            UnitVector vec1 = ed1.StartVertex.Point.VectorTo(ed1.StopVertex.Point).AsUnitVector(),
                vec2 = ed2.StartVertex.Point.VectorTo(ed2.StopVertex.Point).AsUnitVector();
            return vec1.IsEqualTo(vec2, 0.01);
        }

        static public bool leftOrRight(Edge ed1, Edge ed2, UnitVector comp = null)
        {
            UnitVector vec1 = ed1.StartVertex.Point.VectorTo(ed1.StopVertex.Point).AsUnitVector(),
                vec2 = ed2.StartVertex.Point.VectorTo(ed2.StopVertex.Point).AsUnitVector(), vec3;
            //if (eq(ed1.StartVertex, ed2.StopVertex)) vec3 = vec1.CrossProduct(vec2);
            /*else*/
            vec3 = vec2.CrossProduct(vec1);
            //return !minus(vec3);
            double dp = vec3.X + vec3.Y + vec3.Z;
            return (dp < 0) ? true : false;
            //return vec3.IsEqualTo(comp, 0.01);
        }

        static public bool minus(UnitVector vec)
        {
            return (vec.X < 0 || vec.Y < 0 || vec.Z < 0) ? true : false;
        }

        static public void translatePts(Vector2d v, ref List<Point2d> pts)
        {
            foreach (var item in pts)
            {
                item.TranslateBy(v);
            }
        }

        static public List<Point2d> getSplPoints(SketchSpline spl)
        {
            List<Point2d> pts = new List<Point2d>();
            foreach (SplineFitPointConstraint item in spl.Constraints)
            {
                pts.Add(item.Point.Geometry); 
            }
            return pts;
        }

        static public DrawingSketch addSpline(DrawingView dv, List<Point2d> pts, bool trans = false)
        {
            DrawingSketch ds = dv.Sketches.Add();
            ds.Edit();
            Matrix2d mtx = I.getMatrix2d();
            mtx.Cell[1, 1] = 1 / dv.Scale;
            mtx.Cell[2, 2] = 1 / dv.Scale;
            ObjectCollection col = I.COC();
            foreach (var item in pts)
	        {
                if (trans)
                item.TransformBy(mtx);
		        col.Add(item);
	        }
            SketchSpline spl = ds.SketchSplines.Add(col, SplineFitMethodEnum.kSmoothSplineFit);
            spl.Closed = true;
            ds.ExitEdit();
            return ds;
        }

        static public void addCut(DrawingSketch ds)
        {
            DrawingView dv = ds.Parent as DrawingView;
            Profile pr = ds.Profiles.AddForSolid();
            ObjectCollection col = I.COC();
            findOcc(dv, new List<string>() { "Крышка" , "Стенка боковая"}, ref col);
            dv.BreakOutOperations.Add(pr, col);
        }

        static public void findOcc(DrawingView dv, List<string> names, ref ObjectCollection col)
        {
            AssemblyDocument asm = dv.ReferencedDocumentDescriptor.ReferencedDocument as AssemblyDocument;
            if (asm == null) return;
            foreach (ComponentOccurrence occ in asm.ComponentDefinition.Occurrences)
            {
                Property p = getProp(occ.ReferencedDocumentDescriptor.ReferencedDocument as Document, "Description");
                if (p == null) continue;
                foreach (var item in names)
                {
                    if (p.Value.ToString().ToLower() == item.ToLower()) col.Add(occ); 
                }
            }
        }

        static public bool addEdge(ObjectsEnumerator en, UnitVector vec, ref HashSet<Edge> eds)
        {
            bool flag = false;
            //             if (en.Count == 3)
            //             {
            foreach (Edge item in en)
            {

                UnitVector v = item.StartVertex.Point.VectorTo(item.StopVertex.Point).AsUnitVector();
                if (v.IsPerpendicularTo(vec, 0.01))
                {
                    eds.Add(item);
                    flag = true;
                }
            }

            // }
            return flag;
        }

        //         static public double getParam(DrawingCurve dc, Point2d pt)
        //         {
        //             Curve2dEvaluator eval = dc.Evaluator2D;
        //             eval.GetParamAtPoint()
        //         }

        static public bool intersectRect(Point2d pt1, double width1, double height1, Point2d pt2, double width2, double height2)
        {
            return intersectRect(pt1.X - width1 / 2, pt1.Y - height1 / 2, pt1.X + width1 / 2, pt1.Y + height1 / 2,
                pt2.X - width2 / 2, pt2.Y - height2 / 2, pt2.X + width2 / 2, pt2.Y + height2 / 2);
        }

        static public bool intersectRect(double xmin1, double ymin1, double xmax1, double ymax1, double xmin2, double ymin2, double xmax2, double ymax2)
        {
            return (ptInRect(xmin1, ymin1, xmin2, xmax2, ymin2, ymax2) || ptInRect(xmin1, ymax1, xmin2, xmax2, ymin2, ymax2) ||
                ptInRect(xmax1, ymax1, xmin2, xmax2, ymin2, ymax2) || ptInRect(xmax1, ymin1, xmin2, xmax2, ymin2, ymax2));
        }

        static public bool ptInRect(double x, double y, double minX, double maxX, double minY, double maxY)
        {
            return ((inTheInterval(x, minX, maxX) && inTheInterval(y, minY, maxY)));
        }

        static public bool inTheInterval(double cur, double min, double max)
        {
            return (cur <= max && cur >= min);
        }

        static public Vector2d normal(Point2d startPt, Point2d endPt)
        {
            Matrix2d mtx = I.tg.CreateMatrix2d();
            mtx.SetToRotation(Math.PI / 2, startPt);
            Vector2d vec = startPt.VectorTo(endPt);
            vec.TransformBy(mtx);
            return vec;
        }
        public static void setDist(Vector2d v, double d)
        {
            v.Normalize();
            v.ScaleBy(d);
        }
        static public Vector2d rotate(Vector2d vec, Point2d startPt, double angle)
        {
            Matrix2d mtx = I.tg.CreateMatrix2d();
            mtx.SetToRotation(angle, startPt);
            vec.TransformBy(mtx);
            return vec;
        }

        static public Point2d translate(Point2d pt, Vector2d vec)
        {
            Point2d copyPt = pt.Copy();
            copyPt.TranslateBy(vec);
            return copyPt;
        }

        public static double getDist(Point2d pt, DrawingCurve line)
        {
            MeasureTools mt = I.app.MeasureTools;
            double min = mt.GetMinimumDistance(pt, line);
            return min;
        }

        public static Double getDist(Point2d pt, DrawingCurve line, double scale, ref Point2d retPt)
        {
            LineSegment2d l = ray(line.StartPoint, line.EndPoint, scale);
            Vector2d norm = normal(line.StartPoint, line.EndPoint); norm.Normalize(); norm.ScaleBy(scale);
            LineSegment2d n = ray(pt, norm);
            ObjectsEnumerator col = n.IntersectWithCurve(l);
            if (col == null) return 10000;
            retPt = (Point2d)col[1];
            Double dist = pt.DistanceTo((Point2d)col[1]);
            return dist;
        }

        public static LineSegment2d ray(Point2d startPt, Point2d endPt, double scale)
        {
            Vector2d vec = startPt.VectorTo(endPt); vec.Normalize(); vec.ScaleBy(scale);
            return ray(startPt, vec);
        }

        public static LineSegment2d ray(Point2d pt, Vector2d vec)
        {
            Point2d startPt = pt.Copy(), endPt = pt.Copy();
            endPt.TranslateBy(vec); vec.ScaleBy(-1);
            startPt.TranslateBy(vec);
            return I.tg.CreateLineSegment2d(endPt, startPt);
        }

        public static Point getPoint(ComponentOccurrence occ, FaceProxy face)
        {
            if (face.SurfaceType == SurfaceTypeEnum.kCylinderSurface)
            {
                double r = ((Cylinder)face.Geometry).Radius;
                FaceProxy f = occ.SurfaceBodies[1].Faces.OfType<FaceProxy>().FirstOrDefault(e => e.SurfaceType == SurfaceTypeEnum.kCylinderSurface && eq(((Cylinder)e.Geometry).Radius, r));
                return f.Vertices[1].Point;
            }
            else if (face.SurfaceType == SurfaceTypeEnum.kPlaneSurface)
            {
                return occ.SurfaceBodies[1].Vertices[1].Point;
            }
            return null;
        }

        public static bool intersPoint(AssemblyComponentDefinition compDef, ComponentOccurrence occ, FaceProxy face)
        {
            Inventor.Point ptmc = getPoint(occ, face);
            SelectionFilterEnum[] f = { SelectionFilterEnum.kPartFaceFilter };
            ObjectsEnumerator en = compDef.FindUsingPoint(ptmc, ref f, 0.1);
            Face fac = en.OfType<Face>().FirstOrDefault(e => e.InternalName == face.InternalName);
            return (fac == null) ? true : false;
        }

        static public Vector getAxis(Face face)
        {
            Vector vec = null;
            if (face.SurfaceType == SurfaceTypeEnum.kCylinderSurface)
            {
                Cylinder cyl = face.Geometry as Cylinder;
                vec = cyl.AxisVector.AsVector();
            }
            else if (face.SurfaceType == SurfaceTypeEnum.kPlaneSurface)
            {
                Plane pl = face.Geometry as Plane;
                vec = pl.Normal.AsVector();
            }
            return vec;
        }

        public static Edge findStartEdge(PartFeature pf)
        {
            foreach (Face f in pf.Faces)
            {
                if (f.SurfaceType != SurfaceTypeEnum.kCylinderSurface) continue;
                foreach (Edge e in f.Edges)
                {
                    if (e.CurveType != CurveTypeEnum.kLineCurve) continue;
                    foreach (Face item in e.Faces)
                    {
                        if (!item.CreatedByFeature.Equals(pf)) return e;
                    }
                }
            }
            return null;
        }

        public static List<Face> findTangFaces(PartFeature pf, Edge e)
        {
            List<Face> fs = new List<Face>();
            Face f = e.Faces[1].CreatedByFeature.Equals(pf) ? e.Faces[1] : e.Faces[2];
            fs.Add(f);
            foreach (Face item in f.TangentiallyConnectedFaces)
            {
                if (!item.CreatedByFeature.Equals(pf)) continue;
                fs.Add(item);
            }
            return fs;
        }
        public static Edge findConnectEdge(Face f1, Face f2)
        {
            foreach (Edge e in f1.Edges)
            {
                int n = 0;
                foreach (Face f in e.Faces)
                {
                    if (f.Equals(f1) || f.Equals(f2)) n++;
                }
                if (n == 2) return e;
            }
            return null;
        }

        public static Vertex setOrigin(Face f, Point pt, double t, PlanarSketch ps,
            out Point2d o2, out UnitVector2d x2, out UnitVector2d y2,
            SheetMetalComponentDefinition def = null)
        {
            UnitVector x, y;
            var vert = u.getOrigin(f, pt, t, out x, out y, def);
            u.getOrigin2d(ps, pt, x, y, out o2, out x2, out y2);
            return vert;
        }

        static public Vertex getOrigin(Face face, Point pt, double t, out UnitVector x, out UnitVector y,
            SheetMetalComponentDefinition def = null)
        {
            var vert = u.min<Vertex>(face.Vertices, el => el.Point.DistanceTo(pt));
            //if (def != null) u.clientTxt(def as PartComponentDefinition, vert.Point, "near");
            pt = vert.Point;
            var pl = face.Geometry as Plane;
            x = u.createUnitVector(0, 0, 1); y = u.createUnitVector(0, 0, 1);
            foreach (Edge item in vert.Edges)
            {
                var vec = u.getVector(item, pt);
                if (vec == null) continue;
                if (vec.IsParallelTo(pl.Normal.AsVector())) continue;
                if (u.eq(vec.Length, t))
                    y = vec.AsUnitVector();
                else if (vec.IsPerpendicularTo(pl.Normal.AsVector())) x = vec.AsUnitVector();
            }
            return vert;
        }

        static public void getOrigin2d(PlanarSketch ps, Point pt, UnitVector x, UnitVector y,
            out Point2d o, out UnitVector2d x2, out UnitVector2d y2)
        {
            o = ps.ModelToSketchSpace(pt);
            x2 = uv3dto2d(x, ps, pt);
            y2 = uv3dto2d(y, ps, pt);
        }

        static public UnitVector2d uv3dto2d(UnitVector v, PlanarSketch ps, Point pt)
        {
            var ep = pt.Copy(); ep.TranslateBy(v.AsVector());
            var p1 = ps.ModelToSketchSpace(pt);
            var p2 = ps.ModelToSketchSpace(ep);
            return p1.VectorTo(p2).AsUnitVector();
        }

        static public Edge getEdge(object obj, double t, double r)
        {
            dynamic f = obj;
            foreach (Edge e in f.Edges)
            {
                if (e.CurveType != CurveTypeEnum.kLineCurve) continue;
                foreach (Face item in e.Faces)
                {
                    var cyl = item.Geometry as Cylinder;
                    if (cyl == null) continue;
                    if (u.eq(cyl.Radius, r) || u.eq(cyl.Radius, r + t)) return e;
                }
            }
            return null;
        }

        static public Edge getAxis(PlanarSketch ps, double t, double r, Document doc, out Point pt, out UnitVector x, out UnitVector y)
        {
            Face f = ps.PlanarEntity as Face;
            var e1 = getEdge(f, t, r);
            pt = e1.StartVertex.Point;
            x = null; y = null;
            List<object> hs = new List<object>();
            foreach (Edge item in e1.StartVertex.Edges)
            {
                if (item.CurveType != CurveTypeEnum.kLineCurve) continue;
                if (u.eq(u.len(item), t)) continue;
                if (item.Equals(e1)) { x = u.getVector(item, pt).AsUnitVector(); hs.Add(item); }
                else { y = u.getVector(item, pt).AsUnitVector(); hs.Add(item); }
            }
            u.higlight(doc, hs[0], hs[1], null, null);
            return e1;
            //pt = u.midPt(e1);
            //higlight(doc, e1, e1.StartVertex, null, null);
            //x = e1.StartVertex.Point.VectorTo(e1.StopVertex.Point).AsUnitVector();
            //y = 

            //Face f2 = e1.Faces[1].SurfaceType == SurfaceTypeEnum.kCylinderSurface ? e1.Faces[1] : e1.Faces[2];
            //var e2 = u.get<Edge>(f2.Edges, el => !el.Equals(e1) && u.check(el, e1, (a, b) => a.IsParallelTo(b)));
            //var p2 = u.midPt(e2);

            //higlight(doc, e1, e2, null, null);

            //Plane pl = ps.PlanarEntityGeometry;

            //var d = u.distToPlane(pl, p2);
            //var norm = pl.Normal.AsVector();
            //norm.ScaleBy(d);
            //p2.TranslateBy(norm);
            //var pt2 = ps.ModelToSketchSpace(pt);
            //var pt3 = ps.ModelToSketchSpace(p2);

            //p2 = ps.SketchToModelSpace(pt3);
            //pt = ps.SketchToModelSpace(pt2);

            //x = p2.VectorTo(pt).AsUnitVector();
            //y = x.CrossProduct(pl.Normal);
        }

        static public Vector getAxis(ComponentOccurrence occ)
        {
            Document doc = occ.Definition.Document as Document;
            ComponentDefinition compDef = null;
            Vector vec = null;
            if (doc.DocumentType == DocumentTypeEnum.kAssemblyDocumentObject)
            {
                compDef = (ComponentDefinition)((AssemblyDocument)doc).ComponentDefinition;
            }
            else if (doc.DocumentType == DocumentTypeEnum.kPartDocumentObject)
            {
                compDef = (ComponentDefinition)((PartDocument)doc).ComponentDefinition;
            }
            Face f = occ.SurfaceBodies[1].Faces.OfType<Face>().FirstOrDefault(e => e.SurfaceType == SurfaceTypeEnum.kCylinderSurface);
            if (f == null)
            {
                return I.tg.CreateVector(0, 0, 1);
            }
            vec = getAxis(f);
            return vec;
        }

        static public WorkAxis getWAxis(PartComponentDefinition def, UnitVector dir)
        {
            return get<WorkAxis>(def.WorkAxes, f => f.Line.Direction.IsParallelTo(dir));
        }

        public static WorkPlane getPlane(PartDocument doc, string name)
        {
            return doc.ComponentDefinition.WorkPlanes[name];
        }

        public static WorkPlane getPlane(FaceProxy face, UnitVector norm)
        {
            PartDocument pDoc = face.ContainingOccurrence.Definition.Document as PartDocument;
            foreach (WorkPlane item in pDoc.ComponentDefinition.WorkPlanes)
            {
                if (eq(item.Plane.Normal, norm)) return item;
            }
            return null;
        }
        public static void updatePlane(WorkPlane wp, string offset)
        {
            ModelParameter mp = wp.DrivenBy.OfType<ModelParameter>().FirstOrDefault();
            if (mp == null) return;
            mp.Expression = offset;
        }

        public static WorkPlane addPlane(PartComponentDefinition compDef, string name, string baseName, string offset, string rev)
        {
            WorkPlane pl = compDef.WorkPlanes.OfType<WorkPlane>().FirstOrDefault(e => e.Name == name);
            if (pl != null) { updatePlane(pl, offset); return pl; }
            pl = compDef.WorkPlanes.OfType<WorkPlane>().FirstOrDefault(e => e.Name == baseName);
            if (pl == null) return null;
            pl = compDef.WorkPlanes.AddByPlaneAndOffset(pl, offset);
            pl.Name = name;
            if (rev != null) pl.FlipNormal();
            return pl;
        }

        public static CustomTable addCustomTable(CustomTable tbl, string name)
        {
            Point2d insPoint = tbl.Position;
            Sheet sh = tbl.Parent as Sheet;
            tbl.Delete();
            string[] col = new string[] { "BEND ID", "BEND DIRECTION", "BEND ANGLE", "BEND RADIUS" };
            return sh.CustomTables.AddBendTable(sh.DrawingViews[1].ReferencedDocumentDescriptor.FullDocumentName, insPoint, name, col);
        }
        public static void planeToXML(XMLDoc parent, WorkPlane plane)
        {
            //             if (plane.DefinitionType == WorkPlaneDefinitionEnum.kPlaneAndOffsetWorkPlane)
            //                 foreach (WorkPlane item in plane.DrivenBy)
            // 	            {
            // 		            planeToXML(parent, item);
            // 	            }
            XElement el;
            if (plane.DefinitionType == WorkPlaneDefinitionEnum.kPlaneAndOffsetWorkPlane)
            {
                PlaneAndOffsetWorkPlaneDef def = plane.Definition as PlaneAndOffsetWorkPlaneDef;
                Parameter param = def.Offset;
                parameterToXML(parent, param);
                string name = plane.Name;
                if (plane.DrivenBy.Count == 1 && plane.DrivenBy[1] is MirrorFeature) name = (plane.DrivenBy[1] as MirrorFeature).Name;
                el = new XElement("Plane", new XAttribute("Name", name), new XAttribute("BasePlane", (def.Plane as WorkPlane).Name), new XAttribute("Offset", def.Offset.Name));
                parent.El.Add(el);
            }
        }

        public static void sketchToXML(XMLDoc parent, PlanarSketch ps)
        {
            XElement el = new XElement("Sketch", new XAttribute("Name", ps.Name), new XAttribute("BasePlane", (ps.PlanarEntity as WorkPlane).Name));
            parent.El.Add(el);
        }

        public static PlanarSketch addSketch(PartComponentDefinition compDef, string name, string baseName)
        {
            WorkPlane pl = compDef.WorkPlanes.OfType<WorkPlane>().FirstOrDefault(e => e.Name == baseName);
            if (pl == null) return null;
            PlanarSketch ps = compDef.Sketches.Add(pl);
            ps.Name = name;
            return ps;
        }

        public static void joinIMates(Document doc)
        {
            if (isNull(doc)) return;
            AssemblyComponentDefinition acd = I.getACD(doc);
            if (isNull(acd)) return;
            var ss = doc.SelectSet;
            if (ss.Count == 0)
            {
                foreach (ComponentOccurrence oc in acd.Occurrences)
                {
                    joinIMate(oc);
                }
            }
            else
            {
                foreach (var oc in ss)
                {
                    if (oc is ComponentOccurrence)
                    joinIMate(oc as ComponentOccurrence);
                }
            }
        }
        public static void joinIMate(ComponentOccurrence oc)
        {
            var acd = oc.Parent;
            if (oc.Suppressed) return;
            foreach (var im in oc.iMateDefinitions.OfType<CompositeiMateDefinitionProxy>())
            {
                if (isNull(im)) continue;
                if (im.Type != ObjectTypeEnum.kCompositeiMateDefinitionProxyObject) continue;
                if (im.IsConsumed) continue;
                try
                {
                    if (im.MatchList == null) continue;
                    string[] ml = im.MatchList as string[];
                    if (isNull(ml)) continue;
                    joinIMate(acd, ml, im);
                }
                catch (System.Exception ex)
                {
                }

            }
        }

        public static void joinIMate(AssemblyComponentDefinition acd, string[] ml, CompositeiMateDefinitionProxy im)
        {
            //List<string> lst = ml.ToList();
            string imName = im.ContainingOccurrence.ReferencedDocumentDescriptor.FullDocumentName;
            if (imName.IndexOf("Content Center") != -1) return;
            foreach (var mn in ml)
            {
                string mName = mn;
                foreach (ComponentOccurrence oc in acd.Occurrences)
                {
                    if (oc.Suppressed) continue;
                    foreach (var d in oc.iMateDefinitions.OfType<CompositeiMateDefinitionProxy>())
                    {
                        if (isNull(d)) continue;
                        if (d.Type != ObjectTypeEnum.kCompositeiMateDefinitionProxyObject) continue;
                        if (d.IsConsumed) continue;
                        if (im.Equals(d)) continue;
                        string dName = d.ContainingOccurrence.ReferencedDocumentDescriptor.FullDocumentName;
                        if (dName == imName) continue;
                        //if (dName.IndexOf("Content Center") != -1) continue;
                        if (im.Count != d.Count) continue;
                        if (mName.StartsWith(d.Name))
                        {
                            iMateResult r = null;
                            if (im.IsConsumed || d.IsConsumed) continue;
//                             if (im.MatchList == null) continue;
//                             string[] dml = d.MatchList as string[];
//                             if (isNull(dml)) continue;
//                             var lst = dml.ToList();
                            if (!ml.Contains(d.Name)) continue;
                            iMateDefinition d1 = im as iMateDefinition, d2 = d as iMateDefinition;
                            try
                            {
                                r = acd.iMateResults.AddByTwoiMates(d1, d2);
                                //if (mName != null) r.Name = mName;
                            }
                            catch (Exception)
                            {
                                //r.Delete();
                            }
                            bool del = false;
                            if (r == null) return;
                            foreach (AssemblyConstraint item in r.Constraints)
                            {
                                if (item.HealthStatus != HealthStatusEnum.kUpToDateHealth) { del = true; break; }
                            }
                            if (del) r.Delete();
                        }
                    }
                }
            }
        }

        public static double convToDouble(string val, char[] te = null)
        {
            if (isNull(val)) return 0;
            if (te != null) val = val.TrimEnd(te);
            return double.Parse(val.Replace(',', '.'), System.Globalization.CultureInfo.InvariantCulture.NumberFormat);
        }

        public static string convToModelName(string val)
        {
            string r = "";
            char[] sep = new char[] { ' ', '_' };
            if (val.IndexOfAny(sep) != -1)
            {
                string[] spl = val.Split(sep);
                foreach (var item in spl)
                {
                    r += item[0].ToString().ToUpper() + item.Substring(1).ToLower();
                }
                return r;
            }
            return val;
        }

        public static double convToDouble(string val, double scale, int tol)
        {
            if (val == "") return 0;
            return Math.Round(convToDouble(val) * scale, tol);
        }

        public static string convToString(double val, int tol, string format)
        {
            return Math.Round(val, tol).ToString(format);
        }

        public static int convToInt(string val)
        {
            if (isNull(val)) return 0;
            int r;
            int.TryParse(val, out r);
            return r;
        }

        public static void round(Parameter p, int tol = 3)
        {
            double v = radToDeg(p.ModelValue);
            double d = (Math.Abs(v - (int)v));

            if (d < 0.02 || d > 0.98) tol = 1;
            else if (d < 0.002 || d > 0.998) tol = 2;
            
            p.Expression = round(v, tol).ToString();
        }

        public static Point2d addPt2d(double x, double y)
        {
            return I.tg.CreatePoint2d(x, y);
        }

        public static Point2d addPt2d(double x, string y, double scale = 1)
        {
            return addPt2d(x / scale, -convToDouble(y) / scale);
        }

        public static Point2d addPt2d(string x, string y, double scale = 1)
        {
            return addPt2d(convToDouble(x) / 10 / scale, convToDouble(y) / 10 / scale);
        }

        public static Point2d addPt2d(string x, double y, double scale = 1)
        {
            return addPt2d(convToDouble(x) / scale, y / scale);
        }

        public static object getPlane(FaceProxy face)
        {
            Face f = face.NativeObject;
            PartFeature pf = f.CreatedByFeature;
            object plane = null;
            if (f.SurfaceType == SurfaceTypeEnum.kCylinderSurface)
            {
                Cylinder cyl = f.Geometry as Cylinder;
                return getPlane(face, cyl.AxisVector);
            }
            //             if (pf.Type == ObjectTypeEnum.kContourFlangeFeatureObject)
            //             {
            //                 ContourFlangeFeature cffp = pf as ContourFlangeFeature;
            //                 SketchEntity se = (SketchEntity)cffp.Definition.Path[1].SketchEntity;
            //                 UnitVector vec = ((PlanarSketch)se.Parent).PlanarEntityGeometry.Normal;
            //                 return getPlane(face, vec);
            //             }
            return plane;
        }

        public static Point2d midPt(Point2d pt1, Point2d pt2, double offsetDir, double offsetN, double scale = 1, double w = 0.5)
        {
            Vector2d vec = pt1.VectorTo(pt2);
            if (scale != 1)
            {
                if (offsetN != 0) offsetN = scale * vec.Length;
                if (offsetDir != 0) offsetDir = scale * vec.Length;
            }
            vec.ScaleBy(w);
            Vector2d norm = normal(pt1, pt2); norm.Normalize(); norm.ScaleBy(offsetDir);
            pt1.TranslateBy(vec); vec.Normalize(); vec.ScaleBy(offsetN);
            pt1.TranslateBy(vec);
            pt1.TranslateBy(norm);
            return pt1;
        }

        public static Point midPt(Point pt1, Point pt2, double w = 0.5)
        {
            Point pt = pt1.Copy();
            Vector vec = pt.VectorTo(pt2);
            vec.ScaleBy(w);
            pt.TranslateBy(vec);
            return pt;
        }
        public static Point midPt(Edge e)
        {
            return midPt(e.StartVertex.Point, e.StopVertex.Point);
        }
        public static Point2d midPt(Point2d pt1, Point2d pt2, double w = 0.5)
        {
            Point2d pt = pt1.Copy();
            Vector2d vec = pt.VectorTo(pt2);
            vec.ScaleBy(w);
            pt.TranslateBy(vec);
            return pt;
        }

        public static Point centerEdge(Edge e)
        {
            Circle c = (Circle)e.Geometry;
            return c.Center;
        }

        public static bool comparePoint(Inventor.Vertex pt1, Inventor.Vertex pt2, int tolerance)
        {
            return ((int)(pt1.Point.X * tolerance) == (int)(pt2.Point.X * tolerance) &&
                (int)(pt1.Point.Y * tolerance) == (int)(pt2.Point.Y * tolerance) &&
                (int)(pt1.Point.Z * tolerance) == (int)(pt2.Point.Z * tolerance)) ? true : false;
        }

        public static bool comparePoint2d(Inventor.Vertex pt1, Inventor.Vertex pt2, int tolerance)
        {
            return ((int)(pt1.Point.X * tolerance) == (int)(pt2.Point.X * tolerance) &&
                (int)(pt1.Point.Y * tolerance) == (int)(pt2.Point.Y * tolerance)) ? true : false;
        }

        static public double len(Vertex v1, Vertex v2)
        {
            Vector vec = I.tg.CreateVector(v2.Point.X - v1.Point.X, v1.Point.Y - v2.Point.Y, v2.Point.Z - v1.Point.Z);
            double l = Math.Round(vec.Length, 1);
            if (l % 5 != 0)
            {
                l = l / 5;
                l = (int)l; l = (l + 1) * 5;
            }
            return l;
        }

        static public Vector getVector(Edge e, Point o)
        {
            if (e.GeometryType != CurveTypeEnum.kLineSegmentCurve) return null;
            var l = e.Geometry as LineSegment;
            var ep = u.eq(l.StartPoint, o) ? l.EndPoint : l.StartPoint;
            return o.VectorTo(ep);
        }

        static public Vector getVector(Edge e)
        {
            return e.StartVertex.Point.VectorTo(e.StopVertex.Point);
        }

        static public bool check(Edge e1, Edge e2, Func<Vector, Vector, bool> f)
        {
            Vector v1 = getVector(e1), v2 = getVector(e2);
            return f(v1, v2);
        }

        static public double len(Edge e)
        {
            return e.StartVertex.Point.VectorTo(e.StopVertex.Point).Length;
        }

        static public void addClientLine(PartDocument doc, string name, Asset color, Point pt1, Point pt2)
        {
            PartComponentDefinition compDef = doc.ComponentDefinition; ClientGraphics gs;
            try
            {
                ICollection<ClientGraphics> col = compDef.ClientGraphicsCollection as ICollection<ClientGraphics>;
                gs = compDef.ClientGraphicsCollection[name];
            }
            catch (Exception)
            {
                gs = compDef.ClientGraphicsCollection.Add(name);
            }
            //             if (compDef.ClientGraphicsCollection[name] != null)
            //             {
            //                 compDef.ClientGraphicsCollection[name].Delete();
            //                 I.app.ActiveView.Update();
            //             }                       
            GraphicsDataSets data = doc.GraphicsDataSetsCollection.Add(doc.GraphicsDataSetsCollection.Count.ToString());
            GraphicsNode node = gs.AddNode(gs.Count);
            LineGraphics line = node.AddLineGraphics();
            node.Appearance = color;
            GraphicsCoordinateSet coord = data.CreateCoordinateSet(data.Count);
            coord.Add(1, pt1);
            coord.Add(2, pt2);
            line.CoordinateSet = coord;
            //LineStripGraphics lstrip = gs.AddNode(2).AddLineStripGraphics();
            I.app.ActiveView.Update();
        }

        static public void addClientLine(FlatPattern fp, string name, Asset color, Point pt1, Point pt2)
        {
            ClientGraphics gs;
            try
            {
                ICollection<ClientGraphics> col = fp.ClientGraphicsCollection as ICollection<ClientGraphics>;
                gs = fp.ClientGraphicsCollection[name];
            }
            catch (Exception)
            {
                gs = fp.ClientGraphicsCollection.Add(name);
            }
            //             if (compDef.ClientGraphicsCollection[name] != null)
            //             {
            //                 compDef.ClientGraphicsCollection[name].Delete();
            //                 I.app.ActiveView.Update();
            //             }
            PartDocument doc = fp.Document as PartDocument;
            GraphicsDataSets data = doc.GraphicsDataSetsCollection.Add(doc.GraphicsDataSetsCollection.Count.ToString());
            GraphicsNode node = gs.AddNode(gs.Count);
            LineGraphics line = node.AddLineGraphics();
            node.Appearance = color;
            GraphicsCoordinateSet coord = data.CreateCoordinateSet(data.Count);
            coord.Add(1, pt1);
            coord.Add(2, pt2);
            line.CoordinateSet = coord;
            //LineStripGraphics lstrip = gs.AddNode(2).AddLineStripGraphics();
            I.app.ActiveView.Update();
        }

        static public XMLDoc getXMLOFD()
        {
            string p = I.curProjPath();
            var sd = file.subDir(p, new string[] { "OldVersions" }, System.IO.SearchOption.AllDirectories);
            XMLDoc xml = u.XOFD(custom: sd);
            return xml;
        }

        static public string OFD(string iniDir, string filter = "Inventor Files (*.iam;*.ipt)|*.iam;*.ipt|All Files (*.*)|*.*", 
            bool multi = false, string fn = "")
        {
            FileDialog fd;
            I.app.CreateFileDialog(out fd);
            fd.InitialDirectory = iniDir;
            fd.Filter = filter;
            fd.FilterIndex = 1;
            fd.MultiSelectEnabled = multi;
            fd.CancelError = true;
            if (fn != "") fd.FileName = fn;
            //if (filenames != null)
            //{
            //    string fnstr = "\"" + String.Join("\" \"", filenames) + "\"";
            //    fd.FileName = fnstr;
            //}
            try
            {
                fd.ShowOpen();
            }
            catch (Exception ex)
            {
                return "";
            }
            //if (!fd.CancelError) return "";
            return fd.FileName;
        }

        public static string GetLastOpenSaveFile(string extention)
        {
            Microsoft.Win32.RegistryKey regKey = Microsoft.Win32.Registry.CurrentUser;
            string lastUsedFolder = string.Empty;
            regKey = regKey.OpenSubKey("Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\ComDlg32\\OpenSavePidlMRU");

            if (string.IsNullOrEmpty(extention))
                extention = "html";

            Microsoft.Win32.RegistryKey myKey = regKey.OpenSubKey(extention);

            if (myKey == null && regKey.GetSubKeyNames().Length > 0)
                myKey = regKey.OpenSubKey(regKey.GetSubKeyNames()[regKey.GetSubKeyNames().Length - 2]);

            if (myKey != null)
            {
                string[] names = myKey.GetValueNames();
                if (names != null && names.Length > 0)
                {
                    byte[] result = myKey.GetValue(names[names.Length - 1]) as byte[];
//                     for (int i = 0; i < result.Length; ++i)
//                     {
//                         lastUsedFolder = lastUsedFolder + string.Format("{0}", result[i]);
//                     }
                    //lastUsedFolder = System.Text.Encoding.UTF8.GetString(result);
                    lastUsedFolder = unmanaged.GetPathFromPIDL(result);
                }
            }
            return lastUsedFolder;
        }

        static public string WOFD(string iniDir = "", string filter = "XML Files|*.xml", string [] custom = null, bool restore = true)
        {
            System.Windows.Forms.OpenFileDialog ofd = new System.Windows.Forms.OpenFileDialog();
            System.Windows.Forms.FileDialogCustomPlacesCollection col = ofd.CustomPlaces;  
            if (custom != null)
            {
                foreach (var item in custom)
                {
                    col.Add(item);
                }
            }
            if (!isNull(iniDir))
                ofd.InitialDirectory = iniDir;
            ofd.RestoreDirectory = restore;
            ofd.Filter = filter;
//             string ffn = GetLastOpenSaveFile("xml");
//             if (!isNull(ffn))
//             {
//                 ofd.InitialDirectory = file.p(ffn);
//                 ofd.FileName = file.nameWithExt(ffn);
//             }
            System.Windows.Forms.DialogResult dr;
            dr = ofd.ShowDialog();
            if (dr != System.Windows.Forms.DialogResult.OK) return "";
//             try
//             {
//                dr = ofd.ShowDialog();
//             }
//             catch (Exception)
//             {
//                 return "";
//             }
            return ofd.FileName;
        }

        static public XMLDoc XOFD(string iniDir = "", string rootName = "head", string [] custom = null)
        {
            string name = WOFD(iniDir, custom: custom);
            if (name == "") return null;
            XMLDoc xdoc = new XMLDoc(name, rootName);
            return xdoc;
        }

        static public Transaction transAct(object doc, string name)
        {
            Transaction tr = I.app.TransactionManager.StartTransaction((_Document)doc, name);
            return tr;
        }

        static public void copySS(string n, DrawingDocument m_Drw, string drwTemplate)
        {
            if (drwTemplate == null)
            {
                drwTemplate = I.app.DesignProjectManager.ActiveDesignProject.TemplatesPath;
                XML name = new XML(I.p() + @"\TemplatePath.xml");
                System.Collections.Generic.List<string> strs = new System.Collections.Generic.List<string>();
                strs = name.ReadXML("Template", "DrawingTemplate");
                drwTemplate = drwTemplate + strs[0];
            }
            Inventor.DrawingDocument tmpDoc = (Inventor.DrawingDocument)I.app.Documents.Open(drwTemplate, false);
            SketchedSymbolDefinition def = tmpDoc.SketchedSymbolDefinitions[n].CopyTo((Inventor._DrawingDocument)m_Drw, true);
            tmpDoc.Close(true);
        }

        static public Asset createColor(PartDocument doc, string intName, string name, byte red, byte green, byte blue)
        {
            Asset color;
            try
            {
                color = doc.Assets.Add(AssetTypeEnum.kAssetTypeAppearance, "Generic", intName, name);
                ((ColorAssetValue)color["generic_diffuse"]).Value = I.app.TransientObjects.CreateColor(red, green, blue);
            }
            catch (Exception)
            {
                color = doc.Assets[name];
            }
            return color;
        }

        static public UnitVector2d createUnitVector2d(double x, double y)
        {
            return I.tg.CreateUnitVector2d(x, y);
        }

        static public UnitVector createUnitVector(double x, double y, double z)
        {
            return I.tg.CreateUnitVector(x, y, z);
        }

        static public Vector2d createVector2d(double x, double y)
        {
            return I.tg.CreateVector2d(x, y);
        }

        static public Vector createVector(double x, double y, double z)
        {
            return I.tg.CreateVector(x, y, z);
        }

        static public Point2d createPoint2d(double x, double y)
        {
            return I.tg.CreatePoint2d(x, y);
        }

        static public Point createPoint(double x, double y, double z)
        {
            return I.tg.CreatePoint(x, y, z);
        }

        static public void getCurvesToArray(DrawingView dv, DrawingCurve dc, out List<DrawingCurve> dcX, out List<DrawingCurve> dcY)
        {
            dcX = new List<DrawingCurve>(); dcY = new List<DrawingCurve>();
            dcX.Add(dc); dcY.Add(dc);
            Point2d pt = dc.CenterPoint;
            Circle2d originCircle = dc.Segments[1].Geometry as Circle2d;
            foreach (DrawingCurve item in dv.DrawingCurves)
            {
                if (item.ProjectedCurveType != Curve2dTypeEnum.kCircleCurve2d) continue;
                Circle2d circ = item.Segments[1].Geometry as Circle2d;
                if (!eq(originCircle.Radius, circ.Radius)) continue;
                if (eq(pt.X, item.CenterPoint.X) && !eq(pt.Y, item.CenterPoint.Y))
                {
                    dcY.Add(item);
                }
                if (!eq(pt.X, item.CenterPoint.X) && eq(pt.Y, item.CenterPoint.Y))
                {
                    dcX.Add(item);
                }
            }
        }

        static public Centerline addCenterLine(DrawingView dv, DrawingCurve a, DrawingCurve b)
        {
            Sheet sh = dv.Parent; GeometryIntent i1 = sh.CreateGeometryIntent(a, a.MidPoint), i2 = sh.CreateGeometryIntent(b, b.MidPoint);
            ObjectCollection col = I.COC(); col.Add(i1); col.Add(i2);
            if (checkCL(dv, i1, i2)) return null;  
            var cl = sh.Centerlines.Add(col);
            return cl;
        }

        static public double getPoz(GeometryIntent i1, GeometryIntent i2, UnitVector2d u)
        {
            return Math.Abs(scalar(i1.PointOnSheet, u) - scalar(i2.PointOnSheet, u));
        }

        static public bool check(DrawingView dv, GeometryIntent pt1, GeometryIntent pt2, DimensionTypeEnum t)
        {
            if (t == DimensionTypeEnum.kAlignedDimensionType) return false;
            UnitVector2d u = t == DimensionTypeEnum.kVerticalDimensionType ? I.CUV2d(0,1): I.CUV2d(1,0);
            var ie = getDims(dv, t); 
            double poz = getPoz(pt1, pt2, u);
            var d = ie.FirstOrDefault(e => eq(poz, getPoz(e.IntentOne, e.IntentTwo, u)));
            return d != null;
        }

        static public bool check(DrawingView dv, DimensionTypeEnum t, GeometryIntent i1, GeometryIntent i2)
        {
            var dims = getDims(dv, t);
            var d = dims.FirstOrDefault(f => eq(i1.PointOnSheet, f.IntentOne.PointOnSheet) && eq(i2.PointOnSheet, f.IntentTwo.PointOnSheet));
            return d != null;
        }

        static public bool checkCL(DrawingView dv, GeometryIntent i1, GeometryIntent i2)
        {
            Sheet sh = dv.Parent as Sheet;
            var cls = gets<Centerline>(sh.Centerlines, f => true);
            var d = cls.FirstOrDefault(f => eq(i1.PointOnSheet, (f.FitPoints[1] as GeometryIntent).PointOnSheet) && eq(i2.PointOnSheet, (f.FitPoints[2] as GeometryIntent).PointOnSheet));
            return d != null;
        }

        static public IEnumerable<LinearGeneralDimension> getDims(DrawingView dv, DimensionTypeEnum t)
        {
            Sheet sh = dv.Parent as Sheet;
            return gets<LinearGeneralDimension>(sh.DrawingDimensions.GeneralDimensions, e => e.DimensionType == t);
        }
        static public IEnumerable<LinearGeneralDimension> getDims(Sheet sh)
        {
            return gets<LinearGeneralDimension>(sh.DrawingDimensions.GeneralDimensions, fi => true);
        }

        static public IEnumerable<LinearGeneralDimension> getDims(DrawingView dv)
        {
            Box2d dvBox = I.Box(dv);
            dvBox.Expand(0.1);
            foreach (var item in getDims(dv.Parent))
            {
                GeometryIntent gi1 = item.IntentOne, gi2 = null;
                if (gi1 != null)
                {
                    Centerline cl = gi1.Geometry as Centerline;
                    Centermark cm = gi1.Geometry as Centermark;
                    DrawingCurve dc = gi1.Geometry as DrawingCurve;
                    if (cl != null) cm = cl.FitPoints[1] as Centermark;
                    if (cm != null) gi2 = cm.AttachedEntity as GeometryIntent;
                    if (gi2 != null) dc = gi2.Geometry as DrawingCurve;
                    if (dc != null && !(dc.Parent.Equals(dv)))
                        continue;
                }
                yield return item;
            }
        }

        static public T addDim<T>(DrawingView dv, DrawingCurve ent1, DrawingCurve ent2,
            double offset = 1.5, DimensionTypeEnum align = DimensionTypeEnum.kAlignedDimensionType, double min = 0, double direct = 1, DimensionStyle st = null) where T : class
        {
            Sheet sh = dv.Parent as Sheet; GeometryIntent i1, i2; Point2d pt;
            if (ent1.ModelGeometry is SketchLine && ent2.ModelGeometry is SketchLine)
            {
                i1 = sh.CreateGeometryIntent(ent1); i2 = sh.CreateGeometryIntent(ent2);
                pt = midPt(ent1.EndPoint, ent2.EndPoint);
                if (st == null) st = I.getStyleA();
                return sh.DrawingDimensions.GeneralDimensions.AddAngular(pt, i1, i2) as T;
            }
            else if (ent1.CurveType == CurveTypeEnum.kCircleCurve && ent2.CurveType == CurveTypeEnum.kCircleCurve ||
            (ent1.CurveType == CurveTypeEnum.kCircularArcCurve && ent2.CurveType == CurveTypeEnum.kCircularArcCurve))
            {
                i1 = sh.CreateGeometryIntent(ent1, PointIntentEnum.kCenterPointIntent); i2 = sh.CreateGeometryIntent(ent2, PointIntentEnum.kCenterPointIntent);
                pt = midPt(ent1.CenterPoint, ent2.CenterPoint);
                pt = align == DimensionTypeEnum.kHorizontalDimensionType ? I.tg.CreatePoint2d(pt.X, min + direct * offset) :
                    align == DimensionTypeEnum.kVerticalDimensionType ?
                        I.tg.CreatePoint2d(min + direct * offset, pt.Y) : pt;
                if (st == null) st = I.getStyleL();
                return sh.DrawingDimensions.GeneralDimensions.AddLinear(pt, i1, i2, align) as T;
            }
            else if (ent1.CurveType == CurveTypeEnum.kCircleCurve && ent2.CurveType == CurveTypeEnum.kLineSegmentCurve ||
            (ent1.CurveType == CurveTypeEnum.kCircularArcCurve && ent2.CurveType == CurveTypeEnum.kLineSegmentCurve))
            {
                i1 = sh.CreateGeometryIntent(ent1, PointIntentEnum.kCenterPointIntent); i2 = sh.CreateGeometryIntent(ent2);
                pt = midPt(ent1.CenterPoint, ent2.StartPoint);
                pt = align == DimensionTypeEnum.kHorizontalDimensionType ? I.tg.CreatePoint2d(pt.X, min + direct * offset) :
                    align == DimensionTypeEnum.kVerticalDimensionType ?
                        I.tg.CreatePoint2d(min + direct * offset, pt.Y) : pt;
                if (st == null) st = I.getStyleL();
                return sh.DrawingDimensions.GeneralDimensions.AddLinear(pt, i1, i2, align) as T;
            }
            else
            {
                i1 = sh.CreateGeometryIntent(ent1); i2 = sh.CreateGeometryIntent(ent2);
                pt = midPt(ent1.EndPoint, ent2.EndPoint);
                if (st == null) st = I.getStyleA();
                return sh.DrawingDimensions.GeneralDimensions.AddAngular(pt, i1, i2) as T;
            }
        }
        static public T getGeom<T>(DrawingCurve dc) where T : class
        {
            object o = dc.Segments[1].Geometry;
            return o is T ? o as T : null;
        }
        static public T getGeom<T>(Edge e) where T: class
        {
            object o = e.Geometry;
            return o is T ? o as T : null;
        }
        static public List<GeometryIntent> getGeomIntents(Sheet sh, DrawingCurve dc1, DrawingCurve dc2, ref Point2d tp,
            DrawingCurve rDC = null)
        {
            List<GeometryIntent> ints = new List<GeometryIntent>();
            Arc2d a1 = getGeom<Arc2d>(dc1), a2 = getGeom<Arc2d>(dc2);
            LineSegment2d l1 = getGeom<LineSegment2d>(dc1), l2 = getGeom<LineSegment2d>(dc2);
            tp = midPt(dc1.MidPoint, dc2.MidPoint);
            if (rDC != null) tp = midPt(rDC.StartPoint, rDC.EndPoint);
            if (a1 != null && a2 != null)
            {
                ints.Add(getGeomIntent(sh, dc1, PointIntentEnum.kMidPointIntent));
                ints.Add(getGeomIntent(sh, dc2, PointIntentEnum.kMidPointIntent));
                
                I.getStyle("Текст*");
            }
            else if (l1 != null && l2 != null)
            {
                if (rDC != null)
                {
                    Point2d min = getVarDist(dc1, rDC, (a, b) => a <= b);
                    ints.Add(getGeomIntent(sh, dc1, PointIntentEnum.kEndPointIntent, min));
                    min = getVarDist(dc2, rDC, (a, b) => a <= b);
                    ints.Add(getGeomIntent(sh, dc2, PointIntentEnum.kEndPointIntent, min));
                }
                else if (l1.Direction.IsParallelTo(l2.Direction, 0.001))
                {
                    ints.Add(sh.CreateGeometryIntent(dc1));
                    ints.Add(sh.CreateGeometryIntent(dc2));
                }
            }
            else if (l1 != null && a2 != null)
            {
                Point2d max = getVarDist(dc2, dc1, (a,b) => a >= b);
                ints.Add(sh.CreateGeometryIntent(dc1));
                ints.Add(getGeomIntent(sh, dc2, PointIntentEnum.kEndPointIntent, max));
            }
            else if (a1 != null && l2 != null)
            {
                Point2d max = getVarDist(dc1, dc2, (a, b) => a >= b);
                ints.Add(getGeomIntent(sh, dc1, PointIntentEnum.kEndPointIntent, max));
                ints.Add(sh.CreateGeometryIntent(dc2));
            }
            return ints;
        }
        static public Point2d getVarDist(DrawingCurve ent1, DrawingCurve ent2, Func<double, double ,bool> f)
        {
            double d1 = ent1.StartPoint.DistanceTo(ent2.StartPoint),
                d2 = ent1.EndPoint.DistanceTo(ent2.StartPoint);
            return f(d1, d2) ? ent1.StartPoint : ent1.EndPoint;
        }
        static public DrawingCurveSegment getSegm(DrawingCurve dc, Vector2d v, Point2d pt, Func<double, double,bool> f)
        {
            if (dc.Segments.Count == 1) return dc.Segments[1];
            Box2d b = null;
            double sum = 0, d = 0;
            DrawingCurveSegment s = null;
            foreach (DrawingCurveSegment seg in dc.Segments)
            {
                var l = seg.Geometry as LineSegment2d;
                var a = seg.Geometry as Arc2d;
                if (l != null) b = l.Evaluator.RangeBox;
                if (a != null) b = a.Evaluator.RangeBox;
                var mp = u.midPt(b.MinPoint, b.MaxPoint);
                d = Math.Abs(u.dotProduct(v, pt.VectorTo(mp)));
                if (f(d,sum)) { s = seg; sum = d; }
            }
            return s;
        }
        static public DrawingCurve findMinDC(DrawingView dv, Point2d pt, Vector2d v, Func<DrawingCurve, bool> f,
            Func<double, double, bool> cond, out double l)
        {
            DrawingCurve rdc = null;
            l = 0;
            double tol = 0.2;
            Vector2d n = I.CV2d(v); u.normal(n, null);
            Point2d pt2 = I.CP2d(pt, n);
            foreach (DrawingCurve dc in u.gets<DrawingCurve>(dv.DrawingCurves, fi => f(fi)))
            {
                var vec = dc.StartPoint.VectorTo(dc.EndPoint);
                if (vec.Length < tol) continue;
                if (!vec.IsPerpendicularTo(v)) continue;
                double d = normal(pt, pt2, dc.StartPoint).Length;
                if (d < tol) continue;
                if (rdc == null) { l = d; rdc = dc; continue; }
                if (cond(d, l)) { l = d; rdc = dc; }
            }
            return rdc;
        }
        static public DrawingView findDV(object ob)
        {
            Centerline cl = ob as Centerline;
            if (cl != null)
            {
                foreach (DrawingView item in cl.Parent.DrawingViews)
                {
                    Box2d b = I.Box(item);
                    Point2d mp = midPt(cl.StartPoint, cl.EndPoint);
                    if (b.Contains(mp)) return item;
                }
            }
            return null;
        }
        static public DrawingCurve getDC(DrawingView dv, Edge e, Func<DrawingCurve, bool> f)
        {
            foreach (DrawingCurve dc in dv.DrawingCurves)
            {
                if (!f(dc)) continue;
                var me = dc.ModelGeometry as Edge;
                if (me == null) continue;
                if (me.Equals(e)) return dc;
            }
            return null;
        }
        static public GeometryIntent getGeomIntent(Sheet sh, DrawingCurve dc, PointIntentEnum pie, Point2d pt = null)
        {
            GeometryIntent i;
            if (pt != null) pie = getPIE(dc, pt);
            return sh.CreateGeometryIntent(dc, pie);
        }
        static public object findCL(Centermark cm, Vector2d dir)
        {
            foreach (Centerline cl in cm.Centerlines)
            {
                var v = cl.StartPoint.VectorTo(cl.EndPoint);
                if (v.IsPerpendicularTo(dir,0.01)) return cl;
            }
            return cm;
        }
        static public GeometryIntent getGeomIntent(Sheet sh, object cm)
        {
            GeometryIntent gi;
            return (cm is Centermark) ? sh.CreateGeometryIntent(cm, PointIntentEnum.kCenterPointIntent):
                sh.CreateGeometryIntent(cm);
        }
        static public LinearGeneralDimension addDim(DrawingView dv, Centermark cm1, Centermark cm2)
        {
            var sh = dv.Parent;
            var v = cm1.Position.VectorTo(cm2.Position);
            var gi1 = getGeomIntent(sh, findCL(cm1, v));
            var gi2 = getGeomIntent(sh, findCL(cm2, v));
            Point2d pt = midPt(cm1.Position, cm2.Position);
            LinearGeneralDimension dim = sh.DrawingDimensions.GeneralDimensions.AddLinear(pt, gi1, gi2);
            return dim;
        }
        static public LinearGeneralDimension addDim(DrawingView dv, GeometryIntent gi, DrawingCurve dc)
        {
            if (dc.CurveType != CurveTypeEnum.kLineSegmentCurve) return null;
            var sh = dv.Parent;
            var pt = ptIntent(gi);
            var v = normal(dc.StartPoint, dc.EndPoint, ptIntent(gi));
            var gi2 = sh.CreateGeometryIntent(dc);
            v.ScaleBy(-0.5);
            pt.TranslateBy(v);
            var dim = sh.DrawingDimensions.GeneralDimensions.AddLinear(pt, gi, gi2);
            return dim;
        }
        static public LinearGeneralDimension addDim(DrawingView dv, GeometryIntent gi1, GeometryIntent gi2, 
            Vector2d v, double min = 1.2)
        {
            int o = octant(v);
            //DimensionTypeEnum align = DimensionTypeEnum.kAlignedDimensionType;
            DimensionTypeEnum align = o == 1 || o == 4 ? DimensionTypeEnum.kVerticalDimensionType :
                DimensionTypeEnum.kHorizontalDimensionType;
            var b = new BoxBase(); b.set(dv);
            var pt = b.getPoint(v, o, min);
            //sh.DrawingNotes.GeneralNotes.AddFitted(dc1.StartPoint, "s1");
            var dim = dv.Parent.DrawingDimensions.GeneralDimensions.AddLinear(pt, gi1, gi2, align);
            if (containsDim(dv.Parent, dim)) { dim.Delete(); dim = null; }
            disjoint(dv, dim);
            if (v.IsParallelTo(I.CV2d(0,1),0.01))
            {
                if (dim == null) return null;
                var dp1 = dotProduct(v, dim.Text.RangeBox.MaxPoint, dv.Center);
                var dp2 = dotProduct(v, dv.Label.RangeBox.MinPoint, dv.Center);
                if (dp1 >= dp2)
                {

                    setDist(v, dp1 - dp2 + 4);
                    var tp = I.CP2d(dv.Label.Position, v);
                    dv.Label.Position = tp;
                }
            }          
            return dim;
        }
        static public LinearGeneralDimension addDim(DrawingView dv, GeometryIntent gi1, GeometryIntent gi2,
            Vector2d v, Point2d pt)
        {
            var vec = I.CV2d(1, 0);
            //DimensionTypeEnum align = DimensionTypeEnum.kAlignedDimensionType;
            DimensionTypeEnum align = vec.IsParallelTo(v, 0.01) ? DimensionTypeEnum.kVerticalDimensionType :
                vec.IsPerpendicularTo(v, 0.01) ? DimensionTypeEnum.kHorizontalDimensionType :
                DimensionTypeEnum.kAlignedDimensionType;
            //if (gi1.IntentType == IntentTypeEnum.kNoPointIntent || gi2.IntentType == IntentTypeEnum.kNoPointIntent)
            //    align = DimensionTypeEnum.kAlignedDimensionType;
            if (align == DimensionTypeEnum.kAlignedDimensionType)
            {
                //Point2d pt1 = getPointInt(gi1, v), pt2 = getPointInt(gi2, v);
                //if (dotProduct(pt1.VectorTo(pt2), v) > 0)
                //    pt = pt2;
                //else pt = pt1;
            }
            //align = DimensionTypeEnum.kAlignedDimensionType;
            LinearGeneralDimension dim = null;
            try
            {
                dim = gi1.Geometry is Centermark ? dv.Parent.DrawingDimensions.GeneralDimensions.AddLinear(pt, gi2, gi1, align) :
                dv.Parent.DrawingDimensions.GeneralDimensions.AddLinear(pt, gi1, gi2, align);
            }
            catch (Exception)
            {
            }
            //if (align == DimensionTypeEnum.kAlignedDimensionType) dim.CenterText();
            if (dim != null) moveDimText(dim);
            return dim;
        }
        static public bool checkDimBend(GeometryIntent gi)
        {
            DrawingCurve dc = gi.Geometry as DrawingCurve;
            Centerline cl = gi.Geometry as Centerline;
            Centermark cm = gi.Geometry as Centermark;
            if (dc != null && (dc.EdgeType == DrawingEdgeTypeEnum.kBendDownEdge ||
                dc.EdgeType == DrawingEdgeTypeEnum.kBendUpEdge)) return true;
            return false;
        }
        static public Point2d getPointInt(GeometryIntent gi, Vector2d vec)
        {
            if (gi.IntentType == IntentTypeEnum.kPoint2dIntent) return gi.PointOnSheet;
            var cl = gi.Geometry as Centerline;
            var dc = gi.Geometry as DrawingCurve;
            Point2d pt = null;
            Vector2d v;
            if (cl != null)
            {
                v = cl.StartPoint.VectorTo(cl.EndPoint);
                pt = dotProduct(v, vec) < 0 ? cl.StartPoint : cl.EndPoint;
            }
            else if (dc != null)
            {
                v = dc.StartPoint.VectorTo(dc.EndPoint);
                pt = dotProduct(v, vec) < 0 ? dc.StartPoint : dc.EndPoint;
            }
            return pt;
        }
        static public double getParam(DrawingCurve dc, Vector2d v, Point2d pt, Func<double, double ,bool> f)
        {
            var dcs = getSegm(dc, v, pt, (d, sum) => d >= sum);
            return getParamAtPoint(dc.Evaluator2D, dc.EndPoint);
        }
        static public PointIntentEnum getPIE(DrawingCurve dc, Vector2d v, Point2d pt, Func<double,double,bool> f)
        {
            if (dc.Segments.Count == 1) return PointIntentEnum.kPlanarFaceCenterPointIntent;
            var vec1 = pt.VectorTo(dc.StartPoint);
            var vec2 = pt.VectorTo(dc.EndPoint);
            var d1 = dotProduct(v, vec1);
            var d2 = dotProduct(v, vec2);
            return f(d1, d2) ? PointIntentEnum.kStartPointIntent : PointIntentEnum.kEndPointIntent;
        }
        static public void disjoint(DrawingView dv, LinearGeneralDimension dim)
        {
            if (dim == null) return;
            var db = new DimBox(dim, dv.Position);
            List<DimBox> dims = new List<DimBox>();
            foreach (var item in getDims(dv.Parent))
            {
                var db1 = new DimBox(item, dv.Position);
                if (db.check(db1)) dims.Add(db1);
            }
            var ie = dims.OrderByDescending(el => el.d);
            if (ie.Count() == 0) return;
            db.move(ie.ElementAt(0), dv);
        }
        static public bool containsDim(Sheet sh, LinearGeneralDimension dim)
        {
            if (dim == null) return false;
            UnitVector2d dir1 = I.CUV2d(), dir2 = I.CUV2d();
            double d = dimLenght(dim, ref dir1);
            foreach (var item in sh.DrawingDimensions.GeneralDimensions)
            {
                if (dim.Equals(item)) continue;
                var lgd = item as LinearGeneralDimension;
                if (lgd == null) continue;
                double d1 = dimLenght(lgd, ref dir2);
                if (eq(d, d1) && dir1.IsParallelTo(dir2)) return true;
            }
            return false;
        }
        static public double dimLenght(LinearGeneralDimension dim, ref UnitVector2d dir)
        {
            if (dim == null) return 0;
            var l = dim.DimensionLine as LineSegment2d;
            if (l == null) return 0;
            dir = l.Direction;
            var v = l.StartPoint.VectorTo(l.EndPoint);
            return v.Length;
        }
        static public LinearGeneralDimension addDim(DrawingView dv, DrawingCurve dc1, DrawingCurve dc2, DimensionTypeEnum align,
            Vector2d v, string stStr, bool move, Vector2d v1, double min = 1.2, DrawingCurve rDC = null)
        {
            I.getStyle(stStr);
            var sh = dv.Parent;
            Point2d pt = null;
            var gis = getGeomIntents(sh, dc1, dc2, ref pt, rDC);
            //DrawingCurve ent1 = dc1.ProjectedCurveType == Curve2dTypeEnum.kCircularArcCurve2d ? dc1 : dc2;
            Vector2d vec = I.CV2d(dv.Position, pt), vec1 = I.CV2d(dv.Position, pt);
            vec1.SubtractVector(v);
            if (vec1.Length < vec.Length) setDist(v, min);
            else setDist(v, -min);
            pt.TranslateBy(v);
            if (v1 != null)
            {
                var b = new BoxBase(); b.set(dv);
                int o = octant(v1);
                var pt1 = b.getPoint(v, o, min);
                v.Normalize();
                var tmp = pt.VectorTo(pt1);
                v.ScaleBy(tmp.DotProduct(v));
                pt.TranslateBy(v);
            }
            Point2d m = midPt(dc1.MidPoint, dc2.MidPoint);
            if (check(m, pt))
            {
                setDist(v, -0.35); pt.TranslateBy(v);
            }
            if (gis.Count == 0) return null;
            if (gis[0] == null || gis[1] == null) return null;
            try
            {
                if (dv.GeneralDimensionType != GeneralDimensionTypeEnum.kProjectedGeneralDimension)
                    dv.GeneralDimensionType = GeneralDimensionTypeEnum.kProjectedGeneralDimension;
                LinearGeneralDimension dim = sh.DrawingDimensions.GeneralDimensions.AddLinear(pt, gis[0], gis[1], align, true, I.dimStyle);
                if (move) moveDimText(dim);
                return dim;
            }
            catch (Exception)
            {
                return null;
            }
        }
        static public Point2d getDimPoint(LinearGeneralDimension dim, bool start)
        {
            var s = dim.DimensionLine;
            LineSegment2d s2 = s as LineSegment2d;
            if (s2 != null)
                return start ? s2.StartPoint : s2.EndPoint;
            LineSegment s3 = s as LineSegment;
            Point pt = null;
            if (s3 != null)
                pt = start ? s3.StartPoint : s3.EndPoint;
            return I.tg.CreatePoint2d(pt.X, pt.Y);
        }
        static public void moveDimText(LinearGeneralDimension dim)
        {
            Point2d tp = dim.Text.Origin, pt1 = getDimPoint(dim, true), 
                pt2 = getDimPoint(dim, false);
            Vector2d v = pt1.VectorTo(pt2), vec1 = dim.Text.RangeBox.MinPoint.VectorTo(dim.Text.RangeBox.MaxPoint);
            Vector2d vb = pt1.VectorTo(tp);
            var dir = I.CV2d(v);
            dir.Normalize();
            double d = Math.Abs(vec1.DotProduct(dir)*1.1);
            if (v.Length - d > 0)
            {
                //if (!check(m, tp))
                //{
                    dim.CenterText();
                    return;
                //}
            }
            d = vec1.Length * 1;
            d = v.Length * 0.6 + d * 0.5+0.3;
            setDist(v,-1);
            Vector2d v1 = dimVec(vb, v);
            setDist(v, -1);
            Vector2d v2 = dimVec(vb, v);
            d = v1.Length < v2.Length ? d : -d;
            setDist(v,d);
            Point2d pt = dim.Text.Origin;
            pt.TranslateBy(v);
            dim.Text.Origin = pt;
        }
        static public void alignDim(LinearGeneralDimension dim)
        {
            Point2d tp = dim.Text.Origin;
            LineSegment2d l = dim.DimensionLine as LineSegment2d;
            Point2d mp = u.midPt(l.StartPoint, l.EndPoint);
            Vector2d n = u.normal(l, tp);
            mp.TranslateBy(n);
            if (n.Length > 0.4)
            {
                dim.Text.Origin = mp;
                var v = mp.VectorTo(tp);
                Point2d np = dim.Text.Origin;
                v.ScaleBy(1);
                np.TranslateBy(v);
                dim.Text.Origin = np;
            }
        }
        static public bool check(Point2d pt1, Point2d pt2)
        {
            Vector2d v = pt1.VectorTo(pt2);
            var o = octant(v);
            if (o >= 2 && o <= 5) return true;
            return false;
        }
        static public Vector2d dimVec(Vector2d v1, Vector2d v)
        {
            v1.SubtractVector(v);
            return v1;
        }
        static public PointIntentEnum getPIE(DrawingCurve c, Point2d pt)
        {
            return u.eq(pt, c.StartPoint) ? PointIntentEnum.kStartPointIntent : PointIntentEnum.kEndPointIntent;
        }

        static public T addDim<T>(DrawingView dv, DrawingCurve dc, double offset = 1.5,
            DimensionTypeEnum align = DimensionTypeEnum.kAlignedDimensionType, DimensionStyle st = null) where T : class
        {
            Sheet sh = dv.Parent as Sheet; GeometryIntent i1, i2;
            if (st == null) st = I.getStyleL();
            if (dc.ProjectedCurveType == Curve2dTypeEnum.kLineSegmentCurve2d)
            {
                if (st == null) st = I.getStyleL();
                i1 = sh.CreateGeometryIntent(dc, dc.StartPoint); i2 = sh.CreateGeometryIntent(dc, dc.EndPoint);
                Point2d pt = midPt(dc.StartPoint, dc.EndPoint, offset, 0);
                return sh.DrawingDimensions.GeneralDimensions.AddLinear(pt, i1, i2, align, DimensionStyle: st) as T;
            }
            else if (dc.ProjectedCurveType == Curve2dTypeEnum.kCircularArcCurve2d)
            {
                if (st == null) st = I.getStyleR();
                i1 = sh.CreateGeometryIntent(dc);
                Point2d pt = midPt(dc.StartPoint, dc.CenterPoint, offset, 0);
                return sh.DrawingDimensions.GeneralDimensions.AddRadius(pt, i1, DimensionStyle: st) as T;
            }
            return null;
        }

        static public string getDocName(Document doc)
        {
            string docName = file.name(doc.FullFileName);
            string tmp = u.getPropValue(doc, "Description");
            if (tmp != "") docName = tmp;
            return docName;
        }

        static public LinearGeneralDimension addDim(DrawingView dv, object ent1, object ent2, object pi1, object pi2, Point2d mpt1, Point2d mpt2, double offset = 1.5,
            DimensionTypeEnum align = DimensionTypeEnum.kAlignedDimensionType)
        {
            Sheet sh = dv.Parent as Sheet;
            GeometryIntent i1, i2;
            if (pi1 == null) i1 = sh.CreateGeometryIntent(ent1);
            else i1 = sh.CreateGeometryIntent(ent1, pi1);
            if (pi2 == null) i2 = sh.CreateGeometryIntent(ent2);
            else i2 = sh.CreateGeometryIntent(ent2, pi2);
            Point2d pt = midPt(mpt1, mpt2, offset, 0);
            return sh.DrawingDimensions.GeneralDimensions.AddLinear(pt, i1, i2, align);
        }

        static public LinearGeneralDimension addDim(DrawingView dv, DrawingCurve ent1, DrawingCurve ent2, Point2d pt1, Point2d pt2, double offset = 1.5, double offsetDir = 0, DimensionStyle st = null)
        {
            DimensionTypeEnum align = DimensionTypeEnum.kAlignedDimensionType;
            Sheet sh = dv.Parent as Sheet;
            GeometryIntent i1, i2;
            if (st == null) st = I.getStyleL();
            PointIntentEnum pi1 = getPIE(ent1, pt1, pt2), pi2 = getPIE(ent2, pt2, pt1);
            i1 = sh.CreateGeometryIntent(ent1, pi1);
            i2 = sh.CreateGeometryIntent(ent2, pi2);
            var t = u.getTangent(ent2.Evaluator2D);
            if (Math.Abs(t[0]) < 0.001) align = DimensionTypeEnum.kHorizontalDimensionType;
            else if (Math.Abs(t[1]) < 0.001) align = DimensionTypeEnum.kVerticalDimensionType;
            Point2d s = I.CP2d(pt1), e = I.CP2d(pt2);
            Point2d pt = midPt(s, e, offsetDir, offset);
            return sh.DrawingDimensions.GeneralDimensions.AddLinear(pt, i1, i2, align, DimensionStyle: st);
        }

        static public T addDim<T>(DrawingView dv, Tuple<DrawingCurve, DrawingCurve, PointIntentEnum, PointIntentEnum, Point2d, Point2d> tup, double offset = 1.5, double offsetDir = 0,
            DimensionStyle st = null) where T : class
        {
            DimensionTypeEnum align = DimensionTypeEnum.kAlignedDimensionType;
            Sheet sh = dv.Parent as Sheet;
            GeometryIntent i1, i2;
            if (st == null) st = I.getStyleL();
            i1 = sh.CreateGeometryIntent(tup.Item1, tup.Item3);
            i2 = sh.CreateGeometryIntent(tup.Item2, tup.Item4);
            var t = u.getTangent(tup.Item2.Evaluator2D);
            if (Math.Abs(t[0]) < 0.001) align = DimensionTypeEnum.kHorizontalDimensionType;
            else if (Math.Abs(t[1]) < 0.001) align = DimensionTypeEnum.kVerticalDimensionType;
            //Point2d s = I.CP2d(tup.Item5), e = I.CP2d(tup.Item6);
            //Point2d pt = midPt(s, e, offsetDir, offset);
            Point2d pt = tup.Item5;
            if (u.check(dv, align, i1, i2)) return null;
            return sh.DrawingDimensions.GeneralDimensions.AddLinear(pt, i1, i2, align, DimensionStyle: st) as T;
        }

        static public PointIntentEnum getPIE(DrawingCurve dc, Point2d pt, Point2d pt2)
        {
            Vector2d v1 = u.getTangentVec(dc.Evaluator2D, 0.1, dc.StartPoint), v2 = u.getTangentVec(dc.Evaluator2D, 0.1, dc.EndPoint);
            if (v2.IsParallelTo(v1, 0.01) || v2.IsPerpendicularTo(v1, 0.01))
                return checkPIE(dc, pt); 
            double ang = v1.AngleTo(v2);
            if (ang - degToRad(80) < 0.1)
            {
                Point2d tmppt = u.maxDist(pt2, dc);
                return checkPIE(dc, tmppt);
            }
            else
            {
                return PointIntentEnum.kMidPointIntent;
            }
        }

        static public PointIntentEnum checkPIE(DrawingCurve dc, Point2d pt)
        {
            return eq(dc.StartPoint, pt) ? PointIntentEnum.kStartPointIntent : PointIntentEnum.kEndPointIntent;
        }

        static public Point2d maxDist(Point2d pt, Arc2d a)
        {
            return a.StartPoint.DistanceTo(pt) < a.EndPoint.DistanceTo(pt) ? a.EndPoint : a.StartPoint;
        }

        static public Point2d maxDist(Point2d pt, DrawingCurve a)
        {
            return a.StartPoint.DistanceTo(pt) > a.EndPoint.DistanceTo(pt) ? a.EndPoint : a.StartPoint;
        }

        public static double strToDouble(string value)
        {
            value = value.Replace('.', ',');
            return Double.Parse(value);
        }
        public static string getDN(string[] v, string dn, int i)
        {
            int num = 1;
            if (v.Length == 2)
            {
                //num = v[1].Length;
                if (v[1].Length == 2 || v[1].Length == 5) num = 2;
                return dn.Substring(0, dn.Length - num) + v[1];
            }
            if (v.Length == 3)
            {
                //num = v[2].Length;
                return v[1] + dn.Substring(v[1].Length, dn.Length - num - v[1].Length) + v[2];
            }
            return dn.Remove(dn.Length - 2) + i.ToString("00");
        }
        public static string getFN(string sn, string dn, int i)
        {
            var spl = sn.Split('$');
            return getDN(spl, dn, i);
        }
        public static string getFN(XElement el, string path = "")
        {
            var fn = el.Value;
            if (fn.StartsWith("\\")) fn = path + fn.Trim('\\');
            return fn;
        }
        public static bool checkFltr(string[] lst, string f)
        {
            string tmp = f; //u.getSubS(f, 5);
            foreach (var item in lst)
            {
                string e = item.ToLower();
                if (e.StartsWith("-") && u.checkRegex(tmp, e))
                {
                    return true;
                }
                else if (tmp == e) return true;
            }
            return false;
        }

        static public bool replPar(Parameters pars, string name, out Parameter p, bool f = true)
        {
            p = getParameter(pars, name);
            Parameter np;
            if (isNull(p)) return false;
            if (p.ParameterType == ParameterTypeEnum.kDerivedParameter)
            {
                return true;
            }
            string nname; 
            if (f)
            {
                nname = name + "_";
                np = pars.UserParameters.AddByExpression(nname, p.Expression, p.get_Units()) as Parameter;
            }
            else
            {
                nname = name.TrimEnd('_');
            }
            //p.Name = nname;
            foreach (Parameter item in p.Dependents)
            {
                string expr = item.Expression.Replace(name, nname);
                item.Expression = expr;
            }
            p.Delete();
            return false;
        }

        public static Document getFromPath(string dn, ref string path)
        {
            if (isNull(path)) return null;
            if (!path.StartsWith("\\")) path = "\\" + path;
            path = file.newPath(dn, path);
            if (!file.check(path)) return null;
            Document dc = I.open(path, false, false);
            return dc;
        }

        static public List<string> getCBValues(string elName)
        {
            List<string> lst = new List<string>();
            string filePath = I.p() + @"\Imate.xml";
            if (!System.IO.File.Exists(filePath)) return null;
            InvDoc.XMLDoc xml = new InvDoc.XMLDoc(I.p() + @"\Imate.xml", "head");
            foreach (var item in xml.El.Descendants(elName))
            {
                lst.Add(item.Value);
            }
            return lst;
        }
        static public List<string> set_variable(Document def)
        {
             return u.gets<Property>(def.PropertySets[4], f => f.Name.ToLower().StartsWith("type")).Select(e => e.Name).ToList();
        }
        static public int getCBIndex(string elName, string attVal)
        {
            string filePath = I.p() + @"\Imate.xml";
            if (!System.IO.File.Exists(filePath)) return -1;
            InvDoc.XMLDoc xml = new InvDoc.XMLDoc(I.p() + @"\Imate.xml", "head");
            int i = 1;
            foreach (var item in xml.El.Descendants(elName))
            {
                if (item.Attribute("Cur") != null && item.Attribute("Cur").Value == attVal)
                    return i;
                i++;
            }
            return -1;
        }
        static public void setCBValue(string elName, int ind, string val)
        {
            List<string> lst = new List<string>();
            string filePath = I.p() + @"\Imate.xml";
            if (!System.IO.File.Exists(filePath)) return;
            InvDoc.XMLDoc xml = new InvDoc.XMLDoc(I.p() + @"\Imate.xml", "head");
            int i = 1;
            foreach (var item in xml.El.Descendants(elName))
            {
                if (item.Attribute("Cur") != null && 
                    item.Attribute("Cur").Value == val) item.SetAttributeValue("Cur", null);
                if (i == ind)
                {
                    item.SetAttributeValue("Cur", val);
                }
                i++;
            }
            xml.save();
        }

        public static void removeLink(Document doc, XMLDoc xml)
        {
            string dn = doc.FullFileName;
            foreach (var item in xml.find("RemoveLink"))
            {
                string path = XMLDoc.getAttributeValue(item, "from"),
                    //pars = XMLDoc.getAttributeValue(item, "vals"),
                     part = XMLDoc.getAttributeValue(item, "in");
                ObjectCollection obj = null; obj = I.COC(obj);
                //if (!path.StartsWith("\\")) path = "\\" + path;
                //path = file.newPath(doc.FullFileName, path);
                if (path == dn) continue;
                //if (!file.check(path)) return;
                Document dc = getFromPath(dn, ref path); //I.open(path, false, false);
                if (!isNull(part)) doc = getFromPath(dn, ref part);
                if (isNull(doc, dc)) continue;
                Parameters ups = I.getParameters(doc);
                if (isNull(ups)) continue;
                List<DerivedParameterTable> rems = new List<DerivedParameterTable>();
                foreach (DerivedParameterTable dpt in ups.DerivedParameterTables)
                {
                    if (dpt.ReferencedDocumentDescriptor.FullDocumentName == dc.FullDocumentName)
                    {
                        rems.Add(dpt);
                    }
                }
                foreach (var dpt in rems)
                {
                    dpt.Delete(); 
                }
            }
        }

        public static bool checkRegex(string f, string val)
        {
            val = val.TrimStart('-');
            Regex regex = new Regex(val, RegexOptions.IgnoreCase);
            return regex.IsMatch(f);
        }

        public static string findRegexFN(string path, string val, bool noext)
        {
            val = val.TrimStart('-');
            Regex regex = new Regex(val, RegexOptions.IgnoreCase);
            var fs = file.getFiles(path, ".ipt");
            var o = u.get<string>(fs, f => regex.IsMatch(f));
            if (!u.isNull(o)) val = file.nameWithExt(o);
            else
            {
                fs = file.getFiles(path, ".iam");
                o = u.get<string>(fs, f => regex.IsMatch(f));
                if (!u.isNull(o)) val = file.nameWithExt(o);
            }
            if (noext) val = file.trimEnd(val, ".");
            return val;
        }

        public static Document getDoc(ref string p, ref string n, string tn)
        {
            string path = file.newPath(tn, p);
            path = file.p(path);
            p = path;
            if (n.StartsWith("-"))
                n = u.findRegexFN(path, n, false);
            
            return I.open(path + n, false, false);
        }

        public static void linkParameters(string dn, string docName, XMLDoc xml)
        {
            //string dn = doc.FullFileName;
            foreach (var item in xml.find("LinkParam"))
            {
                string pathFrom = XMLDoc.getAttributeValue(item, "path"), nameFrom = XMLDoc.getAttributeValue(item, "name"),
                    se = XMLDoc.getAttributeValue(item, "except");
                List<string> expt = new List<string>();
                if (!isNull(se))
                {
                    expt = getSpl(se, ';').ToList();
                }
                var df = getDoc(ref pathFrom, ref nameFrom, dn);
                if (isNull(df)) continue;
                var pf = I.getParameters(df);
                if (isNull(pf)) continue;
                foreach (var el in item.Elements())
                {
                    string pathIn = XMLDoc.getAttributeValue(el, "path"), nameIn = XMLDoc.getAttributeValue(el, "name");
                    //if (pathFrom == pathIn) continue;
                    ObjectCollection obj = null; obj = I.COC(obj);
                    var di = getDoc(ref pathIn, ref nameIn, dn);
                    if (isNull(di)) continue;
                    var pi = I.getParameters(di);
                    if (isNull(pi)) continue;
                    Parameter par = null;
                    List<string> spl = new List<string>();
                    foreach (Parameter pr in pf)
                    {
                        if (pr.ParameterType != ParameterTypeEnum.kUserParameter) continue;
                        if (expt.Contains(pr.Name)) continue;
                        if (isNull(pr)) continue;
                        if (isNull(u.getParameter(pi.UserParameters, pr.Name))) continue;
                        if (replPar(pi, pr.Name, out par))
                        {
                            DerivedParameterTable t = par.Parent as DerivedParameterTable;
                            if (t.LinkedParameters.Count == 0) continue;
                            if (t.ReferencedDocumentDescriptor.FullDocumentName != df.FullDocumentName) continue;
                        }
                        obj.Add(pr); spl.Add(pr.Name);
                    }
                    if (obj.Count == 0) continue;
                    var path = pathFrom + nameFrom;
                    DerivedParameterTable tab = checkDerivTable(pi, path);
                    if (isNull(tab))
                    {
                        pi.DerivedParameterTables.Add2(path, obj);
                    }
                    else
                    {
                        changeLink(tab, obj);
                    }
                    foreach (var p in spl)
                    {
                        replPar(pi, p + "_", out par, false);
                    }
                }
            }
        }

//         public static void linkParameters(Document doc, XMLDoc xml, string docName)
//         {
//             string dn = doc.FullFileName;
//             foreach (var item in xml.find("LinkParam"))
//             {
// //                 var fltr = XMLDoc.getSpl(item, "Fltr", ';');
// //                 if (docName.ToLower().IndexOf("base") == -1)
// //                 {
// //                     if (fltr != null && !checkFltr(fltr, docName.ToLower())) continue;
// //                 }
//                 string path = XMLDoc.getAttributeValue(item, "from"), 
//                     //pars = XMLDoc.getAttributeValue(item, "vals"),
//                     part = XMLDoc.getAttributeValue(item, "in"),
//                     se = XMLDoc.getAttributeValue(item, "except");
//                 List<string> expt = new List<string>();
//                 if (!isNull(se))
//                 {
//                     expt = getSpl(se, ';').ToList();
//                 }
//                 //var spl = getSpl(pars, ';');
//                 //if (isNull(spl)) continue;
//                 ObjectCollection obj = null; obj = I.COC(obj);
//                 //if (!path.StartsWith("\\")) path = "\\" + path;
//                 //path = file.newPath(doc.FullFileName, path);
//                 if (path == dn) continue;
//                 //if (!file.check(path)) return;
//                 Document dc = getFromPath(dn, ref path); //I.open(path, false, false);
//                 if (!isNull(part)) doc = getFromPath(dn, ref part);
//                 if (isNull(doc, dc)) continue;
//                 Parameters ups = I.getParameters(doc);
//                 if (isNull(ups)) continue;
//                 Parameters dcPar = I.getParameters(dc);
//                 Parameter par = null;
//                 List<string> spl = new List<string>();
//                 foreach (Parameter pr in dcPar)
//                 {
//                     if (expt.Contains(pr.Name)) continue;
//                     if (pr.ParameterType != ParameterTypeEnum.kUserParameter) continue;
//                     //Parameter pr = getParameter(dcPar, p);
//                     if (isNull(pr)) continue;
//                     if (isNull(u.getParameter(ups.UserParameters, pr.Name))) continue;
//                     if (replPar(ups, pr.Name, out par))
//                     {
//                         DerivedParameterTable t = par.Parent as DerivedParameterTable;
//                         if (t.LinkedParameters.Count == 0) continue;
//                         if (t.ReferencedDocumentDescriptor.FullDocumentName != dc.FullDocumentName) continue;
//                     }
//                     obj.Add(pr); spl.Add(pr.Name);
//                 }
// 
//                 if (obj.Count == 0) continue;
// //                 try
// //                 {
//                     DerivedParameterTable tab = checkDerivTable(ups, path);
//                     if (isNull(tab))
//                     {
//                         ups.DerivedParameterTables.Add2(path, obj);
//                     }
//                     else
//                     {
//                         changeLink(tab, obj);
//                     }
//                     foreach (var p in spl)
//                     {
//                         replPar(ups, p + "_", out par, false);
//                     }
//                 //}
// //                 catch (System.Exception ex)
// //                 {
// // 
// //                 } 
//             }  
//         }

        public static void changeLink(DerivedParameterTable tab, ObjectCollection pars)
        {
            tab.LinkedParameters = pars;
        }

        public static DerivedParameterTable checkDerivTable(Parameters ps, string f)
        {
            return get<DerivedParameterTable>(ps.DerivedParameterTables, fi => fi.ReferencedDocumentDescriptor.FullDocumentName == f);
        }

        public static void updateParameter(Parameter p, string f)
        {
            if (p.Expression == f) return;
            p.Expression = f;
        }

        public static void updateParameter(System.Collections.IEnumerable ps, Dictionary<string, XElement> dic)
        {
            foreach (Parameter item in ps)
            {
                if (!dic.ContainsKey(item.Name)) continue;
                string v = XMLDoc.getAttributeValue(dic[item.Name], "Value");
                updateParameter(item, v);
            }
        }

        public static Parameter getParameter(System.Collections.IEnumerable ps, string name)
        {
            return u.get<Parameter>(ps, f => f.Name == name);
        }

        public static Parameter getParameter(Document doc, string name)
        {
            return getParameter(I.getParameters(doc), name);
        }

        public static List<string> getParamsFromString(string expr)
        {
            List<string> lst = new List<string>();
            if (isNull(expr)) return lst;
            var spl = expr.Split(new char[] { '/', '*', '+', '-', '%', ' ' });
            foreach (var str in spl)
            {
                string tmp = str.Trim(new char[] { '(', ')' });  
                Regex r = new Regex(@"^[^0-9]+\w+");
                MatchCollection m = r.Matches(tmp);
                List<string> fltr = new List<string>() { "mm", "ul", "мм", "бр", "rad", "grad", "град", "isolate", "sin", "cos" };
                foreach (Match item in m)
                {
//                     if (isNull(item.Groups[1])) continue;
                    string v = item.Value;
                    v = v.Trim(new char[] { '-' });
                    if (!fltr.Contains(v)) lst.Add(v);
                }
            }
            return lst;
        }

        public static void addParameter(Document doc, XElement el)
        {
            switch (el.Attribute("Type").Value)
            {
                case "ul":
                    break;
                default:
                    break;
            }
            string group = "", comment = "";
            var fltr = XMLDoc.getSpl(el, "Fltr", ';');
            if (!isNull(XMLDoc.getAttributeValue(el, "noFilter"))) fltr = null;
            if (!isNull(fltr) && !(checkFltr(fltr, getDocName(doc).ToLower())))
            {
                return;
            }
            //string t = XMLDoc.getAttributeValue(el, "fType");
//             if (!isNull(t))
//             {
//                 string dName = u.getDocName(doc);
//                 if (dName.IndexOf("base") != -1 && t == "b")
//             }
            string ipt = XMLDoc.getAttributeValue(el, "ipt");
            if (!isNull(ipt)){
                string templ = XMLDoc.getAttributeValue(el, "templ", "");
                if (isNull(templ)) templ = "";
                string ext = XMLDoc.getAttributeValue(el, "ext");
                if (isNull(ext)) ext = ".ipt";
                string ffn = file.p(doc.FullFileName) + ipt + ext, alias = XMLDoc.getAttributeValue(el, "dirName");
                sdoc = sdoc ?? I.newDoc(ffn, templ, true, alias);
                if (isNull(sdoc)) return;
                if (sdoc.FullFileName != ffn)
                {
                    sdoc.Save2();
                    sdoc = I.newDoc(ffn, alias);
                }
                sdoc.UnitsOfMeasure.LengthUnits = UnitsTypeEnum.kMillimeterLengthUnits;
                doc = sdoc;
            }
            if (el.Attribute("Group") != null) group = el.Attribute("Group").Value;
            if (el.Attribute("Comment") != null) comment = el.Attribute("Comment").Value;
            //if (el.Attribute("Formula") != null)
            addParameter(doc, el.Attribute("Name").Value, el.Attribute("Value").Value, el.Attribute("Type").Value, false, group, comment);
            //else
            //addParameter(doc, el.Attribute("Name").Value, el.Attribute("Value").Value, el.Attribute("Type").Value, true, group);
        }

        public static bool addPrPath(XElement el, string name, string p, string pt = "FrequentlyUsedFolder")
        {
            if (isNull(el)) return false;
            if (!isNull(XMLDoc.find(el, "PathName", name))) return false;
            XElement add = new XElement("ProjectPath", new XAttribute("pathtype", pt), new XElement("PathName", name), new XElement("Path", p));
            el.Add(add);
            return true;
        }

        public static Document changeProjProp(XMLDoc xdoc)
        {
            DesignProject pr = I.app.DesignProjectManager.ActiveDesignProject;
            XMLDoc xml = null;
            XElement xmlel = null; string ws = "";
            string ffn = null; bool s = false;
            xml = new XMLDoc(pr.FullFileName, "InventorProject", ".ipj");
            xmlel = XMLDoc.getXElement(xml.El, "ProjectPaths");
            XElement tel = XMLDoc.find("pathtype", "Workspace", xmlel);
            ws = tel.Element("PathName").Value;
            foreach (var el in xdoc.El.Elements("PrPath"))
            {
                var vals = XMLDoc.getXAttributesValues(el, new string[] { "pathtype", "Name", "Path" });
                if (u.isNull(vals[0], vals[1], vals[2])) continue;
                string prp = vals[2];
                switch (vals[0])
                {
                    case "FrequentlyUsedFolder":
                        prp = ws + prp;
                        break;
                    default:
                        break;
                }
                if (u.addPrPath(xmlel, vals[1], prp, vals[0])) s = true;
            }
            if (s)
            {
                if (!I.app.DesignProjectManager.ActiveDesignProject.Equals(I.app.DesignProjectManager.DesignProjects[1]))
                {
                    ffn = I.closeDocs();
                    I.app.DesignProjectManager.DesignProjects[1].Activate(false);
                }
            }
            if (s && !pr.FullFileName.Equals(I.app.DesignProjectManager.ActiveDesignProject.FullFileName))
            {
                xml.save();
                pr.Activate(false);
                return I.open(ffn, false, false);
            }
            return null;
        }

        public static void addParameter(Document doc, string v, Dictionary<string,XElement> xml)
        {
            
            List<string> lst = getParamsFromString(v);
            foreach (var item in lst)
            {
                if (!isNull(u.getParameter(doc, item))) continue;
                if (!xml.ContainsKey(item)) continue;
                addParameter(doc, XMLDoc.getAttributeValue(xml[item], "Value"), xml);
                string no = XMLDoc.getAttributeValue(xml[item], "no");
                if (!isNull(no) && !findParameter(doc, item, true, xml[item]))
                    addParameter(doc, xml[item]);
                else addParameter(doc, xml[item]);
            }
        }

        public static string parameterToXML(XMLDoc parent, Parameter param, string ret = "")
        {
            if (param.DrivenBy.Count != 0)
                foreach (Parameter item in param.DrivenBy)
                {
                    ret = parameterToXML(parent, item, ret);
                }
            string units = ""; double scale = 10; bool flag = false;
            if (param.Expression.EndsWith("бр")) { units = "ul"; scale = 1; }
            else { units = "mm"; }
            if (param.Expression.IndexOfAny(new char[] { '*', '/', '+', '-' }) != -1) flag = true;
            if (ret == "") ret = param.Expression;
            XElement el = new XElement("Parameter", new XAttribute("Name", param.Name), new XAttribute("Value", convToDouble(param.Value.ToString(), scale, 3).ToString()), new XAttribute("Type", units));
            if (flag) el = new XElement("Parameter", new XAttribute("Name", param.Name), new XAttribute("Value", ret), new XAttribute("Type", units), new XAttribute("Formula", "1"));
            XElement tmp = parent.getXElement(el.Attribute("Name").Value, "Name", "Parameter");
            if (tmp == null)
                parent.El.Add(el);
            if (flag) return param.Expression;
            else return "";
        }

        public static void addParameter(Document doc, XNode el)
        {
            addParameter(doc, el as XElement);
        }

        public static object getProxy(ComponentOccurrence oc, object el)
        {
            object ob;
            oc.CreateGeometryProxy(el, out ob);
            return ob;
        }

        static public bool isNull(params object[] objs)
        {
            if (objs == null) return true;
            foreach (var ob in objs)
            {
                if (ob is string && (ob == null || (string)ob == "")) return true;
                if (ob is int && (ob == null || (int)ob == 0)) return true;
                if (ob is double && (ob == null || (double)ob == 0)) return true;
                if (ob == null) return true;
            }
            return false;
        }

        static public bool checkEndOfPart(Document doc, PartComponentDefinition pcd, string name)
        {
            object be = Entity.endOfPart;
            if (isNull(be)) return false;
            var pane = doc.BrowserPanes.ActivePane;
            var bn = pane.GetBrowserNodeFromObject(be);
            BrowserNode br = bn.Parent as BrowserNode;
            if (isNull(br)) return false;
            string n = bn.FullPath;
            int i = 0;
            foreach (BrowserNode item in br.BrowserNodes)
            {
                if (i == 3) { bn = item; break; }
                if (item.FullPath == n) i++;
                if (i != 0) i++;
            }
            if (bn.FullPath.EndsWith(name)) return true;
            return false;
        }

        public static UserParameter addParameter(Document doc, string name, string value, string typ, bool single = true, string group = "", string comment = "")
        {
            Parameters par = null;
            if (isNull(name, value)) return null;
            if (doc.DocumentType == DocumentTypeEnum.kAssemblyDocumentObject)
            {
                par = ((AssemblyDocument)doc).ComponentDefinition.Parameters;
            }
            else if (doc.DocumentType == DocumentTypeEnum.kPartDocumentObject)
            {
                par = ((PartDocument)doc).ComponentDefinition.Parameters;
            }
            value = value.Replace('.', ',');
            UserParameter old = par.OfType<UserParameter>().FirstOrDefault(p => p.Name == name);
            Parameter pold = getParameter(par, name);
            if (!isNull(pold) && pold.ParameterType == ParameterTypeEnum.kDerivedParameter) 
                return null;
            if (single)
            {
                double val;
                val = convToDouble(value, 1, 3);
                //if (typ != UnitsTypeEnum.kUnitlessUnits) { val /= 10; val = Math.Round(val, 3); }
                if (old != null)
                {
                    if ((double)old.Value != val) old.Value = val;
                    return old;
                }
                UserParameter up = par.UserParameters.AddByValue(name, val, typ);
                if (comment != "") up.Comment = comment;
                if (group != "") addParameterToGroup(doc, up as Parameter, group);
                return up;
            }
            else if (old != null) { old.Expression = value; addParameterToGroup(doc, old as Parameter, group); return old; }
            else
            {
                UserParameter up = null;
                try
                {
                    up = par.UserParameters.AddByExpression(name, value, typ);
                }
                catch
                {
//                     if (typ == "mm") typ = "ul";
//                     else typ = "mm";
//                     up = par.UserParameters.AddByExpression(name, value, typ);
//                    System.Windows.Forms.MessageBox.Show("Не удалось создать параметер со значением" + value);
                }
                if (comment != "") up.Comment = comment;
                if (group != "") addParameterToGroup(doc, up as Parameter, group);
                return up;
            }
        }

        public static CustomParameterGroup getGroup(Document doc, string groupName)
        {
            CustomParameterGroup cpg = null; Parameters pars = null;
            //if (doc.DocumentType == DocumentTypeEnum.kPartDocumentObject && doc.ReferencedDocumentDescriptors.Count == 1)
            //    pars = ((PartDocument)(InvDoc.u.referendedDocDesc(doc)).ReferencedDocument).ComponentDefinition.Parameters;
            if (doc.DocumentType == DocumentTypeEnum.kPartDocumentObject) pars = ((PartDocument)doc).ComponentDefinition.Parameters;
            else if (doc.DocumentType == DocumentTypeEnum.kAssemblyDocumentObject) pars = ((AssemblyDocument)doc).ComponentDefinition.Parameters;
            cpg = u.get<CustomParameterGroup>(pars.CustomParameterGroups, f => f.DisplayName == groupName);
            return (cpg == null) ? pars.CustomParameterGroups.Add(groupName, groupName) : cpg;
        }

        public static void addParameterToGroup(Document doc, Parameter par, string groupName)
        {
            if (isNull(groupName)) return;
            if (isNull(par)) return;
            CustomParameterGroup cpg = getGroup(doc, groupName);
            cpg.Add(par);
        }

        public static void getParametersFromGroup(Document doc, string nameGroup, ref HashSet<string> names)
        {
            CustomParameterGroup cpg = getGroup(doc, nameGroup);
            if (cpg.Count == 0) { cpg.Delete(); return; }
            foreach (Parameter item in cpg)
            {
                names.Add(item.Name);
            }
        }

        public static void addParameters(Document doc, XMLDoc xml, bool replace)
        {  
            xml.setRoot();
            HashSet<string> pnames = new HashSet<string>();
            foreach (var item in xml.find("Parameter"))
            {
                string name = XMLDoc.getAttributeValue(item, "Name");

                if (name == null) continue;
                if (pnames.Contains(name) && isNull(XMLDoc.getAttributeValue(item, "ipt"))) continue;
                pnames.Add(name);
               /* bool f = ;*/
                if (replace) addParameter(doc, item);
                else if (!findParameter(doc, name, true, item))
                    addParameter(doc, item);
            }
        }

        public static bool findParameter(Document doc, string name, bool add = false, XElement el = null)
        {
            bool fl = false;
            if (doc.DocumentType == DocumentTypeEnum.kPartDocumentObject)
            {
                PartComponentDefinition compDef = ((PartDocument)doc).ComponentDefinition;
                Parameter param = get<Parameter>(compDef.Parameters, f => f.Name == name);
                if (param != null) return true;
                if (compDef.ReferenceComponents.DerivedPartComponents.Count == 0) return fl;
                DerivedPartDefinition def = compDef.ReferenceComponents.DerivedPartComponents[1].Definition;
                fl = derivParams(ref def, name);
                if (add && !fl && el != null) 
                {
                    addParameter(compDef.ReferenceComponents.DerivedPartComponents[1].ReferencedDocumentDescriptor.ReferencedDocument as Document, el);
                    def = compDef.ReferenceComponents.DerivedPartComponents[1].Definition;
                    fl = derivParams(ref def, name);
                }
                compDef.ReferenceComponents.DerivedPartComponents[1].Definition = def;
            }
            return fl;
        }
        public static string getParameter(Document doc, Regex r)
        {
            string rname = null;
            PartComponentDefinition pdef = I.getPCD(doc);
            if (pdef == null) return null;
            Parameter p = get<Parameter>(pdef.Parameters, f => r.IsMatch(f.Name));
            if (p != null) return p.Name;
            if (pdef.ReferenceComponents.DerivedPartComponents.Count == 0) return null;
            DerivedPartDefinition def = pdef.ReferenceComponents.DerivedPartComponents[1].Definition;
            foreach (DerivedPartEntity item in def.Parameters)
            {
                UserParameter up = item.ReferencedEntity as UserParameter;
                if (up == null) continue;
                if (r.IsMatch(up.Name))
                {
                    item.IncludeEntity = true; break;
                    rname = up.Name;
                }
            }
            pdef.ReferenceComponents.DerivedPartComponents[1].Definition = def;
            return rname;
        }
        public static bool derivParams(ref DerivedPartDefinition def, string name)
        {
             bool fl = false;
            foreach (DerivedPartEntity par in def.Parameters)
            {
                if (par.ReferencedEntity is UserParameter)
                {
                    UserParameter p = par.ReferencedEntity as UserParameter;
                    if (p.Name == name)
                    {
                        par.IncludeEntity = true; fl = true;
                        break;
                    }
                }
            }
            return fl;
        }
        public static string getSubS(string s, int l)
        {
            if (s.Length < l) return s;
            else return s.Substring(0, l);
        }

        public static Point2d addPoint(SketchPoint sp, Vector2d dir)
        {
            Point2d pt = I.CP2d(sp);
            pt.TranslateBy(dir);
            return pt;
        }

        public static SketchLine addLine(PlanarSketch ps, SketchPoint sp, Vector2d dir, bool c = false)
        {
            SketchLine l = ps.SketchLines.AddByTwoPoints(sp, addPoint(sp, dir));
            if (c) l.Construction = true;
            return l;
        }

        public static Vector2d getDirect(SketchEntity se)
        {
            SketchLine sl = se as SketchLine;
            if (isNull(sl)) return null;
            return sl.StartSketchPoint.Geometry.VectorTo(sl.EndSketchPoint.Geometry);
        }

        public static WorkPlane findPlane(Document doc, string name, PartComponentDefinition pcd = null)
        {
            bool r = false;
            PartComponentDefinition def;
            def = isNull(pcd) ? I.getPCD(doc) : pcd;
            WorkPlanes pls = null;
            if (!isNull(def)) pls = def.WorkPlanes;
            else
            {
                var acd = I.getACD(doc); pls = acd.WorkPlanes;
            }
            WorkPlane s = null;
            if (name.Length == 1)
            {
                int num = u.convToInt(name);
                if (!isNull(num)) s = get(pls, num) as WorkPlane;
            }
            else s = get<WorkPlane>(pls, f => f.Name == name);
            if (s != null) return s;
            if (def == null || def.ReferenceComponents.DerivedPartComponents.Count == 0) return null;
            DerivedPartUniformScaleDef der = def.ReferenceComponents.DerivedPartComponents[1].Definition as DerivedPartUniformScaleDef;
            if (der == null) return null;
            action<DerivedPartEntity>(der.WorkFeatures, a => { a.IncludeEntity = true; r = true; }, f => ((WorkPlane)f.ReferencedEntity).Name == name);
            if (r)
            {
                def.ReferenceComponents.DerivedPartComponents[1].Definition = der as DerivedPartDefinition;
                return get<WorkPlane>(def.WorkPlanes, f => f.Name == name);
            }
            return null;
        }

        public static bool findSketch(Document doc, string name)
        {
            bool r = false;
//             string desc = u.getPropValue(doc,"Description");
//             if (desc != "")
//             {
//                 if (!desc.StartsWith(getSubS(name, 4))) return false;
//             }
            var def = I.getPCD(doc);
            PlanarSketch s = get<PlanarSketch>(def.Sketches, f => f.Name == name);
            if (s != null) return true;
            if (def == null || def.ReferenceComponents.DerivedPartComponents.Count == 0) return false;
            DerivedPartUniformScaleDef der = def.ReferenceComponents.DerivedPartComponents[1].Definition as DerivedPartUniformScaleDef;
            if (der == null) return false;
            action<DerivedPartEntity>(der.Sketches, a => { a.IncludeEntity = true; r = true; }, f => ((PlanarSketch)f.ReferencedEntity).Name == name);
            if (r) def.ReferenceComponents.DerivedPartComponents[1].Definition = der as DerivedPartDefinition;
            return r;
        }

        public static void findParameter(Document doc, HashSet<string> names)
        {
            PartComponentDefinition compDef = ((PartDocument)doc).ComponentDefinition;
            if (compDef.ReferenceComponents.DerivedPartComponents.Count == 0) return;
            DerivedPartDefinition def = compDef.ReferenceComponents.DerivedPartComponents[1].Definition;
            if (doc.DocumentType != DocumentTypeEnum.kPartDocumentObject) return;
            foreach (var name in names)
            {
                foreach (DerivedPartEntity par in def.Parameters)
                {
                    if (par.ReferencedEntity is UserParameter)
                    {
                        UserParameter p = par.ReferencedEntity as UserParameter;
                        if (p.Name == name)
                        {
                            par.IncludeEntity = true;
                            break;
                        }
                    }
                }
            }
            compDef.ReferenceComponents.DerivedPartComponents[1].Definition = def;
        }

        public static LeaderNote findLeaderText(DrawingDocument doc, string text)
        {
            foreach (LeaderNote l in doc.ActiveSheet.DrawingNotes.LeaderNotes)
            {
                if (l.Text.IndexOf(text) != -1) return l;
            }
            return null;
        }

        public static void replaceText(DrawingDocument doc, string find, string replace)
        {
            LeaderNote l = findLeaderText(doc, find);
            if (l != null) l.FormattedText = replace;
        }

        public static void autoComplete(System.Windows.Forms.TextBox tb, XMLDoc xmldoc, string name, string attName)
        {
            tb.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend;
            tb.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.CustomSource;
            System.Windows.Forms.AutoCompleteStringCollection col = new System.Windows.Forms.AutoCompleteStringCollection();
            xmldoc.addItems<System.Windows.Forms.AutoCompleteStringCollection>(col, name, attName);
            tb.AutoCompleteCustomSource = col;
        }

        public static void autoComplete<T>(System.Windows.Forms.TextBox tb, T xmldoc, string name) where T : PartDocument
        {
            tb.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend;
            tb.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.CustomSource;
            System.Windows.Forms.AutoCompleteStringCollection col = new System.Windows.Forms.AutoCompleteStringCollection();
            if (xmldoc is PartDocument)
            {
                foreach (UserParameter item in xmldoc.ComponentDefinition.Parameters)
                {
                    if (item.Value.ToString().IndexOf(name) != -1)
                        col.Add(item.Value.ToString());
                }
            }
            tb.AutoCompleteCustomSource = col;
        }

        public static string pathUtil(this Inventor.Document doc)
        {
            return doc.FullFileName.Substring(0, doc.FullFileName.LastIndexOf('\\'));
        }

        public static string createDir(Document doc, string add = "")
        {
            string path = pathDoc(doc, add);
            System.IO.Directory.CreateDirectory(path);
            return path;
        }

        public static string pathDoc(Document doc, string add = "")
        {
            string name = doc.FullDocumentName;
            if (add != "") return pathDoc(name) + add;
            return pathDoc(name);
        }

        public static string pathProj()
        {
            string p = I.app.DesignProjectManager.ActiveDesignProject.WorkspacePath;
            //if (isNull(p)) p = I.app.DesignProjectManager.ActiveDesignProject.WorkgroupPaths[1].
            return p;
        }
        public static void checkSubPath(Document doc, XMLDoc xml, bool first = true)
        {

            string sp = file.p(doc.FullDocumentName);
            sp = sp.TrimEnd(new char[] { '\\' });
            List<XElement> rem = new List<XElement>();
            foreach (var item in xml.El.Elements())
            {
                string s = XMLDoc.getAttributeValue(item, "subPath"), snum = XMLDoc.getAttributeValue(item, "level");
                int level = u.convToInt(snum);
                if (u.isNull(s)) continue;
                s = s.Trim(new char[] { '\\' });
                if (s == ".")
                {
                    if (sp != I.curProjPath())
                        rem.Add(item);
                }
                else if (level != 0)
                {
                    if (!endSubPath(sp, s, level)) rem.Add(item);
                }
                else if (!sp.EndsWith(s)) rem.Add(item);
            }
            foreach (var item in rem)
            {
                item.Remove();
            }
            if (first)
            {
                xml.insert(maxLevel: 2);
                checkSubPath(doc, xml, false);
            }
        }

        public static bool endSubPath(string fn, string f, int level = 1)
        {
            for (int i = 0; i < level; i++)
            {
                if (fn.EndsWith(f)) return true;
                fn = file.trimEnd(fn, "\\");
            }
            return false;
        }

        public static void paramFilter(Document doc, XMLDoc xml)
        {
            List<XElement> rem = new List<XElement>();
            p = null;
            paramFilter(doc, xml.El, ref rem); 
            foreach (var item in rem)
            {
                item.Remove();
            }
            p = null;
        }

        public static void paramFilter(Document doc, XElement xml, ref List<XElement> rem)
        {
            int c = 0;
            foreach (var item in xml.Elements())
            {
                paramFilter(doc, item, ref rem);
            }
            string v = XMLDoc.getAttributeValue(xml, "fParam");
            if (isNull(v)) return ;
            var spl = getSpl(v, ';');
            foreach (var item in spl)
            {
                string n = getElem(item, 0), val = getElem(item, 1);
                if (isNull(n, val)) continue;
                p = (!isNull(p) && p.Name == n) ? p : getParameter(doc, n);
                if (isNull(p)) continue;
                if (!eq(p.ModelValue, convToDouble(val))) c++;
            }
            if (c == spl.Length) rem.Add(xml);
        }

        static public string getElem(string v, int num, char sep = ':')
        {
            var spl = getSpl(v, sep);
            if (spl == null) return null;
            if (spl.Length <= num) return null;
            return spl[num];
        }

        static public double getDoubleElem(string v, int num, char sep = ':', double scale = 1, int tol = 3)
        {
            return convToDouble(getElem(v, num, sep),scale,tol);
        }

        static public int getIntElem(string v, int num, char sep = ':')
        {
            return convToInt(getElem(v, num, sep));
        }

        public static string pathDoc(string doc)
        {
            return doc.Substring(0, doc.LastIndexOf('\\'));
        }

        public static string nameUtil(this Inventor.Document doc)
        {
            return doc.FullFileName.Substring(doc.FullFileName.LastIndexOf('\\') + 1, doc.FullFileName.Length - 1 - doc.FullFileName.LastIndexOf('\\') - 4);
        }

        public static string nameUtil(string doc, bool ext = false)
        {
            int ind = 4;
            if (!ext) ind = 0;
            return doc.Substring(doc.LastIndexOf('\\') + 1, doc.Length - 1 - doc.LastIndexOf('\\') - ind);
        }

        static public bool intersectWithBox(Inventor.Box2d box, ref HashSet<Inventor.Box2d> boxes)
        {
            foreach (Inventor.Box2d item in boxes)
            {
                if (box.IsDisjoint(item))
                {
                    return true;
                }
            }
            return false;
        }

        static public string nameForSave(Document doc, bool drw = false)
        {
            string suf = "", pn = "", desc = "", lit1 = "", lit2 = "";
            switch (doc.DocumentType)
            {
                case DocumentTypeEnum.kAssemblyDocumentObject:
                    suf = ".iam";
                    break;
                case DocumentTypeEnum.kDrawingDocumentObject:
                    suf = ".idw";
                    break;
                case DocumentTypeEnum.kPartDocumentObject:
                    suf = ".ipt";
                    break;
                default:
                    break;
            }
            if (drw) suf = ".idw";
            Property p = null;
            p = getProp(doc, "Part Number");
            if (p != null) pn = p.Value.ToString();
            p = getProp(doc, "Description");
            if (p != null) desc = p.Value.ToString();
            p = getProp(doc, "Литера1");
            if (p != null) lit1 = p.Value.ToString();
            p = getProp(doc, "Литера2");
            if (p != null) lit2 = p.Value.ToString();
            return pn + " (" + desc + ")" /*+ lit1 + lit2*/ + suf;
        }

        static public string nameForSave(XElement el, string typ, string cnt = "")
        {
            if (cnt != "") cnt = "^" + cnt;
            string pn = typ + "." + el.Attribute("DecNumber").Value, desc = el.Attribute("Name").Value, suf = nameSuff(el);
            return pn + " (" + desc + ")" /*+ lit1 + lit2*/ + suf + cnt;
        }

        static public string nameSuff(XElement el)
        {
            string suf = "";
            if (el.Name == "Assembly") suf = ".iam";
            else if (el.Name == "Part") suf = ".ipt";
            return suf;
        }

        static public void partNumber(Document doc, string decNumber, string type)
        {
            Property p = null;
            addProp(doc, "Type", type);
            addProp(doc, "DecNumber", decNumber);
            p = getProp(doc, "Part Number");
            p.Expression = "=<Type>.<DecNumber>";
        }

        static public bool occsContaints(ComponentOccurrences occs, string find, int count)
        {
            int cnt = 0;
            if (!file.check(find)) return true;
            foreach (ComponentOccurrence occ in occs)
            {
                string name = occ.ReferencedDocumentDescriptor.FullDocumentName;
                //                 if (name.IndexOf(':') != -1) {
                //                     var spl = name.Split(':');
                //                     name = spl[0];
                //                 }
                if (find == name) { cnt++; }
            }
            if (cnt > 0 && count <= cnt) { return true; }
            return false;
        }

        static public void release()
        {
            foreach (Document d in I.app.Documents)
            {
                d.ReleaseReference();
            }
            I.app.Documents.CloseAll(true);
        }

        static public void transactStart(Document doc, string name)
        {
            if (tr != null) tr.Abort();
            tr = I.app.TransactionManager.StartTransaction(doc as _Document, name);
        }
        static public void transactEnd()
        {
            if (tr != null) { tr.End(); tr = null; }
        }
        static public void transactUndo()
        {
            if (tr != null)
            {
                I.app.TransactionManager.UndoTransaction();
            }
        }

        static public void setUpdate()
        {
            I.app.ScreenUpdating = !I.app.ScreenUpdating;
        }

        static public void setSilence()
        {
            I.app.SilentOperation = !I.app.SilentOperation;
        }
        static public bool check<T>(T val, T find, Func<T, T, bool> cmp)
        {
            return cmp(val, find);
        }
        static public double scale(double v, double s)
        {
            return v * s;
        }
    }
    static public class Highlight
    {
        static public List<HighlightSet> sets;
        static public List<byte[]> colors;
        static Document d;
        static int i = 0;
        static public void set(Document doc, List<byte[]> col = null)
        {
            d = doc;
            sets = new List<HighlightSet>();
            colors = new List<byte[]>()
            {
                new byte[] {255,0,0}, new byte[] {0,255,0}, new byte[] {0,0,255}, new byte[] {255,255,0},
                new byte[] {0,255,255}, new byte[] {255,0,255}
            };
            if (col != null) colors = col;
        }
        static public void add(object o)
        {
            var to = I.app.TransientObjects;
            var color = to.CreateColor(colors[i][0], colors[i][1], colors[i][2]); i++;
            if (i >= colors.Count) i = 0;
            var hs = d.CreateHighlightSet();
            hs.Color = color;
            hs.AddItem(o);
            sets.Add(hs);
        }
        static public void clear()
        {
            foreach (var set in sets)
            {
                set.Clear();
            }
            sets.Clear();
        }
    }
    public class LambdaComparer<T> : IEqualityComparer<T>
    {
        private readonly Func<T, T, bool> _lambdaComparer;
        private readonly Func<T, int> _lambdaHash;

        public LambdaComparer(Func<T, T, bool> lambdaComparer) :
            this(lambdaComparer, o => 0)
        {
        }

        public LambdaComparer(Func<T, T, bool> lambdaComparer, Func<T, int> lambdaHash)
        {
            if (lambdaComparer == null)
                throw new ArgumentNullException("lambdaComparer");
            if (lambdaHash == null)
                throw new ArgumentNullException("lambdaHash");

            _lambdaComparer = lambdaComparer;
            _lambdaHash = lambdaHash;
        }

        public bool Equals(T x, T y)
        {
            return _lambdaComparer(x, y);
        }

        public int GetHashCode(T obj)
        {
            return _lambdaHash(obj);
        }
    }
    enum pType
    {
        b, s, a, u, c
    }
    public class GetDocs
    {
        public HashSet<Document> docs = new HashSet<Document>();
        public GetDocs()
        {
            foreach (Document item in Macros.StandardAddInServer.m_inventorApplication.Documents.VisibleDocuments)
            {
                add(item);
            }
        }
        public GetDocs(Document doc)
        {
            add(doc);
        }
        public void add(Document doc)
        {
            docs.Add(doc);
            foreach (Document item in doc.ReferencedDocuments)
            {
                add(item);
            }
        }
        public void rename(Dictionary<string, string> dic)
        {
            foreach (Document item in docs)
            {
                if (item.DocumentType != DocumentTypeEnum.kDrawingDocumentObject) continue;
                Sheet sh = (item as DrawingDocument).Sheets[1];
                foreach (DrawingSketch s in sh.Sketches)
                {
                    s.Edit();
                    foreach (TextBox tb in s.TextBoxes)
                    {
                        if (dic.ContainsKey(tb.Text))
                            tb.Text = dic[tb.Text];
                    }
                    s.ExitEdit();
                }
            }
        }
    }
}
