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
                InvDoc.XML n = new InvDoc.XML(@"C:\ProgramData\Autodesk\Inventor Addins\TemplatePath.xml");
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
            if (doc == null) doc = addPrtDoc();
            SheetMetalComponentDefinition partDef = null;
            SheetMetalStyle sms = null;
            if (name.EndsWith(".ipt"))
            {
                if (doc == null)
                {
                    partDef = (SheetMetalComponentDefinition)((PartDocument)invApp.Documents.Open(name, false)).ComponentDefinition;
                    sms = partDef.ActiveSheetMetalStyle;
                }

                derDef = doc.ComponentDefinition.ReferenceComponents.DerivedPartComponents.CreateUniformScaleDef(name);
                //derDef.IncludeAllSketches = DerivedComponentOptionEnum.kDerivedExcludeAll;
                derDef.IncludeAlliMateDefinitions = DerivedComponentOptionEnum.kDerivedIncludeAll;
                derDef.UseColorOverridesFromSource = true;
                derDef.DeriveStyle = DerivedComponentStyleEnum.kDeriveAsSingleBodyWithSeams;
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
                        tmp = nameFind.Split(';');
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
                DerivedPartComponent derComp = doc.ComponentDefinition.ReferenceComponents.DerivedPartComponents.Add((DerivedPartDefinition)derDef);
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

        static public DrawingView addView(Sheet sh, Document doc, XElement e, string list, double s, List<double> scales)
        {
            if (!System.IO.File.Exists(@"C:\ProgramData\Autodesk\Inventor Addins\sheet.xml")) return null;
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
                    view = addDV(doc, pt, sh, TypeViews.BaseView, s, viewEnum, nvm);
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

        static public DrawingView addDV(Document doc, Point2d pt, Sheet sh, TypeViews type, double scale, ViewOrientationTypeEnum viewEnum, NameValueMap nvm, DrawingView parent = null, bool align = false, bool offseV = false)
        {
            DrawingView dv = null;
            switch (type)
            {
                case TypeViews.BaseView:
                    dv = sh.DrawingViews.AddBaseView((_Document)doc, pt, scale, viewEnum, DrawingViewStyleEnum.kHiddenLineRemovedDrawingViewStyle, AdditionalOptions: nvm);
                    dv.IsRasterView = false;
                    //doc.Update();
                    if (nvm.Count == 0)
                    {
                        if (viewEnum != ViewOrientationTypeEnum.kCurrentViewOrientation)
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
            sh = sheet(sh, w, h);
            tr.End();
            return sh;
        }

        public Sheet changeSheet(Sheet sh ,double w = 420, double h = 297)
        {
            tr = invApp.TransactionManager.StartTransaction((_Document)doc, "Добавить лист");
            sh = sheet(sh, w, h);
            tr.End();
            return sh;
        }

        private Sheet sheet(Sheet sh ,double w = 420, double h = 297)
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
            else tbd = drw.TitleBlockDefinitions["ГОСТ - Форма 2a"];
            }
            catch (Exception)
            {
                temp = temp ?? openDrwDoc(pathTemplate[0]);
                copySketchSymbolDefinition(temp, drw, new string [] { "ГОСТ - Форма 1", "ГОСТ - Форма 2a"});
                //temp.TitleBlockDefinitions["ГОСТ - Форма 1"].CopyTo((_DrawingDocument)drw, true);
                //temp.TitleBlockDefinitions["ГОСТ - Форма 2a"].CopyTo((_DrawingDocument)drw, true);
                if (sh.Name == "Лист:1") tbd = drw.TitleBlockDefinitions["ГОСТ - Форма 1"];
                else tbd = drw.TitleBlockDefinitions["ГОСТ - Форма 2a"];
                temp.Close();
            }

            if (sh.TitleBlock == null)
            {
                //sh.TitleBlock.Delete();
                sh.AddTitleBlock(tbd);
            }

            List<string> except = new List<string> { "ГОСТ - Доп. графы 1", "ГОСТ - Доп. графы 2", "ГОСТ - Доп. графы 3", "АР", "Развертка 1:N АР"};
            List<Point2d> poz = new List<Point2d> {invApp.TransientGeometry.CreatePoint2d(2,0.5), invApp.TransientGeometry.CreatePoint2d(2,h*0.1 - 12.5), 
                invApp.TransientGeometry.CreatePoint2d(2,h*0.1 - 0.5)}; 

            for (int i = 0; i < except.Count; i++)
			{
                SketchedSymbol ss = sh.SketchedSymbols.OfType<SketchedSymbol>().FirstOrDefault(e => e.Name == except[i]);
                if (ss == null && i <3)
                    sh.SketchedSymbols.Add(drw.SketchedSymbolDefinitions[except[i]], poz[i]);
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
                            Point2d pt = invApp.TransientGeometry.CreatePoint2d(tb.Origin.X + offsetX, tb.Origin.Y + offsetY);
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
            name = w + "x" + h;
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
            if (vec != null) ss.Position = InvDoc.u.translate(ss.Position, vec);
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

        static public void withElem<T>(List<T> col, Action<T> act)
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
        public static HighlightSet hs = null;
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
        public static string getPropValue(Property p)
        {
            return p == null ? "" : p.Value.ToString();
        }

        public static string getPropValue(Document doc, string name)
        {
            Property p = getProp(doc, name);
            return p == null ? "" : p.Value.ToString();
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

        public static void addProp(Document doc, string name, object val)
        {
            Property p = getProp(doc, name);
            if (p != null && p.Value.ToString() != val.ToString()) p.Value = val;
            else if (p == null) doc.PropertySets["Inventor User Defined Properties"].Add(val, name);
        }

        static public object findAtPoint(SketchedSymbolDefinition ssd, Point2d pt)
        {
            foreach (Inventor.TextBox tb in ssd.Sketch.TextBoxes)
            {
                if (tb.RangeBox.Contains(pt)) return tb;
            }
            return null;
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

        static public bool eq(Point pt1, UnitVector v1, Point pt2, double l = 1.5)
        {
            Vector v = pt1.VectorTo(pt2);
            double d = Math.Abs(scalar(v, v1.AsVector()));
            if (d > l) return false;
            if (isNullVector(v)) return false;
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

        static public T findAtRay<T>(PartComponentDefinition compDef, Point pt, UnitVector dir, SelectionFilterEnum[] filter, double tol = 15, int ind = 0)
        {
            object locPts = null;
            ObjectsEnumerator en = compDef.FindUsingVector(pt, dir, ref filter, true, tol, true, out locPts);
            MeasureTools meas = I.app.MeasureTools;
            if (en.Count == 0) return default(T);
            T ax = en.OfType<T>().OrderBy(e => meas.GetMinimumDistance(pt, e)).ElementAt(ind);
            return ax;
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

        static public bool directionIs(Vector2d v1, Vector2d v2, bool x = true)
        {
            if (x)
                return (Math.Sign(v1.X) == Math.Sign(v2.X));
            else
                return (Math.Sign(v1.Y) == Math.Sign(v2.Y));
        }

        static public bool isCollinear(UnitVector2d vec1, UnitVector2d vec2)
        {
            return vec1.IsEqualTo(vec2);
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

        static public bool eq(Point pt1, Point pt2)
        {
            return (eq(pt1.X, pt2.X) && eq(pt1.Y, pt2.Y) && eq(pt1.Z, pt2.Z));
        }

        static public bool eq(Vertex pt1, Vertex pt2)
        {
            return (eq(pt1.Point.X, pt2.Point.X) && eq(pt1.Point.Y, pt2.Point.Y) && eq(pt1.Point.Z, pt2.Point.Z));
        }

        static public bool eq(UnitVector vec1, UnitVector vec2, bool ab = false)
        {
            if (ab) { abs(ref vec1); abs(ref vec2); }
            return (eq(vec1.X, vec2.X) && eq(vec1.Y, vec2.Y) && eq(vec1.Z, vec2.Z));
        }

        static public bool eq(Vector vec1, Vector vec2)
        {
            return (eq(vec1.X, vec2.X) && eq(vec1.Y, vec2.Y) && eq(vec1.Z, vec2.Z));
        }

        static public bool eq(SketchLine sl1, SketchLine sl2)
        {
            if (eq(sl1.StartSketchPoint.Geometry, sl2.StartSketchPoint.Geometry) && eq(sl1.EndSketchPoint.Geometry, sl2.EndSketchPoint.Geometry)) return true;
            return false;
        }

        static public void abs(ref UnitVector vec)
        {
            vec.X = Math.Abs(vec.X); vec.Y = Math.Abs(vec.Y); vec.Z = Math.Abs(vec.Z);
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


        static public double getLenght(Edge e, int tol = 3)
        {
            double minParam, maxParam, length;
            CurveEvaluator eval = e.Evaluator;
            eval.GetParamExtents(out minParam, out maxParam);
            eval.GetLengthAtParam(minParam, maxParam, out length);
            length = Math.Round(length, tol);
            return length;
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

        static public double[] getPointAtParam(Curve2dEvaluator e, double[] p)
        {
            double[] pts = new double[p.Length * 2];
            e.GetPointAtParam(ref p, ref pts);
            return pts;
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
            foreach (T item in ie)
            {
                if (f(item)) return item as T;
            }
            return null;
        }

        static public IEnumerable<T> gets<T>(System.Collections.IEnumerable ie, Func<T, bool> filter)
        {
            foreach (T item in ie)
            {
                if (filter(item)) yield return item;
            }
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

        static public string getName(string name, char sep, int ind)
        {
            var spl = name.Split(sep);
            return spl[ind];
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

        public static double convToDouble(string val)
        {
            if (val == "") return 0;
            return double.Parse(val.Replace('.', ','));
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
            return Math.Round(double.Parse(val.Replace('.', ',')) * scale, tol);
        }

        public static string convToString(double val, int tol, string format)
        {
            return Math.Round(val, tol).ToString(format);
        }

        public static double convToInt(string val)
        {
            if (val == "") return 0;
            try
            {
                return int.Parse(val);
            }
            catch
            {
                return 0;
            }
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
                offsetN = scale * vec.Length;
                offsetDir = scale * vec.Length;
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

        public static Point2d midPt(Point2d pt1, Point2d pt2, double w = 0.5)
        {
            Point2d pt = pt1.Copy();
            Vector2d vec = pt.VectorTo(pt2);
            vec.ScaleBy(w);
            pt.TranslateBy(vec);
            return pt;
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

        static public string OFD(string iniDir, string filter = "Inventor Files (*.iam;*.ipt)|*.iam;*.ipt|All Files (*.*)|*.*", bool multi = false)
        {
            FileDialog fd;
            I.app.CreateFileDialog(out fd);
            fd.InitialDirectory = iniDir;
            fd.Filter = filter;
            fd.FilterIndex = 1;
            fd.MultiSelectEnabled = multi;
            fd.CancelError = true;
            try
            {
                fd.ShowOpen();
            }
            catch (Exception)
            {
                return "";
            }
            //if (!fd.CancelError) return "";
            return fd.FileName;
        }

        static public string WOFD(string iniDir, string filter = "XML Files|*.xml", string [] custom = null)
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
            ofd.InitialDirectory = iniDir;
            ofd.Filter = filter;
            try
            {
                ofd.ShowDialog();
            }
            catch (Exception)
            {
                return "";
            }
            return ofd.FileName;
        }

        static public XMLDoc XOFD(string iniDir, string rootName = "head")
        {
            string name = WOFD(iniDir);
            if (name == "") return null;
            XMLDoc xdoc = new XMLDoc(name, rootName);
            return xdoc;
        }

        static public Transaction transAct(object doc, string name)
        {
            Transaction tr = I.app.TransactionManager.StartTransaction((_Document)doc, name);
            return tr;
        }

        static public void copySS(string n, DrawingDocument m_Drw)
        {
            string drwTemplate = I.app.DesignProjectManager.ActiveDesignProject.TemplatesPath;
            XML name = new XML(@"C:\ProgramData\Autodesk\Inventor Addins\TemplatePath.xml");
            System.Collections.Generic.List<string> strs = new System.Collections.Generic.List<string>();
            strs = name.ReadXML("Template", "DrawingTemplate");
            drwTemplate = drwTemplate + strs[0];
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
            return sh.Centerlines.Add(col);
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

        public static double strToDouble(string value)
        {
            value = value.Replace('.', ',');
            return Double.Parse(value);
        }

        public static void addParameter(Document doc, XElement el)
        {
            UnitsTypeEnum typ = UnitsTypeEnum.kMillimeterLengthUnits;
            switch (el.Attribute("Type").Value)
            {
                case "ul":
                    typ = UnitsTypeEnum.kUnitlessUnits;
                    break;
                default:
                    break;
            }
            string group = "", comment = "";
            if (el.Attribute("Group") != null) group = el.Attribute("Group").Value;
            if (el.Attribute("Comment") != null) comment = el.Attribute("Comment").Value;
            //if (el.Attribute("Formula") != null)
            addParameter(doc, el.Attribute("Name").Value, el.Attribute("Value").Value, el.Attribute("Type").Value, false, group, comment);
            //else
            //addParameter(doc, el.Attribute("Name").Value, el.Attribute("Value").Value, el.Attribute("Type").Value, true, group);
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

        public static UserParameter addParameter(Document doc, string name, string value, string typ, bool single = true, string group = "", string comment = "")
        {
            Parameters par = null;
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
            else if (old != null) { old.Expression = value; return old; }
            else
            {
                UserParameter up = null;
                try
                {
                    up = par.UserParameters.AddByExpression(name, value, typ);
                }
                catch
                {
                    if (typ == "mm") typ = "ul";
                    else typ = "mm";
                    up = par.UserParameters.AddByExpression(name, value, typ);
                }
                if (comment != "") up.Comment = comment;
                if (group != "") addParameterToGroup(doc, up as Parameter, group);
                return up;
            }
        }

        public static CustomParameterGroup getGroup(Document doc, string groupName)
        {
            CustomParameterGroup cpg = null; Parameters pars = null;
            if (doc.DocumentType == DocumentTypeEnum.kPartDocumentObject && doc.ReferencedDocumentDescriptors.Count == 1)
                pars = ((PartDocument)(InvDoc.u.referendedDocDesc(doc)).ReferencedDocument).ComponentDefinition.Parameters;
            else if (doc.DocumentType == DocumentTypeEnum.kPartDocumentObject) pars = ((PartDocument)doc).ComponentDefinition.Parameters;
            if (doc.DocumentType == DocumentTypeEnum.kAssemblyDocumentObject) pars = ((AssemblyDocument)doc).ComponentDefinition.Parameters;
            cpg = pars.CustomParameterGroups.OfType<CustomParameterGroup>().FirstOrDefault(e => e.DisplayName == groupName);
            return (cpg == null) ? pars.CustomParameterGroups.Add(groupName, groupName) : cpg;
        }

        public static void addParameterToGroup(Document doc, Parameter par, string groupName)
        {
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

        public static bool findParameter(Document doc, string name, bool add = false, XElement el = null)
        {
            bool fl = false;
            if (doc.DocumentType == DocumentTypeEnum.kPartDocumentObject)
            {
                PartComponentDefinition compDef = ((PartDocument)doc).ComponentDefinition;
                if (compDef.ReferenceComponents.DerivedPartComponents.Count == 0) return fl;
                DerivedPartDefinition def = compDef.ReferenceComponents.DerivedPartComponents[1].Definition;
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
                if (add && !fl && el != null) 
                {
                    addParameter(compDef.ReferenceComponents.DerivedPartComponents[1].ReferencedDocumentDescriptor.ReferencedDocument as Document, el);
                }
                compDef.ReferenceComponents.DerivedPartComponents[1].Definition = def;
            }
            return fl;
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
            return doc.FullFileName.Substring(0, doc.FullFileName.LastIndexOf('\\')) + add;
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

        static public string nameForSave(XElement el, string typ)
        {
            string pn = typ + "." + el.Attribute("DecNumber").Value, desc = el.Attribute("Name").Value, suf = ".ipt";
            if (el.Name == "Assembly") suf = ".iam";
            else if (el.Name == "Part") suf = ".ipt";
            return pn + " (" + desc + ")" /*+ lit1 + lit2*/ + suf;
        }

        static public void partNumber(Document doc, string decNumber, string type)
        {
            Property p = null;
            addProp(doc, "Type", type);
            addProp(doc, "DecNumber", decNumber);
            p = getProp(doc, "Part Number");
            p.Expression = "=<Type>.<DecNumber>";
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
