using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Inventor;
using u = InvDoc.u;

namespace InvAddIn
{
    public static class I
    {
        static public readonly Application app;
        static public readonly TransientGeometry tg; 
        static public readonly TransientObjects objs;
        static public readonly NameValueMap nvm, nvmDrw;
        static public string path = "", addInPath = "";
        static public HashSet<SurfaceBody> sb = new HashSet<SurfaceBody>();
        static public List<Document> docs;
        static public DimensionStyle dimStyle;
        static I()
        {
            app = Macros.StandardAddInServer.m_inventorApplication;
            tg = app.TransientGeometry;
            objs = app.TransientObjects;
            nvm = objs.CreateNameValueMap();
            nvmDrw = objs.CreateNameValueMap();
            nvm.Add("SkipAllUnresolvedFiles", true);
            nvmDrw.Add("SkipAllUnresolvedFiles", true);
            nvmDrw.Add("DeferUpdates", true);
        }
        static public string p()
        {
            if (addInPath == "") addInPath = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
            return addInPath;
        }
        static public DocumentsEnumerator visDocs()
        {
            return I.app.Documents.VisibleDocuments;
        }
        static public DrawingDocument addDrw(string fn)
        {
            DrawingDocument drw = null;
            if (InvDoc.file.check(fn)) drw = open(fn) as DrawingDocument;
            else
            {
                drw = app.Documents.Add(DocumentTypeEnum.kDrawingDocumentObject, CreateVisible: false) as DrawingDocument;
                drw.SaveAs(fn, false);
            }
            return drw;
        }
        static public AssemblyDocument addAsm(string fn, bool vis)
        {
            AssemblyDocument asm = null;
            if (InvDoc.file.check(fn)) asm = open(fn) as AssemblyDocument;
            else
            {
                asm = app.Documents.Add(DocumentTypeEnum.kAssemblyDocumentObject, CreateVisible: vis) as AssemblyDocument;
                asm.SaveAs(fn, false);
            }
            return asm;
        }
        static public IEnumerable<Document> getDocs(Func<Document, bool> f)
        {
            return u.gets<Document>(app.Documents.VisibleDocuments, fi => f(fi));
        }
        static public ObjectCollectionByVariant variant(string name ,object val)
        {
            var col = objs.CreateObjectCollectionByVariant();
            col.Add(name, val);
            return col;
        }
        static public Transaction beginTrans(string name, Document d = null)
        {
            if (d == null) d = aDoc();
            return app.TransactionManager.StartTransaction(d as _Document, name);
        }
        static public Matrix getMatrix()
        {
            return tg.CreateMatrix();
        }
        static public Matrix2d getMatrix2d()
        {
            return tg.CreateMatrix2d();
        }
        static public void addDoc(Document doc)
        {
            if (docs == null) docs = new List<Document>();
            docs.Add(doc);
        }
        static public void save(bool dep)
        {
            foreach (var item in docs)
            {
                item.Save2(dep); 
            }
        }
        static public void clearDocs()
        {
            docs.Clear();
        }
        static public Document aDoc()
        {
            return app.ActiveDocument;
        }
        static public string closeDocs()
        {
            string ffn = null;
            foreach (_Document item in app.Documents)
            {
                ffn = item.FullFileName;
                item.Close(); 
            }
            return ffn;
        }
        static public string aPath()
        {
            Document d = app.ActiveDocument;
            return InvDoc.file.p(d.FullFileName);
        }
        static public Document newDoc(string ffn, string templ = "", bool vis = true, string alias = null)
        {
            InvDoc.file.createPath(ffn);
            if (InvDoc.file.check(ffn)) return open(ffn, false, false);
            else
            {
                Document d = app.Documents.Add(DocumentTypeEnum.kPartDocumentObject, templ, vis);
                //if (!InvDoc.u.isNull(alias)) addFLib(InvDoc.file.p(ffn), alias);
                if (ffn.IndexOf("base") != -1)
                {
                    InvDoc.u.addProp(d, "Description", "base");
                }
                d.SaveAs(ffn, false);
                d.Save2();
                return d;
            }
        }
        static public PartDocument findInRef(PartDocument doc, string name)
        {
            foreach (Document item in doc.ReferencingDocuments)
            {
                if (item.DocumentType != DocumentTypeEnum.kPartDocumentObject) continue;
                PartDocument rd = (PartDocument)item;
                var def = rd.ComponentDefinition;
                var dpcs = def.ReferenceComponents.DerivedPartComponents;
                if (dpcs.Count == 0) continue;
                var dpc = dpcs[1];
                foreach (ReferenceFeature s in dpc.SolidBodies)
                {
                    if (s.Name == name) return rd;
                }
            }
            return null;
        }
        static public void setSMS(Document doc, string s)
        {
            try
            {
                var smcd = getSMCD(doc);
                if (smcd == null) return;
                var sms = u.get<SheetMetalStyle>(smcd.SheetMetalStyles, fi => fi.Name == s);
                sms.Activate();
            }
            catch (Exception)
            {
            }
        }
        static public void addSB(Document doc, string ffn, SurfaceBody sb = null, List<string> imNames = null)
        {
            if (doc.DocumentType != DocumentTypeEnum.kPartDocumentObject) return;
            PartComponentDefinition def = ((PartDocument)doc).ComponentDefinition;
            DerivedPartComponents pscs = def.ReferenceComponents.DerivedPartComponents;
            int c = pscs.Count;
            DerivedPartUniformScaleDef usd;

            if (c == 0)
            {
                usd = pscs.CreateUniformScaleDef(ffn);
                usd.ExcludeAll();
            }
            else usd = (DerivedPartUniformScaleDef)pscs[1].Definition;
            if (sb != null)
            {
                foreach (DerivedPartEntity item in usd.Solids)
                {
                    if (((SurfaceBody)item.ReferencedEntity).Equals(sb))
                    {
                        item.IncludeEntity = true;
                        break;
                    }
                    //else item.IncludeEntity = false;
                }
            }
            if (imNames != null)
            {
                foreach (DerivedPartEntity item in usd.iMateDefinitions)
                {
                    var cimd = item.ReferencedEntity as iMateDefinition;
                    if (cimd != null && imNames.Contains(cimd.Identifier)) item.IncludeEntity = true;
                }
            }
            if (c == 0)
                pscs.Add((DerivedPartDefinition)usd);
            else try
                {
                    pscs[1].Definition = (DerivedPartDefinition)usd;
                }
                catch (Exception)
                { 
                }
        }
        static public DesignProject dPr()
        {
            return app.DesignProjectManager.ActiveDesignProject;
        }
        static public void addFLib(string dir, string alias)
        {
            DesignProject pr = dPr();
            ProjectPaths ps = pr.FrequentlyUsedPaths;
            string ws = pr.WorkspacePath;
            //string n = pr.GetCustomSection("ProjectPath");
            if(!u.isNull(u.get<ProjectPath>(ps, f => f.Name == alias))) return;
            InvDoc.XMLDoc xml = new InvDoc.XMLDoc(pr.FullFileName, "InventorProject", ".ipj");
            //var el = xml.find("pathtype","Workspace");
            //string n = el.Elements().ElementAt(0).Value;
            if (dir.EndsWith("\\")) dir = dir.TrimEnd('\\');  
            ProjectPath p = pr.FrequentlyUsedPaths.Add(alias, dir);
            //dir = dir.Replace("C:", "c:");
            //p.Path = "WS\\БВ";
            //p.Path = dir;
        }
        static public Document open(string ffn, bool v = false, bool drw = true)
        {
            if (v)
            {
                silent(true);
            }
            Document doc = null;
            try
            {
                if (!InvDoc.file.check(ffn)) return null;
                if (InvDoc.file.ext(ffn) == "idw" && drw) doc = app.Documents.OpenWithOptions(ffn, nvmDrw, v);
                doc = app.Documents.OpenWithOptions(ffn, nvm, v);
            }
            catch (Exception)
            {
                return null;
            }
            if (v)
            {
                if (doc.RequiresUpdate) doc.Update2(false);
                silent(false);
            }
            return doc; 
        }
        static public void open(string path, string regex)
        {
            var files = InvDoc.file.getFiles(path, regex);
            foreach (var item in files)
            {
                open(item, false, false);
            }
        }
        static public void silent(bool f)
        {
            app.SilentOperation = f;
        }
        static public void screenUpdate(bool f)
        {
            app.ScreenUpdating = f;
        }
        static public void screenSilent(bool f)
        {
            silent(f);
            screenUpdate(!f);
        }
        static public SelectSet getSS(Document doc)
        {
            return doc.SelectSet;
        }
        static public SelectSet getSS()
        {
            return aDoc().SelectSet;
        }
        static public AssemblyComponentDefinition getACD(Document doc)
        {
            return InvDoc.Reflect.getProp<AssemblyDocument, AssemblyComponentDefinition>(doc as AssemblyDocument, "ComponentDefinition");
        }
        static public SheetMetalComponentDefinition getSMCD(Document doc)
        {
            return InvDoc.Reflect.getProp<PartDocument, SheetMetalComponentDefinition>(doc as PartDocument, "ComponentDefinition");
        }
        static public PartComponentDefinition getPCD(Document doc) 
        {
            return InvDoc.Reflect.getProp<PartDocument, PartComponentDefinition>(doc as PartDocument, "ComponentDefinition");
        }
        static public ComponentDefinition getCD(Document doc)
        {
            ComponentDefinition def = null;
            def = getACD(doc) as ComponentDefinition;
            if (def == null) def = getPCD(doc) as ComponentDefinition;
            return def;
        }
        static public iMateDefinitions getIM(Document doc)
        {
            var def = getACD(doc);
            if (def != null) return def.iMateDefinitions;
            var def1 = getPCD(doc);
            if (def1 != null) return def1.iMateDefinitions;
            return null;
        }
        static public SheetMetalComponentDefinition getSMCD()
        {
            return getSMCD(aDoc());
        }
        static public SheetMetalFeatures getSMF()
        {
            return getSMCD().Features as SheetMetalFeatures;
        }
        static public SheetMetalFeatures getSMF(Document doc)
        {
            return getSMCD(doc).Features as SheetMetalFeatures;
        }
        static public FlatPattern getFP(Document doc)
        {
            SheetMetalComponentDefinition smcd = getSMCD(doc);
            if (smcd != null && smcd.HasFlatPattern) return smcd.FlatPattern;
            return null;
        }
        static public void createPunch(string name, Dictionary<string,double> dic, ObjectCollection col, PartComponentDefinition def, double ang = 0)
        {
            string path = app.iFeatureOptions.RootPath;
            SheetMetalFeatures smf = getSMF((Document)def.Document);
            iFeatureDefinition ifd = smf.PunchToolFeatures.CreateiFeatureDefinition(path + name);
            foreach (iFeatureInput input in ifd.iFeatureInputs)
            {
                if (dic.ContainsKey(input.Name))
                {
                    ((iFeatureParameterInput)input).Value = dic[input.Name] / 10;
                }
            }
            smf.PunchToolFeatures.Add(col, ifd, ang);
        }
        static public void createSlot(Point2d insPt, PlanarSketch ps, Vector2d dir, double a, double b)
        {
            dir.ScaleBy(a);
            Point2d ep = insPt.Copy();
            Point2d sp = insPt.Copy();
            ep.TranslateBy(dir);
            dir.ScaleBy(-1);
            sp.TranslateBy(dir);
            ps.AddStraightSlotByCenterToCenter(sp, ep, b);
        }
        static public CutFeature createCut(PlanarSketch ps)
        {
            SheetMetalFeatures smf = getSMF();
            CutDefinition cd = smf.CutFeatures.CreateCutDefinition(ps.Profiles.AddForSolid(false));
            try
            { cd.SetCutAcrossBendsExtent("Толщина"); }
            catch (Exception)
            {
                cd.SetCutAcrossBendsExtent("Thickness");
            }
            return smf.CutFeatures.Add(cd);
        }
        static public ObjectCollection getPoints(List<SketchPoint> pts, Func<SketchPoint,bool> f = null)
        {
            ObjectCollection col = I.COC();
            foreach (SketchPoint item in pts)
            {
                if (f != null && f(item))
                    col.Add(item);
                else col.Add(item);
            }
            return col;
        }
        static public ObjectCollection getSketchPoints(PlanarSketch ps, Func<SketchPoint, bool> f)
        {
            ObjectCollection col = I.COC();
            foreach (SketchPoint item in ps.SketchPoints)
            {
                if (f(item))
                    col.Add(item);
            }
            return col;
        }
        static public Vector CV(double x=0, double y=0, double z=0)
        {
            return tg.CreateVector(x, y, z);
        }
        static public Vector CV(UnitVector v)
        {
            return tg.CreateVector(v.X, v.Y, v.Z);
        }
        static public Vector CV(Vector v, double s = 1)
        {
            return tg.CreateVector(v.X*s, v.Y*s, v.Z*s);
        }
        static public Vector2d CV2d(double x = 0, double y = 0)
        {
            return tg.CreateVector2d(x, y);
        }
        static public UnitVector2d CUV2d(UnitVector2d v)
        {
            return tg.CreateUnitVector2d(v.X, v.Y);
        }
        static public UnitVector2d CUV2d()
        {
            return tg.CreateUnitVector2d(1, 0);
        }
        static public Vector2d CV2d(Vector2d v)
        {
            return tg.CreateVector2d(v.X, v.Y);
        }
        static public Vector2d CV2d(Point2d s, Point2d e)
        {
            var v =  s.VectorTo(e);
            return v;
        }
        static public UnitVector2d CUV2d(double x, double y)
        {
            return tg.CreateUnitVector2d(x,y);
        }
        static public UnitVector CUV(double x, double y, double z)
        {
            return tg.CreateUnitVector(x, y, z);
        }
        static public Point CP(double x = 0, double y = 0, double z = 0)
        {
            return tg.CreatePoint(x, y, z);
        }
        static public Point2d CP2d(double x = 0, double y = 0)
        {
            return tg.CreatePoint2d(x, y);
        }
        static public Point CP(Point pt, Vector v, double scale, int tol = 3)
        {
            double [] a = new double []{};
            pt.GetPointData(ref a);
            v.ScaleBy(scale);
            return CP(a[0] + Math.Round(v.X, tol), a[1] + Math.Round(v.Y, tol), a[2] + Math.Round(v.Z, tol));
        }
        static public Point2d CP2d(Point2d pt, Vector2d v)
        {
            return tg.CreatePoint2d(pt.X + v.X, pt.Y + v.Y);
        }
        static public Point2d CP2d(SketchPoint sp)
        {
            return tg.CreatePoint2d(sp.Geometry.X, sp.Geometry.Y);
        }
        static public Point2d CP2d(Point2d sp)
        {
            return tg.CreatePoint2d(sp.X, sp.Y);
        }
        static public Point2d CP2d(double x, double y, Vector2d v)
        {
            return v.IsParallelTo(I.CV2d(1, 0), 0.01) ? tg.CreatePoint2d(x, y) : tg.CreatePoint2d(y, x);
        }
        static public Box2d Box(double sx, double sy, double ex, double ey)
        {
            Box2d b = tg.CreateBox2d();
            b.MinPoint = CP2d(sx, sy);
            b.MaxPoint = CP2d(ex, ey);
            return b;
        }
        static public Point BoxCenter(Box b)
        {
            return u.midPt(b.MinPoint, b.MaxPoint);
        }
        public static Box2d Box(Point2d pt1, Point2d pt2)
        {
            Box2d b = tg.CreateBox2d();
            b.MinPoint = pt1; b.MaxPoint = pt2;
            return b;
        }
        public static Box2d Box(DrawingView dv)
        {
            Box2d b = tg.CreateBox2d();
            var v = I.CV2d(-dv.Width * 0.5, -dv.Height * 0.5);
            b.MinPoint = I.CP2d(dv.Position, v);
            v.ScaleBy(-1);
            b.MaxPoint = I.CP2d(dv.Position, v);
            return b;
        }
        static public Box Box(Point pt1, Point pt2)
        {
            Box b = tg.CreateBox();
            b.MinPoint = CP(pt1.X, pt1.Y, pt1.Z);
            b.MaxPoint = CP(pt2.X, pt2.Y, pt2.Z);
            return b;
        }
        static public Point2d Rev(Point2d pt)
        {
            return tg.CreatePoint2d(pt.Y, pt.X);
            return tg.CreatePoint2d(pt.Y, pt.X);
        }
        static public object Pick(SelectionFilterEnum f, string prompt)
        {
            return I.app.CommandManager.Pick(f, prompt);
        }
        static public ObjectCollection COC()
        {
            return objs.CreateObjectCollection();
        }
        static public ObjectCollection COC(object ob)
        {
            ObjectCollection col = objs.CreateObjectCollection();
            col.Add(ob);
            return col;
        }
        static public ObjectCollection COC(ObjectCollection col)
        {
            if (col == null)
            return objs.CreateObjectCollection();
            col.Clear();
            return col;
        }
        static public EdgeCollection CEC(EdgeCollection col)
        {
            if (col == null) return objs.CreateEdgeCollection();
            col.Clear();
            return col;
        }
        static public FaceCollection CFC(FaceCollection col)
        {
            if (col == null) return objs.CreateFaceCollection();
            col.Clear();
            return col;
        }
        static public WorkPlane CWP(ComponentDefinition def)
        {
            var acd = def as AssemblyComponentDefinition;
            var pcd = def as PartComponentDefinition;
            return acd != null ?
                acd.WorkPlanes.AddFixed(I.CP(), I.CUV(0, 1, 0), I.CUV(0, 0, 1)) :
                pcd.WorkPlanes.AddByPlaneAndOffset(pcd.WorkPlanes[1], 0);
        }
        static public WorkPlaneProxy CWPP(AssemblyComponentDefinition adef, ComponentOccurrence occ, WorkPlane pl)
        {
            var wpl = CWP(occ.Definition);
            var ad = pl.Definition as AssemblyWorkPlaneDef;
            wpl.Adaptive = true;
            object pr;
            occ.CreateGeometryProxy(wpl, out pr);
            var wpr = pr as WorkPlaneProxy;
            adef.Constraints.AddFlushConstraint(pl, wpr, 0);
            return wpr;
        }
        static public Sheet getSheet()
        {
            DrawingDocument doc = aDoc() as DrawingDocument;
            if (doc == null) return null;
            return InvDoc.Reflect.getProp<DrawingDocument, Sheet>(doc, "ActiveSheet");
        }
        static public IEnumerable<Document> getDocs(Document asm, IEnumerable<Document> docs)
        {
            foreach (Document doc in asm.ReferencedDocuments)
            {
                //if (doc.FullDocumentName.IndexOf(path) != -1)
                    docs = u.add<Document>(docs, doc);
                if (doc.DocumentType == DocumentTypeEnum.kAssemblyDocumentObject)
                    docs = u.add<Document>(getDocs(doc, docs),doc);
            }
            //if (asm.FullDocumentName.IndexOf(path) != -1)
            docs = u.add<Document>(docs, asm as Document);
            return docs;
        }
        static public IEnumerable<T> getFiles<T>(IEnumerable<string> ffn, string ext) where T: class
        {
            foreach (var item in ffn)
            {
                yield return open(item) as T;
            }
        }
        static public IEnumerable<T> getFiles<T>(string path, string ext) where T : class
        {
            return getFiles<T>(InvDoc.file.getFiles(path, ext), ext);
        }
        static public void getStyle(string n)
        {
            DrawingDocument doc = aDoc() as DrawingDocument;
            if (doc == null) return;
            if (n == "") { dimStyle = getStyleL(); return; }
            dimStyle = u.get<DimensionStyle>(doc.StylesManager.DimensionStyles, f => f.Name == n);
        }
        static public DimensionStyle getStyleL()
        {
            DrawingDocument doc = aDoc() as DrawingDocument;
            if (doc == null) return null;
            return doc.StylesManager.ActiveStandardStyle.ActiveObjectDefaults.LinearDimensionStyle;
        }
        static public DimensionStyle getStyleR()
        {
            DrawingDocument doc = aDoc() as DrawingDocument;
            if (doc == null) return null;
            return doc.StylesManager.ActiveStandardStyle.ActiveObjectDefaults.RadialDimensionStyle;
        }
        static public DimensionStyle getStyleA()
        {
            DrawingDocument doc = aDoc() as DrawingDocument;
            if (doc == null) return null;
            return doc.StylesManager.ActiveStandardStyle.ActiveObjectDefaults.AngularDimensionStyle;
        }
        static public double multiPV(Point2d p, Vector2d v)
        {
            return p.X*v.X + p.Y*v.Y;
        }
        static public BOM getBOM(Document doc)
        {
            AssemblyComponentDefinition acd = getACD(doc);
            if (acd == null) return null;
            return acd.BOM;
        }
        static public BOM getBOM()
        {
            return getBOM(aDoc());
        }
        static public BOMView getBOMView(Document doc,int ind = 1)
        {
            BOM bom = getBOM(doc);
            if (bom == null) return null;
            return bom.BOMViews[ind];
        }
        static public BOMView getBOMView(int ind = 1)
        {
            return getBOMView(aDoc(), ind);
        }
        static public Parameters getParameters(Document doc)
        {
            Parameters ps = null;
            if (doc is PartDocument) ps = getPCD(doc).Parameters;
            else if (doc is AssemblyDocument) ps = getACD(doc).Parameters;
            return ps;
        }
        static public UserParameters getUserParameters(Document doc)
        {
            UserParameters ps = null;
            if (doc is PartDocument) ps = getPCD(doc).Parameters.UserParameters;
            else if (doc is AssemblyDocument) ps = getACD(doc).Parameters.UserParameters;
            return ps;
        }
        static public string curProjPath()
        {
            return app.DesignProjectManager.ActiveDesignProject.WorkspacePath;
        }
    }
}
