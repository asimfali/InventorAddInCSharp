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
        static public string path = "";
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
        static public Document aDoc()
        {
            return app.ActiveDocument;
        }
        static public Document open(string ffn, bool v = false)
        {
            if (InvDoc.file.ext(ffn) == "idw") return app.Documents.OpenWithOptions(ffn, nvmDrw, v);
            return app.Documents.OpenWithOptions(ffn, nvm, v);
        }
        static public void silent(bool f)
        {
            app.SilentOperation = f;
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
        static public SheetMetalComponentDefinition getSMCD()
        {
            return getSMCD(aDoc());
        }
        static public FlatPattern getFP(Document doc)
        {
            SheetMetalComponentDefinition smcd = getSMCD(doc);
            if (smcd != null && smcd.HasFlatPattern) return smcd.FlatPattern;
            return null;
        }
        static public Vector CV(double x=0, double y=0, double z=0)
        {
            return tg.CreateVector(x, y, z);
        }
        static public Vector2d CV2d(double x = 0, double y = 0)
        {
            return tg.CreateVector2d(x, y);
        }
        static public Point CP(double x = 0, double y = 0, double z = 0)
        {
            return tg.CreatePoint(x, y, z);
        }
        static public Point2d CP2d(double x = 0, double y = 0)
        {
            return tg.CreatePoint2d(x, y);
        }
        static public Box2d Box(double sx, double sy, double ex, double ey)
        {
            Box2d b = tg.CreateBox2d();
            b.MinPoint = CP2d(sx, sy);
            b.MaxPoint = CP2d(ex, ey);
            return b;
        }
        static public ObjectCollection COC()
        {
            return objs.CreateObjectCollection();
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
    }
}
