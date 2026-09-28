using System;
using System.Collections.Generic;
using System.Windows.Forms;
using System.Linq;
using System.Text;
using Inventor;
using InvDoc;
using ExtensionMethods;
using System.Xml;
using System.Xml.Linq;
using System.Text.RegularExpressions;
using InterfaceDll;
using ut = InvDoc.u;

namespace InvAddIn
{
    enum drawSide
    {
        left, right, top, bottom
    }
    class Curve
    {
        drawSide ds;
        DrawingView dv;
        double max = 0;
        List<DrawingCurve> curves = new List<DrawingCurve>();
        public Curve(DrawingView dv, drawSide ds)
        {
            this.dv = dv; this.ds = ds;
            add();
            curves.Sort(new comp());
        }
        public void add()
        {
            foreach (DrawingCurve item in dv.DrawingCurves)
            {
                if (item.ProjectedCurveType != Curve2dTypeEnum.kLineSegmentCurve2d) continue;
                switch (ds)
                {
                    case drawSide.left:
                    case drawSide.right:
                        if (max == 0) max = dv.Width;
                        if (u.isVertical(item.Segments[1]))
                        {
                            curves.Add(item); 
                        }
                        break;
                    case drawSide.top:
                    case drawSide.bottom:
                        if (max == 0) max = dv.Height;
                        if (u.isHorizontal(item.Segments[1]))
                        {
                            curves.Add(item);
                        }
                        break;
                    default:
                        break;
                }
            }
        }
        public List<DrawingCurve> filter(double val, Func<double, bool> f)
        {
            return curves.Where(e => f(u.getLenght(e))).ToList();
        }
    }
    class Dim
    {
        drawSide f, l;
        //int fnum, lnum;
        string ff, lf;
        Point2d cen;
        DimensionTypeEnum dt;
        string stName;
        public Dim(LinearGeneralDimension d, double max)
        {
            cen = d.Text.Origin;
            DrawingCurve fdc = d.IntentOne.Geometry as DrawingCurve;
            DrawingCurve ldc = d.IntentTwo.Geometry as DrawingCurve;
            dt = d.DimensionType;
            stName = d.Style.Name;
            f = getSide(fdc, u.isVertical(fdc.Segments[1]));
            l = getSide(ldc, u.isVertical(fdc.Segments[1]));
            ff = getFilter(fdc, max);
            lf = getFilter(ldc, max);
        }
        public string getFilter(DrawingCurve dc, double max)
        {
            double l = u.getLenght(dc);
            l = l / max;
            if (l > 0.7) return "70-100";
            else if (l < 0.2) return "0-20";
            else return "20-70";
        }
        public drawSide getSide(DrawingCurve dc, bool v)
        {
            DrawingView dv = dc.Parent;
            if (v)
            {
                return (dc.StartPoint.X <= dv.Position.X) ? drawSide.left: drawSide.right;
            }
            else
            {
                return (dc.StartPoint.Y <= dv.Position.Y) ? drawSide.bottom : drawSide.top;
            }
        }
    }
    class comp : IComparer<DrawingCurve>
    {
        public int Compare(DrawingCurve x, DrawingCurve y)
        {
            if (u.eq(x.StartPoint.X, y.StartPoint.X))
            {
                return x.StartPoint.Y < y.StartPoint.Y ? 1 : -1;
            }
            return x.StartPoint.X < y.StartPoint.X ? 1 : -1;
        }
    }
    class Draw
    {
        MyXML xml;
        DrawingDocument drw;
        public Draw()
        {
            xml = new MyXML("Draw");
        }
        public Draw(DrawingDocument d)
        {
            if (d == null) return;
            drw = d;
            xml = new MyXML("Draw", "Draw");
            fillXML(ref xml.elem);
            xml.save(xml.elem);
        }
        public void draw(List<string> str)
        {
            foreach (var item in str)
            {
                Document doc = I.open(item);
                if (doc != null) draw(doc);
            }
        }
        public void draw(Document doc)
        {

        }
        public void fillXML(ref XElement el)
        {
            Document doc = drw.Sheets[1].DrawingViews[1].ReferencedDocumentDescriptor.ReferencedDocument as Document;
            var props = u.getProps(doc, new string[] { "Part Number", "Description" });
            XElement xel = MyXML.addXElement("Draw", new Dictionary<string, string>() { { "Name", "test" }, { "Desc", props[1].Value.ToString() } });
            int i = 1;
            foreach (Sheet s in drw.Sheets)
            {
                fillXML(s, i, xel);
                i++;
            }
            el = xel;
        }
        public void fillXML(Sheet sh, int i, XElement el)
        {
            XElement tmp = MyXML.addXElement("Sheet", new Dictionary<string, string>() { { "Name", i.ToString() }});
            MyXML.addAtt(tmp, "w", (sh.Width * 10).ToString());
            MyXML.addAtt(tmp, "h", (sh.Height * 10).ToString());
            foreach (DrawingView dv in sh.DrawingViews)
            {
                fillXML(dv, tmp);  
            }
            el.Add(tmp);
        }
        public void fillXML(DrawingView dv, XElement el)
        {
            XElement tmp = MyXML.addXElement("View", new Dictionary<string, string>() { { "Name", dv.Name }, 
            {"x", u.round(dv.Center.X).ToString()}, {"y",u.round(dv.Center.Y).ToString()}});
            if (dv.ParentView != null) MyXML.addAtt(tmp, "Parent", dv.ParentView.Name);
            MyXML.addAtts(tmp, new Dictionary<string, string>() { { "type", dv.ViewType.ToString() }, {"scale", dv.Scale.ToString() },
            { "camera", u.round(dv.Camera.UpVector.X).ToString() + ";" + u.round(dv.Camera.UpVector.Y).ToString() + ";" + u.round(dv.Camera.UpVector.Z).ToString()}});
            if (dv.ViewType == DrawingViewTypeEnum.kStandardDrawingViewType)
            {
                Document doc = dv.ReferencedDocumentDescriptor.ReferencedDocument as Document;
                var props = u.getProps(doc, new string[] { "Part Number", "Description" });
                MyXML.addAtt(tmp, "model", props[1].Value.ToString());
                string dn = props[0].Value.ToString();
                Regex r = new Regex(@".*-(\d\d)$");
                Match m = r.Match(dn);
                if (m.Groups.Count == 2)
                {
                    MyXML.addAtt(tmp, "isp", m.Groups[1].Value);
                }
            }
            el.Add(tmp);
        }
    }
    public abstract class DimsBase
    {
        public PartFeature pf;
        public abstract void set(DrawingCurve dc, DrawingView dv);
        public abstract void addDims();
        public abstract void get();
    }
    public class DimBends
    {
        Document rDoc;
        Sheet sh;
        double t, R;
        SheetMetalComponentDefinition smcd;
        FlatPattern fp = null;
        EdgesCol eCol;
        HashSet<DrawingCurve> dcFilter = new HashSet<DrawingCurve>();
        //List<DimBend> bends = new List<DimBend>();
        List<DimsBase> dims = new List<DimsBase>();
        //List<Holes> holes = new List<Holes>();
        //List<FlatBend> fbends = new List<FlatBend>();
        //List<FlatGab> gabs = new List<FlatGab>();
        HashSet<int> numsBend = new HashSet<int>();
        //List<LinearGeneralDimension> dims = new List<LinearGeneralDimension>();
        public DimBends(Sheet sh)
        {
            this.sh = sh;
            DrawingDocument drw = (DrawingDocument)sh.Parent;
            //dims =  u.getDims(sh);
            foreach (DocumentDescriptor desc in drw.ReferencedDocumentDescriptors)
            {
                rDoc = desc.ReferencedDocument as Document;
                if (rDoc.DocumentType != DocumentTypeEnum.kPartDocumentObject) return;
                smcd = I.getSMCD(rDoc);
                R = (double)u.getParameter(rDoc, "Радиус_гибки").Value;
                t = (double)smcd.Thickness.Value;
                if (sh.DrawingViews[1].IsFlatPatternView)
                    findFlatBends();
                else findModelBends(); 
            }
        }
        public void findFlatBends()
        {
            fp = I.getFP(rDoc);
            eCol = new EdgesCol(fp);
            //var gx = new FlatGab(I.CV(1, 0, 0), eCol);
            //var gy = new FlatGab(I.CV(0, 1, 0), eCol);
            //dims.Add(gx); dims.Add(gy);
            foreach (FlatBendResult fbr in fp.FlatBendResults)
            {
                int bo = 0;
                BendOrderSourceTypeEnum boe;
                Edge e = fbr.Edge;
                fbr.GetBendOrder(out bo, out boe);
                if (numsBend.Contains(bo)) continue;
                numsBend.Add(bo);
                dims.Add(new FlatBend(eCol.bf, e, eCol, fbr.IsDirectionUp));   
            }
            //foreach (var item in u.gets<Face>(fp.SurfaceBodies[1].Faces,
            //    fi => fi.SurfaceType == SurfaceTypeEnum.kCylinderSurface))
            //{
            //    var h = new Holes(item, eCol);
            //    if (dims.Contains(h, new FeatComp())) continue;
            //    dims.Add(h);
            //}
        }
        public void findModelBends()
        {
            foreach (Bend item in smcd.Bends)
            {
                if (item.FrontFaces[1].SurfaceType == SurfaceTypeEnum.kCylinderSurface &&
                    u.eq(((Cylinder)item.FrontFaces[1].Geometry).Radius, R + t))
                    dims.Add(new DimBend(item.FrontFaces[1], t));
                else if (item.BackFaces[1].SurfaceType == SurfaceTypeEnum.kCylinderSurface &&
                    u.eq(((Cylinder)item.BackFaces[1].Geometry).Radius, R + t))
                    dims.Add(new DimBend(item.BackFaces[1], t));
            }
        }
        //public void findModelBends()
        //{
        //    foreach (Face item in smcd.SurfaceBodies[1].Faces)
        //    {
        //        if (item.SurfaceType == SurfaceTypeEnum.kCylinderSurface &&
        //            u.eq(((Cylinder)item.Geometry).Radius, R + t))
        //            dims.Add(new DimBend(item, t));
        //    }
        //}
        public void fill()
        {
            foreach (DrawingView dv in sh.DrawingViews)
            {
                if (dv.IsFlatPatternView) fillFlat(dv);
                else fill(dv);
            }
        }
        public void addDims()
        {
            foreach (DimsBase item in dims)
            {
                item.addDims();
            }
            var g = new Gabs(sh.DrawingViews[1]);
            g.draw();
            if (sh.DrawingViews.Count > 1)
            {
                var ng = new Gabs(sh.DrawingViews[2]);
                ng.draw();
            }
        }
        public void fill(DrawingView dv)
        {
            foreach (DrawingCurve dc in dv.DrawingCurves)
            {
                if (!dcFilter.Contains(dc))
                    addRadius(dc, dv);
                foreach (DimsBase item in dims)
                {
                    item.set(dc, dv);
                }
            }
        }
        public void fillFlat(DrawingView dv)
        {
            //Vector2d vec = I.CV2d(1, 0);
            foreach (DrawingCurve dc in dv.DrawingCurves)
            {
                foreach (DimsBase item in dims)
                {
                    item.set(dc, dv);
                }
                if (dc.EdgeType == DrawingEdgeTypeEnum.kBendDownEdge ||
                    dc.EdgeType == DrawingEdgeTypeEnum.kBendUpEdge)
                {
                    //var v = dc.StartPoint.VectorTo(dc.EndPoint);
                    //if (v.IsParallelTo(vec) || v.IsPerpendicularTo(vec)) continue;
                    addCL(dc, dv, false);
                    addCL(dc, dv, true);
                }
            }
            //foreach (var item in u.gets<DrawingCurve>(dv.DrawingCurves, fi => u.getLenght(fi) > 5))
            //{
            //    addCL(item, dv);
            //}
            var curves = u.gets<DrawingCurve>(dv.DrawingCurves, fi => !dcFilter.Contains(fi) &&
            fi.ProjectedCurveType == Curve2dTypeEnum.kCircleCurve2d);
            foreach (var item in curves.GroupBy(el => ((Circle)((Edge)el.ModelGeometry).Geometry).Radius))
            {
                foreach (var x in item.GroupBy(el => el.CenterPoint.X))
                {
                    Clines cl1 = new Clines(dv, x); dims.Add(cl1);
                }
                foreach (var y in item.GroupBy(el => el.CenterPoint.Y))
                {
                    Clines cl1 = new Clines(dv, y); dims.Add(cl1);
                } 
            }          
        }
        public void addCL(DrawingCurve dc, DrawingView dv, bool rev)
        {
            Clines cl = new Clines(); cl.set(dc, dv); cl.add(dcFilter, rev);
            cl.getDCs(ref dcFilter);
            dims.Add(cl);
        }
        public void addRadius(DrawingCurve dc1, DrawingView dv)
        {
            var g = dc1.Segments[1].Geometry as Arc2d;
            if (g == null) return;
            if (g.Radius < 2) return;
            var dcs = u.gets<DrawingCurve>(dv.DrawingCurves, fi => fi.ProjectedCurveType == Curve2dTypeEnum.kCircularArcCurve2d &&
            u.eq(((Arc2d)fi.Segments[1].Geometry).Center, g.Center));
            if (dcs.Count() != 2) return;
            DrawingCurve dc = dcs.ElementAt(0), dc2 = dcs.ElementAt(1);
            dcFilter.Add(dc); dcFilter.Add(dc2);
            Arc2d a1 = (Arc2d)dc.Segments[1].Geometry, a2 = (Arc2d)dc2.Segments[1].Geometry;
            if (a1.Radius > a2.Radius) dc = dc2;
            GeometryIntent i3 = sh.CreateGeometryIntent(dc);
            Point2d mp = u.midPt(g.StartPoint, g.EndPoint);
            I.getStyle("Текст**");
            sh.DrawingDimensions.GeneralDimensions.AddRadius(mp, i3, DimensionStyle: I.dimStyle); 
        }
    }
    public class Clines : DimsBase
    {
        DrawingView dv;
        DrawingCurve dc;
        const double d = 1.5;
        Projection prv, prn;
        List<EdgePoint> curves = new List<EdgePoint>();
        Vector2d v, n;
        CLine cl;
        public Clines()
        {
        }
        public Clines(DrawingView dv, IEnumerable<DrawingCurve> ds)
        {
            this.dv = dv;
            foreach (var item in ds)
            {
                curves.Add(get(item));
            }
            cl = new CLine(curves, dv, false, false);
        }
        public void add(HashSet<DrawingCurve> f, bool rev)
        {
            v = dc.StartPoint.VectorTo(dc.EndPoint); n = I.CV2d(v);
            u.normal(n, null);
            if (rev) n.ScaleBy(-1);
            prv = new Projection(v, false); prn = new Projection(n, false);
            prv.set(dc.StartPoint, dc.EndPoint);
            u.setDist(n, d * dv.Scale);
            var ptn1 = I.CP2d(dc.StartPoint, n);
            //n.ScaleBy(-1);
            var ptn2 = I.CP2d(dc.EndPoint, n);
            prn.set(ptn1, ptn2);
            //u.addText(dv, "min", prn.pts[0].pt);
            //u.addText(dv, "max", prn.pts[1].pt);
            fill(f);
            cl = new CLine(curves, dv, false, false);
        }
        public override void addDims()
        {
            cl.addDim();
        }
        public void getDCs(ref HashSet<DrawingCurve> ds)
        {
            foreach (var item in curves)
            {
                ds.Add(item.dc);
            }
        }
        public override void get()
        {

        }
        public EdgePoint get(DrawingCurve dc)
        {
            var ep = new EdgePoint(dc.ModelGeometry as Edge);
            ep.dc = dc;
            return ep;
        }

        public void fill(HashSet<DrawingCurve> f)
        {
            foreach (DrawingCurve item in u.gets<DrawingCurve>(dv.DrawingCurves, fi =>
            fi.ProjectedCurveType == Curve2dTypeEnum.kCircleCurve2d && !f.Contains(fi)))
            {
                Point2d cen = item.CenterPoint;
                bool c1 = prn.contains(cen), c2 = prv.contains(cen);
                if (c1 && c2)
                {
                    EdgePoint ep = get(item);
                    curves.Add(ep);
                    //u.addText(dv, "pt", cen);
                }
            }
        }

        public override void set(DrawingCurve dc, DrawingView dv)
        {
            this.dv = dv; this.dc = dc;
        }
    }
    public class FeatComp : IEqualityComparer<DimsBase>
    {
        public bool Equals(DimsBase x, DimsBase y)
        {
            if (x.pf == null || y.pf == null) return false;
            return x.pf.Equals(y.pf);
        }

        public int GetHashCode(DimsBase obj)
        {
            if (obj.pf == null) return obj.GetHashCode();
            return obj.pf.GetHashCode();
        }
    }
    public class Holes : DimsBase
    {
        EdgesCol eCol;
        readonly DimsData dd;
        //List<DrawingCurve> curves = new List<DrawingCurve>();
        bool close = false;
        DrawingView dv = null;
        public Holes(Face f, EdgesCol col)
        {
            eCol = col;
            dd = eCol.add(f);
            base.pf = f.CreatedByFeature;
            if (pf is ReferenceFeature)
            {
                ReferenceFeature rf = pf as ReferenceFeature;
                Box num = null;
                SurfaceBody rsb = null;
                if (rf.Name.StartsWith("Элемент"))
                {
                    FlatPattern fp = rf.Parent as FlatPattern ;
                    var smcd = fp.Parent;
                    rf = u.get<PartFeature>(smcd.Features, fi => fi is ReferenceFeature &&
                    (u.eq(fi.RangeBox.MinPoint, rf.RangeBox.MinPoint) &&
                    u.eq(fi.RangeBox.MaxPoint, rf.RangeBox.MaxPoint))) as ReferenceFeature;
                    
                }
                else
                {
                    rsb = rf.ReferencedEntity as SurfaceBody;
                }
                Face f2 = u.get<Face>(rsb.Faces, fi => u.eq(fi.Evaluator.RangeBox.MinPoint, num.MinPoint) &&
                u.eq(fi.Evaluator.RangeBox.MaxPoint, num.MaxPoint));
                base.pf = f2.CreatedByFeature;
            }
            get();
        }
        public Box getNum(Face f, Faces fs)
        {
            foreach (Face item in fs)
            {
                if (item.Equals(f)) return item.Evaluator.RangeBox;
            }
            return null;
        }
        public override void get()
        {
            if (base.pf.Name.ToLower().StartsWith("круг")) close = true;
        }
        public override void set(DrawingCurve dc, DrawingView dv)
        {
            if (this.dv == null) this.dv = dv;
            var e = dc.ModelGeometry as Edge;
            dd.check(dc);
            
        }
        public override void addDims()
        {
            dd.addDims(dv, close);
            int num = 1;
            foreach (var item in dd.clines)
            {
                var dims = item.addDim();
                if (dims == null || dims.Count == 0) continue;
                var dl = dims[0].DimensionLine as LineSegment2d;
                var v = dl.Direction.AsVector();
                var n = I.CV2d(v);
                u.normal(n, null);
                addDim(dims[0].IntentOne, n, num);
                //addDim(dims[dims.Count - 1].IntentOne, n, num);
                num++;
            }
        }
        public void addDim(GeometryIntent gi, Vector2d v, int num)
        {
            var vm1 = u.viewToModel(v, dv);
            var dd1 = eCol.get(vm1);
            if (dd1 == null) return;
            var pt1 = u.ptIntent(gi);
            var pt2 = I.CP2d(pt1, v);
            var ep = dd1.getMin(pt1, pt2, dv);
            if (ep == null) return;
            DrawingCurve dc = ep.getDC(dv, fi => fi.ProjectedCurveType == Curve2dTypeEnum.kLineSegmentCurve2d &&
            fi.CurveType == CurveTypeEnum.kLineSegmentCurve);
            if (dc == null) return;
            u.addDim(dv, gi, dc);
        }
    }
    public class CLine
    {
        public Centerline l;
        bool dims = false, cl = false;
        DrawingView dv;
        Vector2d v, n;
        List<EdgePoint> eps = new List<EdgePoint>();
        ObjectCollection col;
        public CLine(IEnumerable<EdgePoint> pts, DrawingView dv, bool d, bool close)
        {
            dims = d;
            this.dv = dv;
            cl = close;
            eps.AddRange(pts);
            fill();
            if (col.Count <= 1) return;
            l = dv.Parent.Centerlines.Add(col, Closed: cl);
        }
        public void fill()
        {
            col = I.COC();
            foreach (EdgePoint item in eps)
            {
                if (item.dc == null) continue;
                var cm = findCM(item.dc);
                object i = item.dc;
                if (cm != null) i = cm;
                var gi = dv.Parent.CreateGeometryIntent(i);
                col.Add(gi);
            }
        }
        public Centermark findCM(DrawingCurve dc)
        {
            foreach (Centermark cm in dv.Parent.Centermarks)
            {
                var gi = cm.AttachedEntity as GeometryIntent;
                DrawingCurve c = gi?.Geometry as DrawingCurve;
                if (c != null && c.Equals(dc)) return cm;
            }
            return null;
        }
        public List<LinearGeneralDimension> addDim()
        {
            List<LinearGeneralDimension> r = new List<LinearGeneralDimension>();
            if (col.Count == 0) return null;
            Centermark cmb = null;
            int i = 0;
            if (!dims) return null;
            if (cl) return null;
            foreach (Centermark cm in l.FitPoints)
            {
                i++;
                if (i == 1)
                {
                    cmb = cm;
                    continue;
                }
                if (v == null)
                {
                    v = cmb.Position.VectorTo(cm.Position);
                    n = I.CV2d(v); u.normal(n, null);
                }
                LinearGeneralDimension dim = u.addDim(dv, cmb, cm);
                r.Add(dim);
                cmb = cm;
                move(dim, 1.4);
            }
            return r;
        }
        public void move(LinearGeneralDimension dim, double d)
        {
            u.setDist(n, d);
            dim.Text.Origin = I.CP2d(dim.Text.Origin, n);
        }
    }
    public class EdgesCol
    {
        List<DimsData> dds = new List<DimsData>();
        readonly FlatPattern fp;
        readonly SurfaceBody sb;
        readonly Plane pl;
        readonly public Face bf;
        //public Point cen;
        //public double diag;
        public UnitVector bn;
        public EdgesCol(FlatPattern f)
        {
            fp = f;
            bf = fp.TopFace;
            pl = bf.Geometry as Plane;
            bn = pl.Normal;
            sb = fp.SurfaceBodies[1];
            //cen = u.midPt(bf.Evaluator.RangeBox.MinPoint, bf.Evaluator.RangeBox.MaxPoint);
            //diag = sb.RangeBox.MinPoint.DistanceTo(sb.RangeBox.MaxPoint);
        }
        public DimsData add(Edge e, Func<Edge, bool> f)
        {
            var tmp = get(e);
            if (tmp != null) return tmp;
            var es = u.gets<Edge>(bf.Edges, fi => f(fi));
            var dd = new DimsData(es, bn);
            dd.setDirs(e);
            dds.Add(dd);
            return dd;
        }
        public DimsData add(Face f)
        {
            PartFeature pf = f.CreatedByFeature;
            var tmp = get(pf);
            if (tmp != null) return tmp;
            var fs = u.gets<Face>(sb.Faces, fi => fi.CreatedByFeature.Equals(pf));
            var es = fs.SelectMany(fi => fi.Edges.Cast<Edge>()).Where(fi => u.eq(pl.DistanceTo(fi.PointOnEdge), 0));
            var dd = new DimsData(es, bn, pf);
            dd.getDirs();
            dd.fill();
            dds.Add(dd);
            return dd;
        }
        public DimsData get(PartFeature feat)
        {
            var tmp = u.get<DimsData>(dds, fi => fi.pf != null && fi.pf.Equals(feat));
            return tmp;
        }
        public DimsData get(Edge e)
        {
            var v = u.getDir(e);
            return get(v);
        }
        public DimsData get(Vector v)
        {
            var tmp = u.get<DimsData>(dds, fi => fi.v != null && fi.v.IsParallelTo(v, 0.01));
            return tmp;
        }
    }
    public class DimsData
    {
        List<Edge> es = new List<Edge>();
        List<Vector> vecs = new List<Vector>();
        public List<EdgePoint> eps = new List<EdgePoint>();
        public List<CLine> clines = new List<CLine>();
        readonly UnitVector bn;
        readonly public PartFeature pf;
        public Vector v, n;
        EdgePoint bp = null;
        public DimsData(IEnumerable<Edge> edges, UnitVector bn, PartFeature pf = null)
        {
            es.AddRange(edges);
            this.bn = bn;
            if (pf != null) this.pf = pf;
        }
        public void add(IEnumerable<Edge> edges)
        {
            foreach (Edge item in edges)
            {
                if (!es.Contains(item)) es.Add(item);
            }
        }
        public bool check(DrawingCurve dc)
        {
            foreach (var item in eps)
            {
                if (item.check(dc)) 
                    return true;
            }
            return false;
        }
        public void setDirs(Edge e)
        {
            v = u.getDir(e);
            n = u.normal(bn, v);
        }
        public void setDirs(Vector norm)
        {
            n = norm;
            v = u.normal(bn, n);
        }
        public void getDirs()
        {
            for (int i = 0; i < es.Count; i++)
            {
                bp = new EdgePoint(es[i]);
                eps.Add(bp);
                for (int j = 0; j < es.Count; j++)
                {
                    if (i == j) continue;
                    EdgePoint tp = new EdgePoint(es[j]);
                    eps.Add(tp);
                    Vector v = bp.get(tp.cen);
                    vecs.Add(v);
                }
                if (vecs.Count > 0)
                {
                    //if (check(vecs[0])) return;
                    for (int j = 0; j < vecs.Count; j++)
                    {
                        if (check(vecs[j])) return;
                    }
                }
                eps.Clear();
                vecs.Clear();
            }
        }
        public void fill(Edge be)
        {
            foreach (Edge item in es)
            {
                var ep = new EdgePoint(item, be, bn);
                eps.Add(ep);
            }
            eps.Sort();
        }
        public void fill()
        {
            if (bp == null) return;
            foreach (var item in eps)
            {
                item.set(bp, v, n);
            }
            eps.Sort();
        }
        public EdgePoint getMin(Edge e)
        {
            Vector v1 = u.normal(e, eps[0].cen, bn), v2 = u.normal(e, eps[eps.Count - 1].cen, bn);
            return v1.Length < v2.Length ?eps[0]: eps[eps.Count-1];
        }
        public EdgePoint getMin(Point2d p1, Point2d p2, DrawingView dv)
        {
            var dc1 = eps[0].getDC(dv, fi => fi.ProjectedCurveType == Curve2dTypeEnum.kLineSegmentCurve2d &&
            fi.CurveType == CurveTypeEnum.kLineSegmentCurve);
            var dc2 = eps[eps.Count - 1].getDC(dv, fi => fi.ProjectedCurveType == Curve2dTypeEnum.kLineSegmentCurve2d &&
            fi.CurveType == CurveTypeEnum.kLineSegmentCurve);
            if (dc1 == null || dc2 == null) return null;
            Vector2d v1 = u.normal(p1, p2, u.getCenPoint(dc1)), v2 = u.normal(p1, p2, u.getCenPoint(dc2));
            return v1.Length < v2.Length ? eps[0] : eps[eps.Count - 1];
        }
        public bool check(Vector v)
        {
            bool perp = false;
            int par = 0;
            Vector tn = null;
            foreach (Vector item in vecs)
            {
                if (v.IsEqualTo(item, 0.01)) continue;
                else if (item.IsParallelTo(v, 0.01)) { par++; }
                else if (item.IsPerpendicularTo(v, 0.01)) {tn = item; perp = true; }
            }
            if (par != 0 || perp) 
            { 
                this.v = v; this.n = u.normal(bn, v);

                return true; 
            }
            if (par == vecs.Count-1) 
            { 
                this.v = v; this.n = u.normal(bn, v); 
                return true; 
            }
            return false;
        }
        public void addDims(DrawingView dv, bool close)
        {
            addDims(dv, close, fi => fi.x);
            addDims(dv, close, fi => fi.y);
        }
        public void addDims(DrawingView dv, bool close, Func<EdgePoint, double> f)
        {
            bool first = true;
            foreach (var g in eps.GroupBy(fi => f(fi)))
            {
                if (g.Count() < 2) continue;
                CLine cl = new CLine(g, dv, first, close);
                clines.Add(cl);
                first = false;
            }
        }
    }
    public class EdgePoint : IComparable<EdgePoint>
    {
        public Edge e;
        public Point cen;
        public DrawingCurve dc = null;
        public double x, y;
        public Vector v, n;
        public EdgePoint(Edge ed, Edge be = null, UnitVector bn = null)
        {
            e = ed;
            cen = u.getCenPoint(e);
            if (be != null && bn != null)
            {
                n = u.normal(be, cen, bn);
                v = u.getDir(ed);
                x = 0; y = u.round(n.DotProduct(I.CV(1,1,1)));
            }
        }
        public Vector get(Point bp)
        {
            return bp.VectorTo(cen);
        }
        public void set(EdgePoint bp, Vector v, Vector n)
        {
            Vector dir = get(bp.cen);
            this.x = u.round(dir.DotProduct(v)); 
            this.y = u.round(dir.DotProduct(n));
        }
        public int CompareTo(EdgePoint other)
        {
            if (u.eq(x, other.x))
                return y < other.y ? -1 : y > other.y ? 1 : 0;
            return x < other.x ? -1 : x > other.x ? 1 : 0;
        }
        public bool check(DrawingCurve dc)
        {
            Edge me = dc.ModelGeometry as Edge;
            if (me.Equals(e))
            {
                this.dc = dc;
                return true;
            }
            return false;
        }
        public DrawingCurve getDC(DrawingView dv, Func<DrawingCurve, bool> f)
        {
            if (dc != null) return dc;
            dc = u.getDC(dv, e, fi => f(fi));
            return dc;
        }
    }
    public class FlatGab : DimsBase
    {
        EdgesCol eCol;
        Edge e1, e2;
        Face bf;
        Vector v;
        DrawData dd = null;
        public FlatGab(Vector dir, EdgesCol col)
        {
            eCol = col;
            v = dir;
            get();
        }
        public override void get()
        {
            double d = 0.5;
            var e = u.get<Edge>(eCol.bf.Edges, fi => fi.GeometryType == CurveTypeEnum.kLineSegmentCurve &&
            u.getLenght(fi) > d && u.getDir(fi).IsPerpendicularTo(v, 0.01));
            var dd = eCol.add(e, fi => fi.GeometryType == CurveTypeEnum.kLineSegmentCurve &&
            u.getDir(fi).IsParallelTo(u.getDir(e)));
            dd.fill(e);
            e1 = dd.eps[0].e; e2 = dd.eps[dd.eps.Count - 1].e;
        }
        public override void set(DrawingCurve dc, DrawingView dv)
        {
            Edge e = (Edge)dc.ModelGeometry;
            if (dd == null) dd = new DrawData(dv, bf);
            if (e1 != null && e1.Equals(e)) 
                dd.set(dc, 1);
            if (e2 != null && e2.Equals(e)) 
                dd.set(dc, 2);
        }
        public override void addDims()
        {
            var d = dd.addDim(false, "", 4);
            
        }
    }
    public class FlatBend : DimsBase
    {
        EdgesCol eCol;
        Face bf;
        bool up = true;
        Edge bendEdge, e2;
        DrawData dd = null;
        public FlatBend(Face f, Edge e, EdgesCol col, bool up)
        {
            eCol = col;
            bendEdge = e;
            this.up = up;
            get();
        }
        public override void get()
        {
            double d = 0.5;
            var dd = eCol.add(bendEdge, fi => fi.GeometryType == CurveTypeEnum.kLineSegmentCurve &&
            u.getLenght(fi) > d && u.getDir(fi).IsParallelTo(u.getDir(bendEdge)));
            dd.fill(bendEdge);
            e2 = dd.getMin(bendEdge).e;
        }
        public override void set(DrawingCurve dc, DrawingView dv)
        {
            //if (dc.CurveType == CurveTypeEnum.kBSplineCurve) return;
            if (!(dc.ModelGeometry is Edge)) return;
            Edge e = (Edge)dc.ModelGeometry;
            if (dd == null) dd = new DrawData(dv, bf);
            if (bendEdge != null && bendEdge.Equals(e)) 
                dd.set(dc, 1);
            if (e2 != null && e2.Equals(e))
                dd.set(dc, 2);
        }
        public override void addDims()
        {
            dd.dir = up ? 1 : 2;
            dd.addDim(true, "Гибы", 1.2);
        }
    }
    public class ordDraw : IEquatable<ordDraw>
    {
        DrawingView dv;
        object c1, c2;
        double offset = 0.7;
        string txt = "";
        DrawingCurve d;
        public ordDraw(DrawingView dv, object c1, object c2)
        {
            this.c1 = c1; this.c2 = c2; this.dv = dv;
        }
        public void setDC(DrawingCurve dc)
        {
            d = dc;
        }
        public void setTxt(double d, int c)
        {
            txt = c.ToString() + "x" + u.round(d * 10, 1).ToString() + "=";
        }
        public void setOffset(double d)
        {
            this.offset += d;
        }
        public LinearGeneralDimension addDim(List<LinearGeneralDimension> dims, Vector2d dir = null)
        {
            DrawBox db = new DrawBox(dv);
            db.set(c1); db.set(c2);
            db.setPr();
            var v = db.getDir();
            if (d != null) db.transMP(d, offset);
            GeometryIntent gi1 = db.getGI(0), gi2 = db.getGI(1);
            if (check(gi1, dims) && check(gi2, dims)) return null;
            if (dir != null)
            {
                double norm = u.dotProduct(dir, v);
                if (norm < 0)
                {
                    var v1 = I.CV2d(v); v.ScaleBy(-1);
                    var vec = db.mp.VectorTo(dv.Center);
                    var dist = u.dotProduct(v1, vec);
                    v1.ScaleBy(dist * 2);
                    db.mp.TranslateBy(v1);
                }
            }
            var dim = u.addDim(dv, gi1, gi2, v, db.mp);
            if (txt != "") dim.Text.FormattedText = txt + @"<DimensionValue/>";
            return dim;
        }
        public bool check(GeometryIntent gi, List<LinearGeneralDimension> dims)
        {
            foreach (var item in dims)
            {
                if (item.IntentOne != null && item.IntentOne.Geometry.Equals(gi.Geometry)) return true;
                else if (item.IntentTwo != null && item.IntentTwo.Geometry.Equals(gi.Geometry)) return true;
            }
            return false;
        }
        public bool Equals(ordDraw other)
        {
            if (c1.Equals(other.c1) || c2.Equals(other.c1) &&
                (c1.Equals(other.c2) || c2.Equals(other.c2))) return true;
            return false;
        }
    }
    public class clDraw
    {
        DrawingView dv;
        Centerline cl;
        readonly Vector2d bv, bn;
        Corners corners;
        List<Centermark> ls = new List<Centermark>();
        HashSet<Centerline> clines = new HashSet<Centerline>();
        public List<ordDraw> ord = new List<ordDraw>();
        public clDraw(DrawingView dv, Centerline cl)
        {
            this.dv = dv; this.cl = cl;
            corners = new Corners(dv);
            bv = cl.StartPoint.VectorTo(cl.EndPoint);
            bv.Normalize();
            bn = I.CV2d(bv); u.normal(bn, null);
            fill(cl);
        }
        public void fill(Centerline cl)
        {
            clines.Add(cl);
            foreach (var item in cl.FitPoints)
            {
                var cm = item as Centermark;
                if (cm == null) continue;
                if (ls.Contains(cm)) continue;
                ls.Add(cm);
                if (cm.Centerlines.Count > 1)
                {
                    var cl1 = u.get<Centerline>(cm.Centerlines, fi => !fi.Equals(cl));
                    fill(cl1);
                }
            }   
        }
        public void addDims(Func<Point2d, double> gr, bool x)
        {
            Centermark bcm = null;
            Centermark cm = null;
            Vector2d v = null;
            direct dir;
            bool arr = false;
            double dist = 0;
            if (clines.Count <= 1) return;
            var l = u.gets<Centermark>(ls, fi => true).GroupBy(el => gr(el.Position)).
                OrderBy(el => el.Key);
            foreach (var item in l)
            {
                bcm = null;
                arr = checkArray(item, out dist);
                foreach (var d in item)
                {
                    if (bcm == null) { bcm = d; continue; }
                    v = bcm.Position.VectorTo(d.Position);
                    var cl1 = getCL(bcm, v);
                    var cl2 = getCL(d, v);
                    ord.Add(new ordDraw(dv, cl1, cl2));
                    if (cm == null) {
                        dir = x ? direct.Left : direct.Bottom;
                        corners.add(bcm, v, dir); cm = d;
                    }
                    if (arr)
                    {
                        dist /= dv.Scale;
                        break;
                    }
                }
                if (arr)
                {
                    ordDraw od = new ordDraw(dv, item.First(), item.Last());
                    od.setOffset(0.7); od.setTxt(dist, item.Count() - 1);
                    ord.Add(od);
                }
            }
            var v1 = I.CV2d(v); v1.ScaleBy(-1);
            dir = x ? direct.Right : direct.Up;
            corners.add(cm, v1, dir);
        }
        public bool checkArray(IEnumerable<Centermark> pts, out double d)
        {
            Vector2d v = null;
            Centermark cm = null;
            d = 0;
            if (pts.Count() < 3) return false;
            foreach (var item in pts)
            {
                if (cm == null) { cm = item; continue; }
                if (v == null) { v = item.Position.VectorTo(cm.Position); cm = item; continue; }
                var v1 = item.Position.VectorTo(cm.Position);
                if (!u.eq(v1.Length, v.Length)) return false;
                d = v.Length;
                cm = item;
            }
            return true;
        }
        public void addDim(Centermark cm, bool n, bool rev)
        {
            Vector2d v = n ? bn : bv;
            v.Normalize();
            if (rev) v.ScaleBy(-1);
            corners.add(cm, v, direct.Right);
        }
        public IEnumerable<Centermark> addDim()
        {
            var pts = u.gets<Centermark>(cl.FitPoints, fi => true);
            pts = ordCM(pts);
            if (clines.Count > 1) return null;
            Centermark cm = null; double d = 0;
            bool arr = checkArray(pts, out d);
            int count = cl.FitPoints.Count;
            if (count < 4) arr = false;
            foreach (var pt in pts)
            {
                if (cm == null) { cm = pt; continue; }
                Centermark cm1 = pt;
                ord.Add(new ordDraw(dv, cm, cm1));
                if (arr)
                {
                    d /= dv.Scale;
                    break;
                }
                cm = cm1;
            }

            if (arr)
            {
                ordDraw od = new ordDraw(dv, pts.First(), pts.Last());
                od.setOffset(0.7); od.setTxt(d, count- 1);
                ord.Add(od);
            }
            return pts;
        }
        public IEnumerable<Centermark> ordCM(IEnumerable<Centermark> vals)
        {
            if (vals.Count() <= 2) return vals;
            var e1 = vals.First(); var e2 = vals.Last();
            var v = e1.Position.VectorTo(e2.Position);
            v.Normalize();
            return vals.OrderBy(el => u.dotProduct(v, el.Position));
        }
        public object getCL(Centermark cm, Vector2d v)
        {
            foreach (Centerline item in cm.Centerlines)
            {
                var v1 = item.StartPoint.VectorTo(item.EndPoint);
                if (v1.IsPerpendicularTo(v, 0.01)) return item;
            }
            return cm;
        }
        public DrawingCurve getDC(Centermark cm,Vector2d v)
        {
            var c = u.gets<DrawingCurve>(dv.DrawingCurves, fi => check(fi, v)).Where(
                el => u.dotProduct(v, el.StartPoint) >= 0).OrderBy(
                el => u.dotProduct(v, el.StartPoint));
            return c.FirstOrDefault();
        }
        public bool check(DrawingCurve dc, Vector2d v)
        {
            if (dc.ProjectedCurveType != Curve2dTypeEnum.kLineSegmentCurve2d) return false;
            var v1 = dc.StartPoint.VectorTo(dc.EndPoint);
            return v.IsPerpendicularTo(v1, 0.01);
        }
        public bool check(direct d, List<direct> ds)
        {
            foreach (direct item in ds)
            {
                if (item == d) return true;
            }
            return false;
        }
        public void drawDims(List<direct> directs)
        {
            Vector2d dir = null;
            if (directs.Count > 1)
            {
                var d1 = directs[0]; var d2 = directs[1];
                if (check(direct.Bottom, directs) && check(direct.Left, directs)) dir = I.CV2d(-1, -1);
                else if (check(direct.Bottom, directs) && check(direct.Right, directs)) dir = I.CV2d(1, -1);
                else if (check(direct.Up, directs) && check(direct.Left, directs)) dir = I.CV2d(-1, 1);
                else if (check(direct.Up, directs) && check(direct.Right, directs)) dir = I.CV2d(1, 1);
            }
            List<LinearGeneralDimension> dims = new List<LinearGeneralDimension>();
            foreach (var item in ord)
            {
                var d = item.addDim(dims, dir);
                if (d != null) dims.Add(d);
            }
            dims.Clear();
            foreach (var item in corners.ord)
            {
                item.addDim(dims);
            }
        }
        public void setDC(DrawingCurve dc)
        {
            foreach (var item in ord)
            {
                item.setDC(dc);
            }
        }
        public void draw()
        {
            addDims(fi => u.round(fi.X), false);
            addDims(fi => u.round(fi.Y), true);
            var pts = addDim();
            if (pts != null) ls = pts.ToList();
            addDim(ls[0], false, false);
            addDim(ls[0], true, true);
            addDim(ls[ls.Count - 1], false, false);
            addDim(ls[ls.Count - 1], true, true);
            corners.addDims();
            var cp = u.get<CornerPt>(corners.lst, fi => fi.checkPar(bn));
            if (cp != null) setDC(cp.dc);
            drawDims(corners.filter);
        }
    }
    public enum direct { Left, Right, Up, Bottom}
    public class Corners
    {
        readonly DrawingView dv;
        public List<CornerPt> lst = new List<CornerPt>();
        public List<ordDraw> ord = new List<ordDraw>();
        public List<direct> filter = new List<direct>();
        public Corners(DrawingView dv)
        {
            this.dv = dv;
        }
        public void add(Centermark cm, Vector2d v, direct dir)
        {
            double l;
            var dc = u.findMinDC(dv, cm.Position, v,
                fi => fi.ProjectedCurveType == Curve2dTypeEnum.kLineSegmentCurve2d &&
                fi.EdgeType == DrawingEdgeTypeEnum.kUnknownEdge,
                (a, b) => a < b, out l);
            if (dc != null)
            {
                //u.addText(dv, "pt1", dc.StartPoint);
                var cp = new CornerPt(dc, cm, l, dir);
                if (contains(cp)) return;
                lst.Add(cp);
            }
        }
        public bool contains(CornerPt cp)
        {
            foreach (var item in lst)
            {
                if (item.octant == cp.octant) return true;
            }
            return false;
        }
        public void addDims()
        {
            if (lst.Count > 1 && lst[0].di != lst[1].di) addDimsXY();
            else addOrdDims();
        }
        public void addDimsXY()
        {
            MyForm F = new MyForm("DrawDimsInterface.xml", "Тест");
            F.f.ShowDialog();
            bool left = F.chks[0].Checked;
            bool right = F.chks[1].Checked;
            bool up = F.chks[2].Checked;
            bool bottom = F.chks[3].Checked;
            F.f.Close();
            if (left) filter.Add(direct.Left);
            if (right) filter.Add(direct.Right);
            if (up) filter.Add(direct.Up);
            if (bottom) filter.Add(direct.Bottom);
            foreach (var item in lst)
            {
                if (filter.Contains(item.di))
                    ord.Add(new ordDraw(dv, item.cm, item.dc));
            }
        }
        public void addOrdDims()
        {
            if (lst.Count == 0) return;
            var v = lst[0].dir;
            DrawingCurve dc = null;
            var cp1 = getMin(v);
            var tmp = u.get<CornerPt>(lst, fi => !fi.checkPar(v));
            if (tmp != null) dc = tmp.dc;
            ordDraw od = new ordDraw(dv, cp1.cm, cp1.dc);
            od.setDC(dc); ord.Add(od);
            if (tmp == null) return;
            v = tmp.dir;
            var cp2 = getMin(v);
            od = new ordDraw(dv, cp2.cm, cp2.dc);
            od.setDC(cp1.dc); ord.Add(od);
        }
        public CornerPt getMin(Vector2d v)
        {
            var vals = u.gets<CornerPt>(lst, fi => fi.checkPar(v));
            if (vals.Count() == 1) return vals.First();
            var v1 = vals.ElementAt(0); var v2 = vals.ElementAt(1);
            return v1.d <= v2.d ? v1 : v2;
        }
    }
    public class CornerPt
    {
        public DrawingCurve dc;
        public Vector2d dir;
        public direct di;
        public double d;
        public int octant;
        public Centermark cm;
        public CornerPt(DrawingCurve dc, Centermark cm, double d, direct di)
        {
            this.dc = dc; this.cm = cm; this.d = d; this.di = di;
            getVec();
        }
        public void getVec()
        {
            var v = u.normal(dc.StartPoint, dc.EndPoint, cm.Position);
            v.Normalize(); v.ScaleBy(-1);
            dir = v;
            octant = u.octant(v);
        }
        public bool checkPar(Vector2d v)
        {
            return dir.IsParallelTo(v, 0.01);
        }
        public bool check(List<int []> vals)
        {
            foreach (var item in vals)
            {
                if (check(item)) return true;
            }
            return false;
        }
        public bool check(int [] f)
        {
            int v1 = f[0], v2 = f[1];
            return v1 < v2 ? (octant > v1 && octant < v2) :
                !(octant > v2 && octant < v1);
        }
    }
    public class DrawBox
    {
        public DrawingView dv;
        public Box2d dvBox;
        public Projection prv, prn;
        public Point2d mp;
        public Vector2d v;
        public List<BoxData> bd = new List<BoxData>();
        //int o;
        public DrawBox(DrawingView dv, Vector2d v = null)
        {
            this.dv = dv;
            this.v = v;
            dvBox = I.Box(dv);
        }
        public void set(object dc)
        {
            if (dc is Centerline)
                bd.Add(new BoxData(dc as Centerline));
            else if (dc is DrawingCurve)
                bd.Add(new BoxData(dc as DrawingCurve));
            else if (dc is DrawingCurveSegment)
                bd.Add(new BoxData((dc as DrawingCurveSegment).Parent));
            else if (dc is Centermark)
                bd.Add(new BoxData(dc as Centermark));
        }
        public void setPr()
        {
            bd[0].setType(); bd[1].setType();
            Point2d pt11 = bd[0].getPt(), pt12 = bd[0].getPt(false), 
                pt21 = bd[1].getPt(), pt22 = bd[1].getPt(false);
            var v1 = bd[0].getDir(bd[1]) ?? bd[1].getDir(bd[0]);
            if (v != null) v1 = v;
            var l1 = bd[0].getL(v1); var l2 = bd[1].getL(v1);
            if (l1 == 0 && l2 == 0)
            {
                v1 = pt11.VectorTo(pt21);
                var n = I.CV2d(v1); u.normal(n, null);
                prn = new Projection(n, false);
                prv = new Projection(v1, false);
                prn.set(pt11, pt21);
                prv.set(pt11, pt21);
                mp = u.midPt(pt11, pt21);
                return;
            } 
            else
            {
                var n = I.CV2d(v1); u.normal(n, null);
                prn = new Projection(v1, false);
                prv = new Projection(n, false);
            }
            Point2d pt1, pt2, pt3, pt;
            bd[0].getPoints(prv.v, bd[1], out pt1, out pt2, out pt3);
            prn.set(pt1, pt2);
            pt = I.CP2d(pt1);
            //u.addText(dv, "pt11", pt11); u.addText(dv, "pt12", pt12);
            //u.addText(dv, "pt21", pt21); u.addText(dv, "pt22", pt22);
            Vector2d vec1 = pt1.VectorTo(pt3), norm = null;
            u.normal(prn.v, vec1, out norm);
            var len1 = pt.VectorTo(pt3).Length;
            pt.TranslateBy(norm);
            //u.addText(dv, "pt1", pt1); u.addText(dv, "pt2", pt2); u.addText(dv, "pt3", pt3);
            //u.addText(dv, "tr", pt);
            var len2 = pt.VectorTo(pt21).Length;
            if (len1 < len2)
            {
                norm.ScaleBy(-2); pt.TranslateBy(norm);
            }
            pt = u.midPt(pt2, pt);
            //u.addText(dv, "mp", pt);
            mp = pt;
            prv.set(pt, pt3);
        }
        public Vector2d getDir()
        {
            prn.add(dvBox.MinPoint); prn.add(dvBox.MaxPoint);
            double v1 = prn.getPt(0), v2 = prn.getPt(1);
            Vector2d v = I.CV2d(prn.v);
            double t1 = v1 - prn.C, t2 = v2 - prn.C;
            //u.addText(dv, "p1", pt1); u.addText(dv, "p2", pt2);
            
            Vector2d bvec = I.CV2d(1, 0);
            double dist = u.dotProduct(v, I.CV2d(1, 1));
            if (!prn.v.IsParallelTo(bvec, 0.01) && !prn.v.IsPerpendicularTo(bvec, 0.01))
            {
                dist = 1;
            }
            else if (Math.Abs(Math.Abs(t1) - Math.Abs(t2)) < 0.2)
            {
                dist = dist * (t2 * u.dotProduct(v, I.CV2d(1, 1)) - 1);
            }
            else if (Math.Abs(t1) > Math.Abs(t2))
            {
                dist = dist * (t2 * u.dotProduct(v, I.CV2d(1,1)) - 1);
            }
            else
            {
                dist = dist * (t1 * u.dotProduct(v, I.CV2d(1, 1)) + 1);
            }
            u.setDist(v, dist);
            mp.TranslateBy(v);
            //u.addText(dv, "mp", mp);
            v.Normalize();
            return v;
        }
        public GeometryIntent getGI(int i)
        {
            Vector2d v1 = I.CV2d(prv.v);
            Point2d pt = bd[i].getPt();
            double d = u.dotProduct(v1, mp.VectorTo(pt));
            if (d != 0) v1.ScaleBy(d); 
            v1.Normalize();
            GeometryIntent gi = bd[i].getGI(dv,v1, mp);
            return gi;
        }
        public void transMP(DrawingCurve dc, double dist)
        {
            var v = u.normal(dc.StartPoint, dc.EndPoint, mp);
            v.ScaleBy(-1);
            var d = v.Length;
            v.Normalize();
            double norm = u.dotProduct(v, dc.StartPoint.VectorTo(dv.Center));
            if (norm < 0)
                v.ScaleBy(d + dist);
            else v.ScaleBy(d - dist);
            mp.TranslateBy(v);
        }
    }
    public enum bdType { nullType, cenLine, cenMark, arc, Line}
    public class BoxData
    {
        DrawingCurve dc;
        Centerline cl;
        Centermark cm;
        bdType t;
        public Point2d pt = null;
        public BoxData(DrawingCurve dc)
        {
            this.dc = dc;
        }
        public BoxData(Centermark cm)
        {
            this.cm = cm;
        }
        public BoxData(Centerline cl)
        {
            this.cl = cl;
        }
        public GeometryIntent getGI(DrawingView dv, Vector2d v = null, Point2d mp = null)
        {
            //if (dc != null)
            //{
            //    u.addText(dv, "pt1", dc.MidPoint);
            //}
            PointIntentEnum pie;
            if (v != null)
            {
                int o = u.octant(v);
                if (dc != null && (dc.ProjectedCurveType == Curve2dTypeEnum.kCircleCurve2d ||
                    dc.ProjectedCurveType == Curve2dTypeEnum.kCircularArcCurve2d))
                {
                    pie = o == 1 ? PointIntentEnum.kCircularRightPointIntent :
                        o == 4 ? PointIntentEnum.kCircularLeftPointIntent :
                        o == 3 ? PointIntentEnum.kCircularTopPointIntent :
                        PointIntentEnum.kCircularBottomPointIntent;
                    return dv.Parent.CreateGeometryIntent(dc, pie);
                }
            }
            if (dc != null && pt != null)
            {
                pie = getPIE();
                return dv.Parent.CreateGeometryIntent(dc, pie);
            }
            return cl != null ? dv.Parent.CreateGeometryIntent(cl) :
                cm != null ? dv.Parent.CreateGeometryIntent(cm) :
                dc != null ? u.addGIDC(dv, dc, mp) : null;
        }
        public PointIntentEnum getPIE()
        {
            PointIntentEnum pie = u.eq(pt, dc.StartPoint) ? PointIntentEnum.kStartPointIntent :
                    u.eq(pt, dc.EndPoint) ? PointIntentEnum.kEndPointIntent : PointIntentEnum.kCenterPointIntent;
            return pie;
        }
        public Point2d maxPt(Vector2d v, Point2d pt)
        {
            Point2d pt1, pt2;
            if (dc != null) { pt1 = dc.StartPoint; pt2 = dc.EndPoint; }
            else if (cl != null) { pt1 = cl.StartPoint; pt2 = cl.EndPoint; }
            else if (cm != null) return cm.Position;
            else { return null; }
            double d1 = maxPt(v, pt, pt1), d2 = maxPt(v, pt, pt2);
            return d1 <= d2 ? pt2 : pt1;
        }
        public double maxPt(Vector2d v, Point2d pt, Point2d pt1)
        {
            if (pt == null || pt1 != null) return 0;
            var vec = pt.VectorTo(pt1);
            var v1 = I.CV2d(v); v1.Normalize();
            return Math.Abs(u.dotProduct(v1, vec));
        }
        public void getPoints(Vector2d v, BoxData d, out Point2d pt1, out Point2d pt2, out Point2d pt3)
        {
            double l1 = getL(v), l2 = d.getL(v);
            if (check(v, d) || d.check(v, this))
            {
                d.pt = d.maxPt(v, getPt()); pt = maxPt(v, d.getPt());
            }
            if (t == bdType.arc)
            {
                pt1 = d.getPt(); pt2 = d.getPt(false); pt3 = maxPt(v, pt1); pt = pt3;
                if (d.t == bdType.Line)
                {
                    var vec = d.getDir();
                    if (!vec.IsParallelTo(v, 0.01)) d.pt = d.maxPt(v, pt3);
                }
            }
            else if (d.t == bdType.arc)
            {
                pt1 = getPt(); pt2 = getPt(false); pt3 = d.maxPt(v, pt1); d.pt = pt3;
                if (t == bdType.Line)
                {
                    var vec = getDir();
                    if (!vec.IsParallelTo(v, 0.01)) pt = maxPt(v, pt3);
                }
            }
            else if (l1 == 0)
            {
                pt1 = getPt(); pt2 = getPt(false); pt3 = d.getPt();
            }
            else if (l2 == 0)
            {
                pt1 = d.getPt(); pt2 = d.getPt(false); pt3 = getPt();
            }
            else if (l1 <= l2)
            {
                pt1 = getPt(); pt2 = getPt(false); pt3 = d.getPt();
            }
            else
            {
                pt1 = d.getPt(); pt2 = d.getPt(false); pt3 = getPt();
            }
        }
        public bool check(Vector2d dir, BoxData d)
        {
            //if (d.t == bdType.cenMark || t == bdType.cenMark) return false;
            var d1 = getDir();
            var d2 = d.getDir();
            if (d1 == null) return false;
            if (d2 != null)
            {
                if (d1.IsParallelTo(d2, 0.01)) return false;
            }
            if (!d1.IsParallelTo(dir, 0.01)) return true;
            return false;
        }
        public void setType()
        {
            t = (dc != null && dc.ProjectedCurveType == Curve2dTypeEnum.kLineSegmentCurve2d) ?
                bdType.Line : (cl != null && cl.GeometryType == CurveTypeEnum.kLineSegmentCurve) ?
                bdType.cenLine : (cm != null) ? 
                bdType.cenMark : (dc != null && dc.ProjectedCurveType == Curve2dTypeEnum.kCircularArcCurve2d) ?
                bdType.arc : bdType.nullType;
        }
        public Vector2d getDir(BoxData bd = null)
        {
            if (bd != null)
            {
                if (t == bdType.arc && (bd.t == bdType.Line || bd.t == bdType.cenLine)) return bd.getDir();
            }
            switch (t)
            {
                case bdType.nullType:
                    break;
                case bdType.cenLine:
                    return cl.StartPoint.VectorTo(cl.EndPoint);
                case bdType.Line:
                    return dc.StartPoint.VectorTo(dc.EndPoint);
                case bdType.arc:
                    if (!u.eq(((Arc2d)dc.Segments[1].Geometry).SweepAngle, Math.PI / 2))
                        return dc.StartPoint.VectorTo(dc.EndPoint);
                    break;
                default:
                    break;
            }
            return null;
        }
        public Point2d getPt(bool first = true)
        {
            switch (t)
            {
                case bdType.nullType:
                    break;
                case bdType.cenLine:
                    return first ? cl.StartPoint : cl.EndPoint;
                case bdType.Line:
                    return first ? dc.StartPoint : dc.EndPoint;
                case bdType.cenMark:
                    return cm.Position;
                case bdType.arc:
                    return first ? dc.StartPoint : dc.EndPoint;
                default:
                    break;
            }
            return null;
        }
        public double getL(Vector2d v = null)
        {
            switch (t)
            {
                case bdType.nullType:
                    break;
                case bdType.cenLine:
                    return cl.StartPoint.VectorTo(cl.EndPoint).Length;
                case bdType.Line:
                    return dc.StartPoint.VectorTo(dc.EndPoint).Length;
                case bdType.arc:
                    var vec = dc.StartPoint.VectorTo(dc.EndPoint);
                    if (v == null) return -1000;
                    var v1 = I.CV2d(v); v1.Normalize();
                    return Math.Abs(u.dotProduct(vec, v1));
                default:
                    break;
            }
            return 0;
        }
    }
    public class DrawData
    {
        public Face bf;
        public DrawingCurve dc1, dc2;
        public int dir = 0;
        public double t = 0;
        public DrawingView dv;
        //public Vector2d v = null;
        public bool remove = false;
        public DrawData(DrawingView dv, Face f)
        {
            bf = f;
            this.dv = dv;
            dc1 = null; dc2 = null;
        }
        public void set(DrawingCurve dc, int num)
        {
            if (num == 1) dc1 = dc;
            else if (num == 2) dc2 = dc;
            //v = dc.StartPoint.VectorTo(dc.EndPoint);
        }
        public LinearGeneralDimension addDim(bool move, string st, double min, bool ra = false)
        {
            if (!(dc1 != null && dc2 != null))
            {
                //MessageBox.Show("Непараллельный гиб");
                return null;
            }

            //Vector2d h = I.CV2d(0, 1);
            //DimensionTypeEnum dte = DimensionTypeEnum.kAlignedDimensionType;
            //var v = getVector();
            //if (v.IsParallelTo(h, 0.001)) dte = DimensionTypeEnum.kHorizontalDimensionType;
            //else if (v.IsPerpendicularTo(h, 0.001)) dte = DimensionTypeEnum.kVerticalDimensionType;

            if (t != 0) addRaRz();
            DrawBox db = new DrawBox(dv);
            db.set(dc1); db.set(dc2);
            db.setPr();
            var v = db.getDir();
            var dim = u.addDim(dv, db.getGI(0), db.getGI(1), v, db.mp);
            I.getStyle(st);
            dim.Style = I.dimStyle;
            //var dim = u.addDim(dv, dc1, dc2, dte, v, st, move, setDir(v), min);
            //addAngleDim();
            return dim;
        }
        public void addRaRz()
        {
            DrawingCurve dc = (u.eq(u.getLenght(dc1), t)) ? dc1 : dc2;
            var ra = new RaRz(dv);
            ra.add(dc);
        }
        public Vector2d setDir(Vector2d v)
        {
            if (dir == 0) return null; 
            Vector2d r = null;
            Vector2d[] dirs = { I.CV2d(1, 0), I.CV2d(0, -1) };
            for (int i = 0; i < dirs.Length; i++)
            {
                if (dirs[i].IsParallelTo(v, 0.01)) r = dirs[i];
            }
            if (r == null) return null;
            if (dir == 1) r.ScaleBy(-1);
            return r;
        }
        public void addAngleDim()
        {
            var g = dc1.Segments[1].Geometry as Arc2d;
            if (g == null) return;
            if (u.eq(Math.Abs(g.SweepAngle), Math.PI / 2)) return;
            I.getStyle("Текст**");
            var dc = u.get<DrawingCurve>(dv.DrawingCurves, fi => fi.ProjectedCurveType == Curve2dTypeEnum.kCircularArcCurve2d &&
            u.eq(((Arc2d)fi.Segments[1].Geometry).Center, g.Center) && !fi.Equals(dc1));
            if (dc != null)
            addAngleDim(dc);
        }
        public AngularGeneralDimension addAngleDim(DrawingCurve dc)
        {
            Arc2d g = dc.Segments[1].Geometry as Arc2d;
            if (u.eq(Math.Abs(g.SweepAngle), Math.PI) || u.eq(Math.Abs(g.SweepAngle), Math.PI*0.5)) return null;
            Point2d pt1 = g.StartPoint, pt2 = g.EndPoint;
            Sheet sh = dv.Parent as Sheet;
            var ob1 = sh.FindUsingPoint(pt1, 0.01);
            var l1 = u.get<DrawingCurveSegment>(ob1, fi => fi.GeometryType == Curve2dTypeEnum.kLineSegmentCurve2d);
            var ob2 = sh.FindUsingPoint(pt2, 0.01);
            var l2 = u.get<DrawingCurveSegment>(ob2, fi => fi.GeometryType == Curve2dTypeEnum.kLineSegmentCurve2d);
            var t1 = u.getTangent(g.Evaluator, 0.1, pt1);
            var t2 = u.getTangent(g.Evaluator, 0.1, pt2);
            double d = 2;
            Vector2d v1 = I.CV2d(-t1[0], -t1[1]), v2 = I.CV2d(t2[0], t2[1]);
            u.setDist(v1, d); u.setDist(v2, d);
            Point2d m = u.midPt(I.CP2d(g.Center, v1), I.CP2d(g.Center, v2));
            GeometryIntent i1 = sh.CreateGeometryIntent(l1.Parent), i2 = sh.CreateGeometryIntent(l2.Parent);
            var dim = sh.DrawingDimensions.GeneralDimensions.AddAngular(m, i1, i2, DimensionStyle: I.dimStyle);
            return dim;
        }
        public Vector2d getVector()
        {
            return dc1.ProjectedCurveType == Curve2dTypeEnum.kLineSegmentCurve2d ?
                dc1.StartPoint.VectorTo(dc1.EndPoint) : dc2.StartPoint.VectorTo(dc2.EndPoint);
        }
        public void check()
        {
            if (dc1 == null || dc2 == null) remove = true;
        }
    }
    public enum connectType { tang, normal}
    public class BendData
    {
        Face f;
        LineSegment ls;
        public Dictionary<DrawingView, DrawData> curves = new Dictionary<DrawingView, DrawData>();
        public double dist, t;
        UnitVector bn;
        public List<Edge> edges = new List<Edge>();
        public List<Edge> bedges = new List<Edge>();
        public BendData(Edge e, Face bf, double t)
        {
            edges.Add(e);
            this.t = t;
            ls = (LineSegment)e.Geometry;
            bedges.AddRange(u.gets<Edge>(bf.Edges, fi => true));
            f = u.get<Face>(e.Faces, fi => !fi.Equals(bf));
            findEdge();
        }
        public void set(DrawingCurve dc, DrawingView dv, Face bf)
        {
            if (dc.ProjectedCurveType == Curve2dTypeEnum.kLineSegmentCurve2d &&
                dc.CurveType == CurveTypeEnum.kCircularArcCurve) return;
            DrawData dd = new DrawData(dv, bf);
            if (curves.ContainsKey(dv)) dd = curves[dv];
            else
            {
                curves.Add(dv, dd);
            }
            if (bedges.Contains(dc.ModelGeometry) && dc.CurveType != CurveTypeEnum.kBSplineCurve)
            {
                dd.set(dc, 1); 
            }
            else if (edges.Contains(dc.ModelGeometry) && dc.CurveType != CurveTypeEnum.kBSplineCurve) 
                dd.set(dc, 2);
        }
        public void clear()
        {
            foreach (var item in curves)
            {
                item.Value.check();
            }
        }
        public void findEdge()
        {
            Plane pl = f.Geometry as Plane;
            if (pl == null) return;
            bn = pl.Normal;
            double d = 0.5; Edge e = null;
            Dictionary<double, List<Edge>> des = new Dictionary<double, List<Edge>>();
            foreach (var item in u.gets<Edge>(f.Edges, fi => fi.GeometryType == CurveTypeEnum.kLineSegmentCurve))
            {
                LineSegment l = (LineSegment)item.Geometry;
                if (!l.Direction.IsParallelTo(ls.Direction, 0.01)) continue;
                var d1 = u.normal(item, ls.StartPoint, bn);
                if (d1.Length > d)
                {
                    double len = u.round(d1.Length);
                    if (des.Keys.Contains(len)) des[len].Add(item);
                    else
                    {
                        des.Add(len, new List<Edge> { item });
                    }
                    d = len; e = item;
                }
            }
            if (e != null)
            {
                foreach (var item in des[d])
                {
                    edges.Add(item); addEdges(item);
                }
                 dist = d; 
            }
        }
        public void addEdges(Edge e)
        {
            var fa = u.get<Face>(e.Faces, fi => !fi.Equals(f));
            IEnumerable<Edge> es = null;
            if (fa.SurfaceType == SurfaceTypeEnum.kPlaneSurface)
            {
                es = u.gets<Edge>(fa.Edges, fi => u.eq(fi, t));
            }
            else if (fa.SurfaceType == SurfaceTypeEnum.kCylinderSurface)
            {
                es = u.gets<Edge>(fa.Edges, fi => fi.GeometryType == CurveTypeEnum.kCircularArcCurve);
            }
            else return;
            edges.AddRange(es);
        }
    }
    public class DimBend : DimsBase
    {
        readonly Face f;
        readonly double t;
        public bool added = false;
        BendData bend = null;
        public DimBend(Face f, double t)
        {
            this.f = f;
            this.t = t;
            get();
            //type = (e.CurveType == CurveTypeEnum.kCircleCurve) ? connectType.tang : connectType.normal;
        }
        public override void get()
        {
            BendData tmp;
            foreach (var item in u.gets<Edge>(f.Edges, fi => fi.GeometryType == CurveTypeEnum.kLineSegmentCurve))
            {
                tmp = new BendData(item, f, t);
                if (bend == null) bend = tmp;
                else if (tmp.dist < bend.dist) bend = tmp;
            }
        }
        public override void set(DrawingCurve dc, DrawingView dv)
        {
            if (added) return;
            bend.set(dc, dv, f);
        }
        public override void addDims()
        {
            bend.clear();
            List<DrawData> dd = bend.curves.Values.Where(fi => !fi.remove).ToList();
            setRaRz(dd);
            foreach (var g in dd.GroupBy(fi => fi.bf))
            {
                if (g.Count() == 1)
                {
                    //continue;
                    g.ElementAt(0).addDim(true, "Текст**", 1.2);
                }
                else
                {
                    bool add = false;
                    foreach (var item in g)
                    {
                        if (!add && item.dc1.ProjectedCurveType == Curve2dTypeEnum.kCircularArcCurve2d)
                        {
                            item.addDim(true, "Текст**", 1.2);
                            add = true;
                        }
                    }
                    if (!add) g.ElementAt(0).addDim(true, "Текст**", 1.2);
                }
            }
        }
        public void setRaRz(List<DrawData> dd)
        {
            foreach (var item in dd)
            {
                if (u.eq(u.getLengthModel(item.dc1), t) || u.eq(u.getLengthModel(item.dc2), t))
                {
                    item.t = t; return;
                }
            }
        }
    }
    public class Projections
    {
        const int tol = 1000;
        DrawingView dv;
        Vector2d v, n;
        int y;
        public double Y
        {
            get => (double)y / tol;
        }
        IEnumerable<Box2d> bs;
        IEnumerable<Projection> prs;
        List<Projection> jprs = new List<Projection>();
        public List<Projection> holes = new List<Projection>();
        public Projections(DrawingView dv, Vector2d dir, Point2d sp)
        {
            this.dv = dv; v = dir;
            n = I.CV2d(v); u.normal(n, null);
            y = (int)(u.dotProduct(n, sp) * 1000);
        }
        public void fill()
        {
            bs = u.gets<DrawingCurve>(dv.DrawingCurves, fi => 
            fi.ProjectedCurveType == Curve2dTypeEnum.kCircleCurve2d ||
            fi.ProjectedCurveType == Curve2dTypeEnum.kBSplineCurve2d ||
            fi.ProjectedCurveType == Curve2dTypeEnum.kCircularArcCurve2d).Select(el => el.Evaluator2D.RangeBox);
            bs = (v.IsParallelTo(I.CV2d(1, 0), 0.001)) ?
                bs.OrderBy(el => el.MinPoint.X) : bs.OrderBy(el => el.MinPoint.Y);
            prs = bs.Select(el => new Projection(el, v));
        }
        public void draw(IEnumerable<Projection> pr)
        {
            double o = -0.5;
            int i = 0;
            foreach (var item in pr)
            {
                var pt1 = I.CP2d(item.get(1), Y + o * i);
                var pt2 = I.CP2d(item.get(2), Y + o * i);
                dv.Parent.DrawingNotes.GeneralNotes.AddFitted(pt1, "s");
                dv.Parent.DrawingNotes.GeneralNotes.AddFitted(pt2, "e");
                i++;
            } 
        }
        public void draw(IEnumerable<Box2d> pr)
        {
            int o = -500, i = 0;
            foreach (var item in pr)
            {
                var pt1 = item.MinPoint; var pt2 = item.MaxPoint;
                dv.Parent.DrawingNotes.GeneralNotes.AddFitted(pt1, "s");
                dv.Parent.DrawingNotes.GeneralNotes.AddFitted(pt2, "e");
                i++;
            }
        }
        public void addHoles(Projection p)
        {
            holes.Add(new Projection(p.s, jprs[0].s));
            for (int i = 1; i < jprs.Count; i++)
            {
                holes.Add(new Projection(jprs[i - 1].e, jprs[i].s));
            }
            holes.Add(new Projection(jprs[jprs.Count-1].e, p.e));
            //draw(holes);
        }
        public void setCut(double d)
        {
            double sum = d;
            Projection [] tmp = holes.OrderByDescending(el => el.L).ToArray();
            double min =  0.4 / dv.Scale;
            for (int i = 0; i < tmp.Length; i++)
            {
                var item = tmp[i];
                if (sum <= 0) break;
                var L = item.L;
                item.Cut = L;
                if (sum < L) item.Cut = sum;
                sum -= item.Cut;
                if (sum <= min) { item.Cut = item.Cut < item.L ? item.Cut : item.L; break; }
            }
        }
        public void addBreaks()
        {
            //Point2d dvCen = dv.Center;
            double tol = 0.5, gap = 0.2, sum = 0, cen = 0;
            double z = gap;
            BreakOrientationEnum boe = v.IsParallelTo(I.CV2d(1, 0), 0.01) ?
                BreakOrientationEnum.kHorizontalBreakOrientation : BreakOrientationEnum.kVerticalBreakOrientation;
            for (int i = 0; i < holes.Count; i++)
            {
                double L = holes[i].Cut;
                if (L == 0 || L < tol) continue;
                cen = holes[i].C + sum;
                var pt1 = I.CP2d(cen - ( L * 0.5 - z * 0.5), Y, v); 
                var pt2 = I.CP2d(cen + ( L * 0.5 - z * 0.5), Y, v);
                sum -= (L - gap - z) * 0.5;
                dv.BreakOperations.Add(boe, pt1, pt2, BreakStyleEnum.kRectangularBreakStyle, 10, gap, 1, false);
            }
            //dv.Position = dvCen;
        }
        public Centerline getCL()
        {
            foreach (var item in dv.Parent.FindUsingPoint(dv.Center))
            {
                var cl = item as Centerline;
                if (cl != null && cl.GeometryType == CurveTypeEnum.kLineSegmentCurve
                    && ((LineSegment2d)cl.Geometry).Direction.AsVector().IsPerpendicularTo(v, 0.01))
                    return cl;
            }
            return null;
        }
        public void addCenter()
        {
            double offset = 0.1;
            var cl = getCL();
            if (cl == null) return;
            var d = u.dotProduct(v, cl.StartPoint);
            var p = u.get<Projection>(holes, fi => fi.S <= d && d <= fi.E);
            var ind = holes.IndexOf(p);
            double end = p.E;
            if (p == null) return;
            p.E = d - offset;
            holes.Insert(ind+1, new Projection(d + offset, end));
        }
        public void join()
        {
            var lst = prs.ToList();
            Projection p1, p2, p3;
            for (int i = 0; i < lst.Count; i++)
            {
                p1 = lst[i];
                p3 = p1;
                for (int j = i+1; j < lst.Count; j++)
                {
                    //if (j == i) continue;
                    p2 = lst[j];
                    if (p1.contains(p2))
                    {
                        continue;
                    }
                    var ptmp = p1.join(p2);
                   
                    if (ptmp != null) { p3 = ptmp; p1 = p3; i = j - 1; }
                    else 
                    {
                        i = j-1;
                        break;
                    }
                }
                if (p3 != null) 
                    jprs.Add(p3);
            }
            //draw(jprs);
        }
    }
    public class Projection3d
    {
        const int tol = 1000;
        Vector v;
        public int s, e, c;
        public double S
        {
            get => (double)s / tol;
            set => s = (int)(value * tol);
        }
        public double E
        {
            get => (double)e / tol;
            set => e = (int)(value * tol);
        }
        public double C
        {
            get => ((double)s + (double)e) / (2 * tol);
            set => c = (int)(value * tol);
        }
        public double L
        {
            get => (double)(e - s) / tol;
        }
        public Projection3d(Vector v, Point pt, Point cen = null)
        {
            v.Normalize();
            this.v = v;
            cen = cen ?? I.CP();
            var vec = cen.VectorTo(pt);
            var d = u.dotProduct(v, vec);
            S = 0; E = d;
        }
    }
    public class Projection2dData
    {
        public int v;
        public Point2d pt;
        public Projection2dData(int val, Point2d pt)
        {
            this.pt = pt; this.v = val;
        }
    }
    public class Projection
    {
        const int tol = 1000;
        public Vector2d v;
        public List<Projection2dData> pts = new List<Projection2dData>();
        public int s, e, c;
        public int n;
        public int cut;
        public double Cut
        {
            set => cut = (int)(value * tol);
            get => (double)cut / tol;
        }
        public double S
        {
            get => (double)s / tol;
            set => s = (int)(value * tol);
        }
        public double E
        {
            get => (double)e / tol;
            set => e = (int)(value * tol);
        }
        public double C
        {
            get => ((double)s + (double)e) / (2 * tol);
            set => c = (int)(value * tol);
        }
        public double L
        {
            get => (double)(e - s) / tol;
        }
        public Projection(Vector2d v, bool abs = true)
        {
            var vec = I.CV2d(v);
            vec.Normalize();
            if (abs) u.abs(ref vec);
            this.v = vec;
        }
        public Projection(double s, double e)
        {
            set((int)(s * tol), (int)(e * tol));
        }
        public Projection(int s, int e)
        {
            set(s, e);
        }
        public Projection(Box2d b, Vector2d v, Point2d sp = null)
        {
            this.v = v;
            set((int)(get(b.MinPoint, this.v) * tol), (int)(get(b.MaxPoint, this.v) * tol));
        }
        public Projection(DrawingCurve dc, Vector2d v)
        {
            this.v = v;
            if (dc.CurveType == CurveTypeEnum.kLineCurve)
            {
                set((int)(get(dc.StartPoint, this.v) * tol), (int)(get(dc.EndPoint, this.v) * tol));
            }
            else if (dc.CurveType == CurveTypeEnum.kCircleCurve || dc.CurveType == CurveTypeEnum.kCircularArcCurve)
            {
                set((int)(get(dc.CenterPoint, this.v) * tol), (int)(get(dc.CenterPoint, this.v) * tol));
            }
        }
        public void set(Point2d pt1, Point2d pt2)
        {
            var flag = set((int)(get(pt1, this.v) * tol), (int)(get(pt2, this.v) * tol));
            if (flag)
            { pts.Add(new Projection2dData(s, pt1)); pts.Add(new Projection2dData(e, pt2)); }
            else
            {
                pts.Add(new Projection2dData(s, pt2)); pts.Add(new Projection2dData(e, pt1));
            }
        }
        public void add(Point2d pt)
        {
            var val = (int)(get(pt, this.v) * tol);
            pts.Add(new Projection2dData(val, pt));
        }
        public double get()
        {
            var r = pts[pts.Count-1];
            return (double)r.v / tol;
        }
        public void setD(double p)
        {
            var tmp = (int)(p * tol);
            s = c - tmp /2; e = c + tmp/2;
        }
        public void setC(double cen)
        {
            c = (int)(cen*tol);
        }
        public double dir(Projection p)
        {
            int mid = (e + s) /2;
            int v1 = mid - p.s, v2 = mid - p.e;
            var m = u.minAbs(v1, v2) / tol;
            var t = p.L / m;
            return (t < 2) ? 0 : m;
        }
        public Point2d textPt(Projection p, Point2d pt)
        {
            var d = dir(p);
            if (d == 0) return pt;
            var d1 = u.dotProduct(v, pt);
            double mid = (E + S) / 2;
            return pt;
        }
        public double get(int n)
        {
            return n == 1 ? (double)s / tol : (double)e / tol;
        }
        public double getPt(int n)
        {
            return (double)pts[pts.Count - 1 - n].v / tol;
        }
        public bool set(int s, int e)
        {
            if (s <= e)
            {
                this.s = s; this.e = e;
            } else
            {
                this.e = s; this.s = e;
                return false;
            }
            return true;
        }
        public bool eq(int val)
        {
            return val == s || val == e ;
        }
        public double get(Point2d ep, Vector2d vec)
        {
            return u.dotProduct(vec, ep);
        }
        public bool contains(int val)
        {
            return val >= s && val <= e;
        }
        public bool contains2(int val)
        {
            return val > s && val < e;
        }
        public bool contains(Point2d pt)
        {
            var val = (int)(get(pt, this.v) * tol);
            return contains(val);
        }
        public bool disjoint(Projection p)
        {
            if (contains(p) || p.contains(this)) return false;
            if (p.s >= e || p.e <= s) return false;
            bool b1, b2;
            b1 = contains2(p.s);
            if (!b1) b1 = eq(p.s);
            b2 = contains2(p.e);
            if (!b2) b2 = eq(p.e);
            return b1 != b2;
        }
        public bool contains(Projection p)
        {
            return contains((int)p.s) && contains((int)p.e);
        }
        public Projection getMax(Projection p, Func<double, double ,bool> f)
        {
            var l1 = e - s; var l2 = p.e - p.s;
            return f(l1, l2) ? this : p;
        }
        public Projection join(Projection p)
        {
            Projection proj = null;
            if (contains(p.s) && !contains(p.e)) proj = new Projection(this.s, p.e);
            return proj;
        }
    }
    public class Gabs
    {
        DrawingView dv;
        GabEnt[] ents;
        public LinearGeneralDimension xdim, ydim;
        public Gabs(DrawingView dv)
        {
            this.dv = dv;
            ents = new GabEnt[4];
            fill();
        }
        public void fill()
        {
            Vector2d[] dirs = { I.CV2d(-1,0), I.CV2d(1,0), I.CV2d(0,-1), I.CV2d(0,1)};
            for (int i = 0; i < ents.Length; i++)
            {
                ents[i] = new GabEnt(dv); ents[i].set(dirs[i]);
            }
        }
        public void draw()
        {
            Vector2d[] vecs = { I.CV2d(0, 1), I.CV2d(-1,0)};
            GabEnt x1 = ents[0], x2 = ents[1], y1 = ents[2], y2 = ents[3];
            xdim = draw(x1, x2, vecs[0]);
            ydim = draw(y1, y2, vecs[1]);
        }
        public LinearGeneralDimension draw(GabEnt e1, GabEnt e2, Vector2d v)
        {
            if (e1.dc == null || e2.dc == null) return null;
            DrawBox db = new DrawBox(dv, v);
            db.set(e1.dc); db.set(e2.dc);
            db.setPr();
            var v1 = db.getDir();
            var dim = u.addDim(dv, db.getGI(0), db.getGI(1), v1, db.mp);
            if (u.containsDim(dv.Parent, dim)) { dim.Delete(); dim = null; }
            return dim;
            //var dvec = I.CV2d(v); dvec.Normalize();
            //e1.addGI(dvec); e2.addGI(dvec);
            //var dim = u.addDim(dv, e1.gi, e2.gi, v);
            //return dim;
        }
    }
    public class GabEnt
    {
        public DrawingCurve dc;
        DrawingView dv;
        BoxBase b;
        //Point2d pt;
        public GeometryIntent gi;
        public PointIntentEnum pie;
        Vector2d v;
        double d;
        //int o;
        public GabEnt(DrawingView dv)
        {
            this.dv = dv;
            b = new BoxBase(); b.set(dv);
            //u.addText(dv, "p", b.b.MinPoint);
            //b.draw(dv, 1);
        }
        public void set(Vector2d dir)
        {
            v = dir;
            d = b.get(v, dv.Center);
            //b.draw(dv, 2);
            var curves = u.gets<DrawingCurve>(dv.DrawingCurves, fi => check(fi));
            var l = u.gets<DrawingCurve>(curves, fi => fi.ProjectedCurveType == Curve2dTypeEnum.kLineSegmentCurve2d)
                ?.OrderByDescending( e => u.getLenght(e));
            if (l != null && l.Count() > 0) { dc = l.ElementAt(0); return; }
            var c = u.get<DrawingCurve>(curves, fi => fi.ProjectedCurveType == Curve2dTypeEnum.kCircleCurve2d ||
            fi.ProjectedCurveType == Curve2dTypeEnum.kCircularArcCurve2d);
            if (c != null) { dc = c;  return; }
        }
        public bool check(DrawingCurve dc, double tol = 0.001)
        {
            if (filter(dc)) return false;
            BoxBase tmp = new BoxBase();
            //DrawingCurveSegment ss = dc.Segments[1], se = dc.Segments[dc.Segments.Count];
            if (dc.ProjectedCurveType == Curve2dTypeEnum.kLineSegmentCurve2d && 
                dc.CurveType != CurveTypeEnum.kCircularArcCurve && dc.CurveType != CurveTypeEnum.kCircleCurve)
            {
                //var dcs = u.getSegm(dc, v, dv.Center, (d, sum) => d >= sum);

                var vec = dc.StartPoint.VectorTo(dc.EndPoint);
                var ed = dc.ModelGeometry as Edge;
                if (ed == null) return false;
                if (u.getLenght(ed) < 0.06) return false;
                tmp.set(dc.StartPoint, dc.EndPoint, vec);
                //u.addText(dv, "p", dc.MidPoint);
            }
            else if (dc.ProjectedCurveType == Curve2dTypeEnum.kCircularArcCurve2d)
            {
                var dcs = u.getSegm(dc, v, dv.Center, (d, sum) => d >= sum);
                if (dcs == null) return false;
                var geom = dcs.Geometry as Arc2d;
                tmp.set(geom);
            }
            var d1 = tmp.get(v, dv.Center);
            double dx = Math.Abs(d - d1);
            return dx < tol;
        }
        public bool filter(DrawingCurve dc)
        {
            var g = dc.Segments[1].Geometry as LineSegment2d;
            if (g == null) return false;
            //if (dc.CurveType == CurveTypeEnum.kCircularArcCurve) return true;
            return g.Direction.IsParallelTo(v.AsUnitVector(), 0.01);
        }
        public void addGI(Vector2d dvec)
        {
            int o = u.octant(v);
            var pt = b.getPoint(v, o, 0);
            //dv.Parent.DrawingNotes.GeneralNotes.AddFitted(pt, "pt" + o.ToString());
            if (dc.ProjectedCurveType == Curve2dTypeEnum.kCircleCurve2d || 
                dc.ProjectedCurveType == Curve2dTypeEnum.kCircularArcCurve2d)
            {
                pie = o == 1 ? PointIntentEnum.kCircularRightPointIntent :
                    o == 4 ? PointIntentEnum.kCircularLeftPointIntent :
                    o == 3 ? PointIntentEnum.kCircularTopPointIntent :
                    PointIntentEnum.kCircularBottomPointIntent;
                gi = dv.Parent.CreateGeometryIntent(dc, pie);
            }
            else if (dc.ProjectedCurveType == Curve2dTypeEnum.kLineSegmentCurve2d)
            {
                var v1 = dc.StartPoint.VectorTo(dc.EndPoint);
                if (v1.IsPerpendicularTo(v, 0.01))
                {
                    var pie = u.getPIE(dc, dvec, dv.Center, (d1, d2) => d1 >= d2);
                    if (pie == PointIntentEnum.kPlanarFaceCenterPointIntent)
                        gi = dv.Parent.CreateGeometryIntent(dc);
                    else gi = dv.Parent.CreateGeometryIntent(dc, pie);
                }
                else
                {
                    var box = dc.Evaluator2D.RangeBox;
                    pt = o >= 1 && o <= 3 ? box.MaxPoint : box.MinPoint;
                    gi = dv.Parent.CreateGeometryIntent(dc, pt);
                }
            }
        }
    }
    public class BoxBase
    {
        public Box2d b;
        protected Vector2d v, n;
        protected Point2d min, max, mid;
        public Projection prv, prn;
        public BoxBase()
        {
        }
        public void set(DrawingView dv)
        {
            this.mid = dv.Center;
            this.min = I.CP2d(dv.Center.X - dv.Width * 0.5, dv.Center.Y - dv.Height * 0.5);
            this.max = I.CP2d(dv.Center.X + dv.Width * 0.5, dv.Center.Y + dv.Height * 0.5);
            b = I.Box(min, max);
            setVN();
        }
        public void set(Box2d b)
        {
            this.b = b;
            this.min = b.MinPoint; this.max = b.MaxPoint;
            this.mid = u.midPt(min, max);
            setVN();
        }
        public void setVN()
        {
            v = I.CV2d(1, 0); n = I.CV2d(0, 1);
            prv = new Projection(b, v);
            prn = new Projection(b, n);
        }
        public void setVN(Box2d box,Vector2d v)
        {
            Vector2d vert = I.CV2d(0, 1), hor = I.CV2d(1,0);
            if (v.IsParallelTo(vert, 0.01) || v.IsPerpendicularTo(vert, 0.01))
            {
                setVN();
                return;
            }
            this.v = v;
            this.n = I.CV2d(v); u.normal(n, null);
            Vector2d vec = min.VectorTo(max);
            double a = u.dotProduct(vec, hor), b = u.dotProduct(vec, vert);
            v.Normalize();
            double tg = v.Y / v.X;
            double x = (b - a * tg) / (1 + tg * tg), y = a - x;
            double px = Math.Sqrt(y * y + y * y / (tg * tg)), py = Math.Sqrt(x * x + x * x * (tg * tg));
            prv = new Projection(v);
            prn = new Projection(n);
            var cen = u.midPt(box.MinPoint, box.MaxPoint);
            var vdp = u.dotProduct(v, cen);
            var ndp = u.dotProduct(n, cen);
        }
        public void setVN(Vector2d v, Point2d pt)
        {
            this.v = I.CV2d(Math.Abs(v.X), Math.Abs(v.Y));
            this.v.Normalize();
            n = I.CV2d(this.v); u.normal(n, null);
            n.X = Math.Abs(n.X); n.Y = Math.Abs(n.Y); 
            prv = new Projection(v);
            prn = new Projection(n);
            var vdp = u.dotProduct(this.v, pt);
            var ndp = u.dotProduct(this.n, pt);
            prv.setC(vdp);
            prn.setC(ndp);
            prv.setD(v.Length);
            prn.setD(0);
        }
        public void set(Point2d min, Point2d max, Vector2d v)
        {
            this.min = min; this.max = max; this.mid = u.midPt(min, max);
            b = I.Box(min, max);
            //var m = u.midPt(min, max);
            setVN(v, this.mid);
        }
        public void set(Arc2d a)
        {
            if (a == null) return;
            b = a.Evaluator.RangeBox;
            this.min = a.Evaluator.RangeBox.MinPoint;
            this.max = a.Evaluator.RangeBox.MaxPoint; this.mid = u.midPt(min, max);
            b = I.Box(min, max);
            setVN(b, I.CV2d(1, 0));
        }
        public void set(GeometryIntent gi1, GeometryIntent gi2, Vector2d v)
        {
            var p1 = get(gi1, v); var p2 = get(gi2, v);
            prv = p1.L < p2.L ? p1 : p2;
        }
        public Projection get(GeometryIntent gi, Vector2d v)
        {
            var dc = gi.Geometry as DrawingCurve;
            return new Projection(dc, v);
        }
        public double get(Vector2d v, Point2d c)
        {
            if (b == null) return -1000;
            int o = u.octant(v);
            //var v1 = I.CV2d(v); v1.Normalize();
            //double d1 = u.dotProduct(v1, b.MaxPoint, c), d2 = u.dotProduct(v1, b.MinPoint, c);
            //d1 = Math.Abs(d1); d2 = Math.Abs(d2);
            //double d = d1 >= d2 ? d1 : d2;
            double d = o == 1 ? b.MaxPoint.X : o == 4 ? b.MinPoint.X :
                o == 3 ? b.MaxPoint.Y : b.MinPoint.Y;
            return d;
            //return Math.Abs(Math.Round(d, 3));
        }
        public Point2d getPoint(Vector2d v, int o, double offset)
        {
            v.Normalize();
            Vector2d vec = o >= 1 && o <= 3 ? mid.VectorTo(max) : mid.VectorTo(min);
            double d = vec.DotProduct(v);
            d = d > 0 ? d + offset : d - offset;
            v.ScaleBy(d);
            var pt = I.CP2d(mid, v);
            return pt;
        }
        public void draw(DrawingView dv, int i)
        {
            dv.Parent.DrawingNotes.GeneralNotes.AddFitted(min, "m" + i);
            dv.Parent.DrawingNotes.GeneralNotes.AddFitted(max, "M" + i);
            //var pt1 = I.CP2d(prv.S, prn.C);
            //dv.Parent.DrawingNotes.GeneralNotes.AddFitted(pt1, "pt" + i);
        }
    }
    public class ArrangeDims
    {
        List<DimBox> dims = new List<DimBox>();
        List<DimBox> disj = new List<DimBox>();
        DrawingView dv;
        public ArrangeDims(IEnumerable<LinearGeneralDimension> ds, DrawingView dv)
        {
            this.dv = dv;
            foreach (var item in ds)
            {
                dims.Add(new DimBox(item, dv.Center));
            }
        }
        public ArrangeDims(DrawingView dv, List<DimBox> dims)
        {
            this.dv = dv;
            this.dims = dims;
        }
        public void sort()
        {
            dims = dims.OrderBy(el => el.prv.C).ToList();
        }
        public void setPos()
        {
            foreach (var item in dims)
            {
                item.setPos(dv);
            }
        }
        public bool isDisjoint(DimBox db)
        {
            if (disj.Contains(db)) return false;
            foreach (var item in dims)
            {
                if (item.Equals(db)) continue;
                if (disj.Contains(item)) continue;
                if (item.disjoint(db))
                {
                    var tmdb = item.getMax(db);
                    tmdb.setNullLev();
                    if (!disj.Contains(tmdb))
                    disj.Add(tmdb);
                    return true;
                }
            }
            return false;
        }
        public void setLevel()
        {
            setLevel(dims);
        }
        public void incLevel(List<DimBox> ds)
        {
            foreach (var item in ds)
            {
                item.incLev();
            }
        }
        public int contains(List<DimBox> ds, DimBox b)
        {
            int i = 0;
            foreach (var item in ds)
            {
                if (item.Equals(b)) continue;
                if (b.contains(item)) i++;
            }
            return i;
        }
        public void setLevel(List<DimBox> ds)
        {
            if (ds.Count == 0) return;
            //DimBox db = ds.ElementAt(0);
            List<DimBox> tmp = new List<DimBox>();
            incLevel(ds);
            foreach (var item in ds)
            {
                //if (item.Equals(db)) continue;
                isDisjoint(item);
                if (disj.Contains(item)) continue;
                int c = contains(ds, item);
                if (c > 0) tmp.Add(item);
                //item.incLev();
                //if (item.contains(db))
                //{
                //    tmp.Add(item);
                //}
                //else if (db.contains(item))
                //{
                //    tmp.Add(item);
                //    //if (tmp.Contains(db)) tmp.Add(item);
                //    //else { tmp.Add(db); db = item; }
                //}
                //db = item;
            }
            setLevel(tmp);
        }
    }
    public class DimBox: BoxBase
    {
        LinearGeneralDimension dim;
        LineSegment2d vl;
        public static double offset = 0.8;
        public static double bOffset = 0.1;
        public int lev;
        int o;
        public double d;
        public DimBox (LinearGeneralDimension dim, Point2d pt)
        {
            this.dim = dim;
            vl = dim.DimensionLine as LineSegment2d;
            v = vl.Direction.AsVector();
            Vector2d vec = pt.VectorTo(vl.StartPoint);
            u.normal(v, vec, out n);
            o = u.octant(n);
            d = n.Length;
            prv = new Projection(v);
            prv.set(vl.StartPoint, vl.EndPoint);
            prn = new Projection(n);
            prn.set(vl.StartPoint, pt);
            //draw(d.Parent);
        }
        public void move(DimBox db, DrawingView dv)
        {
            u.setDist(n, db.d - d + offset);
            var pt = I.CP2d(dim.Text.Origin);
            pt.TranslateBy(n);
            dim.Text.Origin = pt;
            //if (n.IsParallelTo(I.CV2d(0,1)))
            //{
            //    u.setDist(n, db.d - d + offset * 2);
            //    pt = I.CP2d(dv.Label.Position);
            //    pt.TranslateBy(n);
            //    dv.Label.Position = pt;
            //}
        }
        public void setPos(DrawingView dv)
        {
            
            if (o == 1)
            {
                setPos(dv.Center, dv.Width, true);
            }
            else if (o == 4)
            {
                setPos(dv.Center, -dv.Width, true);
            }
            else if (o == 3)
            {
                setPos(dv.Center, dv.Height, false);
            }
            else if (o == 7)
            {
                setPos(dv.Center, -dv.Height, false);
            }

        }
        public void setPos(Point2d pt, double w, bool x)
        {
            Point2d tp = dim.Text.Origin;
            double d;
            var pt1 = x ? I.CP2d(pt, I.CV2d(w * 0.5, 0)) : I.CP2d(pt, I.CV2d(0, w * 0.5));
            prn.add(pt1);
            d = prn.get();
            var v1 = x ? u.normal(vl, tp).Length : -u.normal(vl, tp).Length;
            d -= v1;
            d = w < 0 ? d - bOffset - lev * offset : d + bOffset + lev * offset;
            tp = x ? I.CP2d(d, u.dotProduct(prv.v, tp)) : I.CP2d(u.dotProduct(prv.v, tp), d);
            dim.Text.Origin = tp;
            u.alignDim(dim);
        }
        public void decLev()
        {
            lev--;
        }
        public void incLev()
        {
            lev++;
        }
        public void setNullLev()
        {
            lev = -1;
        }
        public bool contains(DimBox db)
        {
            return prv.contains(db.prv);
        }
        public bool disjoint(DimBox db)
        {
            return prv.disjoint(db.prv);
        }
        public bool check(DimBox db)
        {
            return this.o == db.o;
        }
        public DimBox getMax(DimBox db)
        {
            var pr = prv.getMax(db.prv, (a, b) => a >= b);
            return pr.Equals(prv) ? this : db;
        }
        public DimBox getMin(DimBox db)
        {
            var pr = prv.getMax(db.prv, (a, b) => a < b);
            return pr.Equals(prv) ? this : db;
        }
        public void draw(Sheet sh)
        {
            var v = u.normal(vl, dim.Text.Origin);
            d = v.Length;
            base.set(dim.Text.RangeBox);
            var l = I.tg.CreateLine2d(mid, v.AsUnitVector());
            var pt1 = I.CP2d(min); var pt2 = I.CP2d(max);
            var mtx = u.mirrorMtx(l);
            u.mirror(mtx, pt1); u.mirror(mtx, pt2);
            sh.DrawingNotes.GeneralNotes.AddFitted(pt1, "m1");
            sh.DrawingNotes.GeneralNotes.AddFitted(pt2, "M1");
            //sh.DrawingNotes.GeneralNotes.AddFitted(dim.Text.Origin, "c");
        }
    }
    public class MyBox
    {
        public readonly MyColumn col;
        public readonly MyRow row;
        Point2d min, max, mid;
        List<DimBox> dims = new List<DimBox>();
        List<MyBox> sibl = new List<MyBox>();
        Box2d b;
        public MyBox(MyColumn c, MyRow r)
        {
            col = c; row = r;
        }
        public void fill(Point2d pt)
        {
            min = I.CP2d(pt);
            max = I.CP2d(pt.X + col.w, pt.Y + row.h);
            mid = u.midPt(min, max);
            b = I.Box(min, max);
        }
        public void draw(Sheet sh, string num)
        {
            //sh.DrawingNotes.GeneralNotes.AddFitted(min, "min" + num);
            //sh.DrawingNotes.GeneralNotes.AddFitted(max, "max" + num);
        }
        public void fillDims(IEnumerable<LinearGeneralDimension> ds, DrawingView dv)
        {                    
            foreach (var item in ds)
            {
                LineSegment2d dl = item.DimensionLine as LineSegment2d;
                var min = dl.StartPoint;
                var max = dl.EndPoint;
                if (b.Contains(min) && b.Contains(max))
                    dims.Add(new DimBox(item, dv.Position));
            }
            //if (dims.Count > 0) sh.DrawingNotes.GeneralNotes.AddFitted(mid, dims.Count.ToString());
        }
        public void arrange(DrawingView dv)
        {
            ArrangeDims ad = new ArrangeDims(dv, dims);
            ad.sort();
            ad.setLevel();
            ad.setPos();
        }
    }
    public class MyColumn
    {
        public int i;
        public readonly double w;
        public MyColumn(int ind, double w)
        {
            this.w = w; i = ind;
        }
    }
    public class MyRow
    {
        public int i;
        public readonly double h;
        public MyRow(int ind, double h)
        {
            this.h = h; i = ind;
        }
    }
    public class MyBoxes
    {
        Sheet sh;
        DrawingView dv;
        List<MyColumn> cols = new List<MyColumn>();
        List<MyRow> rows = new List<MyRow>();
        Box2d b, b1;
        List<MyBox> boxes = new List<MyBox>();
        IEnumerable<LinearGeneralDimension> dims;
        //Point2d bp;
        public MyBoxes(DrawingView dv)
        {
            this.dv = dv; sh = dv.Parent;
            dims = u.getDims(dv);
        }
        public void setBox(double x1, double y1, double x2, double y2)
        {
            var min = I.CP2d(x1, y1); var max = I.CP2d(sh.Width - x2, sh.Height - y2);
            b = I.Box(min, max);
            var pt1 = I.CP2d(dv.Center, I.CV2d(-dv.Width / 2, -dv.Height / 2));
            var pt2 = I.CP2d(dv.Center, I.CV2d(dv.Width / 2, dv.Height / 2));
            b1 = I.Box(pt1, pt2);
        }
        public void setBox()
        {
            var min = I.CP2d(dv.Center, I.CV2d(-dv.Width*1.5, -dv.Height*1.5));
            var max = I.CP2d(dv.Center, I.CV2d(dv.Width*1.5, dv.Height*1.5));
            b = I.Box(min, max);
            var pt1 = I.CP2d(dv.Center, I.CV2d(-dv.Width / 2, -dv.Height / 2));
            var pt2 = I.CP2d(dv.Center, I.CV2d(dv.Width / 2, dv.Height / 2));
            b1 = I.Box(pt1, pt2);
        }
        public void fill()
        {
            foreach (var item in boxes)
            {
                item.fillDims(dims, dv);
            }
        }
        public void arrange()
        {
            foreach (var item in boxes)
            {
                if (item.col.Equals(cols[1]) && item.row.Equals(rows[1])) continue;
                item.arrange(dv);
            }
        }
        public List<double> getGrid(int i, Vector2d n)
        {
            var v1 = b.MinPoint.VectorTo(b1.MinPoint);
            var d1 = v1.DotProduct(n);

            var lst = new List<double>() { d1};

            v1 = b1.MinPoint.VectorTo(b1.MaxPoint);
            d1 = v1.DotProduct(n);

            for (int j = 0; j < i; j++)
            {
                lst.Add(d1 / i);
            }

            v1 = b1.MaxPoint.VectorTo(b.MaxPoint);
            d1 = v1.DotProduct(n);

            lst.Add(d1);
            //var vec = I.CV2d(dv.Width / 2, dv.Height / 2);
            //var pt = I.CP2d(dv.Center, vec);
            //sh.DrawingNotes.GeneralNotes.AddFitted(pt, "s1");
            return lst;
        }
        public void create(List<double> cs, List<double> rs)
        {
            var pt = I.CP2d(b.MinPoint);
            double y = pt.Y;
            for (int i = 0; i < cs.Count; i++)
            {
                cols.Add(new MyColumn(i, cs[i]));
            }
            for (int j = 0; j < rs.Count; j++)
            {
                rows.Add(new MyRow(j, rs[j]));
            }
            int k = 0;
            foreach (var c in cols)
            {
                int l = 0;
                k++;
                foreach (var r in rows)
                {
                    l++;
                    var b = new MyBox(c, r);
                    b.fill(pt);
                    boxes.Add(b);
                    b.draw(sh, k.ToString() + l.ToString());
                    pt.Y += r.h;
                }
                pt.X += c.w; pt.Y = y;
            }
        }
    }
    internal class BrowserBtn : Button
    {
        public static Drawings m_Drw;
        public static Drawings getDrw { get { return m_Drw; } }
        public BrowserBtn(string displayName, string internalName, string clientId, string description,
            string tooltip)
            : base(displayName, internalName, CommandTypesEnum.kShapeEditCmdType, clientId, description, tooltip,
                  ButtonDisplayEnum.kDisplayTextInLearningMode) { }
        protected override void ButtonDefinition_OnExecute(NameValueMap context)
        {
            findBrowserNode();
        }
        public void findBrowserNode()
        {
            var pDoc = I.aDoc();
            var ss = pDoc.SelectSet;
            if (ss.Count != 1) return;
            var br = pDoc.BrowserPanes.ActivePane;
            var se = ss[1] as SketchEntity;
            var e = ss[1] as Edge; var f = ss[1] as Face;
            PartFeature pf = null;
            if (f != null) pf = f.CreatedByFeature;
            if (e != null)
            {
                var f1 = e.Faces[1]; var f2 = e.Faces[2];
                pf = f1.Evaluator.Area <= f2.Evaluator.Area ? f1.CreatedByFeature : f2.CreatedByFeature;
            }
            if (se != null)
            {
                var ps = se.Parent;
                var node1 = br.GetBrowserNodeFromObject((object)ps);
                node1.EnsureVisible();
                node1.DoSelect();
                return;
            }
            if (pf == null) return;
            var node = br.GetBrowserNodeFromObject((object)pf);
            node.EnsureVisible();
            node.DoSelect();
        }
    }
}
