using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Inventor;
using InvDoc;

namespace InvAddIn
{
    public class CopyPaste
    {
        static public List<feature> from = new List<feature>();
        static public List<feature> to = new List<feature>();
        public defData o, n;
        List<PartFeature> pfs = new List<PartFeature>();
        CommandManager cMgr;
        InteractionEvents intEv, selPlEv;
        KeyboardEvents key;
        UserInterfaceManager uiMgr;
        SelectEvents sEv, sPlEv;
        WorkPlane wp;
        bool doEv = false, stop = false;
        object a, l;
        public CopyPaste(Document doc)
        {
            cMgr = I.app.CommandManager;
            uiMgr = I.app.UserInterfaceManager;
            o = new defData(doc);
        }
        public void add(SelectSet ss)
        {
            clear();
            clearEvts();
            foreach (var item in ss)
            {
                var fl = item as FlangeFeature;
                if (fl != null)
                {
                    var f = new Flange(fl);
                    addDepend(f);
                }
                var cff = item as ContourFlangeFeature;
                if (cff != null)
                {
                    var cf = new ContourFlange(cff);
                    addDepend(cf);
                }
                var fa = item as Face;
                if (fa != null)
                {
                    var f = new fFeat(fa);
                    //f.fillFaces();
                    from.Add(f);
                    //addDepend(f);
                }
                var sp = item as SketchPoint;
                if (sp != null)
                {
                    var sk = new sPoints(sp);
                }
            }
        }
        public void addDepend(feature f)
        {
            f.fillFaces();
            from.Add(f);
            var lst = getFeatures(f); from.AddRange(lst);
            //lst = getPunches(f); from.AddRange(lst);
        }
        public void clear()
        {
            from.Clear();
            to.Clear();
            wp = null;
        }
        public void create()
        {
            to.Clear();
            PartFeature pf = null;
            n = new defData(I.aDoc());
            foreach (var item in from)
            {
                if (item is Flange)
                {
                    pf = addFlange(item as Flange);
                    if (pf != null) pfs.Add(pf);
                }
                if (item is ContourFlange cff)
                {
                    pf = addContourFlange(cff);
                    if (pf != null) pfs.Add(pf);
                }
                if (item is Hole || item is Punch)
                {
                    pf = item.add(to[0]);
                    if (pf != null) pfs.Add(pf);
                }
                if (item is fFeat)
                {
                    pf = addFaceFeat(item as fFeat);
                    if (pf != null) pfs.Add(pf);
                }
            }
            if (pfs.Count > 0 && wp != null)
            {
                u.addMirror(n.doc, wp, pfs);
            }
        }
        public object getValue(FeatureDimensions dims, FeatureDimensionTypeEnum t)
        {
            foreach (FeatureDimension item in dims)
            {
                if (item.DimensionType == t) return item.Parameter.ModelValue;
            }
            return null;
        }
        public PartFeature addContourFlange(ContourFlange from)
        {
            if (n.smcd == null) return null;
            createEvts();
            if (n.eCol.Count == 0) return null;
            var e = n.eCol[1] as Edge;
            var g = e.Geometry as LineSegment;
            var f = u.get<Face>(getVertex(e).Faces, el => el.SurfaceType == SurfaceTypeEnum.kPlaneSurface
            && ((Plane)el.Geometry).Normal.IsParallelTo(g.Direction));
            var cf = from.add(n.doc, f, e);
            var nccf = new ContourFlange(cf);
            to.Add(nccf);
            return cf as PartFeature;
        }
        public Vertex getVertex(Edge e)
        {
            Vertex v = e.StartVertex;
            int n = 0;
            foreach (Edge item in v.Edges)
            {
                if (item.CurveType == CurveTypeEnum.kLineCurve) n++;
            }
            if (n == v.Edges.Count) return v;
            return e.StopVertex;
        }
        public PartFeature addFlange(Flange from)
        {
            if (n.smcd == null) return null;
            createEvts();
            a = getValue(from.fl.FeatureDimensions, FeatureDimensionTypeEnum.kAngleFeatureDimension);
            l = getValue(from.fl.FeatureDimensions, FeatureDimensionTypeEnum.kLinearFeatureDimension);
            var nFl = new Flange(n.doc);
            nFl.add(n.eCol, l, a);
            to.Add(nFl);
            return nFl.fl as PartFeature;
        }
        public PartFeature addFaceFeat(fFeat f)
        {
            if (n.smcd == null) return null;
            if (f.pf is FlangeFeature)
            {
                createEvts();
                a = getValue(f.pf.FeatureDimensions, FeatureDimensionTypeEnum.kAngleFeatureDimension);
                l = getValue(f.pf.FeatureDimensions, FeatureDimensionTypeEnum.kLinearFeatureDimension);
                var nFl = new fFeat(n.doc);
                nFl.add(n.eCol, l, a);
                to.Add(nFl);
                nFl.cp(f.holes);
                nFl.create();
                return nFl.f as PartFeature;
            }
            else
            {
                createEvts();
                var nFl = new fFeat(n.doc);
                nFl.setEdge(n.eCol);
                //nFl.add(n.eCol, l, a);
                to.Add(nFl);
                nFl.cp(f.holes);
                nFl.create();
            }
            return null;
        }
        public void addToList(PartFeature pf, feature feat, ref List<feature> lst)
        {
            Face fa = null;
            if (feat is fFeat)
            {
                fa = ((fFeat)feat).f;
            }
            if (pf is HoleFeature)
            {
                var hole = new Hole(pf as HoleFeature);
                if (hole.check(feat))
                {
                    lst.Add(hole);
                    if (fa == null)
                    {
                        fa = hole.hf.Sketch.PlanarEntity as Face;
                        hole.plIndex = hole.of.findFaceIndex(fa);
                    }
                    else hole.plIndex = 1;
                }
            }
            else if (pf is PunchToolFeature)
            {
                var punch = new Punch(pf as PunchToolFeature);
                if (punch.check(feat))
                {
                    lst.Add(punch);
                    if (fa == null)
                    {
                        fa = punch.ops.PlanarEntity as Face;
                        punch.plIndex = punch.of.findFaceIndex(fa);
                    }
                    else punch.plIndex = 1;
                }
            }
        }
        public List<feature> getFeatures(feature feat)
        {
            List<feature> lst = new List<feature>();
            var f = feat as fFeat;
            if (f != null)
            {
                foreach (Edge e in f.f.Edges)
                {
                    Face f2 = u.get<Face>(e.Faces, el => !el.Equals(f.f));
                    if (f2 == null) continue;
                    addToList(f2.CreatedByFeature, feat, ref lst);
                }
            }
            else
            {
                foreach (PartFeature pf in feat.pf.Parent.SurfaceBodies[1].AffectedByFeatures)
                {
                    addToList(pf, feat, ref lst);
                }
            }
            return lst;
        }

        public void clearEvts()
        {
            if (intEv != null) intEv = null;
        }

        public void createEvts()
        {
            if (intEv == null)
                intEv = cMgr.CreateInteractionEvents();
            intEv.InteractionDisabled = false;
            sEv = intEv.SelectEvents;
            selectEdge();
            sEv.OnPreSelect += SEv_OnPreSelect;
            sEv.OnSelect += SEv_OnSelect;
            key = intEv.KeyboardEvents;
            key.OnKeyDown += Key_OnKeyDown;
            intEv.Start();
            while (!doEv) uiMgr.DoEvents();
            foreach (var item in sEv.SelectedEntities)
            {
                if (item is Edge)
                    n.eCol.Add(item);
            }
            intEv.Stop();
            sEv.OnPreSelect -= SEv_OnPreSelect;
            sEv.OnSelect -= SEv_OnSelect;
            doEv = false;
        }

        public void initEvent(InteractionEvents iEv, string txt, SelectionFilterEnum f)
        {
            if (iEv == null) {
                iEv = cMgr.CreateInteractionEvents();
                iEv.SelectEvents.ClearSelectionFilter();
            }
            iEv.InteractionDisabled = false;
            iEv.StatusBarText = txt;
            var sel = iEv.SelectEvents;
            sel.AddSelectionFilter(f);
            
        }
        public List<object> runEvents(InteractionEvents iEv)
        {
            List<object> lst = new List<object>();
            var sel = iEv.SelectEvents;
            sel.OnPreSelect += Sel_OnPreSelect;
            sel.OnSelect += Sel_OnSelect;
            iEv.Start();
            while (!stop) uiMgr.DoEvents();
            foreach (var item in sel.SelectedEntities)
            {
                lst.Add(item);
            }
            iEv.Stop();
            stop = false;
            sel.OnPreSelect -= Sel_OnPreSelect;
            sel.OnSelect -= Sel_OnSelect;
            return lst;
        }

        private void Sel_OnSelect(ObjectsEnumerator JustSelectedEntities, SelectionDeviceEnum SelectionDevice, 
            Point ModelPosition, Point2d ViewPosition, View View)
        {
        }

        private void Sel_OnPreSelect(ref object PreSelectEntity, out bool DoHighlight, 
            ref ObjectCollection MorePreSelectEntities, SelectionDeviceEnum SelectionDevice, Point ModelPosition,
            Point2d ViewPosition, View View)
        {
            DoHighlight = true;
        }

        public void selectEdge()
        {
            sEv.AddSelectionFilter(SelectionFilterEnum.kPartEdgeLinearFilter);
            //sEv.AddSelectionFilter(SelectionFilterEnum.kWorkPlaneFilter);
            intEv.StatusBarText = "Выберите ребро";
        }

        public void selectPlane()
        {
            selPlEv = cMgr.CreateInteractionEvents();
            selPlEv.StatusBarText = "Выберите плоскость для зеркала";
            sEv.AddSelectionFilter(SelectionFilterEnum.kWorkPlaneFilter);
            intEv.Start();
        }

        private void SEv_OnSelect(ObjectsEnumerator JustSelectedEntities, SelectionDeviceEnum SelectionDevice, 
            Point ModelPosition, Point2d ViewPosition, View View)
        {
        }

        private void Key_OnKeyDown(int Key, ShiftStateEnum ShiftKeys)
        {
            if (Key == 32)
            {
                doEv = true;
            }
            if (Key == 77)
            {
                u.changePlane(ref wp, n.smcd as PartComponentDefinition);       
            }
        }

        private void SEv_OnPreSelect(ref object PreSelectEntity, out bool DoHighlight,
            ref ObjectCollection MorePreSelectEntities, SelectionDeviceEnum SelectionDevice,
            Point ModelPosition, Point2d ViewPosition, View View)
        {
            DoHighlight = true;
        }
    }

    public class defData
    {
        public SheetMetalComponentDefinition smcd;
        public SheetMetalFeatures smf;
        public Document doc;
        public PlanarSketch ps;
        public double t, r;
        public EdgeCollection eCol;
        public defData(Document doc)
        {
            this.doc = doc;
            getSMCD(doc);
        }
        public void getSMCD(Document doc)
        {
            smcd = I.getSMCD(doc);
            if (smcd == null) return;
            smf = smcd.Features as SheetMetalFeatures;
            t = smcd.Thickness.ModelValue;
            r = smcd.BendRadius.ModelValue;
            eCol = I.objs.CreateEdgeCollection();
        }
    }

    public class feature
    {
        protected defData n, o;
        protected PlanarSketch ps;
        public EdgeCollection eCol;
        public int plIndex = -1;
        public PartFeature pf;
        public List<Face> fs = new List<Face>();
        public feature(Document doc)
        {
            o = new defData(doc);
        }
        public object getValue(FeatureDimensions dims, FeatureDimensionTypeEnum t)
        {
            foreach (FeatureDimension item in dims)
            {
                if (item.DimensionType == t) return item.Parameter.ModelValue;
            }
            return null;
        }
        public bool check(Face f, Func<Face, bool> filter)
        {
            foreach (Edge e in f.Edges)
            {
                foreach (Face item in e.Faces)
                {
                    if (item.Equals(f)) continue;
                    if (filter(item)) return true;
                }
            }
            return false;
        }
        public Point getCenter(Face f)
        {
            var b = f.Evaluator.RangeBox;
            var pt = b.MinPoint;
            var v = b.MinPoint.VectorTo(b.MaxPoint);
            v.ScaleBy(0.5);
            pt.TranslateBy(v);
            return pt;
        }
        public void fillFaces()
        {
            var e = u.findStartEdge(pf);
            fs = u.findTangFaces(pf, e);
            //clientDraw();
        }
        public int findFaceIndex(Face f)
        {
            for (int i = 0; i < fs.Count; i++)
            {
                if (fs[i].Equals(f)) return i;
            }
            return -1;
        }
        public Face getFaceByIndex(int ind, out Edge e)
        {
            Face f1 = fs[ind - 1], f2 = fs[ind];
            e = u.findConnectEdge(f1, f2);
            return f2;
        }
        public void clientDraw()
        {
            int n = 0;
            foreach (var item in fs)
            {
                u.clientTxt(o.smcd as PartComponentDefinition, item.PointOnFace, n.ToString());
                n++;
            }
        }
        public Point getCenter(Edge f)
        {
            var b = f.Evaluator.RangeBox;
            var pt = b.MinPoint;
            var v = b.MinPoint.VectorTo(b.MaxPoint);
            v.ScaleBy(0.5);
            pt.TranslateBy(v);
            return pt;
        }
        public Face getFace(PartFeature pf, Point pt,Func<Face, bool> filter)
        {
            foreach (Face item in pf.Faces)
            {
                Face f = null;
                if (check(item, filter)) f = item;
                if (f != null) return item;
                if (pt != null)
                {
                    foreach (Face t in f.TangentiallyConnectedFaces)
                    {
                        if (t.SurfaceType == SurfaceTypeEnum.kCylinderSurface) continue;
                        var pl = t.Geometry as Plane;
                        if (u.eq(u.distToPlane(pl, pt), 0)) return item;
                    }
                }
            }
            return null;
        }
        public Point getCenter(PartFeature pf, Func<Face, bool> filter)
        {
            var tmp = u.get(pf.Faces, filter);
            return getCenter(tmp);
        }
        public Edge find(PartFeature pf, Func<Edge, bool> filter)
        {
            foreach (Face f in pf.Faces)
            {
                foreach (Edge item in f.Edges)
                {
                    if (filter(item)) return item;
                }
            }
            return null;
        }
        public virtual void set(PartFeature pf) { }
        public virtual PartFeature add(feature f, Point pt = null)
        {
            this.pf = f.pf;
            n = new defData(pf.Parent.Document as Document);
            if (eCol == null) eCol = I.objs.CreateEdgeCollection();
            return pf;
        }
        public virtual bool check(feature feat, Face f)
        {
            if (f == null) return false;
            foreach (Face item in pf.Faces)
            {
                if (item.Equals(f)){ return true;}
            }
            return false;
        }
    }

    public class fFeat: feature
    {
        public Face f, nf = null;
        Edge be;
        public List<fPoints> holes = new List<fPoints>();
        public fFeat(Face face): base((Document)face.Parent.ComponentDefinition.Document)
        {
            f = face;
            be = u.getEdge(f, o.t, o.r);
            pf = f.CreatedByFeature;
            var es = u.gets<Edge>(f.Edges, el => el.GeometryType == CurveTypeEnum.kCircleCurve);
            var gr = es.GroupBy(el => ((Circle)el.Geometry).Radius);
            foreach (var item in gr)
            {
                fHoles h = new fHoles(f, item.Key, item, be);
                holes.Add(h);
            }
            var eloops = u.gets<EdgeLoop>(f.EdgeLoops, el => el.Edges.Count == 4);
            var gr1 = eloops.GroupBy(el => u.slotGroup(el));
            foreach (var item in gr1)
            {
                if (item.Key == null) continue;
                fPunch fp = new fPunch(f, 0, item, be, item.Key);
                holes.Add(fp);
            }
            //es = u.gets<Edge>(f.Edges, el => el.GeometryType == CurveTypeEnum.kCircularArcCurve);
            //var gr1 = es.GroupBy(el => u.get<Face>(el.Faces, r => !r.Equals(f)).CreatedByFeature);
            //foreach (var item in gr1)
            //{
            //    if (item.Key is PunchToolFeature) {
            //        fPunch fp = new fPunch(f, 0, item, be, (PunchToolFeature)item.Key);
            //        holes.Add(fp);
            //    }
            //}
        }
        public fFeat(Document doc) : base(doc) { }
        public PartFeature add(EdgeCollection col, object l, object a)
        {
            if (a == null || l == null) return null;
            var def = o.smf.FlangeFeatures.CreateFlangeDefinition(col, a, l);
            var fl = o.smf.FlangeFeatures.Add(def);
            pf = fl as PartFeature;
            return pf;
        }
        public void setEdge(EdgeCollection col)
        {
            be = col[1] as Edge;
            var fs = u.gets<Face>(be.Faces, el => el.SurfaceType == SurfaceTypeEnum.kPlaneSurface);
            f = fs.OrderByDescending(el => el.Evaluator.Area).FirstOrDefault();
        }
        public void cp(List<fPoints> old)
        {
            holes = old;
            foreach (var item in holes)
            {
                item.setNew(o);
            }
        }
        public void create()
        {
            Face fa = null;
            Edge en = null;
            if (pf is FlangeFeature)
                fa = pf.Faces[5];
            else
            {
                fa = f;
                en = be;
            }
            foreach (var item in holes)
            {
                if (en != null) item.ne = en;
                item.copySketch(fa, pf);
                item.add();
            }
        }

    }

    public class fPoints: feature
    {
        public Face of, nf;
        protected double d;
        protected List<Point> pts = new List<Point>();
        protected Matrix mtx;
        protected Point pt;
        protected Vertex ov, nv;
        protected SketchCP scp;
        public Edge ne, oe;
        protected PlanarSketch nps;
        protected Point norigin;
        protected ObjectCollection col;
        public UnitVector ox, oy, nx, ny;
        public UnitVector2d ox2, oy2, nx2, ny2;
        public Point2d o2, n2;
        public fPoints(Face face, double r, Edge be) : base((Document)face.Parent.ComponentDefinition.Document)
        {
            of = face;
            oe = be;
            setMatrix(be);
            this.d = r * 2;
            pt = ov.Point;
            //pt.TransformBy(mtx);
            //foreach (Edge e in hs)
            //{
            //    Circle c = e.Geometry as Circle;
            //    pt = c.Center;
            //    pt.TransformBy(mtx);
            //    pts.Add(pt);
            //}
            //var o = I.CP();
            //pts.Sort((e1, e2) => e1.DistanceTo(o).CompareTo(e2.DistanceTo(o)));

        }
        public void setNew(defData d)
        {
            base.n = d;
        }
        public void setMatrix(Edge e)
        {
            mtx = I.getMatrix();
            ov = e.StartVertex;
            var e2 = this.getEdge(e, ov);
            Vector vx = u.getVector(e, ov.Point), vz = u.getVector(e2, ov.Point);
            Vector vy = vx.CrossProduct(vz);

            vx.Normalize(); vy.Normalize(); vz.Normalize();
            //Point x = ov.Point, y = ov.Point, z = ov.Point;
            //x.TranslateBy(vx); y.TranslateBy(vy); z.TranslateBy(vz);
            //PartComponentDefinition def = o.smcd as PartComponentDefinition;
            //u.clientTxt(def, x, "vx");
            //u.clientTxt(def, y, "vy");
            //u.clientTxt(def, z, "vz");
            //mtx.SetCoordinateSystem(ov.Point, vx, vy, vz);
            mtx.SetToAlignCoordinateSystems(ov.Point, vx, vy, vz, I.CP(), I.CV(1, 0, 0), I.CV(0, 1, 0), I.CV(0, 0, 1));
        }
        public Edge getEdge(Edge e, Vertex v)
        {
            var r = u.get<Edge>(v.Edges, el => e.CurveType == CurveTypeEnum.kLineCurve && !e.Equals(el) && u.eq(u.len(el), o.t));
            return r;
        }
        public fPoints(Document doc) : base(doc) { }
        public PartFeature add(EdgeCollection col, object l, object a)
        {
            if (a == null || l == null) return null;
            var def = o.smf.FlangeFeatures.CreateFlangeDefinition(col, a, l);
            var fl = o.smf.FlangeFeatures.Add(def);
            pf = fl as PartFeature;
            return pf;
        }
        public void copySketch(Face f, PartFeature nf)
        {
            if (ne == null)
                ne = u.getEdge(f, n.t, n.r);
            setNew(f);
            scp = new SketchCP(nps);
            //scp.setVector(o2);
            double l = u.len(ne) / u.len(oe);
            scp.setAlignMtx(l);
            var pt = pts[0];
            var p2 = I.CP2d(pt.X, pt.Y);
            p2.TransformBy(scp.mtx);
            scp.checkDir(p2);
            scp.addPoints(pts);
            scp.addConstrPt();
            scp.addDimsPt();
            //getEnts(nf);

            col = I.COC();
            foreach (SketchPoint item in nps.SketchPoints)
            {
                if (item.HoleCenter)
                    col.Add(item);
            }
        }
        public void setNew(Face f)
        {
            if (ne == null)
                ne = u.getEdge(f, n.t, n.r);
            var norigin = ne.StartVertex;
            this.norigin = norigin.Point;
            nps = n.smcd.Sketches.AddWithOrientation(f, ne, true, true, norigin, false);
        }
        public virtual void add()
        {

        }
    }

    public class fHoles : fPoints
    {
        
        public fHoles(Face face, double r, IEnumerable<Edge> hs, Edge be): base(face, r, be)
        {
            of = face;
            oe = be;
            setMatrix(be);
            this.d = r*2;
            pt = ov.Point;
            pt.TransformBy(mtx);
            foreach (Edge e in hs)
            {
                Circle c = e.Geometry as Circle;
                pt = c.Center;
                pt.TransformBy(mtx);
                pts.Add(pt);
            }
            var o = I.CP();
            pts.Sort((e1, e2) => e1.DistanceTo(o).CompareTo(e2.DistanceTo(o)));
           
        }
        public fHoles(Document doc) : base(doc) { }
        public override void add()
        {
            if (col.Count == 0) return;
            var def = n.smf.HoleFeatures.CreateSketchPlacementDefinition(col);

            var nhf = n.smf.HoleFeatures.AddDrilledByDistanceExtent(def, d, n.t * 3, PartFeatureExtentDirectionEnum.kPositiveExtentDirection);
            //u.changePFName(this.hf as PartFeature, nhf as PartFeature);
            //return nhf as PartFeature;
        }
    }

    public class fPunch : fPoints
    {
        public PunchToolFeature punch, npunch;
        double d, l, ang = 0;
        Vector dir, bdir;

        public fPunch(Face face, double r, Edge be, PunchToolFeature p) : base(face, r, be)
        {
            punch = p;
            of = face;
            oe = be;
            setMatrix(be);
            this.d = r * 2;
            pt = ov.Point;
            pt.TransformBy(mtx);
            foreach (SketchPoint sp in p.PunchCenterPoints)
            {
                pt = sp.Geometry3d;
                pt.TransformBy(mtx);
                pts.Add(pt);
            }
            var o = I.CP();
            pts.Sort((e1, e2) => e1.DistanceTo(o).CompareTo(e2.DistanceTo(o)));

        }
        public fPunch(Face face, double r, IEnumerable<EdgeLoop> l, Edge be, double? gr): base(face, r, be)
        {
            of = face;
            oe = be;
            bdir = be.StartVertex.Point.VectorTo(be.StopVertex.Point);
            setMatrix(be);
            var data = u.slotData(l.ElementAt(0));
            this.d = data[0];
            this.l = data[1];
            dir = u.slotDir(l.ElementAt(0));
            foreach (var item in l)
            {
                pt = u.slotCenter(item);
                pt.TransformBy(mtx);
                pts.Add(pt);
            }
            var o = I.CP();
            pts.Sort((e1, e2) => e1.DistanceTo(o).CompareTo(e2.DistanceTo(o)));
        }
        public fPunch(Document doc) : base(doc) { }
        public void fill(iFeatureDefinition def)
        {
            int i = 1;
            foreach (iFeatureParameterInput item in def.iFeatureInputs)
            {
                var input = punch.iFeatureDefinition.iFeatureInputs;
                item.Expression = (((iFeatureParameterInput)input[i]).Parameter.ModelValue * 10).ToString();
                i++;
            }
        }
        public override void add()
        {
            if (this.d != 0 && this.l != 0)
            {
                Dictionary<string, double> dict = new Dictionary<string, double>()
                {
                    {"ширина", this.d*10}, {"a", this.l*10}
                };
                if (dir.IsParallelTo(bdir)) ang = Math.PI / 2;
                I.createPunch("Punches/Овал.ide", dict, col, (PartComponentDefinition)n.smcd, ang);
                return;
            }
            if (col.Count == 0) return;
            var def = n.smf.PunchToolFeatures.
                CreateiFeatureDefinition(punch.iFeatureTemplateDescriptor.LastKnownSourceFileName);
            fill(def);
            npunch = n.smf.PunchToolFeatures.Add(col, def, AcrossBends: true);
        }
    }

    public class Flange : feature
    {
        public FlangeFeature fl;
        public Flange(FlangeFeature f): base(f.Parent.Document as Document)
        {
            fl = f;
            pf = f as PartFeature;
        }
        public Flange(Document doc) : base(doc) { }
        public PartFeature add(EdgeCollection col ,object l, object a)
        {
            if (a == null || l == null) return null;
            var def = o.smf.FlangeFeatures.CreateFlangeDefinition(col, a, l);
            fl = o.smf.FlangeFeatures.Add(def);
            pf = fl as PartFeature;
            return pf;
        }
    }

    public class sPoints
    {
        PlanarSketch ps;
        SketchCP scp;
        SketchPoint sp;
        List<SketchPoint> pts = new List<SketchPoint>();
        public sPoints(SketchPoint sp)
        {
            this.sp = sp;
            this.ps = sp.Parent as PlanarSketch;
            this.scp = new SketchCP(ps);
            scp.addConstrPt(sp);
            scp.addDimsPt(sp, false);
        }
    }

    public class ContourFlange : feature
    {
        public ContourFlangeFeature cff, ncff;
        public ContourFlangeDefinition def, ndef;
        public PlanarSketch ops, nps;
        public Vertex oorigin, norigin;
        public Face oface, nface;
        public ContourFlange(ContourFlangeFeature f): base(f.Parent.Document as Document)
        {
            cff = f;
            pf = cff as PartFeature;
            def = cff.Definition;
        }
        public ContourFlangeFeature add(Document doc, Face f, Edge e)
        {
            nface = f;
            n = new defData(doc);
            var d = u.len(e);
            norigin = getOrigin(f, e);
            if (norigin == null) return null;

            var x = getThinEdge(f, norigin, n.t);

            setOldPS();
            //var oldSE = (SketchEntity)def.Path[1].SketchEntity;
            //var oldPs = oldSE.Parent as PlanarSketch;
            //var oldFace = oldPs.PlanarEntity as Face;
            //var oldOrigin = getOrigin(oldFace, SketchCP.getOrigin(oldSE));
            //var oldX = getThinEdge(oldFace, oldOrigin, o.t);
            nps = n.smcd.Sketches.AddWithOrientation(f, x, true, true, norigin, false);

            SketchCP sk = new SketchCP(nps);

            Point2d oorigin2d, norigion2d;
            UnitVector2d ox2d, oy2d, nx2d, ny2d;
            oorigin = u.setOrigin(oface, oorigin.Point, o.t, ops, out oorigin2d, out ox2d, out oy2d);
            norigin = u.setOrigin(nface, norigin.Point, n.t, nps, out norigion2d, out nx2d, out ny2d);

            sk.setAlignMtx(oorigin2d, ox2d, oy2d, norigion2d, nx2d, ny2d);
          
            sk.addProj(e);
            sk.addLines(def);
            sk.addConstr();
            sk.addDims();
            var path = n.smf.CreatePath(nps.SketchLines[1]);
            ndef = n.smf.ContourFlangeFeatures.CreateContourFlangeDefinition(path);
            ndef.SetDistanceExtent(d, PartFeatureExtentDirectionEnum.kNegativeExtentDirection);
            ncff = n.smf.ContourFlangeFeatures.Add(ndef);
            return ncff;
        }
        public Vertex setOrigin(Face f, ref Point pt, double t, PlanarSketch ps,
            out Point2d o2, out UnitVector2d x2, out UnitVector2d y2)
        {
            UnitVector x, y;
            var vert = u.getOrigin(f, pt, t, out x, out y);
            u.getOrigin2d(ps, pt, x, y, out o2, out x2, out y2);
            return vert;
        }
        public void setOldPS()
        {
            var se = def.Path[1].SketchEntity as SketchEntity;
            ops = se.Parent as PlanarSketch;
            oface = ops.PlanarEntity as Face;
            oorigin = def.Path.Wires[1].Edges[1].StartVertex;
            //u.clientTxt(o.smcd as PartComponentDefinition, oorigin.Point, "oldWire");
        }
        public Vertex getOrigin(Face f, Edge e)
        {
            return u.get<Vertex>(f.Vertices, el => u.eq(el, e.StartVertex) || u.eq(el, e.StopVertex));
        }
        public Vertex getOrigin(Face f, Point pt)
        {
            return u.get<Vertex>(f.Vertices, el => u.eq(el.Point, pt));
        }
        public static Edge getThinEdge(Face f, Vertex origin, double t)
        {
            return u.get<Edge>(f.Edges, el =>
            (el.StartVertex.Equals(origin) || el.StopVertex.Equals(origin)) &&
               u.eq(u.len(el), t)
            );
        }
    }

    public class Hole : feature
    {
        public HoleFeature hf, nhf;
        public PlanarSketch ops, nps;
        public Point norigin, oorigin;
        public UnitVector ox, oy, nx, ny;
        public UnitVector2d ox2, oy2, nx2, ny2;
        public Point2d o2, n2;
        public feature of, nf; //old / new feature
        public Face f;
        public Edge oe, ne;
        //public bool thick = false;
        ObjectCollection col;
        SketchCP scp;
        public Hole(HoleFeature f): base(f.Parent.Document as Document)
        {
            hf = f;
        }
        public bool check(feature feat)
        {
            of = feat;
            var ps = hf.Sketch;
            var f = ps.PlanarEntity as Face;
            return of.check(this, f);
        }
        public override PartFeature add(feature feat, Point pt = null)
        {
            this.nf = feat;
            nf.fillFaces();
            base.add(nf);
            Edge edge;
            f = nf.getFaceByIndex(plIndex, out edge);
            copySketch(f, nf.pf);
            if (this.hf.PlacementType != HolePlacementTypeEnum.kSketchPlacementType) return null;
            var def = n.smf.HoleFeatures.CreateSketchPlacementDefinition(col);
            
            switch (this.hf.HoleType)
            {
                case HoleTypeEnum.kDrilledHole:
                    nhf = n.smf.HoleFeatures.AddDrilledByDistanceExtent(def,
                        getValue(hf.FeatureDimensions, FeatureDimensionTypeEnum.kHoleFeatureDimension), hf.Depth,
                        ((DistanceExtent)hf.Extent).Direction);
                    break;
                case HoleTypeEnum.kCounterSinkHole:
                    break;
                case HoleTypeEnum.kCounterBoreHole:
                    break;
                case HoleTypeEnum.kSpotFaceHole:
                    break;
                default:
                    break;
            }
            u.changePFName(this.hf as PartFeature, nhf as PartFeature);
            return nhf as PartFeature;
        }
        public void copySketch(Face f, PartFeature nf)
        {
            setNew(f);
            scp = new SketchCP(nps);
            setOld();
            //scp.setVector(o2);
            double l = u.len(ne) / u.len(oe);
            scp.setAlignMtx(o2, ox2, oy2, n2, nx2, ny2, l);

            scp.addPoints(hf as PartFeature, hf.Sketch, this.norigin);
            scp.addConstrPt();
            scp.addDimsPt();
            //getEnts(nf);

            col = I.COC();
            foreach (SketchPoint item in nps.SketchPoints)
            {
                if (item.HoleCenter)
                col.Add(item);
            }
        }
        public void setOld()
        {
            ops = hf.Sketch;
            oe = u.getAxis(ops, o.t, o.r, o.doc, out oorigin, out ox, out oy);
            u.getOrigin2d(ops, oorigin, ox, oy, out o2, out ox2, out oy2);
        }
        public void setNew(Face f)
        {
            var e = u.getEdge(f, n.t, n.r);
            var norigin = e.StartVertex;
            this.norigin = norigin.Point;
            nps = n.smcd.Sketches.AddWithOrientation(f, e, true, true, norigin, false);
            ne = u.getAxis(nps, n.t, n.r, o.doc, out this.norigin, out nx, out ny);
            //u.clientTxt(n.smcd as PartComponentDefinition, this.norigin, "o");
            u.getOrigin2d(nps, norigin.Point, nx, ny, out n2, out nx2, out ny2);
        }
        public double getDist(Edge e, Point pt)
        {
            var tmp = getCenter(e);
            return pt.VectorTo(tmp).Length;
        }
    }

    public class Punch : feature
    {
        public PunchToolFeature punch, npunch;
        public PlanarSketch ops, nps;
        public Point norigin, oorigin;
        public UnitVector ox, oy, nx, ny;
        public UnitVector2d ox2, oy2, nx2, ny2;
        public Point2d o2, n2;
        public Face f;
        public feature of, nf;
        public Edge oe, ne;
        ObjectCollection col;
        SketchCP scp;
        public Punch(PunchToolFeature f): base(f.Parent.Document as Document)
        {
            punch = f;
        }
        public override PartFeature add(feature feat, Point pt = null)
        {
            this.nf = feat;
            nf.fillFaces();
            base.add(nf);
            Edge edge;
            f = nf.getFaceByIndex(plIndex, out edge);
            copySketch(f, nf.pf);
            var def = n.smf.PunchToolFeatures.
                CreateiFeatureDefinition(punch.iFeatureTemplateDescriptor.LastKnownSourceFileName);
            fill(def);
            npunch = n.smf.PunchToolFeatures.Add(col, def, AcrossBends: true);
            u.changePFName(punch as PartFeature, npunch as PartFeature);
            return npunch as PartFeature;
        }
        public void fill(iFeatureDefinition def)
        {
            int i = 1;
            foreach (iFeatureParameterInput item in def.iFeatureInputs)
            {
                var input = punch.iFeatureDefinition.iFeatureInputs;
                item.Expression = (((iFeatureParameterInput)input[i]).Parameter.ModelValue*10).ToString();
                i++;
            }
        }
            public bool check(feature feat)
        {
            of = feat;
            ops = ((SketchPoint)punch.PunchCenterPoints[1]).Parent as PlanarSketch;
            var f = ops.PlanarEntity as Face;
            return of.check(feat, f);
        }
        public void copySketch(Face f, PartFeature nf)
        {
            setNew(f);
            scp = new SketchCP(nps);
            setOld();
            //scp.setVector(o2);
            double l = u.len(ne) / u.len(oe);
            scp.setAlignMtx(o2, ox2, oy2, n2, nx2, ny2, l);

            scp.addPoints(punch as PartFeature, ops, this.norigin);
            scp.addConstrPt();
            scp.addDimsPt();
            //getEnts(nf);

            col = I.COC();
            foreach (SketchPoint item in nps.SketchPoints)
            {
                if (item.HoleCenter)
                    col.Add(item);
            }
        }
        public void setOld()
        {
            oe = u.getAxis(ops, o.t, o.r, o.doc, out oorigin, out ox, out oy);
            u.getOrigin2d(ops, oorigin, ox, oy, out o2, out ox2, out oy2);
        }
        public void setNew(Face f)
        {
            var e = u.getEdge(f, n.t, n.r);
            var norigin = e.StartVertex;
            this.norigin = norigin.Point;
            nps = n.smcd.Sketches.AddWithOrientation(f, e, true, true, norigin, false);
            ne = u.getAxis(nps, n.t, o.r, o.doc, out this.norigin, out nx, out ny);
            //u.clientTxt(n.smcd as PartComponentDefinition, this.norigin, "o");
            u.getOrigin2d(nps, norigin.Point, nx, ny, out n2, out nx2, out ny2);
        }
    }

    public class SketchCP
    {
        List<TwoPointDistanceDimConstraint> dims = new List<TwoPointDistanceDimConstraint>();
        PlanarSketch from, to;
        SketchEntity old;
        Vector2d v;
        public Matrix2d mtx;
        Face f;
        int x = 1, y = 1;
        List<SketchPoint> pts = new List<SketchPoint>();
        
        public SketchCP(PlanarSketch to)
        {
            this.to = to;
            f = this.to.PlanarEntity as Face;
        }
        public Point2d transform(Point2d pt)
        {
            if (v != null)
            {
                pt.TranslateBy(v);
            }
            pt.TransformBy(mtx);
            return pt;
        }

        public void setAlignMtx(Point2d o1, UnitVector2d x1, UnitVector2d y1,
            Point2d o2, UnitVector2d x2, UnitVector2d y2, double dx = 1)
        {
            mtx = I.getMatrix2d();
            mtx.SetToAlignCoordinateSystems(o1, x1.AsVector(), y1.AsVector(), o2, x2.AsVector(), y2.AsVector());
            mtx.Cell[1, 1] = mtx.Cell[1, 1] * dx;
            mtx.Cell[1, 3] = mtx.Cell[1, 3] * dx;
        }
        public void setAlignMtx(double dx = 1)
        {
            mtx = I.getMatrix2d();
            mtx.Cell[1, 1] = mtx.Cell[1, 1] * dx;
            mtx.Cell[1, 3] = mtx.Cell[1, 3] * dx;
        }
        public void setAlignMtx(PlanarSketch oldPS, Vertex o1, Vertex o2, Edge e1, Edge e2)
        {
            from = oldPS;
            Vector2d ox, oy, nx, ny;
            Point2d op, np;
            getVect(o1, e1, out ox, out oy, out op);
            getVect(o2, e2, out nx, out ny, out np);
            
            mtx.SetToAlignCoordinateSystems(op, ox, oy, I.CP2d(), nx, ny);
        }
        public void getVect(Vertex o, Edge e, out Vector2d v, out Vector2d n, out Point2d pt)
        {
            pt = from.ModelToSketchSpace(getPoint(e, o, true));
            var ep = from.ModelToSketchSpace(getPoint(e, o, false));
            v = pt.VectorTo(ep); v.Normalize();
            n = u.normal(pt, ep); n.Normalize();
        }
        public Point getPoint(Edge e, Vertex o, bool eq)
        {
            if (eq) return !u.eq(e.StartVertex, o) ? e.StopVertex.Point : e.StartVertex.Point;
            return u.eq(e.StartVertex, o) ? e.StopVertex.Point : e.StartVertex.Point;
        }

        public void addProj(object o)
        {
            to.AddByProjectingEntity(o);
        }
        public PlanarSketch addLines(ContourFlangeDefinition def)
        {
            int i = 1;
            foreach (PathEntity p in def.Path)
            {
                var e = def.Path.Wires[1].Edges[i];
                var pt1 = transform(getPoint(p.SketchEntity, e, true));
                var pt2 = transform(getPoint(p.SketchEntity, e, false));
                if (e.CurveType == CurveTypeEnum.kLineCurve)
                    addLine(pt1, pt2);
                else if (e.CurveType == CurveTypeEnum.kCircleCurve)
                {
                    var a = p.SketchEntity as SketchArc;
                    addArc(pt1, pt2, transform(a.CenterSketchPoint.Geometry));
                }
                i++;
            }
            return to;
        }
        public IEnumerable<SketchPoint> sort(PartFeature pf, Point pt)
        {
            if (pf is HoleFeature)
                return ((HoleFeature)pf).HoleCenterPoints.Cast<SketchPoint>().OrderByDescending(f => pt.DistanceTo(f.Geometry3d));
            if (pf is PunchToolFeature)
                return ((PunchToolFeature)pf).PunchCenterPoints.Cast<SketchPoint>().OrderByDescending(f => pt.DistanceTo(f.Geometry3d));
            return null;
        }
        public void addPoints(PartFeature pf, PlanarSketch ps, Point norigin)
        {
            from = ps;
            foreach (SketchPoint sp in sort(pf, norigin))
            {
                var pt = transform(sp.Geometry);
                this.pts.Add(addPoint(pt));
            }
        }
        public void addPoints(List<Point> pts)
        {
            
            foreach (var pt in pts)
            {
                this.pts.Add(addPoint(pt));
            }
        }
        public bool checkDir(Point2d p2)
        {
            if (f == null) return true;
            Point pt = this.to.SketchToModelSpace(p2);
            var rb = f.Evaluator.RangeBox;
            if (rb.Contains(pt)) return true;
            p2.Y = -p2.Y;
            pt = this.to.SketchToModelSpace(p2);
            if (rb.Contains(pt))
            {
                y = -1;
                return true;
            }
            p2.Y = -p2.Y;
            p2.X = -p2.X;
            pt = this.to.SketchToModelSpace(p2);
            if (rb.Contains(pt))
            {
                x = -1;
                return true;
            }
            p2.Y = -p2.Y;
            pt = this.to.SketchToModelSpace(p2);
            if (rb.Contains(pt))
            {
                x = -1; y = -1;
                return true;
            }
            return false;
        }
        public Point2d getPoint(Object ent, Edge e, bool start)
        {
            dynamic el = ent;
            SketchPoint pt1 = el.StartSketchPoint, pt2 = el.EndSketchPoint;
            //if (v == null)
            //{
            //    var sp = pt1.Reference ? pt1 : pt2;
            //    v = I.CV2d(-sp.Geometry.X, -sp.Geometry.Y);
            //}


            //if (e.CurveType == CurveTypeEnum.kLineCurve)
            //{
            //    SketchLine l = ent as SketchLine;
            //    pt1 = l.StartSketchPoint; pt2 = l.EndSketchPoint;
            //}
            //else (e.CurveType == CurveTypeEnum.kCircleCurve)
            //{
            //    SketchArc a = ent as SketchArc;
            //    pt1 = a.StartSketchPoint;
            //}
            if (start) return u.eq(pt1.Geometry3d, e.StartVertex.Point) ?
            pt1.Geometry : pt2.Geometry;
            return u.eq(pt1.Geometry3d, e.StopVertex.Point) ?
            pt1.Geometry : pt2.Geometry;
        }
        public void setVector(Point2d pt)
        {
            v = I.CV2d(-pt.X, -pt.Y);
        }
        public SketchPoint getSP(Point2d p1)
        {
            return u.get<SketchPoint>(to.SketchPoints, f => u.eq(f.Geometry, p1));
        }
        public void setPoint(ref object pt)
        {
            var tmp = getSP(pt as Point2d);
            if (tmp != null) pt = tmp;
        }
        SketchPoint addPoint(Point2d pt, bool hole = true)
        {
            return to.SketchPoints.Add(pt, hole);
        }
        SketchPoint addPoint(Point pt, bool hole = true)
        {
            Point2d p2 = I.CP2d(pt.X*x, pt.Y*y);
            p2.TransformBy(mtx);
            return addPoint(p2, hole);
        }
        SketchLine addLine(object pt1, object pt2)
        {
            setPoint(ref pt1); setPoint(ref pt2);
            return to.SketchLines.AddByTwoPoints(pt1, pt2);
        }
        SketchArc addArc(object pt1, object pt2, Point2d cen)
        {
            setPoint(ref pt1); setPoint(ref pt2);
            return to.SketchArcs.AddByCenterStartEndPoint(cen, pt1, pt2);
        }
        public void addConstr()
        {
            constraints.ps = to;
            foreach (SketchLine sl in to.SketchLines)
            {
                if (u.isVertical(sl))
                {
                    constraints.addVert(sl);
                }
                if (u.isHorizontal(sl))
                {
                    constraints.addHor(sl);
                }
            }
        }
        public void addConstrPt(SketchPoint tmp = null)
        {
            constraints.ps = to;
            if (tmp == null)
                tmp = to.SketchPoints[1];
            foreach (SketchPoint item in to.SketchPoints)
            {
                    pts.Add(item);
            }
            if (pts.Contains(tmp))
                pts.Remove(tmp);

            pts = pts.OrderBy(el => el.Geometry.DistanceTo(tmp.Geometry)).ToList();
            SketchPoint prev = null;
            foreach (SketchPoint sp in pts)
            {
                if (u.eq(sp.Geometry, tmp.Geometry)) continue;
                try
                {
                    addConstr(tmp, sp);
                    if (prev != null)
                    {
                        addConstr(sp, prev);
                    }
                    prev = sp;
                }
                catch (Exception)
                {
                }
            }
        }
        public void addConstr(SketchPoint sp1, SketchPoint sp2)
        {
            if (u.eq(sp1.Geometry.X, sp2.Geometry.X)) constraints.addVert(sp1, sp2);
            else if (u.eq(sp1.Geometry.Y, sp2.Geometry.Y)) constraints.addHor(sp1, sp2);
        }
        public void addDims()
        {
            foreach (SketchLine sl in to.SketchLines)
            {
                bool nov = false, noh = false;
                var v = u.get<object>(sl.Constraints, f => f is VerticalConstraint);
                if (v != null) noh = true;
                var h = u.get<object>(sl.Constraints, f => f is HorizontalConstraint);
                if (h != null) nov = true;
                if (!noh)
                {
                    var d = constraints.addTwoPointDist(sl, "", -1, 1, DimensionOrientationEnum.kHorizontalDim);
                    d.Parameter.Value = Math.Round(d.Parameter.ModelValue, 3);
                }
                if (!nov)
                {
                    var d = constraints.addTwoPointDist(sl, "", -1, 1, DimensionOrientationEnum.kVerticalDim);
                    d.Parameter.Value = Math.Round(d.Parameter.ModelValue, 3);
                }
            }
        }
        public SketchPoint addProj()
        {
            Face f = to.PlanarEntity as Face;
            var v = u.get<Vertex>(f.Vertices, el => u.eq(el.Point, to.OriginPointGeometry));
            var sp = to.AddByProjectingEntity(v) as SketchPoint;
            sp.HoleCenter = false;
            return sp;
        }
        public IEnumerable<SketchPoint> rev()
        {
            return to.SketchPoints.Cast<SketchPoint>().Reverse();
        }
        public void addDimsPt(SketchPoint tmp = null, bool hole = true)
        {
            if (tmp == null)
                tmp = addProj();
            foreach (SketchPoint sp in pts)
            {
                if (!hole && !sp.HoleCenter) continue;
                //bool nov = false, noh = false;
                //if (u.eq(tmp.Geometry.X, sp.Geometry.X)) noh = true;
                //if (u.eq(tmp.Geometry.Y, sp.Geometry.Y)) nov = true;
                var dic1 = constraints.check(tmp);
                var dic2 = constraints.check(sp);

                if (!(dic1["v"] && dic2["v"]))
                {
                    var d = constraints.addTwoPointDist(tmp, sp, "", -1, 1, DimensionOrientationEnum.kHorizontalDim);
                    dims.Add(d);
                    setValue(d);
                }
                if (!(dic1["h"] && dic2["h"]))
                {
                    var d = constraints.addTwoPointDist(tmp, sp, "", -1, 1, DimensionOrientationEnum.kVerticalDim);
                    dims.Add(d);
                    setValue(d);
                }
                tmp = sp;
            }
        }
        public void setValue(TwoPointDistanceDimConstraint d)
        {
            foreach (var item in dims)
            {
                if (u.eq(item.Parameter.ModelValue, d.Parameter.ModelValue))
                {
                    if (item.Equals(d)) continue;
                    d.Parameter.Expression = item.Parameter.Name;
                    return;
                }
            }
            d.Parameter.Value = u.round(d.Parameter.ModelValue);
        }
    }

    public class copyModel
    {
        defData def;
        double r, h;
        Face f;
        List<Edge> edges = new List<Edge>();
        Plane pl;
        public copyModel(Edge e)
        { 
        }
    }

    enum cp_enum { copy, paste};

    internal class CopyPasteBtn : Button
    {
        static public CopyPaste cp;
        public cp_enum en;
        public CopyPasteBtn(string displayName, string internalName, string clientId, string description,
            string tooltip)
            : base(displayName, internalName, CommandTypesEnum.kShapeEditCmdType, clientId, description, tooltip,
                  ButtonDisplayEnum.kDisplayTextInLearningMode)
        { }
        protected override void ButtonDefinition_OnExecute(NameValueMap context)
        {
            switch (en)
            {
                case cp_enum.copy:
                    copy();
                    break;
                case cp_enum.paste:
                    paste();
                    break;
                default:
                    break;
            }
        }
        public void copy()
        {
            Document doc = I.aDoc();
            if (cp == null)
            {
                cp = new CopyPaste(doc);
            }
                cp.add(doc.SelectSet);
        }
        public void paste()
        {
            if (cp == null) return;
            cp.create();
        }
    }
}
