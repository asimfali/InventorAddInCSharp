using InvDoc;
using Inventor;
using System;
using System.Collections.Generic;
using System.Linq;
using InterfaceDll;
using System.Xml.Linq;

namespace InvAddIn
{
    class IMates
    {
        ComponentDefinition def;
        CompositeiMateDefinitionProxy comp;
        SelectSet ss;
        PlanarSketch ps;
        IEnumerable<EdgeProxy> edges;
        FaceProxy fp;
        List<EdgeProxy> lstEdges = new List<EdgeProxy>();
        List<int> nums = new List<int>();
        List<Point> pts = new List<Point>();
        List<HoleFeature> hfs = new List<HoleFeature>();
        Dictionary<Face, Face> holes = new Dictionary<Face, Face>();
        List<IMateData> lst = new List<IMateData>();
        public string oHole, iHole, type;
        public ObjectCollection feats;
        public ObjectCollection bodies;

        public IMates(Document doc)
        {
            //Highlight.set(doc);
            init();
            feats = I.COC();
            if (type == null || type == "")
            {
                System.Windows.Forms.MessageBox.Show("Не выбрана сторона");
                return;
            }
            ss = doc.SelectSet;
            //var trans = I.beginTrans("Конструктивная пара", doc);
            //I.screenSilent(true);
            if (doc.DocumentType == DocumentTypeEnum.kAssemblyDocumentObject)
            {
                def = (ComponentDefinition)((AssemblyDocument)doc).ComponentDefinition;
                if (!(ss[1] is CompositeiMateDefinitionProxy)) return;
                comp = (CompositeiMateDefinitionProxy)ss[1];
                getEdgesProxy(comp, 1);
                getEdgesProxy(comp, 2);
                var ed = edges.ElementAt(0);
                var fp = getFaceProxy(ed);
                createSketch(fp.fp, edges);
            }
            else if (doc.DocumentType == DocumentTypeEnum.kPartDocumentObject)
            {
                def = (ComponentDefinition)((PartDocument)doc).ComponentDefinition;
                add();
            }

            //I.screenSilent(false);
            //trans.End();
        }
        public IMates(Document doc, EdgeProxy e)
        {
            init();
            if (doc.DocumentType == DocumentTypeEnum.kAssemblyDocumentObject)
            {
                def = (ComponentDefinition)((AssemblyDocument)doc).ComponentDefinition;
                getEdgesProxy(e);
                var fp = getFaceProxy(e);
                createSketch(fp.fp, edges);
            }
        }
        public void init()
        {
            oHole = Macros.StandardAddInServer.m_SettingsHole1.get();
            iHole = Macros.StandardAddInServer.m_SettingsHole2.get();
            type = Macros.StandardAddInServer.m_SettingsAct.get();
        }
        public void setBodies()
        {
            bodies = I.COC();
            HashSet<SurfaceBody> tmp = new HashSet<SurfaceBody>();
            foreach (PartFeature item in feats)
            {
                foreach (SurfaceBody sb in item.SurfaceBodies)
                {
                    tmp.Add(sb);
                }
            }
            foreach (var item in tmp)
            {
                bodies.Add(item);
            }
        }
        public void changeDiam()
        {
            if (feats.Count == 2)
            {
                HoleFeature hf1 = feats[1] as HoleFeature, hf2 = feats[2] as HoleFeature;
                Point p1 = I.BoxCenter(hf1.RangeBox), p2 = I.BoxCenter(hf2.RangeBox);
                var origin = ((PartComponentDefinition)def).WorkPoints[1].Point;
                Vector v1 = origin.VectorTo(p1), v2 = origin.VectorTo(p2);
                if (v1.Length < v2.Length)
                {
                    hf1.HoleDiameter.Expression = iHole; hf2.HoleDiameter.Expression = oHole;
                } else
                {
                    hf1.HoleDiameter.Expression = oHole; hf2.HoleDiameter.Expression = iHole;
                }
            }
        }
        public void add()
        {
            if (ss[1] is PlanarSketch)
            {
                ps = (PlanarSketch)ss[1];
                findRay(ps);
            }
            else if (ss[1] is HoleFeature || ss[1] is MirrorFeature)
            {
                Faces fs = null;
                if (ss[1] is HoleFeature) fs = ((HoleFeature)ss[1]).Faces;
                else if (ss[1] is MirrorFeature) fs = ((MirrorFeature)ss[1]).Faces;
                findHoles(fs);

                foreach (var item in holes)
                {
                    IMateData d = dist(item.Key, item.Value);
                    lst.Add(d);
                }
                addIMate();
            }
        }
        public void createSketch(FaceProxy f, IEnumerable<EdgeProxy> ie)
        {
            if (f == null) return;
            ComponentOccurrence to = f.ContainingOccurrence;
            if (to.DefinitionDocumentType != DocumentTypeEnum.kPartDocumentObject) return;
            PartComponentDefinition d = (PartComponentDefinition)to.Definition;
            //u.higlight((Document)def.Document, f, f, null, null); 
            //return;
            PlanarSketch ps = d.Sketches.Add(f.NativeObject);

            int i = -1;
            foreach (var e in ie)
            {
                i++;
                Point pt = null;
                Circle c = e.Geometry as Circle;
                if (c != null)
                {
                    pt = c.Center;
                }
                else if (e.TangentiallyConnectedEdges.Count == 4)
                {
                    pt = u.slotCenter(e.EdgeUses[1].EdgeLoop);
                }

                var mtx = u.transformAsmToPart(to);
                pt.TransformBy(mtx);
                if (checkInPts(pt))
                    continue;
                pts.Add(pt);
                ps.SketchPoints.Add(ps.ModelToSketchSpace(pt));
                if (lstEdges.Contains(e)) nums.Add(i);
            }
            if (oHole == null || oHole == "") oHole = (((Circle)ie.ElementAt(0).Geometry).Radius * 20).ToString();
            List<Point> lst = u.gets<SketchPoint>(ps.SketchPoints, fi => fi.HoleCenter == true).Select(el => el.Geometry3d).ToList();
            var fd = u.findNearestFace(d, f, lst[0]);
            if (fd.f == null) return;
            double dist = fd.dist();
            var dir = u.holeDir(fd.pl, ps);
            //double dist = pt1.Key == null ? 20 : lst[0].DistanceTo(pt1.Key);
            //if (dist == 0) dist = 3;
            //PartFeatureExtentDirectionEnum dir = pt1.Key == null ?
            //    PartFeatureExtentDirectionEnum.kNegativeExtentDirection :
            //    u.holeDir(ps.PlanarEntityGeometry, lst[0], pt1.Key);
            var hf = addHole(d, ps, lst, oHole, dist, dir, fd.num(d));
            if (hf == null) return;
            addIMate(d, hf);
            string fDist = Macros.StandardAddInServer.m_FastenerDist.get(),
                fName = Macros.StandardAddInServer.m_FastenerName.get();
            if (fName == null || fName == "") return;
            //Plane pl = getPlane((EdgeProxy)pt1.Value);
            addFastIMate(hf, d, fDist, fName);
        }
        public bool checkInPts(Point pt)
        {
            foreach (Point item in pts)
            {
                if (u.eq(item, pt)) return true;
            }
            return false;
        }
        public void addFastIMate(HoleFeature hf1, PartComponentDefinition def, string fDist, string fName,
            string dir = "Снаружи")
        {
            if (hf1.Faces.Count == 0) return;
            var f1 = hf1.Faces[1];
            var pt = I.CP();
            var e1 = f1.Edges[1]; var e2 = f1.Edges[2];
            double d1 = u.distToEdge(e1, pt), d2 = u.distToEdge(e2, pt);
            int i = (d1 > d2) ? 1 : 2;
            foreach (Face f in hf1.Faces)
            {
                double d = 0;
                if (fDist != null) d = u.convToDouble(fDist);

                var imate = IMate.iMate_(f.Edges[i], def, d);
                imate.Name = fName;
            }
        }
        public void test()
        {
            AssemblyComponentDefinition adef = (AssemblyComponentDefinition)def;
            ComponentOccurrence occ1 = adef.Occurrences[1], occ2 = adef.Occurrences[2];
            PartComponentDefinition pd1 = (PartComponentDefinition)occ1.Definition,
                pd2 = (PartComponentDefinition)occ2.Definition;
            Edge e = pd1.SurfaceBodies[1].Edges[1];
            object oe;
            //occ1.CreateGeometryProxy(e, out oe);
            //EdgeProxy ep = (EdgeProxy)oe;
            PlanarSketch ps = pd2.Sketches[1];
            object os;
            Circle c = (Circle)e.Geometry;
            Point pt = c.Center;
            //occ2.CreateGeometryProxy(ps, out os);
            //PlanarSketchProxy psp = (PlanarSketchProxy)os;
            var mtx = occ1.Transformation;
            var mtx2 = occ2.Transformation;
            mtx.TransformBy(mtx2);
            //double x = mtx.Cell[1, 4], y = mtx.Cell[2, 4], z = mtx.Cell[3, 4];
            //x = u.reverse(x); y = u.reverse(y); z = u.reverse(z);
            //mtx.Cell[1, 1] = 1;
            //mtx.Cell[1, 4] = x; mtx.Cell[2, 4] = y; mtx.Cell[3, 4] = z;

            //mtx.Cell[1, 1] = 1; mtx.Cell[2, 2] = 1; mtx.Cell[3, 3] = 1;
            pt.TransformBy(mtx);
            ps.SketchPoints.Add(ps.ModelToSketchSpace(pt));
            //SketchEntity prE = psp.AddByProjectingEntity(ep);
            //ps.BreakLink();
        }
        public void getEdgesProxy(CompositeiMateDefinitionProxy c, int n)
        {
            EdgeProxy ep;
            if (c.Count != 2) return;
            if (c[n] is InsertiMateDefinitionProxy)
            {
                InsertiMateDefinitionProxy i = (InsertiMateDefinitionProxy)c[n];
                ep = (EdgeProxy)i.Entity;
                lstEdges.Add(ep);
                FaceProxy f = u.get<FaceProxy>(ep.Faces, fi => fi.SurfaceType == SurfaceTypeEnum.kPlaneSurface);
                if (fp == null) fp = f;
                else if (fp.Equals(f)) return;
                getEdgesProxy(f.Edges, fi => u.eq(fi.Radius, ((Circle)ep.Geometry).Radius));
            }
        }
        public void getEdgesProxy(EdgeProxy e)
        {
            dynamic c = e.Geometry;
            if (c == null) return;
            var co = e.ContainingOccurrence;
            var poc = co.ParentOccurrence;
            var pl = getPlane(e);
            if (poc != null && poc.DefinitionDocumentType == DocumentTypeEnum.kAssemblyDocumentObject)
            {
                foreach (ComponentOccurrenceProxy so in poc.SubOccurrences)
                {
                    if (so.SurfaceBodies.Count == 0) continue;
                    var es = so.SurfaceBodies[1].Edges;
                    getEdgesProxy(es, fi => u.eq(fi.Radius, c.Radius) && check(pl, fi.Center));
                }
            }
            else if (co.DefinitionDocumentType == DocumentTypeEnum.kPartDocumentObject)
            {
                var es = co.SurfaceBodies[1].Edges;
                getEdgesProxy(es, fi => u.eq(fi.Radius, c.Radius) && check(pl, fi.Center));
            }
        }
        public Plane getPlane(EdgeProxy e)
        {
            FaceProxy face = u.get<FaceProxy>(e.Faces, fi => fi.SurfaceType == SurfaceTypeEnum.kPlaneSurface);
            Plane pl = (Plane)face.Geometry;
            return pl;
        }
        public bool check(Plane pl, Point pt)
        {
            double d = pl.DistanceTo(pt);
            return u.eq(d, 0) ? true : false;
        }
        public void getEdgesProxy(Edges e, Func<dynamic, bool> f)
        {
            var es = u.gets<EdgeProxy>(e, fi => fi.CurveType == CurveTypeEnum.kCircleCurve && f(fi.Geometry));
            if (edges == null) edges = es;
            else edges = edges.Concat(es);
        }
        public void findRay(PlanarSketch ps)
        {
            if (ps.SketchPoints.Count == 0) return;
            if (type == "Видимые") FaceData.visible = true;
            Plane p = ps.PlanarEntityGeometry;
            var v = p.Normal; v = u.round(v);
            UnitVector r = u.reverse(v);
            addHoles(ps, v, iHole, oHole);
            addHoles(ps, r, iHole, oHole);
            changeDiam();
        }
        public bool checkDep(SketchPoint sp)
        {
            PlanarSketch ps = sp.Parent as PlanarSketch;
            foreach (var d in ps.Dependents)
            {
                var e = d as HoleFeature;
                if (e == null) continue;
                foreach (SketchPoint item in e.HoleCenterPoints)
                {
                    if (item.Equals(sp)) return true;
                }
            }
            return false;
        }
        public void addHoles(PlanarSketch ps, UnitVector v, string iHole, string oHole)
        {
            List<FaceData> fDates = new List<FaceData>();
            PartFeatureExtentDirectionEnum dir =
                u.dotProduct(v.AsVector(), ps.PlanarEntityGeometry.Normal.AsVector()) > 0 ?
                PartFeatureExtentDirectionEnum.kNegativeExtentDirection :
                PartFeatureExtentDirectionEnum.kPositiveExtentDirection;
            foreach (SketchPoint item in u.gets<SketchPoint>(ps.SketchPoints, fi => fi.HoleCenter))
            {
                //if (checkDep(item)) continue;
                var pt = item.Geometry3d;
                //var vec = v.Copy().AsVector(); vec.ScaleBy(-1);
                //pt.TranslateBy(vec);
                List<FaceData> tmp = new List<FaceData>();
                u.findUsingRay((PartComponentDefinition)def, pt, v, (o, l) =>
                {
                    FaceData fd = new FaceData(o, (Point)l);
                    if (fd.f != null)
                        tmp.Add(fd);
                }, 0.1, true);
                if (tmp.Count < 2) continue;
                else if (tmp.Count >= 2)
                {
                    int i = 1;
                    if (!tmp[i - 1].check(tmp[i])) i = 2;
                    tmp[i].setSP(pt);
                    tmp[i].setSide(); tmp[i].setDiam(iHole, oHole);
                    fDates.Add(tmp[i]);
                }
            }
            var ig = fDates.GroupBy(e => e.sb);
            string fast = Macros.StandardAddInServer.m_FastenerName.get(),
            fDist = Macros.StandardAddInServer.m_FastenerDist.get();
            foreach (var item in ig)
            {
                List<Point> pts = new List<Point>();
                FaceData elem = item.ElementAt(0);
                foreach (var vals in item)
                {
                    pts.Add(vals.sp);
                }
                HoleFeature hf1 = addHole((PartComponentDefinition)def, ps, pts, elem.diam,
                    elem.dist() + 0.3, dir, elem.num((PartComponentDefinition)def)); 
                if (hf1 == null) return;
                hf1.SetAffectedBodies(u.findBodies((PartComponentDefinition)def, hf1 as PartFeature));
                feats.Add(hf1 as PartFeature);
                if (elem.s == side.outher && fast != null)
                {
                    foreach (Face f in hf1.Faces)
                    {
                        double d = 0;
                        if (fDist != null) d = u.convToDouble(fDist);
                        var imate = IMate.iMate_(fEdge(f), (PartComponentDefinition)def, d);
                        imate.Name = fast;
                    }
                }
            }
        }
        //public void addHoles(SketchPoints sps, UnitVector v, string iHole, string oHole,
        //    PartFeatureExtentDirectionEnum dir)
        //{
        //    List<FaceData> fDates = new List<FaceData>();
        //    foreach (SketchPoint item in u.gets<SketchPoint>(sps, fi => fi.HoleCenter == true))
        //    {
        //        var pt = item.Geometry3d;
        //        FaceData face = getFace(pt, v, fi => fi.checkPlane());
        //        if (face.f == null) continue;
        //        face.setSide(); face.setDiam(iHole, oHole);
        //        fDates.Add(face);
        //        //Highlight.add(face.f);
        //    }
        //    var ig = fDates.GroupBy(e => e.sb);
        //    string fast = Macros.StandardAddInServer.m_FastenerName.get(),
        //        fDist = Macros.StandardAddInServer.m_FastenerDist.get();
        //    List<HoleFeature> delHF = new List<HoleFeature>();
        //    foreach (var item in ig)
        //    {
        //        List<Point> pts = new List<Point>();
        //        FaceData elem = item.ElementAt(0);
        //        foreach (var vals in item)
        //        {
        //            pts.Add(vals.sp);
        //        }
        //        HoleFeature hf1 = addHole((PartComponentDefinition) def,ps, pts, elem.diam, elem.d + 0.3, dir, elem.sb);
        //        if (hf1.HealthStatus == HealthStatusEnum.kDriverLostHealth)
        //        {
        //            //delHF.Add(hf1); continue;
        //        }
        //        if (elem.s == side.outher && fast != null)
        //        {
        //            foreach (Face f in hf1.Faces)
        //            {
        //                double d = 0;
        //                if (fDist != null) d = u.convToDouble(fDist);
        //                var imate = IMate.iMate_(fEdge(f), (PartComponentDefinition)def, d);
        //                imate.Name = fast;
        //            }
        //        }
        //    }
        //    foreach (HoleFeature item in delHF)
        //    {
        //        item.Delete(true, true, true);
        //    }
        //}
        public Edge fEdge(Face f)
        {
            var pt = I.CP();
            Circle c1 = (Circle)f.Edges[1].Geometry, c2 = (Circle)f.Edges[2].Geometry;
            Vector v1 = pt.VectorTo(c1.Center), v2 = pt.VectorTo(c2.Center);
            return (v1.Length > v2.Length) ? f.Edges[1] : f.Edges[2];
        }
        public FaceData getFace(Point pt, UnitVector v, Func<FaceData, bool> f)
        {
            ObjectsEnumerator objs = null, loc = null;
            Point pt1 = pt; // I.CP(pt, v.AsVector(), 0.1);
            if (def is PartComponentDefinition)
                ((PartComponentDefinition)def).FindUsingRay(pt1, v, 0.1, out objs, out loc);
            else if (def is AssemblyComponentDefinition)
                ((AssemblyComponentDefinition)def).FindUsingRay(pt1, v, 0.1, out objs, out loc);
            int i = 0;
            //if (loc.Count == 1) return new FaceData(null, pt);
            foreach (Point item in loc)
            {
                i++;
                Point p = item;
                if (!(objs[i] is FaceProxy || objs[i] is Face)) continue;
                //if (objs[i] is FaceProxy && !check((FaceProxy)objs[i])) continue;

                FaceData fd = new FaceData(objs[i], p);
                fd.setSP(pt);
                if (f(fd)) return fd;
            }
            return new FaceData(null, pt);
        }
        public bool check(FaceProxy fp)
        {
            var occ = fp.ContainingOccurrence;
            string fn1 = ((Document)def.Document).FullDocumentName,
                fn2 = occ.ReferencedFileDescriptor.FullFileName;
            string p1 = file.p(fn1), p2 = file.p(fn2);
            return p1 == p2;
        }
        public void filter()
        {
        }
        public FaceData getFaceProxy(EdgeProxy e)
        {
            ObjectsEnumerator objs, loc;
            dynamic c = e.Geometry;
            Point pt = c.Center;
            UnitVector v = c.Normal, r = u.reverse(v);
            //Point ptv = I.CP(pt, v.AsVector(), 0.1), ptr = I.CP(pt, r.AsVector(), 0.1);
            FaceData
                //f1 = getFace(pt, v, fi => !fi.fp.Parent.Equals(e.Parent)),
                f2 = getFace(pt, r, fi => !fi.fp.Parent.Equals(e.Parent));
            return f2;
            //if (f1.fp != null && f2.fp == null) return f1;
            //else if (f2.fp != null && f1.fp == null) return f2;
            //else if (f1.fp != null && f2.fp != null)
            //{
            //    double d1 = u.round(f1.loc.DistanceTo(pt)), d2 = u.round(f2.loc.DistanceTo(pt));
            //    if (d1 <= d2) return f1;
            //    else return f2;
            //}
            //return f1;
        }
        //public KeyValuePair<FaceProxy, Point> getFaceProxy(Point pt, UnitVector v, Func<FaceProxy, bool> f)
        //{
        //    ObjectsEnumerator objs, loc;
        //    AssemblyComponentDefinition adef = (AssemblyComponentDefinition)def;
        //    adef.FindUsingRay(pt, v, 0.1, out objs, out loc);
        //    int i = 1;
        //    foreach (Point item in loc)
        //    {
        //        Point p = item;
        //        if (objs[1] is FaceProxy proxy && f(proxy)) return new KeyValuePair<FaceProxy, Point>(proxy, p);
        //        i++;
        //    }
        //    return new KeyValuePair<FaceProxy, Point>();
        //}
        public HoleFeature addHole(PartComponentDefinition def, PlanarSketch sk, List<Point> pts,
            string diam, object d, PartFeatureExtentDirectionEnum dir, int num)
        {
            ObjectCollection col = I.getSketchPoints(sk, fi => fi.HoleCenter == true && contains(pts, fi.Geometry3d));
            if (col.Count == 0) return null;
            if (diam.Contains("x"))
            {
                var spl = diam.Split('x');
                Dictionary<string, double> dict = new Dictionary<string, double>()
                {
                    {"ширина", u.convToDouble(spl[0])}, {"a", u.convToDouble(spl[1])}
                };
                I.createPunch("Punches/Овал.ide", dict, col, def);
                return null;
            }
            SketchHolePlacementDefinition hdef = def.Features.HoleFeatures.CreateSketchPlacementDefinition(col);
            try
            {
                HoleFeature hf = def.Features.HoleFeatures.AddDrilledByDistanceExtent(hdef, diam, d, dir);
                col = I.COC();
                if (num == 0) return hf;
                col.Add(def.SurfaceBodies[num]);
                if (def.SurfaceBodies.Count > 1)
                {
                    hf.SetAffectedBodies(col);
                }
                if (hf.HealthStatus == HealthStatusEnum.kDriverLostHealth)
                {
                    PartFeatureExtentDirectionEnum ext = dir == PartFeatureExtentDirectionEnum.kNegativeExtentDirection ?
                        PartFeatureExtentDirectionEnum.kPositiveExtentDirection :
                        PartFeatureExtentDirectionEnum.kNegativeExtentDirection;
                    hf.SetDistanceExtent(d, ext);
                }
                return hf;
            }
            catch (Exception)
            {
                return null;
            }
            //if (hf.HealthStatus == HealthStatusEnum.kDriverLostHealth)
            //{
            //    PartFeatureExtentDirectionEnum ext = dir == PartFeatureExtentDirectionEnum.kNegativeExtentDirection ?
            //        PartFeatureExtentDirectionEnum.kPositiveExtentDirection : 
            //        PartFeatureExtentDirectionEnum.kNegativeExtentDirection;
            //    hf.SetDistanceExtent(d, ext);
            //}
        }
        //public void fill(iFeatureDefinition def, List<string> p)
        //{
        //    int i = 0;
        //    if (def.iFeatureInputs.Count == p.Count)
        //    {
        //        foreach (iFeatureParameterInput item in def.iFeatureInputs)
        //        {
        //            item.Expression = p[i];
        //            i++;
        //        }
        //    }
        //}
        //public void addPunch(List<string> p, string fn, List<Point> pts)
        //{
        //    var smf = I.getSMF();
        //    var def = smf.PunchToolFeatures.CreateiFeatureDefinition(fn);
        //    fill(def, p);
        //    var col = I.COC();
        //    foreach (var item in pts)
        //    {
        //        col.Add(item);
        //    }
        //    var punch = smf.PunchToolFeatures.Add(col, def, AcrossBends: true);
        //}
        public bool contains(List<Point> pts, Point pt)
        {
            foreach (Point item in pts)
            {
                if (u.eq(pt, item)) return true;
            }
            return false;
        }
        public void findHF(PlanarSketch ps)
        {
            PartComponentDefinition def = (PartComponentDefinition)this.def;
            foreach (HoleFeature item in def.Features.HoleFeatures)
            {
                if (item.Sketch.Equals(ps)) hfs.Add(item);
            }
        }
        public void findHoles(Faces fs)
        {
            foreach (Face f1 in fs)
            {
                Face f2 = findFace(f1);
                if (f2 == null) continue;
                holes.Add(f1, f2);
            }
        }
        public Face findFace(Face f2)
        {
            PartComponentDefinition def = (PartComponentDefinition)this.def;
            if (f2.SurfaceType != SurfaceTypeEnum.kCylinderSurface) return null;
            Cylinder c1 = (Cylinder)f2.Geometry;
            var f = u.findAtPoint<Face>(def, c1.BasePoint,
                new SelectionFilterEnum[] { SelectionFilterEnum.kPartFaceCylindricalFilter },
                15, fi => check(c1, (Cylinder)fi.Geometry) && !fi.SurfaceBody.Equals(f2.SurfaceBody));
            return f;
        }
        public bool check(Cylinder c1, Cylinder c2)
        {
            UnitVector u1 = c1.AxisVector, u2 = c2.AxisVector;
            Point pt1 = c1.BasePoint, pt2 = c2.BasePoint;
            return u1.IsParallelTo(u2) && check(pt1, pt2);
        }
        public bool check(Point pt1, Point pt2)
        {
            int count = 0;
            double[] c1 = { }, c2 = { };
            pt1.GetPointData(ref c1); pt2.GetPointData(ref c2);
            for (int i = 0; i < c1.Count(); i++)
            {
                if (u.eq(c1[i], c2[i])) count++;
            }
            if (count > 1) return true;
            return false;
        }
        public IMateData dist(Face f1, Face f2)
        {
            List<IMateData> d = new List<IMateData>() { new IMateData(f1.Edges[1], f2.Edges[1]), new IMateData(f1.Edges[1], f2.Edges[2]),
            new IMateData(f1.Edges[2], f2.Edges[1]), new IMateData(f1.Edges[2], f2.Edges[2])};
            d.Sort();
            return d.First();
        }
        public void addIMate(PartComponentDefinition def, HoleFeature hf)
        {
            if (nums.Count != 2) return;
            List<DirHole> hs = null;
            List<Edge> edges = new List<Edge>();

            for (int i = 0; i < nums.Count; i++)
            {
                hs = new List<DirHole>();
                hs.Add(new DirHole(hf, 1, pts[nums[i]]));
                hs.Add(new DirHole(hf, 2, pts[nums[i]]));
                hs.Sort();
                edges.Add(hs[0].e);
            }
            InsertiMateDefinition iimd = (InsertiMateDefinition)comp[1];
            string dist = iimd.Distance.Expression;
            string match = comp.Name;
            if (comp.MatchList != null)
            {
                string[] m = comp.MatchList as string[];
                if (m.Length > 0) match = m[0];
            }
            CompositeiMateDefinition im = IMate.iInsComposite(edges[0], edges[1], def, dist, match, comp.Name);
        }
        public void addIMate()
        {
            PartComponentDefinition def = (PartComponentDefinition)this.def;
            var n = getName();
            PartDocument bdoc = (PartDocument)def.Document;
            string type = u.getPropValue((Document)bdoc, "Type");

            CompositeiMateDefinition im1, im2;
            im1 = IMate.iInsComposite(lst[0].e1, findEdge(lst[0].e1, lst[1]), def, u.round(lst[0].d * 100), n[0], n[1]);
            im2 = IMate.iInsComposite(lst[0].e2, findEdge(lst[0].e2, lst[1]), def, u.round(lst[0].d * 100), n[1], n[0]);
            if (type != "")
            {
                I.open(file.p(bdoc.FullDocumentName), type + "*.ipt");
                var pd = I.findInRef(bdoc, refName(lst[0].e1, bdoc));
                if (pd != null)
                {
                    var imList = findInsIMId(lst[0].e1.Parent); imList.Add(im1.Identifier);
                    I.addSB((Document)pd, bdoc.FullDocumentName, null, imList);
                }
                pd = I.findInRef(bdoc, refName(lst[0].e2, bdoc));
                if (pd != null)
                {
                    var imList = findInsIMId(lst[0].e1.Parent); imList.Add(im2.Identifier);
                    I.addSB((Document)pd, bdoc.FullDocumentName, null, imList);
                }
            }
        }
        public List<string> findInsIMId(SurfaceBody sb)
        {
            PartComponentDefinition def = (PartComponentDefinition)this.def;
            List<string> lst = new List<string>();
            foreach (iMateDefinition item in def.iMateDefinitions)
            {
                if (item is InsertiMateDefinition)
                {
                    SurfaceBody s = ((Edge)((InsertiMateDefinition)item).Entity).Parent;
                    if (s.Equals(sb)) lst.Add(item.Identifier);
                }
            }
            return lst;
        }
        public string refName(Edge e, PartDocument doc)
        {
            return e.Parent.Name + "::" + file.nameWithExt(doc.FullDocumentName);
        }
        public List<string> getName()
        {
            string n1 = lst[0].e1.Parent.Name.Split('$')[0], n2 = lst[0].e2.Parent.Name.Split('$')[0];
            return new List<string>() { n1 + "_" + n2, n2 + "_" + n1 };
        }
        public Edge findEdge(Edge e, IMateData d)
        {
            var sb = e.Parent;
            if (d.e1.Parent.Equals(sb)) return d.e1;
            else return d.e2;
        }

    }
    public class IMateData : IComparable
    {
        private Circle c1, c2;
        public Edge e1, e2;
        public double d;
        public IMateData(Edge e1, Edge e2)
        {
            this.e1 = e1; this.e2 = e2;
            c1 = (Circle)e1.Geometry; c2 = (Circle)e2.Geometry;
            d = dist(c1, c2);
        }
        public double dist(Circle c1, Circle c2)
        {
            return c1.Center.DistanceTo(c2.Center);
        }
        public int CompareTo(object obj)
        {
            if (obj == null) return 1;
            IMateData d1 = obj as IMateData;
            return this.d.CompareTo(d1.d);
        }
    }
    public enum side { inner, outher }
    public class FaceData
    {
        public static bool visible = false;
        public side s;
        public Face f = null;
        public FaceProxy fp = null;
        readonly public SurfaceTypeEnum t;
        public readonly Point loc;
        public string diam;
        public SurfaceBody sb;
        public string sbName;
        public Plane pl;
        public Point sp;
        Vector n;
        public double d;
        public FaceData(object fa, Point pt)
        {
            if (fa is Face)
            {
                f = (Face)fa;
                t = f.SurfaceType;
                loc = pt;
            }
            if (fa is FaceProxy)
            {
                fp = (FaceProxy)fa; f = fp.NativeObject;
                t = f.SurfaceType;
                loc = pt;
            }
            if (f != null)
            {
                var sb = f.Parent;
                if (visible && !sb.Visible)
                {
                    f = null;
                    return;
                }
                sbName = sb.Name;
            }
        }
        private void setNull()
        {
            f = null;
        }
        public void setDiam(string iHole, string oHole)
        {
            diam = (s == side.inner) ? iHole : oHole;
        }
        public void setSP(Point pt)
        {
            sp = pt;
        }
        public bool checkPlane()
        {
            if (t != SurfaceTypeEnum.kPlaneSurface) setNull();
            Vector v = sp.VectorTo(loc);
            pl = (Plane)f.Geometry;
            d = u.round(v.Length);
            return f != null;
        }
        public void setSide()
        {
            Point bp = I.CP();
            Vector v = bp.VectorTo(sp), v1 = bp.VectorTo(loc);
            sb = f.Parent;
            sbName = sb.Name;
            double d = v.Length, d1 = v1.Length;
            if (d1 < d) s = side.inner;
            else s = side.outher;
        }
        public bool check(FaceData o)
        {
            sb = f.Parent;
            sbName = sb.Name;
            o.sb = o.f.Parent;
            return o.sb.Equals(sb);
        }
        public double dist()
        {
            pl = (Plane)f.Geometry;
            double d = u.distToPlane(pl, sp);
            if (u.eq(d, 0)) d = 3;
            return d;
        }
        public int num(PartComponentDefinition def)
        {
            for (int i = 1; i < def.SurfaceBodies.Count + 1; i++)
            {
                if (def.SurfaceBodies[i].Name == sbName) return i;
            }
            return 0;
        }
    }
    public class DirHole : IComparable
    {
        public HoleFeature hf;
        public Edge e;
        public Point bp;
        public double d;
        Point pt;
        public DirHole(HoleFeature hf, int edNum, Point bp)
        {
            this.hf = hf;
            this.bp = bp;
            e = find().Edges[edNum];
            pt = u.centerEdge(e);
            d = pt.DistanceTo(bp);
        }
        public int CompareTo(object obj)
        {
            if (obj == null) return 1;
            DirHole other = obj as DirHole;
            return this.d.CompareTo(other.d);
        }
        Face find()
        {
            var uv = hf.Sketch.PlanarEntityGeometry.Normal;
            foreach (Face f in hf.Faces)
            {
                Point ce = u.centerEdge(f.Edges[1]);
                if (u.eq(ce, bp)) return f;
                UnitVector v = bp.VectorTo(ce).AsUnitVector();
                if (v.IsParallelTo(uv)) return f;
            }
            return null;
        }
    }
    public class MoveIM
    {
        public AssemblyDocument doc;
        SelectSet ss;
        public MoveIM(Document d)
        {
            if (d.DocumentType != DocumentTypeEnum.kAssemblyDocumentObject) return;
            doc = d as AssemblyDocument;
            //ss = doc.SelectSet;
            //move();
        }
        public void move(System.Collections.IEnumerable ss)
        {
            foreach (var item in ss)
            {
                var im = item as InsertiMateDefinitionProxy;
                if (im != null)
                {
                    if (check(im)) continue;
                    im.NativeObject.Suppressed = true;
                    im.Suppressed = true;
                    addIMate(im);
                }
                var cm = item as CompositeiMateDefinitionProxy;
                if (cm != null)
                {
                    if (check(cm)) continue;
                    cm.Suppressed = true;
                    cm.NativeObject.Suppressed = true;
                    var col = I.COC();
                    foreach (var m in cm)
                    {
                        var e = addIMate(m);
                        if (e != null) col.Add(e);
                    }
                    doc.ComponentDefinition.iMateDefinitions.AddCompositeiMateDefinition(col,
                        get_name(cm.Name), cm.MatchList);
                }
            }
            doc.Update2();
        }
        public iMateDefinition addIMate(object m)
        {
            object pr;
            var im = m as InsertiMateDefinitionProxy;
            if (im != null)
            {
                im.ContainingOccurrence.CreateGeometryProxy(im.Entity, out pr);
                return doc.ComponentDefinition.iMateDefinitions.AddInsertiMateDefinition(pr, im.AxesOpposed, im.Distance.Value,
                    Name: get_name(im.Name), MatchList: im.MatchList) as iMateDefinition;
            }
            var am = m as AngleiMateDefinitionProxy;
            if (am != null)
            {
                am.ContainingOccurrence.CreateGeometryProxy(am.Entity, out pr);
                return doc.ComponentDefinition.iMateDefinitions.AddAngleiMateDefinition(pr, am.DirectionOneReversed, am.Angle.Value,
                    Name: get_name(am.Name), MatchList: am.MatchList) as iMateDefinition;
            }
            var fm = m as FlushiMateDefinitionProxy;
            if (fm != null)
            {
                fm.ContainingOccurrence.CreateGeometryProxy(fm.Entity, out pr);
                return doc.ComponentDefinition.iMateDefinitions.AddFlushiMateDefinition(pr, fm.Offset.Value,
                    Name: get_name(fm.Name), MatchList: fm.MatchList) as iMateDefinition;
            }
            var mm = m as MateiMateDefinitionProxy;
            if (mm != null)
            {
                mm.ContainingOccurrence.CreateGeometryProxy(mm.Entity, out pr);
                return doc.ComponentDefinition.iMateDefinitions.AddMateiMateDefinition(pr, mm.Offset.Value,
                    Name: get_name(mm.Name), MatchList: mm.MatchList) as iMateDefinition;
            }
            return null;
        }
        public string get_name(string s)
        {
            if (s.IndexOf('^') == -1) return s;
            var spl = s.Split('^');
            return spl[0];
        }
        public bool check(InsertiMateDefinitionProxy i)
        {
            foreach (var item in doc.ComponentDefinition.iMateDefinitions)
            {
                var im = item as InsertiMateDefinition;
                if (im != null)
                {
                    if (eq(i.Entity, im.Entity)) return true;
                }
            }
            return false;
        }
        public bool check(CompositeiMateDefinitionProxy i)
        {
            foreach (var item in doc.ComponentDefinition.iMateDefinitions)
            {
                var cm = item as CompositeiMateDefinition;
                if (cm != null)
                {
                    foreach (var m in cm)
                    {
                        var e = getEnt(m);
                        if (check(e, i)) return true;
                    }
                }
            }
            return false;
        }
        public bool check(object e, CompositeiMateDefinitionProxy i)
        {
            int num = 0;
            foreach (var item in i)
            {
                var ent = getEnt(item);
                if (eq(ent, e)) num++;
            }
            return num > 0 ? true : false;
        }
        public object getEnt(object m)
        {
            var im = m as InsertiMateDefinition;
            if (im != null) return im.Entity;
            var am = m as AngleiMateDefinition;
            if (am != null) return am.Entity;
            var fm = m as FlushiMateDefinition;
            if (fm != null) return fm.Entity;
            var mm = m as MateiMateDefinition;
            if (mm != null) return mm.Entity;
            return null;
        }
        public bool eq(object o1, object o2)
        {
            EdgeProxy e1 = o1 as EdgeProxy, e2 = o2 as EdgeProxy;
            if (e1 == null || e2 == null) return false;
            return e1.NativeObject.Equals(e2.NativeObject);
            //Circle c1 = e1.Geometry as Circle, c2 = e2.Geometry as Circle;
            //if (c1 == null || c2 == null) return false;
            //return u.eq(c1.Center, c2.Center);
        }
    }
    public class MatePlane
    {
        AssemblyDocument doc;
        CompositeiMateDefinitionProxy cm;
        CommandManager mgr;
        public MatePlane(Document d)
        {
            if (d.DocumentType != DocumentTypeEnum.kAssemblyDocumentObject) return;
            doc = d as AssemblyDocument;
            var colors = new List<byte[]>()
            {
                new byte[] {255,0,0}, new byte[] {255,0,0},
                new byte[] {0,255,0}, new byte[] {0,255,0},
                new byte[] {0,0,255}, new byte[] {0,0,255}
            };
            Highlight.set(d, colors);
            cm = doc.SelectSet[1] as CompositeiMateDefinitionProxy;
            if (cm == null)
            {
                System.Windows.Forms.MessageBox.Show("Должна быть выбрана конструктивная группа");
                return;
            }
            mgr = I.app.CommandManager;
            highlight();
        }
        public void highlight()
        {
            Dictionary<object, object> dic = new Dictionary<object, object>();
            //ComponentOccurrence occ = cm.ContainingOccurrence;
            foreach (var item in cm)
            {
                MateiMateDefinitionProxy md = item as MateiMateDefinitionProxy;
                FlushiMateDefinitionProxy fd = item as FlushiMateDefinitionProxy;
                if (md != null)
                {
                    Highlight.add(md.Entity);
                }
                else if (fd != null) Highlight.add(fd.Entity);
                var pr = mgr.Pick(SelectionFilterEnum.kAllPlanarEntities, "Выберите плоскость");
                Highlight.add(pr);
                if (md != null) dic.Add(md, pr);
                else if (fd != null) dic.Add(fd, pr);
            }
            Highlight.clear();
            foreach (var d in dic)
            {
                var ent = getEnt(d.Key);
                var offset = getOffset(d.Key);
                Plane pl1 = getPlane(ent), pl2 = getPlane(d.Value);
                if (pl1 == null || pl2 == null) return;
                if (u.dotProduct(pl1.Normal.AsVector(), pl2.Normal.AsVector()) < 0)
                    doc.ComponentDefinition.Constraints.AddMateConstraint(ent, d.Value, offset);
                else doc.ComponentDefinition.Constraints.AddFlushConstraint(ent, d.Value, offset);
            }
        }
        public Plane getPlane(object pr)
        {
            if (pr == null) return null;
            var wp = pr as WorkPlane;
            var f = pr as Face;
            if (wp == null && f == null) return null;
            Plane pl = wp != null ? wp.Plane : f.Geometry as Plane;
            return pl;
        }
        public object getEnt(object item)
        {
            MateiMateDefinitionProxy md = item as MateiMateDefinitionProxy;
            FlushiMateDefinitionProxy fd = item as FlushiMateDefinitionProxy;
            if (md != null) return md.Entity;
            else if (fd != null) return fd.Entity;
            return null;
        }
        public object getOffset(object item)
        {
            MateiMateDefinitionProxy md = item as MateiMateDefinitionProxy;
            FlushiMateDefinitionProxy fd = item as FlushiMateDefinitionProxy;
            if (md != null) return md.Offset.Value;
            else if (fd != null) return fd.Offset.Value;
            return null;
        }
    }
    public class CopyIMate
    {
        Document doc;
        CommandManager mgr;
        SelectSet ss;
        SelectEvents sel;
        bool flag = false, rev;
        public CopyIMate(Document d, composeIM im, bool rev = true)
        {
            doc = d;
            this.rev = rev;
            mgr = I.app.CommandManager;
            add(im);
        }
        public object getDef()
        {
            if (doc.DocumentType == DocumentTypeEnum.kAssemblyDocumentObject)
                return ((AssemblyDocument)doc).ComponentDefinition;
            else if (doc.DocumentType == DocumentTypeEnum.kPartDocumentObject)
                return ((PartDocument)doc).ComponentDefinition;
            return null;
        }
        public void add(composeIM comp)
        {
            dynamic def = getDef();
            Document d = def.Document;
            if (!d.IsModifiable)
            {
                System.Windows.Forms.MessageBox.Show("Документ нельзя модифицировать");
                return;
            }
            var col = I.COC();
            int i = 0;
            foreach (var item in comp.ims)
            {
                if (item.name == null) continue;
                i++;
                object pr = null;

                if (comp.ims.Count != 1)
                    pr = item.t == imType.Insert ? mgr.Pick(SelectionFilterEnum.kAllCircularEntities, "Выберите отверстие " + i) :
                        mgr.Pick(SelectionFilterEnum.kAllPlanarEntities, "Выберите плоскость ");
                else
                {
                    ss = doc.SelectSet;
                    ss.Clear();
                    var lst = selOp();
                    foreach (var el in lst)
                    {
                        if (check(el, def.iMateDefinitions)) continue;
                        var m = def.iMateDefinitions.AddInsertiMateDefinition(el, item.ao, item.dist);
                        m.Name = comp.name;
                    }
                    return;
                }
                if (pr != null && item.t == imType.Insert)
                {
                    var m = def.iMateDefinitions.AddInsertiMateDefinition(pr, item.ao, item.dist);
                    col.Add(m);
                }
                else if (pr != null && item.t == imType.Mate)
                {
                    iMateDefinitions ims = def.iMateDefinitions;
                    var m = ims.AddMateiMateDefinition(pr, item.dist);
                    col.Add(m);
                }
                else if (pr != null && item.t == imType.Flush)
                {
                    iMateDefinitions ims = def.iMateDefinitions;
                    var m = ims.AddFlushiMateDefinition(pr, item.dist);
                    col.Add(m);
                }
            }
            CompositeiMateDefinition imdef = def.iMateDefinitions.AddCompositeiMateDefinition(col, comp.name, comp.ml);
            if (rev)
                reverse(imdef);
        }
        public void reverse(CompositeiMateDefinition im)
        {
            string n = im.Name;
            if (n.IndexOf("_") == -1) return;
            string[] names = new string[1]; names[0] = n;
            im.MatchList = names;
            var spl = n.Split('_');
            n = $"{spl[1]}_{spl[0]}";
            im.Name = n;
        }
        public bool check(Edge e, iMateDefinitions ims)
        {
            foreach (var item in ims)
            {
                var im = item as InsertiMateDefinition;
                if (im == null) continue;
                if (im.Entity.Equals(e)) return true;
            }
            return false;
        }
        public List<Edge> selOp()
        {
            var lst = new List<Edge>();
            var evts = mgr.CreateInteractionEvents();
            evts.InteractionDisabled = false;
            sel = evts.SelectEvents;
            sel.WindowSelectEnabled = true;
            sel.AddSelectionFilter(SelectionFilterEnum.kPartEdgeCircularFilter);
            sel.OnSelect += Sel_OnSelect;
            var key = evts.KeyboardEvents;
            key.OnKeyPress += Key_OnKeyPress;
            evts.Start();
            evts.StatusBarText = "Выберите отверстия";
            flag = true;
            while (flag) I.app.UserInterfaceManager.DoEvents();

            foreach (var item in sel.SelectedEntities)
            {
                lst.Add(item as Edge);
            }

            evts.Stop();
            sel.OnSelect -= Sel_OnSelect;
            key.OnKeyPress -= Key_OnKeyPress;
            sel = null; key = null; evts = null;
            return lst;
        }

        private void Key_OnKeyPress(int KeyASCII)
        {
            if (KeyASCII == 32)
            {
                Edge ed = (Edge)sel.SelectedEntities[sel.SelectedEntities.Count];
                double r = ((Circle)ed.Geometry).Radius;
                Inventor.Face f = (ed.Faces[1].SurfaceType == SurfaceTypeEnum.kPlaneSurface) ? ed.Faces[1] : ed.Faces[2];
                foreach (Edge e in f.Edges)
                {
                    if (e.GeometryType == CurveTypeEnum.kCircleCurve && ((Circle)e.Geometry).Radius == r)
                    {
                        sel.AddToSelectedEntities(e);
                    }
                }
            }
            else if (KeyASCII == 13)
            {
                flag = false;
            }
        }

        private void Sel_OnSelect(ObjectsEnumerator JustSelectedEntities, SelectionDeviceEnum SelectionDevice,
            Point ModelPosition, Point2d ViewPosition, View View)
        {

        }
    }
    public class CopyProduct
    {
        PartDocument doc;
        string path, nPath;
        MyForm form;
        MyXML xml;
        string oType, nType, oBase, nBase;
        List<Document> docs = new List<Document>();
        Dictionary<Document, Document> rrefs = new Dictionary<Document, Document>();
        List<string> filter = new List<string> { ".ipt" };
        public CopyProduct(PartDocument d)
        {
            doc = d;
            path = file.p(doc.FullDocumentName);
            docs = open(filter);
            var el = getXML(u.gets<Document>(doc.ReferencingDocuments, fi => true));
            if (el == null) return;
            xml = new MyXML(el);
            form = new MyForm(xml, "Копировать");
            form.f.Controls[0].Width = 400;
            var dr = form.f.ShowDialog();
            if (dr != System.Windows.Forms.DialogResult.OK) return;
            var parts = getParts();
            nBase = form.cbs[0].Text;
            nType = form.cbs[1].Text;
            nPath = file.sub(path, "\\", 2) + "\\" + nType + "\\";
            file.dir(nPath);
            create(nPath, open(new List<string> { oBase + ".ipt" }), oBase, nBase, ref rrefs);
            var pr2 = u.getProp(rrefs.ElementAt(0).Value, "Type");
            pr2.Value = nType;
            create(nPath, open(new List<string> { oBase + ".iam" }), oBase, nBase, ref rrefs);
            create(nPath, parts, oType, nType, ref rrefs);
            var bPair = getBase();
            if (bPair.Key == null) return;
            replaceRef(bPair, rrefs);
            replaceAsm(ref rrefs);
            replaceDrw(ref rrefs);
        }
        public List<Document> open(List<string> filter)
        {
            var docs = new List<Document>();
            foreach (var item in file.getFiles(path, filter))
            {
                var d = I.open(item);
                if (d == null) continue;
                docs.Add(d);
            }
            return docs;
        }
        public void replaceAsm(ref Dictionary<Document, Document> asm)
        {
            var ds = open(new List<string> { ".iam" });
            //var asm = new Dictionary<Document, Document>();
            var fds = fDocs(ds);
            create(nPath, fds, oType, nType, ref asm);
            fDoc(asm);
        }
        public void replaceDrw(ref Dictionary<Document, Document> drw)
        {
            var ds = open(new List<string> { ".idw" });
            //var drw = new Dictionary<Document, Document>();
            var fds = fDocs(ds);
            create(nPath, fds, oType, nType, ref drw);
            fDoc(drw);
        }
        public IEnumerable<Document> fDocs(List<Document> ds)
        {
            foreach (var item in ds)
            {
                if (!rrefs.Keys.Contains(item)) yield return item;
            }
        }
        public void fDoc(Dictionary<Document, Document> docs)
        {
            foreach (var item in docs)
            {
                foreach (var r in rrefs)
                {
                    foreach (DocumentDescriptor rd in item.Value.ReferencedDocumentDescriptors)
                    {
                        if (rd.FullDocumentName == r.Key.FullDocumentName)
                        {
                            rd.ReferencedFileDescriptor.ReplaceReference(r.Value.FullDocumentName);
                        }
                    }
                }
            }
        }
        public void create(string nPath, IEnumerable<Document> docs, string o, string n,
            ref Dictionary<Document, Document> dic)
        {
            foreach (Document item in docs)
            {
                var nn = file.nameWithExt(item.FullDocumentName);
                nn = nn.Replace(o, n);
                var nfn = nPath + nn;
                if (!file.check(nfn))
                    item.SaveAs(nfn, true);
                var d = I.open(nfn);
                dic.Add(item, d);
            }
        }
        public KeyValuePair<Document, Document> getBase()
        {
            foreach (var item in rrefs)
            {
                if (item.Key.Equals(doc)) return item;
            }
            return new KeyValuePair<Document, Document>();
        }
        public void replaceRef(KeyValuePair<Document, Document> b, Dictionary<Document, Document> vals)
        {
            foreach (var item in vals)
            {
                if (item.Equals(b)) continue;
                var o = item.Key; var n = item.Value;
                foreach (DocumentDescriptor rd in n.ReferencedDocumentDescriptors)
                {
                    if (rd.FullDocumentName.Equals(b.Key.FullDocumentName))
                    {
                        rd.ReferencedFileDescriptor.ReplaceReference(b.Value.FullDocumentName);
                    }
                }
            }
        }
        public IEnumerable<Document> getParts()
        {
            foreach (var item in form.chks)
            {
                if (item.Checked)
                {
                    yield return getDoc(item.Text);
                }
            }
        }
        public Document getDoc(string n)
        {
            foreach (Document item in docs)
            {
                if (item.FullDocumentName.IndexOf(n) != -1) return item;
            }
            return null;
        }
        public XElement getXML(IEnumerable<Document> docs)
        {
            oBase = file.name(doc.FullDocumentName);
            string w = 500.ToString(), h = (docs.Count() * 15 + 30 + 100).ToString();
            oType = u.getPropValue(doc as Document, "Type");
            if (oType == "") return null;
            XElement el = new XElement("head");
            el.Add(MyXML.addXElement("Form", new Dictionary<string, string> { { "w", w}, { "h", h},
                { "x", "10" }, { "y", "30" } }));
            el.Add(MyXML.addXElement("Lbl", new Dictionary<string, string> { { "w", "100"}, {"h", "15" }, { "val",
                "Выберите детали, которые следует вновь создать"} }));
            el.Add(MyXML.addXElement("Lbl", new Dictionary<string, string> { { "w", "150" }, { "h", "15" }, { "val", oBase + " -> " } }));
            var cb = MyXML.addXElement("CB", new Dictionary<string, string> { { "w", "150" }, { "h", "15" }, { "val", oBase }, { "cur", oBase } });
            cb.Add(MyXML.addEl(cb, "el", "val", oBase));
            el.Add(cb);
            el.Add(MyXML.addXElement("Lbl", new Dictionary<string, string> { { "w", "150" }, { "h", "15" }, { "val", oType + " -> " } }));
            cb = MyXML.addXElement("CB", new Dictionary<string, string> { { "w", "150" }, { "h", "15" }, { "val", oType }, { "cur", oType } });
            cb.Add(MyXML.addEl(cb, "el", "val", oType));
            el.Add(cb);
            int i = 0;
            foreach (Document item in docs)
            {
                i++;
                var ffn = item.FullDocumentName;
                var v = file.nameWithExt(ffn);
                var e = MyXML.addXElement("Chk", new Dictionary<string, string> { { "w", "300" }, { "h", "15" }, { "val", v } });
                e.SetAttributeValue("x", "2"); e.SetAttributeValue("y", "1"); e.SetAttributeValue("pos", 3);
                //else { e.SetAttributeValue("x", "2"); e.SetAttributeValue("y", "1"); e.SetAttributeValue("pos", 2); }
                el.Add(e);
            }
            el.Add(new XElement("Btn", new XAttribute("w", 100), new XAttribute("h", 20), new XAttribute("val", "Создать")));
            return el;
        }
    }
    public class MateData
    {
        public CompositeiMateDefinition comp;
        public InsertiMateDefinition ins;
        PartComponentDefinition def;
        SurfaceBody sb;
        Plane pl;
        public static Dictionary<iMateDefinition, List<InsertiMateDefinition>> dic;
        public static List<iMateDefinition> ims = new List<iMateDefinition>();
        public MateData(Document doc, Edge e)
        {
            def = I.getPCD(doc);
            sb = e.Parent;
            set(e);
        }
        public static List<iMateDefinition> getIM(PartComponentDefinition def)
        {
            if (dic != null && dic.Count > 0) return ims;
            ims = u.gets<iMateDefinition>(def.iMateDefinitions, fi => fi is InsertiMateDefinition || fi is CompositeiMateDefinition).ToList();
            return ims;
        }
        public static iMateDefinition get(PartComponentDefinition def, Edge e)
        {
            var ims = getIM(def);
            foreach (var item in ims)
            {
                var im = item as InsertiMateDefinition;
                if (im != null && im.Entity.Equals(e)) { return item; }
                var cm = item as CompositeiMateDefinition;
                if (cm != null)
                {
                    foreach (var c in u.gets<InsertiMateDefinition>(cm, fi => fi is InsertiMateDefinition))
                    {
                        if (c.Entity.Equals(e))
                        {
                            return item;
                        }
                    }
                }
            }
            return null;
        }
        public void set(Edge e)
        {
            var ims = getIM(def);
            var tmp = get(def, e);
            if (tmp == null) return;
            ins = tmp as InsertiMateDefinition;
            comp = tmp as CompositeiMateDefinition;
        }
        public bool check(Edge e)
        {
            foreach (var item in def.iMateDefinitions)
            {
                var i = item as InsertiMateDefinition;
                if (i != null && e.Equals(i.Entity)) return true;
                var c = item as CompositeiMateDefinition;
                foreach (var ent in c)
                {
                    var i2 = item as InsertiMateDefinition;
                    if (i2 != null && e.Equals(i2.Entity)) return true;
                }
            }
            return false;
        }
        public void add(List<Edge> edges)
        {
            if (edges.Count != 2) return;
            foreach (var item in edges)
            {
                if (check(item)) return;
            }
            if (ins != null)
            {
                add(edges[0], ins);
            }
            else if (comp != null)
            {
                List<InsertiMateDefinition> lst = new List<InsertiMateDefinition>();
                if (comp.Count == 2)
                {
                    for (int i = 0; i < 2; i++)
                    {
                        var im = comp[i + 1] as InsertiMateDefinition;
                        if (im == null) return;
                        var tmp = add(edges[i], im);
                        lst.Add(tmp);
                    }
                    add(lst);
                }
            }
        }
        public CompositeiMateDefinition add(List<InsertiMateDefinition> lst)
        {
            var col = I.COC();
            //lst.Reverse();
            foreach (var item in lst)
            {
                col.Add(item);
            }
            return def.iMateDefinitions.AddCompositeiMateDefinition(col, comp.Name, comp.MatchList);
        }
        public InsertiMateDefinition add(Edge e, InsertiMateDefinition ins)
        {
            return def.iMateDefinitions.AddInsertiMateDefinition(e, false, ins.Distance.Value, null,
                    ins.Name, ins.MatchList);
        }
        public List<Edge> findEdge(object o)
        {
            List<Edge> eds = new List<Edge>();
            pl = getPlane(o);
            if (comp.Count == 2)
            {
                foreach (var im in comp)
                {
                    var insIM = im as InsertiMateDefinition;

                    if (insIM == null || insIM.Suppressed) break;
                    var e = findEdge((Edge)insIM.Entity);
                    if (e == null) break;
                    if (check(e)) break;
                    eds.Add(e);
                }
            }
            return eds;
        }
        public static Plane getPlane(object o)
        {
            WorkPlane wp = o as WorkPlane;
            Face wf = o as Face;
            var pl = wp != null ? wp.Plane : wf.Geometry as Plane;
            return pl;
        }
        public static Edge findEdge(Edge e, Plane pl, SurfaceBody sb)
        {
            Circle c = e.Geometry as Circle;
            if (c == null) return null;
            var n = u.normal(c, pl);
            n.ScaleBy(2);
            var pt = I.CP(c.Center.X, c.Center.Y, c.Center.Z); pt.TranslateBy(n);
            foreach (Edge item in sb.Edges)
            {
                var g = item.Geometry as Circle;
                if (g == null) continue;
                if (u.eq(g.Center, pt)) return item;
            }
            return null;
            //var ed = u.findAtPoint(def, pt, e.Parent, new SelectionFilterEnum[] { SelectionFilterEnum.kPartEdgeCircularFilter }) as Edge;
            //if (ed != null && ed.CurveType != CurveTypeEnum.kCircleCurve) return null;
            //return ed;
        }
        public Edge findEdge(Edge e)
        {
            return findEdge(e, pl, sb);
        }
    }
    public class BodyIMate
    {
        AssemblyDocument asm;
        PartDocument doc;
        PartComponentDefinition smcd;
        SurfaceBody sb;
        SelectSet ss;
        List<iMateDefinition> ims = new List<iMateDefinition>();
        string sb_name;
        public BodyIMate(Document doc)
        {
            if (doc.DocumentType == DocumentTypeEnum.kPartDocumentObject)
                initPart(doc);
            else if (doc.DocumentType == DocumentTypeEnum.kAssemblyDocumentObject)
                initAsm(doc);
        }
        public void initPart(Document doc)
        {
            this.doc = doc as PartDocument;
            smcd = I.getPCD(doc);
            ss = doc.SelectSet;
            add();
        }
        public void initAsm(Document doc)
        {
            asm = doc as AssemblyDocument;
            List<iMateDefinition> ims = new List<iMateDefinition>();
            foreach (var item in asm.ComponentDefinition.Occurrences)
            {
                getIMates(item, ref ims);
            }
            foreach (var item in ims)
            {
                dynamic ent = item;
                string name = ent.Name;
                var spl = name.Split('^');
                int num = u.convToInt(spl[1]);
                ComponentOccurrenceProxy occ = ent.ContainingOccurrence;
                if (occ.OccurrencePath.Count == num)
                {
                    MoveIM m = new MoveIM(doc);
                    m.move(new List<object> { item });
                }
                else if (occ.OccurrencePath.Count > num)
                {
                    ComponentOccurrence occ1 = occ.OccurrencePath[occ.OccurrencePath.Count-num];
                    var def = occ1.DefinitionReference.ReferencedDefinition;
                    var doc1 = def.Document as Document;
                    //BodyIMate b = new BodyIMate(doc1);
                    MoveIM m = new MoveIM(doc1);
                    m.move(new List<object> { item });
                }
            }
        }
        public void getIMates(object ob, ref List<iMateDefinition> ims)
        {
            dynamic ent = ob;
            foreach (var item in ent.SubOccurrences)
            {
                getIMates(item, ref ims);
            }
            foreach (var item in ent.iMateDefinitions)
            {
                string name = item.Name;
                if (name.IndexOf("^") != -1) ims.Add(item);
            }
        }
        public void add()
        {
            sb = ss[1] as SurfaceBody;
            if (sb == null) return;
            sb_name = sb.Name;
            getImates();
            foreach (var im in ims)
            {
                add(im);
            }
        }
        public void getImates()
        {
            foreach (iMateDefinition im in smcd.iMateDefinitions)
            {
                List<bool> b = new List<bool>();
                inBody(im, ref b);
                if (!b.Contains(false)) ims.Add(im);
            }
        }
        public void inBody(iMateDefinition im, ref List<bool> b)
        {
            if (im.Suppressed)
            {
                b.Add(false);
                return;
            }
            dynamic ent = im; 
            SurfaceBody body;
            if (im.Type == ObjectTypeEnum.kCompositeiMateDefinitionObject)
            {
                CompositeiMateDefinition comp = im as CompositeiMateDefinition;
                foreach (iMateDefinition item in comp)
                {
                    inBody(item, ref b); 
                }
            }
            else if (im.Type == ObjectTypeEnum.kInsertiMateDefinitionObject)
            {
                body = ent.Entity.Faces[1].SurfaceBody;
                b.Add(body.Name == sb.Name);
            }
        }
        public void add(iMateDefinition im)
        {
            List<Edge> eds = new List<Edge>();
            List<double> offsets = new List<double>();
            findEdges(im, ref eds, ref offsets);
            foreach (SurfaceBody item in smcd.SurfaceBodies)
            {
                if (item.Name == sb_name) continue;
                var eds1 = findEdges(eds, item, offsets);
                if (eds1.Count != eds.Count) continue;
                if (check(eds1)) continue;
                bool d = dir(eds[0], eds1[0]);
                int i = 0;
                var comp = addIM(im, eds1, ref i);
                if (comp == null) continue;
                string n = d ? im.Name : im.Name + "!";
                if (comp.Type == ObjectTypeEnum.kInsertiMateDefinitionObject)
                {
                    IMate.addName(comp as InsertiMateDefinition, n);
                }
                else if (comp.Type == ObjectTypeEnum.kCompositeiMateDefinitionObject)
                {
                    IMate.addName(comp as CompositeiMateDefinition, n);
                }
            }
        }
        public iMateDefinition addIM(iMateDefinition im, List<Edge> eds, ref int i)
        {
            iMateDefinition def = null;
            if (im.Type == ObjectTypeEnum.kCompositeiMateDefinitionObject)
            {
                var col = I.COC();
                CompositeiMateDefinition comp = im as CompositeiMateDefinition;
                foreach (iMateDefinition item in comp)
                {
                    col.Add(addIM(item, eds, ref i));
                    
                }
                def = smcd.iMateDefinitions.AddCompositeiMateDefinition(col) as iMateDefinition;
            }
            else if (im.Type == ObjectTypeEnum.kInsertiMateDefinitionObject)
            {
                var ins = im as InsertiMateDefinition;
                def = smcd.iMateDefinitions.AddInsertiMateDefinition(eds[i], ins.AxesOpposed, ins.Distance.Expression) as iMateDefinition;
            }
            i++;
            return def;
        }
        public bool dir(Edge e1, Edge e2)
        {
            Circle c1 = e1.Geometry as Circle, c2 = e2.Geometry as Circle;
            return u.isCollinear(c1.Normal, c2.Normal);
        }
        public List<Edge> findEdges(List<Edge> eds, SurfaceBody body, List<double> offsets)
        {
            List<Edge> eds1 = new List<Edge>();
            int i = 0;
            foreach (Edge item in eds)
            {
                double offset = offsets[i];
                var e = u.get<Edge>(body.Edges, f => f.GeometryType == CurveTypeEnum.kCircleCurve && eq(item, f, offset));
                if (e != null) eds1.Add(e);
                i++;
            }
            return eds1;
        }
        public bool eq(Edge e1, Edge e2, double offset)
        {
            Circle c1 = e1.Geometry as Circle, c2 = e2.Geometry as Circle;
            Point p1 = c1.Center, p2 = c2.Center;
            if (offset != 0)
            {
                var v = c1.Normal.AsVector(); v.ScaleBy(-offset);
                p1.TranslateBy(v);
            }
            return u.eq(p1, p2);
        }
        public void findEdges(iMateDefinition im, ref List<Edge> eds, ref List<double> offsets)
        {
            dynamic ent = im;
            Edge ed;
            if (im.Type == ObjectTypeEnum.kCompositeiMateDefinitionObject)
            {
                CompositeiMateDefinition comp = im as CompositeiMateDefinition;
                foreach (iMateDefinition item in comp)
                {
                    findEdges(item, ref eds, ref offsets);
                }
            }
            else if (im.Type == ObjectTypeEnum.kInsertiMateDefinitionObject)
            {
                ed = ent.Entity;
                offsets.Add(ent.Distance.ModelValue);
                eds.Add(ed);
            }
        }
        public SurfaceBody GetBody(string name)
        {
            return u.get<SurfaceBody>(smcd.SurfaceBodies, f => f.Name == name);
        }
        public bool check(List<Edge> eds)
        {
            foreach (iMateDefinition im in doc.ComponentDefinition.iMateDefinitions)
            {
                if (im.Suppressed) continue;
                if (im.Type == ObjectTypeEnum.kCompositeiMateDefinitionObject)
                {
                    var cm = im as CompositeiMateDefinition;
                    if (eds.Count != cm.Count) return false;
                    int i = 0, count = 0;
                    List<bool> lst = new List<bool>();
                    foreach (var m in cm)
                    {
                        var e = getEnt(m);
                        if (e.Equals(eds[i])) count++;
                        
                        i++;
                    }
                    if (count == eds.Count) return true;
                }
                else
                {
                    var e = getEnt(im);
                    if (e.Equals(eds[0])) return true;
                }
            }
            return false;
        }
        public object getEnt(object m)
        {
            var im = m as InsertiMateDefinition;
            if (im != null) return im.Entity;
            var am = m as AngleiMateDefinition;
            if (am != null) return am.Entity;
            var fm = m as FlushiMateDefinition;
            if (fm != null) return fm.Entity;
            var mm = m as MateiMateDefinition;
            if (mm != null) return mm.Entity;
            return null;
        }
    }
    public class MirrorIMate
    {
        PartDocument doc;
        PartComponentDefinition def;
        HashSet<Point> filter = new HashSet<Point>();
        HashSet<iMateDefinition> ims = new HashSet<iMateDefinition>();
        public MirrorIMate(Document d)
        {
            if (d.DocumentType != DocumentTypeEnum.kPartDocumentObject) return;
            doc = d as PartDocument;
            def = doc.ComponentDefinition;
            fill();
            find();
        }
        public void find()
        {
            foreach (MirrorFeature mf in def.Features.MirrorFeatures)
            {
                if (mf.Suppressed) continue;
                var fs = getFaces(mf);
                var pl = mf.Definition.MirrorPlaneEntity;
                foreach (var item in fs)
                {
                    var es = findEdge(item, pl);
                    foreach (var e in es)
                    {
                        var tmp = MateData.get(def, e.Key);
                        if (tmp != null) continue;
                        var md = new MateData((Document)doc, e.Value);
                        if (md.ins != null)
                        {
                            if (ims.Contains(md.ins as iMateDefinition)) continue;
                            md.add(new List<Edge> { e.Key });
                            ims.Add(md.ins as iMateDefinition);
                        }
                        else if (md.comp != null)
                        {
                            if (ims.Contains(md.comp as iMateDefinition)) continue;
                            List<Edge> eds = new List<Edge>();
                            eds = md.findEdge(pl);
                            if (eds.Count == 2)
                            {
                                u.action<Edge>(eds, a => filter_add(a));
                                md.add(eds);
                                ims.Add(md.comp as iMateDefinition);
                            }
                        }
                    }
                }
            }
        }
        public Point get_point(object o)
        {
            Edge e = o as Edge;
            var g = e.Geometry as Circle;
            if (g == null) return null;
            return g.Center;
        }
        public bool filter_check(object o)
        {
            var pt = get_point(o);
            foreach (var item in filter)
            {
                if (u.eq(item, pt)) return true;
            }
            return false;
        }
        public void filter_add(object o)
        {
            filter.Add(get_point(o));
        }
        public void fill()
        {
            foreach (var item in def.iMateDefinitions)
            {
                InsertiMateDefinition i = item as InsertiMateDefinition;
                if (i != null) filter_add(i.Entity);
                var c = item as CompositeiMateDefinition;
                if (c != null)
                {
                    foreach (var im in c)
                    {
                        i = im as InsertiMateDefinition;
                        if (im != null) filter_add(i.Entity);
                    }
                }
            }
        }
        public List<Face> getFaces(MirrorFeature mf)
        {
            List<Face> faces = new List<Face>();
            foreach (Face f in mf.Faces)
            {
                if (f.SurfaceType != SurfaceTypeEnum.kCylinderSurface) continue;
                if (f.TangentiallyConnectedFaces.Count != 0) continue;
                var el = check(f);
                if (el == null) faces.Add(f);
            }
            return faces;
        }
        public Edge check(Face f)
        {
            foreach (Edge item in f.Edges)
            {
                if (check(item) != null) return item;
            }
            return null;
        }
        public InsertiMateDefinition check(Edge e)
        {
            foreach (var item in def.iMateDefinitions)
            {
                InsertiMateDefinition im = item as InsertiMateDefinition;
                if (im == null) continue;
                if (im.Entity.Equals(e)) return im;
            }
            return null;
        }
        public Dictionary<Edge, Edge> findEdge(Face f, object o)
        {
            var pl = MateData.getPlane(o);
            var dic = new Dictionary<Edge, Edge>();
            var eds = u.gets<Edge>(f.Edges, fi => fi.CurveType == CurveTypeEnum.kCircleCurve);
            foreach (Edge e in eds)
            {
                var ed = MateData.findEdge(e, pl, e.Parent);
                if (ed != null) dic[e] = ed;
            }
            return dic;
        }

        public bool checkEdge(Edge e)
        {
            var f = u.get<Face>(e.Faces, fi => fi.SurfaceType == SurfaceTypeEnum.kCylinderSurface);
            if (!(f.CreatedByFeature is MirrorFeature)) return true;
            return false;
        }
    }
    internal class IMatesBtn : Button
    {
        //public static Drawings m_Drw;
        //public static Drawings getDrw { get { return m_Drw; } }
        public IMatesBtn(string displayName, string internalName, string clientId, string description, string tooltip,
            ButtonDisplayEnum buttonDisplayType = ButtonDisplayEnum.kDisplayTextInLearningMode, CommandTypesEnum commandType = CommandTypesEnum.kNonShapeEditCmdType)
            : base(displayName, internalName, commandType, clientId, description, tooltip, buttonDisplayType) { }
        protected override void ButtonDefinition_OnExecute(NameValueMap context)
        {
            IMates ims = new IMates(I.aDoc());
        }
    }
    public class ISettings : InvComboBox
    {
        public List<string> filter = new List<string>() { "Hole" };
        string name, cur;
        public ISettings(string displayName, string internalName, int w, string join = "",
            CommandTypesEnum commandType = CommandTypesEnum.kShapeEditCmdType)
            : base(displayName, internalName, commandType, w)
        {
            name = join; cur = displayName;
        }
        protected override void ComboBoxDefinition_OnSelect(NameValueMap nvm)
        {
            if (!filter.Contains(name)) return;
            int ind = this._ComboBoxDef.ListIndex;
            u.setCBValue(name, ind, cur);
        }
        public void fill(List<string> vals, int index)
        {
            if (filter.Contains(name))
            {
                int i = u.getCBIndex(name, cur);
                if (i != -1) index = i;
            }
            foreach (var item in vals)
            {
                this._ComboBoxDef.AddItem(item);
            }
            if (index != 0)
                this._ComboBoxDef.ListIndex = index;
        }
        public string get()
        {
            return this._ComboBoxDef.Text;
        }
    }
}
