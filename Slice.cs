using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
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
    class Slice
    {
        readonly Document doc;
        readonly SheetMetalComponentDefinition smcd;
        readonly SheetMetalFeatures smf;
        UnfoldFeature uf;
        RefoldFeature rf;
        PlanarSketch ps;
        Point2d bpt = null;
        List<SketchPoint> filterPts = new List<SketchPoint>();
        double t;
        double min, max;
        public Slice(Document d)
        {
            doc = d;
            smcd = I.getSMCD(doc);
            t = (double)smcd.Thickness.Value;
            smf = smcd.Features as SheetMetalFeatures;
        }
        public Face getFace()
        {
            var sb = smcd.SurfaceBodies[1];
            var fs = u.gets<Face>(sb.Faces, fi => fi.SurfaceType == SurfaceTypeEnum.kPlaneSurface).
                OrderByDescending(el => el.Evaluator.Area);
            return fs.ElementAt(1);
        }
        public void unfold()
        {
            var f = getFace();
            if (f == null) return;
            uf = smf.UnfoldFeatures.Add(f);
            try
            {
                rf = smf.RefoldFeatures.Add(uf.StationaryFace);
                rf.SetEndOfPart(true);
            }
            catch (Exception)
            {
            }
        }
        public void findSpl(IEnumerable<SketchEntity> ents)
        {
            List<SketchSpline> filterSpl = new List<SketchSpline>();
            double tol = 5;
            var spls = u.gets<SketchEntity>(ents, fi => fi.Type == ObjectTypeEnum.kSketchSplineObject && !fi.Construction).
                Where(el => el.RangeBox.MinPoint.VectorTo(el.RangeBox.MaxPoint).Length > tol).Select(el => el as SketchSpline);
            for (int i = 0; i < spls.Count(); i++)
            {
                var spl1 = spls.ElementAt(i);
                if (filterSpl.Contains(spl1)) continue;
                var spl2 = findSpl(spl1, spls);
                if (filterSpl.Contains(spl2)) continue;
                filterSpl.Add(spl1); filterSpl.Add(spl2);
                if (spl1 != null && spl2 != null)
                {
                    spl1.Construction = true; spl2.Construction = true;
                    var min = getMinEnt((SketchEntity)spl1, (SketchEntity)spl2);
                    var max = min.Equals((SketchEntity)spl1) ? spl2 : spl1;
                    culv(min, (SketchEntity)max);
                    
                    List<SketchPoint> pts = new List<SketchPoint> { spl1.StartSketchPoint, spl1.EndSketchPoint,
                    spl2.StartSketchPoint, spl2.EndSketchPoint};
                    var p = fill(pts);
                    if (p.Count() == 2)
                    {
                        SketchPoint p1 = p.ElementAt(0), p2 = p.ElementAt(1);
                        Box2d b = max.Geometry.Evaluator.RangeBox;

                        Point2d t1 = I.CP2d(b.MinPoint.X, b.MaxPoint.Y), t2 = I.CP2d(b.MaxPoint.X, b.MinPoint.Y);
                        //ps.TextBoxes.AddFitted(t1, "1"); ps.TextBoxes.AddFitted(t2, "2");
                        var pt1 = ps.SketchToModelSpace(t1);
                        SelectionFilterEnum[] sf = { SelectionFilterEnum.kAllPlanarEntities };
                        var col = smcd.FindUsingPoint(pt1, ref sf);
                        Point2d mp = col.Count != 0 ? t1 : t2;
                        if (inters((SketchEntity)spl1, (SketchEntity)spl2))
                        {
                            var l = ps.SketchLines.AddByTwoPoints(mp, p1);
                            var l2 = ps.SketchLines.AddByTwoPoints(l.StartSketchPoint, p2);
                        }
                        else
                        {
                            var l = ps.SketchLines.AddByTwoPoints(p1, p2);
                            //max.Construction = false;
                        }
                        //var spl = getSpl(p1);
                        //if (spl != null) ps.GeometricConstraints.AddTangent((SketchEntity)spl,(SketchEntity) l);
                        //spl = getSpl(p2);
                        //if (spl != null) ps.GeometricConstraints.AddTangent((SketchEntity)spl, (SketchEntity)l2);
                    }
                }
                //break;
            }
        }
        public List<SketchPoint> fill(List<SketchPoint> pts)
        {
            List<SketchPoint> lst = new List<SketchPoint>();
            bool add = true;
            foreach (var item in pts)
            {
                foreach (var e in item.AttachedEntities)
                {
                    if (e is SketchArc) {add = false; break; }
                }
                if (add)
                lst.Add(item);
                add = true;
            }
            return lst;
        }
        public SketchSpline getSpl(SketchPoint pt)
        {
            foreach (var item in pt.AttachedEntities)
            {
                if (item is SketchSpline) return item as SketchSpline;
            }
            return null;
        }
        public SketchSpline findSpl(SketchSpline spl, IEnumerable<SketchSpline> spls)
        {
            double tol = 0.2;
            var mp = u.midPt(spl.RangeBox);
            foreach (var item in spls)
            {
                if (item.Construction) continue;
                if (item.Equals(spl)) continue;
                var pt = u.midPt(item.RangeBox);
                var v = I.CV2d(pt, mp);
                if (v.Length < tol) return item;
            }
            return null;
        }
        public void getCulv(Curve2dEvaluator ev, int count, ref double[] culv, ref double[] pars, ref double[] dirs)
        {
            double min, max;
            ev.GetParamExtents(out min, out max);
            pars = new double[count];
            pars[0] = min; pars[count-1] = max;
            double step = (max - min) / count;
            for (int i = 1; i < count - 1; i++)
            {
                pars[i] = min + step * i;
            }
            ev.GetCurvature(ref pars, ref dirs, ref culv);
        }
        public List<SketchPoint> culv(SketchEntity ent1, SketchEntity ent2)
        {  
            SketchSpline spl1 = ent1 as SketchSpline, spl2 = ent2 as SketchSpline;
            List<SketchPoint> points = new List<SketchPoint>() { spl1.StartSketchPoint, spl1.EndSketchPoint,
            spl2.StartSketchPoint, spl2.EndSketchPoint};
            List<SketchPoint> filterPts = new List<SketchPoint>(); filterPts.Add(spl1.StartSketchPoint);
            int count = 10;
            var ev1 = spl1.Geometry.Evaluator;
            var ev2 = spl2.Geometry.Evaluator;
            double[] pars1 = { }, pars2 = { };
            double[] dirs1 = { }, dirs2 = { };
            double[] culv1 = { }, culv2 = { };
            double[] fder = { }, sder = { };
            getCulv(ev1, count, ref culv1, ref pars1, ref dirs1);
            getCulv(ev2, count, ref culv2, ref pars2, ref dirs2);
            culv2 = culv2.Reverse().ToArray(); pars2 = pars2.Reverse().ToArray();
            //ev1.GetFirstDerivatives(ref pars1, ref fder);
            ev1.GetSecondDerivatives(ref pars1, ref sder);
            var steps = average(sder);
            //return null;
            SketchArc sa = null;
            int s = 0;
            bool close = false;
            bool first = true;
            for (int i = 1; i < steps.Count; i++)
            {
                var tmpSP = getSP(spl1, first, out sa);
                first = false;
                if (steps[i] > pars1.Length-2)
                {
                    var p = getSP(spl1, spl2, spl1.StartSketchPoint);
                    //ps.TextBoxes.AddFitted(p.Geometry, "p1");
                    addArc(ev1, ev2, ps, pars1, pars2, steps[i-1], steps[i], tmpSP, sa, p);
                    filterPts.Add(tmpSP);
                    close = true;
                    continue;
                }
                var a = addArc(ev1, ev2, ps, pars1, pars2, steps[i-1], steps[i], tmpSP, sa, null);
                ps.DimensionConstraints.AddTwoPointDistance(a.StartSketchPoint, spl1.StartSketchPoint,
                   DimensionOrientationEnum.kAlignedDim, u.midPt(a.StartSketchPoint.Geometry, spl1.StartSketchPoint.Geometry));
            }
            if (!close)
            {
                var tmpSP = getSP(spl1, first, out sa);
                var p = getSP(spl1, spl2, spl1.StartSketchPoint);
                addArc(ev1, ev2, ps, pars1, pars2, s, pars1.Length-1, tmpSP, sa, p);
                filterPts.Add(tmpSP);
            }
            return filterPts;
        }
        public List<int> average(double [] arr)
        {
            double tol = 1.2;
            double m = 0;
            List<int> vals = new List<int>();
            vals.Add(0);
            for (int i = 0; i < arr.Length; i+=2)
            {
                if (m == 0) { m = arr[i]; continue; }
                var tmp = arr[i];
                double diff = Math.Abs(tmp - m);
                if (diff > tol) { vals.Add(i/2); m = tmp; }
            }
            return vals;
        }
        public SketchPoint getSP(SketchSpline e1, SketchSpline e2, SketchPoint sp)
        {
            if (inters((SketchEntity)e1, (SketchEntity)e2))
            {
                SketchSpline spl;
                spl = (check(e1, sp)) ? e2 : e1;
                return getMin(sp, (SketchEntity)spl, (a, b) => a > b);
            }
            else
            {
                return e1.StartSketchPoint.Equals(sp) ? e1.EndSketchPoint : e1.StartSketchPoint;
            }
        }
        public bool check(SketchSpline s, SketchPoint sp)
        {
            return s.StartSketchPoint.Equals(sp) || s.EndSketchPoint.Equals(sp);
        }
        public SketchPoint getSP(SketchSpline spl, bool first, out SketchArc sa)
        {
            sa = null;
            if (ps.SketchArcs.Count > 0 && !first) sa = ps.SketchArcs[ps.SketchArcs.Count];
            return sa == null ? spl.StartSketchPoint : sa.StartSketchPoint;
        }
        public SketchArc addArc(Curve2dEvaluator ev1, Curve2dEvaluator ev2, PlanarSketch ps, double[] pars1, double[] pars2, 
            int s, int e, SketchPoint pt, SketchArc sa, SketchPoint end)
        {
            object spt1, ept1, mpt1, spt2, ept2, mpt2, spt, ept, mpt;
            getPts(pars1, ev1, s, e, out spt1, out mpt1, out ept1);
            getPts(pars2, ev2, s, e, out spt2, out mpt2, out ept2);
            spt = spt1; ept = ept1; mpt = mpt1;
            if (pt != null)
            {
                spt = pt;
                //ps.TextBoxes.AddFitted(pt.Geometry, "start");
            }
            if (end != null)
            {
                //var v = end.Geometry.VectorTo(sa.StartSketchPoint.Geometry);
                //if (v.Length < 0.01)
                //{
                //    //ps.GeometricConstraints.AddCoincident((SketchEntity)sa.StartSketchPoint, (SketchEntity)end);
                //    return;
                //}
                ept = end;
                mpt = getMP(pars1, ev1);
                //ps.TextBoxes.AddFitted(end.Geometry, "end");
            }
            SketchArc a = ps.SketchArcs.AddByThreePoints(spt, mpt as Point2d, ept);
            if (end != null) return a;
            if (sa != null)
            {
                ps.GeometricConstraints.AddTangent((SketchEntity)sa, (SketchEntity)a);
            }
            return a;
        }
        public void getPts(double[] pars, Curve2dEvaluator ev, int s, int e, out object sp, out object mp, out object ep)
        {
            double sp1 = pars[s], ep1 = pars[e], mp1 = (pars[s] + pars[e]) * 0.5;
            sp = u.getPointAtParam(ev, sp1); ep = u.getPointAtParam(ev, ep1);
            mp = u.getPointAtParam(ev, mp1);
        }
        public object getMP(double[] pars, Curve2dEvaluator ev)
        {
            return u.getPointAtParam(ev, pars[pars.Length - 1] - 0.05);
        }
        public void create()
        {
            var ss = doc.SelectSet;
            if (ss.Count != 1) return;
            var e = ss[1] as Edge;
            if (e == null) return;
            var f = u.get<Face>(e.Faces, fi => fi.SurfaceType == SurfaceTypeEnum.kPlaneSurface);
            ps = smcd.Sketches.Add(f);
            var a = ps.AddByProjectingEntity(e); a.Construction = true;
            divide(a);
            addCut();
        }
        public void divide(object a)
        {
            double tol = 0.05;
            var spl = a as SketchSpline;
            var arc = a as SketchEllipticalArc;
            EllipticalArc2d ell = null;
            BSplineCurve2d bspl = null;
            Curve2dEvaluator ev = null;
            object curve = null;

            if (spl != null)
            {
                bspl = spl.Geometry;
                bspl.Evaluator.GetParamExtents(out min, out max);
                ev = spl.Geometry.Evaluator;
                curve = bspl;
            }
            else if (arc != null)
            {
                ell = arc.Geometry;
                ell.Evaluator.GetParamExtents(out min, out max);
                ev = arc.Geometry.Evaluator;
                curve = ell;
            }
            if (curve == null) return;
            
            //int i = 0;
            while (max - min > tol)
            {
                divide(curve, ev, max);
                var l = ps.SketchLines[ps.SketchLines.Count];
                ps.GeometricConstraints.AddCoincident((SketchEntity)l.EndSketchPoint, (SketchEntity)a);
                //i++;
            }
            var col = I.COC(a);
            var en = ps.OffsetSketchEntitiesUsingPoint(col, bpt);
            foreach (var item in en)
            {
                SketchEntity ent = item as SketchEntity; 
                SketchPoint sp1, sp2;
                var el = item as SketchEllipticalArc;
                var el1 = item as SketchOffsetSpline;
                if (el != null) { sp1 = el.StartSketchPoint; sp2 = el.EndSketchPoint; }
                else if (el1 != null)
                {
                    sp1 = el1.StartSketchPoint; sp2 = el1.EndSketchPoint;
                }
                else return;
                var tmp1 = getMin(sp1, (SketchEntity)ps.SketchLines[1], (l1, l2) => l1 < l2);
                var tmp2 = getMin(sp2, (SketchEntity)ps.SketchLines[ps.SketchLines.Count], (l1, l2) => l1 < l2);
                ps.SketchLines.AddByTwoPoints(tmp1, sp1);
                ps.SketchLines.AddByTwoPoints(tmp2, sp2);
            }
        }
        public SketchPoint getMin(SketchPoint sp, SketchEntity sl, Func<double, double, bool> f)
        {
            SketchPoint sp1 = getSP(sl), sp2 = getSP(sl, true);
            Vector2d v1 = sp.Geometry.VectorTo(sp1.Geometry), v2 = sp.Geometry.VectorTo(sp2.Geometry);
            return f(v1.Length,v2.Length) ? sp1 : sp2;
        }
        public void divide(object a,Curve2dEvaluator ev, double p2)
        {            
            double tol = 1;
            double[] pars1 = { min}, pars2 = { p2 };
            double[] coord1 = { }, coord2 = { };
            ev.GetPointAtParam(ref pars1, ref coord1);
            var pt1 = u.getPt(coord1);
            ev.GetPointAtParam(ref pars2, ref coord2);
            var pt2 = u.getPt(coord2);
            
            var mp = u.midPt(pt1, pt2);
            var v = pt1.VectorTo(pt2);
            var n = I.CV2d(v); u.normal(v, null); n.Normalize();
            var l = I.tg.CreateLine2d(mp, v.AsUnitVector());
            var col = l.IntersectWithCurve(a);
            foreach (var item in col)
            {
                Point2d p = item as Point2d;
                v = mp.VectorTo(p);
                p.GetPointData(ref coord1);
                double[] gp = { }, md = { };
                SolutionNatureEnum[] sol = { };
                ev.GetParamAtPoint(ref coord1, ref gp, ref md, ref pars1, ref sol);
                var v1 = pt1.VectorTo(pt2);
                if (v1.Length > tol)
                {
                    divide(a, ev, pars1[0]);
                }
                else
                {
                    //var sp1 = getSP(min);
                    //if (sp1 != null) ps.SketchLines.AddByTwoPoints(sp1, pt1);
                    //else
                    SketchLine sl;
                    if (ps.SketchLines.Count > 0)
                    {
                        sl = ps.SketchLines[ps.SketchLines.Count];
                        ps.SketchLines.AddByTwoPoints(sl.EndSketchPoint, pt2);
                    } else
                    ps.SketchLines.AddByTwoPoints(pt1, pt2);
                    //var sl = ps.SketchLines.AddByTwoPoints(mp, p); sl.Construction = true;
                    if (bpt == null)
                    {
                        var vec = mp.VectorTo(p); vec.ScaleBy(4);
                        bpt = I.CP2d(mp);
                        bpt.TranslateBy(vec);
                    }
                    min = p2;
                }  
            }
        }
        public SketchPoint getSP(Point2d pt, SketchLine l)
        {
            var sl = u.get<SketchLine>(ps.SketchLines, fi => !fi.Construction && !fi.Equals(l) && 
            (u.eq(pt, fi.StartSketchPoint.Geometry) || u.eq(pt, fi.EndSketchPoint.Geometry)));
            if (sl != null) return u.eq(sl.StartSketchPoint.Geometry, pt) ? sl.StartSketchPoint : sl.StartSketchPoint;
            return null;
        }
        public void addSketch()
        {
            if (uf == null)
            {
                uf = smf.UnfoldFeatures[1];
            }
            ps = smcd.Sketches.Add(uf.StationaryFace);
            SurfaceBody sb = smcd.SurfaceBodies[1];
            var vec = ps.PlanarEntityGeometry.Normal.AsVector();
            var es = getEdges(getEnt(sb, ps), vec);
            List<SketchEntity> ents = new List<SketchEntity>();
            List<SketchSpline> spl = new List<SketchSpline>();
            foreach (var item in es)
            {
                SketchEntity se = ps.AddByProjectingEntity(item);
                ents.Add(se);
            }
            //return;
            joinLine(ents, fi => fi.AttachedEntities.Count == 1, true);
            joinLine(ents, fi => fi.AttachedEntities.Count == 2, false);
            findSpl(ents);
            addCut();
        }
        public void addExtrude()
        {
            if (uf == null)
            {
                uf = smf.UnfoldFeatures[1];
            }
            ps = smcd.Sketches.Add(uf.StationaryFace);
            SurfaceBody sb = smcd.SurfaceBodies[1];
            var vec = ps.PlanarEntityGeometry.Normal.AsVector();
            var es = getEnt(sb, ps);
            List<SketchEntity> ents = new List<SketchEntity>();
            List<SketchSpline> spl = new List<SketchSpline>();
            foreach (var item in es)
            {
                foreach (var ed in item.Edges)
                {
                    SketchEntity se = ps.AddByProjectingEntity(ed);
                    ents.Add(se);
                }
            }
            addFace();
        }
        public double getL(SketchEntity se)
        {
            return getVec(se).Length;
        }
        public void joinLine(List<SketchEntity> ls, Func<SketchPoint, bool> f, bool ends)
        {
            IEnumerable<SketchEntity> ents = !ends ? u.gets<SketchEntity>(ps.SketchSplines, fi => true) :
                u.gets<SketchEntity>(ls, fi => true);
            
            foreach (var item in ls)
            {
                join(item, fi => f(fi), ents, ends);
            }
        }
        public void addLines(List<SketchLine> ls)
        {
            List<SketchPoint> addLines = new List<SketchPoint>();
            for (int i = 0; i < ls.Count-3; i += 2)
            {
                var sl1 = ls[i]; var sl2 = ls[i + 2];
                var pts = getPts(sl1, sl2);
                var v = pts[1].Geometry.VectorTo(pts[2].Geometry);
                if (u.eq(v.Length, 0)) continue;
                if (v.Length < 20)
                    ps.SketchLines.AddByTwoPoints(pts[1], pts[2]);
                sl1 = ls[i + 1]; sl2 = ls[i + 3];
                pts = getPts(sl1, sl2);
                v = pts[1].Geometry.VectorTo(pts[2].Geometry);
                if (u.eq(v.Length, 0)) continue;
                if (v.Length < 20)
                    ps.SketchLines.AddByTwoPoints(pts[1], pts[2]);
            }
        }
        public void addFace()
        {
            var pr = ps.Profiles.AddForSolid();
            var def = smf.FaceFeatures.CreateFaceFeatureDefinition(pr);
            def.Direction = PartFeatureExtentDirectionEnum.kNegativeExtentDirection;
            smf.FaceFeatures.Add(def);
        }
        public void addCut()
        {
            var pr = ps.Profiles.AddForSolid();
            var def = smf.CutFeatures.CreateCutDefinition(pr);
            smf.CutFeatures.Add(def);
        }
        public void addEnd(List<SketchLine> ls)
        {
            for (int i = 0; i < ls.Count; i+=2)
            {
                var sp1 = getPt(ls[i],ls); var sp2 = getPt(ls[i + 1],ls);
                if (sp1 == null || sp2 == null) continue;
                if (addGap(sp1, sp2))
                ps.SketchLines.AddByTwoPoints(sp1, sp2);
            }
        }
        public void join(SketchEntity se, Func<SketchPoint, bool> f, IEnumerable<SketchEntity> pts, bool ends)
        {
            SketchPoint sp1 = getSP(se), sp2 = getSP(se, false);
            join(sp1, f, pts, ends); join(sp2, f, pts, ends);
        }
        public void join(SketchPoint sp, Func<SketchPoint, bool> f, IEnumerable<SketchEntity> pts, bool ends)
        {
            SketchEntity se = getAtt(sp);
            var vec = getVec(se);
            double d = 1.5;
            var ents = u.gets<SketchPoint>(getPts(pts), fi => fi.Constraints.Count == 0 && f(fi) &&
            fi.Geometry.VectorTo(sp.Geometry).Length < d);
            var hs = getHS(ents); 
            if (hs.Count == 2 && ends) { /*return;*/ ps.SketchLines.AddByTwoPoints(hs.ElementAt(0), hs.ElementAt(1)); }
            else if (hs.Count >= 4 && !ends)
            {
                IEnumerable<SketchPoint> spts = hs.OrderBy(el => el.Geometry.VectorTo(sp.Geometry).Length).Take(4);
                var p1 = spts.ElementAt(0); var p2 = spts.ElementAt(1); var p3 = spts.ElementAt(2); var p4 = spts.ElementAt(3);
                if (check(spts)) return;
                SketchEntity se1 = null, se2 = null;
                getEnts(spts, out se1, out se2);
                
                if (se1 == null || se2 == null) return;

                var tmp1 = getSP(se1); var tmp2 = getSP(se1, false);
                var tmp3 = getSP(se2); var tmp4 = getSP(se2, false);
                var v1 = tmp1.Geometry.VectorTo(tmp2.Geometry);
                var v2 = tmp3.Geometry.VectorTo(tmp4.Geometry);
                //ps.TextBoxes.AddFitted(tmp1.Geometry, "p1");
                //ps.TextBoxes.AddFitted(tmp3.Geometry, "p2");

                //Projection pr1 = new Projection(v1, false);
                //pr1.set(tmp1.Geometry, tmp2.Geometry);
                var n1 = I.CV2d(v1); u.normal(n1, null); n1.Normalize();
                double d1 = u.dotProduct(n1, tmp3.Geometry.VectorTo(tmp1.Geometry)),
                    d2 = u.dotProduct(n1, tmp4.Geometry.VectorTo(tmp1.Geometry));
                if (inters(se1, se2))
                {
                    if (se1 != null) se1.Construction = true;
                    if (se2 != null) se2.Construction = true;
                    addSL(se2, tmp1); addSL(se2, tmp2);
                }
                else addSL(se1, se2);
            }
        }
        public bool inters(SketchEntity se1, SketchEntity se2)
        {
            var tmp1 = getSP(se1); var tmp2 = getSP(se1, false);
            var tmp3 = getSP(se2); var tmp4 = getSP(se2, false);
            var v1 = tmp1.Geometry.VectorTo(tmp2.Geometry);
            var n1 = I.CV2d(v1); u.normal(n1, null); n1.Normalize();
            double d1 = u.dotProduct(n1, tmp3.Geometry.VectorTo(tmp1.Geometry)),
                d2 = u.dotProduct(n1, tmp4.Geometry.VectorTo(tmp1.Geometry));
            return Math.Sign(d1) != Math.Sign(d2);
        }
        public void getEnts(IEnumerable<SketchPoint> pts, out SketchEntity se1, out SketchEntity se2)
        {
            se1 = null; se2 = null;
            foreach (var item in pts)
            {
                SketchPoint sp = item;
                foreach (var p in pts)
                {
                    if (p.Equals(sp)) continue;
                    if (se1 == null) se1 = getSE(sp, p);
                    else if (se2 == null)
                    {
                        se2 = getSE(sp, p);
                        if (se2 != null && se2.Equals(se1)) se2 = null;
                    }
                }
            }
        }
        public SketchEntity getAtt(SketchPoint sp)
        {
            if (sp.AttachedEntities.Count == 1) return sp.AttachedEntities[1];
            SketchEntity se1 = sp.AttachedEntities[1], se2 = sp.AttachedEntities[2];
            if (se1 is SketchSpline) return se1;
            else if (se2 is SketchSpline) return se2;
            else return se1;
        }
        public SketchEntity getMinEnt(SketchEntity e1, SketchEntity e2)
        {
            Face f = ps.PlanarEntity as Face;
            Point pt = u.midPt(f.Evaluator.RangeBox);
            Point2d spt = ps.ModelToSketchSpace(pt);
            Point2d mp1 = getMP(e1), mp2 = getMP(e2);
            Vector2d v1 = mp1.VectorTo(spt), v2 = mp2.VectorTo(spt);
            return v1.Length < v2.Length ? e1 : e2;

            //Vector n = ps.PlanarEntityGeometry.Normal.AsVector(),
            //    v1 = sp1.Geometry3d.VectorTo(ps.OriginPointGeometry);
            //SketchEntity se = e2;
            //double d = u.dotProduct(v1, n);
            //if (u.eq(d, 0)) se = e1;
            //return se;
        }
        public void addSL(SketchEntity e1, SketchEntity e2)
        {
            var se = getMinEnt(e1, e2);
            se.Construction = true;
            ps.SketchLines.AddByTwoPoints(getSP(se), getSP(se,false));
        }
        public HashSet<SketchPoint> getHS(IEnumerable<SketchPoint> pts)
        {
            HashSet<SketchPoint> h = new HashSet<SketchPoint>();
            foreach (var item in pts)
            {
                h.Add(item);
            }
            return h;
        }
        public bool check(IEnumerable<SketchPoint> pts)
        {
            foreach (var sp in pts)
            {
                if (filterPts.Contains(sp)) return true;
                filterPts.Add(sp);
            }
            return false;
        }
        public void addSL(SketchEntity e, SketchPoint sp)
        {
            var p1 = getSP(e); var p2 = getSP(e, false);
            var v1 = p1.Geometry.VectorTo(sp.Geometry); var v2 = p2.Geometry.VectorTo(sp.Geometry);
            if (v1.Length < v2.Length) ps.SketchLines.AddByTwoPoints(sp, p2);
            else ps.SketchLines.AddByTwoPoints(sp, p1);
        }
        public HashSet<SketchPoint> getPts(IEnumerable<SketchEntity> ents)
        {
            HashSet<SketchPoint> pts = new HashSet<SketchPoint>();
            foreach (var item in ents)
            {
                pts.Add(getSP(item)); pts.Add(getSP(item, false));
            }
            return pts;
        }
        public SketchEntity getSE(SketchPoint sp1, SketchPoint sp2)
        {
            SketchEntity e1 = sp1.AttachedEntities[1], e2 = sp1.AttachedEntities[2],
                e3 = sp2.AttachedEntities[1], e4 = sp2.AttachedEntities[2];
            if (e1.Equals(e3) || e1.Equals(e4)) return e1;
            else if (e2.Equals(e4) || e2.Equals(e3)) return e2;
            return null;
        }
        public Vector2d getVec(SketchEntity se)
        {
            var sp1 = getSP(se); var sp2 = getSP(se, false);
            return sp1.Geometry.VectorTo(sp2.Geometry);
        }
        public SketchPoint getSP(SketchEntity se, bool fl = true)
        {
            SketchLine sl = se as SketchLine;
            if (sl != null) return fl ? sl.StartSketchPoint : sl.EndSketchPoint;
            SketchEllipticalArc a = se as SketchEllipticalArc;
            if (a != null) return fl ? a.StartSketchPoint : a.EndSketchPoint;
            SketchSpline spl = se as SketchSpline;
            if (spl != null) return fl ? spl.StartSketchPoint : spl.EndSketchPoint;
            return null;
        }
        public Point2d getMP(SketchEntity se)
        {
            var s = getSP(se); var e = getSP(se, false);
            return u.midPt(s.Geometry, e.Geometry);
        }
        public bool addGap(SketchPoint sp1, SketchPoint sp2, bool gap = true)
        {
            var re1 = sp1.ReferencedEntity; var re2 = sp2.ReferencedEntity;
            if (re1 == null || re2 == null) return true;
            Vertices vs1 = re1 as Vertices, vs2 = re2 as Vertices;
            Vertex v1 = vs1[1], v2 = vs2[1];
            var e1 = u.get<Edge>(v1.Edges, fi => fi.CurveType == CurveTypeEnum.kBSplineCurve);
            var e2 = u.get<Edge>(v2.Edges, fi => fi.CurveType == CurveTypeEnum.kBSplineCurve);
            if (e1 == null || e2 == null) return true;
            var p1 = ps.AddByProjectingEntity(getVertex(e1, v1)) as SketchPoint; p1.HoleCenter = false;
            var p2 = ps.AddByProjectingEntity(getVertex(e2, v2)) as SketchPoint; p2.HoleCenter = false;
            if (gap) ps.SketchLines.AddByTwoPoints(p1, p2);
            ps.SketchLines.AddByTwoPoints(sp1, p2);
            ps.SketchLines.AddByTwoPoints(sp2, p1);
            return false;
        }
        public SketchPoint getPt(SketchLine sl, List<SketchLine> ls)
        {
            SketchPoint sp = sl.StartSketchPoint, ep = sl.EndSketchPoint;
            var c1 = check(sp, ls);
            var c2 = check(ep, ls);
            return c1 ? sp : c2 ? ep : null;
        }
        public bool check(SketchPoint pt, IEnumerable<SketchLine> ls)
        {
            if (pt.Constraints.Count > 0) return false;
            if (pt.AttachedEntities.Count > 1) return false;
            return true;
        }
        public Vertex getVertex(Edge e, Vertex v)
        {
            return e.StartVertex.Equals(v) ? e.StopVertex : e.StartVertex;
        }
        public List<SketchPoint> getPts(SketchLine sl1, SketchLine sl2)
        {
            List<SketchPoint> pts = new List<SketchPoint>() { sl1.StartSketchPoint, sl1.EndSketchPoint,
            sl2.StartSketchPoint, sl2.EndSketchPoint};
            var v = sl1.StartSketchPoint.Geometry.VectorTo(sl1.EndSketchPoint.Geometry);
            return pts.OrderBy(el => u.dotProduct(v, el.Geometry)).ToList(); 
        }
        public List<Face> getEnt(SurfaceBody sb, PlanarSketch ps)
        {
            List<Face> fs = new List<Face>();
            foreach (var item in u.gets<Face>(sb.Faces, fi => fi.SurfaceType == SurfaceTypeEnum.kPlaneSurface ||
            fi.SurfaceType == SurfaceTypeEnum.kBSplineSurface))
            {
                if (item.SurfaceType == SurfaceTypeEnum.kBSplineSurface)
                {
                    fs.Add(item);
                    continue;
                }
                var pl = item.Geometry as Plane;
                var n = ps.PlanarEntityGeometry.Normal;
                if (pl.Normal.IsParallelTo(n) || pl.Normal.IsPerpendicularTo(n)) continue;
                fs.Add(item);
            }
            return fs;
        }
        public Edge getDir(List<Edge> es)
        {
            foreach (var item in es)
            {
                var e1 = getDir(item.StartVertex);
                if (e1 != null) return e1;
                e1 = getDir(item.StopVertex);
                if (e1 != null) return e1;
            }
            return null;
        }
        public Edge getDir(Vertex v)
        {
            Edge e = null; double min = 10000;
            bool spl = false;
            foreach (Edge item in v.Edges)
            {
                if (item.GeometryType == CurveTypeEnum.kBSplineCurve) { spl = true; continue; }
                Vector vec = item.StartVertex.Point.VectorTo(item.StopVertex.Point);
                double d = vec.Length;
                if (d < min)
                {
                    min = d; e = item;
                }
            }
            if (!spl) return null;
            return e;
        }
        public List<Edge> getEdges(List<Face> fs, Vector vec)
        {
            List<Edge> es = new List<Edge>();
            foreach (var item in fs)
            {
                foreach (Edge e in item.Edges)
                {
                    var v = e.StartVertex.Point.VectorTo(e.StopVertex.Point);
                    double n = u.dotProduct(vec, v);
                    //if (u.eq(n, 0)) continue;
                    if (!u.eq(n,0) && Math.Abs(n) < t*1.1) continue;
                    if (v.Length * 0.7 < Math.Abs(n)) continue;
                    //if (e.GeometryType == CurveTypeEnum.kLineCurve v.Length < t * 2) continue;
                    es.Add(e);
                }
            }
            return es;
        }
    }
}
