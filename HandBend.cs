//#define INV14

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Linq;
using InterfaceDll;
using InvDoc;
using Inventor;

namespace InvAddIn
{
    class HandBend
    {
        public Document doc;
        public PlanarSketch ps;
        int count = 0;
        double w, offset, thick, cut_l;
        MyForm F;
        MyXML xml;
        bool templ = false;
        string cut_l_s, count_s, w_s,
                offset_s, thick_s, plunge,
                template, t;
        bool noParam = false, co = false;

        public HandBend(Document doc)
        {
            this.doc = doc;
            SelectSet ss = I.getSS(doc);
            List<Edge> eds = new List<Edge>();
            foreach (Edge item in ss)
            {
                eds.Add(item);
            }
            F = new MyForm("HandBendInterface.xml", "Ручная гибка");
            //F.bnts[0].Click += fillet_Click;
            F.f.ShowDialog();
            init();
            var tr = I.beginTrans("bend", doc);
            hand_bend(eds);
            tr.End();
        }
        public void init()
        {
            template = F.cbs[7].Text;
            if (init(template))
            {
                templ = true;
                return;
            }

            cut_l_s = F.cbs[5].Text; count_s = F.cbs[2].Text; w_s = F.cbs[1].Text;
            offset_s = F.cbs[0].Text; thick_s = F.cbs[4].Text; plunge = F.cbs[6].Text;
            t = F.cbs[3].Text.ToLower()[0].ToString();

            set_val();
            
        }
        public bool set_val()
        {
            w = u.convToDouble(w_s);
            offset = u.convToDouble(offset_s);
            if (count_s != "")
                count = u.convToInt(count_s);
            thick = u.convToDouble(thick_s);

            if (cut_l_s != "")
            {
                cut_l = u.convToDouble(cut_l_s);
                count = 0;
                offset = cut_l;
            }

            offset *= 0.1;
            w *= 0.1;
            thick *= 0.1;

            if (t == "s")
            {
                w += thick;
                offset += thick * 0.5;
            }
            if (count == 0) co = true;
            if (count == 0) noParam = true;
            return true;
        }
        public bool init(string template)
        {
            if (template == null || template == "") return false;
            xml = new MyXML(template);
            return true;
        }
        public HandBend()
        {
            doc = I.aDoc();
            var ss = I.getSS(doc);
            var sl = ss[1] as SketchLine;
            if (sl == null) return;
            ps = sl.Parent as PlanarSketch;
            SketchArc sa = null;
            var l = add_half_slot(sl.StartSketchPoint, sl.EndSketchPoint, sl, 2, 10, ref sa);
            add_half_slot(sl.EndSketchPoint, sl.StartSketchPoint, sl, 2, 10, ref sa, l);
        }
        int get_count(SketchPoint sp1, SketchPoint sp2, double l, double w, ref double offset)
        {
            var vec = sp1.Geometry.VectorTo(sp2.Geometry);
            double l2 = vec.Length, l3 = l+w;
            int c = (int)(l2 / l3);
            double m = l2 - (c*l3 + w);
            if (m < 0)
            {
                c--;
                m = l2 - (c * l3 + w);
            }
            offset = m / 2;
            return c;
        }
        public bool set_vals(SketchPoint pt1, SketchPoint pt2)
        {
            var vec = pt1.Geometry.VectorTo(pt2.Geometry);
            var smcd = I.getSMCD(doc);
            var th = smcd.ActiveSheetMetalStyle.Thickness;

            foreach (var item in xml.elem.Elements())
            {
                var th_s = MyXML.getAtt(item, "thick");
                var mat = MyXML.getAtt(item, "mat");
                if (th == th_s)
                {
                    return set_vals(item, vec.Length);
                }
            }
            return false;
        }
        public bool set_vals(XElement el, double l)
        {
            l *= 10;
            foreach (var item in el.Elements())
            {
                var min = MyXML.convToDouble(MyXML.getAtt(item, "min"));
                var max = MyXML.convToDouble(MyXML.getAtt(item, "max"));
                if (l > min && l <= max)
                {
                    return set_val(item);
                }
            }
            return false;
        }
        public bool set_val(XElement el)
        {
            cut_l_s = MyXML.getAtt(el, "l"); 
            count_s = MyXML.getAtt(el, "count"); 
            w_s = MyXML.getAtt(el, "w");
            offset_s = MyXML.getAtt(el, "offset"); 
            thick_s = MyXML.getAtt(el, "t");
            t = MyXML.getAtt(el, "type").ToLower()[0].ToString();
            plunge = MyXML.getAtt(el, "plunge");

            set_val();
            return true;
        }
        void hand_bend(List<Edge> ss)
        {
            List<SketchLine> lines = new List<SketchLine>();
            SketchArc sa = null;
            if (doc.ActivatedObject is FlatPattern)
            {
                FlatPattern fp = doc.ActivatedObject as FlatPattern;
                ps = fp.Sketches.Add(fp.TopFace);
                foreach (Edge ed in ss)
                {
                    SketchPoint pt1 = ps.AddByProjectingEntity(ed.StartVertex) as SketchPoint,
                     pt2 = ps.AddByProjectingEntity(ed.StopVertex) as SketchPoint;
                    double lstart = offset;
                    double l = w;
                    if (co) count = get_count(pt1, pt2, offset, w, ref lstart);
                    if (templ)
                    {
                        bool cont = set_vals(pt1, pt2);
                        if (!cont) continue;
                        lstart = offset;
                        l = w;
                        if (co) count = get_count(pt1, pt2, offset, w, ref lstart);
                    }
                    Parameter pStart = null, p = null;
                    SketchLine sl = addLine(pt1, pt1.Geometry.VectorTo(pt2.Geometry), lstart, l, ref pStart, false);
                    if (noParam)
                    {
                        TwoPointDistanceDimConstraint dc = ps.DimensionConstraints.AddTwoPointDistance(sl.StartSketchPoint, sl.EndSketchPoint,
                            DimensionOrientationEnum.kAlignedDim, u.midPt(sl.StartSketchPoint.Geometry, sl.EndSketchPoint.Geometry, 0, 0.2));
                        dc.Parameter.Expression = (offset * 10).ToString();
                        pStart.Expression = ((lstart + w) * 10).ToString();
                    }
                    ps.GeometricConstraints.AddCoincident(pt2 as SketchEntity, sl as SketchEntity);
                    lines.Add(sl);
                    for (int i = 1; i < count; i++)
                    {
                        sl = addLine(sl, l, ref p, false);
                        lines.Add(sl);
                    }
                    if (!noParam)
                    {
                        TwoPointDistanceDimConstraint dc = ps.DimensionConstraints.AddTwoPointDistance(pt2, sl.EndSketchPoint,
                            DimensionOrientationEnum.kAlignedDim, u.midPt(sl.EndSketchPoint.Geometry, pt2.Geometry, 0, 2));
                        dc.Parameter.Expression = pStart.Name;
                    }
                    if (plunge.ToLower() == "yes")
                    {
                        var l1 = add_half_slot(pt1, pt2, sl, thick, pStart.ModelValue - w, ref sa);
                        add_half_slot(pt2, pt1, sl, thick, pStart.ModelValue - w, ref sa, l1);
                    }
                }
            }
            if (lines.Count > 0 && t != "l")
                add_rect(lines, thick, t);
            if (sa != null)
                ps.GeometricConstraints.AddEqualRadius(sa as SketchEntity, ps.SketchArcs[ps.SketchArcs.Count] as SketchEntity);
        }

        public SketchLine add_half_slot(SketchPoint sp1, SketchPoint sp2, SketchLine sl, double w, double l, ref SketchArc sa,
            SketchLine sl2 = null)
        {
            Point2d p1 = sp1.Geometry, p2 = sp2.Geometry;
            var vec = p1.VectorTo(p2); vec.Normalize();
            vec.ScaleBy(l);
            var norm = u.normal(p1, p2); norm.Normalize();
            norm.ScaleBy(w);
            p1.TranslateBy(norm);
            norm.ScaleBy(-2);
            Point2d p3 = I.CP2d(p1, norm);
            SketchPoint mp = ps.SketchPoints.Add(I.CP2d(), false);

            var l1 = ps.SketchLines.AddByTwoPoints(p1, p3);
            ps.GeometricConstraints.AddMidpoint(mp, l1);
            ps.GeometricConstraints.AddPerpendicular(l1 as SketchEntity, sl as SketchEntity);
            ps.GeometricConstraints.AddCoincident(mp as SketchEntity, sp1 as SketchEntity);

            Point2d p4 = I.CP2d(p1, vec), p5 = I.CP2d(p3, vec);
            var l2 = ps.SketchLines.AddByTwoPoints(l1.StartSketchPoint, p4);
            var l3 = ps.SketchLines.AddByTwoPoints(l1.EndSketchPoint, p5);
            if (sl2 != null) ps.GeometricConstraints.AddEqualLength(sl2, l2);
            else
                ps.DimensionConstraints.AddTwoPointDistance(l2.StartSketchPoint, l2.EndSketchPoint,
                            DimensionOrientationEnum.kAlignedDim, u.midPt(l2.StartSketchPoint.Geometry, l2.EndSketchPoint.Geometry, 0, 0.2));

            ps.GeometricConstraints.AddPerpendicular(l1 as SketchEntity, l2 as SketchEntity);
            ps.GeometricConstraints.AddPerpendicular(l1 as SketchEntity, l3 as SketchEntity);

            var mp1 = u.midPt(l2.EndSketchPoint.Geometry, l3.EndSketchPoint.Geometry);

            var a1 = ps.SketchArcs.AddByCenterStartEndPoint(mp1, l2.EndSketchPoint, l3.EndSketchPoint, false);
            ps.GeometricConstraints.AddTangent(a1 as SketchEntity, l2 as SketchEntity);
            ps.GeometricConstraints.AddTangent(a1 as SketchEntity, l3 as SketchEntity);

            if (sa != null) ps.GeometricConstraints.AddEqualRadius(sa as SketchEntity, a1 as SketchEntity);
            sa = a1;
            return l2;
        }

        public void concid_line(SketchEntitiesEnumerator lines, SketchLine sl, int num1, bool start)
        {
            var ps = sl.Parent;
            var sp = ps.SketchPoints.Add(I.CP2d());
            sp.HoleCenter = false;
            var mp = ps.GeometricConstraints.AddMidpoint(sp, (SketchLine)lines[num1]);
            if (start)
                ps.GeometricConstraints.AddCoincident(mp.Point as SketchEntity, sl.StartSketchPoint as SketchEntity);
            else
                ps.GeometricConstraints.AddCoincident(mp.Point as SketchEntity, sl.EndSketchPoint as SketchEntity);
        }

        public void concid_arc(IEnumerable<SketchArc> arcs, SketchLine sl, bool start)
        {
            var ps = sl.Parent;
            if (start)
            {
                arcs = arcs.OrderBy(e => e.CenterSketchPoint.Geometry.DistanceTo(sl.StartSketchPoint.Geometry));
                ps.GeometricConstraints.AddCoincident(sl.StartSketchPoint as SketchEntity, arcs.First().CenterSketchPoint as SketchEntity);
            }
            //ps.GeometricConstraints.AddMidPointToArc(sl.StartSketchPoint, arcs.ElementAt(num1));
            else
            {
                arcs = arcs.OrderBy(e => e.CenterSketchPoint.Geometry.DistanceTo(sl.EndSketchPoint.Geometry));
                ps.GeometricConstraints.AddCoincident(sl.EndSketchPoint as SketchEntity, arcs.First().CenterSketchPoint as SketchEntity);
            }
            //ps.GeometricConstraints.AddMidPointToArc(sl.EndSketchPoint, arcs.ElementAt(num1));
        }

        public void add_rects()
        {
            var doc = I.aDoc();
            var ss = I.getSS(doc);
            var ps = ss[1] as PlanarSketch;
            if (ps == null) return;
            var lines = u.gets<SketchLine>(ps.SketchLines, f => f.Construction == false);
            add_rect(lines, 2, "s");
        }

        public void add_rect(IEnumerable<SketchLine> lines, double w, string t = "r")
        {
            //w = w * 0.1;
            SketchLine sl = lines.First();
            SketchEntity se = null;
            foreach (SketchLine line in lines)
            {
                if (line.Equals(sl))
                {
                    if (t == "r")
                        sl = add_rect(line, w);
                    else if (t == "s")
                        se = add_slot(line, w);
                }
                else
                {
                    if (t == "r")
                        add_rect(line, sl);
                    else
                        add_slot(line, se);
                }
            }
        }

        public SketchEntity add_slot(SketchLine sl, object p)
        {
            double w = 0;
            if (p is double)
                w = ((double)p);
            else if (p is SketchArc)
            {
                w = ((SketchArc)p).Radius*2;
            }

            sl.Construction = true;
            var ps = sl.Parent;
            Point2d p1 = I.CP2d(sl.StartSketchPoint), p2 = I.CP2d(sl.EndSketchPoint);
            var vec = p1.VectorTo(p2); vec.Normalize();

            var lines = ps.AddStraightSlotByOverall(p1, p2, w);
            var arcs = u.gets<SketchArc>(lines, f => f is SketchArc);

            concid_arc(arcs, sl, true);
            concid_arc(arcs, sl, false);

            var line = arcs.ElementAt(0) as SketchEntity;
            var mp = I.CP2d(sl.StartSketchPoint.Geometry, vec);

            if (p is double)
            {
                var d = ps.DimensionConstraints.AddRadius(line, mp);
                d.Parameter.Expression = (w * 5).ToString();
            }
            else if (p is SketchEntity)
            {
                ps.GeometricConstraints.AddEqualRadius(line, p as SketchEntity);
            }
            return line;
        }

        public SketchLine add_rect(SketchLine sl, object p)
        {
            double w = 0;
            if (p is double)
                w = ((double)p);
            else if (p is SketchLine)
            {
                w = ((SketchLine)p).Length;
            }

            sl.Construction = true;
            var ps = sl.Parent;
            Point2d p1 = I.CP2d(sl.StartSketchPoint), p2 = I.CP2d(sl.EndSketchPoint), p3 = I.CP2d(sl.EndSketchPoint);
            var vec = p1.VectorTo(p2); vec.Normalize();
            var norm = u.normal(p1, p2); norm.Normalize();
            norm.ScaleBy(w / 2);
            p1.TranslateBy(norm); p2.TranslateBy(norm);
            norm.ScaleBy(-1);
            p3.TranslateBy(norm);

            var lines = ps.SketchLines.AddAsThreePointRectangle(p1, p2, p3);

            concid_line(lines, sl, 2, true);
            concid_line(lines, sl, 4, false);
            var line = lines[2] as SketchLine;
            var mp = I.CP2d(sl.StartSketchPoint.Geometry, vec);

            if (p is double)
            {
                var d = ps.DimensionConstraints.AddTwoPointDistance(line.StartSketchPoint, line.EndSketchPoint, DimensionOrientationEnum.kAlignedDim, mp);
                d.Parameter.Expression = (w * 10).ToString();
            }
            else if (p is SketchLine)
            {
                ps.GeometricConstraints.AddEqualLength(line, p as SketchLine);
            }
            return line;
        }

        public SketchLine addLine(SketchPoint sp, Vector2d vec, double l1, double l2, ref Parameter param, bool noParam)
        {
            PlanarSketch ps = sp.Parent as PlanarSketch;
            Point2d start, end;
            vec.Normalize();
            vec.ScaleBy(l1);
            start = sp.Geometry;
            start.TranslateBy(vec);
            end = start.Copy();
            vec.Normalize(); vec.ScaleBy(l2);
            end.TranslateBy(vec);
            SketchLine sl = ps.SketchLines.AddByTwoPoints(start, end);
            ps.GeometricConstraints.AddCoincident((SketchEntity)sp, (SketchEntity)sl);
            if (!noParam)
                param = ps.DimensionConstraints.AddTwoPointDistance(sp, sl.StartSketchPoint, DimensionOrientationEnum.kAlignedDim, u.midPt(sp.Geometry, sl.StartSketchPoint.Geometry, 0, 2)).Parameter;
            return sl;
        }

        public SketchLine addLine(SketchLine startsl, double l1, ref Parameter param, bool noParam)
        {
            PlanarSketch ps = startsl.Parent as PlanarSketch;
            Point2d start, end;
            Vector2d vec = startsl.StartSketchPoint.Geometry.VectorTo(startsl.EndSketchPoint.Geometry);
            vec.Normalize();
            vec.ScaleBy(l1);
            start = startsl.EndSketchPoint.Geometry;
            start.TranslateBy(vec);
            end = start.Copy();
            vec.Normalize(); vec.ScaleBy(l1);
            end.TranslateBy(vec);
            SketchLine sl = ps.SketchLines.AddByTwoPoints(start, end);

            if (!noParam)
            {
                TwoPointDistanceDimConstraint constr = ps.DimensionConstraints.AddTwoPointDistance(startsl.EndSketchPoint, sl.StartSketchPoint,
                    DimensionOrientationEnum.kAlignedDim, u.midPt(startsl.EndSketchPoint.Geometry, sl.StartSketchPoint.Geometry, 0, 2));
                if (param == null)
                    param = constr.Parameter;
                else
                    constr.Parameter.Expression = param.Name;
            }
            ps.GeometricConstraints.AddEqualLength(startsl, sl);
            ps.GeometricConstraints.AddCollinear((SketchEntity)startsl, (SketchEntity)sl);
            return sl;
        }
    }

    class MultiCut
    {
        public Document doc;
        public PlanarSketch ps;
        public FlatPattern fp;
        public SheetMetalComponentDefinition smcd;
        public SheetMetalFeatures smf;
        public Face f;
        public Edge ed;
        SurfaceBody sb;
        int count;
        Vector v; 
        Vector2d vx, vy, offset;

        public MultiCut(Document doc)
        {
            this.doc = doc;
            init();
            //fp = smcd.FlatPattern;
            //if (fp == null) return;
            //f = fp.TopFace;
            //ps = fp.Sketches.Add(f, false);
            //set_offset(true);
            this.doc = add_doc();
            init();
            count = 3;
            //for (int i = 0; i < count; i++)
            //{
            //    add_entities(i);
            //}
            //extr();
            add_pattern();
        }

        public Document add_doc()
        {
            string path = file.p(doc.FullDocumentName);
            string templPath = I.dPr().TemplatesPath;
            string name = file.name(doc.FullDocumentName);
            Document d = I.newDoc(path + name + "multi.ipt", templPath + "Листовой.ipt", true);
            I.addSB(d, doc.FullDocumentName, smcd.SurfaceBodies[1]);
            d.Save();
            return d;
        }
        public void init()
        {
            smcd = I.getSMCD(doc);
            sb = smcd.SurfaceBodies[1];
            if (smcd == null) return;
            smf = smcd.Features as SheetMetalFeatures;
        }
        public void find_dir()
        {
            if (smcd.Bends.Count == 0) return;
            var b = smcd.Bends[1];
            f = b.FrontFaces[1];
            ed = u.gets<Edge>(f.Edges, e => e.CurveType == CurveTypeEnum.kLineCurve).First();
        }
        public void add_pattern()
        {
            find_dir();
            var v = ed.StartVertex.Point.VectorTo(ed.StopVertex.Point);
            var col = I.COC();
            col.Add(sb);
            var b = sb.RangeBox;
            Point p1 = u.ProjectPointOntoVector(b.MinPoint, v), p2 = u.ProjectPointOntoVector(b.MaxPoint, v);
            double d = p1.DistanceTo(p2);
#if INV14
            smf.RectangularPatternFeatures.Add(col, ed, true, count, d);
#else
            var def = smf.RectangularPatternFeatures.CreateDefinition(col, ed, true, count, d);
            var pat = smf.RectangularPatternFeatures.AddByDefinition(def);
#endif

        }

        public void set_offset(bool x = true)
        {
            var b = f.Evaluator.RangeBox;
            v = b.MinPoint.VectorTo(b.MaxPoint);
            vx = I.CV2d(v.X, 0);
            vy = I.CV2d(0, v.Y);
            offset = x ? vx : vy;
        }
        public void add_entities(int num)
        {
            var v = offset.Copy();
            v.ScaleBy(num + 1);
            foreach (Edge ed in f.Edges)
            {
                if (ed.CurveType == CurveTypeEnum.kLineSegmentCurve || ed.CurveType == CurveTypeEnum.kLineCurve)
                    add_line(ed, v);
                else if (ed.GeometryType == CurveTypeEnum.kCircleCurve)
                    add_circle(ed, v);
                else if (ed.GeometryType == CurveTypeEnum.kCircularArcCurve)
                    add_arc(ed, v);
            }
        }
        public bool filter_line(Point2d p1, Point2d p2)
        {
            foreach (SketchLine item in ps.SketchLines)
            {
                if (u.eq(item.StartSketchPoint.Geometry, p1) &&
                    u.eq(item.EndSketchPoint.Geometry, p2))
                    return true;
            }
            return false;
        }
        public void extr()
        {
            u.filter_line(ps);
            var merged = new MergePoints(ps);
            var pr = ps.Profiles.AddForSolid(true);
            var def = fp.Features.ExtrudeFeatures.CreateExtrudeDefinition(pr, PartFeatureOperationEnum.kJoinOperation);
            def.SetDistanceExtent(smcd.Thickness, PartFeatureExtentDirectionEnum.kNegativeExtentDirection);
            fp.Features.ExtrudeFeatures.Add(def);
        }
        public void add_line(Edge ed, Vector2d v)
        {
            Point s = ed.StartVertex.Point, e = ed.StopVertex.Point;
            Point2d p1 = ps.ModelToSketchSpace(s), p2 = ps.ModelToSketchSpace(e);
            p1.TranslateBy(v); p2.TranslateBy(v);
            ps.SketchLines.AddByTwoPoints(p1, p2);
        }
        public void add_circle(Edge ed, Vector2d v)
        {
            Circle c = ed.Geometry as Circle;
            Point s = c.Center;
            Point2d p1 = ps.ModelToSketchSpace(s);
            p1.TranslateBy(v);
            ps.SketchCircles.AddByCenterRadius(p1, c.Radius);
        }
        public void add_arc(Edge ed, Vector2d v)
        {
            Arc3d a = ed.Geometry as Arc3d;
            var vec = I.CV(v.X, v.Y, 0);
            Point c = a.Center, s = a.StartPoint, e = a.EndPoint;
            Point mp = get_mp(ed);
            c.TranslateBy(vec); e.TranslateBy(vec); s.TranslateBy(vec); mp.TranslateBy(vec);
            Point2d p1 = ps.ModelToSketchSpace(c), p2 = ps.ModelToSketchSpace(s), p3 = ps.ModelToSketchSpace(e), p4 = ps.ModelToSketchSpace(mp);
            //p1.TranslateBy(v); p2.TranslateBy(v); p3.TranslateBy(v);
            
            //var ar = ps.SketchArcs.AddByCenterStartEndPoint(p1, p2, p3);
            //var ar = ps.SketchArcs.AddByCenterStartSweepAngle(p1, a.Radius, a.StartAngle, a.SweepAngle);
            var ar = ps.SketchArcs.AddByThreePoints(p2, p4, p3);
            double l1 = u.getLenght(ed), l2 = u.getLength(ar);
            //if (!u.eq(l1, l2))
            //{
            //    ar.Delete();
            //    ar = ps.SketchArcs.AddByCenterStartEndPoint(p1, p2, p3, false);
            //}
            //else if (!u.eq(ar.Geometry3d.StartPoint, s))
            //{
                
            //    if (!u.eq(ar.StartAngle, a.StartAngle))
            //    {
            //        ar.Delete();
            //        ps.SketchArcs.AddByCenterStartEndPoint(p1, p2, p3, false);
            //    }
            //    else
            //    {
            //        ar.Delete();
            //        ps.SketchArcs.AddByCenterStartEndPoint(p1, p3, p2, false);
            //    }
            //    //ar = ps.SketchArcs.AddByCenterStartEndPoint(p1, p3, p2, false);
            //}
            //else if (!u.eq(ar.StartAngle, a.StartAngle) || !u.eq(ar.StartAngle, a.StartAngle, 0.1))
            //{
            //    ar.Delete();
            //    ar = ps.SketchArcs.AddByCenterStartEndPoint(p1, p2, p3, false);
            //}
        }
        public Point get_mp(Edge ed)
        {
            var ev = ed.Evaluator;
            double min, max, mid;
            ev.GetParamExtents(out min, out max);
            mid = (max + min) / 2;
            double[] pt = new double[3];
            double[] param = new double[] { mid };
            ev.GetPointAtParam(ref param, ref pt);
            return I.CP(pt[0], pt[1], pt[2]);
        }
    }
    class VertexProjection
    {
        public Vertex start;
        public Vertex v;
        public Vector vec;
        public Point pr;
        public Edge ed;
        public double d;
        public VertexProjection(Vertex v, Vector vec, Vertex s, Edge ed)
        {
            start = s;
            this.ed = ed;
            CutFlatPattern.edges.Add(ed);
            this.v = v;
            pr = u.ProjectPointOntoVector(v.Point, vec);
            this.vec = start.Point.VectorTo(pr);
            d = start.Point.DistanceTo(pr);
            if (vec.DotProduct(this.vec) < 0) 
                d *= -1;
        }
    }
    class SortVertex
    {
        public List<VertexProjection> lst = new List<VertexProjection>();
        public Vector v;
        Point origin;
        Edge s;
        int i, j;
        SheetMetalComponentDefinition smcd;
        public SortVertex(Edge s, SheetMetalComponentDefinition def, Vertex start = null, int j = 0)
        {
            this.s = s;
            this.j = j;
            smcd = def;
            v = s.StartVertex.Point.VectorTo(s.StopVertex.Point);
            origin = s.StartVertex.Point;
            if (start != null)
            {
                double d1 = start.Point.DistanceTo(s.StartVertex.Point);
                double d2 = start.Point.DistanceTo(s.StopVertex.Point);
                if (d1 > d2)
                {
                    origin = s.StopVertex.Point;
                    v = s.StopVertex.Point.VectorTo(s.StartVertex.Point);
                }
                
            }
            //u.clientTxt(smcd as PartComponentDefinition, u.midPt(s.StartVertex.Point, s.StopVertex.Point), $"{j}-{i}-s");
        }
        public void setVector(Vector v)
        {
            this.v = v;
        }
        public void addVertex(Vertex vert, Edge ed, ref HashSet<Vertex> filter)
        {
            if (!check_vertex(filter, vert))
            {
                lst.Add(new VertexProjection(vert, v, s.StartVertex, ed));
                //u.clientTxt(smcd as PartComponentDefinition, vert.Point, $"{j}-{i}-s");
                filter.Add(vert);
            }
        }
        public bool check_vertex(HashSet<Vertex> filter, Vertex vert)
        {
            foreach (var item in filter)
            {
                if (u.eq(vert.Point, item.Point)) 
                    return true;
            }
            return false;
        }
        public void merge(SortVertex other)
        {
            lst.AddRange(other.lst);
        }
        public bool add(Edge ed, ref HashSet<Vertex> filter)
        {

            if (FPOp.check_collinear(origin, ed.StartVertex.Point, v))
            {
                addVertex(ed.StartVertex, ed, ref filter);
            }
            if (FPOp.check_collinear(origin, ed.StopVertex.Point, v))
            {
                addVertex(ed.StopVertex, ed, ref filter);
            }

                //u.clientTxt(smcd as PartComponentDefinition, u.midPt(ed.StartVertex.Point, ed.StopVertex.Point), $"{j}-{i}-s");
                //return true;
            return false;
        }
        public void add(IEnumerable<Edge> eds, ref HashSet<Vertex> filter)
        {
            foreach (var item in eds)
            {
                if (!v.IsParallelTo(item.StartVertex.Point.VectorTo(item.StopVertex.Point))) continue;
                add(item, ref filter);
                i++;
            }
        }
        public void sort()
        {
            lst = lst.OrderBy(e => e.d).ToList();
        }
    }
    class MergePoint
    {
        public SketchPoint sp;
        public double x, y;
        public bool merged = false;
        public MergePoint(SketchPoint sp)
        {
            this.sp = sp;
            x = sp.Geometry.X;
            y = sp.Geometry.Y;
        }
        public bool check(MergePoint mp)
        {
            return u.eq(x, mp.x) && u.eq(y, mp.y);
        }
        public void merge(MergePoint mp)
        {
            sp.Merge(mp.sp);
            merged = true;
        }
    }
    class MergePoints
    {
        List<MergePoint> pts = new List<MergePoint>();
        public MergePoints(PlanarSketch ps)
        {
            foreach (SketchPoint item in ps.SketchPoints)
            {
                pts.Add(new MergePoint(item));
            }

            merge();
        }
        public void merge()
        {
            foreach (var item in pts)
            {
                if (item.merged) continue;
                foreach (var pt in pts)
                {
                    if (pt.merged) continue;
                    if (item.Equals(pt)) continue;
                    if (item.check(pt))
                    {
                        item.merge(pt);
                    }
                }
            }
        }
    }
    class CutFlatPattern
    {
        Document doc;
        SheetMetalComponentDefinition smcd;
        SheetMetalFeatures smf;
        FlatPattern fp;
        PlanarSketch ps;
        SurfaceBody sb;
        UnfoldFeature uf;
        RefoldFeature rf;
        SurfaceBody sborigin;
        ObjectCollection bends;
        public static HashSet<Edge> edges = new HashSet<Edge>();
        string name = "Mark", sbname;
        Face f;
        double r;
        public CutFlatPattern(Document doc)
        {
            this.doc = doc;
            smcd = I.getSMCD(this.doc);
            if (smcd.SurfaceBodies.Count > 1) return;
            smf = smcd.Features as SheetMetalFeatures;
            if (smcd == null) return;
            //fp = I.getFP(this.doc);
            //if (fp == null) return;
            r = 0.1;
            if (exist()) return;
            if (smcd.Bends.Count == 0) return;
            sborigin = smcd.SurfaceBodies[1];
            sbname = sborigin.Name;
            addSketch();
        }
        public CutFlatPattern(string name)
        {
            this.doc = I.aDoc();
            smcd = I.getSMCD(this.doc);
            if (smcd.SurfaceBodies.Count == 1) return;
            foreach (SurfaceBody item in smcd.SurfaceBodies)
            {
                if (name == item.Name)
                {
                    sborigin = item;
                    sbname = sborigin.Name;
                    this.name += sbname;
                    bends = I.COC();
                    break;
                }
            }
            smf = smcd.Features as SheetMetalFeatures;
            if (smcd == null) return;
            //fp = I.getFP(this.doc);
            //if (fp == null) return;
            r = 0.1;
            if (exist()) return;
            if (smcd.Bends.Count == 0) return;
            //getBends();
            addSketch();
        }
        public bool exist()
        {
            foreach (CutFeature item in smf.CutFeatures)
            {
                if (item.Name == name) return true;
            }
            return false;
        }
        public void findFace()
        {
            var fs = u.gets<Face>(sborigin.Faces, e => e.SurfaceType == SurfaceTypeEnum.kPlaneSurface).
                OrderBy(e => e.Evaluator.Area);
            f = fs.Last();
        }
        public void addSketch()
        {
            findFace();
            uf = smf.UnfoldFeatures.Add(f);
            rf = smf.RefoldFeatures.Add(uf.StationaryFace);
            uf.SetEndOfPart(false);
            ps = smcd.Sketches.Add(uf.StationaryFace);
            projectBend();
            rf.SetEndOfPart(false);
        }
        public void getBends()
        {
            foreach (Bend bend in smcd.Bends)
            {
                var fa = bend.FrontFaces[1];
                if (fa.SurfaceBody.Name != sbname) bends.Add(bend);
            }
        }
        public void projectBend()
        {
            HashSet<Vertex> filter = new HashSet<Vertex>();
            var bends = new HashSet<Edge>();
            foreach (Bend bend in smcd.Bends)
            {
                var fa = bend.FrontFaces[1];
                if (fa.SurfaceBody.Name != sbname) continue;
                foreach (Face fs in bend.FrontFaces)
                {
                    var r = u.gets<Edge>(fs.Edges, f => f.CurveType == CurveTypeEnum.kLineCurve && f.TangentiallyConnectedEdges.Count == 0);
                    foreach (var item in r)
                    {
                        bends.Add(item);
                    }
                }

                //fs = bend.BackFaces[1];
                //r = u.gets<Edge>(fs.Edges, f => f.CurveType == CurveTypeEnum.kLineCurve);
                //bends.AddRange(r);
            }
            int i = 1;
            foreach (Bend b in smcd.Bends)
            {
                var fs = b.FrontFaces[1];
                if (fs.SurfaceBody.Name != sbname) continue;
                var eds = u.gets<Edge>(fs.Edges, f => f.TangentiallyConnectedEdges.Count == 0 && !edges.Contains(f));
                if (eds.Count() == 0) continue;
                eds = eds.OrderByDescending(e => u.getLenght(e));
                var e1 = eds.ElementAt(0);
                var e2 = eds.ElementAt(1);
                var sv1 = new SortVertex(e1, smcd);
                sv1.addVertex(e1.StartVertex, e1,  ref filter); sv1.addVertex(e1.StopVertex, e1, ref filter);
                sv1.add(bends, ref filter);
                
                var sv2 = new SortVertex(e2, smcd, e1.StartVertex, 1);
                sv2.addVertex(e2.StartVertex, e2, ref filter); sv2.addVertex(e2.StopVertex, e2, ref filter);
                sv2.add(bends, ref filter);
                //sv1.merge(sv2);
                sv1.sort();
                sv2.sort();
                if (sv1.lst.Count > 0 && sv2.lst.Count > 0)
                {
                    if (!addCircle(sv1.lst[0].v, sv2.lst[0].v))
                    {
                        addCircle(sv1.lst[0].v, sv2.lst[sv2.lst.Count - 1].v);
                    }
                    if (!addCircle(sv1.lst[sv1.lst.Count-1].v, sv2.lst[sv2.lst.Count-1].v))
                    {
                        addCircle(sv1.lst[sv1.lst.Count - 1].v, sv2.lst[0].v);
                    }
                }
                i++;
                filter.Clear();
                //addCircle(eds.ElementAt(0), eds.ElementAt(1));
            }
            edges.Clear();
            cut();
        }
        public void addCircle(Edge ed1, Edge ed2)
        {
            SketchPoint p1 = ps.AddByProjectingEntity(ed1.StartVertex) as SketchPoint;
            SketchPoint p2 = ps.AddByProjectingEntity(ed2.StopVertex) as SketchPoint;
            SketchPoint p3 = ps.AddByProjectingEntity(ed1.StopVertex) as SketchPoint;
            SketchPoint p4 = ps.AddByProjectingEntity(ed2.StartVertex) as SketchPoint;
            p1.HoleCenter = false; p2.HoleCenter = false; p3.HoleCenter = false; p4.HoleCenter = false;
            addCircle(p1, p4);
            addCircle(p2, p3);
        }
        public bool addCircle(Vertex v1, Vertex v2)
        {

            SketchPoint p1 = addSketchPoint(v1);
            SketchPoint p2 = addSketchPoint(v2);
            p1.HoleCenter = false; p2.HoleCenter = false;
            addCircle(p1, p2);
            return true;
        }

        public SketchPoint addSketchPoint(Vertex v)
        {
            foreach (SketchPoint item in ps.SketchPoints)
            {
                if (u.eq(item.Geometry, ps.ModelToSketchSpace(v.Point)))
                    return item;
            }
            return ps.AddByProjectingEntity(v) as SketchPoint;
        }

        public void addCircle(SketchPoint p1, SketchPoint p2)
        {
            var sl = ps.SketchLines.AddByTwoPoints(p1, p2);
            sl.Construction = true;
            var mp = ps.SketchPoints.Add(I.CP2d());
            ps.GeometricConstraints.AddMidpoint(mp, sl);
            var c1 = ps.SketchCircles.AddByCenterRadius(mp, r);
            ps.GeometricConstraints.AddCoincident(mp as SketchEntity, c1.CenterSketchPoint as SketchEntity);
            if (ps.SketchCircles.Count > 1) ps.GeometricConstraints.AddEqualRadius(c1 as SketchEntity, ps.SketchCircles[1] as SketchEntity);
        }
        public void cut()
        {
            var pr = ps.Profiles.AddForSolid(false);
            var def = smf.CutFeatures.CreateCutDefinition(pr);
            var cf = smf.CutFeatures.Add(def);
            cf.Name = name;
        }
    }
}
