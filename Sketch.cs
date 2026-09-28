using System;
using System.Collections.Generic;
using System.Windows.Forms;
using System.Linq;
using System.Drawing;
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
    internal class SketchBtn : Button
    {
        public static SketchOp so;

        public SketchBtn(Icon standardIcon, Icon largeIcon, string displayName = "Центр массива", string internalName = "ArrayCenter", 
            string clientId = "{3FF41256-8915-4D38-A76C-01C80A3A709A}",
            string description = "Начало массива", string tooltip = "Начало массива")
            : base(displayName, internalName, clientId, description, tooltip, standardIcon, largeIcon)
        {
        }
        protected override void ButtonDefinition_OnExecute(NameValueMap context)
        {
            if (Macros.StandardAddInServer.m_inventorApplication.ActiveDocument.DocumentType == DocumentTypeEnum.kPartDocumentObject)
            {
                so = new SketchOp(Macros.StandardAddInServer.m_inventorApplication.ActiveDocument as PartDocument);
            }
        }
    }

    internal class ComboBoxBtn : InvComboBox
    {

        public ComboBoxBtn(Icon standardIcon, Icon largeIcon, string displayName = "", string internalName = "Autodesk:Macros:CB", CommandTypesEnum commandType = CommandTypesEnum.kShapeEditCmdType, string clientId = "{5B47B9A0-EA78-4FD6-BFC1-D475AA510E05}",
            string description = "Данные для массива", string tooltip = "Данные для массива", ButtonDisplayEnum buttonDisplayType = ButtonDisplayEnum.kDisplayTextInLearningMode)
            : base(displayName, internalName, commandType, 200, clientId, standardIcon, largeIcon, description, tooltip)
        {
        }
        protected override void ComboBoxDefinition_OnSelect(NameValueMap context)
        {
            Macros.StandardAddInServer.data.getData();
        }
    }

    public class DataToArray
    {
        public Parameter parX, parY;
        PartDocument doc; ComboBoxDefinition def;
        public DataToArray(PartDocument doc, ComboBoxDefinition cbDef)
        {
            this.doc = doc;
            def = cbDef;
            addData();
        }
        public void addData()
        {
            def.Clear();
            foreach (CustomParameterGroup item in doc.ComponentDefinition.Parameters.CustomParameterGroups)
            {
                def.AddItem(item.DisplayName);
            }
            if (def.ListCount == 0)
            {
                foreach (DerivedParameter item in doc.ComponentDefinition.Parameters.DerivedParameterTables[1].DerivedParameters)
                {
                    if (item.Name.ToLower().EndsWith("смещение"))
                    def.AddItem(item.Name.Remove(item.Name.IndexOf("Смещение"))); 
                }
            }
        }
        public void getData()
        {
            if (doc.ComponentDefinition.Parameters.CustomParameterGroups.Count != 0)
            {
                CustomParameterGroup group = doc.ComponentDefinition.Parameters.CustomParameterGroups[def.Text];
                parX = null; parY = null;
                foreach (Parameter item in group)
                {
                    if (item.Name.EndsWith("Смещение")) parX = item;
                    if (item.Name.EndsWith("СмещениеY")) parY = item;
                }
            }
            if (doc.ComponentDefinition.Parameters.DerivedParameterTables.Count != 0)
            {
                foreach (DerivedParameter item in doc.ComponentDefinition.Parameters.DerivedParameterTables[1].DerivedParameters)
                {
                    if (item.Name == def.Text + "Смещение") parX = item as Parameter;
                    if (item.Name == def.Text + "СмещениеY") parY = item as Parameter;
                }
            }
        }
    }

    public class SketchOp
    {
        PartDocument doc;
        PlanarSketch ps;
        CommandManager cmd;
        SketchLine sl, sl1, sl2, sl3;
        Inventor.Point2d pt;
        SketchPoint sp;
        Vector2d vec;
        public SketchOp(PartDocument doc)
        {
            this.doc = doc;
            if (doc.ActivatedObject is PlanarSketch)
            {
                ps = doc.ActivatedObject as PlanarSketch;
                sl1 = addLine();
                sl2 = findLine(sl1);
                //sl2.Construction = true;
                SketchLine tmp = findLine(sl2);
//                 tmp.Construction = true;
//                 tmp = findLine(tmp);
//                 tmp.Construction = true;
                //addMidConsraint(sl);
                sl = addPerpendicular(sl1, sl2);
                
                //addMidConsraint(sl2);
                sl3 = addPerpendicular(sl2, sl1, sl);
                //ps.GeometricConstraints.AddCoincident(sl.EndSketchPoint as SketchEntity, sl3.EndSketchPoint as SketchEntity);

                //sp = ps.SketchPoints.Add(sl.Geometry.MidPoint);
                sp = addPoint(sl1, sl2);
                addDimConstr(sp, sl, Macros.StandardAddInServer.data.parX);
                addDimConstr(sp, sl3, Macros.StandardAddInServer.data.parY);
                addAttr(doc as Document, sl1, sp, tmp.Construction);
                //doc.Update();
            }
        }

        public void addAttr(Document doc, SketchLine sl, SketchPoint sp, bool contstr)
        {
           VariableData vd = new VariableData(doc as Document);
           UnitVector vec = sl.Geometry3d.Direction;
           AttributeSet attSet = vd.AttribSetAdd(sp, "dir");
           attSet.Add("X", ValueTypeEnum.kDoubleType, vec.X);
           attSet.Add("Y", ValueTypeEnum.kDoubleType, vec.Y);
           attSet.Add("Z", ValueTypeEnum.kDoubleType, vec.Z);
            if (contstr)
            {
                string val = ut.getRefkey<SketchLine>(doc, sl, 0);
                vd.AttribAdd<SketchPoint, string>(sp, "SL", val, ValueTypeEnum.kStringType);
            }
        }
        public SketchLine addLine()
        {
            cmd = (ps.Application as Inventor.Application).CommandManager;
            if (doc.SelectSet.Count != 0 && doc.SelectSet[1] is SketchLine) return doc.SelectSet[1] as SketchLine;
            object ob = cmd.Pick(SelectionFilterEnum.kSketchCurveFilter, "Выберите грань для главного направления");
            SketchEntity se = /*ps.AddByProjectingEntity*/ob as SketchEntity;
            if (se is SketchLine) return se as SketchLine;
            return null;
        }
        public SketchPoint addPoint(SketchLine sl1, SketchLine sl2)
        {
            Vector2d v1 = addVec(sp, sl1), v2 = addVec(sp, sl2);
            pt = sp.Geometry.Copy();
            v1.AddVector(v2);
            v1.ScaleBy(0.4);
            pt.TranslateBy(v1);
            return ps.SketchPoints.Add(pt);
        }
        public void addMidConsraint(SketchLine line)
        {
            pt = line.Geometry.MidPoint; 
            SketchPoint s = ps.SketchPoints.Add(pt,false);
            ps.GeometricConstraints.AddMidpoint(s,sl);
        }
        public void addMidConsraint(SketchLine sl1, SketchLine sl2)
        {
            ps.GeometricConstraints.AddMidpoint(sl2.StartSketchPoint, sl1);
        }
        public Vector2d addVec(SketchPoint pt, SketchLine dir)
        {
            Vector2d v;
            if (ut.eq(sp.Geometry, dir.StartSketchPoint.Geometry))
                v = dir.StartSketchPoint.Geometry.VectorTo(dir.EndSketchPoint.Geometry);
            else v = dir.EndSketchPoint.Geometry.VectorTo(dir.StartSketchPoint.Geometry);
            return v;
        }
        public SketchLine addPerpendicular(SketchLine line, SketchLine dir, SketchLine sl = null)
        {
            vec = addVec(sp, dir);
            vec.ScaleBy(0.5);
            pt = line.Geometry.MidPoint;
            pt.TranslateBy(vec);                                                         
            SketchLine l;
            if (sl == null)
                l = ps.SketchLines.AddByTwoPoints(line.Geometry.MidPoint as object, pt as object);
            else l = ps.SketchLines.AddByTwoPoints(line.Geometry.MidPoint, sl.EndSketchPoint);
            l.Construction = true;
            addMidConsraint(line, l);
            ps.GeometricConstraints.AddParallel(l as SketchEntity, dir as SketchEntity);
            return l;
        }
        public void addDimConstr(SketchPoint p, SketchLine sl, Parameter par = null)
        {
            Vector2d v = ut.normal(sl.StartSketchPoint.Geometry, sl.EndSketchPoint.Geometry), v1 = p.Geometry.VectorTo(sl.StartSketchPoint.Geometry);
            if (v.AngleTo(v1) > Math.PI / 2) v.ScaleBy(-1);
            pt = p.Geometry.Copy();
            pt.TranslateBy(v);
            ObjectsEnumerator en = sl.Geometry.IntersectWithCurve(I.tg.CreateLineSegment2d(p.Geometry, pt), 0.1);
            if (en != null && en.Count != 0)
            {
                pt = en[1] as Point2d;
            }
            OffsetDimConstraint offs = ps.DimensionConstraints.AddOffset(sl, p as SketchEntity, ut.midPt(p.Geometry, pt, p.Geometry.DistanceTo(pt) / 5, 0), false);
            if (par != null) offs.Parameter.Expression = par.Name;
            else offs.Parameter.Value = 0;
        }
        public SketchLine findLine(SketchLine sl)
        {
            sp = sl.StartSketchPoint;
            SketchLine line = findPoint(sp, sl);
            
            if (line != null) return line;
            else { sp = sl.EndSketchPoint;return findPoint(sp, sl); }

        }
        public SketchLine findPoint(SketchPoint sp, SketchLine sl)
        {
            foreach (CoincidentConstraint item in sp.Constraints)  
            {
                if (item.EntityOne is SketchLine && !item.EntityOne.Equals(sl)) {return item.EntityOne as SketchLine;}
                if (item.EntityTwo is SketchLine && !item.EntityTwo.Equals(sl)) {return item.EntityTwo as SketchLine; }
            }
            return null;
        }
    }

    public class SketchInMB
    {
        PartDocument doc;
        PartComponentDefinition def;
        Edge e;
        Face f;
        string offset, l, count;
        string start, end;
        PlanarSketch ps;
        double val = 0;
        TwoPointDistanceDimConstraint rdim = null;
        public List<Parameter> par = new List<Parameter>();
        public SketchPoint[] sp;
        bool cen = true;
        public object dir;
        SketchLine bl, sl;
        public SketchInMB(Document doc, string[] vals)
        {
            def = I.getPCD(doc);
            this.doc = doc as PartDocument;
            if (this.doc.SelectSet.Count == 0) return;
            e = this.doc.SelectSet[1] as Edge;
            if (e == null || def == null) return;
            offset = vals[0]; l = vals[1];
            dir = u.axisFromEdge(def, e);
            addPSArray();
        }
        public SketchInMB(PartDocument doc, Edge e, string [] vals)
        {
            def = doc.ComponentDefinition;
            this.e = e;
            if (vals != null && vals.Length == 3)
            {
                offset = vals[0]; l = vals[1]; count = vals[2];
            }
            addPS();
        }
        public bool isDir()
        {
            WorkAxis wa = dir as WorkAxis;
            if (wa == null) return true;
            Vector v1 = sp[0].Geometry3d.VectorTo(sp[1].Geometry3d);
            var uv = v1.AsUnitVector();
            return u.eq(uv, wa.Line.Direction, 0.001);
        }
        public void addPS()
        {
            var fs = u.gets<Face>(e.Faces, fi => fi.SurfaceType == SurfaceTypeEnum.kPlaneSurface);
            f = fs.First();
            if (fs.Count() == 2) f = getMinAreaFace(fs);
            var pt = def.WorkPoints[1];
            if (f is FaceProxy)
            {
                f = ((FaceProxy)f).NativeObject;
                e = ((EdgeProxy)e).NativeObject;
            }
            ps = def.Sketches.AddWithOrientation(f, e, true, true, pt);
            bl = (SketchLine)ps.AddByProjectingEntity(e); bl.Construction = true;
            sl = addLine(2, 1); sl.Construction = true;
            if (l.EndsWith("%"))
            {
                start = String.Format(@"isolate(round((0.01*{0}*", l.TrimEnd('%'));
                end = String.Format(@")/({0}-1))*({0}-1);ul;mm)",count);
                cen = false; 
                //l = ((bl.Length * u.convToDouble(l.TrimEnd('%')) * 0.01) * 10).ToString();
            }
            if (l.StartsWith("-"))
            {
                val = u.convToDouble(l.TrimStart('-'));
                l = " - 2*" + val + " mm";
                start = String.Format(@"isolate(round((-2*{0}+", val);
                end = String.Format(@")/({0}-1))*({0}-1);ul;mm)", count);
                cen = false;
                //l = ((bl.Length - 2 * val*0.1) * 10).ToString();
            }
            addConstr(sl);
        }
        public void addPSArray()
        {
            var fs = u.gets<Face>(e.Faces, fi => fi.SurfaceType == SurfaceTypeEnum.kPlaneSurface);
            f = fs.First();
            if (fs.Count() == 2) f = getMinAreaFace(fs);
            var pt = def.WorkPoints[1];
            ps = def.Sketches.AddWithOrientation(f, e, true, true, pt);
            constraints.ps = ps;
            bl = (SketchLine)ps.AddByProjectingEntity(e); bl.Construction = true;
            rdim = constraints.addTwoPointDist(bl.StartSketchPoint, bl.EndSketchPoint, null, -0, 3);
            rdim.Driven = true;
            par.Add(rdim.Parameter);
            sp = minSP(bl);
            if (sp == null) return;
            var uv = sp[0].Geometry3d.VectorTo(sp[1].Geometry3d);
            if (!u.eq(uv.X, 0) && u.eq(uv.Y, 0) && u.eq(uv.Z, 0))
            {
                sp = new SketchPoint[] { sp[1], sp[0] };
            }
            var v1 = sp[0].Geometry.VectorTo(sp[1].Geometry);
            var n = v1.Copy();
            u.normal(n, null);
            n.Normalize();
            v1.Normalize();
            var p3 = sp[0].Geometry;
            p3.TranslateBy(v1);
            p3.TranslateBy(n);
            var sp1 = ps.SketchPoints.Add(p3);
            if (!f.Evaluator.RangeBox.Contains(sp1.Geometry3d))
            {
                n.ScaleBy(-2);
                sp1.MoveBy(n);
            }
            var d1 = constraints.addTwoPointDist(bl, sp1, 0.3);
            d1.Parameter.Expression = offset;
            par.Add(d1.Parameter);
            DimensionOrientationEnum align = DimensionOrientationEnum.kHorizontalDim;
            Vector2d v = I.CV2d(sp[0].Geometry, sp[1].Geometry);
            if (v.IsParallelTo(I.CV2d(0, 1), 0.001)) align = DimensionOrientationEnum.kHorizontalDim;
            var d2 = constraints.addTwoPointDist(sp[0], sp1, null, 0.3, 1, align);
            d2.Parameter.Expression = l;
            par.Add(d2.Parameter);
        }
        public SketchPoint[] minSP(SketchLine sl)
        {
            Inventor.Point p1 = sl.StartSketchPoint.Geometry3d, p2 = sl.EndSketchPoint.Geometry3d;
            double[] coord1 = new double[] { }, coord2 = new double[] { };
            p1.GetPointData(ref coord1); p2.GetPointData(ref coord2);
            Vertex v;
            for (int i = 0; i < coord1.Length; i++)
            {
                double c1 = coord1[i], c2 = coord2[i];
                if (u.eq(c1, c2)) continue;
                return c1 < c2 ? new[] { sl.StartSketchPoint, sl.EndSketchPoint } :
                    new[] { sl.EndSketchPoint, sl.StartSketchPoint };
            }
            return null;
        }
        public Face getMinAreaFace(IEnumerable<Face> fs)
        {
            Face f1 = fs.ElementAt(0), f2 = fs.ElementAt(1);
            return f1.Evaluator.Area >= f2.Evaluator.Area ? f1 : f2;
        }
        public SketchLine addLine(double l, double offset)
        {
            LineSegment2d g = bl.Geometry;
            var pt = g.MidPoint;
            var n = u.normal(g.StartPoint, g.EndPoint); u.setDist(n, offset);
            var pt1 = I.CP2d(pt, n);
            if (!f.Evaluator.RangeBox.Contains(ps.SketchToModelSpace(pt1)))
            {
                u.setDist(n, -offset*2); pt1.TranslateBy(n);
            }
            var v = I.CV2d(g.StartPoint, g.EndPoint); u.setDist(v, -l / 2);
            pt1.TranslateBy(v); u.setDist(v,-l);
            var pt2 = I.CP2d(pt1, v);
            return ps.SketchLines.AddByTwoPoints(pt1, pt2);
        }
        public void addConstr(SketchLine sl)
        {
            var sp = ps.SketchPoints.Add(bl.Geometry.MidPoint); sp.HoleCenter = false;
            ps.GeometricConstraints.AddMidpoint(sp, bl);
            var sp2 = ps.SketchPoints.Add(sl.Geometry.MidPoint); sp2.HoleCenter = false;
            ps.GeometricConstraints.AddMidpoint(sp2, sl);
            sp2 = sl.EndSketchPoint;
            constraints.ps = ps;
            DimensionOrientationEnum align = DimensionOrientationEnum.kHorizontalDim;
            Vector2d v = I.CV2d(bl.Geometry.StartPoint, bl.Geometry.EndPoint);
            if (v.IsParallelTo(I.CV2d(0, 1), 0.001)) align = DimensionOrientationEnum.kHorizontalDim;
            if (!cen)
            {
                rdim = constraints.addTwoPointDist(bl.StartSketchPoint, bl.EndSketchPoint, null, -0,3);
                rdim.Driven = true;
                l = start + rdim.Parameter.Name + end;
                ps.GeometricConstraints.AddParallel((SketchEntity)bl, (SketchEntity)sl);
            }
            var d1 = constraints.addTwoPointDist(sl, null, -0.5, 1, align);
            d1.Parameter.Expression = l;
            var dc = constraints.addTwoPointDist(sp, sp2, null, 0.2, 0.1, align);
            dc.Parameter.Expression = d1.Parameter.Name + "/2";
            addPoints(sl, d1);
            var dy = constraints.addTwoPointDist(sl, sp, 0);
            dy.Parameter.Expression = offset;
            //dc.TextPoint = I.CP2d(sp.Geometry);
            //ps.GeometricConstraints.AddVerticalAlign(sp, sp2);
        }
        public void addPoints(SketchLine sl, TwoPointDistanceDimConstraint bc)
        {
            Double len = sl.Length;
            double dx = 0, c = u.convToDouble(count);
            dx = len / c;
            double sum = 0;
            Vector2d v = sl.StartSketchPoint.Geometry.VectorTo(sl.EndSketchPoint.Geometry);
            sl.StartSketchPoint.HoleCenter = true;
            for (int i = 1; i < c; i++)
            {
                sum = i * dx;
                var dc = addPoint(sl, v, sum);
                dc.Parameter.Expression = bc.Parameter.Name + "/(" + c.ToString() + "-1)*" + i.ToString();
            }
        }
        public TwoPointDistanceDimConstraint addPoint(SketchLine sl, Vector2d v, double d)
        {
            SketchPoint bp = sl.StartSketchPoint;
            u.setDist(v, d);
            var pt = I.CP2d(bp.Geometry, v);
            var sp = ps.SketchPoints.Add(pt);
            ps.GeometricConstraints.AddCoincident((SketchEntity)sp, (SketchEntity)sl);
            return constraints.addTwoPointDist(bp, sp, null, -0.2);
        }
    }
    public struct HoleFace
    {
        static public double t, r;
        public Face f;
        public Edge e;
        public double x, y;
        public string count, offset, l;
        public HoleFace(Face fa)
        {
            f = fa;
            count = ""; offset = ""; l = "";
            var rect = f.Evaluator.ParamRangeRect;
            List<double> vals = new List<double>()
            {
                rect.MaxPoint.X - rect.MinPoint.X,
                rect.MaxPoint.Y - rect.MinPoint.Y
            };
            x = vals.Min();
            y = vals.Max();
            e = null;
            findEdge();
        }
        public void findEdge()
        {
            var loop = u.get<EdgeLoop>(f.EdgeLoops, fi => fi.IsOuterEdgeLoop);
            e = u.getEdge(loop, 0, t + r);
        }
    }
    public class AutoHoles
    {
        List<HoleFace> faces = new List<HoleFace>();
        Document doc;
        SheetMetalComponentDefinition def;
        SelectSet ss;
        List<SurfaceBody> bodies = new List<SurfaceBody>();
        MyXML xdoc;
        public AutoHoles(Document doc)
        {
            this.doc = doc;
            def = I.getSMCD(doc);
            I.screenSilent(true);
            ss = doc.SelectSet;
            fill();
            xdoc = new MyXML("ConstraintsFlange.xml");
            var el = xdoc.getEl("Flange");
            foreach (var item in el.Elements())
            {
                var xmin = MyXML.convToDouble(MyXML.getAtt(item, "min"));
                var xmax = MyXML.convToDouble(MyXML.getAtt(item, "max"));
                List<double> vals = new List<double>() { xmin/10, xmax/10 };

                findFaces(vals, item);
            }
            doc.Update2();
            I.screenSilent(false);
        }
        public void fill()
        {
            if (ss.Count > 0)
            {
                foreach (var item in ss)
                {
                    var sb = item as SurfaceBody;
                    if (sb != null) bodies.Add(sb);
                }
            }
            else
            {
                foreach (SurfaceBody item in def.SurfaceBodies)
                {
                    bodies.Add(item);
                }
            }
        }
        public bool checkSketch(Face f)
        {
            foreach (var item in def.Sketches)
            {
                PlanarSketch pl = item as PlanarSketch;
                if (pl == null) continue;
                var fa = pl.PlanarEntity as Face;
                if (fa == null) continue;
                if (fa.Equals(f)) return true;
            }
            return false;
        }
        public void findFaces(List<double> vals, XElement el)
        {
            foreach (SurfaceBody sb in bodies)
            {
                if (!sb.Visible) continue;
                var fs = u.gets<Face>(sb.Faces, fi => fi.SurfaceType == SurfaceTypeEnum.kPlaneSurface);
                var sms = def.GetBodySheetMetalStyle(sb);
                if (sms != null)
                {
                    HoleFace.t = u.convToDouble(sms.Thickness)/10;
                    HoleFace.r = u.convToDouble(sms.BendRadius)/10;
                }
                foreach (var f in fs)
                {
                    var hf = new HoleFace(f);
                    if (hf.e == null) continue;
                    if (hf.x > vals[0] && hf.x <= vals[1])
                    {
                        foreach (var item in el.Elements())
                        {
                            var ymin = MyXML.convToDouble(MyXML.getAtt(item, "min"))/10;
                            var ymax = MyXML.convToDouble(MyXML.getAtt(item, "max"))/10;
                            if (hf.y > ymin && hf.y < ymax)
                            {
                                var e1 = (XElement)item.FirstNode;
                                hf.count = MyXML.getAtt(e1, "count");
                                hf.offset = MyXML.getAtt(e1, "y");
                                hf.l = MyXML.getAtt(e1, "x");
                                faces.Add(hf);
                            }
                        }
                    }
                }
            }
            foreach (var item in faces)
            {
                if (checkSketch(item.f)) continue;
                var sk = new SketchInMB((PartDocument)doc, item.e, new string[] { item.offset, item.l, item.count });
            }
        }
    }
    public struct HoleEdge: IEquatable<HoleEdge>
    {
        public SurfaceBody sb;
        public Edge e;
        Circle g;
        Inventor.Point cen;
        public double d;
        UnitVector n;
        public HoleEdge(Edge e, SurfaceBody sb)
        {
            this.e = e;
            this.sb = sb;
            d = 0;
            g = e.Geometry as Circle;
            cen = g.Center;
            this.n = g.Normal;
        }
        public bool check(HoleEdge he)
        {
            if (he.sb.Equals(sb)) return false;
            if (!u.eq(n.DotProduct(he.n) ,-1)) return false;
            var v = cen.VectorTo(he.cen);
            if (u.eq(v.Length, 0))
            {
                d = 0;
                return true;
            }
            if (n.AsVector().DotProduct(v) > 0) return false;
            d = u.round(v.Length);
            if (d > 0.7) return false;
            return true;
        }

        public bool find(HoleEdge he)
        {
            if (!u.eq(he.d, d)) return false;
            if (!he.sb.Equals(sb)) return false;
            return true;
        }

        public bool Equals(HoleEdge other)
        {
            return e.Equals(other.e);
        }
    }
    public class AutoiMates
    {
        Document doc;
        SheetMetalComponentDefinition def;
        SelectSet ss;
        HashSet<HoleEdge> filter = new HashSet<HoleEdge>();
        List<HoleEdge> edges = new List<HoleEdge>();
        Dictionary<HoleEdge, HoleEdge> pairs = new Dictionary<HoleEdge, HoleEdge>();
        Dictionary<SurfaceBody, SurfaceBody> bodies_ = new Dictionary<SurfaceBody, SurfaceBody>();
        List<SurfaceBody> bodies = new List<SurfaceBody>();

        public AutoiMates(Document doc)
        {
            this.doc = doc;
            def = I.getSMCD(doc);
            ss = doc.SelectSet;
            I.screenSilent(true);
            fill();
            findEdges();
            check();
            create();
            I.screenSilent(false);
        }
        public void findEdges()
        {
            foreach (SurfaceBody sb in bodies)
            {
                if (!sb.Visible) continue;
                var eds = u.gets<Edge>(sb.Edges, fi => fi.GeometryType == CurveTypeEnum.kCircleCurve);
                foreach (var e in eds)
                {
                    var he = new HoleEdge(e, sb);
                    edges.Add(he);
                }
            }
        }
        public List<HoleEdge> find(HoleEdge he, HoleEdge he2)
        {
            List<HoleEdge> f = new List<HoleEdge>();
            foreach (var item in pairs)
            {
                if (item.Key.Equals(he)) continue;
                if (he.find(item.Key))
                {
                    if (item.Value.sb.Equals(he2.sb))
                        f.Add(item.Key);
                }
            }
            return f;
        }
        public void addiMate(HoleEdge e1, HoleEdge e2, string name, double d)
        {
            var obj = I.COC();
            var i1 = def.iMateDefinitions.AddInsertiMateDefinition(e1.e, true, d);
            obj.Add(i1);
            var i2 = def.iMateDefinitions.AddInsertiMateDefinition(e2.e, true, d);
            obj.Add(i2);
            IMate.addName(def.iMateDefinitions.AddCompositeiMateDefinition(obj), name);
        }
        public void create()
        {
            Inventor.ProgressBar pb = I.app.CreateProgressBar(false, pairs.Count, "Создание КП");
            pb.Message = "Создаются КП";
            foreach (var item in pairs)
            {
                if (filter.Contains(item.Key) || filter.Contains(item.Value))
                    continue;
                if ((bodies_.ContainsKey(item.Key.sb) && bodies_[item.Key.sb].Equals(item.Value.sb)) ||
                   (bodies_.ContainsKey(item.Value.sb) && bodies_[item.Value.sb].Equals(item.Key.sb)))
                    continue;
                var tmp = find(item.Key, item.Value);
                if (tmp.Count > 0)
                {
                    filter.Add(item.Key);
                    filter.Add(item.Value);
                    filter.Add(tmp[0]);
                    filter.Add(pairs[tmp[0]]);
                    string n1 = u.shortName(item.Key.sb.Name.Split('$')[0]), n2 = u.shortName(item.Value.sb.Name.Split('$')[0]);
                    if (checkiM(item.Key, $"{n1}_{n2}") || checkiM(item.Value, $"{n2}_{n1}")) continue;
                    var d = getDist(item.Key, item.Value);
                    addiMate(item.Key, tmp[0], $"{n1}_{n2}", d);
                    addiMate(item.Value, pairs[tmp[0]], $"{n2}_{n1}", d);
                    bodies_[item.Key.sb] = item.Value.sb;
                }
                pb.UpdateProgress();
            }
            pb.Close();
        }
        public bool checkiM(HoleEdge he, string name)
        {
            foreach (var item in def.iMateDefinitions)
            {
                CompositeiMateDefinition cdef = item as CompositeiMateDefinition;
                if (cdef == null) continue;
                if (cdef.Name == name) return true;
                foreach (var im in cdef)
                {
                    InsertiMateDefinition i = im as InsertiMateDefinition;
                    if (i.Suppressed) continue;
                    if (i != null)
                    {
                        if (he.e.Equals(i.Entity)) return true;
                    }
                }
            }
            return false;
        }
        public void fill()
        {
            if (ss.Count > 0)
            {
                bodies = CreateComponent.GetBodies(ss).ToList();
                //foreach (var item in ss)
                //{
                //    var sb = item as SurfaceBody;
                //    if (sb != null) bodies.Add(sb);
                //}
            }
            else
            {
                foreach (SurfaceBody item in def.SurfaceBodies)
                {
                    bodies.Add(item);
                }
            }
        }
        public double getDist(HoleEdge e, HoleEdge e2)
        {
            var lst = new List<HoleEdge>() { e, e2 };
            lst = lst.OrderBy(f => f.d).ToList();
            return lst[1].d;
        }
        public void check()
        {
            Inventor.ProgressBar pb = I.app.CreateProgressBar(false, edges.Count, "Проверка ребер");
            pb.Message = "Поиск...";
            foreach (var e in edges)
            {
                foreach (var e2 in edges)
                {
                    if (e.Equals(e2)) continue;
                    if (e.check(e2))
                    {
                        if (pairs.ContainsKey(e) || pairs.ContainsKey(e2)) 
                            continue;
                        pairs[e] = e2;
                    }
                }
                pb.UpdateProgress();
            }
            pb.Close();
        }
    }
    public class PartFromAsm
    {
        AssemblyDocument asm;
        PartDocument docFrom, docTo;
        ComponentOccurrence occFrom, occTo;
        SheetMetalComponentDefinition smcdFrom, smcdTo;
        SurfaceBody sb;
        string sb_name;
        WorkPoint origin;
        SelectSet ss;
        Matrix mtx;
        SketchCopy copy;
        public PartFromAsm(Document doc)
        {
            asm = doc as AssemblyDocument;
            if (asm == null) return;
            init();
            copy = new SketchCopy(docTo as  Document);
            copy.setMtx(mtx, smcdFrom as PartComponentDefinition);
            addUCS();
            //addOrigin();
            foreach (PlanarSketch item in smcdFrom.Sketches)
            {
                copy.add(item, false);
            }
        }
        public void init()
        {
            ss = asm.SelectSet;
            if (ss.Count != 2) return;
            occFrom = ss[1] as ComponentOccurrence;
            occTo = ss[2] as ComponentOccurrence;
            mtx = occFrom.Transformation;
            if (occFrom == null || occTo == null) return;
            docFrom = getDoc(occFrom, ref smcdFrom);
            docTo = getDoc(occTo, ref smcdTo);
        }
        public void addUCS()
        {
            var def = smcdTo.UserCoordinateSystems.CreateDefinition();
            def.Transformation = mtx;
            smcdTo.UserCoordinateSystems.Add(def);
        }
        public PartDocument getDoc(ComponentOccurrence occ, ref SheetMetalComponentDefinition def)
        {
            def = occ.DefinitionReference.ReferencedDefinition as SheetMetalComponentDefinition;
            if (def == null) return null;
            return def.Document as PartDocument;
        }
    }
    public class SketchCopy
    {
        PartDocument doc;
        PartComponentDefinition def, from;
        public PlanarSketch ps, ps_new;
        HashSet<string> bl_names = new HashSet<string>();
        Dictionary<string, string> sketches = new Dictionary<string, string>();
        SurfaceBody body_old, body_new;
        string body_new_name, body_old_name;
        Matrix mtx = null;
        bool link = true, noMtx = false;
        public SketchCopy(Document doc)
        {
            this.doc = doc as PartDocument;
            def = this.doc.ComponentDefinition;
            if (this.doc == null) return;
        }
        public void setMtx(Matrix mtx, PartComponentDefinition def)
        {
            this.mtx = mtx; link = false; from = def;
        }
        public PlanarSketch checkSketch(PlanarSketch ps)
        {
            if (sketches.Keys.Contains(ps.Name)) 
                return u.get<PlanarSketch>(def.Sketches, f => f.Name == sketches[ps.Name]);
            add(ps, true);
            return ps_new;
        }
        public void setBody(SurfaceBody b)
        {
            body_new = b; body_new_name = body_new.Name;
        }
        public SurfaceBody getBody(string name)
        {
            return u.get<SurfaceBody>(def.SurfaceBodies, f => f.Name.ToLower() == name.ToLower());
        }
        public void setOldBody(SurfaceBody b)
        {
            body_old = b; body_old_name = body_old.Name;
        }
        public void add(PlanarSketch ps, bool transform)
        {
            this.ps = ps;
            //noMtx = false;
            addSketch();
            foreach (SketchEntity item in ps.SketchEntities)
            {
                if (mtx == null)
                {
                    if (!addBlock(item))

                        create(item);
                }
                else
                {
                    if (!addBlock(item))
                        createMtx(item, transform);
                }
            }
            //noMtx = true;
            geom();
            dim();
        }
        public List<object> lst(object ob)
        {
            return new List<object>() { ob };
        }
        public void addSketch()
        {
            var pe = ps.PlanarEntity;
            var f = pe as Face;
            if (body_new == null && body_new_name != null)
                body_new = getBody(body_new_name);
            if (f == null || body_new == null)
            {
                addSketch(pe);
            }
            else
            {
                var el = lst(f.PointOnFace);
                if (mtx != null) tr(el);
                var nf = u.findFace(new List<Inventor.Point> { el[0] as Inventor.Point }, body_new);
                if (nf.Count >= 1)
                {
                    addSketch(nf[0]);
                }
            }
            if (ps_new != null) sketches.Add(ps.Name, ps_new.Name);
        }
        public void addSketch(object f)
        {
            object ax = ps.AxisEntity, o = ps.OriginPoint;
            if (mtx == null)
            {
                ps_new = def.Sketches.Add(f);
            }
            else
            {
                ax = findAxis(ps.AxisEntity);
                o = findPoint(ps.OriginPoint);
                if (f is WorkPlane)
                {
                    var pl = findPlane((f as WorkPlane).Plane);
                    ps_new = def.Sketches.Add(pl);
                }
                else if (f is Face)
                {
                    body_new = getBody(body_new_name);
                    var el = lst(((Face)f).PointOnFace);
                    //if (mtx != null) tr(el);
                    var nf = u.findFace(new List<Inventor.Point> { el[0] as Inventor.Point }, body_new);
                    if (nf.Count == 1)
                    {
                        ps_new = def.Sketches.Add(nf[0]);
                    }
                }
            }
            if (ax != null) ps_new.AxisEntity = ax;
            ps_new.AxisIsX = ps.AxisIsX; ps_new.OriginPoint = o; ps_new.NaturalAxisDirection = ps.NaturalAxisDirection;
            if (def.UserCoordinateSystems.Count > 0)
            {
                UnitVector v = ps_new.AxisEntityGeometry.Direction;
                var old = lst(ps.AxisEntityGeometry.Direction); tr(old);
                UnitVector v1 = old[0] as UnitVector;
                if (u.eq(v, v1)) return;

                var ucs = def.UserCoordinateSystems[def.UserCoordinateSystems.Count];
                if (!v.IsParallelTo(ucs.XAxis.Line.Direction))
                {
                    ps_new.AxisIsX = !ps_new.AxisIsX;
                    if (!u.isCollinear(ps_new.AxisEntityGeometry.Direction, ucs.XAxis.Line.Direction))
                        ps_new.NaturalAxisDirection = !ps_new.NaturalAxisDirection;
                }
            }
        }
        public bool addBlock(SketchEntity se)
        {
            if (se.SketchBlockPath.Count == 0) return false;
            SketchBlock bl = se.SketchBlockPath[1];
            SketchBlockDefinition sbd = bl.Definition;
            PlanarSketch ps1 = se.Parent as PlanarSketch;
            var f = lst(ps1.SketchToModelSpace(bl.Position));
            if (mtx != null)
            {
                //tr(f);
                sbd = copyBlock(bl.Definition.Name);
            }
            var name = bl.Name;
            if (bl_names.Contains(name)) return true;
            bl_names.Add(name);

            var nbl = ps_new.SketchBlocks.AddByDefinition(sbd, ps.ModelToSketchSpace(f[0] as Inventor.Point));
            nbl.Transformation = bl.Transformation; 
            return true;
        }
        public SketchBlockDefinition copyBlock(string name)
        {
            var names = u.gets<SketchBlockDefinition>(def.SketchBlockDefinitions, f => true).Select(el => el.Name);
            if (names.Contains(name)) return def.SketchBlockDefinitions[name];
            var sbd = from.SketchBlockDefinitions[name];
            sbd.CopyTo(doc as _Document);
            return def.SketchBlockDefinitions[name];
        }
        public void create(SketchEntity se)
        {
            SketchEntity nse = null;
            dynamic ent = se;
            if (body_new == null && body_new_name != null)
                body_new = getBody(body_new_name);
            var r = findRef(se, body_new);
            if (r != null) { r.Construction = se.Construction; 
                if (r is SketchPoint)
                {
                    ((SketchPoint)r).HoleCenter = ((SketchPoint)se).HoleCenter;
                }
                return; }
            if (se.Type == ObjectTypeEnum.kSketchPointObject)
            {
                //var con = u.get<GeometricConstraint>(se.Constraints, f => f.Type == ObjectTypeEnum.kConcentricConstraintObject);
                //if (con != null) return;
                nse = ps_new.SketchPoints.Add(ent.Geometry, ent.HoleCenter) as SketchEntity;
            } else if (se.Type == ObjectTypeEnum.kSketchLineObject)
            {
                nse = ps_new.SketchLines.AddByTwoPoints(getPoint(ent.StartSketchPoint, ObjectTypeEnum.kSketchPointObject), 
                   getPoint(ent.EndSketchPoint, ObjectTypeEnum.kSketchPointObject)) as SketchEntity;
            } else if (se.Type == ObjectTypeEnum.kSketchCircleObject)
            {
                SketchEntity pt = getPoint(ent.CenterSketchPoint, ObjectTypeEnum.kSketchPointObject);
                nse = ps_new.SketchCircles.AddByCenterRadius(pt, ent.Radius) as SketchEntity;
                pt.Delete();
            } else if (se.Type == ObjectTypeEnum.kSketchArcObject)
            {
                SketchEntity pt = getPoint(ent.CenterSketchPoint, ObjectTypeEnum.kSketchPointObject);
                nse = ps_new.SketchArcs.AddByCenterStartEndPoint(pt,
                    getPoint(ent.StartSketchPoint, ObjectTypeEnum.kSketchPointObject),
                    getPoint(ent.EndSketchPoint, ObjectTypeEnum.kSketchPointObject)) as SketchEntity;
                pt.Delete();
            }
            if (nse != null && se.Construction) nse.Construction = se.Construction;
        }
        public Point2d get3dTo2d(object pt)
        {
            return ps_new.ModelToSketchSpace(pt as Inventor.Point);
        }
        public void createMtx(SketchEntity se, bool transform = true)
        {
            SketchEntity nse = null;
            dynamic ent = se;
            PlanarSketch ps1 = se.Parent as PlanarSketch;
            var r = findRefMtx(se, getBody(body_new_name));
            if (r != null)
            {
                r.Construction = se.Construction;
                if (r is SketchPoint)
                {
                    ((SketchPoint)r).HoleCenter = ((SketchPoint)se).HoleCenter;
                }
                return;
            }
            if (se.Type == ObjectTypeEnum.kSketchPointObject)
            {
                //var con = u.get<GeometricConstraint>(se.Constraints, f => f.Type == ObjectTypeEnum.kConcentricConstraintObject);
                //if (con != null) return;
                var f = lst(ent.Geometry3d);
                if (transform)
                    tr(f);
                nse = ps_new.SketchPoints.Add(get3dTo2d(f[0]), ent.HoleCenter) as SketchEntity;
            }
            else if (se.Type == ObjectTypeEnum.kSketchLineObject)
            {
                var f = lst(ent.StartSketchPoint.Geometry3d); f.Add(ent.EndSketchPoint.Geometry3d);
                if (transform)
                    tr(f);
                SketchEntity p1 = checkPoint(f[0]), p2 = checkPoint(f[1]);
                nse = ps_new.SketchLines.AddByTwoPoints(p1, p2) as SketchEntity;
            }
            else if (se.Type == ObjectTypeEnum.kSketchCircleObject)
            {
                var f = lst(ent.CenterSketchPoint.Geometry3d); if (transform) tr(f);
                SketchEntity pt = getPoint(f[0]);
                nse = ps_new.SketchCircles.AddByCenterRadius(pt, ent.Radius) as SketchEntity;
                deleteEnt(pt);
            }
            else if (se.Type == ObjectTypeEnum.kSketchArcObject)
            {
                var f = lst(ent.CenterSketchPoint.Geometry3d); f.Add(ent.StartSketchPoint.Geometry3d);
                f.Add(ent.EndSketchPoint.Geometry3d); if (transform) tr(f);
                SketchEntity pt = checkPoint(f[0]), sp = checkPoint(f[1]), ep = checkPoint(f[2]);
                nse = ps_new.SketchArcs.AddByCenterStartEndPoint(pt, sp, ep) as SketchEntity;
                deleteEnt(pt); deleteEnt(sp); deleteEnt(ep);
                //pt.Delete();
            }
            //else if (se.Type == ObjectTypeEnum.kSketchEllipticalArcObject)
            //{
            //    var f = lst(ent.Geometry3d.Center); f.Add(ent.Geometry3d.MajorAxis);
            //    f.Add(ent.Geometry3d.MinorAxis); if (transform) tr(f);
            //    SketchEntity pt = checkPoint(f[0]);
            //    nse = ps_new.SketchEllipticalArcs.Add(pt, )
            //    pt.Delete();
            //}
            if (nse != null && se.Construction) nse.Construction = se.Construction;
        }
        public void deleteEnt(SketchEntity pt)
        {
            if (pt.OwnedBy == null || pt.OwnedBy.Count == 0) pt.Delete();
        }
        public SketchPoint checkPoint(object pt)
        {
            SketchPoint sp = getPoint(pt) as SketchPoint;
            if (sp == null) sp = ps_new.SketchPoints.Add(get3dTo2d(pt), false);
            return sp;
        }
        public void geom()
        {
            foreach (GeometricConstraint g in ps.GeometricConstraints)
            {
                con(g);
            }
        }
        public void dim()
        {
            foreach (DimensionConstraint d in ps.DimensionConstraints)
            {
                con(d);
            }
        }
        public SketchEntity getSE(SketchEntity ent1, ObjectTypeEnum t, SketchEntity ent2 = null)
        {
            SketchEntity se = null;
            if (t == ObjectTypeEnum.kSketchCircleObject)
                se = getCircle(ent1);
            else if (t == ObjectTypeEnum.kSketchArcObject)
                se = getArc(ent1);
            else if (t == ObjectTypeEnum.kSketchLineObject)
                se = getLine(ent1) as SketchEntity;
            else if (t == ObjectTypeEnum.kSketchPointObject)
                se = getPoint(ent1, t, true, ent2);
            return se;
        }
        public SketchEntity getSE(GeometricConstraint c, bool one = true)
        {
            dynamic ent = c;
            SketchEntity se = null;
            if (c.Type == ObjectTypeEnum.kVerticalAlignConstraintObject ||
                 c.Type == ObjectTypeEnum.kHorizontalAlignConstraintObject)
            {
                se = one ? ent.PointOne : ent.PointTwo;
            }
            else if (c.Type == ObjectTypeEnum.kVerticalConstraintObject ||
                c.Type == ObjectTypeEnum.kHorizontalConstraintObject || c.Type == ObjectTypeEnum.kGroundConstraintObject)
            {
                if (one)
                se = ent.Entity;
            }
            else if (c.Type == ObjectTypeEnum.kMidpointConstraintObject)
            {
                se = one ? ent.Point : ent.Line;
            }
            else if (c.Type == ObjectTypeEnum.kEqualLengthConstraintObject)
            {
                se = one ? ent.LineOne : ent.LineTwo;
            }
            else
            {
                se = one ? ent.EntityOne : ent.EntityTwo;
            }
            return se;
        }
        public bool checkOwner(SketchEntity se1, SketchEntity se2)
        {
            if (se1.OwnedBy.Count == 0) return false;
            foreach (SketchEntity item in se1.OwnedBy)
            {
                if (item.Equals(se2)) return true;
            }
            return false;
        }
        public void con(GeometricConstraint c)
        {
            if (c == null) return;
            //Line l1 = ps.AxisEntityGeometry, l2 = ps_new.AxisEntityGeometry;
            //UnitVector v1 = l1.Direction, v2 = l2.Direction;
            //if (mtx != null)
            //{
            //    v1.TransformBy(mtx);
            //}
            //bool dir = v1.IsPerpendicularTo(v2);
            dynamic ent = c;
            SketchEntity se = null, se2 = null;
            ObjectTypeEnum t = ObjectTypeEnum.k3dAViewObject, t2 = ObjectTypeEnum.k3dAViewObject;
            //try
            //{
            se = getSE(c, true); se2 = getSE(c, false);
            if (se != null) {
                t = se.Type; 
                se = getSE(se, t);
            }
            if (se2 != null) {
                t2 = se2.Type;
                se2 = getSE(se2, t2);
                if (c.Type == ObjectTypeEnum.kCoincidentConstraintObject && se.Equals(se2))
                {
                    if (mtx == null)
                        se2 = getSE(se2, t2, se);
                    else return;
                }
            }
            if (c.Type == ObjectTypeEnum.kConcentricConstraintObject)
            {
                if (t == ObjectTypeEnum.kSketchPointObject || t2 == ObjectTypeEnum.kSketchPointObject) { }
                else
                    ps_new.GeometricConstraints.AddConcentric(se, se2);
            } else if (c.Type == ObjectTypeEnum.kCoincidentConstraintObject)
            {
                if (t2 == ObjectTypeEnum.kSketchPointObject && (checkOwner(se2, se) || checkOwner(se, se2))) { }
                else
                    ps_new.GeometricConstraints.AddCoincident(se, se2);
            } else if (c.Type == ObjectTypeEnum.kVerticalAlignConstraintObject)
            {
                ps_new.GeometricConstraints.AddVerticalAlign(se as SketchPoint, se2 as SketchPoint);
            } else if (c.Type == ObjectTypeEnum.kHorizontalAlignConstraintObject)
            {
                ps_new.GeometricConstraints.AddHorizontalAlign(se as SketchPoint, se2 as SketchPoint);
            } else if (c.Type == ObjectTypeEnum.kCollinearConstraintObject)
            {
                ps_new.GeometricConstraints.AddCollinear(se, se2);
            }
            else if (c.Type == ObjectTypeEnum.kHorizontalConstraintObject)
            {
                //if (dir) ps_new.GeometricConstraints.AddVertical(se);
                ps_new.GeometricConstraints.AddHorizontal(se);
            }
            else if (c.Type == ObjectTypeEnum.kVerticalConstraintObject)
            {
                //if (dir) ps_new.GeometricConstraints.AddHorizontal(se);
                ps_new.GeometricConstraints.AddVertical(se);
            }
            else if (c.Type == ObjectTypeEnum.kGroundConstraintObject)
            {
                ps_new.GeometricConstraints.AddGround(se);
            } else if (c.Type == ObjectTypeEnum.kMidpointConstraintObject)
            {
                ps_new.GeometricConstraints.AddMidpoint(se as SketchPoint, se2 as SketchLine);
            } else if (c.Type == ObjectTypeEnum.kParallelConstraintObject)
            {
                ps_new.GeometricConstraints.AddParallel(se, se2);
            } else if (c.Type == ObjectTypeEnum.kPerpendicularConstraintObject)
            {
                ps_new.GeometricConstraints.AddPerpendicular(se, se2);
            } else if (c.Type == ObjectTypeEnum.kTangentSketchConstraintObject)
            {
                ps_new.GeometricConstraints.AddTangent(se, se2);
            } else if (c.Type == ObjectTypeEnum.kEqualLengthConstraintObject)
            {
                ps_new.GeometricConstraints.AddEqualLength(se as SketchLine, se2 as SketchLine);
            } else if (c.Type == ObjectTypeEnum.kEqualRadiusConstraintObject)
            {
                ps_new.GeometricConstraints.AddEqualRadius(se, se2);
            } else if (c.Type == ObjectTypeEnum.kSymmetryConstraintObject)
            {
                var sym = c as SymmetryConstraint;
                var l = sym.SymmetryLine;
                var nl = getLine(l as SketchEntity);
                if (nl == null) 
                    nl = getRef(l as SketchEntity) as SketchLine;
                ps_new.GeometricConstraints.AddSymmetry(se, se2, nl);
                }
            //}
            //catch (Exception)
            //{
            //}
        }
        public void con(DimensionConstraint c)
        {
            if (c == null) return;
            dynamic ent = c;
            var f = lst(ps.SketchToModelSpace(c.TextPoint));
            if (mtx != null) tr(f);
            Point2d tp = get3dTo2d(f[0]);
            DimensionConstraint d = null;
            if (c.Type == ObjectTypeEnum.kOffsetDimConstraintObject)
            {
                SketchEntity se2 = null;
                if ((ObjectTypeEnum)ent.Entity.Type == ObjectTypeEnum.kSketchLineObject)
                {
                    se2 = getLine(ent.Entity) as SketchEntity;
                } else if ((ObjectTypeEnum)ent.Entity.Type == ObjectTypeEnum.kSketchPointObject)
                {
                    se2 = getPoint(ent.Entity, ObjectTypeEnum.kSketchPointObject);
                }
                d = ps_new.DimensionConstraints.AddOffset(getLine(ent.Line), se2,
                    tp, ent.LinearDiameter);
                if (ent.Line.Reference && se2.Reference)
                {
                    d.Driven = true;
                }
            } else if (c.Type == ObjectTypeEnum.kDiameterDimConstraintObject)
            {
                SketchEntity cir = null;
                if ((ObjectTypeEnum)ent.Entity.Type == ObjectTypeEnum.kSketchCircleObject) cir = getCircle(ent.Entity);
                else if ((ObjectTypeEnum)ent.Entity.Type == ObjectTypeEnum.kSketchArcObject) cir = getArc(ent.Entity);
                d = ps_new.DimensionConstraints.AddDiameter(cir, tp, ent.Driven);
            } else if (c.Type == ObjectTypeEnum.kRadiusDimConstraintObject)
            {
                d = ps_new.DimensionConstraints.AddDiameter(getArc(ent.Entity), tp, ent.Driven);
            } else if (c.Type == ObjectTypeEnum.kTwoPointDistanceDimConstraintObject)
            {
                d = ps_new.DimensionConstraints.AddTwoPointDistance(getPoint(ent.PointOne, ObjectTypeEnum.kSketchPointObject) as SketchPoint,
                    getPoint(ent.PointTwo, ObjectTypeEnum.kSketchPointObject) as SketchPoint, ent.Orientation,
                    tp, ent.Driven);

            }
            else if (c.Type == ObjectTypeEnum.kTwoLineAngleDimConstraintObject)
            {
                d = ps_new.DimensionConstraints.AddTwoLineAngle(getLine(ent.LineOne),
                    getLine(ent.LineTwo), tp, ent.Driven);
            }
            if (d != null && link)
            {
                if (!d.Driven)
                    d.Parameter.Expression = c.Parameter.Name;
            }
        }
        public SketchEntity getPoint(SketchEntity se, ObjectTypeEnum type, bool first = true, SketchEntity se2 = null)
        {
            dynamic ent = se;
            Inventor.Point pt = null;
            if (type == ObjectTypeEnum.kConcentricConstraintObject)
            {
                if (se.Type == ObjectTypeEnum.kSketchCircleObject || se.Type == ObjectTypeEnum.kSketchArcObject)
                {
                    pt = ent.CenterSketchPoint.Geometry3d;
                }
            }
            else if (type == ObjectTypeEnum.kCoincidentConstraintObject)
            {
                pt = first ? ent.StartSketchPoint.Geometry3d: ent.EndSketchPoint.Geometry3d;
            }
            else if (type == ObjectTypeEnum.kSketchPointObject)
            {
                pt = ((SketchPoint)se).Geometry3d;
            }
            if (mtx != null)
            {
                var f = lst(pt); tr(f);
                pt = f[0] as Inventor.Point;
            }
            return findPoint(pt, se2);
        }
        public SketchEntity getPoint(object pt)
        {
            //Point2d p2 = get3dTo2d(pt);
            return findPoint(pt as Inventor.Point);
        }
        public SketchEntity findPoint(Inventor.Point pt, SketchEntity se2 = null)
        {
            foreach (SketchPoint item in ps_new.SketchPoints)
            {
                if (u.eq(item.Geometry3d, pt))
                {
                    if (se2 != null && item.Equals(se2)) continue;
                    if (se2 != null)
                    {
                        if (se2.Constraints.Count != 0 && item.Constraints.Count != 0)
                        {
                            CoincidentConstraint c1 = se2.Constraints[1] as CoincidentConstraint,
                                c2 = item.Constraints[1] as CoincidentConstraint;
                            if (c1 != null && c2 != null && c1.Equals(c2)) continue;
                        }
                    }
                    return item as SketchEntity;
                }
            }
            return null;
        }
        public SketchLine getLine(SketchEntity se)
        {
            SketchLine sl = se as SketchLine;
            var f = lst(sl.Geometry3d.StartPoint); f.Add(sl.Geometry3d.EndPoint);
            if (mtx != null && !noMtx) tr(f);
            foreach (SketchLine item in ps_new.SketchLines)
            {
                //Point2d p1 = u.midPt(item.Geometry.StartPoint, item.Geometry.EndPoint),
                //    p2 = u.midPt(ent.Geometry.StartPoint, ent.Geometry.EndPoint);
                if (u.eq(item, f))
                    return item;
            }
            return null;
        }
        public SketchEntity getRef(SketchEntity se)
        {
            var r = se.ReferencedEntity;
            foreach (SketchEntity item in ps_new.SketchEntities)
            {
                var r1 = item.ReferencedEntity;
                if (r1 == null) continue;
                if (r.Equals(r1)) return item;
            }
            return null;
        }
        public SketchEntity getCircle(SketchEntity se)
        {
            SketchCircle ent = se as SketchCircle;
            var f = lst(ent.CenterSketchPoint.Geometry3d);
            if (mtx != null && !noMtx) tr(f);
            foreach (SketchCircle item in ps_new.SketchCircles)
            {
                if (u.eq(f[0] as Inventor.Point, item.CenterSketchPoint.Geometry3d) && u.eq(ent.Radius, item.Radius))
                    return item as SketchEntity;
            }
            return null;
        }
        public SketchEntity getArc(SketchEntity se)
        {
            SketchArc ent = se as SketchArc;
            var f = lst(ent.Geometry3d.Center);
            f.Add(ent.Geometry3d.StartPoint); f.Add(ent.Geometry3d.EndPoint);
            if (mtx != null && !noMtx) tr(f);
            foreach (SketchArc item in ps_new.SketchArcs)
            {
                //if (mtx != null)
                //{
                //    if (u.eq(f[0] as Inventor.Point, item.Geometry3d.Center) && u.eq(ent.Radius, item.Radius) &&
                //    u.eq(ent.Geometry3d.StartAngle, item.Geometry3d.StartAngle) &&
                //    u.eq(ent.Geometry3d.SweepAngle, item.Geometry3d.SweepAngle)
                //    )
                //        return item as SketchEntity;
                //}
                //else
                //{
                    if (u.eq(f[0] as Inventor.Point, item.Geometry3d.Center, 4) && u.eq(ent.Radius, item.Radius) &&
                    u.eq(f[1] as Inventor.Point, item.Geometry3d.StartPoint, 4) &&
                    u.eq(f[2] as Inventor.Point, item.Geometry3d.EndPoint, 4)
                    //u.eq(ent.Geometry3d.StartAngle, item.Geometry3d.StartAngle) &&
                    //u.eq(ent.Geometry3d.SweepAngle, item.Geometry3d.SweepAngle))
                    )
                        return item as SketchEntity;
                //}
            }
            return null;
        }
        public SketchEntity findRef(SketchEntity se, SurfaceBody sb = null)
        {
            object ob = null;
            var r = se.ReferencedEntity;
            if (r is ObjectCollection)
            {
                r = ((ObjectCollection)se.ReferencedEntity)[1];
            }
            dynamic ent = r;
            if (r == null) return null;
            if (r is WorkAxis) ob = u.get<WorkAxis>(def.WorkAxes, f => f.Name == ent.Name);
            else if (r is WorkPlane) ob = u.get<WorkPlane>(def.WorkPlanes, f => f.Name == ent.Name);
            else if (r is WorkPoint) ob = u.get<WorkPoint>(def.WorkPoints, f => f.Name == ent.Name);
            else if (r is Vertex) ob = u.get<Vertex>(sb.Vertices, f => u.eq(f.Point, (r as Vertex).Point));
            else if (r is SketchEntity) ob = r;
            else if (r is Edges)
            {
                Edge e_old = (r as Edges)[1]; 
                string n = e_old.Faces[1].SurfaceBody.Name;
                if (n == body_old_name)
                {
                    var mp = u.midPt(e_old);
                    ob = u.get<Edge>(sb.Edges, f => u.isPointOnEdge(mp, f));
                }
                else
                {
                    ob = e_old;
                }
            }
            if (ob == null) return null;
            return ps_new.AddByProjectingEntity(ob);
        }
        public SketchEntity findRefMtx(SketchEntity se, SurfaceBody sb = null)
        {
            object ob = null;
            var r = se.ReferencedEntity;
            if (r is ObjectCollection)
            {
                r = ((ObjectCollection)se.ReferencedEntity)[1];
            }
            dynamic ent = r;
            if (r == null) return null;
            if (r is WorkAxis) ob = findAxis(r);
            else if (r is WorkPlane) ob = findPlane(((WorkPlane)r).Plane);
            else if (r is WorkPoint) ob = findPoint(r);
            else if (r is Vertices) ob = findPoint((r as Vertices)[1]);
            else if (r is SketchEntity) ob = r;
            else if (r is Edges)
            {
                Edge e_old = (r as Edges)[1];
                string n = e_old.Faces[1].SurfaceBody.Name;
                if (n == body_old_name)
                    ob = u.get<Edge>(sb.Edges, f => u.eq(u.midPt(f), u.midPt(e_old)));
                else
                {
                    ob = e_old;
                }
            }
            if (ob == null) return null;
            return ps_new.AddByProjectingEntity(ob);
        }
        public object findPoint(object p)
        {
            WorkPoint wx = p as WorkPoint;
            if (wx != null)
            {
                List<object> f = new List<object>() { wx.Point };
                tr(f);
                foreach (WorkPoint item in def.WorkPoints)
                {
                    List<object> t = new List<object>() { item.Point };
                    if (check(f, t)) return item;
                }
            }
            Vertex ed = p as Vertex;
            if (ed != null)
            {
                List<object> f = new List<object>() { ed.Point }; tr(f);
                foreach (Vertex item in u.gets<Vertex>(getBody(body_new_name).Vertices, fi => true))
                {
                    List<object> t = new List<object>() { item.Point };
                    if (check(f, t)) return item;
                }
            }
            return null;
        }
        public object findPlane(object x)
        {
            WorkPlane wp = x as WorkPlane;
            if (wp != null) return findPlane(wp.Plane);
            Face f = x as Face;
            if (f != null) return findPlane(f.Geometry as Plane);
            return null;
        }
        public object findPlane(Plane pl)
        {
            List<object> f = new List<object>() { pl.RootPoint, pl.Normal };
            tr(f);
            foreach (WorkPlane item in def.WorkPlanes)
            {
                var p = item.Plane;
                List<object> t = new List<object>() { p.RootPoint, p.Normal };
                if (check(f, t)) return item;
            }
            this.body_new = getBody(body_new_name);
            foreach (Face item in body_new.Faces)
            {
                if (item.SurfaceType != SurfaceTypeEnum.kPlaneSurface) continue;
                Plane g = item.Geometry as Plane;
                List<object> t = new List<object>() { g.RootPoint, g.Normal };
                if (check(f, t)) return item;
            }
            return null;
        }
        public object findAxis(object x)
        {
            WorkAxis wx = x as WorkAxis;
            if (wx != null)
            {
                List<object> f = new List<object>() { wx.Line.RootPoint, wx.Line.Direction };
                tr(f);
                foreach (WorkAxis item in def.WorkAxes)
                {
                    List<object> t = new List<object>() { item.Line.RootPoint, item.Line.Direction };
                    if (check(f, t)) return item;
                }
            }
            Edge ed = x as Edge;
            if (ed == null && x is EdgeUse)
            {
                ed = ((EdgeUse)x).Edge;
            }
            if (ed != null)
            {
                var mp = u.midPt(ed);
                List<object> f = new List<object>() { mp }; tr(f);
                this.body_new = getBody(body_new_name);
                foreach (Edge item in body_new.Edges)
                {
                    if (item.CurveType != CurveTypeEnum.kLineCurve) continue;
                    if (u.isPointOnEdge(f[0] as Inventor.Point, item)) return item;
                    //List<object> t = new List<object>() { item.StartVertex.Point, item.StopVertex.Point };
                    //if (check(f, t)) return item;
                }
            }
            Face face = x as Face;
            if (face != null)
            {
                Cylinder g = face.Geometry as Cylinder;
                if (g == null) return null;
                List<object> f = new List<object>() { g.BasePoint }; tr(f);
                this.body_new = getBody(body_new_name);
                foreach (Face item in body_new.Faces)
                {
                    if (item.SurfaceType != SurfaceTypeEnum.kCylinderSurface) continue;
                    Cylinder g1 = item.Geometry as Cylinder;
                    if (u.eq(f[0] as Inventor.Point, g1.BasePoint)) return item;
                }
            }
            return null;
        }
        public bool checkPlane(Plane p, Inventor.Point pt, UnitVector v)
        {
            return u.eq(p.Normal, v) && u.eq(p.RootPoint, pt);
        }
        public void tr(List<object> l)
        {
            for (int i = 0; i < l.Count; i++)
            {
                dynamic ent = l[i];
                ent.TransformBy(mtx);
            }
        }
        public bool check(List<object> f, List<object> t)
        {
            bool r = false;
            for (int i = 0; i < f.Count; i++)
            {
                if (f[i] is Inventor.Point)
                    r = u.eq(f[i] as Inventor.Point, t[i] as Inventor.Point);
                else if (f[i] is Vector)
                    r = u.eq(f[i] as Vector, t[i] as Vector);
                else if (f[i] is UnitVector)
                    r = u.eq(f[i] as UnitVector, t[i] as UnitVector);
                if (!r) return false;
            }
            return r;
        }
    }
    public class DrawDims
    {
        DrawingDocument from, to;
        //Sheet sh;
        List<XMLDim> dims = new List<XMLDim>();
        MyXML xml;

        public DrawDims(Document doc)
        {
            if (doc.DocumentType != DocumentTypeEnum.kDrawingDocumentObject) return;
            to = doc as DrawingDocument;
            init();
            foreach (Sheet item in to.Sheets)
            {
                if (!(item.Status == DrawingSheetStatusBits.kUpToDateDrawingSheet))
                    item.Update();
                findDims(item);
            }
            addXML();
        }
        public DrawDims(XElement el)
        {
            var doc = I.aDoc();
            if (doc.DocumentType != DocumentTypeEnum.kDrawingDocumentObject) return;
            to = doc as DrawingDocument;
            init();
            foreach (Sheet item in to.Sheets)
            {
                readXML(el, item);
            }

        }
        public void init()
        {
            //sh = to.ActiveSheet;
        }
        public void findDims(Sheet sh)
        {
            foreach (DrawingDimension item in sh.DrawingDimensions)
            {
                dims.Add(new XMLDim(item));
            }
            foreach (HoleThreadNote item in sh.DrawingNotes.HoleThreadNotes)
            {
                dims.Add(new XMLDim(item));
            }
        }
        public void addXML()
        {
            string p = file.p(to.FullDocumentName);
            xml = new MyXML($"{p}dims.xml", "head");
            foreach (var item in dims)
            {
                xml.addElem(item.getXML());
            }
            xml.save();
        }
        public void readXML(XElement el, Sheet sh)
        {
            foreach (var item in el.Elements())
            {
                dims.Add(new XMLDim(item, sh));
            }
        }
    }
    public class DCurve
    {
        public DrawingCurve dc;
        DrawingView dv;
        Point2d origin;
        public Point2d pt;
        Vector2d v;
        bool start = true;
        double w, h, sc;
        public double d = 10000;
        public DCurve(DrawingCurve c)
        {
            dc = c;
            dv = dc.Parent;
            origin = dv.Center;
            sc = dv.Scale;
            w = dv.Width; h = dv.Height;
        }
        public Vector2d scale(Point2d pt)
        {
            var v = origin.VectorTo(pt);
            transform(v);
            return v;
        }
        public Vector2d getDir()
        {
            return dc.StartPoint.VectorTo(dc.EndPoint);
        }
        public void transform(Vector2d item)
        {
            //item.ScaleBy(1 / sc);
            item.X /= w / 2;
            item.Y /= h / 2;
        }
        public void add(Vector2d v, UnitVector2d n)
        {
            if (dc.CurveType == CurveTypeEnum.kCircleCurve || dc.CurveType == CurveTypeEnum.kCircularArcCurve)
                pt = dc.CenterPoint;
            else pt = dc.StartPoint;
            this.v = v.Copy();
            var v1 = scale(pt);
            d = Math.Abs(u.dotProduct(v1, n.AsVector()) - u.dotProduct(v, n.AsVector()));

            //add(dc.StartPoint);
            //if (add(dc.EndPoint)) start = false;
        }
        public bool add(Point2d pt)
        {
            var v1 = scale(pt);
            v1.ScaleBy(-1); v1.AddVector(v);
            this.pt = pt;
            if (v1.Length < d)
            {
                d = v1.Length; return true;
            }

            return false;
        }
        public GeometryIntent getInt()
        {
            return dv.Parent.CreateGeometryIntent(dc);
        }
    }
    public class XMLDim
    {
        DrawingDimension d;
        HoleThreadNote ht;
        double w, h, sc, sc1, len;
        string dimText = "", stName = "";
        DrawingView dv;
        UnitVector2d dir, n;
        DimensionTypeEnum al;
        Sheet sh;
        ObjectTypeEnum t;
        DCurve dc1, dc2;
        List<Vector2d> vecs = new List<Vector2d>();
        int dvnum, shnum;
        Point2d p1, p2, txt, origin;
        public XMLDim(DrawingDimension dim)
        {
            d = dim;
            sh = d.Parent;
            t = d.Type;
            getCenter(d);
            if (t == ObjectTypeEnum.kLinearGeneralDimensionObject) addLin();
            else if (t == ObjectTypeEnum.kRadiusGeneralDimensionObject ||
                t == ObjectTypeEnum.kDiameterGeneralDimensionObject ||
                t == ObjectTypeEnum.kHoleThreadNoteObject) addRad();
            else if (t == ObjectTypeEnum.kAngularGeneralDimensionObject)
                addAngle();
            scale();
        }
        public XMLDim(HoleThreadNote dim)
        {
            ht = dim;
            sh = dim.Parent;
            t = dim.Type; 
            getCenter(ht); 
            if (t == ObjectTypeEnum.kRadiusGeneralDimensionObject ||
                t == ObjectTypeEnum.kDiameterGeneralDimensionObject ||
                t == ObjectTypeEnum.kHoleThreadNoteObject) addRad();
            scale();
        }
        public XMLDim(XElement el, Sheet sh)
        {
            this.sh = sh;
            readXML(el);
            if (t == ObjectTypeEnum.kLinearGeneralDimensionObject)
            {
                dc1 = findDC(vecs[0], n);
                dc2 = findDC(vecs[1], n);
                if (dc1 == null && dc2 == null) return;
            }
            else if (t == ObjectTypeEnum.kRadiusGeneralDimensionObject ||
                t == ObjectTypeEnum.kDiameterGeneralDimensionObject ||
                t == ObjectTypeEnum.kHoleThreadNoteObject)
            {
                dc1 = findDC(vecs[0], n);
                if (dc1 == null) return;
            }
            else if (t == ObjectTypeEnum.kAngularGeneralDimensionObject)
            {
                dc1 = findDC(vecs[0], n);
                var dvec = dc1.getDir();
                dc2 = findDC(vecs[1], n, dvec);
                if (dc1 == null && dc2 == null) return;
            }
            //u.addText(dv, "1", dc1.pt);
            //u.addText(dv, "2", dc2.pt);
            //u.higlight(sh.Parent as Document, dc1.dc, dc2.dc, null, null);
            addDim();
        }
        public void addDim()
        {
            DimensionStyle st = getStyle();
            var v = vecs[vecs.Count-1].Copy(); v.X *= w/2; v.Y *= h/2;
            var pt = origin.Copy();
            pt.TranslateBy(v);
            GeometryIntent gi = dc1.getInt();
            if (t == ObjectTypeEnum.kLinearGeneralDimensionObject)
            {
                var dim = sh.DrawingDimensions.GeneralDimensions.AddLinear(pt, gi, dc2.getInt(), al, DimensionStyle: st);
                dim.CenterText();
                if (Math.Abs(v.X) > 1 || Math.Abs(v.Y) > 1)
                {
                    var dist = getDist(dim.ExtensionLineOne as LineSegment2d);
                    var ve = dir.AsVector();
                    ve.ScaleBy(len - dist);
                    pt = dim.Text.Origin.Copy(); pt.TranslateBy(ve);
                    dim.Text.Origin = pt;
                }
                check(dim);
                if (dimText != "") dim.Text.FormattedText = dimText;
                if (sh.SurfaceTextureSymbols.Count > 1) return;
                var rd = dv.ReferencedFile.ReferencedDocument as Document;
                if (rd.DocumentType == DocumentTypeEnum.kPartDocumentObject)
                {
                    var smcd = I.getSMCD(rd);
                    if (smcd == null) return;
                    
                    if (u.eq(getDist(dim.DimensionLine as LineSegment2d)*sc, smcd.Thickness.ModelValue))
                    {
                        Drawings.addSurfaceTextureSymbol(dim, 0.3, true);
                        Drawings.addSurfaceTextureSymbol(dim, 0.3, false);
                    }
                }
            }
            else if (t == ObjectTypeEnum.kRadiusGeneralDimensionObject)
            {
                var arc = gi.Geometry as DrawingCurve;
                var shPt = arc.CenterPoint;
                var ve = dir.AsVector(); ve.ScaleBy(-1);
                ve.ScaleBy(len); shPt.TranslateBy(ve);
                var dim = sh.DrawingDimensions.GeneralDimensions.AddRadius(shPt, gi, DimensionStyle: st);
                if (dimText != "") dim.Text.FormattedText = dimText;
            }
            else if (t == ObjectTypeEnum.kDiameterGeneralDimensionObject)
            {
                var arc = gi.Geometry as DrawingCurve;
                var shPt = arc.CenterPoint;
                var ve = dir.AsVector(); ve.ScaleBy(-1);
                ve.ScaleBy(len); shPt.TranslateBy(ve);
                
                var dim = sh.DrawingDimensions.GeneralDimensions.AddDiameter(shPt, gi, false, true, false, st);
                if (dimText != "") dim.Text.FormattedText = dimText;
            }
            else if (t == ObjectTypeEnum.kHoleThreadNoteObject)
            {
                var arc = gi.Geometry as DrawingCurve;
                var shPt = arc.CenterPoint;
                var ve = dir.AsVector(); ve.ScaleBy(-1);
                ve.ScaleBy(len); shPt.TranslateBy(ve);

                var dim = sh.DrawingNotes.HoleThreadNotes.Add(shPt, gi, false, st);
                dim.ArrowheadsInside = false; dim.SingleDimensionLine = false;
                dim.FirstArrowheadType = ArrowheadTypeEnum.kFilledArrowheadType;
                dim.SecondArrowheadType = ArrowheadTypeEnum.kFilledArrowheadType;
                if (dimText != "") dim.Text.FormattedText = dimText;
            }
            else if (t == ObjectTypeEnum.kAngularGeneralDimensionObject)
            {
                var dim = sh.DrawingDimensions.GeneralDimensions.AddAngular(pt, gi, dc2.getInt(), DimensionStyle: st);
                dim.CenterText();
                var arc = dim.DimensionLine as Arc2d;
                var dist = dim.Text.Origin.VectorTo(arc.Center).Length;
                var ve = dir.AsVector();
                ve.ScaleBy(len - dist);
                pt = dim.Text.Origin.Copy(); pt.TranslateBy(ve);
                dim.Text.Origin = pt;
                if (dimText != "") dim.Text.FormattedText = dimText;
            }
        }
        public void check(LinearGeneralDimension ld)
        {
            LineSegment2d seg = ld.DimensionLine as LineSegment2d;
            var vec = ld.Text.RangeBox.MinPoint.VectorTo(ld.Text.RangeBox.MaxPoint);
            var di = vec.X <= vec.Y ? vec.Y : vec.X;
            var di2 = getDist(seg);
            di *= 2;
            if (di2 <= di)
            {
                vec = seg.Direction.AsVector();
                vec.ScaleBy(di);
                var pt = ld.Text.Origin.Copy(); pt.TranslateBy(vec);
                ld.Text.Origin = pt;
            }
        }
        public double getDist(LineSegment2d seg)
        {
            return seg.StartPoint.VectorTo(seg.EndPoint).Length;
        }
        public void addLin()
        {
            p1 = addPoint("one", true);
            p2 = addPoint("two", true);
            LinearGeneralDimension gd = d as LinearGeneralDimension;
            al = gd.DimensionType;
            txt = d.Text.Origin;
        }
        public void addRad()
        {
            p1 = addPoint("line", false);
            if (t == ObjectTypeEnum.kHoleThreadNoteObject) txt = ht.Text.Origin;
            else txt = d.Text.Origin;
            len = p1.VectorTo(txt).Length;
        }
        public void addAngle()
        {
            p1 = addPoint("one", true);
            p2 = addPoint("two", true);
            AngularGeneralDimension gd = d as AngularGeneralDimension;
            al = gd.DimensionType;
            txt = d.Text.Origin;
            Arc2d arc = gd.DimensionLine as Arc2d;
            var v = arc.Center.VectorTo(txt);
            len = v.Length;
            dir = v.AsUnitVector();
        }
        public Point2d addPoint(string ext, bool start)
        {
            dynamic ent = d;
            LineSegment2d seg = null;
            if (t == ObjectTypeEnum.kHoleThreadNoteObject) ent = ht;
            switch (ext)
            {
                case "one":
                    try
                    {
                        seg = ent.ExtensionLineOne;
                        if (seg == null) throw new Exception();
                    }
                    catch (Exception)
                    {
                        return ent.IntentOne.PointOnSheet;
                    }
                    break;
                case "two":
                    try
                    {
                        seg = ent.ExtensionLineTwo;
                        if (seg == null) throw new Exception();
                    }
                    catch (Exception)
                    {
                        return ent.IntentTwo.PointOnSheet;
                    }
                    break;
                case "line":
                    seg = ent.DimensionLine;
                    break;
                default:
                    break;
            }
            if (dir == null) dir = seg.Direction;
            if (len == 0) len = getDist(seg);
            return start ? seg.StartPoint: seg.EndPoint;
        }
        public LineSegment2d getSeg(GeometryIntent gi)
        {
            var dc = gi.Geometry as DrawingCurve;
            return dc.Segments[1].Geometry as LineSegment2d;
        }
        public void getCenter(object dim)
        {
            dynamic ent = dim;
            DrawingCurve dc = null;
            if (t == ObjectTypeEnum.kLinearGeneralDimensionObject ||
                t == ObjectTypeEnum.kAngularGeneralDimensionObject)
                dc = ent.IntentOne.Geometry;
            else if (t == ObjectTypeEnum.kRadiusGeneralDimensionObject ||
                t == ObjectTypeEnum.kDiameterGeneralDimensionObject ||
                t == ObjectTypeEnum.kHoleThreadNoteObject)
                dc = ent.Intent.Geometry;
            dv = dc.Parent;
            dvnum = getNum(sh.DrawingViews, dv);
            shnum = getNum(((DrawingDocument)sh.Parent).Sheets, sh);
            fillDV();
        }
        public int getNum(System.Collections.IEnumerable ie, object ob)
        {
            int i = 1;
            foreach (var item in ie)
            {
                if (item.Equals(ob))
                {
                    return i;
                }
                i++;
            }
            return 0;
        }
        public void fillDV()
        {
            origin = dv.Center;
            sc = dv.Scale;
            w = dv.Width / sc;
            h = dv.Height / sc;
        }
        public void scale()
        {
            if (p1 != null)
                vecs.Add(origin.VectorTo(p1));
            if (p2 != null)
                vecs.Add(origin.VectorTo(p2));
            if (txt != null)
                vecs.Add(origin.VectorTo(txt));
            foreach (var item in vecs)
            {
                transform(item);
            }
        }
        public Vector2d scale(Point2d pt)
        {
            var v = origin.VectorTo(pt);
            transform(v);
            return v;
        }
        public void transform(Vector2d item)
        {
            item.ScaleBy(1 / sc);
            item.X /= w / 2;
            item.Y /= h / 2;
        }
        public DCurve findDC(Vector2d v, UnitVector2d n, Vector2d v1 = null)
        {
            IEnumerable<DrawingCurve> lst = null;
            if (t == ObjectTypeEnum.kLinearGeneralDimensionObject ||
                t == ObjectTypeEnum.kAngularGeneralDimensionObject)
                lst = u.gets<DrawingCurve>(dv.DrawingCurves, f => f.ProjectedCurveType == Curve2dTypeEnum.kLineSegmentCurve2d);
            else if (t == ObjectTypeEnum.kRadiusGeneralDimensionObject)
                lst = u.gets<DrawingCurve>(dv.DrawingCurves, f => f.ProjectedCurveType == Curve2dTypeEnum.kCircularArcCurve2d);
            else if (t == ObjectTypeEnum.kDiameterGeneralDimensionObject ||
                t == ObjectTypeEnum.kHoleThreadNoteObject)
                lst = u.gets<DrawingCurve>(dv.DrawingCurves, f => f.ProjectedCurveType == Curve2dTypeEnum.kCircleCurve2d);
            List<DCurve> cvs = new List<DCurve>(); 
            foreach (var item in lst)
            {               
                if (t == ObjectTypeEnum.kLinearGeneralDimensionObject)
                {
                    var vec = item.StartPoint.VectorTo(item.EndPoint);
                    if (!vec.IsParallelTo(dir.AsVector(), 0.01)) continue;
                }
                else if (t == ObjectTypeEnum.kAngularGeneralDimensionObject)
                {
                    var vec = item.StartPoint.VectorTo(item.EndPoint);
                    if (v1 != null && vec.IsParallelTo(v1, 0.01)) continue;
                }
                DCurve c = new DCurve(item);
                c.add(v, n);
                cvs.Add(c);
            }
            return cvs.OrderBy(el => el.d).First();
        }
        public DimensionStyle getStyle()
        {
            if (stName == "")
            {
                dynamic ent = d;
                if (t == ObjectTypeEnum.kHoleThreadNoteObject) ent = ht;
                DimensionStyle st = ent.Style;
                var sm = st.Parent as DrawingStylesManager;
                //if (st.Equals(sm.DimensionStyles[1])) return null;
                if (t != ObjectTypeEnum.kHoleThreadNoteObject)
                    if (d.Text.FormattedText != "<DimensionValue/>") dimText = d.Text.FormattedText;
                stName = st.Name;
            }
            else
            {
                var doc = sh.Parent as DrawingDocument;
                return u.get<DimensionStyle>(doc.StylesManager.DimensionStyles, f => f.Name == stName);
            }
            return null;
        }
        public XElement getXML()
        {
            XElement el = new XElement("dim");
            MyXML.addAtt(el, "type", t.ToString());
            MyXML.addAtt(el, "align", al.ToString());
            MyXML.addAtt(el, "dv", dvnum.ToString());
            MyXML.addAtt(el, "sheet", shnum.ToString());
            MyXML.addAtt(el, "dirx", dir.X.ToString("0.000"));
            MyXML.addAtt(el, "diry", dir.Y.ToString("0.000"));
            MyXML.addAtt(el, "len", len.ToString("0.000"));
            getStyle();
            MyXML.addAtt(el, "style", stName);
            MyXML.addAtt(el, "dimTxt", dimText);
            
            foreach (var item in vecs)
            {
                var e = MyXML.addXElement("pt", new Dictionary<string, string>() {
                { "x", item.X.ToString("0.000")}, { "y", item.Y.ToString("0.000")}});
                el.Add(e);
            }
            return el;
        }
        public void readXML(XElement el)
        {
            Enum.TryParse(MyXML.getAtt(el, "align"), out al);
            Enum.TryParse(MyXML.getAtt(el, "type"), out t);
            dvnum = int.Parse(MyXML.getAtt(el, "dv"));
            shnum = int.Parse(MyXML.getAtt(el, "sheet"));
            var s = ((DrawingDocument)sh.Parent).Sheets[shnum];
            if (!s.Equals(this.sh)) this.sh = s;
            double x = 0, y = 0;
            MyXML.getDouble(el, "dirx", ref x, 1);
            MyXML.getDouble(el, "diry", ref y, 1);
            MyXML.getDouble(el, "sc", ref sc1, 1);
            MyXML.getDouble(el, "len", ref len, 1);
            stName = MyXML.getAtt(el, "style");
            dimText = MyXML.getAtt(el, "dimTxt");

            dir = I.CUV2d(x, y);
            n = dir.Copy();
            u.normal(n, null);
            foreach (var item in el.Elements())
            {
                MyXML.getDouble(item, "x", ref x, 1);
                MyXML.getDouble(item, "y", ref y, 1);
                vecs.Add(I.CV2d(x, y));
            }
            dv = sh.DrawingViews[dvnum];
            fillDV();
            w = dv.Width; h = dv.Height;
        }
    }
}
