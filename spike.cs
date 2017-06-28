using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Inventor;
using InvDoc;

namespace InvAddIn
{
    public partial class spike : Form
    {
        private Inventor.Point2d pt1, pt2, pt3, pt4, cen, arcStart, arcEnd, textPt, startPt, endPt;
        private Transaction tr;
        private SketchLine line1, line2, line3, line4, line5, line6, line7, line8;
        private SketchArc arc; Edge endEdge;
        public double R, L, H;
        double startL, endL, length, l1, sum, le;
        int rev;
        Edge edge;
        double[] lengths;
        public UnitVector vec = null; bool norm;
        public Inventor.Application invApp;
        public PlanarSketch ps;
        public SheetMetalComponentDefinition smcd;
        public SketchLine oldLin, newLin, startLin, endLin;
        private SymmetryConstraint sym; TangentSketchConstraint tg; SketchArc oldArc; SketchLine oldLine, oldLine2;
        char separator = System.Globalization.CultureInfo.CurrentCulture.NumberFormat.CurrencyDecimalSeparator[0];
        List<Inventor.Vertex> vertex; List<Inventor.Point> pts;
        List<spickesData> sData = new List<spickesData>();
        public spike(Inventor.Application app)
        {
            invApp = app;
        }
        public spike(Inventor.Application app, PlanarSketch sketch)
        {
            invApp = app;
            ps = sketch;
            InitializeComponent();
        }

        private Edge selectFromSketch()
        {
            CommandManager cmdMgr = invApp.CommandManager;
            Edge edge = (Edge)cmdMgr.Pick(SelectionFilterEnum.kPartEdgeFilter, "Выберите ребро:");
            return edge;
        }

        public double convToDouble(string val)
        {
            return Convert.ToDouble(val.Replace(',', separator));
        }

        public bool colin (UnitVector2d v1, UnitVector2d v2, int tolerance)
        {
            int x1 = (int)(v1.X * tolerance), x2 = (int)(v2.X * tolerance), y1 = (int)(v1.Y * tolerance), y2 = (int)(v2.Y * tolerance);
            if (x1 == x2 && y1 == y2) return false;
            if (x1 != 0 && y1 != 0 && x2 != 0 && y2 != 0)
                return x1 / x2 == y1 / y2;
            else
            {
                int n = (x2 != 0) ? x1 / x2 : y1 / y2;
                return (n * x1 == x2 && n * y1 == y2) ? true : false;
            }
        }

        public bool colin(Point2d pt1, Point2d pt2, Point2d pt)
        {
            if (pt.IsEqualTo(pt1) || pt.IsEqualTo(pt2)) return false;
            Vector2d v1 = pt1.VectorTo(pt2);
            Vector2d v2 = pt.VectorTo(pt1);
            return v1.IsParallelTo(v2);
        }

        public void addSpikeDef(SheetMetalComponentDefinition smcd, double R, Double L, Double H, string name = "Прямой шип")
        {
            try
            {
                SketchBlockDefinition d = smcd.SketchBlockDefinitions[name];
                return;
            }
            catch (Exception)
            {
            }
            SketchBlockDefinition def = smcd.SketchBlockDefinitions.Add(name);
            TransientGeometry tg = invApp.TransientGeometry;
            Point2d pt = tg.CreatePoint2d();
            SketchPoint origin = def.SketchPoints.Add(pt,false);
            def.GeometricConstraints.AddGround((SketchEntity)origin);
            Point2d center = tg.CreatePoint2d(pt.X,pt.Y + H-R);
            SketchLine centerLine = def.SketchLines.AddByTwoPoints(origin, center);
            centerLine.Construction = true;
            centerLine.Centerline = true;
            SketchLine leftLine = def.SketchLines.AddByTwoPoints(tg.CreatePoint2d(pt.X - L / 2, pt.Y),tg.CreatePoint2d(pt.X-R,center.Y));
            SketchLine rightLine = def.SketchLines.AddByTwoPoints(tg.CreatePoint2d(pt.X + L / 2, pt.Y), tg.CreatePoint2d(pt.X + R, center.Y));
            SketchLine bottomLine = def.SketchLines.AddByTwoPoints(leftLine.StartSketchPoint, rightLine.StartSketchPoint);
            SketchArc arc = def.SketchArcs.AddByCenterStartEndPoint(centerLine.EndSketchPoint, rightLine.EndSketchPoint, leftLine.EndSketchPoint);
            //def.GeometricConstraints.AddCoincident((SketchEntity)arc.CenterSketchPoint, (SketchEntity)centerLine.EndSketchPoint);
            bottomLine.Construction = true;
            def.GeometricConstraints.AddMidpoint(origin, bottomLine);
            def.GeometricConstraints.AddHorizontal((SketchEntity)bottomLine);
            def.GeometricConstraints.AddPerpendicular((SketchEntity)centerLine, (SketchEntity)bottomLine);
            def.GeometricConstraints.AddSymmetry((SketchEntity)leftLine, (SketchEntity)rightLine, centerLine);
            def.GeometricConstraints.AddTangent((SketchEntity)arc, (SketchEntity)leftLine);
            def.GeometricConstraints.AddTangent((SketchEntity)arc, (SketchEntity)rightLine);
            def.GeometricConstraints.AddCoincident((SketchEntity)centerLine.EndSketchPoint, (SketchEntity)arc.CenterSketchPoint);
            RadiusDimConstraint rad = def.DimensionConstraints.AddRadius((SketchEntity)arc, arc.Geometry.Center);
            rad.Parameter.Value = R; 
            TwoPointDistanceDimConstraint dist = def.DimensionConstraints.AddTwoPointDistance(origin, arc.CenterSketchPoint, DimensionOrientationEnum.kVerticalDim,arc.Geometry.EndPoint);
            dist.Parameter.Value = H - R;
            dist = def.DimensionConstraints.AddTwoPointDistance(bottomLine.StartSketchPoint, bottomLine.EndSketchPoint, DimensionOrientationEnum.kHorizontalDim, bottomLine.Geometry.EndPoint);
            dist.Parameter.Value = L;
            foreach (DimensionConstraint item in def.DimensionConstraints)
            {
                item.Delete();
            }
        }

        public Edge findEdge(SheetMetalComponentDefinition smcd, ref UnitVector vec)
        {
            SurfaceBody body = smcd.SurfaceBodies[1];
            Face face = body.Faces.OfType<Face>().FirstOrDefault(f => f.SurfaceType == SurfaceTypeEnum.kCylinderSurface && f.TangentiallyConnectedFaces.Count != 0);
            if (face == null) return null;
            Cylinder cyl = face.Geometry as Cylinder;
            UnitVector vec1 = cyl.AxisVector;
            InvDoc.u.abs(ref vec1);
            vec = vec1;
            Edge ed = face.Edges.OfType<Edge>().Where(e => e.TangentiallyConnectedEdges.Count > 1).OrderBy(el => direct(vec1, el)).Last();
            Double thick = (double)smcd.Thickness.Value;
            thick = Math.Round(thick, 3);
            ed = reverse(ed, thick);
            ed = isUse(ed, vec1);
            return ed;
        }

        public Edge reverse(Edge ed, double thick)
        {
            Edge e1 = ed.TangentiallyConnectedEdges[1] as Edge, e2 = ed.TangentiallyConnectedEdges[2] as Edge;
            if (InvDoc.u.eq(e1.StartVertex, e2.StopVertex))
            {
                ed = changeEdge(ed, thick);
            }
            return ed;
        }

        public Edge changeEdge(Edge ed, double thick)
        {
            Face f = ed.Faces.OfType<Face>().OrderBy(a => a.Evaluator.Area).First();
            foreach (Edge e in f.Edges)
            {
                if (!e.Equals(ed) && InvDoc.u.getLenght(e) != thick)
                {
                    ed = e; break;
                }
            }
            return ed;
        }

        public Edge isUse(Edge ed, UnitVector vec)
        {
            EdgeCollection edges = ed.TangentiallyConnectedEdges;
            Edge edStart = null/* = edges[1] as Edge*/,
                edEnd = null/* = edges[edges.Count] as Edge*/;
            UnitVector v = I.tg.CreateUnitVector(vec.Y, vec.Z, vec.X);

                for (int i = 0; i < edges.Count - 1; i++)
                {
                    Edge ed1 = edges[i + 1] as Edge;
                    Edge ed2 = edges[i + 2] as Edge;
                    if (!(InvDoc.u.eq(ed1.StopVertex, ed2.StartVertex) || InvDoc.u.eq(ed1.StartVertex, ed2.StopVertex)))
                    {
                        edStart = ed1;
                        if (ed1.GeometryType == CurveTypeEnum.kLineSegmentCurve && ed2.GeometryType == CurveTypeEnum.kLineSegmentCurve)
                        {
                            if (i - 1 == 0) edEnd = edges[edges.Count] as Edge;
                            else edEnd = edges[i - 1] as Edge;
                        }
                        else
                        {
                            edEnd = ed2;
                        }
                    }
                }
            
                double d1 = direct(v, edEnd.StopVertex), d2 = direct(v, edStart.StartVertex);

                if (vec.Y == 1) return (d1 > d2) ? edEnd : edStart;
                else return (d1 < d2) ? edEnd : edStart;
        }

        public double direct(UnitVector vec, Vertex vert)
        {
            Inventor.Point pt = vert.Point;
            double ret = vec.AsVector().DotProduct(I.tg.CreateVector(pt.X, pt.Y, pt.Z));
            return ret;
        }

        public double direct(UnitVector vec, Edge ed)
        {
            Inventor.Point pt = ed.StartVertex.Point;
            double ret = vec.AsVector().DotProduct(I.tg.CreateVector(pt.X, pt.Y, pt.Z));
            return ret;
        }

        public void addSketch(SheetMetalComponentDefinition smcd, double H, double R, double L)
        {
            edge = findEdge(smcd, ref vec);
            Face face = edge.Faces.OfType<Face>().OrderBy(f => f.Evaluator.Area).Last();
            ps = smcd.Sketches.Add(face);
            ps.Name = "Шип";
            lengths = new double[edge.TangentiallyConnectedEdges.Count];
            addLenght(edge, ref lengths);
            addSData(lengths);
            fillSData();
            spikesForm();
        }

        public SketchPoint getPoint(Point2d insPt, SketchLine sl, bool min = false)
        {
            double d1 = insPt.DistanceTo(sl.StartSketchPoint.Geometry), d2 = insPt.DistanceTo(sl.EndSketchPoint.Geometry);
            if (min)
            {
                return d1 < d2 ? sl.StartSketchPoint : sl.EndSketchPoint;
            }
            else
            {
                return d1 > d2 ? sl.StartSketchPoint : sl.EndSketchPoint;
            }
        }

        public void spikesForm()
        {
            int offsetY = 30, offsetX = 10;
            Form f = new Form();
            f.Height = 150 + sData.Count * 15; f.Width = 400; f.WindowState = FormWindowState.Normal; f.Text = "Шипы"; f.StartPosition = FormStartPosition.CenterScreen;
            f.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink; f.AutoSize = true;
            System.Drawing.Point insPt = new System.Drawing.Point(5, 5); int i = 1;
            InterfaceDll.CB cbs;
            InterfaceDll.Lbl lbs;
            lbs = new InterfaceDll.Lbl(offsetX, offsetY, 100, 15, insPt, f, "Участок ");
            cbs = new InterfaceDll.CB(offsetX, offsetY, 200, 15, insPt, f);
            foreach (spickesData d in sData)
            {
                cbs.position(lbs.last(), true);
                lbs.add("Участок " + i, Y: true);
                cbs.add(d.count.ToString(), new[] { "0", "1", "2", "3", "4", "5" }, Y: true);
                i++; 
            }

            InterfaceDll.Btn btns = new InterfaceDll.Btn(offsetX, offsetY, 100, 20, insPt, f, spike_Click, "Добавить");
            btns.center(cbs.last(), offsetY + 5);
            f.Show();
        }

        private void spike_Click(object arg1, EventArgs arg2)
        {
            norm = false;
            Form f = (arg1 as System.Windows.Forms.Button).Parent as Form;
            for (int i = 0; i < sData.Count; i++)
            {
                spickesData sd = sData[i];
                int c = int.Parse(f.Controls.OfType<ComboBox>().ElementAt(i + 1).Text);
                sData[i].count = c;
            }
            start();
            next();
        }

        public void start()
        {
            Inventor.Point fpt = edge.StopVertex.Point;
            Vector2d dir = null;
            SketchLine projectSL = (SketchLine)ps.AddByProjectingEntity(edge);
            SketchLine arc = (SketchLine)ps.AddByProjectingEntity(edge.TangentiallyConnectedEdges[2]);
            projectSL.Construction = true;
            arc.Construction = true;
            dir = projectSL.StartSketchPoint.Geometry.VectorTo(arc.EndSketchPoint.Geometry);
            //dir.ScaleBy(-1);
            if (dir.Length == 0) dir = projectSL.EndSketchPoint.Geometry.VectorTo(arc.EndSketchPoint.Geometry);
            if (dir.Y > 0 || dir.X > 0) norm = true;
            Vector2d normal;// = getNormal(centr).AsVector();
            Point2d pt = I.tg.CreatePoint2d();// = centr.Geometry.MidPoint;
            double le = 0;
            normal = getNormal(projectSL, dir.AsUnitVector()).AsVector();
            normal.ScaleBy(H);
            if (norm) normal.ScaleBy(-1);
            if (InvDoc.u.eq(projectSL.StartSketchPoint.Geometry, arc.StartSketchPoint.Geometry) || InvDoc.u.eq(projectSL.StartSketchPoint.Geometry, arc.EndSketchPoint.Geometry))
            {
                pt = projectSL.EndSketchPoint.Geometry;
            }
            else
            {
                pt = projectSL.StartSketchPoint.Geometry;
            }
            pt.TranslateBy(normal);
            bool last = false;
            List<int> except = new List<int>(1);
            XMLDoc xmlDoc;
            if (System.IO.File.Exists(@"C:\ProgramData\Autodesk\Inventor Addins\spike.xml"))
            {
                xmlDoc = new XMLDoc(@"C:\ProgramData\Autodesk\Inventor Addins\spike.xml", "head");
                foreach (var el in xmlDoc.Doc.Root.Element("except").Elements())
                {
                    except.Add((int)(double.Parse(el.Value.Replace('.', ',')) * 10));
                }
            }
            else { except.Add(117); }
            int count = 1;
            if (except.Exists(e => e == Math.Round(lengths[lengths.Length - 1], 2) * 100)) count = 3;
            for (int i = 0; i < lengths.Length; i++)
            {
                if (i == lengths.Length - count)
                    last = true;
                if (lengths[i] > 0.7 && !except.Exists(e => e == Math.Round(lengths[i], 2) * 100))
                {
                    spickesData sd = sData.Find(s => InvDoc.u.eq(s.l, lengths[i]));
                    if (sd.count != 0)
                    addSketchBlock(le, lengths[i], pt, projectSL, last, dir, norm, sd);
                }
                le += lengths[i];
            }
            SketchLine newLin = ps.SketchLines.OfType<SketchLine>().LastOrDefault(l => l.Construction == true);
            newLin = addLine(getPoint(pt, newLin, true), getPoint(pt, newLin), 0, le / 2);
            if (!(InvDoc.u.eq(projectSL.StartSketchPoint.Geometry, arc.StartSketchPoint.Geometry) || InvDoc.u.eq(projectSL.StartSketchPoint.Geometry, arc.EndSketchPoint.Geometry)))
            {
                endLin = addLine(getPoint(pt, newLin, true), getPoint(pt, newLin), Math.PI / 2, H + 1);
            }
            else
            {
                endLin = addLine(getPoint(pt, newLin, true), getPoint(pt, newLin), -Math.PI / 2, H + 1);
            }
            ps.GeometricConstraints.AddParallel((SketchEntity)newLin, (SketchEntity)projectSL);
            ps.GeometricConstraints.AddPerpendicular((SketchEntity)newLin, (SketchEntity)endLin);
            ps.DimensionConstraints.AddTwoPointDistance(projectSL.StartSketchPoint, newLin.EndSketchPoint, DimensionOrientationEnum.kHorizontalDim,
                InvAddIn.Drawings.midPoint(projectSL.StartSketchPoint.Geometry, newLin.EndSketchPoint.Geometry, 0));
            if (endLin != null && startLin != null)
            {
                //dir.ScaleBy(-1);
                //move(startLin, dir, le / 2);
                addLine(startLin.EndSketchPoint, endLin.EndSketchPoint, 0, 0);
            }
        }

        public void move(SketchLine sl, Vector2d v, double l)
        {
            v.Normalize(); v.ScaleBy(l); 
            SketchPoint sp = sl.StartSketchPoint;
            sp.MoveBy(v);
            sp = sl.EndSketchPoint;
            sp.MoveBy(v);
        }

        public void next()
        {
            foreach (SketchBlock item in ps.SketchBlocks)
            {
                item.Explode();
            }
            SketchLine newLin = ps.SketchLines.OfType<SketchLine>().FirstOrDefault(l => (int)(l.Length * 1000) == (int)(L * 1000));
            ps.DimensionConstraints.AddTwoPointDistance(newLin.StartSketchPoint, newLin.EndSketchPoint, DimensionOrientationEnum.kAlignedDim, Drawings.midPoint(newLin.StartSketchPoint.Geometry, newLin.EndSketchPoint.Geometry, 0));
            foreach (SketchLine item in ps.SketchLines.OfType<SketchLine>().Where(l => (int)(l.Length * 1000) == (int)(L * 1000)))
            {
                if (item.Equals(newLin) || item.Constraints.OfType<EqualLengthConstraint>().Count() != 0) continue;
                ps.GeometricConstraints.AddEqualLength(item, newLin);
            }
            SketchArc sa = ps.SketchArcs.OfType<SketchArc>().First();
            foreach (SketchArc item in ps.SketchArcs)
            {
                if (item.Equals(sa) || item.Constraints.OfType<EqualRadiusConstraint>().Count() != 0) continue;
                ps.GeometricConstraints.AddEqualRadius((SketchEntity)item, (SketchEntity)sa);
            }
            SketchLine oldLine = null;
            newLin = ps.SketchLines.OfType<SketchLine>().First(l => l.Centerline == true);
            SketchLine startLine = ps.SketchLines.OfType<SketchLine>().First(l => l.Construction == true);
            ps.DimensionConstraints.AddTwoPointDistance(newLin.StartSketchPoint, newLin.EndSketchPoint, DimensionOrientationEnum.kAlignedDim,
                Drawings.midPoint(newLin.StartSketchPoint.Geometry, newLin.EndSketchPoint.Geometry, 0));
            var slines = ps.SketchLines.OfType<SketchLine>().Where(l => l.Centerline == true);
            double offset = 0;
            foreach (SketchLine l in slines)
            {
                if (l.Equals(newLin))
                {
                    oldLine = l;
                    continue;
                }
                else
                {
                    Point2d pt = oldLine.EndSketchPoint.Geometry;
                    offset += 5;
                    ps.GeometricConstraints.AddEqualLength(l, newLin);
                    ps.DimensionConstraints.AddTwoPointDistance(l.StartSketchPoint, startLine.StartSketchPoint, DimensionOrientationEnum.kVerticalDim,
                      Drawings.midPoint(pt, l.EndSketchPoint.Geometry, offset));
                }
                oldLine = l;
            }
            try
            {
                WorkPlane wp = smcd.WorkPlanes["Шип_справа"];
            }
            catch (Exception)
            {
                CutFeature cut = Offset.addCut(smcd, "Шип"); cut.Name = "Шипы";
                spike.addPlane(smcd, vec);
                Offset.project(smcd, "Шип_проекция");
                ObjectCollection objs = Macros.StandardAddInServer.m_inventorApplication.TransientObjects.CreateObjectCollection();
                objs.Add(cut);
                Offset.addMirror(smcd, objs, "Шип_слева", vec);
            }
        }

        public class spickesData
        {
            public double l;
            public int count;
            public spickesData(double l, int c)
            {
                this.l = l; count = c;
            }
        }

        public void addSData(double[] lengths)
        {
            List<int> except = new List<int>(1);
            XMLDoc xmlDoc; bool last = false;
            if (System.IO.File.Exists(@"C:\ProgramData\Autodesk\Inventor Addins\spike.xml"))
            {
                xmlDoc = new XMLDoc(@"C:\ProgramData\Autodesk\Inventor Addins\spike.xml", "head");
                foreach (var el in xmlDoc.Doc.Root.Element("except").Elements())
                {
                    except.Add((int)(double.Parse(el.Value.Replace('.', ',')) * 10));
                }
            }
            else { except.Add(117); }
            int count = 1;
            if (except.Exists(e => e == Math.Round(lengths[lengths.Length - 1], 2) * 100)) count = 3;
            for (int i = 0; i < lengths.Length; i++)
            {
                if (i == lengths.Length - count)
                    last = true;
                if (lengths[i] > 0.7 && !except.Exists(e => e == Math.Round(lengths[i], 2) * 100))
                {
                    sData.Add(new spickesData(lengths[i], 1));
                }
            }
        }

        public void fillSData()
        {
            XMLDoc xmlDoc;
            for (int i = 0; i < sData.Count; i++)
			{
                spickesData sd = sData[i];
                if (System.IO.File.Exists(@"C:\ProgramData\Autodesk\Inventor Addins\spike.xml"))
                {
                    xmlDoc = new XMLDoc(@"C:\ProgramData\Autodesk\Inventor Addins\spike.xml", "head");
                    foreach (var el in xmlDoc.Doc.Root.Element("spikes").Elements())
                    {
                        if (sd.l > double.Parse(el.Attribute("min").Value.Replace('.', ',')) * 0.1 && sd.l <= double.Parse(el.Attribute("max").Value.Replace('.', ',')) * 0.1)
                        {
                            sd.count = int.Parse(el.Attribute("count").Value.Replace('.', ','));
                        }
                    }
                    foreach (var ex in xmlDoc.Doc.Root.Element("except").Elements())
                    {
                        double d = double.Parse(ex.Value.Replace('.', ','));
                        if (InvDoc.u.eq(d, sd.l * 10))
                        {
                            sd.count = 0;
                        }
                    }
                }
                else
                {
                    if (le > 0.7 && le <= 5)
                    {
                        sd.count = 1;
                    }
                    if (le > 5 && le <= 15)
                    {
                        sd.count = 2;
                    }
                    if (le > 15)
                    {
                        sd.count = 3;
                    }
                }
            }
        }

        public void addSketchBlock(double lenStart, double le, Point2d insPt, SketchLine projectSL, bool last, Vector2d dir, bool norm, spickesData sd)
        {
            dir.Normalize();
            if (lenStart != 0)
            dir.ScaleBy(lenStart);
            Point2d pt = invApp.TransientGeometry.CreatePoint2d(insPt.X, insPt.Y);
            Point2d newPt = pt.Copy();
            if (dir.Length != 1)
            pt.TranslateBy(dir);
            double offset = 0;
            offset = sd.l / sd.count;
            dir.Normalize();
            dir.ScaleBy(Math.Round(offset,4));
            Vector2d dir2 = dir.Copy();
            dir2.ScaleBy(0.5);
            for (int i = 0; i < sd.count; i++)
            {
                if (i != 0)
                    pt.TranslateBy(dir);
                else pt.TranslateBy(dir2);
                SketchBlock sb;
                if (ps.SketchBlocks.Count == 0 || (last && i == sd.count-1))
                sb = ps.SketchBlocks.AddByDefinition(smcd.SketchBlockDefinitions["Прямой шип"], pt);
                else
                sb = ps.SketchBlocks.AddByDefinition(smcd.SketchBlockDefinitions["Капелька"], pt);
                newLin = ps.SketchLines.OfType<SketchLine>().LastOrDefault(l => l.Construction == true);
                SketchPoint spt = null;
                if (!norm)
                {
                    if (dir.X > 0 || dir.Y > 0)
                    {
                        rotateVector(dir, newLin, Math.PI, ps);
                        spt = getPoint(insPt, newLin, true);
                    }
                    else
                    {
                        rotateVector(dir, newLin, 0, ps);
                        spt = getPoint(insPt, newLin, true);
                    }
                }
                else
                {
                    if (dir.X < 0 || dir.Y < 0)
                    {
                        rotateVector(dir, newLin, Math.PI, ps);
                        spt = getPoint(insPt, newLin, true);
                    }
                    else
                    {
                        rotateVector(dir, newLin, 0, ps);
                        spt = getPoint(insPt, newLin, true);
                    }
                }
                SketchLine sl = ps.SketchLines.OfType<SketchLine>().LastOrDefault(l => l.Centerline == true);
                Vector2d n = sl.StartSketchPoint.Geometry.VectorTo(sl.EndSketchPoint.Geometry);
                if (oldLin != null)
                {   
                    SketchLine l;
                    l = addLine(getPoint(insPt, oldLin), spt);
                    ps.GeometricConstraints.AddCollinear((SketchEntity)oldLin, (SketchEntity)newLin);
                }
                else
                {
                    Vector2d vec = dir.Copy(); /*rotate(Math.PI, pt, dir); */     
                    SketchLine l;
                    vec.ScaleBy(-le/2);
                    newPt.TranslateBy(vec);
                    vec.Normalize();
                    l = addLine(newPt, getPoint(insPt, newLin, true), 0, 2*(vec.X+vec.Y));
                    ps.GeometricConstraints.AddParallel((SketchEntity)l, (SketchEntity)newLin);
                    ps.GeometricConstraints.AddTangent((SketchEntity)ps.SketchArcs[1], (SketchEntity) projectSL);
                    ps.GeometricConstraints.AddParallel((SketchEntity)newLin, (SketchEntity)projectSL);
                    startLin = addLine(l.EndSketchPoint, l.EndSketchPoint, n, H + 1);
//                     if (dir.X > 0 || dir.Y > 0)
//                     startLin = addLine(getPoint(insPt, l, true), getPoint(insPt, l), Math.PI / 2, H + 1);
//                     else
//                         startLin = addLine(getPoint(insPt, l, true), getPoint(insPt, l), -Math.PI / 2, H + 1);
                }
                oldLin = newLin;
            }
        }

        public SketchLine addLine(SketchPoint startPt, SketchPoint endPoint, Vector2d vec, double l)
        {
            SketchLine sl = null;
            try
            {
                Point2d newPt = (startPt.Geometry).Copy();
                vec.Normalize();
                vec.ScaleBy(l);
                newPt.TranslateBy(vec);
                sl = ps.SketchLines.AddByTwoPoints(endPoint, newPt);
                ps.GeometricConstraints.AddCoincident((SketchEntity)sl.StartSketchPoint, (SketchEntity)endPoint);
            }
            catch (Exception)
            {
                return sl;
            }
            return sl;
        }

        public SketchLine addLine(SketchPoint startPt, SketchPoint endPoint, double angl = 0, double l = 0)
        {
            SketchLine sl = null;
            try
            {
            if (l == 0)
            {
                sl = ps.SketchLines.AddByTwoPoints(startPt, endPoint);
                ps.GeometricConstraints.AddCoincident((SketchEntity)startPt, (SketchEntity)sl.StartSketchPoint);
                ps.GeometricConstraints.AddCoincident((SketchEntity)endPoint, (SketchEntity)sl.EndSketchPoint);
            }
            else
            {
                Vector2d vec = rotate(angl, endPoint.Geometry, (startPt.Geometry).VectorTo(endPoint.Geometry));
                vec.Normalize();
                vec.ScaleBy(l);
                Point2d newPt = (startPt.Geometry).Copy();
                if (angl != 0) newPt = (endPoint.Geometry).Copy();
                newPt.TranslateBy(vec);
                sl = ps.SketchLines.AddByTwoPoints(endPoint,newPt);
                ps.GeometricConstraints.AddCoincident((SketchEntity)sl.StartSketchPoint, (SketchEntity)endPoint);
            }

            }
            catch (Exception)
            {
                return sl;
            }
            return sl;
        }

        public SketchLine addLine(Point2d startPt, SketchPoint endPoint, double angl = 0, double l = 0)
        {
            SketchLine sl = null;
            try
            {
                if (l == 0)
                {
                    sl = ps.SketchLines.AddByTwoPoints(startPt, endPoint);
                    ps.GeometricConstraints.AddCoincident((SketchEntity)startPt, (SketchEntity)sl.StartSketchPoint);
                    ps.GeometricConstraints.AddCoincident((SketchEntity)endPoint, (SketchEntity)sl.EndSketchPoint);
                }
                else
                {
                    Vector2d vec = rotate(angl, endPoint.Geometry, (startPt).VectorTo(endPoint.Geometry));
                    vec.Normalize();
                    vec.ScaleBy(l);
                    Point2d newPt = (startPt).Copy();
                    if (angl != 0) newPt = (endPoint.Geometry).Copy();
                    newPt.TranslateBy(vec);
                    sl = ps.SketchLines.AddByTwoPoints(endPoint, newPt);
                    //if (angl != 0) return sl;
                    ps.GeometricConstraints.AddCoincident((SketchEntity)sl.StartSketchPoint, (SketchEntity)endPoint);
                }

            }
            catch (Exception)
            {
                return sl;
            }
            return sl;
        }

        public Vector2d rotate(double angl, Point2d pt, Vector2d origin)
        {
            Matrix2d mtx = invApp.TransientGeometry.CreateMatrix2d();
            mtx.SetToRotation(angl, pt);
            Vector2d vec = invApp.TransientGeometry.CreateVector2d(origin.X, origin.Y);
            vec.TransformBy(mtx);
            return vec;
        }

        public void addLenght(Edge e, ref double[] arr)
        {
            //bool flag = false;
            Edge ed = null, edComp = null; //Vector vec = null, vec1 = null; Vertex pt = null;
            double kFactor = double.Parse(smcd.UnfoldMethod.kFactor.Remove(5).Replace('.',',')), le;
            double thickness = Math.Round((double)smcd.Thickness.Value,3);
            for (int i = 1; i < e.TangentiallyConnectedEdges.Count+1; i++)
            {
                ed = e.TangentiallyConnectedEdges[i] as Edge;
                getLenght(ed);
                if (ed.GeometryType == CurveTypeEnum.kCircularArcCurve)
                {
                    edComp = changeEdge(ed, thickness);
                    if (InvDoc.u.getLenght(ed) > InvDoc.u.getLenght(edComp)) ed = edComp;
                    Arc3d arc = (Arc3d)ed.Geometry;
                    double r = arc.Radius;
                    double angl = arc.SweepAngle;
                    r += kFactor * thickness;
                    le = r * angl;
                    //vec1 = pt.Point.VectorTo(ed.StopVertex.Point);
                }
                else
                {
                    le = getLenght(ed);
                    //pt = ed.StartVertex;
                    //vec = ed.StartVertex.Point.VectorTo(ed.StopVertex.Point);
                }
//                 if (vec != null && vec1 != null)
//                 {
//                     
//                     double ang = vec.AngleTo(vec1);
//                     if (ang > 0) flag = true;
//                     else flag = false;
//                 }
                arr[i-1] = le;
            }
        }

        public double getLenght(Edge e)
        {
            double minParam, maxParam, length;
            CurveEvaluator eval = e.Evaluator;
            eval.GetParamExtents(out minParam, out maxParam);
            eval.GetLengthAtParam(minParam, maxParam, out length);
            return length;
        }

        public UnitVector2d getNormal(SketchLine l, UnitVector2d dir = null)
        {
            dir = dir??l.Geometry.Direction;
            //if (dir.Y > 0) dir = I.tg.CreateUnitVector2d(0, -1);
            Matrix2d mtx = invApp.TransientGeometry.CreateMatrix2d();
            //if (dir.Y > 0)
            mtx.SetToRotation(Math.PI / 2, l.Geometry.MidPoint);
            //else mtx.SetToRotation(-Math.PI/2, l.Geometry.MidPoint);
            dir.TransformBy(mtx);
            return dir;
        }

        static public void rotateVector(Vector2d vec, SketchLine sl, double pi, PlanarSketch ps)
        {
            //Matrix2d mtx = invApp.TransientGeometry.CreateMatrix2d();
            //mtx.SetToRotateTo(sl.Geometry.Direction.AsVector(), vec);
            double angl = sl.Geometry.Direction.AngleTo(vec.AsUnitVector()) + pi; 
            ObjectCollection col = Macros.StandardAddInServer.m_inventorApplication.TransientObjects.CreateObjectCollection();
            col.Add(ps.SketchBlocks[ps.SketchBlocks.Count]);
            ps.RotateSketchObjects(col, sl.Geometry.MidPoint, angl);
        }

        static public WorkPlane addPlane(SheetMetalComponentDefinition smcd, UnitVector vec)
        {
            string val = "0";
            SheetMetalFeatures smf = smcd.Features as SheetMetalFeatures;
            DistanceExtent pfe = smf.ContourFlangeFeatures[1].Definition.DefaultWidthExtent as DistanceExtent;
            val = pfe.Distance.Name + "/2 - 6.5";
            Parameter p = InvAddIn.CreateComponent.getParameter((Document)smcd.Document, "БВ_длина");
            if (p != null)
            {
            val = "БВ_длина/2 + 1";
            }
                WorkPlane wp = smcd.WorkPlanes.OfType<WorkPlane>().First(w => InvDoc.u.eq(vec, w.Plane.Normal));
                wp = smcd.WorkPlanes.AddByPlaneAndOffset(wp, val);
                wp.Name = "Шип_справа"; wp.Visible = false;
                //val = "-(БВ_длина/2 + 1)";
                //wp = smcd.WorkPlanes.AddByPlaneAndOffset(smcd.WorkPlanes[1], val);
                //wp.Name = "Шип_слева"; wp.Visible = false;
                return wp;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            this.Hide();
            try
            {
                tr = invApp.TransactionManager.StartTransaction((_Document)invApp.ActiveDocument, "Шипы");
                vertex = new List<Inventor.Vertex>();
                pts = new List<Inventor.Point>();
                Edge edge = selectFromSketch();
                bool flag = false;
                //if (edge.TangentiallyConnectedEdges.Count != 1)
                //{
                double kFactor = 0.44;
                smcd = (SheetMetalComponentDefinition)((PartDocument)invApp.ActiveDocument).ComponentDefinition;
                double thickness = (double)smcd.Thickness.Value;
                    foreach (Edge ed in edge.TangentiallyConnectedEdges)
                    {
                        double maxPar, minPar;
                        vertex.Add(ed.StartVertex);
                        vertex.Add(ed.StopVertex);
                        ed.Evaluator.GetParamExtents(out maxPar, out minPar);   
                        ed.Evaluator.GetLengthAtParam(maxPar,minPar,out le);
                        if (ed.GeometryType == CurveTypeEnum.kCircularArcCurve)
                        {
                            Arc3d arc = (Arc3d)ed.Geometry;
                            double r = arc.Radius;
                            double angl = arc.SweepAngle;
                            if ((int)(r * 1000) == 150)
                            {
                                r += kFactor * thickness;
                                flag = false;
                            }
                            else if (flag)
                            {
                                r += kFactor * thickness;
                            }
                            else
                            { 
                                r -= thickness - kFactor * thickness;
                                flag = true;
                            }
                            le = r * angl;
                        }
                        length += le;
                    }
                    pts = removeDub(vertex);
                    //endEdge = (Edge)edge.TangentiallyConnectedEdges[3];
                //}
                SketchEntity ent = ps.AddByProjectingEntity(edge);
                ent.Construction = true;
                SketchLine line = (SketchLine)ent;
                if ((int)(line.Geometry.StartPoint.Y*1000) == (int)(line.Geometry.EndPoint.Y*1000))
                    Macros.StandardAddInServer.Spike = "Вертикальные шипы";
                int count = Convert.ToInt16(textBox4.Text);
                R = convToDouble(textBox1.Text) / 10; L = convToDouble(textBox2.Text) / 10; H = convToDouble(textBox3.Text) / 10;
                startL = convToDouble(textBox5.Text) / 10; endL = convToDouble(textBox6.Text) / 10;
                //addSpike(smcd, invApp.TransientGeometry.CreatePoint2d(), invApp.TransientGeometry.CreateVector2d(1, 0), R, R*2, H);
                //addSpike(smcd, invApp.TransientGeometry.CreatePoint2d(), invApp.TransientGeometry.CreateVector2d(1, 0), R, L, H, "Капелька");
                return;

                double lineLength = length;
                length = length - startL - endL - R * 2; l1 = length / (count - 1); sum = 0;
                UnitVector2d dir = line.Geometry.Direction; 
                UnitVector2d normal = dir.Copy();
                Matrix2d mtx = invApp.TransientGeometry.CreateMatrix2d();
                mtx.SetToRotation(3.1415926 / 2, invApp.TransientGeometry.CreatePoint2d(0, 0));
                normal.TransformBy(mtx);
                Vector2d vecNormal = normal.AsVector(), vecDirect = dir.AsVector();
                vecDirect.ScaleBy(startL + R);
                //vecDirect.ScaleBy(-1);
                if (comparePoint(pts[0], line.StartSketchPoint.Geometry3d, 1000))
                {
                    startPt = line.StartSketchPoint.Geometry;
                    Vector2d lineVec = line.Geometry.StartPoint.VectorTo(line.Geometry.EndPoint);
                    lineVec.Normalize(); lineVec.ScaleBy(lineLength);
                    endPt = line.Geometry.StartPoint; endPt.TranslateBy(lineVec);
                    //rev = 1;
                }
                else if (comparePoint(pts[0], line.EndSketchPoint.Geometry3d, 1000))
                {
                    Vector2d lineVec = line.Geometry.EndPoint.VectorTo(line.Geometry.StartPoint);
                    lineVec.Normalize(); lineVec.ScaleBy(lineLength);
                    startPt = line.Geometry.EndPoint; startPt.TranslateBy(lineVec);
                    endPt = line.Geometry.EndPoint;
                    //rev = -1;
                    //vecDirect.ScaleBy(rev);
                }
                startPt.TranslateBy(vecDirect);
                vecDirect.Normalize();    
                pt1 = startPt.Copy();

                for (int i = 0; i < count; i++)
                {
                    //sum += l1;
                    if (i != 0)
                    {
                        vecDirect.ScaleBy(l1);
                        pt1.TranslateBy(vecDirect);
                    }
                    pt2 = pt1.Copy();
                    vecNormal.ScaleBy(-H);
                    pt2.TranslateBy(vecNormal);
                    line1 = ps.SketchLines.AddByTwoPoints(pt1, pt2);
                    line1.Construction = true;
                    line1.Centerline = true;

                    ps.GeometricConstraints.AddPerpendicular((SketchEntity)line1, (SketchEntity)line);
                    ps.GeometricConstraints.AddCoincident((SketchEntity)line1.StartSketchPoint, (SketchEntity)line);
                    cen = pt1.Copy();
                    vecDirect.Normalize();
                    vecNormal.Normalize();
                    //vecNormal.ScaleBy(rev);
                    vecNormal.ScaleBy(-1);
                    vecNormal.ScaleBy(-R); 
                    vecDirect.ScaleBy(-R);
                    cen.TranslateBy(vecNormal);
                    arcStart = cen.Copy(); arcStart.TranslateBy(vecDirect);
                    vecDirect.Normalize();
                    //vecDirect.ScaleBy(rev);
                    vecDirect.ScaleBy(-1);
                    vecDirect.ScaleBy(R); 
                    arcEnd = cen.Copy(); arcEnd.TranslateBy(vecDirect);
                    arc = ps.SketchArcs.AddByThreePoints(arcStart,pt1, arcEnd);
                    vecDirect.Normalize();
                    vecNormal.Normalize();
                    pt3 = pt2.Copy();
                    pt4 = pt2.Copy();
                    if (checkBox1.Checked && (i == 0 || i == (count-1)))
                    {
                        vecDirect.ScaleBy(-R); pt3.TranslateBy(vecDirect); vecDirect.Normalize();
                        //vecDirect.ScaleBy(rev);
                        vecDirect.ScaleBy(-1);
                        vecDirect.ScaleBy(R); pt4.TranslateBy(vecDirect); vecDirect.Normalize();
                    }
                    else
                    {
                        vecDirect.ScaleBy(-L / 2); pt3.TranslateBy(vecDirect); vecDirect.Normalize();
                        //vecDirect.ScaleBy(rev);
                        vecDirect.ScaleBy(-1);
                        vecDirect.ScaleBy(L / 2); pt4.TranslateBy(vecDirect); vecDirect.Normalize();
                    }
                    line2 = ps.SketchLines.AddByTwoPoints(arc.StartSketchPoint, pt4);
                    line3 = ps.SketchLines.AddByTwoPoints(arc.EndSketchPoint, pt3);
                    if (line5 != null) oldLine = line5;
                    line5 = ps.SketchLines.AddByTwoPoints(line3.EndSketchPoint, line2.EndSketchPoint);
                    ps.GeometricConstraints.AddCoincident((SketchEntity)line1.EndSketchPoint, (SketchEntity)line5);
                    if (i==0)
                    { 
                        textPt = line1.Geometry.MidPoint;
                        ps.DimensionConstraints.AddTwoPointDistance(line1.StartSketchPoint, line1.EndSketchPoint, DimensionOrientationEnum.kAlignedDim, textPt);
                        ps.DimensionConstraints.AddRadius((SketchEntity)arc, arc.Geometry.Center);
                        oldLine2 = line5;
                        vecDirect.Normalize();
                        //vecDirect.ScaleBy(rev);
                        vecDirect.ScaleBy(-0.9);
                        startPt.TranslateBy(vecDirect);
                        //vecDirect.ScaleBy(rev);
                        vecDirect.ScaleBy(-1);
                        line6 = ps.SketchLines.AddByTwoPoints(line3.EndSketchPoint, startPt);
                        ps.GeometricConstraints.AddCollinear((SketchEntity)line5, (SketchEntity)line6);
                        vecNormal.Normalize();
                        //vecDirect.ScaleBy(rev);
                        vecNormal.ScaleBy(-1);
                        vecNormal.ScaleBy(0.1);
                        startPt.TranslateBy(vecNormal); vecNormal.ScaleBy(-1); vecNormal.Normalize();
                        line6 = ps.SketchLines.AddByTwoPoints(line6.EndSketchPoint, startPt);   
                        ps.GeometricConstraints.AddPerpendicular((SketchEntity)line6, (SketchEntity)line);
                        vecDirect.ScaleBy(0.1);
                        endPt.TranslateBy(vecDirect); vecDirect.Normalize();
                        line6 = ps.SketchLines.AddByTwoPoints(line6.EndSketchPoint, endPt);
                        ps.GeometricConstraints.AddParallel((SketchEntity)line6, (SketchEntity)line);
                        vecNormal.ScaleBy(H + 0.1);
                        endPt.TranslateBy(vecNormal); vecNormal.Normalize();
                        line6 = ps.SketchLines.AddByTwoPoints(line6.EndSketchPoint, endPt);
                        ps.GeometricConstraints.AddPerpendicular((SketchEntity)line6, (SketchEntity)line);
                    }
                    if (i == 0 || i == (count - 1)) ps.GeometricConstraints.AddParallel((SketchEntity)line5, (SketchEntity)line);

                    if (i == 0)
                    {
                        textPt = line5.Geometry.MidPoint;
                        ps.DimensionConstraints.AddTwoPointDistance(line5.StartSketchPoint, line5.EndSketchPoint, DimensionOrientationEnum.kAlignedDim, textPt);
                    }

                    if (checkBox1.Checked && i == 1)
                    {
                        textPt = line5.Geometry.MidPoint;
                        ps.DimensionConstraints.AddTwoPointDistance(line5.StartSketchPoint, line5.EndSketchPoint, DimensionOrientationEnum.kAlignedDim, textPt);
                    }
                  
                    if (checkBox1.Checked && (i == 1 || i == (count - 1)))
                    {
                        oldLine = null;
                    }

                    if (checkBox1.Checked && count > 2 && i == (count-1))
                    {
                        ps.GeometricConstraints.AddEqualLength(line5, oldLine2);
                    }
                    if (oldLine != null) ps.GeometricConstraints.AddEqualLength(oldLine, line5);

                    line5.Construction = true;

                    if (i == 1)
                    {
                        line4 = ps.SketchLines.AddByTwoPoints(((SketchLine)sym.EntityOne).EndSketchPoint, line3.EndSketchPoint);
                        ps.GeometricConstraints.AddParallel((SketchEntity)line4, (SketchEntity)line);
                        ps.GeometricConstraints.AddEqualRadius((SketchEntity)tg.EntityTwo, (SketchEntity)arc);
                    }
                    if (i > 1)
                    {
                        SketchLine oldLine4 = line4;
                        line4 = ps.SketchLines.AddByTwoPoints(((SketchLine)sym.EntityOne).EndSketchPoint, line3.EndSketchPoint);
                        ps.GeometricConstraints.AddCollinear((SketchEntity)line4, (SketchEntity)oldLine4);
                        ps.GeometricConstraints.AddEqualRadius((SketchEntity)tg.EntityTwo, (SketchEntity)arc);
                    }

                    tg = ps.GeometricConstraints.AddTangent((SketchEntity)arc, (SketchEntity)line2);
                    tg = ps.GeometricConstraints.AddTangent((SketchEntity)arc, (SketchEntity)line3);
                    tg = ps.GeometricConstraints.AddTangent((SketchEntity)line, (SketchEntity)arc);
                    sym = ps.GeometricConstraints.AddSymmetry((SketchEntity)line2, (SketchEntity)line3, line1);
                    //CoincidentConstraint cc = ps.GeometricConstraints.AddCoincident((SketchEntity)arc.StartSketchPoint, (SketchEntity)line3.StartSketchPoint);
                    //cc = ps.GeometricConstraints.AddCoincident((SketchEntity)arc.EndSketchPoint, (SketchEntity)line2.StartSketchPoint);

                    vecDirect.Normalize();
                    vecNormal.Normalize();
                    //vecDirect.ScaleBy(rev);
                    vecNormal.ScaleBy(-1);

                    
                }
                    //UnitVector norm = ps.PlanarEntityGeometry.Normal;
                line7 = ps.SketchLines.AddByTwoPoints(line2.EndSketchPoint, line6.EndSketchPoint);
                ps.GeometricConstraints.AddParallel((SketchEntity)line7, (SketchEntity)line);
                var ie = ps.SketchLines.OfType<SketchLine>().Where(l => l.Construction && l.Centerline);
                foreach (SketchLine item in ie)
                {
                    IEnumerable<SketchLine> l;
                    if (Macros.StandardAddInServer.Spike == "Вертикальные шипы")
                    l = ps.SketchLines.OfType<SketchLine>().Where(lin => lin.Construction && lin.Centerline &&
                        ((int)lin.Geometry.StartPoint.X*1000 == (int)item.Geometry.StartPoint.X*1000)/*colin(lin.Geometry.StartPoint,lin.Geometry.EndPoint, item.Geometry.EndPoint)*/);
                    else
                    l = ps.SketchLines.OfType<SketchLine>().Where(lin => lin.Construction && lin.Centerline &&
                        ((int)lin.Geometry.StartPoint.Y*1000 == (int)item.Geometry.StartPoint.Y*1000));
                    if (l != null/* && !comparePoint(l.Geometry.StartPoint, item.Geometry.StartPoint,1000) &&*/
                        /*!comparePoint(l.Geometry.EndPoint,item.Geometry.EndPoint,1000*/)
                    {
                        try
                        {
                            foreach (var elem in l)
                            {
                                ps.GeometricConstraints.AddCollinear((SketchEntity)item, (SketchEntity)elem);
                            }
                        }
                        catch 
                        {
                        }
                    } 
                }
                    tr.End();
                this.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
                tr.End();
                this.Close();
            }
        }

        public double len(Inventor.Point v1, Inventor.Point v2)
        {
            Vector vec = invApp.TransientGeometry.CreateVector(v2.X - v1.X, v1.Y - v2.Y, v2.Z - v1.Z);
            return vec.Length;
        }

        private List<Inventor.Point> removeDub(List<Inventor.Vertex> lstPt)
        {
            List<Vertex> tmpVtx; List<Inventor.Point> tmpPts;
            tmpPts = new List<Inventor.Point>();
            for (int i=0; i<lstPt.Count; i++)
            {
            tmpVtx = lstPt.FindAll(delegate (Vertex v){return v == lstPt[i];});
            if (tmpVtx != null && tmpVtx.Count == 1)
            {
                tmpPts.Add(tmpVtx[0].Point);
            }
            }
            return tmpPts;
        }
        bool comparePoint(Point2d pt1, Point2d pt2, int tolerance)
        {
            return ((int)(pt1.X * tolerance) == (int)(pt2.X * tolerance) &&
                (int)(pt1.Y * tolerance) == (int)(pt2.Y * tolerance)) ? true : false;
        }
        bool comparePoint(Inventor.Point pt1, Inventor.Point pt2, int tolerance)
        {
            return ((int)(pt1.X * tolerance) == (int)(pt2.X * tolerance) &&
                (int)(pt1.Y * tolerance) == (int)(pt2.Y * tolerance) &&
                (int)(pt1.Z * tolerance) == (int)(pt2.Z * tolerance)) ? true : false;
        }
    }
}
