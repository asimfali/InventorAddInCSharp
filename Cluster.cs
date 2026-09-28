using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using InterfaceDll;
using Inventor;
using InvDoc;
//using System.Xml.Linq;

namespace InvAddIn
{
    enum clusterDirection{ x, y };

    class Cluster : Entity
    {
        double h = 0, w = 0;
        SketchLine slh, slw;
        Point2d insPt; 
        //bool one = true;
        //double offsetX = 0, offsetY = 0;
        //double x, y;
        int countX = 0, countY = 0;
        new MyXML xml; new PlanarSketch ps;
        PlanarSketch psOut;
        new PartDocument doc;
        PartComponentDefinition compDef;
        List<ClusterData> data = new List<ClusterData>();
        public bool full = false;
        CutFeature ob = null;
        public Cluster()
            : base(entTypes.Sketch)
        {
            if (I.aDoc().DocumentType != DocumentTypeEnum.kPartDocumentObject) return;
            MyXML xmlInterface = new MyXML("ClusterInterface.xml");
            MyForm F = new MyForm(xmlInterface, "Кластер");
            F.f.ShowDialog();
            if (F.xmls.Count != 2) return;
            xml = F.xmls[1];
            Dictionary<string,string> dic = new Dictionary<string,string>();
            xml.set("val", F.cbs[0].Text);
            ClusterData.startx = 0; ClusterData.starty = 0;
            xml.getInt("countX", ref ClusterData.countX); xml.getInt("countY", ref ClusterData.countY);
            xml.getDouble("D", ref ClusterData.D, ClusterData.tol); xml.getDouble("dx", ref ClusterData.dx, ClusterData.tol);
            xml.getDouble("dy", ref ClusterData.dy, ClusterData.tol);  xml.getInt("rotate", ref ClusterData.rotate);
            if (xml.getAtt("full") != "") full = true;
            foreach (var item in xml.elem.Elements())
            {
                data.Add(new ClusterData(new MyXML(item))); 
            }
//             if (ClusterData.rotate == 1)     
//             {
//                 data.Insert(0, data[1]);
//                 data.Remove(data[2]);
//             }
            doc = I.aDoc() as PartDocument;
            compDef = doc.ComponentDefinition;
            var ss = I.getSS();
            if (ss.Count == 1 && ss[1] is PlanarSketch) ps = ss[1] as PlanarSketch;
            else ps = compDef.Sketches[compDef.Sketches.Count];
            if (ps.Consumed) return;
            psOut = compDef.Sketches.Add(ps.PlanarEntity);
            psOut.OriginPoint = ps.OriginPoint;
            if (ps.AxisEntity != null)
                psOut.AxisEntity = ps.AxisEntity;
            psOut.AxisIsX = ps.AxisIsX;
            psOut.NaturalAxisDirection = ps.NaturalAxisDirection;
            get();
        }

        public override void add()
        {
                      
        }

        public void add(ClusterData cd, double offsetX, double offsetY)
        {
            List<Point2d> pts;
            pts = cd.get(insPt, offsetX, offsetY);
            foreach (Point2d pt in pts)
            {
               psOut.SketchPoints.Add(pt); 
            }
        }

        public void setCount(int val)
        {
            foreach (var item in data)
            {
                item.count = val; 
            }
        }

        public override void draw()
        {
            if (data.Count == 4)
            {
//                 if (ClusterData.rotate == 1)
//                 {
//                     ClusterData.swap();
//                 }
                x /= -2; y /= -2;
                if (ClusterData.dir.X != 0)
                {
                    setCount(countX * ClusterData.countX);
                    if (ClusterData.rotate == 0)
                    {
                        draw(data, new List<int> { 0, 1, 2, 3 });
                    }
                    else if (ClusterData.rotate == 1)
                    {
                        data[1].count = 1;
                        draw(data, new List<int> { 0, 1, 2 });
                    }
                }
                else
                {
                    setCount(countY * ClusterData.countY);
                    if (ClusterData.rotate == 0)
                    {
                        data[1].count = 1;
                        draw(data, new List<int> { 0, 1, 2 });
                    }
                    else if (ClusterData.rotate == 1)
                    {
                        draw(data, new List<int> { 0, 1, 2, 3 });
                    }
                }
                if (ClusterData.dir.X != 0)
                {
                    setCount(countX * ClusterData.countX);
                    y = y * -1 - ClusterData.starty;
                    if (ClusterData.rotate == 0)
                    {
                        draw(data, new List<int> { 0, 1, 2, 3 });
                    }
                    else if (ClusterData.rotate == 1)
                    {
                        //data[1].count = countX * ClusterData.countX;
                        draw(data, new List<int> { 1, 3 });
                    }
                }
                else
                {
                    setCount(countY * ClusterData.countY);
                    x = x * -1 - ClusterData.startx;
                    if (ClusterData.rotate == 0)
                    {
                        draw(data, new List<int> { 1, 3 });
                    }
                    else if (ClusterData.rotate == 1)
                    {
                        draw(data, new List<int> { 0, 1, 2, 3 });
                    }
                }
                CreateComponent.hole(compDef, (ClusterData.D*ClusterData.tol).ToString(), psOut);
            }
            else if (data.Count == 2 || data.Count == 1)
            {

                x /= -2; y /= -2;
                setCount(countX * ClusterData.countX);
                if (ClusterData.dir.X != 0)
                {
                    draw(data, new List<int> { 0,1});
                }
                else
                {
                    setCount(countY * ClusterData.countY);
                    draw(data, new List<int> { 0,1});
                }

                //double maxy = y *-1 - y, maxx = x *-1 - x;
                if (!this.full)
                {
                    //draw(data, new List<int> { 1 });
                    if (ClusterData.dir.X != 0)
                    {
                        y = y * -1 - ClusterData.starty;
                        draw(data, new List<int> { 0, 1 });
                    }
                    else
                    {
                        setCount(countY * ClusterData.countY);
                        x = x * -1 - ClusterData.startx;
                        draw(data, new List<int> { 0, 1 });
                    }

                }
                //Dictionary<string, double> dic = new Dictionary<string, double>() { { "a", ClusterData.l*ClusterData.tol }, { "ширина", ClusterData.D*ClusterData.tol } };
                //I.createPunch("Punches\\Овал.ide", psOut, dic, 0);
                if (data[0].name == "Slot")
                {
                    foreach (SketchPoint item in psOut.SketchPoints)
                    {
                        I.createSlot(item.Geometry, psOut, I.CV2d(0, 1), (ClusterData.l - ClusterData.D) / 2, ClusterData.D);
                    }
                    ob = I.createCut(psOut);
                }
                else if (data[0].name == "Flower")
                {
                    CreateComponent.hole(compDef, (ClusterData.D * ClusterData.tol).ToString(), psOut);
                }
            }
        }

        public void draw(List<ClusterData> lst, List<int> filter, int xsing = 1, int ysing = 1)
        {
            for (int i = 0; i < lst.Count; i++)
            {
                if (filter.Contains(i)/*true*/) add(lst[i], x*xsing, y*ysing); 
            }
        }

        public void txt()
        {
            string fn = I.aDoc().FullFileName;
            fn = fn.Replace(".ipt", "(Кластер).txt");
            int cx = (countX * ClusterData.countX) * data.Count;
            int cy = (countY * ClusterData.countY) * data.Count;
            if (data.Count == 2) cy = countY * ClusterData.countY;
            double a = 0;
            xml.getDouble("area", ref a, 1);
            using (System.IO.StreamWriter sw = System.IO.File.CreateText(fn))
            {
                sw.Write("Кластер: " + MyXML.getAtt(xml.elem, "val") + "\n");
                sw.Write("Количество ударов: " + cx*cy/(ClusterData.countX*ClusterData.countY) +"\n");
                sw.Write("Количество отверстий по горизонтали: " + cx + " шт.\n");
                sw.Write("Количество отверстий по вертикали: " + cy + " шт.\n");
                sw.Write("Общее количество отверстий: " + cx * cy + " шт.\n");
                sw.Write("Площадь: " + (cx * cy * a).ToString("##.####") + "мм2.");
            }
            if (this.full)
            {
                object ax = null;
                UnitVector v = null;
                ObjectCollection col = I.COC(ob);
                int c = cx;
                double dist = ClusterData.dx;
                v = psOut.AxisEntityGeometry.Direction;
                if (ClusterData.dir.X == 0)
                {
                    
                }
                else
                {
                    c = cy;
                    dist = ClusterData.dy;
                    if (psOut.AxisIsX)
                    {
                        double[] d = { };
                        double[] dpl = { };
                        v.GetUnitVectorData(ref d);
                        psOut.PlanarEntityGeometry.Normal.GetUnitVectorData(ref dpl);

                        for (int i = 0; i < d.Length; i++)
                        {
                            if (d[i] != 0 || dpl[i] != 0)
                            {
                                d[i] = 0;
                            }
                            else d[i] = -1;
                        }
                        v = u.createUnitVector(d[0], d[1], d[2]);
                    }
                }
                ax = u.getWAxis(compDef, v);
                compDef.Features.RectangularPatternFeatures.Add(col, ax, psOut.NaturalAxisDirection, c, dist);
            }
        }

        public override void upd()
        {
            throw new NotImplementedException();
        }

        public void change(double max, double start, double dx, int count, bool dir, double c, double z)
        {
            //if (rotate) { z++; c++; }
            double offset = start;
            int i = 0; double a = dx/z * (count - c);
            while (offset <= max - ClusterData.D)
            {
                offset += a;
                i++;
            }
            offset -= dx;
            if (offset > max)
            {
                offset -= a; i--;
            }
//             if (ClusterData.rotate == 1) 
//             {
//                dir = !dir;
//             }
            if (dir == true) { countX = i; x = offset; }
            else { countY = i; y = offset; }
        }

        public override void get()
        {
            IEnumerable<SketchLine> slX = u.gets<SketchLine>(ps.SketchLines, f => u.eq(f.Geometry.StartPoint.Y, f.Geometry.EndPoint.Y) && f.Construction == false);
            if (slX.Count() != 2) return;
            IEnumerable<SketchLine> slY = u.gets<SketchLine>(ps.SketchLines, f => u.eq(f.Geometry.StartPoint.X, f.Geometry.EndPoint.X) && f.Construction == false);
            if (slY.Count() != 2) return;
            slX.OrderBy(e => e.Geometry.StartPoint.Y); slY.OrderBy(e =>  e.Geometry.StartPoint.X);
            slh = slY.ElementAt(1); slw = slX.ElementAt(1);
            //double l1 = slh.Length, l2 = slw.Length;
            //SketchLine d = I.Pick(SelectionFilterEnum.kSketchCurveLinearFilter, "Выберите направление") as SketchLine;
            //ClusterData.dir = d.Geometry.Direction;         
            //             if (!(d.Geometry.Direction.X == 0 && slh.Geometry.Direction.X == 0))
            //             {
            //                 slh = slw; slw = slY.ElementAt(1);
            //             }
            //insPt = I.CP2d(Math.Min(slX.ElementAt(0).Geometry.StartPoint.X, slX.ElementAt(1).Geometry.StartPoint.X),
            //    Math.Min(slY.ElementAt(0).Geometry.StartPoint.Y, slY.ElementAt(1).Geometry.StartPoint.Y));
            Point2d p1 = slX.ElementAt(0).Geometry.StartPoint, p2 = slX.ElementAt(0).Geometry.EndPoint, p3 = slX.ElementAt(1).Geometry.StartPoint,
                p4 = slX.ElementAt(1).Geometry.EndPoint;
            Point2d mp1 = u.midPt(p1, p2), mp2 = u.midPt(p3, p4);
            insPt = u.midPt(mp1, mp2);
//             insPt = slw.StartSketchPoint.Geometry.X < slw.EndSketchPoint.Geometry.X ? slw.StartSketchPoint.Geometry : slw.EndSketchPoint.Geometry;
//             insPt = slw.StartSketchPoint.Geometry.Y < slw.EndSketchPoint.Geometry.Y ? insPt : 
            int ix = 1, iy = 1;
//             if (insPt.X < 0) ix = -1; 
//             if (insPt.Y < 0) iy = -1;
            if (u.eq(mp1.X, mp2.X)) { /*insPt.X += slw.Length / 2*ix; insPt.Y += slh.Length / 2*iy;*/ h = u.round(slh.Length); w = u.round(slw.Length); }
            else if (u.eq(mp1.Y, mp2.Y)) { /*insPt.X += slh.Length / 2*ix; insPt.Y += slw.Length / 2*iy;*/ h = u.round(slw.Length); w = u.round(slh.Length); }
            //ClusterData.dir = slw.Geometry.Direction;
            if (h > w) ClusterData.dir = I.CUV2d(1, 0);
            else ClusterData.dir = I.CUV2d(0, 1);
            if (data.Count == 0) return;
            //ClusterData.dir = I.CUV2d(1, 0);
            double dx = ClusterData.dx, dy = ClusterData.dy;
            change(w, ClusterData.startx, dx, ClusterData.countX, true, 0, 1);
            double old = 0;
            if (ClusterData.l != 0)
            {
                old = ClusterData.D;
                ClusterData.D = ClusterData.l;
            }
            change(h, ClusterData.starty, dy, ClusterData.countY, false, 0, 1);
            if (ClusterData.l != 0)
            {
                ClusterData.D = old;
            }
            //ps.SketchPoints.Add(insPt);
            //psOut.SketchLines.AddAsTwoPointCenteredRectangle(insPt, I.CP2d(insPt.X-x/2,insPt.Y-y/2));
        }
    }

    class ClusterData
    {
        static public double tol = 10;
        public string name;
        static public int rotate = 0;
        static public int countX = 0, countY = 0;
        static public double D = 0, dx = 0, dy = 0;
        static public double startx = 0, starty = 0;
        public double x, y;
        static public UnitVector2d dir = I.CUV2d(0,1);
        public int count;
        static public double l = 0;        
        public ClusterData(MyXML xml)
        {
            name = xml.elem.Name.ToString();
            //xml.getDouble("D", ref D, tol); 
            xml.getDouble("x", ref x, tol); 
            xml.getDouble("y", ref y, tol);
            xml.getDouble("l", ref l, tol);
            if (startx < x) startx = x;
            if (starty < y) starty = y;
            //xml.getDouble("dx", ref dx, tol); xml.getDouble("dy", ref dy, tol);
        }
//         public Point2d get(Point2d insPt, double offsetX, double offsetY)
//         {
//             return I.CP2d(insPt.X + x + offsetX, insPt.Y + y + offsetY);
//          }
        public List<Point2d> get(Point2d insPt, double offsetX, double offsetY)
        {
            List<Point2d> pst = new List<Point2d>();
//             if (insPt.X < 0) offsetX = -offsetX;
//             if (insPt.Y < 0) offsetY = -offsetY;
            for (int i = 0; i < count; i++)
            {
                pst.Add(I.CP2d(insPt.X + x + offsetX + dx * dir.X * i, insPt.Y + y + offsetY + dy * dir.Y * i)); 
            }
            return pst;
        }
        static public void swap()
        {
            double xd = dir.X, yd = dir.Y;
            dir = I.CUV2d(xd, yd);
        }
    }
}
