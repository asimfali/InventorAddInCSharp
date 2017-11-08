using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml;
using System.Xml.Linq;
using System.Text;
using InvDoc;
using Inventor;
using InterfaceDll;
using System.Diagnostics;


namespace InvAddIn
{
    enum typeFastener { hole, slot}

    class Fastener
    {
        AssemblyDocument doc;
        AssemblyComponentDefinition acd;
        public double maxR = 0.75;
        public List<fastenerSource> edges = new List<fastenerSource>();
        //public List<fastenerSlot> slEdges = new List<fastenerSlot>();
        public List<FaceProxy> except = new List<FaceProxy>();
        public List<slot> slots = new List<slot>();
        MyXML xml, cont;
        //public IEnumerable<FaceProxy> faces;
        public List<FaceProxy> faces;

        public Fastener(Document doc)
        {
            this.doc = doc as AssemblyDocument;
            if (this.doc == null) return;
            acd = I.getACD(doc);
            if (Macros.StandardAddInServer.xml == null)
                xml = new MyXML("Fast.xml");
            else xml = Macros.StandardAddInServer.xml;
            cont = new MyXML("ContentCenter.xml");
            find(acd.Occurrences);
            //faces = faces.Distinct();
            int count = faces.Count;
            //addSlot();
            foreach (var item in slots)
            {
                findSlots(item);
            }
            for (int i = 0; i < count; i++)
            {
                i++;
                List<EdgeProxy> lst = findEdges();
                if (lst != null)
                {
                    fastenerSource ed = new fastenerSource(lst);
                    if (ed != null)
                    {
                        edges.Add(ed);
                    }
                }
                if (faces.Count() == 0) break;
                /*lst = */
            }
            cleanEdges();
        }

//         public void addSlot()
//         {
//             foreach (FaceProxy item in faces)
//             {
//                 if (item.TangentiallyConnectedFaces.Count == 3)
//                 {
//                     slots.Add(new slot(item));
//                 }
//             } 
//         }

        public void add()
        {
            foreach (fastenerSource item in edges)
            {
                if (item.e.Count > 2)
                    find(item);
            }
        }

        private void find(System.Collections.IEnumerable coe)
        {
            foreach (ComponentOccurrence item in coe)
            {
                if (item.Suppressed) continue;
                if (item.ReferencedDocumentDescriptor.FullDocumentName.IndexOf("Content Center Files") != -1) continue;
                if (item.ReferencedDocumentDescriptor.FullDocumentName.EndsWith(".ipt"))
                {
                    faces = u.add<FaceProxy>(faces, u.gets<FaceProxy>(item.SurfaceBodies[1].Faces, f => check(f, maxR))).ToList();
                }
                else if (item.ReferencedDocumentDescriptor.FullDocumentName.EndsWith(".iam"))
                {
                    find(item.SubOccurrences);
                }
            }
        }

        private bool check(FaceProxy f, double maxR)
        {
            if (f.SurfaceType != SurfaceTypeEnum.kCylinderSurface) return false;
            Cylinder c = f.Geometry as Cylinder;
            if (c == null) return false;
            Point pt = c.BasePoint.Copy();
            //pt.TranslateBy(c.AxisVector.AsVector());
            if (f.TangentiallyConnectedFaces.Count == 3)
            {
                if (c.Radius == 9.3 / 20 || c.Radius == 0.2) return false;
                slots.Add(new slot(f));
                return false;
            }
            return check(pt) && f.SurfaceType == SurfaceTypeEnum.kCylinderSurface && (f.CreatedByFeature is HoleFeatureProxy || f.CreatedByFeature is CutFeatureProxy || 
                f.CreatedByFeature is ReferenceFeatureProxy || 
                f.CreatedByFeature is PunchToolFeatureProxy || f.CreatedByFeature is MirrorFeatureProxy
                    || f.CreatedByFeature is RectangularPatternFeatureProxy) 
               && (f.TangentiallyConnectedFaces.Count == 3 || f.TangentiallyConnectedFaces.Count == 0) && ((Cylinder)f.Geometry).Radius < maxR;
        }

        private bool check(Point pt)
        {
            int count = 0; bool f = false;
            foreach (ComponentOccurrence item in acd.Occurrences)
            {
                f = item.RangeBox.Contains(pt);
                if (f && item.ReferencedDocumentDescriptor.FullDocumentName.IndexOf("Content Center Files") != -1) return false;
                //if (f) count++;
            }
            return true;
        }

        private void find(fastenerSource fast)
        {
            Occ.rev = 1;
            fast.create();
            if (fast.filter()) return;
            if (fast.l.Length < 3) return;       
            xml.setRoot();
            cont.setRoot();
            string nameAtt = "hole";
            if (fast.tf == typeFastener.slot) nameAtt = "slot";
            if (!xml.set(xml.getEl(new Dictionary<string, string>() { { "d1", u.convToString(fast.d.First(),3,"##.##") }, { "d2", u.convToString(fast.d.Last(),3,"##.##") }, { "name", nameAtt } }, false))) return;
            Occ old = default(Occ); bool first = true; bool change = false;
            fastEdge ed = fast.fe;
            if (MyXML.getAtt(xml.elem, "rev") != "") ed = fast.le;
            if (contentCenter.check(ed.e as EdgeProxy)) return;
            foreach (var item in xml.elem.Elements())
            {
                string v = MyXML.getAtt(item, "val");
                string strL = MyXML.getAtt(item, "l");
                string r = MyXML.getAtt(item, "rev");
                if (v == "") continue;
                if (v == "rev")
                { Occ.rev = -1; first = true; change = true; continue; }
                xml.setRoot();
                var spl = v.Split(':');
                xml.set("abbr", spl[0]);
                cont.set("name", MyXML.getAtt(xml.elem, "name"));
                int n = int.Parse(spl[1]) - 1;
                cont.set(n);
                string name = cont.getAtt("Id");
                contentCenter cc = new contentCenter(doc as Document);
                Occ co = cc.place(name);
                string fname = xml.getAtt("folder");
                addToFolder(co, fname);
                xml.setRoot();
                cont.setRoot();
                if ((spl[0] == "ГЗ" || spl[0] == "ГЗП") && !first) 
                {
                    //co.dist = ls + lm;
                    if (strL == "" || !u.eq(strL, fast.l[1] * 10)) { co.remove(); continue; }
                    if (fast.fe.Equals(fast.e.First())) Array.Reverse(fast.l);
                    double d = -(fast.l[0] + fast.l[1]);
                    bool axis = false;
                    if (r != "") { axis = !axis; d = fast.l.Sum(); }
                    if (nameAtt == "hole")
                    { cc.add(ed.e, co, d, !axis); }
                    else if (nameAtt == "slot")
                    {
                        fastenerSlot fs = fast as fastenerSlot;
                        FaceProxy fac = u.get<FaceProxy>(ed.e.Faces, fi => fi.SurfaceType == SurfaceTypeEnum.kPlaneSurface);
                        if (fac != null) cc.add(fac, co, d, axis);
                        if (fs.fp != null) cc.add(fs.fp, co);
                    }
                    continue;
                }
                double dist = 0;
                if (change) dist = co.dist;
                if (first)
                {
                    if (change) dist = (fast.l.Sum()) * Occ.rev + co.dist;
                    switch (fast.tf)
                    {
                        case typeFastener.hole:
                            cc.add(ed.e, co, dist);
                            break;
                        case typeFastener.slot:
                            fastenerSlot fs = fast as fastenerSlot;
                            FaceProxy fac = u.get<FaceProxy>(ed.e.Faces, fi => fi.SurfaceType == SurfaceTypeEnum.kPlaneSurface);
                            if (fac != null) cc.add(fac, co, dist);
                            if (fs.fp != null) cc.add(fs.fp, co);
                            break;
                        default:
                            break;
                    }
                    first = false; Occ.rev = -Occ.rev;
                }
                     
                else
                {
                    cc.add(old, co, dist);
                }                
                old = co;
                
            } 
        }

        private void cleanEdges()
        {
            List<fastenerSource> sum = new List<fastenerSource>();
            for (int i = 0; i < edges.Count; i++)
            {
                List<fastenerSource> lst = edges[i].divide();
                if (lst != null)
                {
                    sum.AddRange(lst);
                }
            }
            if (sum.Count != 0)
            {
                for (int i = 0; i < edges.Count; i++)
                {
                    if (!edges[i].remove) sum.Add(edges[i]);
                }
                edges = sum;
            }
        }

        private List<EdgeProxy> findEdges()
        {
            if (faces.Count == 0) return null;
            //FaceProxy face = faces.ElementAt(0);
            FaceProxy face = u.get<FaceProxy>(faces, fi => !except.Contains(fi));
            if (face == null) return null;
            //except.Add(face);
            Cylinder cylfind = face.Geometry as Cylinder;
            UnitVector vb = cylfind.AxisVector;
            Point ptb = cylfind.BasePoint;
            bool isHole = true;
            if (face.TangentiallyConnectedFaces.Count == 3) isHole = false;
            List<EdgeProxy> lst = new List<EdgeProxy>();
            IEnumerable<FaceProxy> dist = null;
            dist = u.add<FaceProxy>(dist, face);
            if (face.Edges.Count == 2)
            {
                lst.AddRange(face.Edges.OfType<EdgeProxy>());
            }
            foreach (FaceProxy item in faces)
            {
                if (item.Equals(face)) continue;
                Cylinder c = item.Geometry as Cylinder;
                if (c == null) continue;
                UnitVector v = c.AxisVector;
                Point pt = c.BasePoint;
                if (u.eq(ptb, vb, pt))
                {
                    if (item.Edges.Count == 2)
                    {lst.AddRange(item.Edges.OfType<EdgeProxy>());}
                    dist = u.add<FaceProxy>(dist, item);
//                     if (isHole && ptb.DistanceTo(pt) < 0.7)
//                         dist = u.add<Face>(dist, item);
//                     else
                        //except.Add(item);
                }
            }
            if (dist != null) { /*faces = faces.Except(dist).ToList();*/ except.AddRange(dist); }
            
            if (lst.Count > 2)
            {
                lst.AddRange(face.Edges.OfType<EdgeProxy>());
                return lst;
            }
            return null;
        }

        private void findSlots(slot item)
        {
            if (faces.Count == 0) return;
            if (item.missing) return;
            IEnumerable<Face> dist = null;
            //Debug.Print(faces.Count().ToString());
            bool add = false;
            List<EdgeProxy> lst = new List<EdgeProxy>();
            UnitVector v = null;
            if (item.bpt == null || item.dir == null) return;
            foreach (var f in faces)
            {
                Cylinder c = f.Geometry as Cylinder;
                v = c.AxisVector;
                Point pt = c.BasePoint;
                if (u.eq(item.bpt, item.dir, pt))
                {
                    //if (item.bpt.DistanceTo(pt) > 0.7) continue;
                    lst.AddRange(u.gets<EdgeProxy>(f.Edges, fi => fi.GeometryType == CurveTypeEnum.kCircleCurve));
                    add = true;
                    //dist = u.add<Face>(dist, f);
                }   
            }
            if (add)
            {
                lst.AddRange(item.e);
                fastenerSlot fs = new fastenerSlot(lst, item);
                fs.tf = typeFastener.slot;
                fs.cen = item.bpt;
                fs.dir = v;
                edges.Add(fs);
            }
            //if (dist != null) faces = faces.Except(dist);
        }

        public BrowserFolder addFolder(string name)
        {
            BrowserPane pane = doc.BrowserPanes["Модель"];

            foreach (BrowserFolder item in pane.TopNode.BrowserFolders)
	        {
		        if (item.Name == name) return item;
	        }
            return pane.AddBrowserFolder(name);
        }

        
        public bool containtsNode(BrowserNode node, BrowserFolder folder)
        {
            string[] spl = node.FullPath.Split(':');
            string name = spl[spl.Length - 2] + ":" + spl[spl.Length - 1];
            BrowserNode tmp = folder.BrowserNode.BrowserNodes.OfType<BrowserNode>().FirstOrDefault(e => e.FullPath.EndsWith(name));
            return tmp == null ? false : true;
        }

        public void addToFolder(Occ o, string nameFolder)
        {
            BrowserPane pane = doc.BrowserPanes["Модель"];
            ObjectCollection occNodes = I.app.TransientObjects.CreateObjectCollection();
            BrowserNode parentNode = pane.TopNode;
            BrowserFolder folder = addFolder(nameFolder);
            BrowserNode node = pane.GetBrowserNodeFromObject(o.occ);
            if (!containtsNode(node, folder)) folder.Add(node);
        }
    }

    class fastenerSource
    {
        public List<fastEdge> e = new List<fastEdge>();
        public UnitVector dir;
        public double[] d;
        public double[] l;
        public bool remove = false;
        public fastEdge fe, le;
        public typeFastener tf = typeFastener.hole;
        public fastenerSource(List<EdgeProxy> lst)
        {
            this.dir = (u.get<Edge>(lst, f => f.GeometryType == CurveTypeEnum.kCircleCurve).Geometry as Circle).Normal;
            u.abs(ref dir);
            //this.dir = (u.get<Edge>(lst, f => f.GeometryType == CurveTypeEnum.kCircularArcCurve).Geometry as Arc3d).Normal;
            lst = lst.Distinct().ToList();
            add(lst);
            e.Sort(new comparer<fastEdge>());
        }

        public void add(List<EdgeProxy> lst)
        {
            foreach (EdgeProxy item in lst)
            {
                e.Add(new fastEdge(item, dir));
            }
        }

        public bool filter()
        {
            ComponentOccurrence d1 = fe.e.ContainingOccurrence.OccurrencePath[1], d2 = le.e.ContainingOccurrence.OccurrencePath[1];
            return d1.Name.Equals(d2.Name) ? true : false;
        }

        public bool filter(fastEdge e1, fastEdge e2)
        {
            ComponentOccurrence d1 = e1.e.ContainingOccurrence.OccurrencePath[1], d2 = e1.e.ContainingOccurrence.OccurrencePath[1];
            return d1.Name.Equals(d2.Name) ? true : false;
        }

        virtual public void create()
        {
            getR(e);
            Array.Sort(d);
            if (u.eq(d[0], getR(e[0])*2))
            {
                e.Reverse();
            }
            l = getL(e);
            if (u.eq(d.First(), d.Last()))
            {
                if (l.First() > l.Last())
                {
                    e.Reverse(); l.Reverse();
                }
            }
            fe = e.First(); le = e.Last();
//             else if (d.First() < d.Last())
//             {
//                 fe = e.First(); le = e.Last();
//             }
//             else
//             {
//                 le = e.First(); fe = e.Last();
//             }
        }

        virtual public List<fastenerSource> divide()
        {
            if (e.Count <= 4) return null;
            List<fastenerSource> lst = new List<fastenerSource>();
            l = getL(e);
            List<EdgeProxy> ed = new List<EdgeProxy>();
            for (int i = 0; i < l.Length; i++)
            {
                ed.Add(e[i].e);
                if (l[i] > 0.3)
                {
                    lst.Add(new fastenerSource(ed));
                    ed.Clear();
                    //ed.Add(e[i].e);
                }
            }
            ed.Add(e[e.Count-1].e);
            if (ed.Count > 2)
            {
                lst.Add(new fastenerSource(ed));
            }
            if (lst.Count > 1)
            {
                remove = true;
                return lst;
            }
            return null;
        }

        virtual protected void getR(List<fastEdge> lst)
        {
            d = new double[lst.Count];
            for (int i = 0; i < lst.Count; i++)
            {
                double r = getR(lst[i]);
                d[i] = r * 2;
            }
        }

        virtual protected double getR(fastEdge e, double tol = 10)
        {
            Circle c = e.e.Geometry as Circle;
            if (c == null) return 0;
            return c.Radius * tol;
        }

        virtual protected double[] getL(List<fastEdge> lst)
        {
            double[] d = new double[lst.Count - 1];
            for (int i = 1; i < lst.Count; i++)
            {
                d[i - 1] = lst[i - 1].pt.DistanceTo(lst[i].pt);
//                 Circle c1 = lst[i - 1].e.Geometry as Circle, c2 = lst[i].e.Geometry as Circle;
//                 d[i - 1] = c1.Center.DistanceTo(c2.Center);
            }
            return d;
        }    
    }

    class fastenerSlot : fastenerSource
    {
        public double L;
        public FaceProxy fp;
        public double s = 0;
        public Point cen;
        public slot sl;
        public fastenerSlot(List<EdgeProxy> lst, slot sl)
            : base(lst)
        {
            cen = sl.bpt; this.sl = sl;
            EdgeProxy e = u.get<EdgeProxy>(lst, f => f.GeometryType == CurveTypeEnum.kCircleCurve);
            if (e != null)
            {
                fp = u.get<FaceProxy>(e.Faces, f => f.SurfaceType == SurfaceTypeEnum.kCylinderSurface);
            }
        }

        public override void create()
        {
            getR(e);
            Array.Sort(d);
            if (u.eq(d[0], getR(e[0]) * 2))
            {
                e.Reverse();
            }
            l = getL(e);
            if (u.eq(d.First(), d.Last()))
            {
                if (l.First() > l.Last())
                {
                    e.Reverse(); l.Reverse();
                }
            }
            fe = e.First(); le = e.Last();
        }

        protected override double[] getL(List<fastEdge> lst)
        {
            double[] d = new double[lst.Count - 1];
            for (int i = 1; i < lst.Count; i++)
            {
                Vector v = lst[i - 1].pt.VectorTo(lst[i].pt);
                double dist = Math.Abs(u.scalar(v, lst[i].dir.AsVector()));
                d[i - 1] = dist;
                //                 Circle c1 = lst[i - 1].e.Geometry as Circle, c2 = lst[i].e.Geometry as Circle;
                //                 d[i - 1] = c1.Center.DistanceTo(c2.Center);
            }
            return d;
        }

        protected override void getR(List<fastEdge> lst)
        {
            base.getR(lst);
        }

        protected override double getR(fastEdge e, double tol = 10)
        {
            if (e.e.GeometryType == CurveTypeEnum.kCircularArcCurve)
            {
                Arc3d a = e.e.Geometry as Arc3d;
                if (a == null) return 0;
                return (a.Radius + s) * tol;
            }
            else if (e.e.GeometryType == CurveTypeEnum.kCircleCurve)
            {
                Circle c = e.e.Geometry as Circle;
                if (c == null) return 0;
                return c.Radius * tol;
            }
            return 0;
        }

        public override List<fastenerSource> divide()
        {
            if (e.Count <= 4) return null;
            List<fastenerSource> lst = new List<fastenerSource>();
            l = getL(e);
            List<EdgeProxy> ed = new List<EdgeProxy>();
            for (int i = 0; i < l.Length; i++)
            {
                ed.Add(e[i].e);
                if (l[i] > 0.3)
                {
                    if (ed.Count > 2)
                        lst.Add(new fastenerSlot(ed ,sl)); 
                    ed.Clear();
                    //ed.Add(e[i].e);
                }
            }
            ed.Add(e[e.Count - 1].e);
            if (ed.Count > 2)
            {
                lst.Add(new fastenerSlot(ed, sl));
            }
            if (lst.Count >= 1)
            {
                foreach (var item in lst)
                {
                    item.tf = typeFastener.slot;
                }
                remove = true;
                return lst;
            }
            return null;
        }
    }

    class fastEdge
    {
        public EdgeProxy e;
        public UnitVector dir;
        public Point pt;
        public fastEdge(EdgeProxy ed, UnitVector d)
        {
            e = ed; dir = d;
            if (ed.GeometryType == CurveTypeEnum.kCircleCurve) pt = (e.Geometry as Circle).Center;
            else if (ed.GeometryType == CurveTypeEnum.kCircularArcCurve) 
            { 
                pt = (e.Geometry as Arc3d).Center;
            }
        }
    }

    class slot
    {
        public Point bpt;
        public UnitVector dir;
        public List<EdgeProxy> e = new List<EdgeProxy>();
        static public HashSet<FaceProxy> except = new HashSet<FaceProxy>();
        public bool missing = false;

        public slot(FaceProxy f)
        {
            if (except.Contains(f) || f.TangentiallyConnectedFaces.Count != 3) { missing = true; return; }
            List<FaceProxy> sf = new List<FaceProxy>();
            sf.Add(f);
            sf.AddRange(u.gets<FaceProxy>(f.TangentiallyConnectedFaces, fi => fi.SurfaceType == SurfaceTypeEnum.kCylinderSurface));
            if (sf.Count() == 2)
            {
                Cylinder c1 = sf.ElementAt(0).Geometry as Cylinder, c2 = sf.ElementAt(1).Geometry as Cylinder;
                if (u.eq(c1.Radius, c2.Radius))
                {
                    bpt = u.midPt(c1.BasePoint, c2.BasePoint);
                    e.AddRange(u.gets<EdgeProxy>(f.Edges, fi => fi.GeometryType == CurveTypeEnum.kCircularArcCurve));
//                     foreach (EdgeProxy item in u.gets<EdgeProxy>(f.Edges, fi => fi.GeometryType == CurveTypeEnum.kCircularArcCurve))
//                     {
//                         e.Add(new fastEdge(item, dir));
//                     } 
                    dir = c1.AxisVector;
                    except.Add(sf.ElementAt(1));
                }
            }
        }
    }

    class comparer<T> : IComparer<T> 
        where T: fastEdge
    {

        public int Compare(T x, T y)
        {
            double d1 = u.scalar(x.pt, x.dir), d2 = u.scalar(y.pt, x.dir);
            if (u.eq(d1, d2)) return 0;
            return d1 < d2 ? -1 : 1;
//             if (!u.eq(c1.X,c2.X))
//                 return c1.Center.X < c2.X ? -1 : 1;
//             if (!u.eq(c1.Center.Y, c2.Center.Y))
//                 return c1.Center.Y < c2.Center.Y ? -1 : 1;
//             if (!u.eq(c1.Center.Z, c2.Center.Z))
//                 return c1.Center.Z < c2.Center.Z ? -1 : 1;
            //return 0;
        }
    }

    struct Occ
    {
        public ComponentOccurrence occ;
        public InsertiMateDefinitionProxy def;
        public Circle cir;
        public double dist;
        public UnitVector dir;
        static public int rev = 1;
        public Occ(ComponentOccurrence c, double dist = -100)
        {
            occ = c;
            def = u.get<InsertiMateDefinitionProxy>(c.iMateDefinitions, i => i.Type == ObjectTypeEnum.kInsertiMateDefinitionProxyObject);
            cir = ((EdgeProxy)def.Entity).Geometry as Circle;
            if (cir != null) dir = cir.Normal;
            else
            {
                Arc3d a = ((EdgeProxy)def.Entity).Geometry as Arc3d;
                dir = a.Normal;
            }
            Vector v = c.RangeBox.MinPoint.VectorTo(c.RangeBox.MaxPoint);
            if (dist != -100) this.dist = dist;
            else this.dist = Math.Abs(v.DotProduct(dir.AsVector()))*rev;
        }
        public void remove()
        {
            occ.Delete();
        }
    }

    class contentCenter
    {
        public string memberfilename;
        Matrix mtx;
        object row;
        MemberManagerErrorsEnum err;
        string failuremessage;
        AssemblyComponentDefinition acd;
        ComponentOccurrence co;

        public contentCenter(Document doc)
        {
            acd = I.getACD(doc);
        }

        public Occ place(string id)
        {
            mtx = I.tg.CreateMatrix();
            ContentCenter c = I.app.ContentCenter;
            ContentFamily cf = (ContentFamily)c.GetContentObject(id.Substring(0, id.LastIndexOf('#') + 1));
            ContentTableRow row = (ContentTableRow)c.GetContentObject(id);
            memberfilename = cf.CreateMember(row, out err, out failuremessage);
            co = acd.Occurrences.Add(memberfilename, mtx);
            return new Occ(co);
        }

        public void add(EdgeProxy ed, Occ occ, double dist = 0, bool axis = true)
        {
            acd.Constraints.AddInsertConstraint(occ.def.Entity as object, ed as object, axis, dist);
        }

        public void add(Occ o1, Occ o2, double dist = 0, bool axis = false)
        {
            double d = dist;
            if (d == 0) d = -o1.dist;
            acd.Constraints.AddInsertConstraint(o1.def.Entity as object, o2.def.Entity as object, axis, d);
        }

        public void add(FaceProxy fp, Occ occ, double dist = 0, bool axis = false)
        {
            if (fp.SurfaceType == SurfaceTypeEnum.kPlaneSurface)
            {
                FaceProxy f = u.get<FaceProxy>((occ.def.Entity as Edge).Faces, fi => fi.SurfaceType == SurfaceTypeEnum.kPlaneSurface);
                if (!axis) acd.Constraints.AddMateConstraint(f as object, fp as object, dist);
                else acd.Constraints.AddFlushConstraint(f as object, fp as object, dist);
            }
            else if (fp.SurfaceType == SurfaceTypeEnum.kCylinderSurface)
            {
                FaceProxy f = u.get<FaceProxy>((occ.def.Entity as Edge).Faces, fi => fi.SurfaceType != SurfaceTypeEnum.kPlaneSurface);
                acd.Constraints.AddMateConstraint(f as object, fp as object, dist, InferredTypeEnum.kInferredLine, InferredTypeEnum.kInferredLine);
            }
        }
                                                                                     
        static public bool check(EdgeProxy ed)
        {
            bool con = false;
            foreach (AssemblyConstraint constr in ed.ContainingOccurrence.OccurrencePath[1].Constraints)
            {
                try
                {
                    if (constr.ResultOfiMate == true) continue;
                    if (constr.EntityTwo != null && constr.EntityTwo.Equals(ed))
                    {
                        con = true;
                        break;
                    }
                }
                catch { }
            }
            return con;
        }
    }
}
