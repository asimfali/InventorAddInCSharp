using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using u = InvDoc.u;
using System.Xml.Linq;
using Inventor;
using InvDoc;
using System.Text.RegularExpressions;

namespace InvAddIn
{
    class Sketches : Button
    {
        public Sketches(string displayName, string internalName, string clientId, string description, string tooltip, 
        ButtonDisplayEnum buttonDisplayType = ButtonDisplayEnum.kDisplayTextInLearningMode, CommandTypesEnum commandType = CommandTypesEnum.kNonShapeEditCmdType)
        : base(displayName, internalName, commandType, clientId, description, tooltip, buttonDisplayType) { }

        protected override void ButtonDefinition_OnExecute(NameValueMap context)
        {
            new Elements(Macros.StandardAddInServer.m_inventorApplication.ActiveDocument);
        }  
    }

    class Elements
    {
        Application app;
        SelectSet ss;
        PlanarSketch ps;
        string docName;
        static public bool[] chks;
        static public string path = null;
        public Elements(Document doc)
        {
            app = Macros.StandardAddInServer.m_inventorApplication;
            I.silent(true);
            Transaction tr = I.beginTrans("Эскизы");
            if (doc.DocumentType == DocumentTypeEnum.kPartDocumentObject)
            {
                Entity.doc = doc; Entity.pcd = ((PartDocument)doc).ComponentDefinition;
                constraints.pars = Entity.pcd.Parameters;
                u.getDerSolid(Entity.pcd);
            }
            if (doc.DocumentType == DocumentTypeEnum.kAssemblyDocumentObject)
            {
                Entity.doc = doc; Entity.acd = ((AssemblyDocument)doc).ComponentDefinition;
                constraints.pars = Entity.acd.Parameters;
            }
            docName = u.getDocName(doc);

            if (path == null)
                path = I.p() + @"\base.xml";
            readXML(path);
            
            I.silent(false);
            Entity.clearStatic();
            tr.End();
        }
        public void checkSubPath(Document doc, XMLDoc xml)
        {
            string sp = file.p(doc.FullDocumentName);
            sp = sp.TrimEnd(new char[] {'\\'});
            List<XElement> rem = new List<XElement>();
            foreach (var item in xml.El.Elements())
            {
                string s = XMLDoc.getAttributeValue(item, "subPath");
                if (u.isNull(s)) continue;
                s = s.Trim(new char[] {'\\'});
                if (!sp.EndsWith(s)) rem.Add(item);
            }
            foreach (var item in rem)
	        {
		        item.Remove();
	        }
        }
        public void removeBase(XMLDoc xml, string el, string name)
        {
            List<XElement> rem = new List<XElement>();
            u.action<XElement>(xml.find(el), a => rem.Add(a), f => !u.isNull(f.Attribute(name)));
            foreach (var item in rem)
            {
                item.Remove();  
            }
        }
        public PlanarSketch addSketch()
        {
            return (ss[1] as SketchPoint).Parent as PlanarSketch;
        }
        public void readXML(string fn)
        {
            Entity.xml = new XMLDoc(fn, "head");
            u.checkSubPath(Entity.doc, Entity.xml);
            XMLDoc.L = 1;
            Entity.xml.insert(f: new List<string>() { "Assembly", "Add", "Part"});
            u.paramFilter(Entity.doc, Entity.xml);
            //removeBase(Entity.xml, "Plane", "BasePlane");
            removeBase(Entity.xml, "Sketch", "BasePlane");
            XMLDoc.removeFltr(Entity.xml.El);
            XMLDoc.removeAttrib(Entity.xml.El);
            Entity.xml.copy("Name");
            string n = "Sketch";
            Entity.xml.copy(n);
            Entity.pars = addParameters();
            if (chks[16]) { u.updateParameter(I.getParameters(Entity.doc).UserParameters, Entity.pars); return; }
            if (chks[0]) read<PlaneFeat>("Plane");
            if (chks[1]) read<UnfoldFeat>("Unfold");
            if (chks[2]) read<SketchInv>(n);
            if (chks[3]) read<CFlangeFeat>("CFlange");
            if (chks[4]) read<LoftFlangeFeat>("Loft");
            if (chks[5]) read<FaceFeat>("Face");
            if (chks[6]) read<FlangeFeat>("Flange");
            if (chks[7]) read<CutFeat>("Cut");
            if (chks[8]) read<HoleInv>("Hole");
            if (chks[9]) read<FilletFeat>("Fillet");
            if (chks[10]) read<IFeatFeat>("IFeature");
            if (chks[11]) read<PunchFeat>("Punch"); 
            if (chks[12]) read<ArrayFeat>("Array");
            if (chks[13]) read<MirrorFeat>("Mirror");
            if (chks[14]) read<IMatesInv>("IMates");
            if (chks[15]) read<FlatPatternFeat>("FP");
            if (!u.isNull(Entity.acd)) read<AssemblyConstrInv>("AsmConstr");
            Entity.pars = null;
        }
        public Dictionary<string, XElement> addParameters()
        {
            Dictionary<string, XElement> dic = new Dictionary<string, XElement>();
            //u.linkParameters(Entity.doc, Entity.xml, docName);
            //XMLDoc comm = new XMLDoc(I.p() + @"\Common.xml", "head");
            addToDic(Entity.xml, ref dic);
            //if (!u.isNull(comm)) addToDic(comm, ref dic);
            return dic;
            //u.addParameters(Entity.doc, Entity.xml, false);
//             Entity.xml.setRoot();
//             HashSet<string> pnames = new HashSet<string>();
//             foreach (var item in Entity.xml.find("Parameter"))
//             {
//                 string name = XMLDoc.getAttributeValue(item, "Name");
//                 if (name == null) continue;
//                 if (pnames.Contains(name)) continue;
//                 pnames.Add(name);
//                 if (!u.findParameter(Entity.doc,name, true, item))
//                     u.addParameter(Entity.doc, item);
//             }
        }

        public void addToDic(XMLDoc xml, ref Dictionary<string, XElement> dic)
        {
            foreach (var item in xml.find("Parameter"))
            {
                var fltr = XMLDoc.getSpl(item, "Fltr", ';');                
                string name = XMLDoc.getAttributeValue(item, "Name");
                if (dic.ContainsKey(name))
                {
                    if (fltr != null && u.checkFltr(fltr, docName.ToLower()))
                    {
                        dic[name] = item;
                    }
                    continue;
                } 
                dic.Add(name, item);
            }
        }
        
        public void read<T>(string n) where T: Entity, new()  
        {
            Entity.xml.setRoot();
            if (typeof(T) == typeof(UnfoldFeat)) Entity.setLastPos();
            foreach (var el in Entity.xml.find(n))
            {
                Entity.setUpd(false);
                
                var fltr = XMLDoc.getSpl(el, "Fltr", ';');
                if (docName.ToLower().IndexOf("base") == -1)
                {
                    if (fltr != null && !u.checkFltr(fltr, docName.ToLower())) continue;
                }
                else if (!(typeof(T) == typeof(SketchInv) || typeof(T) == typeof(PlaneFeat)))
                {
                    continue; 
                }
                if ((typeof(T) == typeof(SketchInv) || typeof(T) == typeof(PlaneFeat)) && !u.isNull(XMLDoc.getAttributeValue(el, "no")) && docName.ToLower().IndexOf("base") != -1) 
                    continue;
                if (u.isNull(XMLDoc.getAttributeValue(el, "Unfold")))
                {
                    if (!u.isNull(Entity.endOfPart)) continue;
                }
//                 else if (typeof(T) != typeof(UnfoldFeat))
//                 {
//                     var fe = I.getSMF(Entity.doc);
//                     if (fe.UnfoldFeatures.Count != 1) continue;
//                     if (!u.checkEndOfPart(Entity.doc, Entity.pcd, "Разв")) continue;
//                 }
                Entity.br = false;
                Entity.xml.El = el;
                string nme = Entity.xml.getAttributeValue("Name");
                //if (!(XMLDoc.isNull(nme)) && nme == "SketchName.x") continue;   
                Entity.hasFirst = false;
                T sk = new T();
                if (Entity.br) continue;
                sk.add();
                
            }
            Entity.br = false;
        }
        public List<string> getNames(string n)
        {
            List<string> names = new List<string>();
            string name = Entity.xml.getAttributeValue(n);
            if (name.IndexOf(";") == -1)
            {
                names.Add(name);
            }
            else
            {
                var spl = name.Split(';');
                names = spl.ToList();
            }
            return names;
        }
        public void createXML()
        {
            string p = u.pathDoc(Entity.doc), name = ps.Name + ".xml";
            Entity.xml = new XMLDoc(p + "\\" + name, "head");
            SketchInv sk = new SketchInv(ps);
            sk.get();
            if (Entity.closed)
            {
                XElement el = Entity.xml.findLast("Sketch");
                if (el != null) el.SetAttributeValue("Constr", "c");
            }
            Entity.xml.save();
        }
    }

    enum entTypes
    {
        Point, Line, Arc, Circle, Slot, Sketch, Proj, Plane, Hole, Flange, Fillet, Cut, Face, Part, Asm, Features, CountourFlange,
        Parameter, IMate, IMates, Mirror, Array, iFeature, Punch, Unfold, Loft, FlatPattern, Extrude, Revolve, AsmConstr, Cluster
    }
    enum cType
    {
        n = 0,
        t = 0x01, // Tangent
        c = 0x02, // Close
        hole = 0x04, // HoleCenter
        cen = 0x08, // Center
        v = 0x10, // Vertical
        h = 0x20, // Horizontal
        eq = 0x40, // Equal
        f = 0x80, // First
        par = 0x100, // Parallel
        perp = 0x200, // Perpendicular
        cl = 0x400, // CenterLine
        con = 0x800, // Coinsident
        col = 0x1000, // Collinear
    }
    enum prType
    {
        WA, WP, Line, Arc, Circle, Point
    }

    enum posType
    {
        s, m, e,
    }

    enum xDirType
    {
        s, a, e, f,
    }

    enum edgeType
    {
        l, a, c,
    }
    enum faceType
    {
        cyl, pl
    }

    abstract class Entity
    {
        public entTypes t;
        public cType con;
        protected string alias;
        protected Vector2d dir;
        static protected bool update = false;
        static public bool closed = false;
        static public bool br = false;
        protected double x, y, bx = 0, by = 0;
        protected string name = "", form = "0.###"/*, con*/;
        public SketchEntity ent/*, prev*/;
        public Entity prev;
        static public Entity fEnt;
        protected SketchPoint sp;
        static protected PlanarSketch ps;
        public SketchPoint startPt, endPt;
        protected SketchLine sl;
        protected PlanarSketch ips;
        protected WorkPoint wp;
        protected WorkAxis wa;
        protected WorkPlane wpl;
        public bool addAngle = false, unfold = false;
        static public object endOfPart;
        public List<DimensionConstraint> constrts = new List<DimensionConstraint>();
        static public Dictionary<string, XElement> pars;

        public SketchLine getSl
        {
            get { return sl; }
            set { sl = value; }
        }
        public string getName
        {
            get { return name; }
        }
        public static Document doc;
        public static PartComponentDefinition pcd;
        public static AssemblyComponentDefinition acd;
        public static XMLDoc xml;
        public bool first = false, last = false;
        public static bool hasFirst = false;
        public SketchPoint bp = null;
        public SketchEntity xent, yent;
        protected string angs;
        public string noDim;

        public virtual void add()
        {
            if (!update)
                draw();
            else
                upd();
        }
        public abstract void draw();
        public abstract void upd();
        public virtual void get()
        {
            Con();
            if (con.HasFlag(cType.c)) last = true;
            setName();
        }
        static public void setUpd(bool v)
        {
            update = v;
        }
        static public object setLastPos()
        {
            object a;
            pcd.GetEndOfPartPosition(out endOfPart, out a);
            return a;
        }
        protected bool checkAsm()
        {
            if (doc.DocumentType == DocumentTypeEnum.kAssemblyDocumentObject)
            {
                br = true;
            }
            return doc.DocumentType != DocumentTypeEnum.kAssemblyDocumentObject;
        }
        static public void retLastPos()
        {
            if (endOfPart == null) return;
            PartFeature tf = endOfPart as PartFeature;
            if (!u.isNull(tf)) tf.SetEndOfPart(false);
            PlanarSketch ps = endOfPart as PlanarSketch;
            if (!u.isNull(ps)) ps.SetEndOfPart(false);
            WorkPlane wp = endOfPart as WorkPlane;
            if (!u.isNull(wp)) wp.SetEndOfPart(false);
        }
        protected void addNoDim()
        {
            noDim = xml.getAttributeValue("NoDim");
        }
        static public void clearStatic()
        {
            xml = null; doc = null; hasFirst = false; fEnt = null; ps = null; update = false; closed = false; br = false; acd = null; pcd = null;
            retLastPos();
            endOfPart = null;
        }
        public virtual void setName()
        {
            name = xml.getAttributeValue("Name", true);
            if (alias == null || alias == "") return;
            if (name == null || name == "")
            {
                name = xml.getAttributeValue("Sketch");
                if (name == null) return;
                name = alias + name;
            }
        }
        protected bool checkSketch()
        {
            if (ips == null || ips.SketchEntities.Count == 0) return true;
            return false;
        }
        public Entity(entTypes t, double x = 0, double y = 0)
        {
            this.x = x; this.y = y; this.t = t;
        }
        public void Con()
        {
            string tmp = xml.getAttributeValue("Constr");
            if (isNull(tmp)) return;
            tmp = tmp.Trim();
            tmp = tmp.Replace(' ',',');
            tmp = replace(tmp, ',');
            Enum.TryParse(tmp, out con);
        }
        public posType[] Pos()
        {
            string[] tmp = u.getSpl(xml.getAttributeValue("Pos"), ';');
            if (isNull(tmp)) return null;
            posType[] r = new posType[tmp.Length];
            for (int i = 0; i < tmp.Length; i++)
            {
                posType pt;
                Enum.TryParse(tmp[i], out pt);
                r[i] = pt;
            }
            return r;
        }
        public string replace(string val, char sym)
        {
            List<char> lst = new List<char>();
            bool add = true;
            int i = 0;
            foreach (var item in val)
            {
                if (item.Equals(sym))
                {
                    i++;
                }
                else i = 0;
                if (i > 1) add = false;
                else add = true;
                if (add) lst.Add(item);
            }
            string r = new string(lst.ToArray());
            return r;
        }
        public IEnumerable<Face> getPlaneFaces(Faces fs)
        {
            var smcd = I.getSMCD(doc);
            if (isNull(smcd)) return null;
            double thick = smcd.Thickness.ModelValue;
            if (isNull(fs)) return null;
            var fa = u.gets<Face>(fs, fi => fi.SurfaceType == SurfaceTypeEnum.kPlaneSurface);
            fa = u.gets<Face>(fa, fi =>
            {
                Edge ed = u.get<Edge>(fi.Edges, e =>
                {
                    return (e.GeometryType == CurveTypeEnum.kLineSegmentCurve) && u.eq(u.getLenght(e), thick);
                });
                return isNull(ed);
            });
            return fa;
        }
        public WorkPlane findWP(string name, PartComponentDefinition pcd = null)
        {
            return u.findPlane(doc, name, pcd);
            //return u.get<WorkPlane>(pcd.WorkPlanes, f => f.Name == name);
        }

        public WorkAxis findWA(string name)
        {
            int i = u.getIntElem(name, 0);
            if (isNull(i))
            return u.get<WorkAxis>(pcd.WorkAxes, f => f.Name == name);
            else
            {
                return u.get(pcd.WorkAxes, i) as WorkAxis;
            }
        }
        public virtual void getCoord(string name)
        {
            if (!isNull(angs))
            {
                dir = getVector(angs);
            }
            else if (xml.getCoord(name))
            {
                xml.getDir(out x, out y);
                dir = I.CV2d(x, y);
            }
            else
            {
                xml.getDir(out x, out y);
                dir = I.CV2d(x, y);
            }
        }
        public Vector2d getVector(string ang)
        {
            double r = 1;
            double a = u.degToRad(u.convToDouble(ang));
            double x = r * Math.Cos(a), y = r * Math.Sin(a);
            return I.CV2d(x, y);
        }
        protected void merge(Entity e, SketchPoint bp, ObjectCollection col = null)
        {
            if (isNull(e)) return;
           
            SketchPoint sp = null;
            if(e.t == entTypes.Line)
                sp = e.startPt;
            bp.HoleCenter = false;
            if (isNull(bp, sp)) return;
            Vector2d v = sp.Geometry.VectorTo(bp.Geometry);
            if (!isNull(col) && v.Length > 1)
            {
                ps.MoveSketchObjects(col, v);
            }
            sp.Merge(bp);
        }
        public void getBase()
        {
            string bsketch = xml.getAttributeValue("bSketch");
            if (!isNull(bsketch)) getFromSketch(bsketch);
            string bplane = xml.getAttributeValue("bPlane");
            if (!isNull(bplane) && bplane != ";") getFromPlane(bplane);
            string bent = xml.getAttributeValue("bEnt");
            if (!isNull(bent)) getFromEntity(bent);
        }
        public List<SketchEntity> getFromFace(string bface)
        {
            Face f = ps.PlanarEntity as Face;
            if (isNull(f)) return null;
            var spl = u.getSpl(bface, ';');
            List<SketchEntity> r = new List<SketchEntity>();
            foreach (var item in spl)
            {
                string snum = getElem(item, 0);
                List<CurveTypeEnum> lst = getCurveType(item);
                var teds = u.gets<Edge>(f.Edges, fil => filterEdge(fil, lst));
                teds = filterMinMax(item, teds);
                int num; int.TryParse(snum, out num);
                if (isNull(num))
                {
                    u.action<Edge>(teds, a => addSE(a, ref r));
                }
                else
                {
                    Edge ed = u.get(teds, num) as Edge;
                    addSE(ed, ref r);
                }
            }
            return r;
        }
        public void addSE(object ed, ref List<SketchEntity> r)
        {
            //if (isNull(ed)) return;
            SketchEntity pr = ps.AddByProjectingEntity(ed);
            SketchLine li = pr as SketchLine;
            if (isNull(li)) return;
            li.Construction = true; li.StartSketchPoint.HoleCenter = false; li.EndSketchPoint.HoleCenter = false;
            r.Add(pr);
        }
        public IEnumerable<Edge> filterMinMax(string f, IEnumerable<Edge> eds)
        {
            if (isNull(f) || f.IndexOf(":") == -1) return eds;
            double min = getDouble(getElem(f, 2), 0.1, 2), max = getDouble(getElem(f, 3), 0.1, 2);
            return u.gets<Edge>(eds, ted => {
                double l = u.getParam(ted);
                return l >= min && l <= max; 
            });
        }
        protected void checkEnt(SketchEntity pr)
        {
            if (pr.Type == ObjectTypeEnum.kSketchLineObject)
            {
                pr.Construction = true;
                if (isX(pr.RangeBox)) { xent = pr; bx = pr.RangeBox.MaxPoint.X; }
                else { yent = pr; by = pr.RangeBox.MaxPoint.Y; }
            }
        }
        public void getFromPlane(string bplane)
        {
            var spl = u.getSpl(bplane, ';');
            foreach (var item in spl)
            {
                WorkPlane twp = u.findPlane(doc, item);
                if (isNull(twp)) { br = true; return; }
                twp.Visible = false;
                SketchEntity pr = ps.AddByProjectingEntity(twp);
                if (isNull(pr)) continue;
                checkEnt(pr);
            }
        }
        public bool isX(Box2d b)
        {
            return u.eq(b.MaxPoint.X, b.MinPoint.X);
        }
        public void getFromSketch(string bsketch)
        {
            u.findSketch(doc, bsketch);
            PlanarSketch tps = u.findInCol<PlanarSketch>(pcd.Sketches, e => e.Name == bsketch);
            string num = xml.getAttributeValue("num");
            if (isNull(tps, num)) { br = true; return; }

            var tmp = u.gets<SketchPoint>(tps.SketchPoints, num, ';');
            if (tmp != null && tmp.Count() == 1)
            {
                SketchPoint tsp = tmp.ElementAt(0);
                if (isNull(tsp)) return;
                SketchEntity pr = u.get<SketchEntity>(ps.SketchEntities, e => !u.isNull(e.ReferencedEntity) && e.ReferencedEntity.Equals(tsp));
                if (isNull(pr))
                pr = ps.AddByProjectingEntity(tsp);
                if (isNull(pr)) return;
                if (pr.Type == ObjectTypeEnum.kSketchPointObject)
                {
                    bp = pr as SketchPoint; getBP(); bp.HoleCenter = false;
                }
            }
        }
        public void getFromEntity(string bent)
        {
            IEnumerable<Face> pf = find(bent);
            IEnumerable<Face> fs = null;
            if (isNull(pf)) return;
            string snum = xml.getAttributeValue("eNum");
            int num = u.getIntElem(snum, 0); faceType ft;
            string t = u.getElem(snum, 1);
            Enum.TryParse(t, out ft);
            if (isNull(t)) return;
            switch (ft)
            {
                case faceType.cyl:
                    fs = u.gets<Face>(pf, el => el.SurfaceType == SurfaceTypeEnum.kCylinderSurface);
                    if (isNull(fs.Count())) return;
                    Face tmpf = u.get(fs, num) as Face;
                    Edge ed = tmpf.Edges[1];
                    SketchEntity pr = ps.AddByProjectingEntity(ed);
                    if (isNull(pr)) return;
                    if (pr.Type == ObjectTypeEnum.kSketchCircleObject)
                    {
                        bp = ((SketchCircle)pr).CenterSketchPoint; getBP(); bp.HoleCenter = false;
                    }
                    break;
                case faceType.pl:
                    fs = u.gets<Face>(pf, el => el.SurfaceType == SurfaceTypeEnum.kPlaneSurface);
                    break;
                default:
                    break;
            }
            
        }
        public IEnumerable<Face> find(string name)
        {
            IEnumerable<Face> fs = null;
            PartFeature pf = InvDoc.Reflect.exist<PartFeature>(pcd.Features, fi => fi.Name == name);
            if (isNull(pf))
            {
                fs = refEntity(pcd.SurfaceBodies[1].Faces, name);
                return fs;
            }
            return u.gets<Face>(pf.Faces, el => true);
        }
        public virtual IEnumerable<T> lastEntity<T>(System.Collections.IEnumerable col = null, string count = "1", bool last = true) where T: class
        {
            if (isNull(col)) col = pcd.Sketches;
            IEnumerable<T> ie = col.OfType<T>(), fie;
            if (isNull(ie)) yield break;
            var spl = XMLDoc.getSpl(count, ';');
            if (isNull(spl)) yield break;
            foreach (var item in spl)
            {
                List<ObjectTypeEnum> lst = getObType(item);
                fie = ie.Where(el => filterEnt(el, lst));
                fie = fie.Where(e => (e as PartFeature).HealthStatus != HealthStatusEnum.kBeyondStopNodeHealth);
                int c = fie.Count();
                int i = int.Parse(getElem(item, 0));
                if (last && checkLast<T>(fie)) 
                    i++;
                if (isNull(i) || i > c) continue;
                yield return fie.ElementAt(c - i);
            }
        }
        public virtual IEnumerable<Face> refEntity(IEnumerable<Face> col = null, string count = "1")
        {
            var spl = XMLDoc.getSpl(count, ';');
            if (isNull(spl)) yield break;
            foreach (var item in spl)
	        {
                List<ObjectTypeEnum> lst = getObType(item);
                foreach (Face fa in col)
                {
                    Face f = fa;
                    if (!isNull(f.ReferencedEntity)) f = fa.ReferencedEntity as Face;
                    if (lst.Contains(f.CreatedByFeature.Type))
                        yield return fa;
                }
            }
        }

        public virtual IEnumerable<Face> refEntity(Faces fs, string name)
        {
            foreach (Face item in fs)
            {
                Face f = item;
                if (!isNull(f.ReferencedEntity)) f = item.ReferencedEntity as Face;
                if (f.CreatedByFeature.Name == name) yield return item; 
            }
        }

        public bool checkLast<T>(IEnumerable<T> ie)
        {
            object ob = ie.LastOrDefault();
            if (isNull(ob)) return false;
            return ob is FilletFeature && t == entTypes.Fillet ? true:
                ob is RectangularPatternFeature && t == entTypes.Array ? true:
                ob is CutFeature && t == entTypes.Cut ? true:
                ob is HoleFeature && t == entTypes.Hole ? true:
                ob is FaceFeature && t == entTypes.Face ? true:
                ob is MirrorFeature && t == entTypes.Mirror ? true:
                ob is FlangeFeature && t == entTypes.Flange ? true:
                ob is FilletFeature && t == entTypes.Fillet ? true:
                ob is iFeature && t == entTypes.iFeature ? true:
                ob is PunchToolFeature && t == entTypes.Punch ? true:
                ob is UnfoldFeature && t == entTypes.Unfold ? true:
                ob is ContourFlangeFeature && t == entTypes.CountourFlange ? true:
                ob is LoftedFlangeFeature && t == entTypes.Loft ? true:
                ob is ExtrudeFeature && t == entTypes.Extrude ? true :
                ob is RevolveFeature && t == entTypes.Revolve ? true :
                false;
        }

        public string getElem(string v, int num, char sep = ':')
        {
            var spl = XMLDoc.getSpl(v, sep);
            if (spl == null) return null;
            if (spl.Length <= num) return null;
            return spl[num];         
        }

        public List<CurveTypeEnum> getCurveType(string item)
        {
            List<CurveTypeEnum> lst = new List<CurveTypeEnum>();
            string v = getElem(item, 1);
            if (isNull(v)) { lst.Add(CurveTypeEnum.kLineSegmentCurve); return lst; }
            edgeType e; Enum.TryParse(v, out e);
            switch (e)
            {
                case edgeType.l:
                    lst.Add(CurveTypeEnum.kLineSegmentCurve);
                    break;
                case edgeType.a:
                    lst.Add(CurveTypeEnum.kCircularArcCurve);
                    break;
                case edgeType.c:
                    lst.Add(CurveTypeEnum.kCircleCurve);
                    break;
                default:
                    break;
            }
            return lst;
        }

        public List<ObjectTypeEnum> getObType(string item)
        {
            List<ObjectTypeEnum> lst = new List<ObjectTypeEnum>();
            string v = getElem(item, 1);
            if (isNull(v)) return lst;
            switch (v)
            {
                case "Г":
                    lst.Add(ObjectTypeEnum.kFaceFeatureObject);
                    break;
                case "Ф":
                    lst.Add(ObjectTypeEnum.kFlangeFeatureObject);
                    break;
                case "З":
                    lst.Add(ObjectTypeEnum.kMirrorFeatureObject);
                    break;
                case "М":
                    lst.Add(ObjectTypeEnum.kRectangularPatternFeatureObject);
                    lst.Add(ObjectTypeEnum.kCircularPatternFeatureObject);
                    break;
                case "К":
                    lst.Add(ObjectTypeEnum.kContourFlangeFeatureObject);
                    break;
                case "В":
                    lst.Add(ObjectTypeEnum.kCutFeatureObject);
                    break;
                case "Пр":
                    lst.Add(ObjectTypeEnum.kPunchToolFeatureObject);
                    break;
                case "П":
                    lst.Add(ObjectTypeEnum.kiFeatureObject);
                    break;
                case "С":
                    lst.Add(ObjectTypeEnum.kFilletFeatureObject);
                    break;
                case "О":
                    lst.Add(ObjectTypeEnum.kHoleFeatureObject);
                    break;
                case "Р":
                    lst.Add(ObjectTypeEnum.kUnfoldFeatureObject);
                    break;
                case "Л":
                    lst.Add(ObjectTypeEnum.kLoftedFlangeFeatureObject);
                    break;
                case "Вр":
                    lst.Add(ObjectTypeEnum.kRevolveFeatureObject);
                    break;
                case "Выд":
                    lst.Add(ObjectTypeEnum.kExtrudeFeatureObject);
                    break;
                default:
                    break;
            }
            return lst;
        }

        public bool filterEdge(Edge e, List<CurveTypeEnum> lst)
        {
            if (isNull(e)) return false;
            return (lst.Contains(e.GeometryType));
        }

        public bool filterEnt(object ob, List<ObjectTypeEnum> lst)
        {
            PartFeature pf = ob as PartFeature;
            if (!isNull(pf))
            return lst.Contains(pf.Type);
            Face f = ob as Face;
            if (f.HasReferenceComponent) f = f.ReferencedEntity as Face;
            if (!isNull(f))
                return lst.Contains(f.CreatedByFeature.Type);
            return false;
        }
        public string conToString()
        {
            return con.ToString().Replace(",","");
        }
        public bool check<T>(T ob)
        {
            return (ob != null) ? true : false;
        }
        protected void getBP()
        {
            if (bp != null)
            {
                bx = bp.Geometry.X; by = bp.Geometry.Y;
            }
        }
        protected PlanarSketch getSketch()
        {
            string tn;
            if (name == null || name == "") tn = xml.getAttributeValue("Name");
            else tn = name;

            if (t == entTypes.iFeature || t == entTypes.Punch || t == entTypes.Hole || t == entTypes.Cut)
            {
                PlanarSketch tsk = u.findInCol<PlanarSketch>(pcd.Sketches, e => e.Name == "_" + tn);
                if (!isNull(tsk)) return null;
            }
            u.findSketch(doc, tn);
            return u.findInCol<PlanarSketch>(pcd.Sketches, e => e.Name == tn);
        }
        protected bool isAlfa(string ts, out bool sign)
        {
            sign = false;
            if (isNull(ts)) return false;
            if (ts.StartsWith("-"))
            {
                sign = true;
                return Char.IsLetter(ts[1]);
            }
            return Char.IsLetter(ts[0]);
        }
        protected double getDouble(string ts, double scale, int count)
        {
            bool sign;
            if (isAlfa(ts, out sign))
            {
                ts = ts.TrimStart('-');
                u.findParameter(doc, ts);
                Parameter p = u.get<Parameter>(pcd.Parameters, e => e.Name == ts);
                if (p == null)
                {
                    u.addParameter(Entity.doc, ts, Entity.pars);
                    p = u.get<Parameter>(pcd.Parameters, e => e.Name == ts);
                    if (p == null) return 0;
                }
                double v = p.ModelValue * 10;
                return sign ? -v : v;
            }
            else return u.convToDouble(ts, scale, count);
        }
        protected int getInt(string val)
        {
            bool sign;
            string ts = xml.getAttributeValue(val);
            if (ts == null) return default(int);
            if (isAlfa(ts, out sign))
            {
                Parameter p = u.get<Parameter>(pcd.Parameters, e => e.Name == ts);
                if (p == null) return 0;
                return (int)p.ModelValue;
            }
            else return u.convToInt(ts);
        }
        protected bool isNull(params object[] objs)
        {
            if (objs == null) return true;
            foreach (var ob in objs)
            {
               if (ob is string && (ob == null || (string)ob == "")) return true;
               if (ob is int && (ob == null || (int)ob == 0)) return true;
               if (ob is double && (ob == null || (double)ob == 0)) return true;
               if (ob == null) return true;
            }
            return false;
        }
        protected void vis()
        {
            if (!isNull(ips)) ips.Visible = false;
            if (!isNull(wp)) wp.Visible = false;
            if (!isNull(wa)) wa.Visible = false;
            if (!isNull(wpl)) wpl.Visible = false;
        }
        protected bool checkPF(string v)
        {
            if (isNull(v)) return true;
            PartFeature mf = u.get<PartFeature>(pcd.Features, fi => fi.Name == v);
            if (isNull(mf)) return false;
            return true;
        }
        protected Face numFace(XElement el)
        {
            string n = XMLDoc.getAttributeValue(el,"num"), l = XMLDoc.getAttributeValue(el,"Last"); //new
            if (isNull(l)) return null;
            IEnumerable<Face> fa = null;
            PartFeature pf = null;

            if (!isNull(getElem(l, 2)))
            {
                fa = refEntity(u.gets<Face>(pcd.SurfaceBodies[1].Faces, fi => fi.CreatedByFeature.Type == ObjectTypeEnum.kReferenceFeatureObject), l);
                fa.Count();
                if (fa.Count() == 0) return null;
            }
            else
            {
                pf = lastEntity<PartFeature>(pcd.Features, l).LastOrDefault();
                if (isNull(pf)) return null;
            }
            if (isNull(fa)) fa = getPlaneFaces(pf.Faces);
            if (isNull(fa)) return null;
            int nu = u.getNum(n, fa.Count());
            string s = u.getElem(n, 1);
            if (!isNull(s)) fa = u.sort<Face>(fa, s);
            //u.clientTxt(fa, pcd);
            //I.app.ActiveView.Update();
            if (isNull(nu)) return null;
            if (fa.Count() < nu) return null;
            Face fp = fa.ElementAt(nu - 1);
            return fp;
        }
    }

    enum units
    {
        mm, ul,
    }

    class ParamInv
    {
        Parameter p;
        const string elName = "Parameter";
        static public XMLDoc xml;
        string name;
        string val;
        units uts = units.mm;
        static public string group;
        public string comment;

        public string Val
        {
            get { return val; }
            set { val = value; }
        }

        public string Name
        {
            get { return name; }
            set { name = value; }
        }

        public Parameter P
        {
            get { return p; }
            set { p = value; }
        }
        List<ParamInv> ps = new List<ParamInv>();
        public entTypes t;
        public ParamInv this[int index]
        {
            get { return ps[index]; }
            set { ps.Insert(index, value); }
        }

        public ParamInv this[string name]
        {
            get { return ps.FirstOrDefault(p => p.Name == name); }
        }
        public ParamInv(Parameter p)
        {
            this.t = entTypes.Parameter; Name = p.Name; val = p.Expression;
            get(p);
        }
        public void get(Parameter p)
        {
            foreach (Parameter item in p.DrivenBy)
            {
                get(item);
            }
            ps.Add(new ParamInv(p));
        }
        
        public override string ToString()
        {
            if (xml.El.Name != elName)
                xml.El = xml.findLast(elName);
            xml.addXElement(elName, new Dictionary<string, string> { { "Name", "" }, { "Value", val }, 
            {"Type", uts.ToString()} ,{"Group", group},{"Comment",comment}});
            return xml.El.ToString();
        }
    }

    class PointInv : Entity
    {
        bool hole;
        double[] Ls;
        double offset;

        public PointInv():    
            base(entTypes.Point)
        {
            get();
        }
        public PointInv(SketchPoint sp): base(entTypes.Point)
        {
            base.sp = sp; startPt = sp; endPt = sp; ent = (SketchEntity)sp; 
        }

        public SketchPoint draw(SketchPoint sp,double L)
        {
            dir.Normalize();
            dir.ScaleBy(L);
            Point2d pt = sp.Geometry;
            pt.TranslateBy(dir);
            return ps.SketchPoints.Add(pt);
        }
        public void constr(SketchLine sl)
        {
            constraints.addConsid(sl as SketchEntity, ent);
        }
        public void dimConstr(SketchPoint prev)
        {
            sp = ent as SketchPoint;
            constraints.addTwoPointDist(sp, prev, name, offset);
        }
        public override void draw()
        {
            sp = ps.SketchPoints.Add(u.createPoint2d(x, y));
            SketchPoint nsp = null;
            sl = sl = ps.SketchLines.AddByTwoPoints(sp, draw(sp, Ls.Sum()));
            startPt = sl.StartSketchPoint; endPt = sl.EndSketchPoint;
            offset = sl.Length * (-0.25);
            constraints.addTwoPointDist(sl, name, offset);
            sl.Construction = true;
            offset = sl.Length * (-0.2);
            for (int i = 0; i < Ls.Length - 1; i++)
            {
                nsp = draw(sp, Ls[i]);
                ent = (SketchEntity)nsp;
                constr(sl);
                dimConstr(sp);
                sp = nsp;
            }
        }

        public override void upd()
        {
            throw new NotImplementedException();
        }

        public override void get()
        {
//             string s = xml.getAttributeValue("Hole");
//             hole = s == "" || s == "0" ? false : true;
            if (update) return;
            getBase();
            if (xml.getCoord(bx, by))
            {
                //first = true; 
                double tx = getDouble(xml.xname, 1, 3), ty = getDouble(xml.yname, 1, 3);
                sp = ps.SketchPoints.Add(u.createPoint2d(tx / 10, ty / 10));
                sp.HoleCenter = true;
                if (isNull(bp)) bp = sp;
                //prev = new PointInv(sp);
            }
//             getCoord("Dir");
//             if (xml.getLen()) Ls = xml.Ls;
        }
    }

    class CircleInv : Entity
    {
        string rs;
        double r;
        SketchCircle cir;
        public CircleInv()
            : base(entTypes.Circle)
        {
            get();
            if (isNull(ps)) return;
        }

        public override void draw()
        {
            if (isNull(ps)) return;
            cir = ps.SketchCircles.AddByCenterRadius(prev.endPt, r/10);
            ent = cir as SketchEntity;
            endPt = cir.CenterSketchPoint;
            constraints.addRadius(ent ,u.getPointAtParam(cir.Geometry.Evaluator, 0.7), cir.Geometry.Center, name, 1, 0, 0.2);
            constraints.setVal(rs);
        }

        public override void upd()
        {

        }

        public override void get()
        {
            base.get();
            addNoDim();
            rs = xml.getAttributeValue("R");
            r = getDouble(rs, 0.1, 3);
            getCoord("Dir");
            if (!hasFirst) { getBase(); hasFirst = true; }
            if (xml.getCoord(bx, by))
            {
                first = true;
                double tx = getDouble(xml.xname, 1, 3), ty = getDouble(xml.yname, 1, 3);
                sp = ps.SketchPoints.Add(u.createPoint2d(tx / 10 + bx, ty / 10 + by));
                sp.HoleCenter = false;
                if (isNull(bp)) bp = ps.SketchPoints.Add(I.CP2d(tx / 10 + bx, ty / 10 + by));
                else if (bp.ReferencedEntity != null)
                {
                    SketchInv.origin = bp;
                    bp = ps.SketchPoints.Add(I.CP2d(tx / 10 + bx, ty / 10 + by));
                }
                prev = new PointInv(sp);
            }
        }
    }

    class SlotInv : Entity
    {
        string rs, ls;
        double r, L;
        public SlotInv()
            : base(entTypes.Slot)
        {
            if (update) return;
            get();
            //ps = getSketch();
            if (isNull(ps)) return;
        }

        public override void draw()
        {
            if (isNull(ps)) return;
            L *= 0.5;
            dir.Normalize();
                if (L != 0)
                    dir.ScaleBy(L);
                Point2d pt = prev.endPt.Geometry; pt.TranslateBy(dir);
                var slo = ps.AddStraightSlotBySlotCenter(prev.endPt, pt, r);
                var gr = constraints.addFixed(prev.endPt as SketchEntity);
                bool fst = true;  
                foreach (SketchEntity item in slo)
                {
                    SketchLine sl = item as SketchLine;
                    if (!isNull(sl) && sl.Construction)
                    {
                        constraints.addTwoPointDist(sl, "", 2);
                        constraints.setVal(ls);
                        if (!isNull(bp))
                        {
                            
                            if (gr.Deletable) gr.Delete();
                            gr = constraints.addFixed(bp as SketchEntity);
                            constraints.addMidPoint(bp, sl);
                            if (gr.Deletable) gr.Delete();
                        }
                        if (con.HasFlag(cType.h)) constraints.addHor(item);
                        else if (con.HasFlag(cType.v)) constraints.addVert(item);
                    }
                    SketchArc sa = item as SketchArc;
                    if (!isNull(sa) && fst)
                    {
                        constraints.addRadius(sa as SketchEntity, sa.StartSketchPoint.Geometry, sa.EndSketchPoint.Geometry , "");
                        constraints.setVal(rs);
                        fst = false;
                    }
                }
        }

        public override void upd()
        {
        }

        public override void get()
        {
            base.get();
            rs = xml.getAttributeValue("R");
            ls = xml.getAttributeValue("L");
            r = getDouble(rs, 0.1, 3);
            L = getDouble(ls, 0.1, 3);
            getCoord("Dir");
            if (!hasFirst) { getBase(); hasFirst = true; }
            if (xml.getCoord(bx, by))
            {
                first = true;
                double tx = getDouble(xml.xname, 1, 3), ty = getDouble(xml.yname, 1, 3);
                sp = ps.SketchPoints.Add(u.createPoint2d(tx / 10 + bx, ty / 10 + by));
                sp.HoleCenter = false;
                if (isNull(bp)) bp = ps.SketchPoints.Add(I.CP2d(tx / 10 + bx, ty / 10 + by));
                else if (bp.ReferencedEntity != null)
                {
                    SketchInv.origin = bp;
                    bp = ps.SketchPoints.Add(I.CP2d(tx / 10 + bx, ty / 10 + by));
                }
                prev = new PointInv(sp);
            }
        }
    }
    
    class ProjInv : Entity
    {
        string snum;
        prType type;
        int num; 
        public ProjInv(): base(entTypes.Proj)
        {          
            get();
        }
        public bool check(PlanarSketch sk)
        {
            return sw(sk, a => checkPr(a));
        }

        public bool sw(PlanarSketch sk, Func<object, bool> a)
        {
            switch (type)
            {
                case prType.WA:
                    return a(getEl<WorkAxis>(pcd.WorkAxes));
                case prType.WP:
                    return a(getEl<WorkPlane>(pcd.WorkPlanes));
                case prType.Line:
                    return a(getEl<SketchLine>(sk.SketchLines));
                case prType.Arc:
                    return a(getEl<SketchArc>(sk.SketchArcs));
                case prType.Circle:
                    return a(getEl<SketchCircle>(sk.SketchCircles));
                case prType.Point:
                    return a(getEl<SketchPoint>(sk.SketchPoints));
                default:
                    return false;
            }
        }

        public override void draw()
        {
            if(isNull(ps)) return;
            PlanarSketch sk;
            if (!isNull(name))
                sk = getSketch();
            else sk = ps;
            if (check(sk)) return;
            sw(sk, a => addPr(a));
        }
        public object getEl<T>(System.Collections.IEnumerable col) where T: class
        {
            if (typeof(T) == typeof(SketchLine) || typeof(T) == typeof(SketchArc) || typeof(T) == typeof(SketchCircle)
                || typeof(T) == typeof(SketchPoint))
            {
                var ie = u.gets<T>(col, fi => ((bool)Reflect.getProp_<T>(fi, "Construction")) == false);
                return u.get(ie, num);
            }
            return isNull(num) ? u.get<T>(col, snum) :
                u.get(col, num);
        }

        public bool addPr(object el)
        {
            if (isNull(el)) return false;
            xent = ps.AddByProjectingEntity(el); xent.Construction = true;
            SketchPoint tsp = xent as SketchPoint;
            if (!isNull(tsp)) tsp.HoleCenter = false;
            return true;
        }

        public bool checkPr(object el)
        {
            var ents = u.gets<SketchEntity>(ps.SketchEntities, f => f.ReferencedEntity != null);
            SketchEntity se = u.get<SketchEntity>(ents, f => f.ReferencedEntity.Equals(el));
            if (isNull(se)) return false;
            //if (se.Construction == true) return false;
            return true;
        }

        public override void upd()
        {
        }

        public override void get()
        {
            snum = xml.getAttributeValue("Name");
            name = xml.getAttributeValue("Sketch");
            if (!isNull(snum)) int.TryParse(snum, out num);
            string ty = xml.getAttributeValue("Type");
            if (isNull(ty)) return;
            Enum.TryParse(ty, out type);
        }
    }
    class LineInv : Entity
    {
        public double L;
        public string ls, num, noAng;
        public int count;
        public int numSE;
        public LineInv(SketchPoint bp = null)
            : base(entTypes.Line)
        {
            this.bp = bp;
            get();
        }
        public LineInv(SketchLine line, bool f = false, bool close = false)
            : base(entTypes.Line)
        {
            L = line.Length*10;
            dir = line.Geometry.Direction.AsVector();
            if (line.Constraints.OfType<TangentSketchConstraint>().Count() != 0)
            {
                con = cType.t;
            }
            Point2d pt = line.StartSketchPoint.Geometry;
            if (f) { x = pt.X; y = pt.Y; first = true;}
        }
        static public double scalar(SketchLine sl,SketchLine prev)
        {
            return scalar(sl.Geometry.Direction, prev.Geometry.Direction); 
        }
        static public double scalar(UnitVector2d v1, UnitVector2d v2)
        {
            return v1.X * v2.X + v1.Y * v2.Y;
        }
        static public double scalarX(SketchLine sl)
        {
            return scalar(sl.Geometry.Direction, u.createUnitVector2d(0, 1));
        }
        static public double scalarY(SketchLine sl)
        {
            return scalar(sl.Geometry.Direction, u.createUnitVector2d(1, 0));
        }
        public void constr(Entity prev, bool hw)
        {
            sl = ent as SketchLine;
            if (prev.t == entTypes.Line)
            {
                if (u.eq(scalar(sl, prev.ent as SketchLine),0))
                {
                    constraints.addPerp(ent, prev.ent);
                }
                else if (isNull(noAng) && !last && !hw)
                {
                    TwoLineAngleDimConstraint ac = constraints.addTwoLineAngle(prev.ent as SketchLine, ent as SketchLine, name: name, w: 0.5);
                    constrts.Add(ac as DimensionConstraint);
                    //Parameter p = ac.Parameter;
                    //if (!isNull(angs)) u.round(p);
                }
            }
            else if (prev.t == entTypes.Arc)
            {
                if (con.HasFlag(cType.t)) constraints.addTangent(prev.ent, ent);
                else if (isNull(noAng))
                {
                    addAngle = true;
                    //SketchLine l = addTangLine(prev);
                    //u.round(constraints.addTwoLineAngle(l, ent as SketchLine).Parameter);
                }
            }
        }
        public SketchLine addTangLine(Entity prev)
        {
            Arc2d arc = (prev.ent as SketchArc).Geometry;
            Curve2dEvaluator ev = arc.Evaluator;
            double[] par = u.getParam(ev, prev.endPt.Geometry);
            double[] o = { };
            ev.GetTangent(ref par, ref o);
            Vector2d vec = u.createVector2d(o[0], o[1]);
            if (!(prev as ArcInv).clockwise)
            {
                vec.ScaleBy(-1);
            }
            SketchLine l = addLine(prev.endPt, 1, vec); l.Construction = true;
            constraints.addTangent(prev.ent, l as SketchEntity);
            return l;
        }
        public SketchLine addLine(SketchPoint skp, double L, Vector2d d)
        {
            d.Normalize();
            if (L != 0)
                d.ScaleBy(L);
            Point2d pt = skp.Geometry; pt.TranslateBy(d);
            return ps.SketchLines.AddByTwoPoints(skp, pt);
        }
        public string dimConstr(bool hw)
        {
            sl = ent as SketchLine;
            DimensionConstraint dc = null;
            if (!hw)
            {
                dc = constraints.addTwoPointDist(sl, name, 1, scale: 0.2) as DimensionConstraint;
                constrts.Add(dc);
                constraints.setVal(ls);
            }
            else
            {
                dc = constraints.addTwoPointDist(sl, name, 1, scale: 0.2, o:DimensionOrientationEnum.kHorizontalDim) as DimensionConstraint;
                constrts.Add(dc);
                dc = constraints.addTwoPointDist(sl, name, 1, scale: 0.2, o: DimensionOrientationEnum.kVerticalDim) as DimensionConstraint;
                constrts.Add(dc);
            }
            return constraints.Val.Parameter.Name;
        }
        public void direct(Entity prev)
        {
            if (prev.ent.Type != ObjectTypeEnum.kSketchArcObject) return;
            Arc2d arc = (prev.ent as SketchArc).Geometry;
            Curve2dEvaluator ev = arc.Evaluator;
            double[] par = u.getParam(ev, prev.endPt.Geometry);
            double[] o = { };
            ev.GetTangent(ref par, ref o);
            double pdirl = dir.Length;
            dir = u.createVector2d(o[0], o[1]);
            if (prev.t == entTypes.Arc && !(prev as ArcInv).clockwise)
            {
                dir.ScaleBy(-1);
            }
            if (u.eq(pdirl, 1)) dir.Normalize();
        }

        public override void draw()
        {
            if (prev == null) return;
            bool hw = false;
            if (ls == "1") hw = true;
            if (last)
            {
                drawLast();
            }
            else
            {
                sp = prev.endPt;
                if (con.HasFlag(cType.t)) direct(prev);
                double dl = dir.Length;
                
                if (!u.eq(dl, 1))
                {
                    L = dl; ls = dl.ToString("#.###");
                }
                dir.Normalize();
                if (L != 0)
                    dir.ScaleBy(L);
                Point2d pt = sp.Geometry; pt.TranslateBy(dir);
                sl = ps.SketchLines.AddByTwoPoints(sp, pt);
            }
            startPt = sl.StartSketchPoint; endPt = sl.EndSketchPoint;
            if (con.HasFlag(cType.hole))
            {
                startPt.HoleCenter = true;
                if (con.HasFlag(cType.cl)) sl.Centerline = true;
                if (!con.HasFlag(cType.f)) endPt.HoleCenter = true;
                //sl.Construction = true;
            }
            ent = (SketchEntity)sl;
            string n = "";
            if (L != 0 && !con.HasFlag(cType.eq))       
                n = dimConstr(hw);
            constr(prev, hw);
            if (count != 0 && con.HasFlag(cType.hole)) addPoints(n);
        }

        public void addPoints(string name)
        {
            Curve2dEvaluator ev = sl.Geometry.Evaluator;
            double min, max, l;
            ev.GetParamExtents(out min, out max);
            l = max - min;
            l = l / count;
            string n = "";
            //string n = u.addParameter(doc, "", (l*10).ToString(), "mm").Name;
            for (int i = 1; i < count; i++)
            {
                double [] pt = {}, ls = {l};
                ev.GetPointAtParam(ref ls, ref pt);
                double lv;
                ev.GetLengthAtParam(min, l, out lv);
                SketchPoint tsp = ps.SketchPoints.Add(I.CP2d(pt[0], pt[1]));
                constraints.addConsid(sl as SketchEntity, tsp as SketchEntity);
                TwoPointDistanceDimConstraint dist = constraints.addTwoPointDist(sl.StartSketchPoint, tsp, "", 0, -(l * i));
                if (i == 1)
                {
                    dist.Parameter.Expression = name + "/" + count;
                    n = dist.Parameter.Name;
                }
                else
                {
                    dist.Parameter.Expression = n + "*" + i;
                }
            }
        }

        public void drawLast()
        {
            sl = ps.SketchLines.AddByTwoPoints(Entity.fEnt.startPt, prev.endPt);
//             if (ll != null)
//             {
//                 if (LineInv.scalar(sl, ll.getSl) == 0)
//                 {
//                     constraints.removeTwoPointDist(ll.getSl);
//                     constraints.addPerp(sl as SketchEntity, ll.getSl as SketchEntity);
//                 }
//             }
        }

        public override void upd()
        {
            Parameter p = u.findInCol<Parameter>(pcd.Parameters, e => e.Name == name);
            if (p != null && !u.eq(p.Value, L)) p.Expression = ls; 
        }

        public override string ToString()
        {
            XElement el = new XElement("Line");
            XMLDoc.addXAttributes(el, new Dictionary<string, string>() { { "Name", "" }, { "Dir", dir.X.ToString(form) + " " + dir.Y.ToString(form) },
            { "L", L.ToString(form) }, { "Constr", conToString() } });
            if (first) el.Add(new XAttribute("XY", (x*10).ToString(form) + " " + (y*10).ToString(form)));
            xml.El.Add(el);
            return el.ToString();
        }

        public void norm()
        {
            if (Math.Abs(dir.X) > 2) dir.X = dir.X / 10;
            if (Math.Abs(dir.Y) > 2) dir.Y = dir.Y / 10;
        }

        public override void get()
        {
            base.get();
            if (update) return;
            noAng = xml.getAttributeValue("NoAng");
            addNoDim();
            if (last) return;
            angs = xml.getAttributeValue("Ang");
            getCoord("Dir");
            norm();
            num = xml.getAttributeValue("Num");
            if (!isNull(num))
            {
                int n; int.TryParse(num, out n);
                if (!isNull(n)) numSE = n;
            }       
            ls = xml.getAttributeValue("L");
            L = getDouble(ls, 0.1, 3)/10;
            count = getInt("Count");
            br = false;
            if (!hasFirst) { getBase(); hasFirst = true; }
            if (br) return;
            if (xml.getCoord(bx, by))
            {
                first = true;
                double tx = getDouble(xml.xname, 1, 3), ty = getDouble(xml.yname, 1, 3);
                sp = ps.SketchPoints.Add(u.createPoint2d(tx/10 + bx, ty/10 + by));
                sp.HoleCenter = false;
                if (isNull(bp)) bp = ps.SketchPoints.Add(I.CP2d(tx/10 + bx,ty/10 + by));
                else if (bp.ReferencedEntity != null)
                {
                    SketchInv.origin = bp;
                    bp = ps.SketchPoints.Add(I.CP2d(tx / 10 + bx, ty / 10 + by));
                }
                prev = new PointInv(sp);
            }
        }
    }
    class ArcInv : Entity
    {
        string rs;
        double R, ang;
        string angx, noAng;
        Point2d cen;
        public bool clockwise = true;
        public ArcInv()
            : base(entTypes.Arc)
        {
            get();
        }
        public ArcInv(SketchArc arc)
            : base(entTypes.Arc)
        {
            R = arc.Radius*10;
            ang = arc.SweepAngle / Math.PI * 180;
            checkArc(arc.Geometry);
            if (arc.Constraints.OfType<TangentSketchConstraint>().FirstOrDefault() != null)
            {
                con = cType.t;
            }
            else
            {
                x = arc.CenterSketchPoint.Geometry.X; y = arc.CenterSketchPoint.Geometry.Y;
                cen = u.createPoint2d(x, y);
            }
        }
        void checkArc(Arc2d a)
        {
            Vector2d d1, d2;
            d1 = u.getTangentVec(a.Evaluator);
            d2 = u.getTangentVec(a.Evaluator, 0.15);
            double a1 = d1.AngleTo(u.createVector2d(-1, -1)), a2 = d2.AngleTo(u.createVector2d(-1,-1));
            if (a2-a1 < 0) ang = -ang;
        }
        void setCenter(SketchLine sl)
        {
            Vector2d v = sl.Geometry.Direction.AsVector();
            double a = ang < 0 ? -Math.PI/2: Math.PI/2;
            u.rotate(v, sl.EndSketchPoint.Geometry, a);
            v.ScaleBy(R);
            Point2d pt = sl.EndSketchPoint.Geometry;
            pt.TranslateBy(v);
            x = pt.X; y = pt.Y;
            cen = u.createPoint2d(x, y);
        }
        void setCenter()
        {
            if (xml.getCoord()) u.createPoint2d(xml.x, xml.y);
        }
        public Point2d pointToArc(double x, double y, Point2d pt, double ang)
        {
            ang -= Math.PI / 2;
            Point2d r = u.createPoint2d(0, 0);  
            r.X = R * Math.Cos(ang) + x;
            r.Y = R * Math.Sin(ang) + y;
            return r;
        }
        public void constr(Entity prev)
        {
            SketchArc arc = ent as SketchArc;
            if (prev.ent.Type != ObjectTypeEnum.kSketchLineObject) return;
            if (con.HasFlag(cType.t))
                constraints.addTangent(prev.ent, arc as SketchEntity);
        }
        public void dimConstr()
        {
            SketchArc arc = ent as SketchArc;
            if (u.eq(ang, Math.PI))
            {
                constraints.addRadius(ent, arc.CenterSketchPoint.Geometry, arc.EndSketchPoint.Geometry, name);
                constrts.Add(constraints.Val);
                constraints.setVal(rs);
            }
            else
            {
                SketchLine sl = ps.SketchLines.AddByTwoPoints(arc.EndSketchPoint, arc.CenterSketchPoint); sl.Construction = true;
                SketchLine sl1 = ps.SketchLines.AddByTwoPoints(arc.StartSketchPoint, arc.CenterSketchPoint); sl1.Construction = true;
                Parameter p;
                DimensionConstraint r = constraints.addRadius(ent, sl.StartSketchPoint.Geometry, sl1.StartSketchPoint.Geometry, name) as DimensionConstraint;
                constrts.Add(r);
                constraints.setVal(rs);
                if (isNull(noAng))
                {
                    DimensionConstraint dc = constraints.addTwoLineAngle(sl, sl1, arc.CenterSketchPoint, name) as DimensionConstraint;
                    constrts.Add(dc);
                    constraints.setVal(angs);
                }
                r.TextPoint = u.midPt(sl.StartSketchPoint.Geometry, sl.EndSketchPoint.Geometry, 0, 1, 0.1, 0.3);
                if (!isNull(angx) && !con.HasFlag(cType.t))
                {
                    if (endPt.Equals(arc.StartSketchPoint))
                        constraints.addTwoLineAngle(prev.ent as SketchLine, sl, arc.CenterSketchPoint);
                    else
                        constraints.addTwoLineAngle(prev.ent as SketchLine, sl1, arc.CenterSketchPoint);
                    constraints.setVal(angx);
                }
            }
        }

        public override void draw()
        {
            sp = null; UnitVector2d v = null; double sAng = 0;
            if (prev.ent.Type == ObjectTypeEnum.kSketchPointObject)
                sp = prev.ent as SketchPoint;
            else if (prev.ent.Type == ObjectTypeEnum.kSketchLineObject)
            {
                sl = prev.ent as SketchLine;
                //setCenter();
                if (cen == null)
                    setCenter(sl);
                sp = sl.EndSketchPoint;
                v = sl.Geometry.Direction;
                sAng = v.AngleTo(u.createUnitVector2d(1, 0));
            }
            //SketchArc arc = ps.SketchArcs.AddByCenterStartSweepAngle(cen, R, sAng, -ang);
            //constraints.addConsid(arc.StartSketchPoint as SketchEntity, sl.EndSketchPoint as SketchEntity);
            if (ang < 0) clockwise = false;
            if (ang < 0) sAng = sAng - Math.PI;
            SketchArc arc = ps.SketchArcs.AddByThreePoints(sp, pointToArc(x, y, sp.Geometry, ang/2 + sAng), pointToArc(x, y, sp.Geometry, ang + sAng));
            //SketchArc arc = ps.SketchArcs.AddByCenterStartEndPoint(cen, sp, pointToArc(x, y, sp.Geometry, ang + sAng), clockwise);
            if (ang > 0)
            {
                startPt = arc.StartSketchPoint; endPt = arc.EndSketchPoint;
            }
            else
            {
                startPt = arc.EndSketchPoint; endPt = arc.StartSketchPoint;
            }
            ent = (SketchEntity)arc;
            constr(prev);
            dimConstr();
        }

        public override void upd()
        {
            Parameter p = u.findInCol<Parameter>(pcd.Parameters, e => e.Name == name+"R");
            if (p != null && !u.eq(p.Value, R)) p.Value = R;
            p = u.findInCol<Parameter>(pcd.Parameters, e => e.Name == name + "A");
            if (p != null && !u.eq(p.Value, ang)) p.Value = ang; 
        }
        public override string ToString()
        {
            XElement el = new XElement("Arc");
            XMLDoc.addXAttributes(el, new Dictionary<string, string>() { { "Name", "" }, { "Ang", ang.ToString(form) }, { "R", R.ToString(form) }, { "Constr", conToString() } });
            if (cen != null) el.Add(new XAttribute("XY", x.ToString(form) + " " + y.ToString(form)));
            xml.El.Add(el);
            return el.ToString();
        }

        public override void get()
        {
            base.get();
            addNoDim();
            if (update) return;
            rs = xml.getAttributeValue("R");
            angs = xml.getAttributeValue("Ang");
            angx = xml.getAttributeValue("AngX");
            noAng = xml.getAttributeValue("NoAng");
            R = getDouble(rs, 1, 3)/10;
            ang = getDouble(angs, 1, 3) * Math.PI / 180;
            if (xml.getCoord())
            {
                x = getDouble(xml.xname, 0.1, 3); y = getDouble(xml.yname, 0.1, 3);
                cen = u.createPoint2d(x, y);
            }
        }
    }
    class IMatesInv : Entity
    {
        List<IMateInv> ims;
        static public double d;
        string mName;
        string no;
        iMateDefinition imate;
        iMateDefinitions idefs;
        ObjectCollection col;
        public IMatesInv() : base(entTypes.IMates)
        {
            if (isNull(pcd)) pcd = I.getPCD(doc);
            idefs = isNull(pcd) ? acd.iMateDefinitions : pcd.iMateDefinitions;
            ims = new List<IMateInv>();
            alias = "КП";
            setAlias();
            get();
            setName();
            readXML();
        }
        public void readXML()
        {
            foreach (var el in xml.El.Elements())
            {
                xml.El = el;
                if (no != null && no != "")
                    ims.Add(new IMateInv(name));
                else ims.Add(new IMateInv());
                if (br) return;
            }
        }

        protected virtual void setAlias()
        {
            string tp = xml.getAttributeValue("Alias");
            if (tp != null) alias = tp;
        }

        public bool check()
        {
            if (!isNull(acd)) idefs = acd.iMateDefinitions;
            iMateDefinition tim = u.get<iMateDefinition>(idefs, f => f.Name == name);
            if (isNull(tim)) return true;
            CompositeiMateDefinition tcim = tim as CompositeiMateDefinition; 
            if (isNull(tcim)) return true;
            return false;
        }

        public override void draw()
        {
            col = I.COC(col);
            if (!check()) return;
            foreach (var item in ims)
            {
                item.add();
                item.addInCol(ref col);
            }
            d = default(double);
            if (!isNull(no)) return;
            imate = idefs.AddCompositeiMateDefinition(col, name) as iMateDefinition;
            addToMatch();
            pcd = null;
        }

        public void addToMatch()
        {
            if (mName == null) return;
            string[] m = imate.MatchList as string[];
            var spl = u.getSpl(mName, ';');
            IEnumerable<string> ie = m;
            u.action<string>(spl, a => ie = ie.Concat(u.add<string>(a)));
            imate.MatchList = ie.ToArray();
        }

        public override void upd()
        {
           
        }

        public override void get()
        {
            name = xml.getAttributeValue("Name");
            d = xml.getDoubleAttValue("d");
            mName = xml.getAttributeValue("Match");
            no = xml.getAttributeValue("no");
        }
    }
    class AssemblyConstrInv : FeaturesInv
    {
        List<IMateInv> ims = new List<IMateInv>();
        AssemblyConstraint ac = null;
        string d, type, srev, ang;
        object ent1, ent2;

        public AssemblyConstrInv()
        {
            t = entTypes.AsmConstr;
            if (isNull(acd)) return;
            get();
            readXML();
        }

        public override void draw()
        {
            if (ims.Count != 2) return;
            switch (type)
            {
                case "вставка":
                    addIns();
                    break;
                case "совмещение":
                    addMate();
                    break;
                case "угол":
                    if (!isNull(d))
                        addAngle();
                    break;
                case "концентрично":
                    //addConcentric();
                    break;
                case "ось":
                    //addAxisImate();
                    break;
                default:
                    break;
            }
        }

        public void addIns()
        {
            ent1 = ims[0].getEnt(); ent2 = ims[1].getEnt();
            bool rev = false;
            if (isNull(srev)) rev = true;
            if (check(ent1, ent2)) return;
            acd.Constraints.AddInsertConstraint(ent1, ent2, rev, u.convToDouble(d, 0.1, 3));
        }

        public void addMate()
        {
            ent1 = ims[0].getEnt(); ent2 = ims[1].getEnt();
            bool rev = false;
            InferredTypeEnum ite1 = InferredTypeEnum.kNoInference, ite2 = InferredTypeEnum.kNoInference;
            if (ims[0].type == "ось") ite1 = InferredTypeEnum.kInferredLine;
            if (ims[1].type == "ось") ite2 = InferredTypeEnum.kInferredLine;
            if (isNull(srev)) rev = true;
            if (check(ent1, ent2)) return;
            if (rev) acd.Constraints.AddMateConstraint(ent1, ent2, u.convToDouble(d, 0.1, 3), ite1, ite2);
            else acd.Constraints.AddFlushConstraint(ent1, ent2, u.convToDouble(d, 0.1, 3), ite1, ite2);
        }

        new public void addAngle()
        {
            ent1 = ims[0].getEnt(); ent2 = ims[1].getEnt();
            if (check(ent1, ent2)) return;
            acd.Constraints.AddAngleConstraint(ent1, ent2, u.degToRad(u.convToDouble(d, 1, 3)), AngleConstraintSolutionTypeEnum.kDirectedSolution);
        }

        public bool check(object ent1, object ent2)
        {
            foreach (AssemblyConstraint item in acd.Constraints)
            {
                if (item.EntityOne.Equals(ent1))
                    if (item.EntityTwo.Equals(ent2)) return true;
            }
            return false;
        }

        public void readXML()
        {
            foreach (var el in xml.El.Elements())
            {
                xml.El = el;
                ims.Add(new IMateInv());
                if (br) return;
            }
        }

        public override void upd()
        {
           
        }

        public override void get()
        {
            type = xml.getAttributeValue("type");
            d = xml.getAttributeValue("d");
            srev = xml.getAttributeValue("rev");
        }

        protected override void addDef()
        {

        }
    }
    class IMateInv : FeaturesInv
    {
        new iMateDefinition def;
        iMateDefinitions idefs;
        public List<iMateDefinition> defs;
        public string type;
        string n;
        string p, pl, ang, asm, fltr, adpt, reg, imate;
        bool align = true;
        UnitVector norm;
        new ComponentOccurrence oc;
        object ob;
        bool rev = true;
        string d;

        public IMateInv(string n = null)
        {
            idefs = isNull(pcd) ? acd.iMateDefinitions : pcd.iMateDefinitions;
            defs = new List<iMateDefinition>();
            num = 1; FeaturesInv.oc = null;
            get();
            this.n = n; 
        }

        public override void draw()
        {
            if (!isNull(d)) IMatesInv.d = u.convToDouble(d);
            switch (type)
            {
                case "вставка":
                    addInsIMate();
                    break;
                case "совмещение":
                    addMateImate();
                    break;
                case "угол":
                    if (!isNull(ang))
                        addAngleMate();
                    break;
                case "концентрично":
                    addConcentric();
                    break;
                case "ось":
                    addAxisImate();
                    break;
                default:
                    break;
            }
        }

        public override void upd()
        {
          
        }

        protected override void addDef()
        {
            
        }

        public object getProxy(object el)
        {
            object tmp;
            if (doc.DocumentType == DocumentTypeEnum.kAssemblyDocumentObject)
            {
                var acd = I.getACD(doc); idefs = acd.iMateDefinitions;
            }
            if (isNull(oc))
            {
                return el;
            }
            oc.CreateGeometryProxy(el, out tmp);
            return tmp;
        }

        public void addInCol(ref ObjectCollection col)
        {
            foreach (var item in defs)
            {
                col.Add(item); 
            }
        }

        public void addInsIMate()
        {
            foreach (Edge item in edges)
            {
                if (item.GeometryType == CurveTypeEnum.kLineSegmentCurve) continue;
                ob = getProxy(item);
                def = idefs.AddInsertiMateDefinition(ob, align, IMatesInv.d / 10) as iMateDefinition;
                if (n != null) def.Name = n;
                defs.Add(def);  
            }
        }

        public void addAngleMate()
        {
            if (isNull(wpl))
                foreach (var item in fs)
                {
                    ob = getProxy(item);
                    def = idefs.AddAngleiMateDefinition(ob, rev, ang) as iMateDefinition;
                    if (n != null) def.Name = n;
                    defs.Add(def);
                }
            else
            {
                ob = getProxy(wpl);
                def = idefs.AddAngleiMateDefinition(ob, rev, ang) as iMateDefinition;
                if (n != null) def.Name = n;
                defs.Add(def);
            }
        }

        public void addConcentric()
        {
            foreach (var item in fs)
            {
                ob = getProxy(item);
                def = idefs.AddMateiMateDefinition(ob, IMatesInv.d / 10, InferredTypeEnum.kInferredLine) as iMateDefinition;
                if (n != null) def.Name = n;
                defs.Add(def);
            }
        }

        public void addAxisImate()
        {
            foreach (var item in fs)
            {
                ob = getProxy(item);
                def = idefs.AddMateiMateDefinition(ob, IMatesInv.d / 10, InferredTypeEnum.kInferredLine) as iMateDefinition;
                defs.Add(def);
            }
        }

        public void addMateImate()
        {
            if (isNull(wpl))
            foreach (var item in fs)
            {
                ob = getProxy(item);
                if (rev) def = idefs.AddFlushiMateDefinition(getProxy(item), IMatesInv.d / 10) as iMateDefinition;
                else def = idefs.AddMateiMateDefinition(ob, IMatesInv.d / 10) as iMateDefinition;
                if (n != null) def.Name = n;
                defs.Add(def);
            }
            else
            {
                ob = getProxy(wpl);
                if (rev) def = idefs.AddFlushiMateDefinition(ob, IMatesInv.d / 10) as iMateDefinition;
                else def = idefs.AddMateiMateDefinition(ob, IMatesInv.d / 10) as iMateDefinition;
                if (n != null) def.Name = n;
                defs.Add(def);
            }
        }

        public object getFace()
        {
            return getProxy(fs.ElementAt(0));
        }

        public object getEdge()
        {
            return getProxy(edges[1]);
        }

        public object getWP()
        {
            return getProxy(wpl);
        }

        public object getFromiMate()
        {
            int n = u.getIntElem(asm,1); string name = u.getElem(asm, 0);
            int num = u.getIntElem(imate,0);
            if (isNull(n, name, num)) return null;
            iMateDefinition d = findImatesDef(acd.Occurrences, name, reg, n, num);
            oc = FeaturesInv.oc;
            if (d is InsertiMateDefinition) return getProxy(((InsertiMateDefinition)d).Entity);
            else if (d is CompositeiMateDefinition) 
            {
                var id = d as CompositeiMateDefinition;
                var nu = u.getIntElem(imate,1);
                if (isNull(nu)) return null;
                if (id[nu] is InsertiMateDefinition) return getProxy(((InsertiMateDefinition)id[nu]).Entity);
                else if (id[nu] is MateiMateDefinition) return getProxy(((MateiMateDefinition)id[nu]).Entity);
            }
            return null;
        }

        public object getEnt()
        {
            if (!isNull(wpl)) return getWP();
            if (!isNull(edges)) return getEdge();
            if (!isNull(fs)) return getFace();
            if (!isNull(imate)) return getFromiMate();
            return null;
        }

        public override bool addEdge()
        {
            base.addEdge();
            if (type == "вставка")
            {
                getEdges();
                base.getNumEdges();
            }
            else if (type == "ось")
            {
                getCylSurf();
            }
            else if ((type == "совмещение" || type == "угол")&& wpl == null)
            {
                getFaces();
            }
            else if ((type == "концентрично"))
            {
                getCylSurf();
            }
            return true;
        }

        public void getCylSurf()
        {
            base.getFaces(f => f.SurfaceType == SurfaceTypeEnum.kCylinderSurface);
            if (!isNull(fltr))
            {
                fs.RemoveWhere(e => !filterR(e, u.getDoubleElem(fltr, 0, scale: 0.1), u.getDoubleElem(fltr, 1, scale: 0.1)));
            }
            fs = filter(fs);
        }

        public void getFaces()
        {
            base.getFaces(f => f.SurfaceType == SurfaceTypeEnum.kPlaneSurface);
            if (!isNull(fltr))
            {
                string typ = null;
                norm = getNorm(ref typ);
                if (!isNull(norm)) {
                    fs.RemoveWhere(e => filterFace(e, typ));
                }
            }
            fs = filter(fs);
        }

        public bool filterFace(Face f, string typ)
        {
            Plane pl = f.Geometry as Plane;
            if (typ == "ff")
                return !pl.Normal.IsParallelTo(norm);
            else if (typ == "cf")
                return !pl.Normal.IsPerpendicularTo(norm);
            else return false;
        }

        public override void getEdges()
        {
            int i = 1;
            set(1, ref i);
            getFaces(fi => fi.SurfaceType == SurfaceTypeEnum.kCylinderSurface);
            getEdges((fi, num) => filter(fi, num, ref i));
        }

        public void set(int j, ref int i)
        {
            if (j == 1)
            {
                i = 1;
                if (rev) i = 2;
            }
        }

        public bool filter(Edge ed, int j ,ref int i)
        {
            set(j, ref i);
            if (ed.GeometryType == CurveTypeEnum.kLineSegmentCurve)
            {
                i++; return false;
            }
            else if (j == i) return true;
            return false;
        }

        public override void get()
        {
            type = xml.getAttributeValue("type");
            imate = xml.getAttributeValue("imate");
            p = xml.getAttributeValue("Parent");
            fltr = xml.getAttributeValue("Filter");
            adpt = xml.getAttributeValue("Adaptive");
            ang = xml.getAttributeValue("ang");
            asm = xml.getAttributeValue("Asm");
            d = xml.getAttributeValue("d");
            reg = xml.getAttributeValue("reg");
            if (!isNull(imate)) return;
            string salign = xml.getAttributeValue("align");
            if (!isNull(salign)) align = false;
            if (!isNull(asm, acd))
            {
                string tmp = getElem(asm, 1);
                if (!isNull(tmp)) num = int.Parse(tmp);
                asm = getElem(asm, 0);
                findPCD(acd.Occurrences, asm, reg);
                if (isNull(asm)) FeaturesInv.oc = null;
                oc = FeaturesInv.oc;
                if (!isNull(oc) && oc.Grounded) oc.Grounded = false;
            }
            
            if (!isNull(adpt))
            {
                PartDocument pdoc = pcd.Document as PartDocument;
                if (!isNull(pdoc)) pdoc.ModelingSettings.AdaptivelyUsedInAssembly = true;
            }
            if (!find(p)) findFaces(p);
            pl = xml.getAttributeValue("Plane");
            if (!isNull(acd) && isNull(asm))
                pcd = null;
            if (!isNull(pl)) wpl = findWP(pl, pcd);
            if (isNull(xml.getAttributeValue("rev"))) rev = false;
            addEdge();
        }
        
    }

    static class constraints
    {
        public static PlanarSketch ps;
        static DimensionConstraint val;
        static public Inventor.DimensionConstraint Val
        {
            get { return val; }
            set { val = value; }
        }
        static Point2d mp;
        public static Parameters pars;
        static public void addName(string name, string suff = "")
        {
            if (!u.isNull(name))
            {
                name += suff;
//                 Parameter p = InvDoc.Reflect.getProp<T,Parameter>(val as T, "Parameter");
//                 if (p.Name == name) name += "1";
//                 p.Name = name;
                Parameter p = u.get<Parameter>(pars, fi => fi.Name == name);
                DimensionConstraint cur = val as DimensionConstraint;
                if (p != null && p.ParameterType == ParameterTypeEnum.kUserParameter) cur.Parameter.Expression = p.Name;
                else cur.Parameter.Name = name;
            }
        }
        static public void create(string name, string val, units u = units.mm)
        {
            pars.UserParameters.AddByExpression(name, val, u.ToString());
        }
        static public RadiusDimConstraint addRadius(SketchEntity ent, Point2d sp1, Point2d sp2, string name, double on = 0, double od = 0, double s = 1, double w = 0.5)
        {
            mp = u.midPt(sp1, sp2, od, on, s, w );
            val = ps.DimensionConstraints.AddRadius(ent, mp) as DimensionConstraint;
            addName(name,"R");
            return val as RadiusDimConstraint;
        }
        static public TwoLineAngleDimConstraint addTwoLineAngle(SketchLine sl1, SketchLine sl2, SketchPoint sp = null, string name = "", double w = 0.5)
        {
            if (sp != null)
                mp = u.midPt(sp.Geometry, mp, w);
            else
            {
                mp = u.midPt(sl1.StartSketchPoint.Geometry, sl2.EndSketchPoint.Geometry, w);
            }
            val = ps.DimensionConstraints.AddTwoLineAngle(sl1, sl2, mp) as DimensionConstraint;
            TwoLineAngleDimConstraint cons = val as TwoLineAngleDimConstraint;
            Parameter p = cons.Parameter;
            u.round(p);
            addName(name,"A");
            return cons;
        }
        static public TwoPointDistanceDimConstraint addTwoPointDist(SketchLine sl, string name, double offset = 0, double scale = 1, DimensionOrientationEnum o = DimensionOrientationEnum.kAlignedDim)
        {
            mp = u.midPt(sl.StartSketchPoint.Geometry, sl.EndSketchPoint.Geometry, offset, 0, scale);
            val = ps.DimensionConstraints.AddTwoPointDistance(sl.StartSketchPoint, sl.EndSketchPoint, o, mp) as DimensionConstraint;
            addName(name);
            return val as TwoPointDistanceDimConstraint;
        }
        static public OffsetDimConstraint addTwoPointDist(SketchLine ent1, SketchPoint ent2, double offset = 0, double scale = 1)
        {
            mp = u.midPt(ent1.RangeBox.MinPoint, ent2.Geometry, offset, 0, scale);
            val = ps.DimensionConstraints.AddOffset(ent1, ent2 as SketchEntity, mp, false) as DimensionConstraint;
            return val as OffsetDimConstraint;
        }
        static public TwoPointDistanceDimConstraint addTwoPointDist(SketchPoint sp1, SketchPoint sp2, string name, double offset = 0, double scale = 1, DimensionOrientationEnum orient = DimensionOrientationEnum.kAlignedDim)
        {
            mp = u.midPt(sp1.Geometry, sp2.Geometry, offset, 0, scale);
            try
            {
                val = ps.DimensionConstraints.AddTwoPointDistance(sp1, sp2, orient, mp) as DimensionConstraint;   
                addName(name);
                return val as TwoPointDistanceDimConstraint;
            }
            catch (System.Exception)
            {
                return null;
            }
            
        }
        
        static public MidpointConstraint addMidPoint(SketchPoint sp, SketchLine sl)
        {
            return ps.GeometricConstraints.AddMidpoint(sp, sl);
        }
        static public VerticalAlignConstraint addVert(SketchPoint sp1, SketchPoint sp2)
        {
            return ps.GeometricConstraints.AddVerticalAlign(sp1, sp2);
        }
        static public HorizontalAlignConstraint addHor(SketchPoint sp1, SketchPoint sp2)
        {
            return ps.GeometricConstraints.AddHorizontalAlign(sp1, sp2);
        }
        static public HorizontalConstraint addHor(object ob)
        {
            SketchEntity se = ob as SketchEntity;
            if (se.ConstraintStatus != ConstraintStatusEnum.kFullyConstrainedConstraintStatus)
            return ps.GeometricConstraints.AddHorizontal(se);
            return null;
        }
        static public VerticalConstraint addVert(object ob)
        {
            SketchEntity se = ob as SketchEntity;
            if (se.ConstraintStatus != ConstraintStatusEnum.kFullyConstrainedConstraintStatus)
            return ps.GeometricConstraints.AddVertical(se);
            return null;
        }
        static public TangentSketchConstraint addTangent(SketchEntity se1, SketchEntity se2)
        {
            return ps.GeometricConstraints.AddTangent(se1, se2);
        }
        static public PerpendicularConstraint addPerp(SketchEntity se1, SketchEntity se2)
        {
            if (se1 == null || se2 == null) return null;
            return ps.GeometricConstraints.AddPerpendicular(se1, se2);
        }
        static public ParallelConstraint addPar(SketchEntity se1, SketchEntity se2)
        {
            if (se1 == null || se2 == null) return null;
            return ps.GeometricConstraints.AddParallel(se1, se2);
        }
        static public CoincidentConstraint addConsid(SketchEntity se1, SketchEntity se2)
        {
            return ps.GeometricConstraints.AddCoincident(se1, se2);
        }
        static public CollinearConstraint addCollinear(object ob1, object ob2)
        {
            SketchEntity se1 = ob1 as SketchEntity, se2 = ob2 as SketchEntity;
            if (u.isNull(se1, se2)) return null;
            return ps.GeometricConstraints.AddCollinear(se1, se2);
        }
        static public GroundConstraint addFixed(object ob)
        {
            SketchEntity se = ob as SketchEntity;
            return ps.GeometricConstraints.AddGround(se);
        }
        static public void removeFixed(object ob)
        {
            SketchEntity se = ob as SketchEntity;
            u.action<GeometricConstraint>(se.Constraints, a => a.Delete(), f => f.Type == ObjectTypeEnum.kGroundConstraintObject);
            SketchLine sl = ob as SketchLine;
            if (!u.isNull(sl))
            {
                u.action<GeometricConstraint>(sl.StartSketchPoint.Constraints, a => a.Delete(), f => f.Type == ObjectTypeEnum.kGroundConstraintObject);
                u.action<GeometricConstraint>(sl.EndSketchPoint.Constraints, a => a.Delete(), f => f.Type == ObjectTypeEnum.kGroundConstraintObject);
            }
        }
        static public void removeTwoPointDist(SketchLine sl)
        {
            foreach (TwoPointDistanceDimConstraint item in ps.DimensionConstraints)
            {
                if ((item.PointOne.Equals(sl.StartSketchPoint) && item.PointTwo.Equals(sl.EndSketchPoint)) ||
                    (item.PointTwo.Equals(sl.StartSketchPoint) && item.PointOne.Equals(sl.EndSketchPoint)))
                    item.Delete();
            }
        }
        static public EqualLengthConstraint addEq(SketchLine sl1, SketchLine sl2)
        {
            try
            {
                return ps.GeometricConstraints.AddEqualLength(sl1, sl2);
            }
            catch (System.Exception ex)
            {
                return null;
            }
        }
        static public void setVal(string v)
        {
            double d; double.TryParse(v, out d);
            if (d != 0)
            {
                val.Parameter.Expression = d < 0 ? (-d).ToString() : d.ToString();
            }
            else
            {
                u.addParameter(Entity.doc, v, Entity.pars);
                val.Parameter.Expression = v.StartsWith("-") ? v.TrimStart('-') : v;
            }
        }
        static public Dictionary<string, bool> check(SketchPoint sp)
        {
            Dictionary<string, bool> dic = new Dictionary<string, bool>() { { "v", false }, { "h", false } };
            foreach (var item in sp.Constraints)
            {
                var gc = item as GeometricConstraint;
                if (gc != null)
                {
                    if (gc is VerticalAlignConstraint) dic["v"] = true;
                    else if (gc is HorizontalAlignConstraint) dic["h"] = true;
                    continue;
                }
                var dc = item as TwoPointDistanceDimConstraint;
                if (dc != null) 
                { 
                    if (dc.Orientation == DimensionOrientationEnum.kHorizontalDim) dic["v"] = true;
                    else if (dc.Orientation == DimensionOrientationEnum.kVerticalDim) dic["h"] = true;
                    continue;
                }
            }
            return dic;
        }
    }
    class SketchInv : Entity
    {
        string angs, cNum, me, rdir;
        WorkPlane pl;
        public static SketchPoint origin;
        string namePlane;
        List<Entity> ents = new List<Entity>();
        List<ParamInv> prms = new List<ParamInv>();
        IEnumerable<Entity> fents;
        string[] dims;
        string[] constr;
        posType[] pos;
        List<SketchEntity> eds;

        public SketchInv(): base(entTypes.Sketch)
        {
            alias = "Э";
            if (!checkAsm()) return;
            origin = null;
            string fltr = xml.getAttributeValue("Fltr");
            if (isNull(fltr) && pcd.ReferenceComponents.DerivedPartComponents.Count != 0) return;
            setAlias();
            ps = getSketch();
            if (!isNull(ps)) update = true;
            addDims();
            addConstr();
            angs = xml.getAttributeValue("Rot");
            rdir = xml.getAttributeValue("rAlign");
            cNum = xml.getAttributeValue("CenterNum");
            me = xml.getAttributeValue("MustExist");
            if ((file.name(doc.FullFileName)).IndexOf("base") != -1) return;
            if (!checkPF(me)) return; 
            if (isNull(ps) && !checkBSketch() && !checkFeat())
            if (isNull(ps) && checkPlane()) return;
           
            if (!update)
            {
                setXDir();
                string bface = xml.getAttributeValue("bFace");
                pos = Pos();
                if (!isNull(bface))
                {
                    eds = getFromFace(bface);
                }
            }
            constraints.ps = ps;
            //getBaseSP();
            SketchPoint bsp = u.get<SketchPoint>(ps.SketchPoints, fi => fi.HoleCenter && fi.ConstraintStatus == ConstraintStatusEnum.kFullyConstrainedConstraintStatus);
            //if (u.gets<SketchPoint>(ps.SketchPoints, fi => fi.HoleCenter == true).Count() <= 1 &&
            //    u.gets<SketchLine>(ps.SketchLines, fi => fi.ReferencedEntity == null).Count() == 0) update = false;
            if (check(bsp)) update = true;
            else update = false;
            setName();
            readXML();
        }
        public SketchInv(PlanarSketch sketch): base(entTypes.Sketch)
        {
            Profile p = sketch.Profiles.AddForSurface();
            Entity.ps = sketch;
            if (p == null) return; bool f = true;
            ParamInv.xml = Entity.xml;
            foreach (DimensionConstraint item in sketch.DimensionConstraints)
            {
                prms.Add(new ParamInv(item.Parameter)); 
            }
            foreach (ProfilePath pp in p)
            {
                closed = pp.Closed;
                foreach (ProfileEntity pe in pp)
                {
                    switch (pe.SketchEntity.Type)
                    {
                        case ObjectTypeEnum.kSketchLineObject:
                            ents.Add(new LineInv(pe.SketchEntity as SketchLine, f));
                            if (f) f = false;
                            break;
                        case ObjectTypeEnum.kSketchArcObject:
                            ents.Add(new ArcInv(pe.SketchEntity as SketchArc));
                            break;
                        default:
                            break;
                    } 
                }
            }
        }

        public void setXDir()
        {
            string xd = xml.getAttributeValue("xDir");
            if (isNull(xd)) return;
            string elType = getElem(xd, 0), elName = getElem(xd, 1), elNum = getElem(xd, 2), elRev = getElem(xd, 3), elXRev= getElem(xd, 4);
            if (isNull(elType)) return;
            int i;
            int.TryParse(elNum, out i);
            xDirType ty; Enum.TryParse(elType, out ty);
            if (isNull(ty)) return;
            object wa = null;
            switch (ty)
            {
                case xDirType.s:
                    name = elName;
                    PlanarSketch ts = getSketch();
                    if (isNull(ts,i)) return;
                    wa = u.get(ts.SketchLines, i);
                    break;
                case xDirType.a:
                    if (isNull(elName)) return;
                    wa = findWA(elName);
                    break;
                case xDirType.e:
                    Face fa = ps.PlanarEntity as Face;
                    if (isNull(fa,i)) return;
                    var ie = u.gets<Edge>(fa.Edges, fi => fi.GeometryType == CurveTypeEnum.kLineSegmentCurve);
                    if (!isNull(elName)) ie = u.gets<Edge>(ie, fi => 
                    {
                        return fi.Faces[1].SurfaceType == SurfaceTypeEnum.kCylinderSurface ||
                        fi.Faces[2].SurfaceType == SurfaceTypeEnum.kCylinderSurface;
                    });
                    wa = u.get(ie, i);
                    break;
                case xDirType.f:

                    break;
                default:
                    break;
            }
            if (!isNull(elRev))
                ps.NaturalAxisDirection = true;
            if (!isNull(elXRev))
                ps.AxisIsX = true;
            ps.AxisEntity = wa;
        }

        public void addDims()
        {
            string tdims = xml.getAttributeValue("Dim");
            if (isNull(tdims)) return;
            dims = XMLDoc.getSpl(tdims, '#');
        }

        public void addConstr()
        {
            string tdims = xml.getAttributeValue("Constr");
            if (isNull(tdims)) return;
            constr = XMLDoc.getSpl(tdims, '#');
        }

        public bool checkPlane()
        {
            if (ps == null)
            {
                namePlane = xml.getAttributeValue("Plane");
                if (namePlane == null)
                {
                    return true;
                }
                pl = u.findPlane(doc, namePlane);
                if (pl != null) ps = pcd.Sketches.Add(pl);
                else return true;
            }
            name = xml.getAttributeValue("Name");
            ps.Name = name;
            return false;
        }

        public bool checkFeat()
        {
            

//             string n = xml.getAttributeValue("num"), l = xml.getAttributeValue("Last"); //new
//             if (isNull(l)) return true;
//             IEnumerable<Face> fa = null;
//             PartFeature pf = null;
// 
//             if (!isNull(getElem(l, 2)))
//             {
//                 fa = refEntity(u.gets<Face>(pcd.SurfaceBodies[1].Faces, fi => fi.CreatedByFeature.Type == ObjectTypeEnum.kReferenceFeatureObject), l);
//                 fa.Count();
//                 if (fa.Count() == 0) return false;
//             }
//             else
//             {
//                 pf = lastEntity<PartFeature>(pcd.Features, l).LastOrDefault();
//                 if (isNull(pf)) return false;
//             }
//             if (isNull(fa)) fa = getPlaneFaces(pf.Faces);
//             if (isNull(fa)) return false;
//             int nu = u.getNum(n, fa.Count());
//             string s = u.getElem(n, 1);
//             if (!isNull(s)) fa = u.sort<Face>(fa, s);
//             if (isNull(nu)) return false;
//             if (fa.Count() < nu) return false;
            Face fp = numFace(xml.El);
            if (isNull(fp)) return false;
            if (fp.SurfaceType != SurfaceTypeEnum.kPlaneSurface) return false;
            setName();
            ps = pcd.Sketches.Add(fp);
            if (!isNull(name))
            ps.Name = name;
            return true;
        }

        public bool checkBSketch()
        {
            if (isNull(xml.getAttributeValue("BProject"))) return false;
            if (xml.El.HasElements)
            {
                XElement tel = xml.El.Elements().ElementAt(0);
                string bs = XMLDoc.getAttributeValue(tel, "bSketch");
                if (isNull(bs)) return false;
                name = getElem(bs, 0, ';');
                PlanarSketch tsk = getSketch();
                if (isNull(tsk)) return false;
                ps = pcd.Sketches.Add(tsk.PlanarEntity);
            }
            return true;
        }

        public bool check(SketchPoint bsp)
        {
            if (!isNull(u.get<SketchLine>(ps.SketchLines, fi => !fi.Construction))) return true;
            if (isNull(bsp)) return false;
            foreach (var item in bsp.Constraints)
            {
                CoincidentConstraint cc = item as CoincidentConstraint;
                if (isNull(cc)) continue;
                SketchEntity se = cc.EntityOne.Equals(bsp) ? cc.EntityTwo : cc.EntityOne;
                SketchLine sl = se as SketchLine;
                if (isNull(sl)) continue;
                if (!sl.Construction) return true;
            }
            return false;
        }

        public void setAlias()
        {
            string tp = xml.getAttributeValue("Alias");
            if (tp != null) alias = tp;
        }

        public bool getBaseSP()
        {
            foreach (SketchPoint item in ps.SketchPoints)
            {
                if (item.HoleCenter == true) { bp = item; return true; }
            }
            return false;
        }
        public void setOrig()
        {
            if (!isNull(pos))
            {
                if(isNull(eds) || eds.Count == 0) return;
                SketchLine se = eds[0] as SketchLine;
                if (isNull(se)) return;
                if (pos.Length >= 1)
                {
                   origin = getFromPos(se, pos[0]); 
                }
                if (pos.Length == 2)
                {
                    se = eds[1] as SketchLine;
                    if (isNull(se)) return;
                    SketchPoint sp2 = getFromPos(se, pos[1]);
                    se = ps.SketchLines.AddByTwoPoints(origin, sp2);
                    se.Construction = true;
                    origin = ps.SketchPoints.Add(I.CP2d(), false);
                    constraints.addMidPoint(origin, se);
                }
                if (!isNull(origin))
                {
                    Vector2d v = fEnt.bp.Geometry.VectorTo(origin.Geometry);
                    v.X += fEnt.bp.Geometry.X; v.Y += fEnt.bp.Geometry.Y;
                    fEnt.bp.MoveBy(v);
                }
            }
        }
        public SketchPoint getFromPos(SketchLine se, posType pos)
        {
            SketchPoint pt = null;
            switch (pos)
            {
                case posType.s:
                    pt = se.StartSketchPoint;
                    break;
                case posType.m:
                    pt = ps.SketchPoints.Add(I.CP2d(), false);
                    constraints.addMidPoint(pt, se);
                    break;
                case posType.e:
                    pt = se.EndSketchPoint;
                    break;
                default:
                    break;
            }
            return pt;
        }
        public void setOrigin()
        {
            if (isNull(origin))
            {
                origin = ps.SketchPoints.Add(ps.ModelToSketchSpace(ps.OriginPointGeometry), false);
                constraints.addFixed(origin as SketchEntity);
            }
        }
        public void addConstraint()
        {
            if (fEnt.bp.ConstraintStatus != ConstraintStatusEnum.kFullyConstrainedConstraintStatus && fEnt.bp.ReferencedEntity == null)
            {
                setOrigin();
                if (!isNull(eds) && eds.Count != 0)
                {
                    if (eds.Count == 1)
                    {
                        SketchEntity se = eds[0];
                        bool x = isX(se.RangeBox);
                        if (!x)
                        {
                            addDim(se, xml.yname);
                            addDim(fEnt.bp as SketchEntity, "x", origin);
                        }
                        else
                        {
                            addDim(se, xml.xname);
                            addDim(fEnt.bp as SketchEntity, "y", origin);
                        }
                        return;
                    } 
                }
                if (fEnt.xent != null)
                {
                    addDim(fEnt.xent, xml.xname);
                }
                if (fEnt.yent != null) addDim(fEnt.yent, xml.yname);
                if (fEnt.xent == null && fEnt.yent == null)
                {
                    
                    if (fEnt.bp.Constraints.Count > 1)
                    {
                        GroundConstraint gr = fEnt.bp.Constraints[1] as GroundConstraint;
                        if (!isNull(gr))
                        {
                            if (gr.Deletable) gr.Delete();
                        }
                    }
                    addDim(fEnt.bp as SketchEntity, "x", origin);
                    addDim(fEnt.bp as SketchEntity, "y", origin);
                }
            }
        }
        public SketchPoint findSP()
        {
            if (ps.OriginPoint != null)
            {
                ps.OriginPoint = pcd.WorkPoints[1];
                return ps.OriginPoint as SketchPoint;
            }
                
            return u.get<SketchPoint>(ps.SketchPoints, f => f.ReferencedEntity != null && f.ReferencedEntity is WorkPoint);
        }
        public void addDim(SketchEntity ent, string name, SketchPoint origin = null)
        {
            if (ent.Type == ObjectTypeEnum.kSketchLineObject)
            {
                if (name == "0") constraints.addConsid(ent, fEnt.bp as SketchEntity);
                else
                {
                    constraints.addTwoPointDist(ent as SketchLine, fEnt.bp);
                    constraints.setVal(name);
                }
            }
            else if (ent.Type == ObjectTypeEnum.kSketchPointObject)
            {
                TwoPointDistanceDimConstraint d;
                bool sign;
                if (u.eq(origin.Geometry, (ent as SketchPoint).Geometry))
                {
                    if (ent.Constraints.Count == 0) constraints.addConsid(origin as SketchEntity, ent);
                    return;
                }
                if (name == "x")
                {
                    d = constraints.addTwoPointDist(origin, ent as SketchPoint, "", 1, 0.2, orient: DimensionOrientationEnum.kHorizontalDim);
                    if (xml.xname != "0" && isAlfa(xml.xname, out sign)) constraints.setVal(xml.xname);
                }
                else if (name == "y")
                {
                    d = constraints.addTwoPointDist(origin, ent as  SketchPoint, "", 1, 0.2, orient: DimensionOrientationEnum.kVerticalDim);
                    if (xml.yname != "0" && isAlfa(xml.yname, out sign)) constraints.setVal(xml.yname);
                }
            }
        }

        public void readXML()
        {
            foreach (var el in xml.El.Elements())
            {
                xml.El = el;
                entTypes ty = (entTypes)Enum.Parse(typeof(entTypes), el.Name.ToString());
                if (br) return;
                switch (ty)
                {
                    case entTypes.Point:
                        ents.Add(new PointInv());
                        break;
                    case entTypes.Line:
                        ents.Add(new LineInv(bp));
                        break;
                    case entTypes.Arc:
                        ents.Add(new ArcInv());
                        break;
                    case entTypes.Proj:
                        ents.Add(new ProjInv());
                        break;
                    case entTypes.Slot:
                        ents.Add(new SlotInv());
                        break;
                    case entTypes.Circle:
                        ents.Add(new CircleInv());
                        break;
                    default:
                        break;
                }
            }
        }
        public bool addCenter()
        {
            fents = u.gets<Entity>(ents, fi => fi.con.HasFlag(cType.cen));
            if (fents == null) return false;
            if (fents.Count() == 1)
            {
                addCenter(fents.ElementAt(0)); return true;
            }
            else if (fents.Count() == 2)
            {
                addCenter(fents);
                return true;
            }

            return false;
        }
        public void addCenter(IEnumerable<Entity> el)
        {
            Entity e1 = el.ElementAt(0), e2 = el.ElementAt(1);
            if (e1.t == entTypes.Line && e2.t == entTypes.Line)
            {
                SketchLine sl = ps.SketchLines.AddByTwoPoints(e1.startPt, e2.endPt);
                sl.Construction = true;
                addCenter(sl);
            }
            else if (e1.t == entTypes.Arc && e2.t == entTypes.Arc)
            {
                SketchLine sl = ps.SketchLines.AddByTwoPoints(((SketchArc)e1.ent).CenterSketchPoint, ((SketchArc)e2.ent).CenterSketchPoint);
                sl.Construction = true;
                addCenter(sl);
            }
        }
        public void addCenter(Entity el)
        {
            addCenter(el.ent);
        }
        public void addCenter(SketchLine sl)
        {
            if (fEnt.bp == null) fEnt.bp = ps.SketchPoints.Add(I.CP2d(), false);
            fEnt.bp.HoleCenter = false;
            constraints.addMidPoint(fEnt.bp, sl);
        }
        public void addCenter(SketchEntity ent)
        {
            if (fEnt.bp == null) fEnt.bp = ps.SketchPoints.Add(I.CP2d(), false);
            fEnt.bp.HoleCenter = false;
            if (ent.Type == ObjectTypeEnum.kSketchLineObject)
            {
                constraints.addMidPoint(fEnt.bp, (SketchLine)ent);
            }
            else if (ent.Type == ObjectTypeEnum.kSketchArcObject)
            {
                SketchArc sa = (SketchArc)ent;
                constraints.addConsid(sa.CenterSketchPoint as SketchEntity, fEnt.bp as SketchEntity);
            } 
        }
        public void addEq()
        {
            IEnumerable<LineInv> ls = u.gets<LineInv>(ents, fi => fi.con.HasFlag(cType.eq));
            var gr = ls.GroupBy(e => e.L); 
            foreach (var item in gr)
            {
                bool f = true;
                LineInv s = null;
                foreach (var e in item)
                {
                    if (item.Count() == 1) continue;
                    if (f)
                    {
                        s = e;
                        constraints.addTwoPointDist(s.getSl, s.getName, 1, 0.2);
                        if (!isNull(s.ls)) constraints.setVal(s.ls);
                        f = false;
                    }
                    else
                    {
                        constraints.addEq(s.getSl, e.getSl);
                    }
                }    
            }     
        }
        public void aConstr()
        {
            if (constr == null) return;
            foreach (var item in constr)
            {
                string s1 = getElem(item, 0, ';'), s2 = getElem(item, 1, ';'), s3 = getElem(item, 2, ';');
                SketchEntity se1 = null, se2 = null;
                se1 = getEnt<SketchEntity>(s1);
                se2 = getEnt<SketchEntity>(s2);
                //u.higlight(doc, se1, se2, I.objs.CreateColor(0,255,0), I.objs.CreateColor(0, 0, 255));
                if (isNull(se1)) continue;
                cType ct; Enum.TryParse(getElem(s3, 0), out ct);
                if (isNull(ct)) continue;
                if (ct.HasFlag(cType.con)) constraints.addConsid(se1, se2);
                if (ct.HasFlag(cType.col)) constraints.addCollinear(se1, se2);
                if (ct.HasFlag(cType.par)) constraints.addPar(se1, se2);
                if (ct.HasFlag(cType.perp)) constraints.addPerp(se1, se2);
                if (ct.HasFlag(cType.h)) constraints.addHor(se1);
                if (ct.HasFlag(cType.v)) constraints.addVert(se1);
            }
        }
        public T getEnt<T>(string s) where T: class
        {
            string t = getElem(s, 1), n = getElem(s, 0), snum = getElem(s, 2), filter = getElem(s, 3);
            bool f = false;
            if (!isNull(filter)) f = true;
            if (typeof(T) == typeof(SketchEntity))
            {
                return isNull(snum) ? getSE(n, t, fi => fi.Construction == f) as T : getSP(n, t, u.convToInt(snum), fi => fi.Construction == f) as T;
            }
            else if (typeof(T) == typeof(SketchPoint))
            {
                return getSP(n, t, u.convToInt(snum), fi => true) as T;
            }
            return null;
        }
        public void aDims()
        {
            if (dims == null) return;
            foreach (var item in dims)
            {
                string s1 = getElem(item, 0, ';'), s2 = getElem(item, 1, ';'), s3 = getElem(item, 2, ';');
                SketchLine sl1 = getSln(s1, fi => true), sl2 = getSln(s2, fi => true);
                if (isNull(sl1) || isNull(sl2))
                {
                    SketchPoint sp = getEnt<SketchPoint>(s1);
                    SketchPoint ep = getEnt<SketchPoint>(s2);
                    if (isNull(sp, ep)) continue;
                    string vtype = getElem(s3, 1), d = getElem(s3, 0);
                    if (isNull(d, vtype)) continue;
                    if (vtype.ToLower() == "v") constraints.addTwoPointDist(sp, ep, "", 1, 0.8, DimensionOrientationEnum.kVerticalDim);
                    else if (vtype.ToLower() == "h") constraints.addTwoPointDist(sp, ep, "", 1, 0.8, DimensionOrientationEnum.kHorizontalDim);
                    else constraints.addTwoPointDist(sp, ep, "", 1, 0.8);
                    constraints.setVal(d);
                    continue;
                }
                //double angl = sl1.Geometry.Direction.AngleTo(sl2.Geometry.Direction);
               // if (u.eq(angl, Math.PI, 1) || u.eq(angl, 0, 1))
               //     constraints.addPar(sl1 as SketchEntity, sl2 as SketchEntity);
                constraints.addTwoPointDist(sl1, sl2.StartSketchPoint);
                if (!isNull(s3)) constraints.setVal(s3);
            }
        }

        public SketchEntity getSE(string n, string t, Func<SketchEntity,bool> f)
        {
            if (isNull(t)) return null;
            switch (t)
            {
                case "a":
                    SketchArc sa = getSA(n, fi => f(fi as SketchEntity));
                    if (isNull(sa)) return null;
                    return sa as SketchEntity;
                case "l":
                    SketchLine sl = getSln(n, fi => f(fi as SketchEntity));
                    if (isNull(sl)) return null;
                    return sl as SketchEntity;
                case "p":
                    SketchPoint sp = getSP(n);
                    return sp as SketchEntity;
                default:
                    break;
            }
            return null;
        }
        public SketchPoint getSP(string n, string t, int num, Func<SketchEntity, bool> f)
        {
            if (isNull(t, num)) return null;
            switch (t)
            {
                case "a":
                    SketchArc sa = getSA(n, fi => f(fi as SketchEntity));
                    if (isNull(sa)) return null;
                    return num == 1 ? sa.StartSketchPoint: sa.EndSketchPoint;
                case "l":
                    SketchLine sl = getSln(n, fi => f(fi as SketchEntity));
                    if (isNull(sl)) return null;
                    return num == 1 ? sl.StartSketchPoint: sl.EndSketchPoint;
                case "p":
                    SketchPoint sp = getSP(n);
                    return sp;
                default:
                    break;
            }
            return null;
        }
        public SketchLine getSln(string n, Func<SketchLine,bool> f)
        {
            int i = u.convToInt(n);
            if (i == 0) return null;
            var ie = u.gets<SketchLine>(ps.SketchLines, fi => f(fi));
            if (ie.Count() >= i)
            return ie.ElementAt(i-1);
            return null;
        }
        public SketchArc getSA(string n, Func<SketchArc,bool> f)
        {
            int i = u.convToInt(n);
            if (i == 0) return null;
            var ie = u.gets<SketchArc>(ps.SketchArcs, fi => f(fi));
            if (ie.Count() >= i)
                return ie.ElementAt(i - 1);
            return null;
        }
        public SketchPoint getSP(string n)
        {
            int i = u.convToInt(n);
            if (i == 0) return null;
            var ie = u.gets<SketchPoint>(ps.SketchPoints, fi => true);
            if (ie.Count() > i)
                return ie.ElementAt(i);
            return null;
        }
        public void remDims()
        {
            fents = u.gets<Entity>(ents, fi => !isNull(fi.noDim));
            u.action<Entity>(fents, a =>
            {
                remDim(a);
            });
        }
        public void remDim(Entity a)
        {
            var spl = u.getSpl(a.noDim, ';');
            foreach (var item in spl)
            {
                int i = u.convToInt(item);
                if (i != 0 && a.constrts.Count >= i)
                    a.constrts[i - 1].Delete(); 
            } 
        }
        public void addHor()
        {
            fents = u.gets<LineInv>(ents, fi => fi.con.HasFlag(cType.h));
            u.action<LineInv>(fents, a => constraints.addHor(a.ent));
        }
        public void addVert()
        {
            fents = u.gets<LineInv>(ents, fi => fi.con.HasFlag(cType.v));
            u.action<LineInv>(fents, a => constraints.addVert(a.ent));
        }
        public void addPar()
        {
            fents = u.gets<LineInv>(ents, fi => fi.con.HasFlag(cType.par));
            u.action<LineInv>(fents, a => constraints.addPar(getSE(a.numSE), a.ent));
        }
        public void addPerp()
        {
            fents = u.gets<LineInv>(ents, fi => fi.con.HasFlag(cType.perp));
            u.action<LineInv>(fents, a => constraints.addPerp(getSE(a.numSE), a.ent));
        }
        public SketchEntity getSE(int n)
        {
            if (isNull(n) || ps.SketchLines.Count < n) return null;
            return ps.SketchLines[n] as SketchEntity;
        }
        public void setDimCen()
        {
            u.action<DimensionConstraint>(constraints.ps.DimensionConstraints, 
                a => setDimCen(a), fi => fi.Type == ObjectTypeEnum.kTwoPointDistanceDimConstraintObject);
        }
        public void setDimCen(DimensionConstraint a)
        {
            TwoPointDistanceDimConstraint d = a as TwoPointDistanceDimConstraint;
            var mp = u.midPt(d.PointOne.Geometry, d.PointTwo.Geometry, 1, 0, -0.2);
            a.TextPoint = mp;
        }
        public override void draw()
        {
            if (ents.Count() == 0) return;
            ObjectCollection col = null; col = I.COC(col);
            for (int i = 0; i < ents.Count(); i++)
            {
                if (ents[i].first)
                {
                    Entity.fEnt = ents[i];
                    setOrig();
                }
                if (i == 0) 
                {
                    ents[i].add(); 
                }
                else
                {
                    if (isNull(ents[i].prev))
                    ents[i].prev = ents[i - 1];
                    ents[i].add();
                }
                //if (!isNull(angs))
                if (!isNull(ents[i].ent) && !ents[i].ent.Reference)
                   col.Add(ents[i].ent);
                if (ents[i].addAngle && Entity.fEnt.t == entTypes.Line)
                {
                    constraints.addTwoLineAngle(Entity.fEnt.ent as SketchLine, ents[i].ent as SketchLine);
                }
            }
            
            addConstraint();
            if (!addCenter())
            {
                if (!isNull(cNum))
                {
                    int n = int.Parse(cNum);
                    merge(ents[n - 1], Entity.fEnt.bp, col);
                }
                else merge(Entity.fEnt, Entity.fEnt.bp, col);
            }
            remDims();
            addEq();
            addPar(); addPerp();
            addHor(); addVert();
            setDimCen();
            aConstr();
            aDims();
            if (!isNull(angs))
            {
                SketchLine l = u.addLine(ps ,Entity.fEnt.bp, u.getDirect(Entity.fEnt.ent), true);
                constraints.addFixed(l.EndSketchPoint);
                ps.RotateSketchObjects(col, Entity.fEnt.bp.Geometry, u.degToRad(u.convToDouble(angs)));
                constraints.addTwoLineAngle((SketchLine)Entity.fEnt.ent, l);
                constraints.setVal(angs);
                constraints.removeFixed(l);
                if (!isNull(rdir))
                {
                    if (rdir == "h") constraints.addHor(l);
                    else if (rdir == "v") constraints.addVert(l);
                }
                
            }
            update = false;
        }

        public override void upd()
        {
            for (int i = 0; i < ents.Count(); i++)
            {
                if (i == 0)
                    ents[i].add();
                else
                {
                    ents[i].prev = ents[i - 1];
                    ents[i].add();
                }
            }
        }

        public override string ToString()
        {
            XElement el = new XElement("Sketch");
            Dictionary<string,string> dic = new Dictionary<string, string>() { { "Name", ps.Name }};
            if (ps.PlanarEntity is WorkPlane) dic.Add("Plane", (ps.PlanarEntity as WorkPlane).Name);
            else 
            {
                Face f = ps.PlanarEntity as Face;
                Point pt = f.PointOnFace;
                dic.Add("PlanePoint", pt.X.ToString("#.###") + " " + pt.Y.ToString("#.###") + " " + pt.Z.ToString("#.###"));
            }
            XMLDoc.addXAttributes(el, dic);
            xml.addXElement(el);
            xml.El = el;
            return el.ToString();
        }

        public override void get()
        {
            string tmp = "";
            this.ToString();
            if (closed)
            {
                ents.Remove(ents[ents.Count - 1]);
            }
            foreach (Entity item in ents)
            {
                tmp += item.ToString();  
            }
        }
    }

    class ClusterInv : FeaturesInv
    {

        HoleFeature f;
        string d;
        string dx, dy;
        new SketchHolePlacementDefinition def;

        public ClusterInv()
        {
            t = entTypes.Cluster;
            if (!checkAsm()) return;
            alias = "Кл";
            get();
        }

        protected override void addDef()
        {
            throw new NotImplementedException();
        }

        public override void draw()
        {
            throw new NotImplementedException();
        }

        public override void setName()
        {
            throw new NotImplementedException();
        }

        public override void upd()
        {
            throw new NotImplementedException();
        }

        public override void get()
        {
            base.get();
            d = xml.getAttributeValue("d");
        }
    }

    class HoleInv : FeaturesInv
    {
        HoleFeature f;
        string d;
//         Point spoint;
//         ObjectCollection col;
//         UnitVector normal;
//         PlanarSketch nSketch;
        new SketchHolePlacementDefinition def;
        public HoleInv()
        {
            t = entTypes.Hole;
            if (!checkAsm()) return;
            alias = "О";
            get();
            ips = getSketch();
            if (checkSketch()) return;
            setName();
            if (!isNull(name) && find(name)) return;
            addDef();                    
        }

        public override void draw()
        {
            if (def == null) return;
            f = smf.HoleFeatures.AddDrilledByDistanceExtent(def, d, thick.Name, PartFeatureExtentDirectionEnum.kPositiveExtentDirection);
            doc.Update();
            if (f.HealthStatus == HealthStatusEnum.kDriverLostHealth) ((DistanceExtent)f.Extent).Direction = PartFeatureExtentDirectionEnum.kNegativeExtentDirection;
            f.Name = fname();
            //addIMate();
        }

        public override void setName()
        {
            name = xml.getAttributeValue("Name");
            if (isNull(name))
            {
                if (isNull(ips)) return;
                name = alias + getDouble(d, 1, 2) + ips.Name;
            }
        }

        public void addIMate()
        {
            var atts = xml.getXAttributesValues(new string[] { "IMate", "IMateName", "IMateOffset", "rev", "IMateMatch" });
            if (atts[0] == null) return;
            var spl = atts[0].Split(';');
            if (spl == null) return;
            object e1 = getEdge(spl[0], atts[3]), e2 = getEdge(spl[1], atts[3]);
            IMate.iInsComposite(e1, e2, base.def as PartComponentDefinition, u.convToDouble(atts[2]), atts[1], atts[4]);
        }

        public object getEdge(string num, string rev)
        {
            base.addEdge();
            getEdges();
            base.getEdges();
            return true;
//             int i = u.convToInt(num), e = 2;
//             if (rev == null || rev == "0") e = 1;
//             return f.SideFaces[i].Edges[e];
        }

        public override void getEdges()
        {
            int i = 1;
            getFaces(fi => true);
            getEdges((fi, num) => num == i);
        }

//         public bool find()
//         {
//             if (ips.PlanarEntity is Face) return false;
//             normal = ips.PlanarEntityGeometry.Normal;
//             Face f1 = findFace();
//             Vector v = normal.AsVector();
//             v.ScaleBy(-1);
//             normal = v.AsUnitVector();
//             Face f2 = findFace();
//             if (f1 == null && f2 == null) return false;
//             f1 = distToFace(f1, spoint) < distToFace(f2, spoint) ? f1: f2;
//             return addSketch(f1);
//         }

//         public double distToFace(Face f, Point pt)
//         {
//             if (f == null) return 100000;
//             return I.app.MeasureTools.GetMinimumDistance(f, pt);
//         }

//         public Face findFace()
//         {
//             ObjectsEnumerator founds, locs;
//             base.def.FindUsingRay(spoint, normal, 0.1, out founds, out locs, true);
//             if (founds.Count != 0 && founds[1] is Face) return founds[1] as Face;
//             return null;
//         }
// 
//         public bool addSketch(Face f)
//         {
//             nSketch = base.def.Sketches.Add(f);
//             nSketch.Visible = false;
//             if (col.Count > 0)
//             {
//                 for (int i = 0; i < col.Count; i++)
//                 {
//                     nSketch.AddByProjectingEntity(col[i+1]);
//                 }
//                 ips = nSketch;
//                 return true;
//             }
//             return false;
//         }

        public override void upd()
        {
          
        }

        public override void get()
        {
            base.get();
            d = xml.getAttributeValue("d");
        }

        protected override void addDef()
        {
            col = I.COC(col);
            foreach (SketchPoint pt in ips.SketchPoints)
            {
                if (pt.HoleCenter == true)
                {
                    if (spoint == null)
                    {
                        spoint = pt.Geometry3d;
                    }
                    col.Add(pt);
                }
            }
            if (findPlaneBySketch()) addDef();
            if (base.addEdge())
            {
                List<string> num = spl.ToList();
                addEdges(fi => num.Contains(fi.ToString()));
            }
            def = smf.HoleFeatures.CreateSketchPlacementDefinition(col);
        }
        protected override void addEdges(Func<int, bool> f)
        {
            ObjectCollection tcol = null; tcol = I.COC(tcol);
            int i = 0;
            foreach (var item in col)
            {
                i++;
                if (f(i)) tcol.Add(item);
            }
            col = tcol;
        }
    }

    class PunchFeat : IFeatFeat
    {
        PunchToolFeature f;
        string a;
        ObjectCollection col;
        public PunchFeat()
        {
            path = @"C:\Users\Public\Documents\Autodesk\Inventor 2014\Catalog\Punches\";
            t = entTypes.Punch;
            if (!checkAsm()) return;
            alias = "Пр";
            get();
            a = xml.getAttributeValue("a");
            if (isNull(a)) a = "0";
            //a += " град";

            if (checkSketch() || fn == null) return;
            addDef();
        }
        protected override void addDef()
        {
            if (findPlaneBySketch()) addDef();
            def = smf.PunchToolFeatures.CreateiFeatureDefinition(fn);
            set();
            col = I.COC(col);
            u.action<SketchPoint>(ips.SketchPoints, a => col.Add(a), f => f.HoleCenter);
        }

        public override void draw()
        {
            if (def == null) return;
            f = smf.PunchToolFeatures.Add(col, def, a);
            f.Name = fname();
        }

        public override void upd()
        {
        }
    }

    class IFeatFeat : FeaturesInv
    {
        iFeature f;
        protected new iFeatureDefinition def;
        protected string[] pars;
        protected string fn;
        protected string path;
        public IFeatFeat()
        {
            path = @"C:\Users\Public\Documents\Autodesk\Inventor 2014\Catalog\";
            t = entTypes.iFeature;
            if (!checkAsm()) return;
            alias = "П";
            if (!checkPF(xml.getAttributeValue("MustExist"))) { update = true; return; }
            get();
            if (checkSketch() && fn == null) return;
            if (!getGeom()) return;
            addDef();
        }

        protected override void addDef()
        {
            if (findPlaneBySketch()) addDef();
            def = smf.iFeatures.CreateiFeatureDefinition(fn);
            u.get<iFeatureSketchPlaneInput>(def.iFeatureInputs, f => true).PlaneInput = ips.PlanarEntity;
            set();
        }

        protected void set()
        {
            int i = 0;
            try
            {
                foreach (var item in u.gets<iFeatureParameterInput>(def.iFeatureInputs, f => true))
                {
                    if (pars.Length > i && !isNull(pars[i])) item.Expression = pars[i];
                    i++;
                }
            }
            catch (System.Exception ex)
            {
            	
            }
        }

        public override void get()
        {
            base.get();
            ips = getSketch();
            if (isNull(ips)) return;
            try
            {
                foreach (var item in ips.Dependents)
                {
                    if (item is PunchToolFeature || item is iFeature) { ips = null; return; }
                }
            }
            catch (Exception)
            {
            }
            pars = getSpl("Elements");
            fn = path + xml.getAttributeValue("FeatName");
        }

        protected bool getGeom()
        {
            if (isNull(ips)) return false;
            if (ips.SketchLines.Count == 0) return false;
            sl = u.get<SketchLine>(ips.SketchLines, f => f.Centerline);
            if (isNull(sl)) return false;
            startPt = u.get<SketchPoint>(ips.SketchPoints, f => f.HoleCenter);
            if (startPt == null) return false;
            wp = base.def.WorkPoints.AddByPoint(startPt);
            if (wp == null) return false;
            return true;
        }

        public override void draw()
        {
            if (def == null) return;
            f = smf.iFeatures.Add(def);
            ips = f.Sketches[1] as PlanarSketch;
            ips.OriginPoint = wp;
            wp.Visible = false;
            ips.AxisEntity = sl;
            f.Name = fname();
        }

        public override void upd()
        {
        }
    }

    class ArrayFeat : FeaturesInv
    {
        RectangularPatternFeature f;
        ObjectCollection col = null;
        string axisXName, axisYName;
        WorkAxis wax, way;
        string dx, dy, cx, cy, rx, ry, mx, my, ang, asm, fl;
        string[] elems;
        object ppf = null;
        public ArrayFeat()
        {
            t = entTypes.Array;
            alias = "M";
            setAlias();
            get();
        }

        protected override void addDef()
        {
        }

        public override void draw()
        {
            if (isNull(col) || col.Count == 0) return;
            if (isNull(asm))
            {
                drawRect();
                drawCirc();
            }
            else
            {
                drawAsm();
            }
        }

        void drawAsm()
        {
            oc = null; pf = null; num = 0;
            if (isNull(ppf)) return;
            if (!isNull(u.get<FeatureBasedOccurrencePattern>(acd.OccurrencePatterns, fi => fi.Name == name))) return;
            FeatureBasedOccurrencePattern fb = acd.OccurrencePatterns.AddFeatureBasedPattern(col, ppf as PartFeature);
            if (isNull(name)) return;
            fb.Name = name;
        }

        void drawRect()
        {
            if (isNull(dx)) return;
            bool x = getRev(rx), y = getRev(rx), mirx = getRev(mx), miry = getRev(my);
            PatternSpacingTypeEnum sp = PatternSpacingTypeEnum.kDefault;
            if (isNull(dy, way, cy))
            {
                f = smf.RectangularPatternFeatures.Add(col, wax, x, cx, dx, sp, ComputeType: PatternComputeTypeEnum.kAdjustToModelCompute);
                f.XDirectionMidPlanePattern = mirx;
            }
            else
            {
                f = smf.RectangularPatternFeatures.Add(col, wax, x, cx, dx, YDirectionEntity: way, NaturalYDirection: y, YCount: cy, YSpacing: dy,
                     ComputeType: PatternComputeTypeEnum.kAdjustToModelCompute);
                f.XDirectionMidPlanePattern = mirx;
                f.YDirectionMidPlanePattern = miry;
            }
            if (f.HealthStatus == HealthStatusEnum.kDriverLostHealth) f.ComputeType = PatternComputeTypeEnum.kIdenticalCompute;
            if (!isNull(fl)) f.XDirectionSpacingType = PatternSpacingTypeEnum.kFitted;
            f.Name = fname();
        }

        void drawCirc()
        {
            if (isNull(ang, wax, cx)) return;
            bool x = getRev(rx), mirx = getRev(mx);
            CircularPatternFeature cf = smf.CircularPatternFeatures.Add(col, wax, x, cx, ang, mirx, PatternComputeTypeEnum.kAdjustToModelCompute);
            if (cf.HealthStatus == HealthStatusEnum.kDriverLostHealth) cf.ComputeType = PatternComputeTypeEnum.kIdenticalCompute;
            cf.Name = fname();
        }

        public override void setName()
        {
            name = xml.getAttributeValue("Name");
            if (isNull(name))
            {
                PartFeature pf = col[1] as PartFeature;
                if (isNull(pf)) return;
                name = alias + pf.Name;
            }
        }

        public override void upd()
        {
        }

        public bool getRev(string val)
        {
            if (isNull(val)) return false;
            return val == "1";
        }

        protected bool check(System.Collections.IEnumerable cl ,string name)
        {
            if (!checkPF(xml.getAttributeValue("MustExist"))) return false;
            List<bool> r = new List<bool>();
            foreach (var item in cl)
            {
                RectangularPatternFeature rpf = item as RectangularPatternFeature;
                CircularPatternFeature cpf = item as CircularPatternFeature;
                ObjectCollection par = !isNull(rpf) ? rpf.ParentFeatures : !isNull(cpf) ? cpf.ParentFeatures : null;
                if (isNull(par)) return false;
                if (col.Count != par.Count) continue;
                //int i = 0;
                if (((PartFeature)par[1]).Name == ((PartFeature)col[1]).Name) return true;
//                 foreach (PartFeature feat in par)
//                 {
//                     i++;
//                     if (feat.Name != ((PartFeature)col[i]).Name) return false;
//                 }
                //return true;
            }
            return false;
        }

//         public bool findInArr()
//         {
//             foreach (RectangularPatternFeature item in smf.RectangularPatternFeatures)
//             {
//                 if (col.Count != item.ParentFeatures.Count) continue;
//                 int i = 0;
//                 foreach (PartFeature feat in item.ParentFeatures)
//                 {
//                     i++;
//                     if (feat.Name != ((PartFeature)col[i]).Name) return false;
//                 }
//                 return true;
//             }
//             return false;
//         }

        public override void get()
        {
            getCustom();
            name = xml.getAttributeValue("Name");
            axisXName = xml.getAttributeValue("ax");
            axisYName = xml.getAttributeValue("ay");
            ang = xml.getAttributeValue("Ang");
            dx = xml.getAttributeValue("dx");
            dy = xml.getAttributeValue("dy");
            cx = xml.getAttributeValue("cx");
            cy = xml.getAttributeValue("cy");
            fl = xml.getAttributeValue("FullL");
            asm = xml.getAttributeValue("Asm");
            if (!isNull(asm, acd))
            {
                string tmp = getElem(asm, 1);
                string n = getElem(asm, 2);
                if (!isNull(tmp)) num = int.Parse(tmp);
                asm = getElem(asm, 0);
                findPCD(acd.Occurrences, asm, null);
                if (isNull(pcd)) return;
                if (isNull(n)) return;
                pf = lastEntity<PartFeature>(pcd.Features, n + ":М", false).LastOrDefault();
                if (isNull(oc)) return;
                oc.CreateGeometryProxy(pf, out ppf);
                if (isNull(pf)) return;
                elems = getSpl("Parent");
                col = I.COC(col);
                findOcc(elems, ref col, null);
                return;
            }

            if (!getParam(dx, dy, cx, cy)) { br = true; return; }
            mx = xml.getAttributeValue("mx");
            my = xml.getAttributeValue("my");
            rx = xml.getAttributeValue("revx");
            ry = xml.getAttributeValue("revy");

            elems = getSpl("Parent");
            col = I.COC(col);
            findSpl(elems, ref col);
            if (check(smf.RectangularPatternFeatures, "a") || check(smf.CircularPatternFeatures, "a")) { col = I.COC(col); return; }
            if (!isNull(axisXName)) wax = findWA(axisXName);
            if (!isNull(axisYName)) way = findWA(axisYName);
        }
    }

    class PlaneFeat : FeaturesInv
    {
        string type, p, plName, axis, d, rev;
        Edge ed;
        public PlaneFeat()
        {
            t = entTypes.Plane;
            if (!checkAsm()) return;
            alias = "Пл";
            setAlias();
            get();
            if (fs == null) { return; }
        }

        protected override void addDef()
        {
        }

        public override void draw()
        {
            if (isNull(type)) return;
            switch (type)
            {
                case "между":
                    if (fs == null || fs.Count != 2) return;
                    addBetween();
                    break;
                case "оси":
                    if (fs == null || fs.Count != 2) return;
                    addByAxes();
                    break;
                case "центр ребра":
                    addByCenter();
                    break;
                case "смещение":
                    addByOffset();
                    break;
                case "угол":
                    addByAngle();
                    break;
                default:
                    break;
            }
            if (!isNull(rev)) wpl.FlipNormal();
            if (!isNull(wpl)) wpl.Name = name; 
        }
        public void addBetween()
        {
            wpl = pcd.WorkPlanes.AddByTwoPlanes(fs.ElementAt(0), fs.ElementAt(1));
            wpl.Name = name;
        }

        public void addByCenter()
        {
            if (isNull(ed)) return;
            wp = pcd.WorkPoints.AddByMidPoint(ed);
            wpl = pcd.WorkPlanes.AddByNormalToCurve(ed, wp);
        }

        public void addByOffset()
        {
            if (isNull(wpl)) return;
            wpl = pcd.WorkPlanes.AddByPlaneAndOffset(wpl, d);
        }

        public void addByAngle()
        {
            if (isNull(wa, wpl)) return;
            wpl = pcd.WorkPlanes.AddByLinePlaneAndAngle(wa, wpl, d);
        }

        public void addByAxes()
        {
            wa = pcd.WorkAxes.AddByRevolvedFace(fs.ElementAt(0));
            WorkAxis wa1 = base.def.WorkAxes.AddByRevolvedFace(fs.ElementAt(1));
            if (isNull(wa,wa1)) return;
            wpl = pcd.WorkPlanes.AddByTwoLines(wa, wa1);
            wa1.Visible = false;
        }

        public override void upd()
        {
        }

        public override bool addEdge()
        {
            base.addEdge();
            getEdges();
            return true;
        }

        public override void getEdges()
        {
            if (type == "смещение" || type == "угол") return;
            if (type == "между" || type == "центр ребра")
                getFaces(fi => fi.SurfaceType == SurfaceTypeEnum.kPlaneSurface);
            else if (type == "оси")
                getFaces(fi => fi.SurfaceType == SurfaceTypeEnum.kCylinderSurface);
            fs = filter(fs);
            if (fs.Count() == 0) return;
            if (type == "центр ребра")
            {
                ed = u.get<Edge>(fs.ElementAt(0).Edges, f => !u.eq(u.getLenght(f), thick.ModelValue));
            }
        }

        public override void get()
        {
            name = xml.getAttributeValue("Name");
            if (!isNull(u.findPlane(doc, name))) { br = true; return; }
            //if (!isNull(u.get<WorkPlane>(def.WorkPlanes, fi => fi.Name == name))) return;
            type = xml.getAttributeValue("type");
            if (isNull(type)) return;
            p = xml.getAttributeValue("Parent");
            plName = xml.getAttributeValue("Plane");
            axis = xml.getAttributeValue("axis");
            rev = xml.getAttributeValue("rev");
            if (!isNull(axis)) wa = findWA(axis);
            d = xml.getAttributeValue("d");
            u.addParameter(doc, d, pars);
            if (!isNull(plName)) wpl = findWP(plName);
            find(p);

            addEdge();
        }
    }

    class FlatPatternFeat : FeaturesInv
    {
        FlatPattern f;
        Face bf;
        object e;
        string a, snum, srev, flip, bends;
        int num; bool rev = false;
        AlignmentTypeEnum al = AlignmentTypeEnum.kHorizontalAlignment;

        public FlatPatternFeat()
        {
            t = entTypes.FlatPattern;
            bends = xml.getAttributeValue("bends");
            if (def.HasFlatPattern) { update = true; f = def.FlatPattern; return; }
            get();
        }

        protected override void addDef()
        {
           
        }

        public override void draw()
        {
            if (def.HasFlatPattern) return;
            def.Unfold2(bf);
            f = def.FlatPattern;
            f.Edit();
            num = u.getNum(snum, f.FlatBendResults.Count);
            if (isNull(num)) num = 1;
            if (f.FlatBendResults.Count >= num)
                e = f.FlatBendResults[num].Edge;
            else
                f.GetAlignment(out al, out e, out rev);
            if (!isNull(a)) al = AlignmentTypeEnum.kVerticalAlignment;
            if (!isNull(srev)) rev = true;
            f.SetAlignment(al, e, rev);
            f.ExitEdit();
        }

        public override void upd()
        {
            if (isNull(bends)) return;
            var spl = u.getSpl(bends, ';');
            int b;
            for (int i = 0; i < f.FlatBendResults.Count; i++)
            {
                if (spl.Length <= i) continue;
                b = u.convToInt(spl[i]);
                if (b == 0) continue;
                f.FlatBendResults[i+1].SetBendOrder(b, true); 
            }
        }

        public override void get()
        {
            pf = u.get(smf, 1) as PartFeature;
            Faces fs = pf.Faces;
            flip = xml.getAttributeValue("flip");
            if (pf is FaceFeature)
            {
                FaceFeature ff = pf as FaceFeature;
                Point pt = ((PlanarSketch)ff.Definition.Profile.Parent).OriginPointGeometry;
                var ie = u.gets<Face>(fs, fi =>
                {
                    return isNull(u.get<Edge>(fi.Edges, e => flen(e)));
                });

                bf = u.get<Face>(ie, fi => !u.eq(fi, pt));
            }
            else if (pf is ContourFlangeFeature)
            {
                bf = u.get<Face>(fs, e => e.TangentiallyConnectedFaces.Count == 0);
                var eds = u.gets<Edge>(bf.Edges, e => !flen(e));
                bf = getFace(eds);
            }
            else if (pf is ReferenceFeature)
            {
                var ie = u.gets<Face>(fs, e => ((Face)e.ReferencedEntity).CreatedByFeature.Type == ObjectTypeEnum.kFaceFeatureObject);
                FaceFeature ff = ((Face)ie.ElementAt(0).ReferencedEntity).CreatedByFeature as FaceFeature;
                Point pt = ((PlanarSketch)ff.Definition.Profile.Parent).OriginPointGeometry;
                ie = u.gets<Face>(ie, fi =>
                {
                    return isNull(u.get<Edge>(fi.Edges, e => flen(e)));
                });
                bf = u.get<Face>(ie, fi => !u.eq(fi, pt));
            }
            a = xml.getAttributeValue("Align");
            snum = xml.getAttributeValue("num"); //new
            srev = xml.getAttributeValue("rev");
        }

        public Face getFace(IEnumerable<Edge> eds)
        {
            Edge e = eds.ElementAt(0);
            if (!isNull(flip)) e = eds.ElementAt(1);
            return getFace(e);
        }

        public Face getFace(Edge e)
        {
            return e.Faces[1].Equals(bf) ? e.Faces[2] : e.Faces[1];
        }
    }

    class MirrorFeat : FeaturesInv
    {
        MirrorFeature f;
        ObjectCollection col = null;
        string plName, pfName;
        WorkPlane pl;
        string[] elems;
        public MirrorFeat()
        {
            t = entTypes.Mirror;
            if (!checkAsm()) return;
            alias = "З";
            setAlias();
            get();
            if (isNull(pl, col)) return;
        }

        protected override void addDef()
        {
            
        }

        protected override bool check(string name)
        {
            if (!checkPF(xml.getAttributeValue("MustExist"))) return false;
            PartFeature tpf = u.get<PartFeature>(smf, e => e.Name == name);
            if (isNull(tpf)) return true;
            return false;
//             MirrorFeature mpf = tpf as MirrorFeature;
//             if (isNull(mpf)) return true;
//             List<string> els = elems.ToList();
//             var c = u.gets<PartFeature>(mpf.ParentFeatures, e => els.Contains(e.Name));
//             if (c.Count() == mpf.ParentFeatures.Count) return false;
//             return true;

            
        }

        public void addCol()
        {
//             foreach (var item in cl)
//             {
//                 MirrorFeature rpf = item as MirrorFeature;
//                 ObjectCollection par = !isNull(rpf) ? rpf.ParentFeatures : null;
//                 if (isNull(par)) return false;
//                 if (col.Count != par.Count) continue;
//                 int i = 0;
//                 foreach (PartFeature feat in par)
//                 {
//                     i++;
//                     if (feat.Name != ((PartFeature)col[i]).Name) return false;
//                 }
//                 return true;
//             }
        }

        public override void setName()
        {
            name = xml.getAttributeValue("Name");
            if (isNull(name))
            {
                if (col.Count == 0) return;
                PartFeature pf = col[1] as PartFeature;
                if (isNull(pf)) return;
                name = alias + pf.Name;
            }
        }

        public override void draw()
        {
            if (col.Count == 0) return;
            f = smf.MirrorFeatures.Add(col, pl, false, PatternComputeTypeEnum.kAdjustToModelCompute);
            if (f.HealthStatus == HealthStatusEnum.kDriverLostHealth) f.ComputeType = PatternComputeTypeEnum.kIdenticalCompute;
            f.Name = fname();
        }

        public override void upd()
        {
        }

        public override void get()
        {
            getCustom();
            name = xml.getAttributeValue("Name");
            plName = xml.getAttributeValue("Plane");
            elems = getSpl("Elements");
            col = I.COC(col);
            findSpl(elems, ref col);
            pl = findWP(plName);
            setName();
            if (!check(name))
            {
                col.Clear();
            }
        }
    }

    class UnfoldFeat : FeaturesInv
    {
        UnfoldFeature f;
        RefoldFeature r;
        Face fa;
        string p, snum;
        string[] numBends;
        int num;
        public UnfoldFeat()
        {
            t = entTypes.Unfold;
            if (!checkAsm()) return;
            alias = "Раз";
            setName();
            string me = xml.getAttributeValue("MustExist");
            if (!checkPF(me)) return;
            if (!isNull(name) && find(name)) { update = true; return; }
            get();
        }

        protected override void addDef()
        {

        }

        public Bend getBend(int i)
        {
            if (i == 0 || i > base.def.Bends.Count) return null;
            return base.def.Bends[i];
        }

        public override void draw()
        {
            if (isNull(fa)) return;
            object tm = Type.Missing;
            if (!isNull(numBends))
            {
                col = I.COC(col);
                u.action(numBends, b => { 
                    Bend be = getBend(u.convToInt(b));
                    if (!isNull(be)) col.Add(be);
                });
                if (col.Count != 0) tm = col; 
            }

            f = smf.UnfoldFeatures.Add(fa, tm);
            f.Name = fname();
            r = smf.RefoldFeatures.Add(f.StationaryFace);
            r.SetEndOfPart(true);
            alias = "Св";
            r.Name = fname();
        }

        public override void upd()
        {
            pf = u.get<PartFeature>(smf, f => f.Type == ObjectTypeEnum.kRefoldFeatureObject);
            if (isNull(pf)) return;
            pf.SetEndOfPart(true);
        }

        public override void get()
        {
            base.get();
            snum = xml.getAttributeValue("num"); //new
            fa = numFace(xml.El);
            numBends = getSpl("Bends");
            if (isNull(fa))
                fa = getFace();
        }
        public override void setName()
        {
            name = xml.getAttributeValue("Name");
            if (isNull(name))
            {
                if (isNull(p))
                {
                    pf = lastEntity<PartFeature>(count: lplus()).LastOrDefault();
                    if (isNull(pf)) return;
                }
                else if (!find(p)) return;
                name = alias + pf.Name;
            }
            else name = alias + name;
        }
        public Face getFace(Faces col = null)
        {
            if (isNull(col))
            {
                SurfaceBody sb = pcd.SurfaceBodies[1];
                if (sb.CreatedByFeature is ContourRollFeature)
                    return sb.Faces[sb.Faces.Count];
                col = sb.Faces;
            }
            var fs = getPlaneFaces(col);
            if (isNull(fs)) return null;
            num = u.getNum(snum, fs.Count());
            if (isNull(num)) num = 1;
            return fs.ElementAt(num - 1);
        }
    }

    class LoftFlangeFeat : FeaturesInv
    {
        LoftedFlangeFeature f;
        new LoftedFlangeDefinition def;
        Path p1, p2;
        string d;
        string dir, ss, rev, sty, w;
        PlanarSketch sk1, sk2;
        Parameter pr = null;
        PartFeatureExtentDirectionEnum ext = PartFeatureExtentDirectionEnum.kPositiveExtentDirection;
        LoftedFlangeOutputTypeEnum ty = LoftedFlangeOutputTypeEnum.kDieFormedLoftedFlange;

        public LoftFlangeFeat()
        {
            t = entTypes.Loft;
            if (!checkAsm()) return;
            alias = "Л";
            get();
            ips = getSketch();
            if (checkSketch()) return;
            ss = xml.getAttributeValue("SecondSketch");
            sk1 = ips;
            if (!isNull(ss))
            {
                name = ss;
                sk2 = getSketch();
            }
            if (!isNull(u.get<object>(ips.Dependents, f => f is LoftedFlangeFeature))) return;
            addDef();
        }

        protected override void addDef()
        {
            if (isNull(ips)) return;

            if (dir == "1") d = "-" + d;
            else if (dir == "2") d = d + "/2";
            if (isNull(sk2))
            {
                if (!findPlaneBySketch(d)) return;
                sk2 = ips;
                if (dir == "2")
                {
                    d = "-" + d; sk1 = sk2;
                    if (!findPlaneBySketch(d)) return;
                    sk2 = ips;
                }
            }
            if (isNull(sk1, sk2)) return;
            p1 = getPath(sk1); p2 = getPath(sk2);
            if (isNull(p1,p2)) return;
            def = smf.LoftedFlangeFeatures.CreateLoftedFlangeDefinition(p1, p2);
            def.ExtentDirection = ext;
            if (!isNull(w)) def.SetOutputType(ty, w);
        }

        public override void draw()
        {
            if (isNull(def)) return;
            f = smf.LoftedFlangeFeatures.Add(def);
            f.Name = fname();
        }

        public override void upd()
        {

        }

        public override void get()
        {
            base.get();
            d = xml.getAttributeValue("d");
            dir = xml.getAttributeValue("dir");
            rev = xml.getAttributeValue("rev");
            w = xml.getAttributeValue("w");
            if (rev == "1") ext = PartFeatureExtentDirectionEnum.kNegativeExtentDirection;
            else if (rev == "2") ext = PartFeatureExtentDirectionEnum.kSymmetricExtentDirection;
            sty = xml.getAttributeValue("type");
            if (sty == "2") ty = LoftedFlangeOutputTypeEnum.kPressBrakeFacetAngleLoftedFlange;
            else if (sty == "3") ty = LoftedFlangeOutputTypeEnum.kPressBrakeFacetDistanceLoftedFlange;
            else if (sty == "1") ty = LoftedFlangeOutputTypeEnum.kPressBrakeChordToleranceLoftedFlange;
        }
    }

    class CFlangeFeat : FeaturesInv
    {
        ContourFlangeFeature f;
        new ContourFlangeDefinition def;
        Path p;
        string d;
        string rev, w;
        PartFeatureExtentDirectionEnum ext = PartFeatureExtentDirectionEnum.kPositiveExtentDirection;
        PartFeatureExtentDirectionEnum dist = PartFeatureExtentDirectionEnum.kSymmetricExtentDirection;
        Parameter pr = null;
        public CFlangeFeat()
        {
            t = entTypes.CountourFlange;
            alias = "К";
            get();
            ips = getSketch();
            if (checkSketch()) return;
            try
            {
                if (!isNull(u.get<object>(ips.Dependents, f => f is ContourFlangeFeature))) return;
            }
            catch (System.Exception ex)
            {
            }
            addDef();
        }

        protected override void addDef()
        {
            p = getPath(ips);
            if (p == null) return;
            def = smf.ContourFlangeFeatures.CreateContourFlangeDefinition(p);
            if (pr != null) d = pr.Name;
            def.ExtentDirection = ext;
            def.SetDistanceExtent(d, dist);
        }

        public override string ToString()
        {
            return "";
        }

        public override void draw()
        {
            if (def == null) return;
            f = smf.ContourFlangeFeatures.Add(def);
            f.Name = fname();
        }

        public override void get()
        {
            base.get();
            d = xml.getAttributeValue("d");
            u.addParameter(doc, d, pars);
            rev = xml.getAttributeValue("rev");
            if (rev == null || rev == "") ext = PartFeatureExtentDirectionEnum.kNegativeExtentDirection;
            else if (rev == "1") ext = PartFeatureExtentDirectionEnum.kPositiveExtentDirection;
            else if (rev == "2") ext = PartFeatureExtentDirectionEnum.kSymmetricExtentDirection;
            w = xml.getAttributeValue("dir");
            if (w == null || w == "") dist = PartFeatureExtentDirectionEnum.kSymmetricExtentDirection;
            else if (w == "1") dist = PartFeatureExtentDirectionEnum.kPositiveExtentDirection;
            else if (w == "2") dist = PartFeatureExtentDirectionEnum.kNegativeExtentDirection;
        }

        public override void upd()
        {
        }
    }

    class FlangeFeat : FeaturesInv
    {
        FlangeFeature f;
        FlangeDefinition def;
        string p, me;
        string d;
        string a;
        string[] fil, opt; 
        bool rev = false;
        public FlangeFeat()
        {
            t = entTypes.Flange;
            if (!checkAsm()) return;
            alias = "Фл";
            me = xml.getAttributeValue("MustExist");
            if (!checkPF(me)) return;
            setAlias();
            setName();
            if (!isNull(name) && find(name)) { f = (FlangeFeature)pf; def = f.Definition; update = true; return; }
            if (!check()) return;
            get();
            if (isNull(pf)) return;
            addDef();
        }

        public bool check()
        {
            if (!isNull(pf)) return true;
            if (pcd.ReferenceComponents.DerivedPartComponents.Count == 0) return true;
            DerivedPartComponent der = pcd.ReferenceComponents.DerivedPartComponents[1];
            if (isNull(der.SolidBodies)) return true;
            if (der.SolidBodies.Count == 0) return true;
            return false;
        }

        public override string ToString()
        {
            throw new System.NotImplementedException();
        }

        public override void setName()
        {
            name = xml.getAttributeValue("Name");
            if (isNull(name))
            {
                if (isNull(p))
                {
                    pf = lastEntity<PartFeature>(count: lplus()).LastOrDefault();
                    if (isNull(pf)) return;
                }
                else if (!find(p)) return;
                name = alias + pf.Name;
            }
        }

        public override void draw()
        {
            if (isNull(def)) return;
            BendOptions o;
//             if (!isNull(opt[0]))
//             {
//                 o = def.BendOptions;
//                 switch (opt[0])
//                 {
//                     case "Пр":
//                         o.BendReliefShape = BendReliefShapeEnum.kStraightBendReliefShape;
//                         break;
//                     case "Скр":
//                         o.BendReliefShape = BendReliefShapeEnum.kRoundBendReliefShape;
//                         break;
//                     default:
//                         break;
//                 }
//             }
            f = smf.FlangeFeatures.Add(def);

            o = f.Definition.BendOptions;
            if (!isNull(opt[1])) setPar(o.BendReliefWidth, opt[1]);
            if (!isNull(opt[2])) setPar(o.BendReliefDepth, opt[2]);
            if (!isNull(opt[3])) setPar(o.MinimumRemnant, opt[3]);
//             if (!isNull(opt[4])) 
//             {
//                 setPar(f.Definition.BendRadius, opt[4]);
//             }
            readXML();
            f.Name = fname();
        }

        public void setPar(object o, string val)
        {
            Parameter p = o as Parameter;
            if (isNull(p)) { o = val; return; }
            p.Expression = val;
        }

        public void readXML()
        {
            XElement p = xml.El;
            f.SetEndOfPart(true);
            //object e1 = null, e2 = null;
            //int i = s0;
            if (!isNull(opt[4]))
            {
                switch (opt[4])
	            {
                    case "o":
                        def.BendPosition = BendPositionEnum.kBendPositionAdjacentFace;
                        break;
		            default:
                     break;
	            }
            }
            foreach (var el in xml.El.Elements())
            {
                xml.El = el;
                var vals = XMLDoc.getXAttributesValues(el, new string[] { "num", "type", "d1", "d2" }); //new
                if (isNull(def.Edges)) return;
                int num = u.getNum(vals[0], def.Edges.Count);
                IEnumerable<Edge> eds = def.Edges.OfType<Edge>();
                string s = getElem(vals[0], 1);
                if (!isNull(s)) eds = u.sort<Edge>(eds, s);
                if (num == 0 || num > eds.Count()) return;
                Edge ed = eds.ElementAt(num-1);
                if (isNull(ed)) return;
                //i++;
                //u.clientLine(pcd, ed, i.ToString());
                //if (isNull(e1)) e1 = ed;
                //else e2 = ed;
                switch (vals[1])
                {
                    case "смещение":
                        def.SetOffsetWidthExtent(ed, ed.StartVertex, vals[2], ed.StopVertex, vals[3]);
                        break;
                    case "ширина":
                        if (def.DefaultWidthExtentType == PartFeatureExtentEnum.kCenteredWidthExtent) def.RemoveOverrideWidthExtent(ed);
                        def.SetCenteredWidthExtent(ed, vals[2]);
                        break;
                    default:
                        break;
                }
                
            }
            f.Definition = def;
            
            xml.El = p;
            base.def.SetEndOfPartToTopOrBottom(false);
            update = false;
            //u.higlight(doc, e1, e2, I.objs.CreateColor(255, 0, 0), I.objs.CreateColor(0, 255, 0));
        }

        public override void upd()
        {
            //readXML();
        }

        protected override void addDef()
        {
            def = smf.FlangeFeatures.CreateFlangeDefinition(edges, a, d);
            def.ApplyAutoMitering = false;
        }

        public override void get()
        {
            name = xml.getAttributeValue("Name");
            string s = xml.getAttributeValue("Sort");
            if (!isNull(s)) sort = s;
            s = xml.getAttributeValue("SortFaces");
            if (!isNull(s)) fsort = s;
            d = xml.getAttributeValue("d");
            u.addParameter(doc, d, pars);
            if (!getParam(d)) return;
            a = xml.getAttributeValue("a");
            if (a == null) a = "90";
/*            a += " град";*/
            p = xml.getAttributeValue("Parent");
            opt = XMLDoc.getXAttributesValues(xml.El ,new string[] { "Forma", "A", "B", "C", "pos"});
            fil = getSpl("filter");
            find(p);
            if (xml.getAttributeValue("rev") != null && xml.getAttributeValue("rev") != "") rev = true;
            addEdge();
        }

        public override bool addEdge()
        {
            base.addEdge();
            if (pf is CutFeature)
            {
                if (isNull(es)) es = new HashSet<Edge>();
                foreach (Face item in pf.Faces)
                {
                    if (filterCut(item))
                    {
                        u.action<Edge>(item.Edges, a => es.Add(a), fi => !flen(fi));
                    }
                }
            }
            else
            {
                getFaces();
            }
                base.filter();
                base.getEdges();
            return true;
        }

        public void getFaces()
        {
            Face fa = null;
//             if (pf.Type == ObjectTypeEnum.kContourFlangeFeatureObject)
//             {
            List<Face> lfs = new List<Face>();
            foreach (Face item in pf.Faces)
            {
                if (filter(item))
                    {
                        if (isNull(fsort)) { fa = item; break; }
                        else
                        {
                            lfs.Add(item);
                        }
                    }
                }
            if (lfs.Count != 0)
            {
                int snum, sel; int.TryParse(getElem(fsort,0), out snum); int.TryParse(getElem(fsort,1), out sel);
                if (!isNull(snum,sel))
                {
                    fa = lfs.OrderBy(fac => getPoint(fac.Edges[1], snum-1)).ElementAt(sel-1);
                }
                getFace(fa);
                return;
            }
            if (pf.Faces.Count == 1 && pf.Faces[1].SurfaceType == SurfaceTypeEnum.kPlaneSurface)
                fa = pf.Faces[1];
            if (rev) fa = changeSide(fa);
            getFace(fa);
            //}
        }
        public Face changeSide(Face f)
        {
            Face fa;
            Vertex v = f.Vertices[1];
            Edge ed = u.get<Edge>(v.Edges, fi => flen(fi));
            v = getVertex(ed, v);
            fa = u.get<Face>(v.Faces, fi => filter(fi));
            return fa;
        }
        public bool filter(Face f)
        {
            if (f.SurfaceType != SurfaceTypeEnum.kPlaneSurface) return false;
            if (u.get<Edge>(f.Edges, fi => flen(fi)) != null) return false;
            return true;
        }
        public bool filterCut(Face f)
        {
            if (f.SurfaceType != SurfaceTypeEnum.kPlaneSurface) return false;
            if (u.get<Edge>(f.Edges, fi => fi.GeometryType != CurveTypeEnum.kLineSegmentCurve) != null) return false;
            return true;
        }
        public Vertex getVertex(Edge e, Vertex v)
        {
            return e.StartVertex.Equals(v) ? e.StopVertex : e.StartVertex;
        }

        public void getFace(Face f)
        {
            if (fs == null) fs = new HashSet<Face>();
            fs.Add(f);
            u.action<Face>(f.TangentiallyConnectedFaces, a => fs.Add(a), fi => fi.SurfaceType == SurfaceTypeEnum.kPlaneSurface && fi.CreatedByFeature.Equals(pf));
            u.action<Face>(fs, a => getEdge(a));
        }

        public void getEdge(Face fa)
        {
            if (es == null) es = new HashSet<Edge>();
            u.action<Edge>(fa.Edges, a => es.Add(a), fi => tangentFace(fi));
        }
    }

    class CutFeat : FeaturesInv
    {
        CutFeature f;
        string noBend;
        new CutDefinition def;
        public CutFeat()
        {
            t = entTypes.Cut;
            if (!checkAsm()) return;
            alias = "В";
            get();
            noBend = xml.getAttributeValue("noBend");
            ips = getSketch();
            if (checkSketch()) return;
            setName();
            if (!check(name)) return;
            addDef();
        }

        protected override void addDef()
        {
            if (findPlaneBySketch()) addDef();
            Profile p = ips.Profiles.AddForSolid();
            ips.Visible = false;
            def = smf.CutFeatures.CreateCutDefinition(p);
            if (isNull(noBend))
            def.SetCutAcrossBendsExtent(thick.Name);
        }

        public override void draw()
        {
            if (def == null) return;
            f = smf.CutFeatures.Add(def);
            f.Name = fname();
        }
    }

    class FilletFeat : FeaturesInv
    {
        FilletFeature f;
        string r, pfName;
        new FilletDefinition def;
        public FilletFeat()
        {
            t = entTypes.Fillet;
            if (!checkAsm()) return;
            alias = "Скр";
            string me = xml.getAttributeValue("MustExist");
            if (!checkPF(me)) return;
            edges = I.objs.CreateEdgeCollection();
            getParentEl();
            if (pf == null) return;
            pfName = pf.Name;
            setName();
            if (check(name)) return;
            get();
            addDef();
        }

        protected override void addDef()
        {
            if (edges.Count == 0) return;
            def = smf.FilletFeatures.CreateFilletDefinition();
            def.AddConstantRadiusEdgeSet(edges, r);
        }

        protected override bool check(string name)
        {
            PartFeature tpf = u.get<PartFeature>(smf, e => e.Name == name);
            if (isNull(tpf)) return false;
            return true;
        }

        public override bool addEdge()
        {
            base.addEdge();
            getEdges();
            filterEdge();
            base.getEdges();
            return true;
        }
        public void filterEdge()
        {
            var spl = XMLDoc.getSpl(xml.El, "filter", ':');
            if (isNull(spl) || spl.Length != 2) return;
            double min = u.convToDouble(spl[0], 0.1, 3), max = u.convToDouble(spl[1], 0.1, 3);
            es.RemoveWhere(e =>
            {
                foreach (Edge item in e.StartVertex.Edges)
                {
                    if (item.Equals(e)) continue;
                    if (!filter(item, min, max)) return true;
                }
                return false;
            });
        }

        public override void getEdges()
        {
            getFaces(fi => fi.SurfaceType == SurfaceTypeEnum.kPlaneSurface);
            getEdges((fi,i) => u.eq(u.getLenght(fi),thick.ModelValue) && tangent(fi));
        }

        public override void setName()
        {
            name = xml.getAttributeValue("Name", true);
            r = xml.getAttributeValue("r");
            if (isNull(r)) return;
            double dr = getDouble(r, 1, 2);
            if (isNull(name))
                name = alias + dr + pfName;
        }

        public override void draw()
        {
            if (def == null) return;
            f = smf.FilletFeatures.Add(def);
            f.Name = fname();
        }

        public override void get()
        {
            base.get();
            r = xml.getAttributeValue("r");
            addEdge();
        }
    }

    class FaceFeat : FeaturesInv
    {
        FaceFeature f;
        new FaceFeatureDefinition def;
        string prj;
        bool rev = false;
        public FaceFeat()
        {
            t = entTypes.Face;
            if (!checkAsm()) return;
            alias = "Г";
            setAlias();
            prj = xml.getAttributeValue("Projection");
            get();
            if (!isNull(xml.getAttributeValue("rev"))) rev = true;
            ips = getSketch();
            if (checkSketch()) return;
            setName();
            if (!check(name)) return;
            //if (!isNull(u.get<object>(ips.Dependents, f => f is FaceFeature))) return;
            addDef();
        }

        public override string ToString()
        {
            return "";
        }

        public override void draw()
        {
            if (def == null) return;    
            f = smf.FaceFeatures.Add(def);
            f.Name = fname();
        }

        public override void upd()
        {
           
        }

        protected override void addDef()
        {
            if (!isNull(prj) && findPlaneBySketch()) addDef();
            Profile p = ips.Profiles.AddForSolid();
            def = smf.FaceFeatures.CreateFaceFeatureDefinition(p);
            if (ips.ReferencedEntity == null && ips.PlanarEntity is Face)
            {
                def.Direction = PartFeatureExtentDirectionEnum.kNegativeExtentDirection;
            }
            if (rev)
                def.Direction = def.Direction == PartFeatureExtentDirectionEnum.kNegativeExtentDirection ?
                    PartFeatureExtentDirectionEnum.kPositiveExtentDirection: PartFeatureExtentDirectionEnum.kNegativeExtentDirection;
        }
    }

    abstract class FeaturesInv : Entity
    {
        protected SheetMetalComponentDefinition def;
        protected SheetMetalFeatures smf;
        protected EdgeCollection edges;
        protected FaceCollection faces;
        protected ObjectCollection col;
        protected UnitVector normal;
        protected Point spoint;
        protected PlanarSketch nSketch;
        protected PartFeature pf;
        protected IEnumerable<Face> fa;
        protected HashSet<Face> fs;
        protected HashSet<Edge> es;
        protected List<ObjectTypeEnum> obFltr;
        protected bool lst = true;
        static public Parameter thick;
        static public SheetMetalStyles sms;
        protected static int num;
        protected static ComponentOccurrence oc;
        protected string[] spl = null;
        protected string sort = null, fsort = null;
        protected abstract void addDef();
        public FeaturesInv() : base (entTypes.Features)
        {
            if (doc.SubType == "{9C464203-9BAE-11D3-8BAD-0060B0CE6BB4}")
            {
                def = pcd as SheetMetalComponentDefinition;
                smf = def.Features as SheetMetalFeatures;
                thick = def.Thickness;
                sms = def.SheetMetalStyles;    
            }
        }
    
        public override void add()
        {
            exist();
            if (!update)
            {
                draw();
                vis();
            }
            else
                upd();
        }

        protected virtual void setAlias()
        {
            string tp = xml.getAttributeValue("Alias");
            if (tp != null) alias = tp;
        }

        public override void draw()
        {
           
        }

        public override void upd()
        {
           
        }

        public Path getPath(PlanarSketch sk)
        {
            var lins = u.get<SketchLine>(sk.SketchLines, fi => !fi.Construction);
            if (isNull(lins)) return null;
            return smf.CreatePath(lins);
        }

        public virtual bool getParam(params string[] names)
        {
            bool sign;
            foreach (var name in names)
            { 
                u.addParameter(Entity.doc, name, Entity.pars);
                if (isAlfa(name, out sign) && !u.findParameter(doc, name)) return false; 
            }
            return true; 
        }

        public bool find(string name)
        {
            if (isNull(name))
            {
                string l = xml.getAttributeValue("Last");
                if (!isNull(l))
                {
                    pf = lastEntity<PartFeature>(count: l).LastOrDefault();
                    return pf == null ? false : true;
                }
            }
            pf = InvDoc.Reflect.exist<PartFeature>(pcd.Features, fi => fi.Name == name);
            return pf == null ? false: true;
        }

        public bool findFaces(string name)
        {
            if (isNull(name))
            {
                string l = xml.getAttributeValue("Last");
                if (!isNull(l))
                {
                    fa = refEntity(u.gets<Face>(pcd.SurfaceBodies[1].Faces, fi => fi.CreatedByFeature.Type == ObjectTypeEnum.kReferenceFeatureObject), l);
                    var gr = fa.GroupBy(g => ((Face)g.ReferencedEntity).CreatedByFeature.Name);
                    int i = XMLDoc.getIntElem(l,0)-1;
                    if (gr.Count() >= i) fa = gr.ElementAt(i);
                    return fa.Count() == 0 ? false : true;
                }
            }
            fa = refEntity(pcd.SurfaceBodies[1].Faces, name);
            return fa.Count() == 0 ? false : true;
        }

        public string lplus()
        {
            string l = xml.getAttributeValue("Last");
            return l;
        }

        protected void getCustom()
        {
            string lst1 = xml.getAttributeValue("No");
            lst = isNull(lst1) ? true : false;
        }                       

        public bool findPlaneBySketch(string d = null)
        {
            if (isNull(ips)) return false;
            if (!ips.HasReferenceComponent && ips.PlanarEntity is Face) return false;
            normal = ips.PlanarEntityGeometry.Normal;
            spoint = ips.SketchPoints[1].Geometry3d;
            object ob;
            if (isNull(d)) 
            {
                double tol = 0.1;
                //if (t == entTypes.Punch) tol = 1;
                Face f1 = null, f2 = null;
                foreach (SketchPoint item in ips.SketchPoints)
                {
                    spoint = item.Geometry3d;
                    if (findFace(ref f1, ref f2, tol)) break;
                }
                if (f1 == null && f2 == null) return false;
                double d1 = distToFace(f1, spoint), d2 = distToFace(f2, spoint);
                f1 = d1 == d2 ? f1 : d1 < d2 ? f1: f2;
                ob = f1;
            }
            else
            {
                WorkPlane twp = pcd.WorkPlanes.AddByPlaneAndOffset(ips.PlanarEntity, d);
                twp.Visible = false;
                ob = twp;
            }
            return addSketch(ob);
        }

        public bool findFace(ref Face f1, ref Face f2, double tol)
        {
            f1 = findFace(tol);
            Vector v = normal.AsVector();
            v.ScaleBy(-1);
            normal = v.AsUnitVector();
            f2 = findFace(tol);
            if (f1 == null && f2 == null) return false;
            return true;
        }

        public Face findFace(double tol = 0.1)
        {
            ObjectsEnumerator founds, locs;
            def.FindUsingRay(spoint, normal, tol, out founds, out locs, true);
            if (founds.Count != 0 && founds[1] is Face) return founds[1] as Face;
            foreach (var item in founds)
            {
                Edge ed = item as Edge;
                if (!isNull(ed))
                return u.get<Face>(ed.Faces, fi =>
                {
                    Plane pl = fi.Geometry as Plane; if (isNull(pl)) return false;
                    if (pl.Normal.IsParallelTo(normal)) return true;
                    return false;
                });
                Vertex v = item as Vertex;
                if (!isNull(v))
                    return u.get<Face>(v.Faces, fi =>
                        {
                            Plane pl = fi.Geometry as Plane; if (isNull(pl)) return false;
                            if (pl.Normal.IsParallelTo(normal)) return true;
                            return false;
                        });
            }
            return null;
        }

        public double distToFace(Face f, Point pt)
        {
            if (f == null) return 100000;
            return Math.Round(I.app.MeasureTools.GetMinimumDistance(f, pt), 3);
        }

        protected virtual bool check(string name)
        {
            PartFeature tpf = u.get<PartFeature>(smf, e => e.Name == name);
            return isNull(tpf);
        }

        public override IEnumerable<PartFeature> lastEntity<PartFeature>(System.Collections.IEnumerable col = null, string count = "1", bool last = true)
        {
            return base.lastEntity<PartFeature>(pcd.Features, count, last);
        }

        public bool addSketch(object f)
        {
            nSketch = def.Sketches.Add(f);
            nSketch.Name = "_" + ips.Name;
            nSketch.Visible = false;
            if (col == null)
            {
                col = I.COC(col);
                if (t == entTypes.Cut || t == entTypes.Face || t == entTypes.Loft)
                {
                    u.action<SketchEntity>(ips.SketchEntities, a => col.Add(a));
                }
                else if (t == entTypes.Punch)
                {
                    u.action<SketchPoint>(ips.SketchPoints, a => col.Add(a), fi => fi.HoleCenter);
                }
                else if (t == entTypes.iFeature)
                {
                    u.action<SketchPoint>(ips.SketchPoints, a => col.Add(a), fi => fi.HoleCenter);
                    u.action<SketchLine>(ips.SketchLines, a => col.Add(a), fi => fi.Centerline);
                }
            }
            if (col.Count > 0)
            {
                for (int i = 0; i < col.Count; i++)
                {
                    nSketch.AddByProjectingEntity(col[i + 1]);
                }
                ips.Visible = false;
                ips = nSketch;
                return true;
            }
            return false;
        }

        public void findOcc(string[] els, ref ObjectCollection col, string reg)
        {
            //var els = getSpl(name);
            if (isNull(els)) return;
            foreach (var item in els)
            {
                string n = getElem(item, 0), n1 = getElem(item, 1);
                if (isNull(n, n1)) continue;
                int.TryParse(n1, out num);
                if (num == 0) continue;
                oc = null;
                findPCD(acd.Occurrences, n, reg);
                if (!isNull(oc)) col.Add(oc);
            }
        }

        public PartFeature findSpl(string[] els, ref ObjectCollection col)
        {
            if (isNull(els))
            {
                string l = xml.getAttributeValue("Last");
                if (isNull(l)) return null;
                bool first = true; PartFeature tpf = null;
                foreach (var item in lastEntity<PartFeature>(count: l, last: lst))
                {
                    PartFeature f = item;
                    if (first) tpf = f;
                    first = false;
                    findSpl(f, ref col);
                }
                return tpf;
            }
            foreach (var item in els)
	        {   
                PartFeature f = InvDoc.Reflect.exist<PartFeature>(smf, fi => fi.Name == item);
                findSpl(f, ref col);
	        }
            return null;
        }

        public void findSpl(PartFeature f, ref ObjectCollection col)
        {
            if (f != null)
            {
                col.Add(f);
                if (f is RectangularPatternFeature)
                {
                    RectangularPatternFeature r = f as RectangularPatternFeature;
                    foreach (var tpf in r.ParentFeatures)
                    {
                        col.Add(tpf);
                    }
                }
            }
        }

        public virtual void getFaces(Func<Face,bool> f)
        {
            if (fs == null) fs = new HashSet<Face>();
            IEnumerable<Face> faces = null;
            if (!isNull(pf))
            faces = pf.Faces.OfType<Face>();
            else if (!isNull(fa) && fa.Count() != 0)
            {
                faces = fa;
            }
            if (isNull(faces))  { br = true; return; }
            foreach (Face item in faces)
            {
                if (f(item)) fs.Add(item);  
            }
        }
        public virtual HashSet<T> filter<T>(HashSet<T> els) where T: class
        {
            if (spl == null) return els;
            HashSet<T> tmp = new HashSet<T>();
            string tstr = String.Join(";", spl);
            IEnumerable<Face> tfs = null;
            string s = u.getElem(tstr, 1);
            if (!isNull(s)) tfs = u.sort<Face>(els as IEnumerable<Face>, s);
//             u.clientTxt(tfs, pcd);
//             I.app.ActiveView.Update();
            if (u.isNull(tfs)) tfs = els as IEnumerable<Face>;

            tfs = u.gets<Face>(tfs, tstr, ';');
            
            
            foreach (var item in tfs)
            {
                tmp.Add(item as T);
            }
//             List<string> num = spl.ToList();
//             int i = 0;
//             foreach (var item in els)
//             {
//                 i++;
//                 if (num.Contains(i.ToString())) tmp.Add(item); 
//             }
            return tmp;
        }
        public void getEdges(Func<Edge, int, bool> f)
        {
            if (es == null) es = new HashSet<Edge>();
            foreach (Face item in fs)
            {
                int i = 0;
                foreach (Edge e in item.Edges)
                {
                    i++;
                    if (f(e, i)) es.Add(e);
                }
            }
        }

        public override void get()
        {
            base.name = xml.getAttributeValue("Sketch");
            setAlias();
        }
        public virtual string fname()
        {
            setName();
            while (find(name))
            {
                name = CreateComponent.nextName(name); 
            }
            return name;
        }
        public virtual void exist()
        {
            if (update) return;
            name = xml.getAttributeValue("Name");
            if (isNull(pcd)) return;
            pf = InvDoc.Reflect.exist<PartFeature>(pcd.Features, fi => fi.Name == name);
            if (pf != null) update = true;
            else update = false;
        }
        public virtual bool addEdge()
        {
            spl = getSpl("num");
            if (isNull(spl)) return false;
            return true;
        }
        public string[] getSpl(string name)
        {
            var atts = xml.getXAttributesValues(new string[] { name });
            if (atts[0] == null) return null;
            spl = atts[0].Split(';');
            if (spl == null) return null;
            return spl;
        }
        public virtual void getEdges()
        {
            edges = I.CEC(edges);
            if (isNull(spl)) addEdges(fi => true);
            else
            {
                List<string> num = spl.ToList();
                addEdges(fi => num.Contains(fi.ToString()));
            }
        }
        public virtual void getNumEdges()
        {
            edges = I.CEC(edges);
            if (isNull(spl)) addEdges(fi => true);
            else
            {
                foreach (var item in spl)
	            {
		           int i = u.convToInt(item);
                   edges.Add(es.ElementAt(i-1));
	            }
            }
        }
        public IOrderedEnumerable<Edge> sortEdges()
        {
            if (es.Count < 2) return null;
            int snum; int.TryParse(sort, out snum);
            if (snum == 0) return null;
            return es.OrderBy(e => getPoint(e, snum-1));
        }
        public double getPoint(Edge e, int n)
        {
            double [] d = new double []{};
            e.StartVertex.GetPoint(ref d);
            return d[n];
        }
        protected virtual void addEdges(Func<int, bool> f)
        {
            int i = 0;
            obFltr = getFilterEdges();
            filterEdges();
            if (!isNull(sort))
            {
                foreach (var item in sortEdges())
                {
                    i++;
                    if (f(i)) edges.Add(item);
                }
                return;
            }

            foreach (var item in es)
            {
                i++;
                if (f(i)) edges.Add(item);
            }
        }
//         protected virtual void addEdges(string f)
//         {
//             int i = 0;
//             obFltr = getFilterEdges();
//             filterEdges();
//             var eds = u.gets<Edge>(es, f, ':');
//             if (!isNull(sort))
//             {
//                 foreach (var item in sortEdges())
//                 {
//                     i++;
//                     if (f(i)) edges.Add(item);
//                 }
//                 return;
//             }
// 
//             foreach (var item in es)
//             {
//                 i++;
//                 if (f(i)) edges.Add(item);
//             }
//         }
        protected List<ObjectTypeEnum> getFilterEdges()
        {
            var spl = XMLDoc.getSpl(xml.El ,"EdgeFltr", ';');
            if (isNull(spl)) return null;
            List<ObjectTypeEnum> lst = new List<ObjectTypeEnum>();
            foreach (var item in spl)
            {
                lst.AddRange(getObType("1:" + item));
            }
            if (lst.Count == 0) return null;
            return lst;
        }
        protected bool filterEdge(Edge e)
        {
            if (isNull(obFltr)) return true;
            foreach (Face item in e.Faces)
            {
                Face f = item;
                /*if (f.TangentiallyConnectedFaces.Count != 0) f = f.TangentiallyConnectedFaces[2] as Face;*/
                if (!obFltr.Contains(f.CreatedByFeature.Type)) return false; 
            }
            return true;
        }
        protected void filterEdges()
        {
            HashSet<Edge> rem = new HashSet<Edge>();
            foreach (Edge item in es)
            {
                if (!filterEdge(item)) rem.Add(item);
            }
            es.RemoveWhere(e => rem.Contains(e));
        }
        protected virtual void getParentEl()
        {
            string n = xml.getAttributeValue("Parent");
            if (n == null)
            {
                XElement el = xml.El.Parent;
                if (el != null && el.Name != "head")
                {
                    n = XMLDoc.getAttributeValue(el, "Name");

                }
            }
            find(n);
        }
        public bool tangent(Edge ed)
        {
            foreach (Face item in ed.StartVertex.Faces)
            {
                if (item.SurfaceType != SurfaceTypeEnum.kPlaneSurface) return false;
            }
            return true;
        }
        public bool tangentFace(Edge ed)
        {
            foreach (Face item in ed.Faces)
            {
                if (item.SurfaceType != SurfaceTypeEnum.kPlaneSurface) return false;
            }
            return true;
        }
        public bool flen(Edge ed)
        {
            if (ed.GeometryType != CurveTypeEnum.kLineSegmentCurve) return false;
            return u.eq(u.getLenght(ed), thick.ModelValue);
        }
        public virtual void filter()
        {
            var spit = getSpl("filter");
            if (spit == null) return;
            double min = u.convToDouble(spit[0], 0.1, 3), max = u.convToDouble(spit[1], 0.1, 3);
            es.RemoveWhere(el => !filter(el, min, max));
            spl = getSpl("num");
        }
        public virtual bool filter(Edge ed, double min, double max)
        {
            double l = u.getLenght(ed);
            return u.filter(l, min, max);
        }
        public bool filterR(Face f, double min, double max)
        {
            Cylinder cyl = f.Geometry as Cylinder;
            if (isNull(cyl)) return false;
            return u.filter(cyl.Radius, min, max);
        }
        public UnitVector getNorm(ref string typ)
        {
            if (isNull(pf)) return null;
            FaceFeature ff = pf as FaceFeature;
            typ = "ff";
            PlanarSketch tps;
            if (isNull(ff))
            {
                ContourFlangeFeature cf = pf as ContourFlangeFeature;
                if (isNull(cf)) return null;
                tps = (PlanarSketch)(cf.Definition.Path[1].SketchEntity as SketchEntity).Parent;
                typ = "cf";
            }
            else tps = (PlanarSketch)ff.Definition.Profile.Parent;
            if (isNull(tps)) return null;
            return tps.PlanarEntityGeometry.Normal;
        }
        public bool eq(string s1, string s2, string reg)
        {
            if (isNull(reg)) return s1 == s2;
            Regex r = new Regex(s2);
            return r.IsMatch(s1);
        }
        public void findPCD(System.Collections.IEnumerable occs, string name, string reg)
        {
            foreach (ComponentOccurrence item in occs)
            {
                if (item.SubOccurrences.Count != 0) findPCD(item.SubOccurrences, name, reg);
                Document doc = item.ReferencedDocumentDescriptor.ReferencedDocument as Document;
                if (eq(u.getPropValue(doc, "Description"),name, reg))
                {
                    PartDocument pdoc = doc as PartDocument;
                    if (isNull(pdoc)) continue;
                    num--;
                    if (num == 0)
                    {
                        pcd = pdoc.ComponentDefinition;
                        oc = item;
                        break;
                    }
                }
            }
        }
        public iMateDefinition findImatesDef(System.Collections.IEnumerable occs, string name, string reg, int n, int num)
        {
            foreach (ComponentOccurrence item in occs)
            {
                Document doc = item.ReferencedDocumentDescriptor.ReferencedDocument as Document;
                if (eq(u.getPropValue(doc, "Description"), name, reg))
                {
                    n--;
                    if (n == 0)
                    {
                        AssemblyComponentDefinition a = I.getACD(doc);
                        iMateDefinitions defs = null;
                        if (!isNull(a)) defs = a.iMateDefinitions;
                        PartComponentDefinition p = I.getPCD(doc);
                        if (!isNull(p)) defs = p.iMateDefinitions;
                        if (isNull(defs)) return null;
                        oc = item;
                        return defs[num];
                    }
                }
            }
            return null;
        }
    }
    public class Drws
    {
        List<Drw> drws = new List<Drw>();
        public Drws()
        {

        }
    }
    public class Drw
    {
        protected DrawingDocument drw;
        Document doc;
        string tmpl = "";
        List<DrwView> views = new List<DrwView>();
        public Drw(string ffn)
        {

        }
        public void add()
        {
            drw = I.app.Documents.Add(DocumentTypeEnum.kDrawingDocumentObject, tmpl, false) as DrawingDocument;
        }
    }
    public class DrwView : Drw
    {
        List<DrwView> views = new List<DrwView>();
        DrawingView dv;
        double x, y;
        double scale, offsetX, offsetY;
        string strScale;
        public DrwView(string path) : base(path)
        {

        }
        public void add()
        {
           // dv = drw.Drawin
        }
    }
}
