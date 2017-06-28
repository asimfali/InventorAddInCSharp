using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using u = InvDoc.u;
using System.Xml.Linq;
using Inventor;
using InvDoc;

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
        public Elements(Document doc)
        {
            app = Macros.StandardAddInServer.m_inventorApplication;
            if (doc.DocumentType == DocumentTypeEnum.kPartDocumentObject)
            {
                Entity.doc = doc as PartDocument;
                ss = Entity.doc.SelectSet;
                if (ss.Count != 1) return;
                //cen = ss[1] as SketchPoint;
                //colect = app.TransientObjects.CreateObjectCollection();
                //add(u.pathUtil(doc) + "\\AC700-53A.00.006 (Люк автомата).ipt");
                //to = cen.Parent as PlanarSketch;
                ps = ss[1] as PlanarSketch;
                if (ps.SketchEntities.Count < 3)
                    readXML(@"c:\WORK\Завесы\Модули\asm.xml");
                else
                    createXML();

            }
        }
        public PlanarSketch addSketch()
        {
            return (ss[1] as SketchPoint).Parent as PlanarSketch;
        }
        public void readXML(string fn)
        {
            Entity.xml = new XMLDoc(fn, "head");
            Entity.xml.insert();
            string n = "Sketch";
            foreach (var el in Entity.xml.find(n))
            {
                Entity.xml.El = el;
                SketchInv sk = new SketchInv();
                sk.add();
            }
        }
        public void createXML()
        {
            string p = u.pathDoc(Entity.doc as Document), name = ps.Name + ".xml";
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
        Point, Line, Arc, Sketch, Hole, Flange, Face, Part, Asm, Features, CountourFlange, Parameter 
    }
    enum cType
    {
        n = 0,
        t = 0x01, // Tangent
        c = 0x02, // Close
        hole = 0x04, // HoleCenter
    }

    abstract class Entity
    {
        public entTypes t;
        public cType con;
        static protected bool update = false;
        static public bool closed = false;
        protected double x, y;
        protected string name = "", form = "0.###"/*, con*/;
        public SketchEntity ent/*, prev*/;
        public Entity prev;
        protected SketchPoint sp;
        static protected PlanarSketch ps;
        public SketchPoint startPt, endPt;
        protected SketchLine sl;
        public static PartDocument doc;
        public static XMLDoc xml;
        public bool first = false, last = false;

        public abstract void add();
        public abstract void draw();
        public abstract void upd();
        public abstract void get();
        public Entity(entTypes t, double x = 0, double y = 0)
        {
            this.x = x; this.y = y; this.t = t;
        }
        public void Con()
        {
            string tmp = xml.getAttributeValue("Constr").Replace(' ',',');
            Enum.TryParse(tmp, out con);
        }
        public string conToString()
        {
            return con.ToString().Replace(",","");
        }
        public bool check<T>(T ob)
        {
            return (ob != null) ? true : false;
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
        Vector2d dir;
        double[] Ls;
        double offset;

        public PointInv():    
            base(entTypes.Point)
        {
            string s = xml.getAttributeValue("Hole");
            hole = s == "" || s == "0" ? false: true;
            xml.getCoord();
            x = xml.x; y = xml.y;
            if (xml.getCoord("Dir"))
            {
                dir = u.createVector2d(xml.x, xml.y);
            }
            if (xml.getLen()) Ls = xml.Ls;
        }
        public PointInv(SketchPoint sp): base(entTypes.Point)
        {
            base.sp = sp; startPt = sp; endPt = sp; ent = (SketchEntity)sp; 
        }

        public override void add()
        {
            if (!update)
                draw();
            else
                upd();
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
            throw new NotImplementedException();
        }
    }
    class LineInv : Entity
    {
        Vector2d dir; double L;
        public LineInv()
            : base(entTypes.Line)
        {
            Con();
            if (con.HasFlag(cType.c)) last = true;
            if (xml.getCoord("Dir")) dir = u.createVector2d(xml.x, xml.y);
            L = u.convToDouble(xml.getAttributeValue("L"),0.1,3);  
            name = u.convToModelName(xml.getAttributeValue("Name", true));
            if (xml.getCoord())
            {
                first = true;
                sp = ps.SketchPoints.Add(u.createPoint2d(xml.x, xml.y));
                sp.HoleCenter = false;
                prev = new PointInv(sp);
            }
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
            if (f) { x = pt.X; y = pt.Y; first = true; }
        }
        public override void add()
        {
            if (!update)
                draw();
            else
                upd();
        }
        public double scalar(SketchLine sl,SketchLine prev)
        {
            return scalar(sl.Geometry.Direction, prev.Geometry.Direction); 
        }
        public double scalar(UnitVector2d v1, UnitVector2d v2)
        {
            return v1.X * v2.X + v1.Y * v2.Y;
        }
        public double scalarX(SketchLine sl)
        {
            return scalar(sl.Geometry.Direction, u.createUnitVector2d(0, 1));
        }
        public double scalarY(SketchLine sl)
        {
            return scalar(sl.Geometry.Direction, u.createUnitVector2d(1, 0));
        }
        public void constr(Entity prev)
        {
            sl = ent as SketchLine;
            if (prev.ent.Type == ObjectTypeEnum.kSketchLineObject)
            {
                if (scalar(sl, prev.ent as SketchLine) == 0)
                {
                    constraints.addPerp(ent, prev.ent);
                }
                else
                {
                    constraints.addTwoLineAngle(prev.ent as SketchLine, ent as SketchLine, name: name, w: 0.5);
                }
            }
            else if (prev.ent.Type == ObjectTypeEnum.kSketchArcObject)
            {
                if (con.HasFlag(cType.t)) constraints.addTangent(prev.ent, ent);
            }
        }
        public void dimConstr()
        {
            sl = ent as SketchLine;
            constraints.addTwoPointDist(sl, name, scale:-0.2);
        }
        public void direct(Entity prev)
        {
            if (prev.ent.Type != ObjectTypeEnum.kSketchArcObject) return;
            Arc2d arc = (prev.ent as SketchArc).Geometry;
            Curve2dEvaluator ev = arc.Evaluator;
            double[] par = u.getParam(ev, prev.endPt.Geometry);
            double[] o = { };
            ev.GetTangent(ref par, ref o);
            dir = u.createVector2d(o[0], o[1]);
            if (prev.t == entTypes.Arc && !(prev as ArcInv).clockwise)
            {
                dir.ScaleBy(-1);
            }
        }

        public override void draw()
        {
            if (prev == null) return;
            sp = prev.endPt;
            if (con.HasFlag(cType.t)) direct(prev);
            dir.Normalize();
            if (L != 0)
                dir.ScaleBy(L);
            Point2d pt = sp.Geometry; pt.TranslateBy(dir);
            sl = ps.SketchLines.AddByTwoPoints(sp, pt);
            startPt = sl.StartSketchPoint; endPt = sl.EndSketchPoint;
            if (con.HasFlag(cType.hole))
            {
                startPt.HoleCenter = true; endPt.HoleCenter = true;
                sl.Construction = true;
            }
            ent = (SketchEntity)sl;
            if (L != 0)
                dimConstr();
            constr(prev);
        }

        public override void upd()
        {
            Parameter p = u.findInCol<Parameter>(doc.ComponentDefinition.Parameters, e => e.Name == name);
            if (p != null && !u.eq(p.Value, L)) p.Value = L; 
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

        public override void get()
        {
            throw new NotImplementedException();
        }
    }
    class ArcInv : Entity
    {
        double R, ang;
        Point2d cen;
        public bool clockwise = true;
        public ArcInv()
            : base(entTypes.Arc)
        {
            R = u.convToDouble(xml.getAttributeValue("R"), 0.1, 3);
            ang = u.convToDouble(xml.getAttributeValue("Ang"))*Math.PI/180;
            if (xml.getCoord())
            {
                x = xml.x; y = xml.y;
                cen = u.createPoint2d(x, y);
            }
            Con();
            if (con.HasFlag(cType.c)) last = true;
            name = u.convToModelName(xml.getAttributeValue("Name", true));
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

        public override void add()
        {
            if (!update)
                draw();
            else
                upd();
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
            SketchLine sl = ps.SketchLines.AddByTwoPoints(arc.EndSketchPoint, arc.CenterSketchPoint); sl.Construction = true;
            SketchLine sl1 = ps.SketchLines.AddByTwoPoints(arc.StartSketchPoint, arc.CenterSketchPoint); sl1.Construction = true;
            RadiusDimConstraint r = constraints.addRadius(ent, sl.StartSketchPoint, sl1.StartSketchPoint, name);
            constraints.addTwoLineAngle(sl, sl1, arc.CenterSketchPoint, name);
            r.TextPoint = u.midPt(sl.StartSketchPoint.Geometry, sl.EndSketchPoint.Geometry, 0.4);
        }

        public override void draw()
        {
            sp = null; UnitVector2d v = null; double sAng = 0;
            if (prev.ent.Type == ObjectTypeEnum.kSketchPointObject)
                sp = prev.ent as SketchPoint;
            else if (prev.ent.Type == ObjectTypeEnum.kSketchLineObject)
            {
                sl = prev.ent as SketchLine;
                setCenter();
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
            Parameter p = u.findInCol<Parameter>(doc.ComponentDefinition.Parameters, e => e.Name == name+"R");
            if (p != null && !u.eq(p.Value, R)) p.Value = R;
            p = u.findInCol<Parameter>(doc.ComponentDefinition.Parameters, e => e.Name == name + "A");
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
            throw new NotImplementedException();
        }
    }
    static class constraints
    {
        public static PlanarSketch ps;
        static object val;
        static Point2d mp;
        static public void addName<T>(string name, string suff = "") where T: class 
        {
            if (name != "")
            {
                name += suff;
                Parameter p = InvDoc.Reflect.getProp<T,Parameter>(val as T, "Parameter");
                if (p.Name == name) name += "1";
                p.Name = name;
            }
        }
        static public RadiusDimConstraint addRadius(SketchEntity ent, SketchPoint sp1, SketchPoint sp2, string name)
        {
            mp = u.midPt(sp1.Geometry, sp2.Geometry);
            val = ps.DimensionConstraints.AddRadius(ent, mp);
            addName<RadiusDimConstraint>(name,"R");
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
            val = ps.DimensionConstraints.AddTwoLineAngle(sl1, sl2, mp);
            addName<TwoLineAngleDimConstraint>(name,"A");
            return val as TwoLineAngleDimConstraint;
        }
        static public TwoPointDistanceDimConstraint addTwoPointDist(SketchLine sl, string name, double offset = 0, double scale = 1)
        {
            mp = u.midPt(sl.StartSketchPoint.Geometry, sl.EndSketchPoint.Geometry, offset, 0, scale);
            val = ps.DimensionConstraints.AddTwoPointDistance(sl.StartSketchPoint, sl.EndSketchPoint, DimensionOrientationEnum.kAlignedDim, mp);
            addName<TwoPointDistanceDimConstraint>(name);
            return val as TwoPointDistanceDimConstraint;
        }
        static public TwoPointDistanceDimConstraint addTwoPointDist(SketchPoint sp1, SketchPoint sp2, string name, double offset = 0, double scale = 1)
        {
            mp = u.midPt(sp1.Geometry, sp2.Geometry, offset, 0, scale);
            try
            {
                val = ps.DimensionConstraints.AddTwoPointDistance(sp1, sp2, DimensionOrientationEnum.kAlignedDim, mp);
                addName<TwoPointDistanceDimConstraint>(name);
                return val as TwoPointDistanceDimConstraint;
            }
            catch (System.Exception)
            {
                return null;
            }
            
        }
        static public TangentSketchConstraint addTangent(SketchEntity se1, SketchEntity se2)
        {
            return ps.GeometricConstraints.AddTangent(se1, se2);
        }
        static public PerpendicularConstraint addPerp(SketchEntity se1, SketchEntity se2)
        {
            return ps.GeometricConstraints.AddPerpendicular(se1, se2);
        }
        static public CoincidentConstraint addConsid(SketchEntity se1, SketchEntity se2)
        {
            return ps.GeometricConstraints.AddCoincident(se1, se2);
        }
    }
    class SketchInv : Entity
    {
        WorkPlane pl;
        string namePlane;
        PartComponentDefinition def;
        List<Entity> ents = new List<Entity>();
        List<ParamInv> prms = new List<ParamInv>();

        public SketchInv(): base(entTypes.Sketch)
        {
            base.name = xml.getAttributeValue("Name");
            def = doc.ComponentDefinition;
            ps = u.findInCol<PlanarSketch>(def.Sketches, e => e.Name == name);
            if (ps == null)
            {
                namePlane = xml.getAttributeValue("Plane");
                pl = u.findInCol<WorkPlane>(def.WorkPlanes, e => e.Name == namePlane);
                if (pl != null) ps = def.Sketches.Add(pl);
                ps.Name = name;
                update = false;
            }
            else update = true;
            readXML();
            constraints.ps = ps;
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

        public void readXML()
        {
            foreach (var el in xml.El.Elements())
            {
                xml.El = el;
                entTypes ty = (entTypes)Enum.Parse(typeof(entTypes), el.Name.ToString());
                switch (ty)
                {
                    case entTypes.Point:
                        ents.Add(new PointInv());
                        break;
                    case entTypes.Line:
                        ents.Add(new LineInv());
                        break;
                    case entTypes.Arc:
                        ents.Add(new ArcInv());
                        break;
                    default:
                        break;
                }
            }
        }
        public override void draw()
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
            Entity f = ents.First(e => e.first);
            Entity l = ents.First(e => e.last);
            ps.SketchLines.AddByTwoPoints(f.startPt, l.endPt);
            //constraints.addConsid(f.startPt as SketchEntity, l.endPt as SketchEntity);
        }

        public override void add()
        {
            if (!update)
                draw();
            else
                upd();
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

    class HoleInv : FeaturesInv
    {
        HoleFeature f;
        public HoleInv()
        {
            throw new System.NotImplementedException();
        }

        public override void add()
        {
            throw new NotImplementedException();
        }

        public override void draw()
        {
            throw new NotImplementedException();
        }

        public override void upd()
        {
            throw new NotImplementedException();
        }

        public override void get()
        {
            throw new NotImplementedException();
        }

        protected override void addDef()
        {
            throw new NotImplementedException();
        }
    }

    class PartInv : Entity
    {
        public PartInv() : base (entTypes.Part)
        {
            throw new System.NotImplementedException();
        }
    
        public override void add()
        {
            throw new NotImplementedException();
        }

        public override void draw()
        {
            throw new NotImplementedException();
        }

        public override void upd()
        {
            throw new NotImplementedException();
        }

        public override void get()
        {
            throw new NotImplementedException();
        }

        public override string ToString()
        {
            throw new System.NotImplementedException();
        }
    }

    class AsmInv : Entity
    {
        public AsmInv() : base (entTypes.Part)
        {
            throw new System.NotImplementedException();
        }
    
        public override void add()
        {
            throw new NotImplementedException();
        }

        public override void draw()
        {
            throw new NotImplementedException();
        }

        public override void upd()
        {
            throw new NotImplementedException();
        }

        public override void get()
        {
            throw new NotImplementedException();
        }

        public override string ToString()
        {
            throw new System.NotImplementedException();
        }
    }

    class CFlangeFeat : FeaturesInv
    {
        ContourFlangeFeature f;
        new ContourFlangeDefinition def;
        new PlanarSketch ps;
        Path p;
        Parameter pr = null;
        public CFlangeFeat(PlanarSketch sketch)
        {
            t = entTypes.CountourFlange;
            ps = sketch;
        }

        protected override void addDef()
        {
            p = smf.CreatePath(ps.SketchEntities[1]);
            if (check<Path>(p)) return;
            def = smf.ContourFlangeFeatures.CreateContourFlangeDefinition(p);
            def.SetDistanceExtent(pr.Name, PartFeatureExtentDirectionEnum.kNegativeExtentDirection);
        }

        public override string ToString()
        {
            throw new System.NotImplementedException();
        }

        public override void add()
        {
            f = smf.ContourFlangeFeatures.Add(def);
        }

        public override void upd()
        {
            throw new System.NotImplementedException();
        }
    }

    class FlangeFeat : FeaturesInv
    {
        FlangeFeature f;
        public FlangeFeat()
        {
            t = entTypes.Flange;
        }

        public override string ToString()
        {
            throw new System.NotImplementedException();
        }

        public override void add()
        {
        }

        public override void upd()
        {
            throw new System.NotImplementedException();
        }

        protected override void addDef()
        {
            throw new NotImplementedException();
        }
    }

    class FaceFeat : FeaturesInv
    {
        FaceFeature f;
        new FaceFeatureDefinition def;
        new PlanarSketch ps;
        public FaceFeat()
        {
            t = entTypes.Face;
            addDef();
            f = smf.FaceFeatures.Add(def);
        }

        public override string ToString()
        {
            throw new System.NotImplementedException();
        }

        public override void add()
        {
            throw new System.NotImplementedException();
        }

        public override void upd()
        {
            throw new System.NotImplementedException();
        }

        protected override void addDef()
        {
            Profile p = ps.Profiles.AddForSolid();
            def = smf.FaceFeatures.CreateFaceFeatureDefinition(p);
        }
    }

    abstract class FeaturesInv : Entity
    {
        protected SheetMetalComponentDefinition def;
        protected SheetMetalFeatures smf;
        static public Parameter thick;
        static public SheetMetalStyles sms;
        protected abstract void addDef();
        public FeaturesInv() : base (entTypes.Features)
        {
            if (doc.SubType == "{9C464203-9BAE-11D3-8BAD-0060B0CE6BB4}")
            {
                def = doc.ComponentDefinition as SheetMetalComponentDefinition;
                smf = def.Features as SheetMetalFeatures;
                thick = def.Thickness;
                sms = def.SheetMetalStyles;
            }
        }
    
        public override void add()
        {
            throw new NotImplementedException();
        }

        public override void draw()
        {
            throw new NotImplementedException();
        }

        public override void upd()
        {
            throw new NotImplementedException();
        }

        public override void get()
        {
            throw new NotImplementedException();
        }
    }
}
