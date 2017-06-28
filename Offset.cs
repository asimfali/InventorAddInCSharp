using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Inventor;

namespace InvAddIn
{
    public partial class Offset : Form
    {
        private PartDocument m_PartDoc;
        private AssemblyComponentDefinition m_asmCompDef;
        private AssemblyDocument asmDoc;
        private SheetMetalComponentDefinition smcd;
        private SheetMetalFeatures smf;
        private Inventor.SketchPoint pt1, pt2, pt3;
        bool flag;
        private Transaction tr;
        private Profile profiles;
        private Inventor.Application invApp;
        private TransientGeometry tg;
        public PlanarSketch ps;
        public PlanarSketchProxy psp;
        private ObjectCollection objCol;
        private string name;
        char separator = System.Globalization.CultureInfo.CurrentCulture.NumberFormat.CurrencyDecimalSeparator[0];
        double a, b;

        public Offset(Inventor.Document newDoc)
        {
            if (newDoc.DocumentType == DocumentTypeEnum.kPartDocumentObject) m_PartDoc = (PartDocument)newDoc;
            if (newDoc.DocumentType == DocumentTypeEnum.kAssemblyDocumentObject)
            {
                asmDoc = (AssemblyDocument)newDoc;
                m_PartDoc = (PartDocument)asmDoc.ActivatedObject;
            }
            //addProject();
            //return;
            tg = ((Inventor.Application)m_PartDoc.Parent).TransientGeometry;
            InitializeComponent();
        }

        private Profile selectFromSketch()
        {
            CommandManager cmdMgr = ((Inventor.Application)m_PartDoc.Parent).CommandManager;
            ps = (PlanarSketch)cmdMgr.Pick(SelectionFilterEnum.kAllPlanarEntities, "Выберите эскиз:");
            name = ps.Name;
            return ps.Profiles.AddForSolid();
        }

        private void rect(Point2d pt1, Point2d pt2, Point2d pt3)
        {
            Vector2d vec = pt1.VectorTo(pt3);
            Vector2d vec2 = pt2.VectorTo(pt3);
            if (vec.Length < vec2.Length) vec = vec2;
            Point2d pt;
            vec.Normalize();
            Vector2d normal = vec.Copy();
            Matrix2d mtx = invApp.TransientGeometry.CreateMatrix2d();
            mtx.SetToRotation(3.1415926 / 2, invApp.TransientGeometry.CreatePoint2d(0, 0));
            normal.TransformBy(mtx);
            vec.ScaleBy(b / 2);
            normal.ScaleBy(a / 2);
            pt = mid(pt1, pt2);
            pt.TranslateBy(vec);
            pt3 = pt.Copy();
            pt3.TranslateBy(normal);
            ps.SketchLines.AddAsThreePointCenteredRectangle(mid(pt1, pt2), pt,pt3);
        }

        static public void rect(SketchLine sl, SketchPoint mp, double a, double b)
        {
            TransientGeometry tg = Macros.StandardAddInServer.m_inventorApplication.TransientGeometry;
            PlanarSketch ps = (PlanarSketch)sl.Parent;
            Point2d pt = tg.CreatePoint2d(mp.Geometry.X,mp.Geometry.Y);
            SketchEntitiesEnumerator en = ps.SketchLines.AddAsThreePointRectangle(pt, tg.CreatePoint2d(pt.X + a, pt.Y), tg.CreatePoint2d(pt.X, pt.Y + b));
            SketchLine sl1 = (SketchLine)en[1];
            SketchLine sl2 = (SketchLine)en[2];
            SketchLine sl3;
            TwoPointDistanceDimConstraint tpdc = ps.DimensionConstraints.AddTwoPointDistance(sl1.StartSketchPoint, sl1.EndSketchPoint,DimensionOrientationEnum.kAlignedDim, sl1.Geometry.MidPoint);
            tpdc = ps.DimensionConstraints.AddTwoPointDistance(sl2.StartSketchPoint, sl2.EndSketchPoint, DimensionOrientationEnum.kAlignedDim, sl2.Geometry.MidPoint);
            if ((int)(sl1.Length*1000 ) <= (int)(a*2000))
            {
                sl3 = midleLine((SketchLine)en[1], (SketchLine)en[3]);
            }
            else
            {
                sl3 = midleLine((SketchLine)en[2], (SketchLine)en[4]);
            }
            ps.GeometricConstraints.AddMidpoint(mp, sl3);
            ps.GeometricConstraints.AddParallel((SketchEntity)sl, (SketchEntity)sl3);
        }

        private Point2d mid (Point2d pt1, Point2d pt2)
        {
            return invApp.TransientGeometry.CreatePoint2d((pt1.X + pt2.X) / 2, (pt1.Y + pt2.Y) / 2);
        }

        static public bool cmpLen(ProfilePath pp, double b = 0.6)
        {
            SketchEntity se1 = (SketchEntity)((ProfileEntity)pp[1]).SketchEntity;
            SketchEntity se2 = (SketchEntity)((ProfileEntity)pp[2]).SketchEntity;
            if (ParamLen(se1) == b || ParamLen(se2) == b) return true;
            return false;
        }

        static public double ParamLen(SketchEntity se)
        {
            double min, max, length;
            Curve2dEvaluator eval = null;
            if (se.Type == ObjectTypeEnum.kSketchLineObject || se.Type == ObjectTypeEnum.kSketchLineProxyObject)
                {eval = ((SketchLine)se).Geometry.Evaluator;}
            else if (se.Type == ObjectTypeEnum.kSketchArcObject || se.Type == ObjectTypeEnum.kSketchArcProxyObject)
                {eval = ((SketchArc)se).Geometry.Evaluator;}
            else if (se.Type == ObjectTypeEnum.kSketchSplineObject || se.Type == ObjectTypeEnum.kSketchSplineProxyObject)
                {eval = ((SketchSpline)se).Geometry.Evaluator;}
            eval.GetParamExtents(out min, out max);
            eval.GetLengthAtParam(min, max, out length);
            return Math.Round(length, 2);
        }

        public void addProject()
        {
            try
            {
            m_asmCompDef = asmDoc.ComponentDefinition;
            invApp = (Inventor.Application)asmDoc.Parent;
            tg = invApp.TransientGeometry;
            smcd = (SheetMetalComponentDefinition)m_asmCompDef.Occurrences[1].Definition;
            SheetMetalComponentDefinition smcd2 = (SheetMetalComponentDefinition)m_asmCompDef.Occurrences[2].Definition;
            //m_asmCompDef.Occurrences[1].Edit();
            CommandManager cmdMgr = ((Inventor.Application)asmDoc.Parent).CommandManager;
            FaceProxy face;
            object obj = cmdMgr.Pick(SelectionFilterEnum.kPartFacePlanarFilter, "Выберите плоскость эскиза:");
            objCol = invApp.TransientObjects.CreateObjectCollection();
            face = (FaceProxy)obj;
            asmDoc.SelectSet.Select(m_asmCompDef.Occurrences[2]);
            //((PartDocument)smcd.Document).SelectSet.Select(m_asmCompDef.Occurrences[2].CreateGeometryProxy());
                //double[] pts = {face.PointOnFace.X,face.PointOnFace.Y,face.PointOnFace.Z};
                //double[] normals = {0,0,0};
                //face.Evaluator.GetNormalAtPoint(ref pts, ref normals);
                //Plane pl = invApp.TransientGeometry.CreatePlane(face.PointOnFace, invApp.TransientGeometry.CreateVector(normals[0], normals[1], normals[2]));
                //obj = m_asmCompDef.Occurrences[2].Definition.SurfaceBodies[1];
                //m_PartDoc.SelectSet.Select()
                obj = null;
                WorkPlane wp = smcd2.WorkPlanes[3];
                m_asmCompDef.Occurrences[1].CreateGeometryProxy(face, out obj);
                FaceProxy face2 = (FaceProxy)obj;
                ps = smcd.Sketches.Add(face2);
                ProjectedCut cut = ps.ProjectedCuts.Add();
                PlanarSketch pls = smcd2.Sketches.Add(smcd2.WorkPlanes[3]);
                m_asmCompDef.Occurrences[2].CreateGeometryProxy(pls, out obj);
                psp = (PlanarSketchProxy)obj;
                foreach (var item in ps.SketchLines)
                {
                    m_asmCompDef.Occurrences[1].CreateGeometryProxy(item,out obj);
                    psp.AddByProjectingEntity(obj);
                }
                //m_asmCompDef.Occurrences[2].CreateGeometryProxy(smcd.WorkPlanes[3],out obj);
                //m_asmCompDef.Occurrences[1].CreateGeometryProxy(ps,out obj);

                //psp.ProjectedCuts.Add();
                TransientGeometry tr = invApp.TransientGeometry;
                //cmdMgr.ControlDefinitions["SketchProjectCutEdgesCmd"].Execute();     
            //enumerat = invApp.TransientGeometry.SurfaceSurfaceIntersection(m_asmCompDef.Occurrences[2].Definition.SurfaceBodies[1], pl);
            }
            catch (Exception)
            {

                throw;
            }
        }

        static public void project(SheetMetalComponentDefinition smcd, string name)
        {
            WorkPlane wp = smcd.WorkPlanes["Шип_справа"];
            PlanarSketch ps = smcd.Sketches.Add(wp);
            ps.Name = name; ps.Visible = false; 
            ps.ProjectedCuts.Add();
        }

        private void offset(Profile profiles)
        {
            foreach (ProfilePath pr in profiles)
            {
                flag = cmpLen(pr);
                if (!flag)
                {
                    b = Convert.ToDouble(this.textBox2.Text.Replace(',', separator)) / 10;
                }
                else b = Convert.ToDouble(this.textBox3.Text.Replace(',', separator)) / 10;
                if (pr.Count == 4)
                {
                    ProfileEntity pe1 = (ProfileEntity)pr[1];
                    ProfileEntity pe3 = (ProfileEntity)pr[3];

                    if (pe1.SketchEntity.Type == ObjectTypeEnum.kSketchLineObject || pe1.SketchEntity.Type == ObjectTypeEnum.kSketchLineProxyObject)
                    {
                        ((SketchLine)pe1.SketchEntity).Construction = true;
                        pt1 = ((SketchLine)pe1.SketchEntity).StartSketchPoint;
                        pt3 = ((SketchLine)pe1.SketchEntity).EndSketchPoint;
                    }
                    else if (pe1.SketchEntity.Type == ObjectTypeEnum.kSketchArcObject || pe1.SketchEntity.Type == ObjectTypeEnum.kSketchArcProxyObject)
                    {
                        ((SketchArc)pe1.SketchEntity).Construction = true;
                        pt1 = ((SketchArc)pe1.SketchEntity).StartSketchPoint;
                        pt3 = ((SketchArc)pe1.SketchEntity).EndSketchPoint;
                    }
                    else if (pe1.SketchEntity.Type == ObjectTypeEnum.kSketchSplineObject || pe1.SketchEntity.Type == ObjectTypeEnum.kSketchSplineProxyObject)
                    {
                        ((SketchSpline)pe1.SketchEntity).Construction = true;
                        pt1 = ((SketchSpline)pe1.SketchEntity).StartSketchPoint;
                        pt3 = ((SketchSpline)pe1.SketchEntity).EndSketchPoint;
                    }
                    if (pe3.SketchEntity.Type == ObjectTypeEnum.kSketchLineObject || pe3.SketchEntity.Type == ObjectTypeEnum.kSketchLineProxyObject)
                    {
                        pt2 = ((SketchLine)pe3.SketchEntity).StartSketchPoint;
                        if (pt2.Geometry.VectorTo(pt1.Geometry).Length < pt2.Geometry.VectorTo(pt3.Geometry).Length)   
                            pt2 = ((SketchLine)pe3.SketchEntity).EndSketchPoint;
                        //pt3 = ((SketchLine)pe3.SketchEntity).Geometry.EndPoint;
                    }
                    else if (pe3.SketchEntity.Type == ObjectTypeEnum.kSketchArcObject || pe3.SketchEntity.Type == ObjectTypeEnum.kSketchArcProxyObject)
                    {
                        pt2 = ((SketchArc)pe3.SketchEntity).StartSketchPoint;
                        if (pt2.Geometry.VectorTo(pt1.Geometry).Length < pt2.Geometry.VectorTo(pt3.Geometry).Length)
                            pt2 = ((SketchArc)pe3.SketchEntity).EndSketchPoint;
                        //pt3 = ((SketchArc)pe3.SketchEntity).Geometry.EndPoint;
                    }
                    else if (pe3.SketchEntity.Type == ObjectTypeEnum.kSketchSplineObject || pe3.SketchEntity.Type == ObjectTypeEnum.kSketchSplineProxyObject)
                    {
                        pt2 = ((SketchSpline)pe3.SketchEntity).StartSketchPoint;
                        if (pt2.Geometry.VectorTo(pt1.Geometry).Length < pt2.Geometry.VectorTo(pt3.Geometry).Length)
                            pt2 = ((SketchSpline)pe3.SketchEntity).EndSketchPoint;
                        //pt3 = ((SketchSpline)pe3.SketchEntity).EndSketchPoint.Geometry;
                    }
                    //ps.Edit();
                    rect(pt1.Geometry, pt2.Geometry, pt3.Geometry);
                    //ps.ExitEdit();
                }
            }

            smf = (SheetMetalFeatures)smcd.Features;
            ps = smcd.Sketches[name];
            CutDefinition cd = smf.CutFeatures.CreateCutDefinition(ps.Profiles.AddForSolid());
            try
            { cd.SetCutAcrossBendsExtent("Толщина"); }
            catch (Exception)
            {
                cd.SetCutAcrossBendsExtent("Thickness");
            }
            smf.CutFeatures.Add(cd);
        }

        static public void offsetAdaptive(SheetMetalComponentDefinition smcd, string name, double valA, double valB, double valB1)
        {
            double a = valA, b = valB;
            PlanarSketch ps = smcd.Sketches[name];
            PlanarSketch newSketch = smcd.Sketches.Add(ps.PlanarEntity);
            name = newSketch.Name;
            Profile profiles = ps.Profiles.AddForSolid();
            foreach (ProfilePath pr in profiles)
            {
                bool flag = cmpLen(pr, valB1-0.03);
                if (flag)
                b = valB1;/*Convert.ToDouble(this.textBox3.Text.Replace(',', separator)) / 10;*/
                if (pr.Count == 4)
                {
                    ProfileEntity pe1 = (ProfileEntity)pr[1];
                    ProfileEntity pe3 = (ProfileEntity)pr[3];
                    if (pe1.CurveType != Curve2dTypeEnum.kLineSegmentCurve2d || pe1.EndSketchPoint.Geometry.DistanceTo(pe1.StartSketchPoint.Geometry) > valA*2)
                    {
                        pe1 = (ProfileEntity)pr[2]; pe3 = (ProfileEntity)pr[4];
                    }
                    //pe1.SketchEntity.Construction = true;

                    //ProfileEntity pe1 = pr.OfType<ProfileEntity>().FirstOrDefault(e1 => e1.CurveType == Curve2dTypeEnum.kLineCurve2d &&
                    //    (e1.EndSketchPoint.Geometry.DistanceTo(e1.StartSketchPoint.Geometry) <= 0.1));
                    //ProfileEntity pe3 = pr.OfType<ProfileEntity>().LastOrDefault(e1 => e1.CurveType == Curve2dTypeEnum.kLineCurve2d &&
                    //    (e1.EndSketchPoint.Geometry.DistanceTo(e1.StartSketchPoint.Geometry) <= 0.1));
                    if (pe1 == null || pe3 == null) continue;


                    //SketchLine sl1 = (SketchLine)pe1.SketchEntity;
                    //SketchLine sl2 = (SketchLine)pe3.SketchEntity;

                    SketchLine sl1 = (SketchLine)newSketch.AddByProjectingEntity(pe1.SketchEntity);
                    SketchLine sl2 = (SketchLine)newSketch.AddByProjectingEntity(pe3.SketchEntity);

                    //sl2.Construction = true;
                    SketchLine sl3 = midleLine(sl1, sl2);
                    SketchPoint mp = newSketch.SketchPoints.Add(sl3.Geometry.MidPoint, false);
                    MidpointConstraint mpc = newSketch.GeometricConstraints.AddMidpoint(mp, sl3);
                    //SketchBlock sb;
                    //SketchLine newLin;
                    //if (b == valB)
                    //{
                    //    sb = ps.SketchBlocks.AddByDefinition(smcd.SketchBlockDefinitions["Паз"], mp.Geometry);
                    //}
                    //else
                    //{
                    //    sb = ps.SketchBlocks.AddByDefinition(smcd.SketchBlockDefinitions["Паз_замыкающий"], mp.Geometry);
                    //}
                    //    newLin = ps.SketchLines.OfType<SketchLine>().LastOrDefault(l => l.Construction == true);
                    //    ps.GeometricConstraints.AddCollinear((SketchEntity)newLin, (SketchEntity)sl3);

                    rect(sl3, mp, a, b);       
                }
            }
            //foreach (SketchBlock item in ps.SketchBlocks)
            //{
            //    item.Explode();
            //}
            //foreach (DimensionConstraint item in newSketch.DimensionConstraints)
            //{
            //    item.Delete();
            //}
            //addEqualLength(newSketch, valA);
            //addEqualLength(newSketch, valB);
            //addEqualLength(newSketch, valB1);

            CutFeature cut = addCut(smcd, name); cut.Name = "Пазы";
        }

        static public void addCutDef(SheetMetalComponentDefinition smcd, Double a, Double b, string name = "Паз")
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
            TransientGeometry tg = Macros.StandardAddInServer.m_inventorApplication.TransientGeometry;
            Point2d pt = tg.CreatePoint2d();
            SketchPoint origin = def.SketchPoints.Add(pt, false);
            def.GeometricConstraints.AddGround((SketchEntity)origin);
            SketchLine sl1, sl2, sl3;
            SketchEntitiesEnumerator en = def.SketchLines.AddAsTwoPointRectangle(tg.CreatePoint2d(-b / 2, -a / 2), tg.CreatePoint2d(b / 2, a / 2));
            sl1 = (SketchLine)en[2]; sl2 = (SketchLine)en[4];
            sl3 = midleLine(sl1, sl2);

            //foreach (DimensionConstraint item in def.DimensionConstraints)
            //{
            //    item.Delete();
            //}
        }

        static public void addEqualLength(PlanarSketch ps,double val)
        {
            SketchLine newLin = ps.SketchLines.OfType<SketchLine>().FirstOrDefault(l => (int)(l.Length * 1000) == (int)(val * 1000));
            ps.DimensionConstraints.AddTwoPointDistance(newLin.StartSketchPoint, newLin.EndSketchPoint, DimensionOrientationEnum.kAlignedDim, newLin.Geometry.MidPoint);
            foreach (SketchLine item in ps.SketchLines.OfType<SketchLine>().Where(l => (int)(l.Length * 1000) == (int)(val * 1000)))
            {
                if (item.Equals(newLin) || item.Constraints.OfType<EqualLengthConstraint>().Count() != 0) continue;
                try
                {
                ps.GeometricConstraints.AddEqualLength(item, newLin);
                }
                catch (Exception)
                {
                }
            }
        }

        static public CutFeature addCut(SheetMetalComponentDefinition smcd, string name)
        {
            SheetMetalFeatures smf = (SheetMetalFeatures)smcd.Features;
            PlanarSketch ps = smcd.Sketches[name];
            CutDefinition cd = smf.CutFeatures.CreateCutDefinition(ps.Profiles.AddForSolid());
            try
            { cd.SetCutAcrossBendsExtent("Толщина"); }
            catch (Exception)
            {
                cd.SetCutAcrossBendsExtent("Thickness");
            }
            return smf.CutFeatures.Add(cd);
        }

        static public MirrorFeature addMirror(SheetMetalComponentDefinition smcd, ObjectCollection objs, string name, UnitVector vec)
        {
            SheetMetalFeatures smf = (SheetMetalFeatures)smcd.Features;
            WorkPlane wp = smcd.WorkPlanes.OfType<WorkPlane>().First(e => InvDoc.u.eq(vec, e.Plane.Normal));
            MirrorFeature mir = smf.MirrorFeatures.Add(objs, wp, false, PatternComputeTypeEnum.kAdjustToModelCompute);
            mir.Name = name;
            return mir;
        }

        static public string getProp<T>(T ob, string nameProp)
        {
            Type t = typeof(T);
            System.Reflection.PropertyInfo pi = t.GetProperty(nameProp);
            if (pi == null) return null;
            return "Yes";
        }

        static public PlanarSketch projectAcrosParts(ref AssemblyComponentDefinition acd, string [] nameIn, string nameOut, string sketchName, ref SheetMetalComponentDefinition smcd)
        {
            PlanarSketch psIn, psOut;
            ComponentOccurrence occ2;
            SheetMetalComponentDefinition smcd1;
            Profile pr;
            object sketchLineProxy;
            object psp;
            PlanarSketchProxy plsP;
            SketchEntity se;
            occ2 = findFlange(acd);
            PartComponentDefinition def = (PartComponentDefinition)occ2.Definition;
            psOut = def.Sketches.OfType<PlanarSketch>().FirstOrDefault(s => s.Name == sketchName);
            if (psOut == null)
            {
                psOut = def.Sketches.Add(def.SurfaceBodies[1].Faces.OfType<Face>().Where(f => f.CreatedByFeature is FaceFeature).OrderByDescending(o => o.Evaluator.Area).ElementAt(1));
            }
            occ2.CreateGeometryProxy(psOut, out psp);
            plsP = (PlanarSketchProxy)psp;
            if (((Document)occ2.Definition.Document).SubType == "{9C464203-9BAE-11D3-8BAD-0060B0CE6BB4}")
            {
                smcd = (SheetMetalComponentDefinition)occ2.Definition;
            }

            foreach (ComponentOccurrence occ1 in acd.Occurrences)
            {
                if (occ1 == null) continue;
                if (((Document)occ1.Definition.Document).SubType == "{9C464203-9BAE-11D3-8BAD-0060B0CE6BB4}")
                {
                    smcd1 = (SheetMetalComponentDefinition)occ1.Definition;
                    if (!(smcd1.Features[1] is ContourFlangeFeature)) continue;
                }
                psIn = ((PartComponentDefinition)occ1.Definition).Sketches[sketchName];

                pr = psIn.Profiles.AddForSolid();

                psIn.Visible = false; 

                //ProjectedCut pc = psIn.ProjectedCuts[1];

                occ1.CreateGeometryProxy(pr, out sketchLineProxy);

                //plsP.AddByProjectingEntity(sketchLineProxy);

                //foreach (SketchEntity s in ((ProjectedCut)sketchLineProxy).SketchEntities)
                //{
                //    if (s.)
                //    plsP.AddByProjectingEntity(s);
                //}

                foreach (ProfilePath pp in (Profile)sketchLineProxy)
                {
                    foreach (ProfileEntity ent in pp)
                    {
                        se = plsP.AddByProjectingEntity(ent.SketchEntity);
                    }
                }
            }
            return psOut;
        }

        static public ComponentOccurrence findOcc(ref AssemblyComponentDefinition acd ,string name)
        {
            return acd.Occurrences.OfType<ComponentOccurrence>().FirstOrDefault(occ => ((Document)occ.Definition.Document).PropertySets[3][14].Value.ToString().ToLower().IndexOf(name) != -1);
        }

        static public ComponentOccurrence findFlange(AssemblyComponentDefinition acd)
        {
            foreach (ComponentOccurrence occ1 in acd.Occurrences)
            {
                if (((Document)occ1.Definition.Document).SubType == "{9C464203-9BAE-11D3-8BAD-0060B0CE6BB4}")
                {
                    SheetMetalComponentDefinition smcd = (SheetMetalComponentDefinition)occ1.Definition;
                    if (smcd.ReferenceComponents.DerivedPartComponents.Count == 1
                        && (smcd.ReferenceComponents.DerivedPartComponents[1].Definition as DerivedPartUniformScaleDef).Mirror) continue;
                    if (smcd.Features[1] is FaceFeature) return occ1;
                }
            }
            return null;
        }

        static public SketchLine midleLine(SketchLine sl1, SketchLine sl2)             
        {
            PlanarSketch ps = (PlanarSketch)sl1.Parent;
            SketchPoint sp1 = ps.SketchPoints.Add(sl1.Geometry.MidPoint, false);
            SketchPoint sp2 = ps.SketchPoints.Add(sl2.Geometry.MidPoint, false);
            SketchLine sl3 = ps.SketchLines.AddByTwoPoints(sp1, sp2);
            MidpointConstraint mpc = ps.GeometricConstraints.AddMidpoint(sp1, sl1);
            mpc = ps.GeometricConstraints.AddMidpoint(sp2, sl2);
            sl3.Construction = true;
            return sl3;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            a = Convert.ToDouble(this.textBox1.Text.Replace(',', separator))/10;
            b = Convert.ToDouble(this.textBox2.Text.Replace(',', separator))/10;
            this.Hide();
                try
                { 
                    invApp = (Inventor.Application)m_PartDoc.Parent;
                    if (m_PartDoc.SubType == "{9C464203-9BAE-11D3-8BAD-0060B0CE6BB4}") smcd = (SheetMetalComponentDefinition)m_PartDoc.ComponentDefinition;
                    tr = invApp.TransactionManager.StartTransaction((_Document)m_PartDoc, "Шипы");
                    objCol = invApp.TransientObjects.CreateObjectCollection();
                    if (asmDoc != null)
                    {
                        ComponentOccurrence co = asmDoc.ComponentDefinition.Occurrences.OfType<ComponentOccurrence>().FirstOrDefault(c => c.Name.ToLower().IndexOf("обтекатель") != -1);
                        if (co != null)
                        {
                            //PartDocument pDoc;
                            if (((Document)co.Definition.Document).SubType == "{9C464203-9BAE-11D3-8BAD-0060B0CE6BB4}") 
                                smcd = (SheetMetalComponentDefinition)((PartDocument)co.Definition.Document).ComponentDefinition;
                            PlanarSketch plsk = smcd.Sketches.OfType<PlanarSketch>().FirstOrDefault(p => p.Name == "Шип");
                            if (plsk == null)
                            {
                                project(smcd, "Шип");
                            }
                        }
                        return;
                    }
                    profiles = selectFromSketch();
                    //offsetAdaptive(profiles);
                    offset(profiles);
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
    }

    internal class OffsetBtn : Button
    {
        public static Offset m_Offset;
        public static spike m_spike;
        public Inventor.Document pDoc { get; set; }
        public static Offset getOffset
        {
            get
            {
                return m_Offset;
            }
        }
        public static spike getspike
        {
            get
            {
                return m_spike;
            }
        }

        #region "Methods"
        public OffsetBtn(string displayName, string internalName, string clientId, string description, string tooltip, Icon standardIcon, Icon largeIcon)
            : base(displayName, internalName, clientId, description, tooltip, standardIcon, largeIcon)
        {

        }

        protected override void ButtonDefinition_OnExecute(NameValueMap context)
        {
            //System.Windows.Forms.Application.Run(m_Break = new BreakOp(Inventor.Document oDoc));
            try
            {
                Macros.StandardAddInServer.forms.Add(m_spike);
                PlanarSketch sketch = (PlanarSketch)InventorApplication.ActiveEditObject;
                if (Macros.StandardAddInServer.activeteForm()) System.Windows.Forms.Application.Run(m_spike = new InvAddIn.spike(InventorApplication,sketch));
            }
            catch (Exception)
            {
               Macros.StandardAddInServer.forms.Add(m_Offset);
               if (Macros.StandardAddInServer.activeteForm()) System.Windows.Forms.Application.Run(m_Offset = new Offset(InventorApplication.ActiveDocument));
            } 
        }

        #endregion
    }
}

