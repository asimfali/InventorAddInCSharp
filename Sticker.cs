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
    public partial class Sticker : Form
    {
        public Inventor.Application invApp;
        public Inventor.PartDocument m_Doc;
        public Inventor.AssemblyDocument m_Asm;
        public AssemblyComponentDefinition compDef;
        private Inventor.Box box;
        private Inventor.TransientGeometry m_TG;
        private string[] spl;
        private Inventor.Point insPt;
        private Vector vec;
        private UnitVector vecX, vecY, vecZ;
        public int cc, cr;
        public string value = "";
        public string textheader, columnname, val = "";
        public int rowheight, columnwidth;
        private InvDoc.XML tbl;
        private List<string> attrXML, valXML;
        private List<System.Drawing.Point> lstTxt;
        private Dictionary<float, float> textComp, textCompRev;
        private List<string> styleParam;
        Face face; bool flag; double R; object edge1, edge2; bool triangle;
        Edge ent;
        private UserInputEvents input;
        private InteractionEvents intEvts;
        private Inventor.MouseEvents mouse;
        private SelectEvents sel;
        private Matrix pos, rot;
        private double offset1 = 0, offset2 = 0;
        private bool mirror = false;
        char separator = System.Globalization.CultureInfo.CurrentCulture.NumberFormat.CurrencyDecimalSeparator[0];
        public Sticker(Inventor.Document pDoc, Inventor.Application m_InvApp)
        {
            InitializeComponent();
            invApp = m_InvApp;
            m_TG = invApp.TransientGeometry;
            if (pDoc.DocumentType == DocumentTypeEnum.kPartDocumentObject)
            {
                m_Doc = (PartDocument)pDoc;
            }
            if (pDoc.DocumentType == DocumentTypeEnum.kAssemblyDocumentObject)
            {
                m_Asm = (AssemblyDocument)pDoc;
                compDef = m_Asm.ComponentDefinition;
                box = compDef.RangeBox;
            }
            valXML = new List<string>();
            attrXML = new List<string>();
            if (System.IO.File.Exists(@"C:\ProgramData\Autodesk\Inventor Addins\Stickers.xml"))
            {
                tbl = new InvDoc.XML(@"C:\ProgramData\Autodesk\Inventor Addins\Stickers.xml");
                tbl.ReadXML("Sticker", ref valXML, ref attrXML);
                if (attrXML.Count != 0)
                {
                    for (int i = 0; i < attrXML.Count; i++)
                    {
                        if (attrXML[i].StartsWith("Name="))
                        {
                            string XMLVal = tbl.substring(attrXML[i], "Name=");
                            this.comboBox1.Items.Add(XMLVal);
                        }
                    }
                }
                
            }
        }

        public class gab
        {
            public Edge left, right, bottom, top;

            public gab(Face f)
            {
                IEnumerable<Edge> edges = f.Edges.OfType<Edge>().OrderByDescending(l => InvDoc.u.getLenght(l)).Take(4);
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {   
            try
            {
                this.Hide();
                double h = 1, w = 1;
                if (m_Asm != null)
                {
                    CommandManager cmdMgr = invApp.CommandManager;
                    string txt = "";
                    //input = cmdMgr.UserInputEvents;       
                    intEvts = cmdMgr.CreateInteractionEvents();
                    intEvts.InteractionDisabled = false;
                    sel = intEvts.SelectEvents;
                    sel.AddSelectionFilter(SelectionFilterEnum.kPartFaceFilter);
                    sel.OnSelect += new SelectEventsSink_OnSelectEventHandler(select);
                    mouse = intEvts.MouseEvents;
                    mouse.OnMouseClick += new MouseEventsSink_OnMouseClickEventHandler(selPt);
                    intEvts.Start();
                    intEvts.StatusBarText = "Выберите поверхность:";
                    //input.OnSelect += new UserInputEventsSink_OnSelectEventHandler(select);
                    //face = (Face)cmdMgr.Pick(SelectionFilterEnum.kPartFaceFilter, "Выберите поверхность:");
                    //input.OnSelect -= new UserInputEventsSink_OnSelectEventHandler(select);
                    flag = true;
                    while (flag)
                    {
                        invApp.UserInterfaceManager.DoEvents();
                    }
                    intEvts.Stop();      
                    sel.OnSelect -= new SelectEventsSink_OnSelectEventHandler(select);
                    intEvts.MouseEvents.OnMouseClick -= new MouseEventsSink_OnMouseClickEventHandler(selPt);
                    List<string> tmp = new List<string>();
                    string name = comboBox1.Text;
                    m_Doc = (PartDocument)invApp.Documents.Add(DocumentTypeEnum.kPartDocumentObject, CreateVisible: false);
                    m_Doc.ComponentDefinition.BOMStructure = BOMStructureEnum.kPurchasedBOMStructure;
                    m_Doc.UnitsOfMeasure.LengthUnits = UnitsTypeEnum.kMillimeterLengthUnits;
                    string filename="";
                    for (int i = 0; i < attrXML.Count; i++)
                    {
                        if (attrXML[i].StartsWith("Name=" + name))
                        {
                            tmp.Add(attrXML[i]);
                            tmp.Add(valXML[i]);
                            i++;
                            while (i < attrXML.Count && attrXML[i].StartsWith("RowHeight="))
                            {
                                tmp.Add(attrXML[i]);
                                tmp.Add(valXML[i]);
                                i++;
                            }
                            break;
                        }
                    }
                    string nameStricker = "";
                    string strTmp = tbl.substring(tmp[0], "PartNumber=");
                    nameStricker = strTmp;
                    if (strTmp != "") { m_Doc.PropertySets[3][2].Value = strTmp; filename += strTmp; }
                    strTmp = tbl.substring(tmp[0], "Description=");
                    if (strTmp != "") 
                    {
                        m_Doc.PropertySets[3][14].Value = strTmp;
                        filename += (filename == "") ? strTmp : "-" + strTmp;
                    }

                    Func<string, double> conv = s =>
                    {
                        char rep = (separator == ',') ? '.' : ',';
                        return (s.IndexOf(separator) != -1) ? Convert.ToDouble(s) : Convert.ToDouble(s.Replace(rep, separator));
                    };

                    strTmp = tbl.substring(tmp[0], "Height=");
                    if (strTmp != "") h = conv(strTmp);
                    strTmp = tbl.substring(tmp[0], "Width=");
                    if (strTmp != "") w = conv(strTmp);
                    strTmp = tbl.substring(tmp[0], "txt=");
                    if (strTmp != "") txt = strTmp;
                    strTmp = tbl.substring(tmp[0], "Sort=");
                    if (strTmp != "") m_Doc.PropertySets[1][4].Value = strTmp;
                    strTmp = tbl.substring(tmp[0], "Note=");
                    if (strTmp != "") m_Doc.PropertySets[1][5].Value = strTmp;
                    strTmp = tbl.substring(tmp[0], "Triangle=");
                    if (strTmp == "False") triangle = false;
                    else triangle = true;
                    strTmp = tbl.substring(tmp[0], "Offset1=");
                    if (strTmp != "") offset1 = conv(strTmp) / 10;
                    strTmp = tbl.substring(tmp[0], "Offset2=");
                    if (strTmp != "") offset2 = conv(strTmp) / 10;
                    strTmp = tbl.substring(tmp[0], "Mirror=");
                    if (strTmp != "") mirror = true;

                    if (face.SurfaceType == SurfaceTypeEnum.kPlaneSurface)
                    {
                        filename += "-" + w.ToString() + "x" + h.ToString();
                        if (!System.IO.File.Exists(path(m_Asm.FullFileName) + filename + ".ipt"))
                        {
                            Plane plane = (Plane)face.Geometry;
                            vecZ = plane.Normal;
                            if (!triangle)
                                planarSticker(ref m_Doc, h, w, txt);
                            else triangleSticker(ref m_Doc, h, w, txt);
                        }
                    }
                    else if (face.SurfaceType == SurfaceTypeEnum.kCylinderSurface)
                    {
                        R = Math.Round(((Cylinder)face.Geometry).Radius,2);
                        filename += "-" + w.ToString() + "x" + h.ToString() + "x" + (R*10).ToString();
                        if (!System.IO.File.Exists(path(m_Asm.FullFileName) + filename + ".ipt"))
                        {
                            if (this.radioButton1.Checked)
                            { circleSticker(ref m_Doc, w, h, R, txt); filename += "(верт)"; }
                            else circleSticker(ref m_Doc, h, w, R, txt);
                        }
                    }
                    filename = path(m_Asm.FullFileName) + filename + ".ipt";
                    if (!System.IO.File.Exists(filename))
                    {
                        m_Doc.SaveAs(filename, false);
                        Document tmp_Doc = invApp.Documents.Open(filename, false);
                        tmp_Doc.PropertySets[3][2].Value = nameStricker;
                        tmp_Doc.Save();
                        tmp_Doc.Close();
                        m_Doc.Close();
                    }

                        Inventor.Point pt = m_TG.CreatePoint();
                        //pos = m_TG.CreateMatrix();
                        
                        
                        //rot.SetCoordinateSystem(insPt, m_TG.CreateVector(YCoord: 1), m_TG.CreateVector(1) , m_TG.CreateVector(ZCoord: 1));
                        
                        //pos.TransformBy(rot);
//                         if (face.SurfaceType == SurfaceTypeEnum.kPlaneSurface && !System.IO.File.Exists(filename))
//                         {
                        
                        //}

                    AssemblyComponentDefinition compDef = m_Asm.ComponentDefinition;
                    //pos.Cell[4, 1] = insPt.X; pos.Cell[4, 2] = insPt.Y; pos.Cell[4, 3] = insPt.Z;
                    rot = m_TG.CreateMatrix();
                    ComponentOccurrence occ = compDef.Occurrences.Add(filename,rot);
                    Vector vec1 = InvDoc.u.getAxis(face);
                    Vector vec2 =  InvDoc.u.getAxis(occ);
                    if (!InvDoc.u.eq(vec1, vec2)) rot.SetToRotateTo(vec1,vec2);
                    //m_Asm.ComponentDefinition.WorkPoints.AddFixed(insPt);
                    vec = pt.VectorTo(insPt);
                    rot.SetTranslation(vec);
                    //rot.Cell[1, 4] = insPt.X; rot.Cell[2, 4] = insPt.Y; rot.Cell[3, 4] = insPt.Z;
                    //occ.Transformation = rot; 
                    occ.SetTransformWithoutConstraints(rot);
                            //rot.SetToRotateTo(m_TG.CreateVector(ZCoord: 1), vecZ.AsVector(), m_TG.CreateVector(XCoord: 1));
                            //rot.SetTranslation(vec);
                    //return;
                    //Face face2 = occ.SurfaceBodies[1].Faces[5];
                    object fac = null;
                    occ.CreateGeometryProxy(((PartComponentDefinition)occ.Definition).WorkPlanes[3], out fac);
                    WorkPlaneProxy face2 = (WorkPlaneProxy)fac;
                    PartComponentDefinition def = (PartComponentDefinition)occ.Definition;
                    //Inventor.Point pt2 = def.WorkPoints[1].Point;
                    //pos.Invert();
                    //pt2.TransformBy(pos);
                    if (face.SurfaceType == SurfaceTypeEnum.kPlaneSurface)
                    {
                        compDef.Constraints.AddFlushConstraint(face, face2, 0);
                    }
                    else if (face.SurfaceType == SurfaceTypeEnum.kCylinderSurface)
                    {
                        //foreach (Face f in occ.SurfaceBodies[1].Faces)
                        //{
                        //    if (f.SurfaceType == SurfaceTypeEnum.kCylinderSurface)
                        //    {
                        //        if (Math.Round(((Cylinder)f.Geometry).Radius, 2) == R)
                        //        {
                        //            face2 = f; break;
                        //        }
                        //    }
                        //}
                        occ.CreateGeometryProxy(((PartComponentDefinition)occ.Definition).WorkAxes[2], out fac);
                        compDef.Constraints.AddMateConstraint((WorkAxisProxy)fac, face, 0, InferredTypeEnum.kInferredLine, InferredTypeEnum.kInferredLine);
                    }
                    //Matrix trans = m_TG.CreateMatrix();
                    //trans.SetTranslation(vec,true);
                    //occ.Transformation.Cell[1, 4] = insPt.X; occ.Transformation.Cell[2, 4] = insPt.Y; occ.Transformation.Cell[3, 4] = insPt.Z;
                    
                    object face1 = null;

                    if (face.SurfaceType == SurfaceTypeEnum.kCylinderSurface)
                    {
                        if (offset1 > 0.05 && offset1 < 0.4)
                        {
                            //PartComponentDefinition pDef = ((PartDocument)((FaceProxy)face).ContainingOccurrence.Definition.Document).ComponentDefinition;
                            ((FaceProxy)face).ContainingOccurrence.CreateGeometryProxy(InvDoc.u.getPlane((FaceProxy)face), out fac);
                            //createProxy((FaceProxy)face, (int)(offset1 * 10), ref fac);
                            //occ.CreateGeometryProxy(((PartComponentDefinition)occ.Definition).WorkPlanes[(int)(offset1 * 10)], out fac);
                            face1 = fac;
                            offset1 = 0;
                        }
                        else /*if (offset1 != 0)*/
                        {
                            offset1 += w / 20;
                            ent = (Edge)cmdMgr.Pick(SelectionFilterEnum.kPartEdgeCircularFilter, "Выберите торцевое ребро:");
                            if (ent.Faces[1].SurfaceType == SurfaceTypeEnum.kPlaneSurface)
                            {
                                face1 = ent.Faces[1];
                            }
                            else
                            {
                                face1 = ent.Faces[2];
                            }
                        }
//                         if (offset2 != 0)
//                         {
//                             offset2 += h / 20;
//                         }
                        //face1 = cmdMgr.Pick(SelectionFilterEnum.kPartEdgeCircularFilter, "Выберите тор"
                    }
                    else
                    {
                        if (offset1 > 0.05 && offset1 < 0.4)
                        {
                            createProxy((FaceProxy)face, (int)(offset1 * 10), ref fac);
                            //occ.CreateGeometryProxy(((PartComponentDefinition)occ.Definition).WorkPlanes[(int)(offset1 * 10)], out fac);
                            edge1 = fac;
                            offset1 = 0;
                        }
                        else /*if (offset1 != 0)*/
                        {
                            offset1 += h / 20;
                            edge1 = cmdMgr.Pick(SelectionFilterEnum.kAllEntitiesFilter, "Выберите ребро или плоскость:");
                        }
                    }
                    //face2 = occ.SurfaceBodies[1].Faces[2];
                    occ.CreateGeometryProxy(((PartComponentDefinition)occ.Definition).WorkPlanes[2], out fac);
                    face2 = (WorkPlaneProxy)fac;
                    MateConstraint mc = null; FlushConstraint fc = null;
                    
//                     if (offset2 < 0.4)
//                     {
//                         face1 = createProxy(occ, (int)(offset2 * 10)); //compDef.WorkPlanes[(int)(offset2*10)];
//                     }
                    //else if (offset2 != 0) offset2 += h / 20;
                    if (face.SurfaceType == SurfaceTypeEnum.kCylinderSurface)
                        fc = compDef.Constraints.AddFlushConstraint(face2, face1, offset1);
                    else
                        mc = compDef.Constraints.AddMateConstraint(face2, edge1, offset1);   
//                     Inventor.Point ptmc = occ.RangeBox.MinPoint;
//                     SelectionFilterEnum [] f = {SelectionFilterEnum.kPartFaceFilter};
//                     ObjectsEnumerator en = compDef.FindUsingPoint(ptmc, ref f, 0.1);
                    if (/*en.Count < 3*/ InvDoc.u.intersPoint(compDef, occ, (FaceProxy)face))
                    {
                        if (mc != null)
                            mc.Offset.Value = - (double)mc.Offset.Value;
                        //else fc.Offset.Value = - (double)fc.Offset.Value;
                    }
                    if (offset2 > 0.05 && offset2 < 0.4)
                    {
                        createProxy((FaceProxy)face, (int)(offset2 * 10), ref fac);
                        //occ.CreateGeometryProxy(((PartComponentDefinition)face.ReferenceComponent.ReferencedDocumentDescriptor.ReferencedDocument.Definition).WorkPlanes[(int)(offset2 * 10)], out fac);
                        edge2 = fac;
                        offset2 = 0;
                    }
                    else /*if (offset2 != 0)*/
                    {
                        if (face.SurfaceType == SurfaceTypeEnum.kCylinderSurface)
                        {
                            offset2 += h / 20;
                        }
                        else offset2 += w / 20;
                        edge2 = cmdMgr.Pick(SelectionFilterEnum.kAllEntitiesFilter, "Выберите ребро или плоскость:");
                    }

                    //face2 = occ.SurfaceBodies[1].Faces[1];
                    occ.CreateGeometryProxy(((PartComponentDefinition)occ.Definition).WorkPlanes[1], out fac);
                    face2 = (WorkPlaneProxy)fac;
                    mc = compDef.Constraints.AddMateConstraint(face2, edge2, offset2);
                    if (mc.HealthStatus == HealthStatusEnum.kInconsistentHealth)
                    {
                        occ.Constraints[2].ConvertToFlushConstraint(occ.Constraints[2].EntityOne, occ.Constraints[2].EntityTwo, offset1);
                    }
                    //en = compDef.FindUsingPoint(ptmc, ref f, 0.1);
                    m_Asm.Update();
                    if (/*en.Count < 3*/ InvDoc.u.intersPoint(compDef, occ, (FaceProxy)face)) 
                    {
                        mc.Delete();
                        mc = compDef.Constraints.AddMateConstraint(face2, edge2, offset2);
                        m_Asm.Update();
                        if (InvDoc.u.intersPoint(compDef, occ, (FaceProxy)face))
                        {
                            //mc.Delete();
                            //mc = compDef.Constraints.AddMateConstraint(face2, edge1, -offset2);
                            double dist = Macros.StandardAddInServer.m_inventorApplication.MeasureTools.GetMinimumDistance(occ.RangeBox.MinPoint, edge2);
                            if (dist > offset2*2)
                            {
                                rot = m_TG.CreateMatrix();
                                Vector tr = occ.Transformation.Translation;
                                rot.SetToRotation(Math.PI, InvDoc.u.getAxis(occ), m_TG.CreatePoint(tr.X, tr.Y, tr.Z));
                                mc.Delete();
                                occ.Transformation = rot;
                                mc = compDef.Constraints.AddMateConstraint(face2, edge2, offset2);
                                if (InvDoc.u.intersPoint(compDef, occ, (FaceProxy)face))
                                {
                                    mc.Offset.Value = -(double)mc.Offset.Value;
                                }
                            }
                            else if (InvDoc.u.intersPoint(compDef, occ, (FaceProxy)face))
                                mc.Offset.Value = -(double)mc.Offset.Value;
                            m_Asm.Update();
                        }
                        //mc.Offset.Value = -(double)mc.Offset.Value;
                    }
                    //occ.Transformation = trans;
                }
                
                this.Close();
            }
            catch (Exception)
            {
                //MessageBox.Show(ex.ToString());
                this.Close();
            }
        }

        private void createProxy(FaceProxy face, int num, ref object fac)
        {
            ComponentOccurrence occ = face.ContainingOccurrence;
            occ.CreateGeometryProxy(((PartComponentDefinition)occ.Definition).WorkPlanes[num], out fac);
        }

        private void select(ObjectsEnumerator JustSelectedEntities, SelectionDeviceEnum SelectionDevice, Inventor.Point ModelPosition, Point2d ViewPosition, Inventor.View View)
        {
            insPt = ModelPosition;
            face = (Face)JustSelectedEntities[1];
            flag = false;
        }

        private void selPt(MouseButtonEnum Button, ShiftStateEnum ShiftKeys, Inventor.Point ModelPosition, Point2d ViewPosition, Inventor.View View)
        {
            insPt = ModelPosition;
            flag = false;            
        }


        private void planarSticker(ref PartDocument pDoc, double heigth, double width, string txt="", double thickness = 0.01)
        {
            PartComponentDefinition compDef = pDoc.ComponentDefinition;
            PlanarSketch ps = compDef.Sketches.Add(compDef.WorkPlanes[3]);
            ps.SketchLines.AddAsTwoPointCenteredRectangle(m_TG.CreatePoint2d(), m_TG.CreatePoint2d(width / 20, heigth / 20));
            Profile pr = ps.Profiles.AddForSolid();
            ExtrudeDefinition extDef = compDef.Features.ExtrudeFeatures.CreateExtrudeDefinition(pr, PartFeatureOperationEnum.kNewBodyOperation);
            extDef.SetDistanceExtent(thickness/10, PartFeatureExtentDirectionEnum.kPositiveExtentDirection);  
            ExtrudeFeature extr = compDef.Features.ExtrudeFeatures.Add(extDef);
            extr.Appearance = InvDoc.u.createColor(pDoc, "white_", "белый_", 255, 255, 255);
            if (txt != "")
            {
                ps = compDef.Sketches.Add(extr.EndFaces[1]);
                Inventor.TextBox tb = ps.TextBoxes.AddByRectangle(pts(-heigth, -width), pts(heigth, width), txt);
                tb.SingleLineText = false;
                tb.HorizontalJustification = HorizontalTextAlignmentEnum.kAlignTextCenter;
                tb.VerticalJustification = VerticalTextAlignmentEnum.kAlignTextMiddle;
                //tb.Rotation = Math.PI;
                if (mirror) tb.Rotation = Math.PI;
                EmbossFeature emb = compDef.Features.EmbossFeatures.AddEmbossFromFace(ps.Profiles.AddForSolid(false, ps.TextBoxes), 0.0005, PartFeatureExtentDirectionEnum.kPositiveExtentDirection);
                emb.Appearance = InvDoc.u.createColor(pDoc, "black_", "черный_", 0, 0, 0);
            }
        }

        private void triangleSticker(ref PartDocument pDoc, double heigth, double width, string txt = "", double thickness = 0.01)
        {
            PartComponentDefinition compDef = pDoc.ComponentDefinition;
            PlanarSketch ps = compDef.Sketches.Add(compDef.WorkPlanes[3]);
            SketchLine line1 = ps.SketchLines.AddByTwoPoints(m_TG.CreatePoint2d(-width/20,-heigth/30),m_TG.CreatePoint2d(width/20, -heigth/30));
            SketchLine line2 = ps.SketchLines.AddByTwoPoints(line1.StartSketchPoint, m_TG.CreatePoint2d(0, heigth / 30*2));
            ps.SketchLines.AddByTwoPoints(line2.EndSketchPoint, line1.EndSketchPoint);
            Profile pr = ps.Profiles.AddForSolid();
            ExtrudeDefinition extDef = compDef.Features.ExtrudeFeatures.CreateExtrudeDefinition(pr, PartFeatureOperationEnum.kNewBodyOperation);
            extDef.SetDistanceExtent(thickness / 10, PartFeatureExtentDirectionEnum.kPositiveExtentDirection);
            ExtrudeFeature extr = compDef.Features.ExtrudeFeatures.Add(extDef);
            extr.Appearance = InvDoc.u.createColor(pDoc, "red_", "красный_", 255, 0, 0);
            if (txt != "")
            {
                ps = compDef.Sketches.Add(extr.EndFaces[1]);
                Inventor.TextBox tb = ps.TextBoxes.AddFitted(m_TG.CreatePoint2d(), txt);
                tb.HorizontalJustification = HorizontalTextAlignmentEnum.kAlignTextCenter;
                tb.VerticalJustification = VerticalTextAlignmentEnum.kAlignTextMiddle;
                tb.SingleLineText = false;
                //tb.Rotation = Math.PI;
                if (mirror) tb.Rotation = Math.PI;
                EmbossFeature emb = compDef.Features.EmbossFeatures.AddEmbossFromFace(ps.Profiles.AddForSolid(false, ps.TextBoxes), 0.0005, PartFeatureExtentDirectionEnum.kPositiveExtentDirection);
                emb.Appearance = InvDoc.u.createColor(pDoc, "black_", "черный_", 0, 0, 0);
            }
        }

        private void circleSticker(ref PartDocument pDoc, double heigth, double width, double R, string txt = "", double thickness = 0.1)
        {
            PartComponentDefinition compDef = pDoc.ComponentDefinition;
            heigth /= 10; width /= 10; thickness /= 10;
            PlanarSketch ps = compDef.Sketches.Add(compDef.WorkPlanes[2]);
            //SketchEntity origin = ps.AddByProjectingEntity(compDef.WorkPoints[1]);
            double alfa = alfaFromL(heigth, R);
            Point2d startPt = m_TG.CreatePoint2d();

            SketchArc arc = ps.SketchArcs.AddByCenterStartSweepAngle(startPt,R,3.1415926/2-alfa/2,alfa);
            //ObjectCollection objs = invApp.TransientObjects.CreateObjectCollection();
            //objs.Add(arc);
            //ps.RotateSketchObjects(objs, arc.CenterSketchPoint.Geometry, 1,RemoveConstraints:true);
            //arc = (SketchArc)en[1];
            //ps.GeometricConstraints.AddVerticalAlign((SketchPoint)origin, arc.CenterSketchPoint);
            ObjectCollection objCol = invApp.TransientObjects.CreateObjectCollection();
            objCol.Add(arc);
            SketchEntitiesEnumerator sArc = ps.OffsetSketchEntitiesUsingDistance(objCol, thickness, true);
            SketchArc arc2 = (SketchArc)sArc[1];
            ps.SketchLines.AddByTwoPoints(arc2.StartSketchPoint, arc.StartSketchPoint);
            ps.SketchLines.AddByTwoPoints(arc2.EndSketchPoint, arc.EndSketchPoint);
            Profile pr = ps.Profiles.AddForSolid();
            ExtrudeDefinition extDef = compDef.Features.ExtrudeFeatures.CreateExtrudeDefinition(pr, PartFeatureOperationEnum.kNewBodyOperation);
            extDef.SetDistanceExtent(width/2, PartFeatureExtentDirectionEnum.kNegativeExtentDirection);
            extDef.SetDistanceExtentTwo(width/2);
            ExtrudeFeature extr = compDef.Features.ExtrudeFeatures.Add(extDef);
            extr.Appearance = InvDoc.u.createColor(pDoc, "white_", "белый_", 255, 255, 255);
            if (txt != "")
            {
                ps = compDef.Sketches.Add(compDef.WorkPlanes[3]);
                Inventor.TextBox tb = ps.TextBoxes.AddByRectangle(pts(-heigth, - width), pts(heigth, width), txt);
                tb.HorizontalJustification = HorizontalTextAlignmentEnum.kAlignTextCenter;
                tb.VerticalJustification = VerticalTextAlignmentEnum.kAlignTextMiddle;
                tb.SingleLineText = false;
                tb.Rotation = Math.PI;
                if (mirror) tb.Rotation = Math.PI;
                if (this.radioButton1.Checked == false)
                {
                    //ObjectCollection col = invApp.TransientObjects.CreateObjectCollection();
                    //col.Add(tb);
                    //ps.RotateSketchObjects(col, pt, Math.PI / 4);
                    tb.Rotation -= Math.PI / 2;
                }
                Face face2 = null;
                foreach (Face f in compDef.SurfaceBodies[1].Faces)
                {
                    if (f.SurfaceType == SurfaceTypeEnum.kCylinderSurface)
                    {
                        double val = Math.Round(((Cylinder)f.Geometry).Radius, 3)*100;
                        val = Math.Round(val, 0);

                        if /*((int)val == (int)((R + thickness)*100))*/   (InvDoc.u.eq(val,(R + thickness)*100))   
                        {
                            face2 = f; break;
                        }
                    }
                }
                EmbossFeature emb = compDef.Features.EmbossFeatures.AddEmbossFromFace(ps.Profiles.AddForSolid(false, ps.TextBoxes), 0.0005, PartFeatureExtentDirectionEnum.kNegativeExtentDirection, WrapFace: face2);
                emb.Appearance = InvDoc.u.createColor(pDoc, "black_", "черный_", 0, 0, 0);
            }
        }

        private Point2d pts(double height, double width)
        {
            Point2d pt = m_TG.CreatePoint2d();
            Point2d pt1 = pt.Copy();
            pt1.X = height / 2; pt1.Y = width / 2;
            return pt1;
        }

        private double alfaFromL(double L, double R)
        {
            return L / (R);
        }

        public string path(string file)
        {
            return file.Substring(0,file.LastIndexOf('\\')+1);
        }
    }

    internal class StickerBtn : Button
    {
        public static Sticker m_Sticker;
        public Inventor.Document pDoc { get; set; }
        public static Sticker getSticker
        {
            get
            {
                return m_Sticker;
            }
        }

        #region "Methods"
        public StickerBtn(string displayName, string internalName, CommandTypesEnum commandType, string clientId, string description, string tooltip, Icon standardIcon, Icon largeIcon, ButtonDisplayEnum buttonDisplayType)
            : base(displayName, internalName, commandType, clientId, description, tooltip, standardIcon, largeIcon, buttonDisplayType)
        {

        }
        public StickerBtn(string displayName, string internalName, CommandTypesEnum commandType, string clientId, string description, string tooltip, ButtonDisplayEnum buttonDisplayType)
            : base(displayName, internalName, commandType, clientId, description, tooltip, buttonDisplayType)
        {

        }

        protected override void ButtonDefinition_OnExecute(NameValueMap context)
        {
            //System.Windows.Forms.Application.Run(m_Break = new BreakOp(Inventor.Document oDoc));
            try
            {
                Macros.StandardAddInServer.forms.Add(m_Sticker);
                if (Macros.StandardAddInServer.activeteForm()) System.Windows.Forms.Application.Run(m_Sticker = new InvAddIn.Sticker(InventorApplication.ActiveDocument, InventorApplication));
            }
            catch (Exception)
            {
            }
        }

        #endregion
    }
}
