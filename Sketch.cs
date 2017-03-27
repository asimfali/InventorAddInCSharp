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

        public SketchBtn(Icon standardIcon, Icon largeIcon, string displayName = "Центр массива", string internalName = "Autodesk:Macros:ArrayCenter", CommandTypesEnum commandType = CommandTypesEnum.kShapeEditCmdType, string clientId = "{3FF41256-8915-4D38-A76C-01C80A3A709A}",
            string description = "Начало массива", string tooltip = "Начало массива",  ButtonDisplayEnum buttonDisplayType = ButtonDisplayEnum.kDisplayTextInLearningMode)
            : base(displayName, internalName, commandType, clientId, description, tooltip, standardIcon, largeIcon, buttonDisplayType)
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

    internal class UpdateBtn : Button
    {
        public static UpdateOp uo;

        public UpdateBtn(Icon standardIcon, Icon largeIcon, string displayName = "Обновить", string internalName = "Autodesk:Macros:Update", CommandTypesEnum commandType = CommandTypesEnum.kNonShapeEditCmdType, string clientId = "{69952C06-27CD-4E89-B216-BAE9F318FB45}",
            string description = "Обновить", string tooltip = "Обновить", ButtonDisplayEnum buttonDisplayType = ButtonDisplayEnum.kDisplayTextInLearningMode)
            : base(displayName, internalName, commandType, clientId, description, tooltip, standardIcon, largeIcon, buttonDisplayType)
        {
        }
        protected override void ButtonDefinition_OnExecute(NameValueMap context)
        {
            if (Macros.StandardAddInServer.m_inventorApplication.ActiveDocument.DocumentType == DocumentTypeEnum.kPartDocumentObject)
            {
                uo = new UpdateOp();
            }
        }
    }


    public class UpdateOp
    {
        public UpdateOp()
        {
            Document DocumentObject = Macros.StandardAddInServer.m_inventorApplication.ActiveDocument;
            if (DocumentObject.DocumentType == DocumentTypeEnum.kPartDocumentObject)
            {
                ComboBoxDefinition cbd = Macros.StandardAddInServer.m_inventorApplication.CommandManager.ControlDefinitions.OfType<ComboBoxDefinition>().FirstOrDefault(b => b.InternalName == "Autodesk:Macros:CB");
                if (cbd == null) return;
                Macros.StandardAddInServer.data = new InvAddIn.DataToArray(DocumentObject as PartDocument, cbd);
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
}
