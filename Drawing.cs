using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml;
using System.Xml.Linq;
using System.Text;
using System.Text.RegularExpressions;
using Inventor;
using InvDoc;
using System.Windows.Forms;
using TableDll;

namespace InvAddIn
{
    internal class DrawingBtn : Button
    {
        public static Drawings m_Drw;
        public static Drawings getDrw { get { return m_Drw; } }
        public DrawingBtn(string displayName, string internalName, CommandTypesEnum commandType, string clientId, string description,
            string tooltip, System.Drawing.Icon standardIcon, System.Drawing.Icon largeIcon, ButtonDisplayEnum buttonDisplayType)
            : base(displayName, internalName, commandType, clientId, description, tooltip, standardIcon, largeIcon, buttonDisplayType) { }
        protected override void ButtonDefinition_OnExecute(NameValueMap context)
        {
            if (I.getSS().Count == 0)
            m_Drw = new Drawings((DrawingDocument)InventorApplication.ActiveDocument);
            else
            {
                DrawingView dv = I.getSS().OfType<DrawingView>().FirstOrDefault();
                if (dv == null) return;
                List<LinearGeneralDimension> dims = I.getSS().OfType<LinearGeneralDimension>().ToList();
                if (dims == null) return;
                OrderDims dimsL = new OrderDims(dv, dims, poz.left);
                dimsL.add();
                dimsL = new OrderDims(dv, dims, poz.right);
                dimsL.add();
                dimsL = new OrderDims(dv, dims, poz.top);
                dimsL.add();
                dimsL = new OrderDims(dv, dims, poz.bottom);
                dimsL.add();
            }
    
        }
    }

    internal class ListBtn : Button
    {
        //public static Drawings m_Drw;
        //public static Drawings getDrw { get { return m_Drw; } }
        public ListBtn(string displayName, string internalName, string clientId, string description,string tooltip, 
            ButtonDisplayEnum buttonDisplayType = ButtonDisplayEnum.kDisplayTextInLearningMode, CommandTypesEnum commandType = CommandTypesEnum.kNonShapeEditCmdType)
            : base(displayName, internalName, commandType, clientId, description, tooltip, buttonDisplayType) { }
        protected override void ButtonDefinition_OnExecute(NameValueMap context)
        {
            if (Macros.StandardAddInServer.m_inventorApplication.ActiveDocument.DocumentType == DocumentTypeEnum.kDrawingDocumentObject)
            {
                sheet sh = new sheet();
                sh.ShowDialog();
            }
        }
    }

    internal class SketchInModel : Button
    {
        PartDocument doc, refDoc; PartComponentDefinition compDef; PlanarSketch ps; Face f; SelectSet ss; HoleFeature hf;
        public SketchInModel(string displayName, string internalName, string clientId, string description, string tooltip,
            ButtonDisplayEnum buttonDisplayType = ButtonDisplayEnum.kDisplayTextInLearningMode, CommandTypesEnum commandType = CommandTypesEnum.kNonShapeEditCmdType)
            : base(displayName, internalName, commandType, clientId, description, tooltip, buttonDisplayType) { }
        protected override void ButtonDefinition_OnExecute(NameValueMap context)
        {
            Document d = Macros.StandardAddInServer.m_inventorApplication.ActiveEditDocument;
            if (!(d is PartDocument)) return;
            compDef = (doc = d as PartDocument).ComponentDefinition;
            ss = d.SelectSet;
            if (ss[1] is HoleFeature) { hf = ss[1] as HoleFeature; holeDiamForm(); }
            f = (ss[1] is Face) ? ss[1] as Face:(ss[1] as FaceProxy).NativeObject;
            ObjectCollection col = I.objs.CreateObjectCollection();
            foreach (var item in ss)
            {
                col.Add(item); 
            }
            ps = compDef.Sketches.Add(f);
            if (col.Count > 1)
            {
                //ps.Edit();
                for (int i = 1; i < col.Count; i++)
                {
                    ps.AddByProjectingEntity(col[i+1]);
                }
                holeDiamForm();
                //ps.ExitEdit();
            }
        }
        public void holeDiamForm()
        {
            int offsetY = 30, offsetX = 10;
            Form f = new Form();
            f.Height = 150; f.Width = 400; f.WindowState = FormWindowState.Normal; f.Text = "Шипы"; f.StartPosition = FormStartPosition.CenterScreen;
            f.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink; f.AutoSize = true;
            System.Drawing.Point insPt = new System.Drawing.Point(5, 5);
            InterfaceDll.CB cbs;
            InterfaceDll.Lbl lbs;
            lbs = new InterfaceDll.Lbl(offsetX, offsetY, 100, 15, insPt, f, "Диаметр ");
            cbs = new InterfaceDll.CB(offsetX, offsetY, 200, 15, insPt, f);
            cbs.position(lbs.last(), true);
            ComboBox cb = cbs[0];
            foreach (UserParameter p in doc.ComponentDefinition.Parameters.UserParameters)
            {
                cb.Items.Add(p.Name);  
            }
            if (doc.ReferencedDocuments.Count == 1 && doc.ReferencedDocuments[1] is PartDocument)
            {
                foreach (UserParameter p in (refDoc = doc.ReferencedDocuments[1] as PartDocument).ComponentDefinition.Parameters.UserParameters)
                {
                    cb.Items.Add(p.Name);
                }
            }

            InterfaceDll.Btn btns = new InterfaceDll.Btn(offsetX, offsetY, 100, 20, insPt, f, spike_Click, "Добавить");
            btns.center(cbs.last(), offsetY + 5);
            f.Show();
        }

        private void spike_Click(object arg1, EventArgs arg2)
        {
            Form f = (arg1 as System.Windows.Forms.Button).Parent as Form;
            string s = f.Controls[1].Text; 
            if (doc != null && refDoc != null)
            {
                InvDoc.u.findParameter(doc as Document, s);
                //CreateComponent.addLinkParam(s, doc as Document);
                //CreateComponent.addLinkParam(doc as Document, refDoc as Document, new string[] {s});
            }
            if (hf != null)
            {
                hf.HoleDiameter.Expression = s;
            }
            else CreateComponent.hole(compDef, s, ps);
            f.Close();
        }
    }

    internal class SurfaceBtn : Button
    {
        DrawingView dv;
        DrawingDocument drw;
        DrawingCurve dc, dc1, dc2;
        LinearGeneralDimension dim;
        public SurfaceBtn(string displayName, string internalName, string clientId, string description, string tooltip, 
            ButtonDisplayEnum buttonDisplayType = ButtonDisplayEnum.kDisplayTextInLearningMode, CommandTypesEnum commandType = CommandTypesEnum.kNonShapeEditCmdType)
            : base(displayName, internalName, commandType, clientId, description, tooltip, buttonDisplayType) { }
        protected override void ButtonDefinition_OnExecute(NameValueMap context)
        {
            drw = (DrawingDocument)Macros.StandardAddInServer.m_inventorApplication.ActiveDocument;
            if (drw.SelectSet.Count == 0)
            dv = (DrawingView)Macros.StandardAddInServer.m_inventorApplication.CommandManager.Pick(SelectionFilterEnum.kDrawingViewFilter, "Выберите вид");
            else if (drw.SelectSet[1] is DrawingCurveSegment)
            {
                dv = ((DrawingCurveSegment)drw.SelectSet[1]).Parent.Parent;
            }
            if (drw.SelectSet.Count == 1 && drw.SelectSet[1] is DrawingCurveSegment)
            {
                dc = ((DrawingCurveSegment)drw.SelectSet[1]).Parent;
            }
            else if (drw.SelectSet.Count == 2 && drw.SelectSet[1] is DrawingCurveSegment && drw.SelectSet[2] is DrawingCurveSegment)
            {
                dc1 = ((DrawingCurveSegment)drw.SelectSet[1]).Parent;
                dc2 = ((DrawingCurveSegment)drw.SelectSet[2]).Parent;
            }
            add();
        }
        private void add()
        {
            if (dv == null) return;
            if (dc != null)
            {
                ObjectCollection col = I.objs.CreateObjectCollection();
                foreach (DrawingCurveSegment item in dv.Parent.FindUsingPoint(dc.StartPoint))
                {
                    if (!item.Parent.Equals(dc))
                        col.Add(item); 
                }
                if (col.Count == 0) return;
                dc1 = (col[1] as DrawingCurveSegment).Parent;
                Vector2d v = (dc.StartPoint.IsEqualTo(dc1.StartPoint)) ? dc1.EndPoint.VectorTo(dc.StartPoint): dc1.StartPoint.VectorTo(dc.StartPoint);
                dim = Drawings.addDim(dc, v);
                dim.Text.FormattedText = dim.Text.FormattedText + "*";
                Drawings.addSurfaceTextureSymbol(dim, 0.3, true);
                Drawings.addSurfaceTextureSymbol(dim, 0.3, false);
            }
//             if (dv != null)
//             {
//                 if (dc1 == null && dc2 == null)
//                 {
//                     Drawings.surfCurve(dv, ref dc, ref dc1, ref dc2);
//                     dim = Drawings.addDim(dc, -15);
//                 }
//                 else
//                 {
//                     dim = Drawings.addDim(dc1, dc2, -15, 7);
//                 }
//                 dim.Text.FormattedText = dim.Text.FormattedText + "*";
//                 double val = 0.5; Vector2d vec;
//                 if (dc1 != null)
//                 {
//                     vec = dc2.MidPoint.VectorTo(dc1.MidPoint); vec.Normalize(); vec.ScaleBy(0.1);
//                     Drawings.addSurfaceTextureSymbol(dv, dc1, val, vec);
//                 }
//                 if (dc2 != null)
//                 {
//                     vec = dc1.MidPoint.VectorTo(dc2.MidPoint); vec.Normalize(); vec.ScaleBy(0.1);
//                     Drawings.addSurfaceTextureSymbol(dv, dc2, val, vec);
//                 }
//             }
        }
    }

        internal class GabButton : Button
    {
        //public static Drawings m_Drw;
        //public static Drawings getDrw { get { return m_Drw; } }
            public GabButton(string displayName, string internalName, string clientId, string description, string tooltip, 
            ButtonDisplayEnum buttonDisplayType = ButtonDisplayEnum.kDisplayTextInLearningMode, CommandTypesEnum commandType = CommandTypesEnum.kNonShapeEditCmdType)
            : base(displayName, internalName, commandType, clientId, description, tooltip, buttonDisplayType) { }
        protected override void ButtonDefinition_OnExecute(NameValueMap context)
        {
            DrawingDocument drw = (DrawingDocument)Macros.StandardAddInServer.m_inventorApplication.ActiveDocument;
            DrawingView dv = null;
            if (drw.SelectSet.Count == 0)
            dv = (DrawingView)Macros.StandardAddInServer.m_inventorApplication.CommandManager.Pick(SelectionFilterEnum.kDrawingViewFilter, "Выберите вид");
            else if (drw.SelectSet[1] is DrawingView)
            {
                dv = (DrawingView)drw.SelectSet[1];
                gabs.add(dv); gabs.set();
                if (dv != null) { new CheckReflect(dv, I.CV2d(1)); new CheckReflect(dv, I.CV2d(0, 1)); }
            }
            else if (drw.SelectSet[1] is DrawingCurveSegment && ((drw.SelectSet[1] as DrawingCurveSegment).Parent.CurveType == CurveTypeEnum.kCircleCurve || 
                (drw.SelectSet[1] as DrawingCurveSegment).Parent.CurveType == CurveTypeEnum.kCircularArcCurve))
            {
                //util.transactStart(drw as Document, "Массив размеры");
                DrawingCurve dc = (drw.SelectSet[1] as DrawingCurveSegment).Parent as DrawingCurve;
                gabs.add(dc.Parent); gabs.set(); gabs.center();
                if (dc.CurveType == CurveTypeEnum.kCircleCurve)
                {
                    DrwArr dims = new DrwArr(dc, e => e.CurveType == CurveTypeEnum.kCircleCurve);
                    dims.add();
                }
                else if (dc.CurveType == CurveTypeEnum.kCircularArcCurve)
                {
                    DrwArr dims = new DrwArr(dc, e => e.CurveType == CurveTypeEnum.kCircularArcCurve);
                    dims.add();
                }
   
                //util.transactEnd();
            }
            if (dv != null)
            {
                gabs.add(dv); gabs.set(); gabs.center();
                Dimensions dimens = new Dimensions(dv);
                //Drawings.gab(dv);
            }
        }
    }

        internal class SplineButton : Button
        {
            public List<Point2d> pts = new List<Point2d>();
            InventorEvents ev;
            ObjectCollection col;
            Point2d vp;
            Point mp;
            DrawingView dv;

            public SplineButton(string displayName, string internalName, string clientId, string description, string tooltip, 
            ButtonDisplayEnum buttonDisplayType = ButtonDisplayEnum.kDisplayTextInLearningMode, CommandTypesEnum commandType = CommandTypesEnum.kNonShapeEditCmdType)
            : base(displayName, internalName, commandType, clientId, description, tooltip, buttonDisplayType) { }
            protected override void ButtonDefinition_OnExecute(NameValueMap context)
            {
                col = I.objs.CreateObjectCollection();
                dv = Macros.StandardAddInServer.m_inventorApplication.ActiveDocument.SelectSet[1] as DrawingView;
                if (dv == null) return;
                ev = new InventorEvents();
                ev.createEvents();
                ev.iEv.MouseEvents.OnMouseClick += MouseEvents_OnMouseClick;
                ev.iEv.KeyboardEvents.OnKeyUp += KeyboardEvents_OnKeyUp;
                ev.iEv.StatusBarText = "Выберите точки для контура";
                ev.iEv.Start();
            }

            void KeyboardEvents_OnKeyUp(int Key, ShiftStateEnum ShiftKeys)
            {
                if (Key == 32)
                {
                    ev.iEv.Stop();
                    ev.iEv.MouseEvents.OnMouseClick -= MouseEvents_OnMouseClick;
                    ev.iEv.KeyboardEvents.OnKeyUp -= KeyboardEvents_OnKeyUp;
                    addSpline(dv);
                    col.Clear();
                }
            }

//             void MouseEvents_OnMouseDoubleClick(MouseButtonEnum Button, ShiftStateEnum ShiftKeys, Point ModelPosition, Point2d ViewPosition, Inventor.View View)
//             {
//                 
//             }

            void MouseEvents_OnMouseClick(MouseButtonEnum Button, ShiftStateEnum ShiftKeys, Point ModelPosition, Point2d ViewPosition, Inventor.View View)
            {
                if (Button == MouseButtonEnum.kLeftMouseButton)
                {
                    mp = ModelPosition;
                    vp = ViewPosition;
                    //Point2d p = dv.ModelToDrawingViewSpace(ModelPosition)
                    vp.X = mp.X - dv.Position.X; vp.Y = mp.Y - dv.Position.Y;
                    Matrix2d mtx = I.tg.CreateMatrix2d();
                    //Vector2d v = I.tg.CreatePoint2d().VectorTo(dv.Position);
                    //v.ScaleBy(1/dv.Scale);
                    //vp.TranslateBy(v);
                    mtx.Cell[1, 1] = 1/dv.Scale;
                    mtx.Cell[2, 2] = 1/dv.Scale;
                    //mtx.Cell[3, 1] = dv.Position.X;
                    //mtx.Cell[3, 2] = dv.Position.Y;
                    vp.TransformBy(mtx);
                    //vp = p;
                    col.Add(vp);
                }
            }
            public void addSpline(DrawingView dv)
            {
                if (col.Count < 2) return;
                Macros.StandardAddInServer.m_inventorApplication.SilentOperation = true;
                Macros.StandardAddInServer.m_inventorApplication.ScreenUpdating = false;
                DrawingSketch ds = dv.Sketches.Add();
                ds.Edit();
                SketchSpline spl = ds.SketchSplines.Add(col, SplineFitMethodEnum.kSmoothSplineFit);
                spl.Closed = true;
                ds.ExitEdit();
                Macros.StandardAddInServer.m_inventorApplication.SilentOperation = false;
                Macros.StandardAddInServer.m_inventorApplication.ScreenUpdating = true;
                addCut(ds, dv);
            }
            public void addCut(DrawingSketch ds, DrawingView dv)
            {
                Profile pr = ds.Profiles.AddForSolid();
                ComponentOccurrence occ = findOcc("Крышка", dv);
                if (occ != null)
                    dv.BreakOutOperations.Add(pr, occ);

            }
            public ComponentOccurrence findOcc(string nameForSearch, DrawingView dv)
            {
                Document doc = dv.ReferencedDocumentDescriptor.ReferencedDocument as Document;
                if (doc.DocumentType != DocumentTypeEnum.kAssemblyDocumentObject) return null;
                AssemblyComponentDefinition compDef = (doc as AssemblyDocument).ComponentDefinition;
                foreach (ComponentOccurrence occ in compDef.Occurrences)
                {
                    if (occ.ReferencedDocumentDescriptor.FullDocumentName.IndexOf("(" + nameForSearch + ")") != -1)
                        return occ;
                }
                return null;
            }
        }

        internal class BreakButton : Button
        {
            public BreakButton(string displayName, string internalName, string clientId, string description, string tooltip, 
            ButtonDisplayEnum buttonDisplayType = ButtonDisplayEnum.kDisplayTextInLearningMode, CommandTypesEnum commandType = CommandTypesEnum.kNonShapeEditCmdType)
            : base(displayName, internalName, commandType, clientId, description, tooltip, buttonDisplayType) { }
            protected override void ButtonDefinition_OnExecute(NameValueMap context)
            {
                brOp();
            }
            public void brOp()
            {
                XMLDoc xmlDoc = new XMLDoc(@"C:\ProgramData\Autodesk\Inventor Addins\break.xml","head");
                List<string> lst = new List<string>();
                foreach (var el in xmlDoc.El.Elements())
	            {
		           lst.Add(el.Attribute("name").Value.ToString());
	            }
                int offsetY = 30, offsetX = 10;
                Form f = new Form();
                f.Height = 150; f.Width = 400; f.WindowState = FormWindowState.Normal; f.Text = "Разрывы"; f.StartPosition = FormStartPosition.CenterScreen;
                f.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink; f.AutoSize = true;
                System.Drawing.Point insPt = new System.Drawing.Point(5, 5);
                InterfaceDll.Lbl lbs = new InterfaceDll.Lbl(offsetX, offsetY, 100, 15, insPt, f, "Тип");
                InterfaceDll.CB cbs = new InterfaceDll.CB(offsetX, offsetY, 200, 15, insPt, f); 
                cbs.position(lbs.last(), true);
                cbs.last().Items.AddRange(lst.ToArray());
                InterfaceDll.Btn btns = new InterfaceDll.Btn(offsetX, offsetY, 100, 20, insPt, f, break_Click, "Добавить");
                btns.center(cbs.last(), offsetY + 5);
                f.Show();

               
            }

            private void break_Click(object arg1, EventArgs arg2)
            {
                Form f = (Form)((System.Windows.Forms.Button)arg1).Parent;
                string name = f.Controls.OfType<ComboBox>().ElementAt(0).Text;
                DrawingDocument drw = Macros.StandardAddInServer.m_inventorApplication.ActiveDocument as DrawingDocument;
                XMLDoc xmlDoc = new XMLDoc(@"C:\ProgramData\Autodesk\Inventor Addins\break.xml","head");
                XElement el = xmlDoc.getXElement(name, "name", "breakOp");
                DrawingView dv; double w, h, cx, cy, gap = 0.2, wb, hb, l, b, minus = 0;
                int i = 0;
                if (!el.HasElements && el.Attribute("name") != null && el.Attribute("name").Value == "по данным чертежа горизонтально")
                {
                    //SelectSet ss = drw.SelectSet;
                    List<DrawingCurveSegment> ss = drw.SelectSet.OfType<DrawingCurveSegment>().ToList();
                    if (ss.Count > 1)
                    {
                        dv = (ss[0]).Parent.Parent;
                        double mn = 2;
                        wb = dv.Width;
                        for (int j = 0; j < ss.Count-1; j++)
                        {
                            if (j == 0 || j == ss.Count - 2) mn = 1;
                            else mn = 2;
                            DrawingCurveSegment seg1 = ss[j] as DrawingCurveSegment, seg2 = ss[j+1] as DrawingCurveSegment;
                            DrawingCurve dc1 = seg1.Parent, dc2 = seg2.Parent;
                            double x1 = (dc1.Evaluator2D.RangeBox.MaxPoint.X - dc1.Evaluator2D.RangeBox.MinPoint.X),
                                x2 = (dc2.Evaluator2D.RangeBox.MaxPoint.X - dc2.Evaluator2D.RangeBox.MinPoint.X),
                                c1 = (dc1.Evaluator2D.RangeBox.MaxPoint.X + dc1.Evaluator2D.RangeBox.MinPoint.X)/2,
                                c2 = (dc2.Evaluator2D.RangeBox.MaxPoint.X + dc2.Evaluator2D.RangeBox.MinPoint.X)/2;
                            w = (x1 + x2)*mn;
                            if (w < 1.2) w = 1.2;
                            cx = (c1 + c2)/2;
                            w = (c2-c1) - w;
                            dv.BreakOperations.Add(BreakOrientationEnum.kHorizontalBreakOrientation, I.tg.CreatePoint2d((cx - minus) - w / 2, 0),
                            I.tg.CreatePoint2d((cx - minus) + w / 2, 0), BreakStyleEnum.kRectangularBreakStyle, 10, gap, 1);
                            minus = (wb - dv.Width)/2;
                        }
                    }
                    return;
                }
                else if (!el.HasElements && el.Attribute("name") != null && el.Attribute("name").Value == "по данным чертежа вертикально")
                {
                    //SelectSet ss = drw.SelectSet;
                    List<DrawingCurveSegment> ss = drw.SelectSet.OfType<DrawingCurveSegment>().ToList();
                    if (ss.Count > 1)
                    {
                        dv = (ss[0]).Parent.Parent;
                        double mn = 2;
                        hb = dv.Height;
                        for (int j = 0; j < ss.Count - 1; j++)
                        {
                            if (j == 0 || j == ss.Count - 2) mn = 1;
                            else mn = 2;
                            DrawingCurveSegment seg1 = ss[j] as DrawingCurveSegment, seg2 = ss[j + 1] as DrawingCurveSegment;
                            DrawingCurve dc1 = seg1.Parent, dc2 = seg2.Parent;
                            double y1 = (dc1.Evaluator2D.RangeBox.MaxPoint.Y - dc1.Evaluator2D.RangeBox.MinPoint.Y),
                                y2 = (dc2.Evaluator2D.RangeBox.MaxPoint.Y - dc2.Evaluator2D.RangeBox.MinPoint.Y),
                                c1 = (dc1.Evaluator2D.RangeBox.MaxPoint.Y + dc1.Evaluator2D.RangeBox.MinPoint.Y) / 2,
                                c2 = (dc2.Evaluator2D.RangeBox.MaxPoint.Y + dc2.Evaluator2D.RangeBox.MinPoint.Y) / 2;
                            h = (y1 + y2) * mn;
                            if (h < 1.2) h = 1.2;
                            cy = (c1 + c2) / 2;
                            h = (c2 - c1) - h;
                            dv.BreakOperations.Add(BreakOrientationEnum.kVerticalBreakOrientation, I.tg.CreatePoint2d(0, (cy - minus) - h / 2),
                            I.tg.CreatePoint2d(0, (cy - minus) + h / 2), BreakStyleEnum.kRectangularBreakStyle, 10, gap, 1);
                            minus = (hb - dv.Height) / 2;
                        }
                    }
                    return;
                }
                dv = drw.SelectSet[1] as DrawingView;

                wb = dv.Width; hb = dv.Height;
                l = dv.Left; b = dv.Top - dv.Height;
                minus = 0;
                
                foreach (var item in el.Elements())
                {
                    w = InvDoc.u.convToDouble(item.Attribute("W").Value);
                    h = InvDoc.u.convToDouble(item.Attribute("H").Value);
                    cx = InvDoc.u.convToDouble(item.Attribute("CX").Value);
                    cy = InvDoc.u.convToDouble(item.Attribute("CY").Value);
                    if (w != 0)
                    {
                        w *= wb;
                        cx *= wb;
                        dv.BreakOperations.Add(BreakOrientationEnum.kHorizontalBreakOrientation, I.tg.CreatePoint2d((l + cx + i*gap -minus/2) -w/2, 0),
                        I.tg.CreatePoint2d((l + cx + i*gap - minus/2) + w / 2, 0), BreakStyleEnum.kRectangularBreakStyle, 10, gap, 1);
                        minus += w;
                    }
                    else if (h != 0)
                    {
                        h *= hb;
                        cy *= hb;
                        dv.BreakOperations.Add(BreakOrientationEnum.kVerticalBreakOrientation, I.tg.CreatePoint2d(0, (b + cy + i * gap - minus / 2) - h / 2),
                        I.tg.CreatePoint2d(0, (b + cy + i * gap - minus / 2) + h / 2), BreakStyleEnum.kRectangularBreakStyle, 10, gap, 1);
                        minus += h;
                    }
                    i=1;
                }
            }
        }

        public enum orient { horizontal, vertical, none }
        public enum action { add, remove, none}

        abstract public class DrwBase
        {

            abstract public void set();
            abstract public void add();
        }

        public class DrwArr
        {
            List<ArrDim> dims = new List<ArrDim>();
            List<ArrDim> draw = new List<ArrDim>();
            //List<gab> gabs = new List<gab>();
            //gab gv = null, gh = null;
            ArrDim origin;
/*            double min;*/
            DrawingView dv;
            int num = 1;
            public DrwArr(DrawingCurve dc, Func<DrawingCurve,bool> check)
            {
                if (dc == null) return;
                dv = dc.Parent;
                origin = new ArrDim(dc);
                dims.Add(origin);
                //draw.Add(origin);
//                 if (dc.CurveType == CurveTypeEnum.kCircularArcCurve)
//                 {
//                     ObjectsEnumerator col = dc.Parent.Parent.FindUsingPoint(dc.StartPoint);
//                     foreach (DrawingCurveSegment item in col)
// 	                {
//                         if (!item.Parent.Equals(dc))
//                         {
//                             ArrDim.dist = util.getLenght(item.Parent);
//                             ArrDim.v = (util.eq(item.StartPoint.X, item.EndPoint.X)) ? orient.vertical : orient.horizontal;
//                         }
// 	                }
//                 }
                u.action<DrawingCurve>(u.gets<DrawingCurve>(dv.DrawingCurves, check), a => { 
                    if (!a.Equals(dc) && this.check(a,dc)) 
                        dims.Add(new ArrDim(a));
                });
            }
            
            public bool check(DrawingCurve a, DrawingCurve b)
            {
                if (a.CurveType == CurveTypeEnum.kCircleCurve) return true;
                Vector2d s = a.StartPoint.VectorTo(a.EndPoint), e = b.StartPoint.VectorTo(b.EndPoint);
                e.SubtractVector(s);
                return u.isNullVector(e);
            }
            public void remove()
            {
                u.action<ArrDim>(dims, a => a.rem = action.add, f => f.check(origin));
                array();      
                check();
                //if (ArrDim.dist != 0 && ArrDim.o == ArrDim.v) check();
            }
            public void add()
            {
                gabs.draw(origin.get(), I.CV2d(1));
                ArrDim.min = gabs.min;
                ArrDim.direct = gabs.pt; /*I.CV2d(1);*/
                ArrDim.set();
                remove();
                add(num);

                draw.Clear(); //draw.Add(origin);
                u.action<ArrDim>(dims, a => a.clear());
                num = 1;
                gabs.draw(origin.get(), I.CV2d(0, 1));
                ArrDim.min = gabs.min;
                ArrDim.direct = gabs.pt;/*I.CV2d(0, 1);*/
                ArrDim.set();
                remove();
                add(num);
            }

            void add(int num)
            {
                if (num <= 2)
                {
                    double center = 1;
                    if (gabs.centerX && ArrDim.direct.Y == 0) center = 0.5;
                    else if (gabs.centerY && ArrDim.direct.X == 0) center = 0.5;
                    u.actionFor<ArrDim>(draw, (a, b) => a.add(dv, b),center: center);
                }
                else
                {
                    LinearGeneralDimension m, d;
                    if (draw[0].Equals(origin))
                    m = draw[0].add(dv, draw[1]);
                    else m = draw[draw.Count-2].add(dv, draw[draw.Count-1]);
                    d = draw[0].add(dv, draw[draw.Count - 1], 1.5);
                    d.Text.FormattedText = m.Text.Text + "x" + (draw.Count - 1) + "=" + d.Text.FormattedText;
                }
            }

            void array()
            {
                draw = u.actionFor<ArrDim>(dims, (a, b) => v(), f => f.rem == action.add);
                draw.Sort();
                u.actionFor<ArrDim>(draw, (a, b) => setL(a,b)); 
            }

            void setL(ArrDim a, ArrDim b)
            {
                a.l = a.pt.DistanceTo(b.pt);
//                 if (num == 1) a.l = ArrDim.direct.Y == 0 ? Math.Abs(b.x - a.x) : Math.Abs(b.y - a.y); 
//                 b.l = ArrDim.direct.X == 0 ? Math.Abs(b.x - a.x) : Math.Abs(b.y - a.y);
            }

            void v() { }

            public void check(ArrDim a, ArrDim b)
            {
                //Point2d p1 = I.tg.CreatePoint2d(a.x, a.y), p2 = I.tg.CreatePoint2d(b.x, b.y);
                if (a.l != 0 && u.eq(a.l, b.l)) 
                {
                    num++; //b.rem = true; 
                };
                //if (util.eq(p1.DistanceTo(p2), ArrDim.dist)) b.rem = true;
            }

            public void check()
            {
                u.actionFor<ArrDim>(draw, (a, b) => check(a,b)); 
            }
        }

        public class ArrDim : IComparable<ArrDim> 
        {
            DrawingCurve dc;
            readonly public double x,y;
            readonly public Point2d pt;
            public action rem = action.none;
            static public double min = 0;
            static public Vector2d direct;
            static public double dist = 0;
            static public orient v = orient.none;
            public double l = 0;
            static public DimensionTypeEnum al = DimensionTypeEnum.kAlignedDimensionType;
            public ArrDim (DrawingCurve dc)
            {
                this.dc = dc; pt = dc.CenterPoint;
                x = dc.CenterPoint.X; y = dc.CenterPoint.Y;
            }

            static public void unset()
            {
                min = 0; direct = null; dist = 0; v = orient.none; al = DimensionTypeEnum.kAlignedDimensionType;
            }

            public bool check(ArrDim d)
            {
                bool r = false;
                if(u.eq(I.multiPV(pt, direct), I.multiPV(d.pt, direct))) r = !r; 
//                 switch (o)
//                 {
//                     case orient.horizontal:
//                         if (util.eq(d.y, y)) r = !r;
//                         break;
//                     case orient.vertical:
//                         if (util.eq(d.x, x)) r = !r;
//                         break;
//                     default:
//                         break;
//                 }
                return r;
            }
            public static void set()
            {
                if (direct.X == 0)
                {
                    al = DimensionTypeEnum.kHorizontalDimensionType;
                }
                else if (direct.Y == 0)
                {
                    al = DimensionTypeEnum.kVerticalDimensionType;
                }
            }
            public void clear()
            {
                rem = action.none;
            }
            public DrawingCurve get()
            {
                return dc;
            }
            public LinearGeneralDimension add(DrawingView dv ,ArrDim d, double offset = 1)
            {
                return u.addDim<LinearGeneralDimension>(dv, dc, d.dc, offset, al, min, direct.X + direct.Y);
            }
            public LinearGeneralDimension add(DrawingView dv, DrawingCurve d, double min, int direct, double offset = 1)
            {
                return u.addDim<LinearGeneralDimension>(dv, dc, d, offset, al, min, direct);
            }

//             public override bool Equals(object obj)
//             {
//                 return util.eq(pt, (obj as ArrDim).pt);
//             }

            public int CompareTo(ArrDim other)
            {
                if (rem != action.add) return 1;
                if (direct.X == 0) return (this.x - other.x) * direct.Y > 0 ? -1 : 1;
                else if (direct.Y == 0) return (this.y - other.y) * direct.X > 0 ? -1 : 1;
                    else return -1;
//                 if (util.eq(this.x, other.x)) return this.y < other.y ? -1 : util.eq(this.y, other.y) ? 0 : 1;
//                 return this.x < other.x ? -1 : 1;
            }
        }


        static public class gabs 
        {
            static public List<gab> gbs = new List<gab>();
            static public DrawingView dv;
            static public double l, r, t, b;
            static public Point2d c;
            static public bool centerX = false, centerY = false;
            static public double min = 0;
            static public Vector2d pt;
            static public bool exist = false;

            static public void set()
            {
                if (!gabs.exist)
                {
                    IEnumerable<DrawingCurve> linesX, linesY; double minLenght = 3;
                    linesX = u.gets<DrawingCurve>(dv.DrawingCurves, f => check(f));
                    linesY = u.gets<DrawingCurve>(dv.DrawingCurves, f => check(f, false));
                    //linesX = linesX.OrderBy(e => e.StartPoint.X); linesY = linesY.OrderBy(e => e.StartPoint.Y);
                    DrawingCurve minx = linesX.FirstOrDefault(e => e.StartPoint.X == linesX.Min(m => m.StartPoint.X)),
                        miny = linesY.FirstOrDefault(e => e.StartPoint.Y == linesY.Min(m => m.StartPoint.Y)),
                        maxx = linesX.FirstOrDefault(e => e.StartPoint.X == linesX.Max(m => m.StartPoint.X)),
                        maxy = linesY.FirstOrDefault(e => e.StartPoint.Y == linesY.Max(m => m.StartPoint.Y));

                    new gab(minx, I.CV2d(-1));
                    new gab(maxx, I.CV2d(1));
                    new gab(miny, I.CV2d(0, -1));
                    new gab(maxy, I.CV2d(0, 1));
                }
            }
            static bool check(DrawingCurve dc, bool x = true)
            {
                if (dc.CurveType != CurveTypeEnum.kLineSegmentCurve) return false;
                if (dc.EdgeType != DrawingEdgeTypeEnum.kUnknownEdge) return false;
                if (x && !u.eq(dc.StartPoint.X, dc.EndPoint.X)) return false;
                if (!x && !u.eq(dc.StartPoint.Y, dc.EndPoint.Y)) return false;
                return true;
            }
            public static void add(gab g)
            {
                gabs.gbs.Add(g);
            }
            public static void add(DrawingView d)
            {
                if (dv != null && dv.Equals(d)) { exist = true; return; }
                gabs.dv = d;
                l = dv.Left; r = dv.Left + dv.Width; t = dv.Top; b = dv.Top - dv.Height;
                c = dv.Position;
                center();
            }
            static public void unset()
            {
                gbs.Clear(); dv = null; min = 0; exist = false; pt = null;
            }
            static public void center()
            {
                Sheet sh = dv.Parent;
                ObjectsEnumerator col = sh.FindUsingPoint(c);
                foreach (var item in col)
	            {
		            if (!(item is Centerline)) return;
                    Centerline l = item as Centerline;
                    Vector2d v = l.StartPoint.VectorTo(l.EndPoint);     
                    if (u.eq(v.X, 0) && v.Length > dv.Height)
                    {
                       centerY = true;
                    }
                    else if (u.eq(v.Y , 0) && v.Length > dv.Width)
                    {
                       centerX = true;
                    }
	            }
            }
            static public LinearGeneralDimension draw(DrawingCurve dc, Vector2d v, double o = 1)
            {
                pt = I.CV2d((dc.CenterPoint.X - c.X), (dc.CenterPoint.Y - c.Y));
                gab g = check(pt, v);
                DimensionTypeEnum al = v.X == 0 ? DimensionTypeEnum.kVerticalDimensionType : DimensionTypeEnum.kHorizontalDimensionType;
                
                if (al == DimensionTypeEnum.kVerticalDimensionType) min = pt.X < 0 ? l : r;
                else if (al == DimensionTypeEnum.kHorizontalDimensionType) min = pt.Y < 0 ? b : t;
                if (g == null) return null;
                pt = I.CV2d(pt.X * v.Y, pt.Y * v.X);
                pt.Normalize();
                LinearGeneralDimension dim = u.addDim<LinearGeneralDimension>(dv, dc, g.dc, o, al, min, pt.X + pt.Y);
                if (g.p == poz.bottom || g.p == poz.left) offsetN(dim, -1);
                else offsetN(dim);
                return dim;
            }
            static public void offsetN(LinearGeneralDimension dim, int dir = 1)
            {
                Vector2d v = dim.Text.RangeBox.MinPoint.VectorTo(dim.Text.RangeBox.MaxPoint);
                LineSegment2d dl = dim.DimensionLine as LineSegment2d;
                UnitVector2d d = dl.Direction;
                if (v.Length > dl.StartPoint.VectorTo(dl.EndPoint).Length)
                {
                    dim.Text.Origin = I.CP2d(Math.Abs(d.X) * dir * v.Length + dim.Text.Origin.X, Math.Abs(d.Y) * dir * v.Length + dim.Text.Origin.Y);
                }
            }
            static public gab get(poz pz)
            {
                return gbs.FirstOrDefault(e => e.p == pz);
            }
            static public gab check(Vector2d pt, Vector2d d)
            {
                Point2d p = I.CP2d(pt.X * d.X, pt.Y * d.Y);
                if (p.X == 0) return p.Y < 0 ? gabs.get(poz.bottom) : gabs.get(poz.top);
                if (p.Y == 0) return p.X < 0 ? gabs.get(poz.left) : gabs.get(poz.right);
                return null;
            }
        }

        public class gab
        {
            public DrawingCurve dc = null;
            public poz p;
            public Vector2d direct;
            public double min;
            public gab(DrawingCurve c, Vector2d direct)
            {
                dc = c; this.direct = direct;
                if (direct.X == 0 && direct.Y < 0) p = poz.bottom;
                else if (direct.X == 0 && direct.Y > 0) p = poz.top;
                else if (direct.Y == 0 && direct.X < 0) p = poz.left;
                else if (direct.Y == 0 && direct.X > 0) p = poz.right;
                gabs.add(this);
            }

        }

        public class CheckReflect
        {
            public List<DrawingCurve> curves = new List<DrawingCurve>();
            public List<DrawingCurve> curvesReflect = new List<DrawingCurve>();
            public List<DrawingCurve> curvesIntent = new List<DrawingCurve>();
            Point2d center;
            DrawingView dv;
            Vector2d vs, direct;

            public CheckReflect(DrawingView dv, Vector2d d)
            {
                this.dv = dv; center = dv.Center; direct = d;
                curves = getCurves(); d.ScaleBy(-1);
                curvesReflect = getCurves();
                curvesIntent = u.gets<DrawingCurve>(dv.DrawingCurves, cur => check(cur)).ToList();
                if (curves.Count != curvesReflect.Count) return;
                if (!check(CurveTypeEnum.kCircleCurve)) return; 
                if (!check(CurveTypeEnum.kLineSegmentCurve)) return;
                if (curvesIntent.Count > 1)
                {
                    sort(curvesIntent);
                    DrawingCurve min = curvesIntent[0], max = curvesIntent[curvesIntent.Count - 1];
                    u.addCenterLine(dv, min, max);
                }
            }

            public void sort(List<DrawingCurve> curves)
            {
                if (direct.X == 0)
                {
                    curves.Sort(new CurveComparerX());
                }
                else if (direct.Y == 0)
                {
                    curves.Sort(new CurveComparerY());
                }
            }

            public void sort(List<DrawingCurve> curves, List<DrawingCurve> curvesReflect)
            {
                sort(curves); sort(curvesReflect);
            }

            public bool check(DrawingCurve c)
            {
                if (!(c.EdgeType == DrawingEdgeTypeEnum.kUnknownEdge && c.CurveType == CurveTypeEnum.kLineSegmentCurve)) return false;
                if (curves.Exists(e => e.Equals(c)) || curvesReflect.Exists(e => e.Equals(c))) return false;
                return true;
            }
            public List<DrawingCurve> getCurves()
            {
                return u.gets<DrawingCurve>(dv.DrawingCurves, c =>
                {
                    return get(c); 
                }).ToList();
            }
            public bool get(DrawingCurve c)
            {
                if (c.EdgeType != DrawingEdgeTypeEnum.kUnknownEdge) return false;
                if (c.CurveType == CurveTypeEnum.kCircleCurve || c.CurveType == CurveTypeEnum.kCircularArcCurve)
                {
                    return get(c.CenterPoint);
                }
                else if (c.CurveType == CurveTypeEnum.kLineSegmentCurve)
                {
                    return (get(c.StartPoint) && get(c.EndPoint));
                }
                return false;
            }
            public bool get(Point2d pt)
            {
                vs = pt.VectorTo(dv.Center);
                if (vs.DotProduct(direct) > 0) return true;
                return false;
            }
            public bool check(CurveTypeEnum typ)
            {
                List<DrawingCurve> cL = u.gets<DrawingCurve>(curves, c => c.CurveType == typ).ToList();
                List<DrawingCurve> cLR = u.gets<DrawingCurve>(curvesReflect, c => c.CurveType == typ).ToList();
                if (cL.Count != cLR.Count) return false;
                sort(cL, cLR);
                for (int i = 0; i < cL.Count; i++)
                {
                    if (!check(cL[i], cLR[i])) return false; 
                }
                return true;
            }
            public bool check(Point2d a, Point2d b)
            {
                Vector2d x = a.VectorTo(center), y = b.VectorTo(center);
                if (u.eq(0.1,-x.DotProduct(direct),y.DotProduct(direct))) return true;
                return false;
            }
            public bool check(DrawingCurve a, DrawingCurve b)
            {
                if ((a.CurveType == CurveTypeEnum.kCircleCurve && b.CurveType == CurveTypeEnum.kCircleCurve) ||
                    (a.CurveType == CurveTypeEnum.kCircularArcCurve && b.CurveType == CurveTypeEnum.kCircularArcCurve))
                {
                    return check(a.CenterPoint, b.CenterPoint);
                }
                else if (a.CurveType == CurveTypeEnum.kLineSegmentCurve && b.CurveType == CurveTypeEnum.kLineSegmentCurve)
                {
                    return (check(a.MidPoint, b.MidPoint));
                }
                return false;
            }
        }

        public class CurveComparerX : IComparer<DrawingCurve>
        {
            public int Compare(DrawingCurve x, DrawingCurve y)
            {
                Point2d mpx = null, mpy = null;
                if (x.CurveType == CurveTypeEnum.kCircleCurve || x.CurveType == CurveTypeEnum.kCircularArcCurve) mpx = x.CenterPoint;
                else if (x.CurveType == CurveTypeEnum.kLineSegmentCurve) mpx = x.MidPoint;
                if (y.CurveType == CurveTypeEnum.kCircleCurve || y.CurveType == CurveTypeEnum.kCircularArcCurve) mpy = y.CenterPoint;
                else if (y.CurveType == CurveTypeEnum.kLineSegmentCurve) mpy = y.MidPoint;

                return u.eq(mpx.X, mpy.X) ? 0 : mpx.X < mpy.X ? -1 : 
                    mpx.X == mpy.X ? mpx.Y <= mpy.Y ? -1: 1 : 1;
            }
        }

        public class CurveComparerY : IComparer<DrawingCurve>
        {
            public int Compare(DrawingCurve x, DrawingCurve y)
            {
                Point2d mpx = null, mpy = null;
                if (x.CurveType == CurveTypeEnum.kCircleCurve || x.CurveType == CurveTypeEnum.kCircularArcCurve) mpx = x.CenterPoint;
                else if (x.CurveType == CurveTypeEnum.kLineSegmentCurve) mpx = x.MidPoint;
                if (y.CurveType == CurveTypeEnum.kCircleCurve || y.CurveType == CurveTypeEnum.kCircularArcCurve) mpy = y.CenterPoint;
                else if (y.CurveType == CurveTypeEnum.kLineSegmentCurve) mpy = y.MidPoint;

                return u.eq(mpx.Y, mpy.Y) ? 0 : mpx.Y < mpy.Y ? -1 :
                    mpx.Y == mpy.Y ? mpx.X <= mpy.X ? -1 : 1 : 1;
            }
        }

        internal class DrwArrBtn : Button
        {
            DrawingDocument drw;
            DrawingCurveSegment dcs = null;
            DrawingCurve dc = null;
            DrawingView dv = null;
            public DrwArrBtn(string displayName, string internalName, string clientId, string description, string tooltip,
            ButtonDisplayEnum buttonDisplayType = ButtonDisplayEnum.kDisplayTextInLearningMode, CommandTypesEnum commandType = CommandTypesEnum.kNonShapeEditCmdType)
                : base(displayName, internalName, commandType, clientId, description, tooltip, buttonDisplayType) { }
            protected override void ButtonDefinition_OnExecute(NameValueMap context)
            {
                drw = (DrawingDocument)Macros.StandardAddInServer.m_inventorApplication.ActiveDocument;
                u.transactStart(drw as Document, "Массив размеры");
                
//                 if (drw.SelectSet.Count == 0)
//                     dcs = (DrawingCurveSegment)Macros.StandardAddInServer.m_inventorApplication.CommandManager.Pick(SelectionFilterEnum.kDrawingCurveSegmentFilter, "Выберите отверстие");
//                 else if (drw.SelectSet[1] is DrawingCurveSegment)
//                 {
//                     dcs = (DrawingCurveSegment)drw.SelectSet[1];
//                 }
//                 add(dcs);
                u.transactEnd();
            }
            private void add(DrawingCurveSegment dcs)
            {
                if (dcs != null)
                {
                    dc = dcs.Parent as DrawingCurve;
                    dv = dc.Parent as DrawingView;
                    List<DrawingCurve> dcX, dcY;
                    LinearGeneralDimension dim;
                    string oldText = ""; double dist = 0, distX = 0, distY = 0;
                    u.getCurvesToArray(dv, dc, out dcX, out dcY);
                    Point2d retPt = null; Vector2d vx = null, vy = null; double x = 1, y = 1;
                    object ob = u.findAtPoint(dv.Parent, dc.CenterPoint, ref retPt, ref distY, 25, true), obx = null, oby = null;
                    //ob = util.findAtPoint(dv.Parent, dc.CenterPoint, ref retPt, ref distX, 25, false);
                    obx = u.findAtVector(dv, dc.CenterPoint, ref retPt, ref distX, ref vy, 2, false);
                    oby = u.findAtVector(dv, dc.CenterPoint, ref retPt, ref distY, ref vx, 2, true);
                    if (obx != null)
                    {
                        ob = (obx as DrawingCurveSegment).Parent;
                        vy.Normalize(); vy.ScaleBy(x);
                        u.addDim<LinearGeneralDimension>(dv, dc, (ob as DrawingCurve)); 
                        //util.addDim(dv, dc, ob, PointIntentEnum.kCenterPointIntent, null, dc.CenterPoint, retPt, util.sumDist(-1, distY), align: DimensionTypeEnum.kVerticalDimensionType);
                    }
                    //ob = util.findAtPoint(dv.Parent, dc.CenterPoint, ref retPt, ref distY, 25, true);
                    if (oby != null)
                    {
                        ob = (oby as DrawingCurveSegment).Parent;
                        vx.Normalize(); vx.ScaleBy(y);
                        u.addDim<LinearGeneralDimension>(dv, dc, ((ob as DrawingCurve))/*, vx, retPt*/);
                        //util.addDim(dv, dc, ob, PointIntentEnum.kCenterPointIntent, null, dc.CenterPoint, retPt, -util.sumDist(-1, distX), align: DimensionTypeEnum.kHorizontalDimensionType);
                    }

                    if (dcX.Count > 1)
                    {
                        dcX = dcX.OrderBy(e => e.CenterPoint.X).ToList();
                        dim = u.addDim(dv, dcX[0], dcX[dcX.Count - 1], PointIntentEnum.kCenterPointIntent, PointIntentEnum.kCenterPointIntent,
                            dcX[0].CenterPoint, dcX[dcX.Count - 1].CenterPoint, u.sumDist(-1, distX), DimensionTypeEnum.kHorizontalDimensionType);
                        if (dcX.Count > 2)
                        {
                            dist = dcX[0].CenterPoint.DistanceTo(dcX[1].CenterPoint);
                            oldText = dim.Text.FormattedText;
                            dim.Text.FormattedText = dcX.Count - 1 + "x" + (dist * 10 / dv.Scale).ToString("#.#") + "=" + oldText;
                        }
                        else dim.Delete();
                    }
                    if (dcY.Count > 1)
                    {
                        dcY = dcY.OrderBy(e => e.CenterPoint.Y).ToList();
                        dim = u.addDim(dv, dcY[0], dcY[dcY.Count - 1], PointIntentEnum.kCenterPointIntent, PointIntentEnum.kCenterPointIntent,
                            dcY[0].CenterPoint, dcY[dcY.Count - 1].CenterPoint, -u.sumDist(-1, distY), DimensionTypeEnum.kVerticalDimensionType);
                        if (dcY.Count > 2)
                        {
                            dist = dcY[0].CenterPoint.DistanceTo(dcY[1].CenterPoint);
                            oldText = dim.Text.FormattedText;
                            dim.Text.FormattedText = dcY.Count - 1 + "x" + (dist * 10 / dv.Scale).ToString("#.#") + "=" + oldText;
                        }
                        else dim.Delete();
                    }
                }
            }
        }

        public enum poz { left, right, bottom, top}

    public class drawingCurves
    {
        public Point2d min, max;
        public bool rev = false, vert;
        public double minX = 10000, minY = 10000, maxX = -10000, maxY = -10000;
        public DrawingCurve curve;
        public void set(DrawingCurve c, Point2d a, Point2d i, double ix, double iy, double ax, double ay)
        {
            curve = c; min = i; max = a; minX = ix; minY = iy; maxX = ax; maxY = ay; /*rev = r;*/
        }
        public void set(DrawingCurveSegment seg)
        {
            this.curve = seg.Parent;
            DrawingCurve item = curve;
            if (item.ProjectedCurveType == Curve2dTypeEnum.kLineSegmentCurve2d && item.CurveType == CurveTypeEnum.kLineSegmentCurve)
            {
                min = item.Evaluator2D.RangeBox.MinPoint;
                max = item.Evaluator2D.RangeBox.MaxPoint;
                minX = min.X; minY = min.Y;
                maxX = max.X; maxY = max.Y;
                if (item.StartPoint.X == item.EndPoint.X)
                {
                    if (item.StartPoint.Y < item.EndPoint.Y)
                    {
                        //cur.rev = true;
                        min = item.StartPoint; max = item.EndPoint;
                    }
                    else
                    {
                        //cur.rev = false;
                        min = item.EndPoint; max = item.StartPoint;
                    }
                }
                else if (item.StartPoint.X < item.EndPoint.X)
                {
                    //cur.rev = true;
                    min = item.StartPoint; max = item.EndPoint;
                }
                else
                {
                    //cur.rev = false;
                    min = item.EndPoint; max = item.StartPoint;
                }
            }
            else if (item.ProjectedCurveType == Curve2dTypeEnum.kCircularArcCurve2d)
            {
                double r = ((Arc2d)(item.Segments[1].Geometry)).Radius; Point2d cen = ((Arc2d)(item.Segments[1].Geometry)).Center;
                double[] pts = new double[4];
                pts = Dimensions.intersect(item); 
                minX = (pts[0] != 0) ? cen.X - r : (item.StartPoint.X < item.EndPoint.X) ? item.StartPoint.X : item.EndPoint.X;
                minY = (pts[1] != 0) ? cen.Y - r : (item.StartPoint.Y < item.EndPoint.Y) ? item.StartPoint.Y : item.EndPoint.Y;
                maxX = (pts[2] != 0) ? cen.X + r : (item.StartPoint.X > item.EndPoint.X) ? item.StartPoint.X : item.EndPoint.X;
                maxY = (pts[3] != 0) ? cen.Y + r : (item.StartPoint.Y > item.EndPoint.Y) ? item.StartPoint.Y : item.EndPoint.Y;
                min = I.tg.CreatePoint2d(minX, minY);
                max = I.tg.CreatePoint2d(maxX, maxY);
            }
        }
        public void setRev(bool X = true)
        {
            if (curve == null) return;
            Vector2d vec = curve.StartPoint.VectorTo(curve.EndPoint);
            Vector2d axis;
            if (X)
            {
               axis = I.tg.CreateVector2d(1,0);
            }
            else axis = I.tg.CreateVector2d(0, 1);
            double angle = vec.AngleTo(axis);
            if (Math.Round(Math.Abs(angle), 4) == Math.Round(Math.PI/2, 4))
            {
                setRev(!X);
            }
            if (Math.Abs(angle) > Math.PI / 2) rev = true;
        }
        public void setVert()
        {
            if (u.eq(curve.StartPoint.X, curve.EndPoint.X)) vert = true;
            else vert = false;
        }
        public void chCurve(poz p)
        {
            Sheet sh = curve.Parent.Parent;
            ObjectsEnumerator col; double tol = 0.01; DrawingCurve c = null;
            Point2d pt = null;
            setVert();
            switch (p)
	        {
		        case poz.left:
                    pt = (rev)? curve.EndPoint: curve.StartPoint;
                    if (vert) return;
                 break;
                case poz.right:
                    pt = (rev)? curve.StartPoint: curve.EndPoint;
                    if (vert) return;
                 break;
                case poz.bottom:
                    pt = (rev)? curve.EndPoint: curve.StartPoint;
                    if (!vert) return;
                 break;
                case poz.top:
                    pt = (rev)? curve.StartPoint: curve.EndPoint;
                    if (!vert) return;
                 break;
                default:
                 break;
	        }
            col = sh.FindUsingPoint(pt, tol);
            if (col != null || col.Count > 1)
            {
                foreach (DrawingCurveSegment seg in col)
                {
                    c = seg.Parent;
                    if (p == poz.left || p == poz.right)
                    {
                        if (u.eq(c.StartPoint.Y, c.EndPoint.Y)) { curve = c; }
                    }
                    else
                    {
                        if (u.eq(c.StartPoint.X, c.EndPoint.X)) { curve = c; }
                    }
                }
            }
            setVert();
        }
    }

    public class OrderDims
    {
        List<Order> dims = new List<Order>();
        DrawingView dv;
        readonly poz direction;
        double min;
        const double offset = 0.7;
        
        public OrderDims(DrawingView dv, IEnumerable<LinearGeneralDimension> d, poz direction)
        {
            this.direction = direction;
            this.dv = dv;
            switch (direction)
            {
                case poz.left:
                    Order.direction = -1;
                    min = dv.Center.X - dv.Width/2 - offset;
                    break;
                case poz.right:
                    Order.direction = 1;
                    min = dv.Center.X + dv.Width / 2 - offset / 2;
                    break;
                case poz.bottom:
                    Order.direction = -1;
                    min = dv.Center.Y - dv.Height / 2 + offset/2;
                    break;
                case poz.top:
                    Order.direction = 1;
                    min = dv.Center.Y + dv.Height / 2 + offset;
                    break;
                default:
                    break;
            }
            if (dv == null) return;

            foreach (var item in d)
            {
                switch (direction)
                {
                    case poz.left:
                        if (item.Text.Origin.X < min && (item.DimensionLine as LineSegment2d).Direction.X == 0) dims.Add(new Order(item, direction));
                        break;
                    case poz.right:
                        if (item.Text.Origin.X > min && (item.DimensionLine as LineSegment2d).Direction.X == 0) dims.Add(new Order(item, direction));
                        break;
                    case poz.bottom:
                        if (item.Text.Origin.Y < min && (item.DimensionLine as LineSegment2d).Direction.Y == 0) dims.Add(new Order(item, direction));
                        break;
                    case poz.top:
                        if (item.Text.Origin.Y > min && (item.DimensionLine as LineSegment2d).Direction.Y == 0) dims.Add(new Order(item, direction));
                        break;
                    default:
                        break;
                } 
            }
            dims.Sort();
            setMinPt();
        }
        public void setMinPt()
        {
            Order.min = min;
            //Order.minV = dims.Where(d => d.al == DimensionAlignmentTypeEnum.kHorizontalAlignmentType).Min(p => p.origin.Y);
            //Order.minH = dims.Where(d => d.al == DimensionAlignmentTypeEnum.kVerticalAlignmentType).Min(p => p.origin.X);
        }
        public void add()
        {
            if (dims.Count == 0) return;
            while (dims.Count > 1)
            {
                for (int i = 0; i < dims.Count; i++)
                {
                    dims[i].check(dims);
                }
                for (int i = 0; i < dims.Count; i++)
                {
                    if (dims[i].create)
                    {
                        dims[i].setOrigin();
                        dims.Remove(dims[i]);
                    }
                }

            }
            dims[0].setOrigin();


//             foreach (var item in dims)
//             {
//                 item.draw();
//             }
        }
    }

    public class Order: IComparable<Order>
    {
        public Box2d box;
        public DimensionAlignmentTypeEnum al = DimensionAlignmentTypeEnum.kDefaultAlignmentType;
        public double l = 0;
        LineSegment2d dl;
        Point2d pt;
        public double offsetN = 0, b = 0.35;
        readonly double offset = 0.85, tboffset = 0;
        static public int direction = 1;
        public int level = 1;
        public Point2d origin;
        double tol = 0.01;
        LinearGeneralDimension dim;
        public bool create = false;
        public double x1, x2;
        static public double min;
        public Order(LinearGeneralDimension dim, poz p = poz.bottom)
        {
            this.dim = dim;
//             el1 = dim.ExtensionLineOne as LineSegment2d;
//             el2 = dim.ExtensionLineTwo as LineSegment2d;
            dl = dim.DimensionLine as LineSegment2d;
            al = dl.Direction.Y == 0 ? DimensionAlignmentTypeEnum.kHorizontalAlignmentType :
                dl.Direction.X == 0 ? DimensionAlignmentTypeEnum.kVerticalAlignmentType : 
                DimensionAlignmentTypeEnum.kAlignedAlignmentType;
            if (check(dl, dim.Text.Origin))
            {
                //Point2d mp = util.midPt(dl.StartPoint, dl.EndPoint);
                Point2d sp = dim.Text.Origin;
                dim.CenterText();
                if (al == DimensionAlignmentTypeEnum.kHorizontalAlignmentType) 
                { 
                    offsetN = - dim.Text.Origin.X + sp.X;
                }
                else if (al == DimensionAlignmentTypeEnum.kVerticalAlignmentType) 
                {
                    offsetN = - dim.Text.Origin.Y + sp.Y;
                }
            }
            box = dim.Text.RangeBox;
            origin = dim.Text.Origin;
            tol *= (dl.Direction.X + dl.Direction.Y);
            double o = 0;
            switch (al)
            {
                case DimensionAlignmentTypeEnum.kHorizontalAlignmentType:
                    o = box.MaxPoint.Y - box.MinPoint.Y;
                    if (tol > 0)
                        box = I.Box(dl.StartPoint.X + tol, 1, dl.EndPoint.X - tol, 0);
                    else
                        box = I.Box(dl.EndPoint.X - tol, 0, dl.StartPoint.X + tol, 1);
                    x1 = box.MinPoint.X; x2 = box.MaxPoint.X; 
                    l = x2 - x1;
                    break;
                case DimensionAlignmentTypeEnum.kVerticalAlignmentType:
                    o = box.MaxPoint.X - box.MinPoint.X;
                    if (tol > 0)
                        box = I.Box(1, dl.StartPoint.Y + tol, 0, dl.EndPoint.Y-tol);
                    else
                        box = I.Box(0, dl.EndPoint.Y - tol,1, dl.StartPoint.Y + tol);
                    x1 = box.MinPoint.Y; x2 = box.MaxPoint.Y; 
                    l = x2 - x1;
                    break;
                default:
                    break;
            }
            if (!u.eq(o, b))
                tboffset += (o - b) / 2;
            if (dim.Text.Text.IndexOfAny(new char[] { '*', 'x' }) != -1 || dim.Style.Name.IndexOfAny(new char[] { '*' }) != -1 || dim.Style.Name == "Гибы")
            {
                tboffset += -0.142;
                //tboffset = offset / 2;
            }
//             else tboffset = -offset / 2;
            //if (p == poz.left || p == poz.top) leftTop -= offset/2;
        }

        public bool check(LineSegment2d dl, Point2d c)
        {
            if (dl.Direction.X == 0) return ((dl.StartPoint.Y - c.Y) * (dl.EndPoint.Y - c.Y) > 0) ? true : false;
            else return ((dl.StartPoint.X - c.X) * (dl.EndPoint.X - c.X) > 0) ? true : false;
        }

        public bool check(List<Order> vals)
        {
            foreach (var item in vals)
            {
                if (item.Equals(this)) continue;
                if (level <= item.level && !box.IsDisjoint(item.box) //&& l > item.l)
                    && (x1 <= item.x1 && x2 >= item.x2))
                {
                    setLevel();
                    return false;
                }
            }
            create = true;
            return true;
        }

        public void draw()
        {
            DrawingSketch ds = dim.Parent.Sketches.Add();
            ds.Edit();
            ds.SketchLines.AddAsTwoPointRectangle(box.MinPoint, box.MaxPoint);
            ds.ExitEdit();
        }

        public void setLevel()
        {
            level++;
        }

        public void setOrigin()
        {

            dim.Text.Origin = al == DimensionAlignmentTypeEnum.kHorizontalAlignmentType ? I.tg.CreatePoint2d(dim.Text.Origin.X, min + (level) * offset * direction + tboffset) :
                I.tg.CreatePoint2d(min + (level) * offset * direction + tboffset, dim.Text.Origin.Y);

            dim.Text.Origin = al == DimensionAlignmentTypeEnum.kHorizontalAlignmentType ? I.tg.CreatePoint2d(dim.Text.Origin.X + offsetN, dim.Text.Origin.Y) :
                I.tg.CreatePoint2d(dim.Text.Origin.X, dim.Text.Origin.Y + offsetN);
            //draw();
        }

        public int CompareTo(Order other)
        {
            return u.eq(this.l, other.l) ? 0 : this.l < other.l ? 1 : 0; 
        }
    }

    public class Dimensions
    {
        private DrawingView dv;
        Point2d rangeMin, rangeMax, cenLeft, cenRight, cenTop, cenBottom;
        Inventor.Box2d box;
        bool leftCurv = true, rightCurv = true, topCurv = true, bottomCurv = true;
        public drawingCurves leftDC = new drawingCurves(), rightDC = new drawingCurves(), bottomDC = new drawingCurves(), topDC = new drawingCurves();
        public LinearGeneralDimension leftDim, rightDim, bottomDim, topDim;
        private TransientGeometry tg = I.tg;
        private ObjectsEnumerator objs;
        private UnitVector2d uvecV, uvecH;
        double r, t;
        int startInt = 1;
        DimensionStyle st1, st2;
        public Dimensions(DrawingView drwView)
        {
            u.transAct(drwView.Parent.Parent, "Размеры");
            dv = drwView;

            if (u.isFirstSheet(dv, 1) && u.isTDoc<PartDocument>(dv.ReferencedDocumentDescriptor.ReferencedDocument) != null)
            {
                st1 = u.getStyle<DimensionStyle>(dv, "Текст*");
                st2 = u.getStyle<DimensionStyle>(dv, "Текст**");
                if (st1 != null && st2 != null)
                {
                    u.setSilence();
                    u.setUpdate();
                    tangent();
                    radius();
                    u.setSilence();
                    u.setUpdate();
                }
            }
            uvecV = tg.CreateUnitVector2d(0, 1);
            uvecH = tg.CreateUnitVector2d(1, 0);
            rangeMin = tg.CreatePoint2d(dv.Position.X - dv.Width/2, dv.Position.Y - dv.Height/2);
            rangeMax = tg.CreatePoint2d(rangeMin.X + dv.Width, rangeMin.Y + dv.Height);
            box = I.tg.CreateBox2d();   
            box.MaxPoint = rangeMax; box.MinPoint = rangeMin;
            cenLeft = tg.CreatePoint2d(dv.Position.X - dv.Width / 2, dv.Position.Y);
            cenRight = tg.CreatePoint2d(dv.Position.X + dv.Width / 2, dv.Position.Y);
            cenTop = tg.CreatePoint2d(dv.Position.X, dv.Position.Y + dv.Height / 2);
            cenBottom = tg.CreatePoint2d(dv.Position.X, dv.Position.Y - dv.Height / 2);
            cenCurv();
            getCurves();
            //getPoints();
            addDims();
            u.transactEnd();
        }

        public void radius()
        {
            t = t == 0 ? u.getPar(dv, e => e.Thickness.Value):t;
            List<DrawingCurve> dcs = u.getCurves<DrawingCurve>(dv, e => check(e, a => a.Radius > 2.1));
            if (dcs.Count() == 0) return;
            DrawingSketch ds = dv.Sketches.Add();
            ds.Edit();
            List<SketchLine> lines = new List<SketchLine>();
            for (int i = 1; i < dcs.Count(); i+=2)
            {
                SketchArc arc = ds.AddByProjectingEntity(dcs[i]) as SketchArc;
                lines.Add(ds.SketchLines.AddByTwoPoints(arc.CenterSketchPoint, arc.EndSketchPoint));
                lines[lines.Count - 1].Centerline = true;
                lines.Add(ds.SketchLines.AddByTwoPoints(arc.CenterSketchPoint, arc.StartSketchPoint));
                lines[lines.Count - 1].Centerline = true;
                
            }
            ds.ExitEdit();

            List<DrawingCurve> ls = u.getCurves<DrawingCurve>(dv, e => e.ModelGeometry is SketchLine && (e.ModelGeometry as SketchLine).Centerline);

            for (int i = startInt; i < dcs.Count(); i += 2)
            {
                u.addDim<RadiusGeneralDimension>(dv, dcs[i], 0, st: st1);      
                u.addDim<AngularGeneralDimension>(dv, ls[i-1], ls[i], st: st2);
            }
        }

        public void tangent()
        {
            HashSet<DrawingCurve> exp = new HashSet<DrawingCurve>();
            r = u.getPar(dv, e => e.BendRadius.Value);
            t = t == 0 ? u.getPar(dv, e => e.Thickness.Value):t;
            List<DrawingCurve> dcs = u.getCurves<DrawingCurve>(dv, e => check(e, a => u.eq(a.Radius, r)));
            foreach (DrawingCurve dc in dcs)
            {
                DrawingCurve start = u.findAtPoint<DrawingCurve>
                    (dv, e => check(e,dc.StartPoint));
                DrawingCurve end = u.findAtPoint<DrawingCurve>
                    (dv, e => check(e, dc.EndPoint));
                bool f = check(dc);
                if (start == null) start = add(dc, f);
                if (end == null) end = add(dc, !f); 
                DrawingCurve [] cs = {start,end};
                if (except(exp, cs  /*start != null && end != null*/))
                {
                    if (!u.eq(u.scalar(start, end), 0))
                    {
                        if (start.ModelGeometry is Edge) { 
                            u.addDim<LinearGeneralDimension>(dv, start, st: st1);
                            exp.Add(start);}
                        if (end.ModelGeometry is Edge)
                        {
                            u.addDim<LinearGeneralDimension>(dv, end, st: st1);
                            exp.Add(end);
                        }
                        u.addDim<AngularGeneralDimension>(dv, start, end, st: st2);
                    }
                    else
                    {
                        double scale = dv.Scale;
                        DrawingCurve mn = u.min<DrawingCurve>(cs, e => u.getLenght(e));
                        DrawingCurve mx = u.max<DrawingCurve>(cs, e => u.getLenght(e));
                        DrawingCurve i = u.findAtPoint(dv, dc.CenterPoint, (r+t)*scale/10, e => u.eq(u.getLenght(e),u.getLenght(mx)) && !mx.Equals(e));
                        Point2d ip = u.max<Point2d>(new Point2d[] { mn.EndPoint, mn.StartPoint }, pt => pt.VectorTo(dc.CenterPoint).Length);
                        LinearGeneralDimension dim = u.addDim(dv, i, mn, null, ip, mn.StartPoint, mn.EndPoint);
                        dim.Style = st2;
                        //if (!check(dim)) dim.Delete();
                        exp.Add(mn); exp.Add(i);
                    }
                }
            }
        }

        public DrawingCurve add(DrawingCurve dc, bool end)
        {
            
            double inc = 0;
            DrawingSketch ds = dv.Sketches.Add();
            ds.Edit();
            SketchArc se = ds.AddByProjectingEntity(dc) as SketchArc;
            Point2d pt = !end ? se.StartSketchPoint.Geometry : se.EndSketchPoint.Geometry;
            Point2d pt2 = pt.Copy();
            Vector2d v = u.getTangentVec(se.Geometry.Evaluator, inc, pt);
            v.Normalize(); //v.ScaleBy(-1);
            pt2.TranslateBy(v);
            SketchLine sl = end ? ds.SketchLines.AddByTwoPoints(se.EndSketchPoint, pt2): ds.SketchLines.AddByTwoPoints(se.StartSketchPoint, pt2);
            ds.ExitEdit();
            return u.getCurves<DrawingCurve>(dv, e => e.ModelGeometry.Equals(sl)).FirstOrDefault();
        }

        public bool check(LinearGeneralDimension dim)
        {
            LineSegment2d ext1 = dim.ExtensionLineOne as LineSegment2d, ext2 = dim.ExtensionLineTwo as LineSegment2d;
            Vector2d x = ext1.StartPoint.VectorTo(ext2.StartPoint), y = ext1.EndPoint.VectorTo(ext2.EndPoint);
            if (x.AngleTo(y) > Math.PI/8) return false;
            return true;
        }

        public bool check(DrawingCurve dc)
        {
            return dc.StartPoint.X < dc.CenterPoint.X;
//             Vector2d v1 = dc.CenterPoint.VectorTo(dc.StartPoint), v2 = dc.CenterPoint.VectorTo(dc.EndPoint);
//             double a1 = v1.AngleTo(util.createVector2d(-1, -1)), a2 = v2.AngleTo(util.createVector2d(-1, -1));
//             return (a2 - a1 < 0) ? true : false;
        }

        public bool check(DrawingCurve dc, Point2d pt)
        {
           return ((dc.StartPoint.IsEqualTo(pt) || dc.EndPoint.IsEqualTo(pt)) && dc.ProjectedCurveType == Curve2dTypeEnum.kLineSegmentCurve2d);
        }

        public bool check(DrawingCurve dc, Func<Arc3d,bool>f)
        {
            Edge e; Arc3d a;
            if (dc.CurveType == CurveTypeEnum.kCircularArcCurve && dc.ModelGeometry != null)
            {
                e = dc.ModelGeometry as Edge;
                if (e == null) return false;
                a = e.Geometry as Arc3d;
                if (a == null) return false;
                if (f(a)) return true;
            }
            return false;
        }

        public bool except(HashSet<DrawingCurve> hs, DrawingCurve[] cs)
        {
            bool r = false;
            foreach (var item in cs)
            {
                if (item == null) return false;
                if (except(hs, item)) { r = true; }
            }
            return r;
        }
                        
        public bool except(HashSet<DrawingCurve> hs, DrawingCurve dc)
        {
            return !hs.Contains(dc);
        }

        public void cenCurv(double tol = 0.01)
        {
            ObjectsEnumerator col;    
            col = dv.Parent.FindUsingPoint(cenLeft, tol);
            if (col != null && col.Count != 0) { leftDC.set((DrawingCurveSegment)col[1]); leftCurv = false;}
            col = dv.Parent.FindUsingPoint(cenRight, tol);
            if (col != null && col.Count != 0) { rightDC.set((DrawingCurveSegment)col[1]); rightCurv = false; }
            col = dv.Parent.FindUsingPoint(cenTop, tol);
            if (col != null && col.Count != 0) { topDC.set((DrawingCurveSegment)col[1]); topCurv = false; }
            col = dv.Parent.FindUsingPoint(cenBottom, tol);
            if (col != null && col.Count != 0) { bottomDC.set((DrawingCurveSegment)col[1]); bottomCurv = false; }
        }

        public void getCurves(int round = 2)
        {
            //curves = dv.DrawingCurves.OfType<DrawingCurve>().OrderBy(c => Math.Round(c.Evaluator2D.RangeBox.MinPoint.X,round)).
            //    OrderBy(c => Math.Round(c.Evaluator2D.RangeBox.MinPoint.Y,round));
            //if (curves == null) return;
            drawingCurves cur = new drawingCurves();

            foreach (DrawingCurve item in dv.DrawingCurves)
            {
                cur.set(item.Segments[1]);
                
                if (leftCurv)
                {
                    if (cur.minX < leftDC.minX) leftDC.set(cur.curve, cur.max, cur.min, cur.minX, cur.minY, cur.maxX, cur.maxY);
                    else if (u.eq(cur.minX, leftDC.minX))
                    {
                        if (cur.minY < leftDC.minY) leftDC.set(cur.curve, cur.max, cur.min, cur.minX, cur.minY, cur.maxX, cur.maxY);
                    }
                }
                if (bottomCurv)
                {
                    if (cur.minY < bottomDC.minY) bottomDC.set(cur.curve, cur.max, cur.min, cur.minX, cur.minY, cur.maxX, cur.maxY);
                    else if (u.eq(cur.minY, bottomDC.minY) && cur.minX < bottomDC.minX) bottomDC.set(cur.curve, cur.max, cur.min, cur.minX, cur.minY, cur.maxX, cur.maxY);
                }
                if (rightCurv)
                {
                    if (cur.maxX > rightDC.maxX) rightDC.set(cur.curve, cur.max, cur.min, cur.minX, cur.minY, cur.maxX, cur.maxY);
                    else if (u.eq(cur.maxX, rightDC.maxX) && cur.maxY > rightDC.maxY) rightDC.set(cur.curve, cur.max, cur.min, cur.minX, cur.minY, cur.maxX, cur.maxY);
                }
                if (topCurv)
                {
                    if (cur.maxY > topDC.maxY) topDC.set(cur.curve, cur.max, cur.min, cur.minX, cur.minY, cur.maxX, cur.maxY);
                    else if (u.eq(cur.maxY, topDC.maxY) && cur.maxX > topDC.maxX) topDC.set(cur.curve, cur.max, cur.min, cur.minX, cur.minY, cur.maxX, cur.maxY);
                }
            }
            //leftDC = curves.First();
            //curves = dv.DrawingCurves.OfType<DrawingCurve>().OrderBy(c => Math.Round(c.Evaluator2D.RangeBox.MaxPoint.X, round)).
            //    OrderBy(c => Math.Round(c.Evaluator2D.RangeBox.MaxPoint.Y, round));
            //rightDC = curves.Last();
            //curves = dv.DrawingCurves.OfType<DrawingCurve>().OrderBy(c => Math.Round(c.Evaluator2D.RangeBox.MinPoint.Y, round)).
            //    OrderBy(c => Math.Round(c.Evaluator2D.RangeBox.MinPoint.X, round));
            //bottomDC = curves.First();
            //curves = dv.DrawingCurves.OfType<DrawingCurve>().OrderBy(c => Math.Round(c.Evaluator2D.RangeBox.MaxPoint.Y, round)).
            //    OrderBy(c => Math.Round(c.Evaluator2D.RangeBox.MaxPoint.X, round));
            //topDC = curves.Last();
            leftDC.setRev(); rightDC.setRev(); topDC.setRev(false); bottomDC.setRev(false);
            leftDC.chCurve(poz.left); rightDC.chCurve(poz.right); topDC.chCurve(poz.top); bottomDC.chCurve(poz.bottom);
        }



        public static double[] intersect(DrawingCurve dc)
        {
            TransientGeometry tg = I.tg;
            double[] pts = new double[4];
            Arc2d arc = (Arc2d)dc.Segments[1].Geometry;
            Point2d cen = arc.Center; Point2d pt = null;
            Line2d l = tg.CreateLine2d(tg.CreatePoint2d(cen.X - (arc.Evaluator.RangeBox.MaxPoint.X - arc.Evaluator.RangeBox.MinPoint.X), cen.Y), tg.CreateUnitVector2d(1, 0));
            ObjectsEnumerator o;
            o = l.IntersectWithCurve(arc);
            if (o != null)
            {
                if (o.Count >= 1) pt = (Point2d)o[1];
                if (pt.X < cen.X) pts[0] = pt.X;
                else pts[1] = pt.X;
                if (o.Count == 2) { pts[0] = ((Point2d)o[1]).X; pts[1] = ((Point2d)o[2]).X; }
            }

            l = tg.CreateLine2d(tg.CreatePoint2d(cen.X, cen.Y - (arc.Evaluator.RangeBox.MaxPoint.Y - arc.Evaluator.RangeBox.MinPoint.Y)), tg.CreateUnitVector2d(0, 1));
            o = l.IntersectWithCurve(arc);
            if (o != null)
            {
                if (o.Count >= 1) pt = (Point2d)o[1];
                if (pt.Y < cen.Y) pts[2] = pt.Y;
                else pts[3] = pt.Y;
                if (o.Count == 2) { pts[2] = ((Point2d)o[1]).Y; pts[3] = ((Point2d)o[2]).Y; }
            }
            return pts;
        }

        public double getParamAtPoint(DrawingCurve dc, Point2d pt)
        {
            double[] param = new double[1];
            double[] pts = new double[2] {pt.X, pt.Y };
            double[] guessParam = new double[1];
            double[] maxDeviat = new double[1];
            SolutionNatureEnum[] sne = new SolutionNatureEnum[1];

            dc.Evaluator2D.GetParamAtPoint(ref pts, ref guessParam, ref maxDeviat, ref param, ref sne);
            return param[0];
        }

        public void addDims()
        {
            leftDim = addDim(dv,bottomDC.curve, topDC.curve, PointIntentEnum.kCircularBottomPointIntent, PointIntentEnum.kCircularTopPointIntent,
                bottomDC.rev, topDC.rev, /*bottomDC.min, topDC.max,*/rangeMin, tg.CreatePoint2d(rangeMin.X, rangeMax.Y), 15, DimensionTypeEnum.kVerticalDimensionType);
            bottomDim = addDim(dv,leftDC.curve, rightDC.curve, PointIntentEnum.kCircularLeftPointIntent, PointIntentEnum.kCircularRightPointIntent,
                leftDC.rev, rightDC.rev, /*leftDC.min, leftDC.max ,*/rangeMin, tg.CreatePoint2d(rangeMax.X, rangeMin.Y), -15, DimensionTypeEnum.kHorizontalDimensionType);
            if (dv.IsFlatPatternView == false)
            {
                dimStyle(leftDim);
                dimStyle(bottomDim);
            }
            dimToBend();
        }

        public void dimStyle(LinearGeneralDimension dim, string styl = "Текст*", int prec = 0)
        {
            DimensionStyle st = null;
            try
            {
                st = ((DrawingDocument)dv.Parent.Parent).StylesManager.DimensionStyles[styl];
            }
            catch
            {

            }
            if (st != null) dim.Style = st;
            else { dim.Precision = prec; dim.Text.FormattedText = dim.Text.FormattedText + "*"; }
        }

        public static LinearGeneralDimension addDim(DrawingView dv,object ent1, object ent2, PointIntentEnum pi1, PointIntentEnum pi2, bool rev1, bool rev2, /*Point2d pt1, Point2d pt2,*/ Point2d mpt1, Point2d mpt2, double offset=15, 
            DimensionTypeEnum align = DimensionTypeEnum.kAlignedDimensionType)
        {
            //double [] param1 = new double[2] {pt1.X, pt1.Y}, param2 = new double[2] {pt2.X, pt2.Y};
            //double [] p1,p2;
            //((DrawingCurve)ent1).Evaluator2D.GetParamAtPoint(param1, p1,)
            PointIntentEnum pie = PointIntentEnum.kStartPointIntent;
            if (rev1) pie = PointIntentEnum.kEndPointIntent;
            
            GeometryIntent intent1 = (((DrawingCurve)ent1).ProjectedCurveType == Curve2dTypeEnum.kCircularArcCurve2d) ?           
                dv.Parent.CreateGeometryIntent(ent1, pi1):
                dv.Parent.CreateGeometryIntent(ent1, pie);
            pie = PointIntentEnum.kStartPointIntent;
            if (!rev2) pie = PointIntentEnum.kEndPointIntent;
            GeometryIntent intent2 = (((DrawingCurve)ent2).ProjectedCurveType == Curve2dTypeEnum.kCircularArcCurve2d) ?
                dv.Parent.CreateGeometryIntent(ent2, pi2) : 
                dv.Parent.CreateGeometryIntent(ent2, pie);
            Vector2d vec = mpt1.VectorTo(mpt2);
            vec.ScaleBy(0.5);
            Point2d mpt = mpt1.Copy();
            mpt.TranslateBy(vec);
            Matrix2d mtx = I.tg.CreateMatrix2d();
            mtx.SetToRotation(Math.PI/2, mpt);
            vec.TransformBy(mtx);
            vec.Normalize();
            vec.ScaleBy(offset/10);
            mpt.TranslateBy(vec);
            return dv.Parent.DrawingDimensions.GeneralDimensions.AddLinear(mpt, intent1, intent2, align);
        }

        public static void addHoleAnnotation(DrawingView dv,Sheet sh)
        {
            Point2d pt = null; DrawingCurve dc = null;
            HashSet<Inventor.Box2d> boxes = new HashSet<Box2d>();
            List<Point2d> ptExct = new List<Point2d>();
            pt = u.spiralPosition(dv, sh, 5, 5, dc.CenterPoint,ref ptExct);
            HoleThreadNote htn = sh.DrawingNotes.HoleThreadNotes.Add(pt, dc);
            addBox(sh, ref boxes);
        }

        public static void addBox(Sheet sh, ref HashSet<Inventor.Box2d> boxes)
        {
            foreach (LinearGeneralDimension item in sh.DrawingDimensions)
            {
                boxes.Add(item.Text.RangeBox);
            }
        }

        public void dimToBend()
        {
            double offset = 12, offsetN = 0;
            List<drawingCurves> rang = new List<drawingCurves> { leftDC, rightDC, topDC, bottomDC };
            IEnumerable<DrawingCurve> curves = dv.DrawingCurves.OfType<DrawingCurve>().Where(c => c.EdgeType == DrawingEdgeTypeEnum.kBendDownEdge || c.EdgeType == DrawingEdgeTypeEnum.kBendUpEdge);
            curves = curves.Distinct();
            I.getStyle("Гибы");
            foreach (DrawingCurve cur in curves)
            {
                double dist = 10000;
                foreach (DrawingCurveSegment seg in cur.Segments)
                {
                    double dCur = seg.StartPoint.DistanceTo(seg.EndPoint);
                    if (dCur < dist) dist = dCur;
                }
                if (dist < 0.1) continue;
                drawingCurves n = new drawingCurves();
                n.curve = cur; bool X = true;
                if (u.eq(cur.StartPoint.X, cur.EndPoint.X)) X = false;
                n.setRev(X);
                drawingCurves r; Point2d retPt = null;
                    r = rang.OrderBy(c => u.getDist(c.curve.StartPoint, cur, 1000, ref retPt)).First(en => en.vert != X);
                    double d = u.getDist(r.curve.StartPoint, n.curve, 1000, ref retPt);

                    if (d < 1.2) offsetN = -10;
                    else offsetN = 0;

//                     offs = rang.OrderBy(c => util.getDist(cur.StartPoint, c.curve, 1000)).First(en => en.vert == X);
//                     d = util.getDist(n.curve.StartPoint, offs.curve, 1000);
                    addBendDim(r.curve, n.curve, offset, offsetN, X);
                //if (X)
                //addDim(r.curve, n.curve, PointIntentEnum.kCircularBottomPointIntent, PointIntentEnum.kCircularTopPointIntent,
                //r.rev, n.rev, /*bottomDC.min, topDC.max,*/rangeMin, tg.CreatePoint2d(rangeMin.X, rangeMax.Y), 15, DimensionTypeEnum.kVerticalDimensionType);
                //else
                //    addDim(r.curve, n.curve, PointIntentEnum.kCircularBottomPointIntent, PointIntentEnum.kCircularTopPointIntent,
                //r.rev, n.rev, /*bottomDC.min, topDC.max,*/rangeMin, tg.CreatePoint2d(rangeMin.X, rangeMax.Y), 15, DimensionTypeEnum.kHorizontalDimensionType);
            }
            //I.setStyle(I.oldDimStyle.Name);
        }

        public void addBendDim(DrawingCurve dc1, DrawingCurve dc2, double offset, double offsetN, bool X)
        {
            LinearGeneralDimension dim;
            if (dc1.Segments[1].GeometryType == Curve2dTypeEnum.kLineSegmentCurve2d && dc2.Segments[1].GeometryType == Curve2dTypeEnum.kLineSegmentCurve2d)
            {
                LineSegment2d ls1 = (LineSegment2d)dc1.Segments[1].Geometry, ls2 = (LineSegment2d)dc2.Segments[1].Geometry;
                if (ls1.Direction.IsParallelTo(ls2.Direction))
                {
                    Point2d pt1 = minPt(dc1, X:X), pt2 = minPt(dc2, X:X);
                    if (X)
                    {
                        if (pt1.X < dv.Position.X) offsetN = -offsetN;
                        if (pt1.Y > dv.Position.Y) offset = -offset;
                    }
                    else
                    {
                        if (pt1.X < dv.Position.X) offset = -offset;
                        if (pt1.Y < dv.Position.Y) offsetN = -offsetN;
                    }
                    Point2d midpt = u.midPt(pt1, pt2, offset / 10, offsetN / 10);
                        Point2d intPt = minSeg(dc1, midpt);
//                         if (midpt.DistanceTo(dc1.StartPoint) < midpt.DistanceTo(dc1.EndPoint))
//                         {
//                             intPt = dc1.EndPoint;
//                         }
//                         else intPt = dc1.StartPoint;
                        GeometryIntent i1, i2;
                        if (intPt != null)
                        i1 = dv.Parent.CreateGeometryIntent(dc1, intPt);
                        else i1 = dv.Parent.CreateGeometryIntent(dc1);
//                         if (midpt.DistanceTo(dc2.StartPoint) < midpt.DistanceTo(dc2.EndPoint))
//                         {
//                             intPt = dc2.EndPoint;
//                         }
//                         else intPt = dc2.StartPoint;
                        intPt = minSeg(dc2, midpt);
                        if (intPt != null)
                            i2 = dv.Parent.CreateGeometryIntent(dc2, intPt);
                        else i2 = dv.Parent.CreateGeometryIntent(dc2);
                        if (u.eq(dc1.StartPoint.X, dc1.EndPoint.X))   
                            dim = dv.Parent.DrawingDimensions.GeneralDimensions.AddLinear(midpt, i1, i2, DimensionTypeEnum.kHorizontalDimensionType,DimensionStyle: I.dimStyle);
                        else dim = dv.Parent.DrawingDimensions.GeneralDimensions.AddLinear(midpt, i1, i2, DimensionTypeEnum.kVerticalDimensionType, DimensionStyle: I.dimStyle);
                    dim.CenterText();
                        if (offset < 0)
                            gabs.offsetN(dim);
                        else gabs.offsetN(dim, -1);
                        //dim = dv.Parent.DrawingDimensions.GeneralDimensions.AddLinear(midpt, i1, i2, DimensionTypeEnum.kAlignedDimensionType);
                        //dimStyle(dim, "Гибы", 1);
                }
            }
        }

        private Point2d minSeg(DrawingCurve dc, Point2d pt)
        {
            DrawingCurveSegment dcs = null; double d = 10000;
            foreach (DrawingCurveSegment item in dc.Segments)
            {
                double dSeg = item.StartPoint.DistanceTo(pt);
                if (dSeg < d) { dcs = item; d = dSeg; }
            }
            if (dc.StartPoint.IsEqualTo(dcs.StartPoint)) return dc.StartPoint;
            if (dc.StartPoint.IsEqualTo(dcs.EndPoint)) return dc.StartPoint;
            if (dc.EndPoint.IsEqualTo(dcs.StartPoint)) return dc.EndPoint;
            if (dc.EndPoint.IsEqualTo(dcs.EndPoint)) return dc.EndPoint;
            return null;
            //return (pt.DistanceTo(dcs.StartPoint) > pt.DistanceTo(dcs.EndPoint)) ? dcs.EndPoint: dcs.StartPoint;
        }

        private Point2d intersectWith(DrawingCurve dc, Point2d pt, ref Point2d pt2, bool X = false, bool rev = false)
        {
            Line2d seg;
            UnitVector2d v1 = uvecH.Copy(), v2 = uvecV.Copy();
            if (rev)
            {
                Vector2d vec = v1.AsVector(); vec.ScaleBy(-1); v1 = vec.AsUnitVector();
                vec = v2.AsVector(); vec.ScaleBy(-1); v2 = vec.AsUnitVector();
            }
            seg = (X) ? tg.CreateLine2d(pt, v1) : tg.CreateLine2d(pt, v2);

            if (dc.ProjectedCurveType == Curve2dTypeEnum.kCircularArcCurve2d || dc.ProjectedCurveType == Curve2dTypeEnum.kCircleCurve2d)
            {
                objs = seg.IntersectWithCurve(dc.Segments[1].Geometry, 0.1);
                if (objs != null)
                {
                    Point2d pt1 = (Point2d)objs[1];
                    if (objs.Count == 2) pt2 = (Point2d)objs[2];
                    return pt1;
                }
            }
            else if (dc.ProjectedCurveType == Curve2dTypeEnum.kLineSegmentCurve2d)
            {
                return minPt(dc, X: X);
            }
            return null;
        }

        static public Point2d minPt(DrawingCurve dc, int round = 2, bool X = false)
        {
            Point2d pt1 = dc.StartPoint, pt2 = dc.EndPoint;
            Vector2d vec = pt1.VectorTo(pt2);
            if (X)
            {
                if (Math.Round(vec.X, round) == 0)
                    return (Math.Round(vec.Y, round) > 0) ? pt1 : pt2;
                return (Math.Round(vec.X, round) > 0) ? pt1 : pt2;
            }
            else
            {
                if (Math.Round(vec.Y, round) == 0)
                    return (Math.Round(vec.X, round) > 0) ? pt1 : pt2;
                return (Math.Round(vec.Y, round) > 0) ? pt1 : pt2;
            }
        }
    }

    public class Drawings : InvDoc.InvDocument<DrawingDocument>
    {
        public Sheet sheet;
        public DrawingView view;
        public GeometryIntent intent;
        public DrawingDimension dim;
        public Centermark center;
        public HashSet<Centermark> centers = new HashSet<Centermark>();
        public HashSet<Centermark> remove = new HashSet<Centermark>();
        public Centerline centerline;
        public ObjectCollection col;
        public TransientGeometry tg;
        public Drawings(DrawingDocument drw, bool center = false)
            : base(drw)
        {
            sheet = drw.ActiveSheet;
            Inventor.Application app =  Macros.StandardAddInServer.m_inventorApplication;
            col = app.TransientObjects.CreateObjectCollection();
            tg = app.TransientGeometry;
            try
            {
                //clearCenter(sheet);
            foreach (DrawingView view in sheet.DrawingViews)
            {
                foreach (DrawingCurve cur in view.DrawingCurves)
                {
                    if (cur.CurveType == CurveTypeEnum.kCircleCurve && !hasCenter(cur,sheet))
                    {
                        intent = sheet.CreateGeometryIntent(cur);
                        sheet.Centermarks.Add(intent);                                                   
                    }
                    //if (cur.CurveType == CurveTypeEnum.kCircularArcCurve)
                    //{
                    //    DrawingCurveSegment seg;
                    //    if ((seg = cur.Segments[1]).GeometryType == Curve2dTypeEnum.kCircularArcCurve2d)
                    //    {
                    //        double ang = Math.Abs(((Arc2d)(cur.Segments[1].Geometry)).SweepAngle);
                    //        int v = (int)(ang * 1000);
                    //        if (v == 3141 && !hasCenter(cur, sheet))
                    //        {
                    //            intent = sheet.CreateGeometryIntent(cur);
                    //            sheet.Centermarks.Add(intent);
                    //        }
                    //    }
                    //}
                }
                if (center) return;
                removeFunc(centers, sheet);
                centers.Clear();
                removeFunc(centers, sheet, false);
            }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
            //IEnumerable<Sheet> ie = getSheets.OfType<Sheet>().Where(s => s.Name == "Лист:1");
            //foreach (Sheet s in ie)
            //{
            //    DrawingSketch sk = s.Sketches.Add();
            //    sk.Edit();
            //    sk.SketchCircles.AddByCenterRadius(invApp.TransientGeometry.CreatePoint2d(), 12);
            //    sk.ExitEdit();
            //}

        }

        public void clearCenter(Sheet sheet)
        {
            foreach (Centerline item in sheet.Centerlines)
            {
                item.Delete();
            }
        }

        static public void addDim(DrawingCurve dc1, Point2d pt1, DrawingCurve dc2, Point2d pt2, double offset = 10)
        {
            GeometryIntent intent1 = dc1.Parent.Parent.CreateGeometryIntent(dc1);
            GeometryIntent intent2 = dc1.Parent.Parent.CreateGeometryIntent(dc2);
            Point2d mpt = midPoint(pt1, pt2, offset/10);
            dc1.Parent.Parent.DrawingDimensions.GeneralDimensions.AddLinear(mpt, intent1, intent2);
        }

        static public LinearGeneralDimension addDim(DrawingCurve dc1, double offset = 10)
        {
            GeometryIntent intent1 = dc1.Parent.Parent.CreateGeometryIntent(dc1, dc1.StartPoint);
            GeometryIntent intent2 = dc1.Parent.Parent.CreateGeometryIntent(dc1, dc1.EndPoint);
            Point2d mpt = midPoint(dc1.StartPoint, dc1.EndPoint, offset / 10, 0.7);
            return dc1.Parent.Parent.DrawingDimensions.GeneralDimensions.AddLinear(mpt, intent1, intent2, DimensionTypeEnum.kAlignedDimensionType);
        }

        static public LinearGeneralDimension addDim(DrawingCurve dc, Vector2d v)
        {
            Sheet sh = dc.Parent.Parent;
            GeometryIntent i1 = sh.CreateGeometryIntent(dc, dc.StartPoint), i2 = sh.CreateGeometryIntent(dc, dc.EndPoint);
            Point2d pt = midPoint(dc.StartPoint, dc.EndPoint, v);
            if (InvDoc.u.eq(dc.StartPoint.X, dc.EndPoint.X))
                return sh.DrawingDimensions.GeneralDimensions.AddLinear(pt, i1, i2, DimensionTypeEnum.kVerticalDimensionType);
            else if (InvDoc.u.eq(dc.StartPoint.Y, dc.EndPoint.Y))
                return sh.DrawingDimensions.GeneralDimensions.AddLinear(pt, i1, i2, DimensionTypeEnum.kHorizontalDimensionType);
            else
                return sh.DrawingDimensions.GeneralDimensions.AddLinear(pt, i1, i2, DimensionTypeEnum.kAlignedDimensionType);
        }

        static public LinearGeneralDimension addDim(DrawingCurve dc1, DrawingCurve dc2, double offset = 10, double offset2 = 0, Point2d pt1 = null, Point2d pt2 = null,
            DimensionTypeEnum type = DimensionTypeEnum.kAlignedDimensionType, Vector2d vec = null)
        {
            GeometryIntent intent1, intent2;
            if (pt1 != null && pt2 != null)
            {
                intent1 = dc1.Parent.Parent.CreateGeometryIntent(dc1, pt1);
                intent2 = dc1.Parent.Parent.CreateGeometryIntent(dc2, pt2);
            }
            else
            {
                intent1 = dc1.Parent.Parent.CreateGeometryIntent(dc1, dc1.MidPoint);
                intent2 = dc1.Parent.Parent.CreateGeometryIntent(dc2, dc2.MidPoint);
                pt1 = dc1.MidPoint; pt2 = dc2.MidPoint;
            }
            Point2d mpt = midPoint(pt1, pt2, offset / 10, offset2 / 10);
            if (vec != null) mpt.TranslateBy(vec);
            return dc1.Parent.Parent.DrawingDimensions.GeneralDimensions.AddLinear(mpt, intent1, intent2, type);
        }

        static public void surfCurve(DrawingView dv, ref DrawingCurve dc, ref DrawingCurve dc1, ref DrawingCurve dc2)
        {
            Document doc = (Document)dv.ReferencedDocumentDescriptor.ReferencedDocument;
            if (doc.DocumentType == DocumentTypeEnum.kPartDocumentObject && doc.SubType == "{9C464203-9BAE-11D3-8BAD-0060B0CE6BB4}")
            {
                double thickness = Math.Round((double)((SheetMetalComponentDefinition)((PartDocument)doc).ComponentDefinition).Thickness.Value, 3);
                IEnumerable<DrawingCurve> ie = dv.DrawingCurves.OfType<DrawingCurve>().OrderBy(o => o.Evaluator2D.RangeBox.MinPoint.Y).OrderBy(o1 => o1.Evaluator2D.RangeBox.MinPoint.X);
                if (dc == null)
                {
                    foreach (DrawingCurve curv in ie)
                    {
                        double l = getLength(curv) / dv.Scale;
                        l = Math.Round(l, 3);
                        if (l == thickness) { dc = curv; break; }
                    }
                }
                if (dc1 == null)
                foreach (DrawingCurve curv in ie)
                {
                    if (curv.Equals(dc)) continue;
                    if (isEqualPoint(curv.Segments[1].StartPoint, dc.Segments[1].StartPoint) || isEqualPoint(curv.Segments[1].EndPoint, dc.Segments[1].StartPoint)) { dc1 = curv; break; }
                }
                if (dc2 == null)
                foreach (DrawingCurve curv in ie)
                {
                    if (curv.Equals(dc)) continue;
                    if (isEqualPoint(curv.Segments[1].StartPoint, dc.Segments[1].EndPoint) || isEqualPoint(curv.Segments[1].EndPoint, dc.Segments[1].EndPoint)) { dc2 = curv; break; }
                }
            }
        }

        static public void gab(DrawingView dv)
        {
            TransientGeometry tg = I.tg;
            Line2d left, top, right, bottom; 
            DrawingCurve [] curves = new DrawingCurve[4];
            Line2d [] seg = new Line2d[4];
            Point2d[] pts = new Point2d[4];
            left = tg.CreateLine2d(tg.CreatePoint2d(dv.Position.X - dv.Width/2, dv.Position.Y - dv.Height/2),
                tg.CreateUnitVector2d(0, 1));
            seg[0] = left;
            top = tg.CreateLine2d(tg.CreatePoint2d(dv.Position.X - dv.Width / 2, dv.Position.Y + dv.Height / 2),
                tg.CreateUnitVector2d(1, 0));
            seg[1] = top;
            right = tg.CreateLine2d(tg.CreatePoint2d(dv.Position.X + dv.Width / 2, dv.Position.Y - dv.Height / 2),
                tg.CreateUnitVector2d(0, 1));
            seg[2] = right;
            bottom = tg.CreateLine2d(tg.CreatePoint2d(dv.Position.X - dv.Width / 2, dv.Position.Y - dv.Height / 2),
                tg.CreateUnitVector2d(1, 0));
            seg[3] = bottom;
            IEnumerable<DrawingCurve> dcs = dv.DrawingCurves.OfType<DrawingCurve>().Where(c => c.ProjectedCurveType == Curve2dTypeEnum.kLineSegmentCurve2d)
                .OrderBy(e => Math.Round(e.Evaluator2D.RangeBox.MinPoint.X, 2)).OrderBy(e => Math.Round(e.Evaluator2D.RangeBox.MinPoint.Y, 2)); curves[0] = dcs.First();
            dcs = dcs.OrderBy(e => Math.Round(e.Evaluator2D.RangeBox.MaxPoint.X, 2)).OrderBy(e => Math.Round(e.Evaluator2D.RangeBox.MaxPoint.Y, 2)); curves[2] = dcs.Last();
            dcs = dcs.OrderBy(e => Math.Round(e.Evaluator2D.RangeBox.MinPoint.Y, 2)).OrderBy(e => Math.Round(e.Evaluator2D.RangeBox.MinPoint.X, 2)); curves[1] = dcs.First();
            dcs = dcs.OrderBy(e => Math.Round(e.Evaluator2D.RangeBox.MaxPoint.Y, 2)).OrderBy(e => Math.Round(e.Evaluator2D.RangeBox.MaxPoint.X, 2)); curves[3] = dcs.Last();
            //objs = intersectWith(left, curves[0]); objs = intersectWith(left, curves[1]); objs = intersectWith(left, curves[2]);
            //Line2d l = tg.CreateLine2d(tg.CreatePoint2d(0,-5), tg.CreateUnitVector2d(1,0));

            //objs = l.IntersectWithCurve(tg.CreateCircle2d(tg.CreatePoint2d(), 5));
            //if (objs != null)
            //{

            //}

            //for (int i = 0; i < curves.Length; i++)
            //{
            //    for (int j = 0; j < seg.Length; j++)
            //    {
            //        objs = intersectWith(seg[j], curves[i]);
            //        if (objs != null)
            //        {
            //            pts[i] = (Point2d)objs[1];
            //        }
            //    }  
            //}

            Point2d pt1 = minPt(curves[0]), pt2 = minPt(curves[2]);
            //if (curves[0].ProjectedCurveType == Curve2dTypeEnum.kCircularArcCurve2d && pts[3] != null)
            //{
            //    pt1 = pts[3];
            //}
            //if (curves[2].ProjectedCurveType == Curve2dTypeEnum.kCircularArcCurve2d && pts[1] != null)
            //{
            //    pt1 = pts[1];
            //}
            Vector2d vec = pt1.VectorTo(bottom.RootPoint); vec.Y = 0;
            LinearGeneralDimension dim = addDim(curves[0],curves[2], -15, 0, pt2, pt1, DimensionTypeEnum.kVerticalDimensionType, vec);
            DimensionStyle st = null;
            try
            {
                st = ((DrawingDocument)dim.Parent.Parent).StylesManager.DimensionStyles["Текст*"];
            }
            catch
            {

            }
            if (st != null) dim.Style = st;
            else { dim.Precision = 0; dim.Text.FormattedText = dim.Text.FormattedText + "*"; }

            pt1 =  minPt(curves[1],true); pt2 = minPt(curves[3], true);
            //if (curves[1].ProjectedCurveType == Curve2dTypeEnum.kCircularArcCurve2d && pts[2] != null)
            //{
            //    pt1 = pts[2];
            //}
            //if (curves[3].ProjectedCurveType == Curve2dTypeEnum.kCircularArcCurve2d && pts[0] != null)
            //{
            //    pt1 = pts[0];
            //}
            vec = pt1.VectorTo(left.RootPoint); vec.X = 0;
            dim = addDim(curves[1], curves[3], -15, 0, pt1, pt2, DimensionTypeEnum.kHorizontalDimensionType, vec);
            if (st != null) dim.Style = st;
            else { dim.Precision = 0; dim.Text.FormattedText = dim.Text.FormattedText + "*"; }
            //LinearGeneralDimension dim = addDim(curves[0], curves[2], 10, 0, curves[0].p)
            //foreach (DrawingCurve dc in dv.DrawingCurves)
            //{
            //    for (int i = 0; i < seg.Length; i++)
            //    {
            //        objs = seg[i].IntersectWithCurve(dc);
            //        if (objs != null)
            //        {
            //            pts[i] = (Point2d)objs[1];
            //            curves[i] = ((DrawingCurveSegment)dv.Parent.FindUsingPoint(pts[i])).Parent;
            //        }
            //    }
            //}
        }

        static public Point2d minPt(DrawingCurve dc, bool X = false)
        {
            Point2d pt1 = dc.StartPoint, pt2 = dc.EndPoint;
            Vector2d vec = pt1.VectorTo(pt2);
            if (X)
            {
                if (Math.Round(vec.X, 2) == 0)
                    return (Math.Round(vec.Y, 2) > 0) ? pt1 : pt2;
                return (Math.Round(vec.X, 2) > 0) ? pt1 : pt2;
            }
            else
            {
                if (Math.Round(vec.Y, 2) == 0)
                    return (Math.Round(vec.X, 2) > 0) ? pt1 : pt2;
                return (Math.Round(vec.Y, 2) > 0) ? pt1 : pt2;
            }
        }

        static public bool isEqualPoint(Point2d pt1, Point2d pt2)
        {
            return (((int)(pt1.X * 1000) == (int)(pt2.X * 1000)) && ((int)(pt1.Y * 1000) == (int)(pt2.Y * 1000))) ? true : false;
        }

        static public double getLength(DrawingCurve dc)
        {
            double maxParam, minParam, lenght;
            dc.Evaluator2D.GetParamExtents(out maxParam, out minParam);
            dc.Evaluator2D.GetLengthAtParam(minParam, maxParam, out lenght);
            return lenght;
        }

        static public Point2d getPointAtParam(LineSegment2d seg, double p)
        {
            double[] param = new double[1];
            param[0] = p;
            double[] pts = new double[2];
            seg.Evaluator.GetPointAtParam(ref param, ref pts);
            return Macros.StandardAddInServer.m_inventorApplication.TransientGeometry.CreatePoint2d(pts[0], pts[1]);
        }

        static public Point2d getPointAtParam(DrawingCurve seg, double p)
        {
            double[] param = new double[1];
            param[0] = p;
            double[] pts = new double[2];
            seg.Evaluator2D.GetPointAtParam(ref param, ref pts);
            return Macros.StandardAddInServer.m_inventorApplication.TransientGeometry.CreatePoint2d(pts[0], pts[1]);
        }

        //public void gabarit(DrawingView dv)
        //{
        //    Point2d maxPt, minPt;    dv.
        //    maxPt = tg.CreatePoint2d(sheet.,sheet.DrawingViews[1].Top)
        //    inters
        //    DrawingCurve dc1 = sheet.DrawingViews[1].DrawingCurves.OfType<DrawingCurve>().FirstOrDefault
        //    foreach (DrawingCurve cur in sheet.DrawingViews[1].DrawingCurves)
        //    {
                
        //    }
        //}

        static public Point2d midPoint(Point2d pt1, Point2d pt2, double offset, double offset2 = 0)
        {
            LineSegment2d l = Macros.StandardAddInServer.m_inventorApplication.TransientGeometry.CreateLineSegment2d(pt1, pt2);
            Point2d mpt = l.MidPoint;
            UnitVector2d uv = l.Direction;
            if (offset2 != 0)
            {
                mpt = l.EndPoint;
                Vector2d v = uv.AsVector(); v.ScaleBy(offset2);
                mpt.TranslateBy(v);
            }

            if ((int)pt1.X * 1000 == (int)pt1.X * 1000)
                rotateByVec(ref uv, 90, mpt);
            else rotateByVec(ref uv, 180, mpt);
            Vector2d vec = uv.AsVector();
            vec.ScaleBy(offset);
            mpt.TranslateBy(vec);
            return mpt;
        }

        static public Point2d midPoint(Point2d pt1, Point2d pt2, Vector2d v, double o1 = 1.5, double o2 = 0.7)
        {
            Vector2d v1 = pt2.VectorTo(pt1);
            v1.ScaleBy(0.5);
            pt1.TranslateBy(v1);
            if (o2 != 0) 
            {
                v1.Normalize(); v1.ScaleBy(o2);
                pt1.TranslateBy(v1);
            }
            v.Normalize(); v.ScaleBy(o1); pt1.TranslateBy(v);
            return pt1;
        }

        static public void rotateByVec(ref UnitVector2d vec, double angl, Point2d pt)
        {
            angl = Math.PI * angl / 180;
            Matrix2d mtx = Macros.StandardAddInServer.m_inventorApplication.TransientGeometry.CreateMatrix2d();
            mtx.SetToRotation(angl, pt);
            vec.TransformBy(mtx);
        }

        static public void addSurfaceTextureSymbol(LinearGeneralDimension dim, double param, bool one)
        {
            ObjectCollection col = Macros.StandardAddInServer.m_inventorApplication.TransientObjects.CreateObjectCollection();
            LineSegment2d seg, seg1, seg2;
            seg1 = (LineSegment2d)dim.ExtensionLineOne; seg2 = (LineSegment2d)dim.ExtensionLineTwo;
            if (one) seg = seg1;
            else seg = seg2;
            Point2d mp = getPointAtParam(seg, param);
            GeometryIntent intent = dim.Parent.CreateGeometryIntent(dim, mp);
            Vector2d v = (one) ? seg1.StartPoint.VectorTo(seg2.StartPoint) : seg2.StartPoint.VectorTo(seg1.StartPoint);
            v.Normalize(); v.ScaleBy(-0.05);
            mp.TranslateBy(v);
            col.Add(mp); col.Add(intent);
            SurfaceTextureSymbol st = dim.Parent.SurfaceTextureSymbols.Add(col, SurfaceTextureTypeEnum.kMaterialRemovalProhibitedSurfaceType);
        }

        static public void addSurfaceTextureSymbol(DrawingView dv ,DrawingCurve dc, double param, Vector2d dir)
        {
            ObjectCollection col = Macros.StandardAddInServer.m_inventorApplication.TransientObjects.CreateObjectCollection();
            Point2d mp = getPointAtParam(dc, param);
            //Vector2d dir = dc.StartPoint.VectorTo(dc.EndPoint);
            //dir.Normalize();
            //Matrix2d mtx = Macros.StandardAddInServer.m_inventorApplication.TransientGeometry.CreateMatrix2d();
            //mtx.SetToRotation(-Math.PI / 2, dc.StartPoint);
            //dir.TransformBy(mtx);
            //dir.ScaleBy(0.1);
            mp.TranslateBy(dir);
            GeometryIntent intent = dv.Parent.CreateGeometryIntent(dc, param);
            col.Add(mp); col.Add(intent);
            SurfaceTextureSymbol st = dv.Parent.SurfaceTextureSymbols.Add(col, SurfaceTextureTypeEnum.kMaterialRemovalProhibitedSurfaceType);
        }

        static public void addSSymbol(Sheet sh, string name, Point2d pt, string[] prompt = null)
        {
           SketchedSymbolDefinition ssd = ((DrawingDocument)sh.Parent).SketchedSymbolDefinitions[name];
           sh.SketchedSymbols.Add(ssd, pt, 0, 1, prompt);
        }

        static public void addIzv(DrawingDocument drw, string name, string date)
        {
            SketchedSymbolDefinition ssd = null;
            ssd = drw.SketchedSymbolDefinitions.OfType<SketchedSymbolDefinition>().FirstOrDefault(e => e.Name == "Изв");
            if (ssd == null) InvDoc.u.copySS("Изв", drw);
            Sheet sh = drw.Sheets[1];
            string num = "1";
            bool add = true;
            Point2d pt = sh.TitleBlock.Position;
            pt.Y = pt.Y+3.5;
            SketchedSymbol ss = getIzv(sh, name, date, ref num, ref add);
            if (ss != null)
            {
                pt = ss.Position; pt.Y = pt.Y + 0.5;
                num = (int.Parse(num) + 1).ToString();
            }
            if (add)
            {
                addSSymbol(sh, "Изв", pt, new string[] { num, name, date });
            }
        }

        static public SketchedSymbol getIzv(Sheet sh, string name, string date, ref string num, ref bool add)
        {
            SketchedSymbol ss = sh.SketchedSymbols.OfType<SketchedSymbol>().LastOrDefault(s => s.Name == "Изв");
            if (ss != null)
            {
                SketchedSymbolDefinition ssd = ss.Definition;
                Inventor.TextBox tb = ssd.Sketch.TextBoxes[3];
                string n = ss.GetResultText(tb);
                tb = ssd.Sketch.TextBoxes[4];
                string d = ss.GetResultText(tb);
                tb = ssd.Sketch.TextBoxes[1];
                num = ss.GetResultText(tb);
                if (n == name && d == date) { add = false; return null; }
            }
            return ss;
        }

        public int removeFunc(HashSet<Centermark> centers, Sheet sheet, bool x = true)
        {
            bool flagX = false, flagY = false; int count = 0;
            //IEnumerable<Centerline> cenLines = null;
            //if (x)
            //cenLines = sheet.Centerlines.OfType<Centerline>().Where(e => (int)(e.EndPoint.X * 1000) == (int)(e.StartPoint.X * 1000));
            //else cenLines = sheet.Centerlines.OfType<Centerline>().Where(e => (int)(e.EndPoint.Y * 1000) == (int)(e.StartPoint.Y * 1000));
            foreach (Centermark item in sheet.Centermarks)
            {
                if (item.Centerlines != null)
                {
                    foreach (Centerline cl in item.Centerlines)
                    {
                        if (x && (int)(cl.EndPoint.X * 1000) == (int)(cl.StartPoint.X * 1000))
                        {
                            flagX = true;
                        }
                        else if ((int)(cl.EndPoint.Y * 1000) == (int)(cl.StartPoint.Y * 1000))
                        {
                            flagY = true;
                        }
                    }
                }
                if (!flagX && !flagY)
                {
                    centers.Add(item);
                }
                else if (x && !flagX) centers.Add(item);
                else if (!x && !flagY) centers.Add(item);
                flagX = false; flagY = false;

            }
            foreach (Centermark cen in sheet.Centermarks)
            {
                if (!centers.Contains(cen))
                {
                    continue;
                }
                centers.Remove(cen);
                col.Add(cen);
                IEnumerable<Centermark> ie = null;
                if (x)
                ie = centers.Where(c => (int)(c.Position.X * 1000) == (int)(cen.Position.X * 1000));
                else ie = centers.Where(c => (int)(c.Position.Y * 1000) == (int)(cen.Position.Y * 1000));
                foreach (var item in ie)
                {
                    if (!cen.Position.IsEqualTo(item.Position))
                    {
                        col.Add(item);
                        remove.Add(item);
                    }
                }
                centers.ExceptWith(remove);
                remove.Clear();
                if (col.Count > 1)
                {
                    centerline = sheet.Centerlines.Add(col);
                    count++;
                }
                col.Clear();
            }
            return count;
        }

        public Centermark minCent(ObjectCollection col, bool x = true)
        {
            int min = 10000;
            Centermark c = null;
            foreach (Centermark item in col)
            {
                if (x && (int)item.Position.X * 1000 < min)
                {
                    min = (int)item.Position.X * 1000;
                    c = item;
                }
                else if ((int)item.Position.Y * 1000 < min)
                {
                    min = (int)item.Position.Y * 1000;
                    c = item;
                }
            }
            return c;
        }

        public Centermark maxCent(ObjectCollection col, bool x = true)
        {
            int max = 0;
            Centermark c = null;
            foreach (Centermark item in col)
            {
                if (x && (int)item.Position.X * 1000 > max)
                {
                    max = (int)item.Position.X * 1000;
                    c = item;
                }
                else if ((int)item.Position.Y * 1000 > max)
                {
                    max = (int)item.Position.Y * 1000;
                    c = item;
                }
            }
            return c;
        }

        public bool hasCenter(DrawingCurve cur, Sheet sheet)
        {
            foreach (Centermark cen in sheet.Centermarks)
            {
                try
                {
                if (cen.AttachedEntity != null && cen.AttachedEntity is GeometryIntent && ((GeometryIntent)(cen.AttachedEntity)).Geometry.Equals(cur)) return true;
                //else if (cur.CurveType == CurveTypeEnum.kCircleCurve && cur.CenterPoint.IsEqualTo(cen.Position)) return true;
                else if (cen.AttachedEntity != null && cen.AttachedEntity is FlatPunchResult)
                {
                    cen.Delete();
                }
                }
                catch
                {
                }
            }
            return false;
        }
    }

    public class Izv
    {
        int numString = 0;

        public SketchedSymbol findSS(DrawingDocument drw, string name)
        {
            return drw.ActiveSheet.SketchedSymbols.OfType<SketchedSymbol>().FirstOrDefault(e => e.Name == name);
        }
    }

    public class TableInv : MyTable
    {
        DrawingDocument doc { get; set; }
        public AssemblyDocument asmDoc { get; set; }
        public TransientGeometry tg;
        public DrawingSketch ds;
        public SketchedSymbolDefinition ssd;
        public SketchLine sl;
        public Point2d pt;
        public Inventor.TextBox tb;
        public BOMView bv;
        public BOMRowsEnumerator rowsEnum;
        string pathXML;
        public ComponentDefinition oCompDef;
        public TableInv() { }
        public TableInv(AssemblyDocument asmDoc, String pathXML) : base(pathXML)
        {
            tg = Macros.StandardAddInServer.m_inventorApplication.TransientGeometry;
            this.asmDoc = asmDoc;
            namePerfBase = asmDoc.PropertySets[3][2].Value.ToString();
            countPerf = loadFromAsm(asmDoc);
            if (countPerf != 0)
            {
                int count = countPerf; int start = 0;
                countPerf = 0;
                if (namePerfBase[namePerfBase.Length - 3] == '-')
                {
                    start = int.Parse(namePerfBase.Substring(namePerfBase.Length - 2));
                    namePerfBase = namePerfBase.Substring(0, namePerfBase.Length - 3);
                }

                //namePerfBase = namePerfBase.Substring(0, namePerfBase.IndexOf("0"));

                for (int i = 0; i < count; i++)
                {
                    addPerf(namePerfBase + '-' + String.Format("{0:00}", start));
                    start++;
                }
                if (perfNames[0].EndsWith("-00")) perfNames[0] = perfNames[0].Substring(0, perfNames[0].Length - 3);
                countPerf = count;
            }
        }
        public TableInv(DrawingDocument drw, string pathXML) : base(pathXML)
        {
            doc = drw;
            asmDoc = (AssemblyDocument)InvDoc.u.referendedDoc(doc as Document);//[doc.ReferencedDocuments.Count];
            tg = ((Inventor.Application)doc.Parent).TransientGeometry;
            namePerfBase = asmDoc.PropertySets[3][2].Value.ToString();
            countPerf = loadFromAsm(asmDoc);
            if (countPerf != 0)
            {
                int count = countPerf; int start = 0;
                countPerf = 0;
                if (namePerfBase[namePerfBase.Length - 3] == '-')
                {
                    start = int.Parse(namePerfBase.Substring(namePerfBase.Length - 2));
                    namePerfBase = namePerfBase.Substring(0, namePerfBase.Length - 3);
                }

                //namePerfBase = namePerfBase.Substring(0, namePerfBase.IndexOf("0"));

                for (int i = 0; i < count; i++)
                {
                    addPerf(namePerfBase + '-' + String.Format("{0:00}", start));
                    start++;
                }
                if (perfNames[0].EndsWith("-00")) perfNames[0] = perfNames[0].Substring(0, perfNames[0].Length - 3);
                countPerf = count;
            }
        }
        public void renumberBom(List<TableRow> rows, BOMView bv)
        {
            int i = 1; int val = 0;
            Dictionary<int, int> ins = new Dictionary<int, int>();
            dic = dic ?? new Dictionary<string, string>();

            foreach (var row in rows)
            {
                if (row.cells[2].value == "") continue;
                val = int.Parse(row.cells[2].value);
                if (!ins.ContainsKey(val))
                {
                    ins.Add(val, i);
                    i++;
                }
            }

            foreach (var row in rows)
            {
                foreach (Inventor.BOMRow BOMrow in bv.BOMRows)
                {
                    dic.Clear();
                    dic = propDoc(BOMrow);
                    if (dic["PartNumber"] == row.cells[3].value && dic["desc"] == row.cells[4].value)
                    {
                        if (row.cells[2].value == "") continue;
                        val = int.Parse(row.cells[2].value);
                        BOMrow.ItemNumber = ins[val].ToString();
                        break;
                    }
                }
            }
        }
        public Dictionary<string, string> propDoc(BOMRow row)
        {
            string partNumber, desc, note, length = "", width = "", thickness = "", dxf = "";
            if (row.ComponentDefinitions[1].Type != ObjectTypeEnum.kVirtualComponentDefinitionObject)
            {
                Inventor.Document oDoc = (Inventor.Document)row.ComponentDefinitions[1].Document;
                partNumber = oDoc.PropertySets[3][2].Value.ToString();
                desc = oDoc.PropertySets[3][14].Value.ToString();
                note = oDoc.PropertySets[1][5].Value.ToString();
                if (oDoc.SubType == "{9C464203-9BAE-11D3-8BAD-0060B0CE6BB4}")
                {
                    SheetMetalComponentDefinition compDef = (SheetMetalComponentDefinition)((PartDocument)oDoc).ComponentDefinition;
                    dxf = oDoc.PropertySets[1][7].Value.ToString();
                    thickness = ((double)compDef.Thickness.Value * 10).ToString();
                    if (compDef.HasFlatPattern)
                    {
                        FlatPattern fp = compDef.FlatPattern;
                        length = ((double)fp.Length*10).ToString("#.#");
                        width = ((double)fp.Width*10).ToString("#.#");
                    }
                }
            }
            else
            {
                VirtualComponentDefinition oVirtCompDef = (VirtualComponentDefinition)row.ComponentDefinitions[1];
                partNumber = oVirtCompDef.PropertySets[3][2].Value.ToString();
                desc = oVirtCompDef.PropertySets[3][14].Value.ToString();
                note = oVirtCompDef.PropertySets[1][5].Value.ToString();
            }
            partNumber = partNumber.Trim();
            Dictionary<string, string> dic = new Dictionary<string, string>();
            dic.Add("PartNumber", partNumber); dic.Add("desc", desc); dic.Add("note", note);
            dic.Add("dxf", dxf); dic.Add("thickness", thickness); dic.Add("FPLength", length); dic.Add("FPWidth", width);
            return dic;
        }

        public static Dictionary<string, string> propDoc(Document oDoc)
        {
            string partNumber, desc, note, length = "", width = "", thickness = "", dxf = "";
                partNumber = oDoc.PropertySets[3][2].Value.ToString();
                desc = oDoc.PropertySets[3][14].Value.ToString();
                note = oDoc.PropertySets[1][5].Value.ToString();
                if (oDoc.SubType == "{9C464203-9BAE-11D3-8BAD-0060B0CE6BB4}")
                {
                    SheetMetalComponentDefinition compDef = (SheetMetalComponentDefinition)((PartDocument)oDoc).ComponentDefinition;
                    dxf = oDoc.PropertySets[1][7].Value.ToString();
                    thickness = ((double)compDef.Thickness.Value * 10).ToString();
                    if (compDef.HasFlatPattern)
                    {
                        FlatPattern fp = compDef.FlatPattern;
                        length = ((double)fp.Length * 10).ToString("#.#");
                        width = ((double)fp.Width * 10).ToString("#.#");
                    }
                }
            partNumber = partNumber.Trim();
            Dictionary<string, string> dic = new Dictionary<string, string>();
            dic.Add("PartNumber", partNumber); dic.Add("desc", desc); dic.Add("note", note);
            dic.Add("dxf", dxf); dic.Add("thickness", thickness); dic.Add("FPLength", length); dic.Add("FPWidth", width);
            return dic;
        }

        public void addTable(BOMView bv)
        {
            this.bv = bv;
            string ss = namePerfBase;
            int ind = ss.IndexOf('.');
            string ss1 = ss.Substring(ind, ss.Length - ind);
            string type = namePerfBase.Substring(0, ss1.IndexOf("0")+ind);
            XElement el = new XElement("Root");
            el = addRows(bv.BOMRows, el, namePerfBase, type);
            //el.Save("C:\\Temp.xml");
            addRows(el, ss, type, null, this);
        }

        public XElement addRows(BOMRowsEnumerator rowsEnum, XElement el ,string namePerfBase, string type)
        {
            string partNumber, desc, note, dxf, thickness, FPLength, FPWidth;
            List<string[]> rs = new List<string[]>();
            foreach (BOMRow row in rowsEnum)
            {
                dic = dic ?? new Dictionary<string, string>();
                dic.Clear();
                dic = propDoc(row);
                if (row.BOMStructure == BOMStructureEnum.kPhantomBOMStructure) continue;
                partNumber = dic["PartNumber"]; desc = dic["desc"]; note = dic["note"]; dxf = dic["dxf"]; thickness = dic["thickness"];
                FPLength = dic["FPLength"]; FPWidth = dic["FPWidth"];

                //if (el.Element("VarData") != null && row.ItemNumber.EndsWith()) continue;
                string number = row.ItemNumber;
                if (number.IndexOf(".") != -1) number = number.Substring(number.LastIndexOf('.') + 1, number.Length - number.LastIndexOf('.') - 1);
                XElement data = new XElement("data", new XAttribute("Position", number),
                    new XAttribute("PartNumber", partNumber), new XAttribute("Description", desc), new XAttribute("Count", row.ItemQuantity.ToString()),
                    new XAttribute("Note", note), new XAttribute("Dxf", dxf), new XAttribute("Thickness", thickness), new XAttribute("FPLength", FPLength), new XAttribute("FPWidth", FPWidth));
                //TableRow tr = perfRows.FirstOrDefault(r => r.cells[3].value == data.Attribute("PartNumber").Value && r.cells[4].value == data.Attribute("Description").Value);


                bool add = true;
//                 if (partNumber.IndexOf('.') != -1)
//                 {
//                     XElement tmpEl = XMLDoc.find(el, new Dictionary<string, string>() { { "PartNumber", partNumber }, { "Description", desc } });
//                     if (tmpEl != null)
//                     {
//                         add = false;
//                         XMLDoc.changeVal(tmpEl, new Dictionary<string, string>() { { "Count", (row.ItemQuantity + int.Parse(tmpEl.Attribute("Count").Value)).ToString() } });
//                     }
//                 }
                if (add) el.Add(data);
                if (row.ChildRows != null)
                {
                    //this.asmDoc = (AssemblyDocument)row.ComponentDefinitions[1].Document;
                    if (loadFromAsm((AssemblyDocument)row.ComponentDefinitions[1].Document, data) != 0) data.Add(new XAttribute("Var", "Yes"));
                    //loadFromAsm((AssemblyDocument)row.ComponentDefinitions[1].Document);
                    addRows(row.ChildRows, data, namePerfBase, type);
                }
            }
            return el;
        }

        public void addToSheet(string name, int [] str)
        {
            Inventor.Application app = (Inventor.Application)(doc.Parent);
            app.ScreenUpdating = false;
            try
            {
                Sheet sheet = doc.ActiveSheet;

                pt = tg.CreatePoint2d();
                pt.X = sheet.Width - 0.5; pt.Y = sheet.Height -0.5;
                renumberBom(rows, bv);
                Dictionary<int, string> drawGroup = new Dictionary<int, string>();
                foreach (var item in group)
                {
                    drawGroup.Add(item.Key, item.Value);
                }
//                 if (groupVar != null)
//                 {
//                     foreach (var item in groupVar)
//                     {
//                         drawGroup.Add(item.Key, item.Value);
//                     }
//                 }
                if (cols.Count == 0) colIni();
                formatRows.Clear();
                addHeader();
                forShow(0, 5000, group, true);
                foreach (var item in formatRows)
                {
                    if (item.cells[4].underline == true) item.cells[4].paddingBottom = 2;
                    if (item.cells[3].value.IndexOf('$') != -1 || item.cells[4].value.IndexOf('$') != -1)
                    {
                        foreach (var c in item.cells)
                        {
                            c.paddingBottom = item.height*5;
                        }
                        item.height *= 15;
                    }
                }
                //double h = formatRows.Sum(e => e.height);
                // double y = maxY(sheet);
                int sum = 0;
                    if (str[0] > formatRows.Count)
                    {
                        //pt.X = sheet.Width - 0.5; pt.Y = sheet.Height - 0.5;
                        drawTable(name, formatRows, 0, formatRows.Count, tg.CreatePoint2d(0,0),ref pt,lastLineW: true);
                        addToSheet(sheet, name, ref pt);
                    }
                    else
                    {
                        int i = 0;
                        int val = 0;
                        while (sum <= formatRows.Count)
                        //for (int i = 0; i < str.Length; i++)
                        {
                            try
                            {
                                val = str[i];
                            }
                            catch 
                            {
                            }
                            int start = sum;
                            int end = val + sum + 1;
                            if (end >= formatRows.Count) end = formatRows.Count;
                            else
                            while (formatRows[end - 2].cells[4].underline == true) { end--; }
                            if (i == 0) addOrder(start + 1);
                            else addOrder(start);
                            if (end == formatRows.Count-1) end = formatRows.Count;
                            sum = end;
                        if (i == 0)
                        {
                            //pt.X = sheet.Width - 0.5-0.8; pt.Y = sheet.Height - 0.5;
                            drawTable(name + '_' + i, formatRows, start, end, tg.CreatePoint2d(0,0),ref pt, 3);
                            addToSheet(sheet, name + '_' + i, ref pt);
                        }
                        else
                        {
                            pt.X -= 18.5;
                            if (i != 1) pt.X -= 0.8;
                            if (end == formatRows.Count) drawTable(name + '_' + i, formatRows, start, end, tg.CreatePoint2d(0,0), ref pt, offset: 8,lastLineW: true);
                            else drawTable(name + '_' + i, formatRows, start, end, tg.CreatePoint2d(0,0), ref pt, offset: 8);
                            addToSheet(sheet, name + '_' + i, ref pt);
                            if (end == formatRows.Count) break;
                        }
                        i++;
                        //int split = formatRows.Count - countRow(28.7 - h, y);
                        //while (formatRows[split - 1].cells[4].underline == true) { split--; }
                        //drawTable(name, formatRows, 0, split, tg.CreatePoint2d(), 3);
                        //addToSheet(sheet, name, tg.CreatePoint2d(sheet.Width - 0.5, sheet.Height - 0.5));
                        //addOrder(split);
                        //drawTable(name + '_' + 1, formatRows, split, formatRows.Count, tg.CreatePoint2d(), offset: 8);
                        //addToSheet(sheet, name + '_' + 1, tg.CreatePoint2d(sheet.Width - 19, sheet.Height - 0.5));
                    }
                }
            }
                catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
            finally
            {
                app.ScreenUpdating = true;
            }
        }
        public void addToSheet(Sheet sh,string name,ref Inventor.Point2d pos, bool del = false)
        {
            try 
	        {
               var ss = sh.SketchedSymbols.OfType<SketchedSymbol>().Where(e => e.Name == name);
                if (ss != null)
                    foreach (var item in ss)
                    {
                        pos = item.Position;
                        item.Delete(); 
                    }
	        }
	        catch
	        {

	        }
            if (!del) sh.SketchedSymbols.Add(doc.SketchedSymbolDefinitions[name], pos);
        }

        public double maxY(Sheet sh)
        {
            var ie = sh.Sketches.OfType<DrawingSketch>().First(e => e.Name == "Технические требования");
            return (ie != null) ? ie.TextBoxes.OfType<Inventor.TextBox>().Max(e => e.RangeBox.MaxPoint.Y) + 1: 6.5;
        }
        public void saveInAsm(bool safeCustom = true)
        {
            VariableDataForSpec vd = new VariableDataForSpec((Document)asmDoc);
            Dictionary<string, string> forSafe = new Dictionary<string, string>() { { "Row", "" }, { "Power", "" }, { "Split", "" } };

            var cust = rows.Where(r => r.group == 3 || r.group == 4);
            vd.AttribsDelete("Custom");
            if (safeCustom && cust.Count() > 0)
            {
                int n = 0;
                
                string[] str = new string[cust.Count()];
                string val; int k = 0; string name = "Custom";
                foreach (var item in cust)
                {
                    val = name + ';' + item.group + ';' + item.ToString();
                    str[k] = val; k++;
                    n++;
                }
                vd.AttribAdd(name, str);
            }
            vd.AttribsDelete("Var");
            if (this.countPerf != 0)
            {
                int n = 0;

                for (int i = 0; i < countPerf; i++)
			    {

                    var ie = rows.Where(r => r.group > (i+1)*100 && r.group < (i+2)*100);
                    string [] str = new string[ie.Count()];
                    string val; int k = 0; string name = "Var" + String.Format("{0:00}",i);
                    foreach (var item in ie)
                    {
                        val = name + ';' + item.group + ';' + item.ToString();
                        str[k] = val; k++;
                        n++;
                    }
                    vd.AttribAdd(name, str);
                }
            }
        }

        public int loadFromAsm(AssemblyDocument asmDoc)
        {
            VariableDataForSpec vd = new VariableDataForSpec((Document)asmDoc);
            List<string> vals = new List<string>();
            List<string> cust = new List<string>();
            List<TableRow> rs = new List<TableRow>(12);
            vals = vd.AttribLoad("Var"); 
            if (vals.Count == 0)
            vals = getVar();
            cust = vd.AttribLoad("Custom");
            if (vals.Count == 0 && cust.Count == 0) return 0;
            foreach (var item in cust)
            {
                vals.Add(item); 
            }
            TableRow tr, trBase = null; int count = 0;
            for (int i = 0; i < vals.Count; i++)
            {
                string[] spl = vals[i].Split(';');
                tr = addRow(spl.Skip(2).Take(11).ToArray(),rows);
                if (tr.cells[2].value == "") tr.number = 0;
                else
                tr.number = int.Parse(tr.cells[2].value);
                tr.group = int.Parse(spl[1]);
                if (spl[0] == "Custom") tr.notNumber = true;

                trBase = rs.Find(e1 => e1.number == tr.number);
                if (trBase == null)
                {
                    rs.Add(tr);
                }
                else
                {
                    trBase.addVar(tr); tr.varRow = true;
                }

                //if (spl[0] == "Var00")
                //{
                //    rs.Add(tr);
                //}
                //else
                //{
                //    trBase = rs.Find(e => e.number == tr.number);
                //    if (trBase != null) {trBase.addVar(tr); tr.varRow = true;}
                //}
                //if (perf != spl[0])
                //{
                //    count++;
                //}
                //else
                //{

                //}
                if (spl[0].StartsWith("Var"))
                {
                    count = int.Parse(spl[0].Substring(3));
                }
            }
            return (count == 0)? count: count+1;
        }

        public static List<string> getVar()
        {
            string name = InvDoc.u.referendedDoc(Macros.StandardAddInServer.m_inventorApplication.ActiveDocument).FullDocumentName;
            //name = Regex.Replace(name, @"-\d\d", "");
            //name = Regex.Replace(name, @"\^\d\d", "");
            List<string> asms = TableInv.getAsms(name);
            if (asms.Count == 1) return new List<string>();
            return TableInv.loadFromAsm(asms);
        }

        public static List<string> getAsms(string name)
        {
            List<string> asms = new List<string>();
            asms.Add(name);
            Regex regex = new Regex(@"(-\d{0,3}\w{0,3}\d{0,3}\w{0,2}\.\d*\.*\d*)(-*\d*).*(^*\d*)");
            string v = "";
            Match m = regex.Match(name);
            if (m.Groups.Count > 1)
            v = m.Groups[1].Value;
            string p = System.IO.Path.GetDirectoryName(name);
            IEnumerable<string> files = System.IO.Directory.EnumerateFiles(p,"*.iam", System.IO.SearchOption.TopDirectoryOnly);
            foreach (var item in files)
            {
                m = regex.Match(item);
                if (m.Groups[1].Value == v && m.Groups[0].Value.IndexOf("^") != -1) 
                    asms.Add(item);
                //if (item.IndexOf(System.IO.Path.GetFileNameWithoutExtension(name) + "^") != -1)
                //if (regex.IsMatch(item)) 
                //    asms.Add(item); 
            }
            asms.Sort();
            return asms;
        }

        public static List<string> loadFromAsm(List<string> asmDocs)
        {
            int count = asmDocs.Count;
            Documents docs = Macros.StandardAddInServer.m_inventorApplication.Documents;
            NameValueMap nvmOptions = I.objs.CreateNameValueMap();
            nvmOptions.Add("SkipAllUnresolvedFiles", true);
            List<List<string>> lst = new List<List<string>>();
            foreach (var item in asmDocs)
            {
                AssemblyDocument asm = docs.OpenWithOptions(item, nvmOptions, false) as AssemblyDocument;
                BOM bom = asm.ComponentDefinition.BOM;
                BOMView bView = bom.BOMViews[1];
                lst.Add(loadFromBOM(bView));
            }
            List<string> fc = getVarFromAsm(lst);
            return fc;
        }

        public static List<string> loadFromBOM(BOMView bView)
        {
            List<string> lst = new List<string>();
            foreach (BOMRow row in bView.BOMRows)
            {
                String name = row.ReferencedFileDescriptor.FullFileName;  
                int count = row.ItemQuantity;
                lst.Add(name + ";" + count.ToString() + ";" + row.BOMStructure);
            }
            return lst;
        }

        public static List<string> getVarFromAsm(List<List<string>> lst)
        {
            int k = 0;
            List<string> first = lst[k];
            //HashSet<string> hs = new HashSet<string>();
            List<string> retLst = new List<string>();
            List<List<string>> tmp = new List<List<string>>();
            List<string> distFirst = first.Except<string>(lst[1]).ToList();
            tmp.Add(distFirst);
            //List<string> distFirst = uni.Except<string>(lst[0]).ToList();
            int num = first.Count - distFirst.Count()+2;
            //changeList(distFirst, num, 0);
            //retLst.AddRange(distFirst);
            for (int i = 1; i < lst.Count; i++)
            {
                //num++;
                List<string> dist = lst[i].Except<string>(first).ToList();
                tmp.Add(dist);
                if (dist.Count > distFirst.Count) { k = i; first = lst[k]; i++; }
                //changeList(dist, num, i);
                //retLst.AddRange(dist);
            }
            if (k != 0)
            {
                tmp.Clear();
                for (int i = 0; i < lst.Count; i++)
                {   
                    List<string> dist = lst[i].Except<string>(first).ToList();
                    if (i == k)
                    {
                        dist = first.Except<string>(lst[0]).ToList();
                    }
                    tmp.Add(dist);
                }
            }
            
            IEnumerable<string> uni = new List<string>();
             for (int j = 0; j < tmp.Count; j++)
             {
                 uni = uni.Union<string>(tmp[j]);
             }
            string s = ""; 
            for (int j = 0; j < tmp.Count; j++)
            {
                //retLst.AddRange(tmp[j]);
                foreach (var item in uni)
                {
                    s = lst[j].Find(el => el.StartsWith(item.Substring(0, item.IndexOf(";"))));
                    if (s != null && !tmp[j].Exists(e => e == s)) tmp[j].Add(s);
                }
            }
            for (int i = 0; i < tmp.Count; i++)
            {
                //num++;
                tmp[i].Sort(new compare());
                changeList(tmp[i], num, i);
                retLst.AddRange(tmp[i]);
            }
            return retLst;
        }

        public class compare: IComparer<string>
        {
            public int Compare(string x, string y)
            {
                var splx = x.Split(';');
                var sply = y.Split(';');
                x = splx[2] + System.IO.Path.GetFileNameWithoutExtension(splx[0]);
                y = sply[2] + System.IO.Path.GetFileNameWithoutExtension(sply[0]);
                return String.Compare(x, y);
            }
        }

        public static void changeList(List<string> lst, int num, int isp)
        {
            Documents docs = Macros.StandardAddInServer.m_inventorApplication.Documents;
            NameValueMap nvmOptions = I.objs.CreateNameValueMap();
            nvmOptions.Add("SkipAllUnresolvedFiles", true);
            for (int i = 0; i < lst.Count; i++)
			{
                var spl = lst[i].Split(';');
                string n = (isp + 1).ToString(), str = "";
                Document doc = docs.OpenWithOptions(spl[0],nvmOptions,false);
                string pn = InvDoc.u.getProp(doc, "Part Number").Value.ToString(),
                    desc = InvDoc.u.getProp(doc, "Description").Value.ToString(),
                    note = InvDoc.u.getProp(doc, "Comments").Value.ToString();
                str = num + ";" + pn + ";" + desc + ";" + spl[1];
                if (pn == "") n += "30";
                else 
                {
                    int l = pn.Length;
                    if (pn[l-3] == '-') pn = pn.Remove(l-3);
                    if (pn.EndsWith("0")) n += "10";
                    else n += "20";
                }
			    lst[i] = "Var" + isp.ToString("00") + ";" + n + ";;;" + str + ";" + note + ";;;;";
                num++;
			}
        }

        public int loadFromAsm(AssemblyDocument asmDoc, XElement el)
        {
            VariableDataForSpec vd = new VariableDataForSpec((Document)asmDoc);
            List<string> vals = new List<string>();
            vals = vd.AttribLoad("Var");
            if (vals.Count == 0) return 0;
            int count = 0;
            for (int i = 0; i < vals.Count; i++)
            {
                string[] spl = vals[i].Split(';');
                el.Add(new XElement ("VarData", new XAttribute("Group", spl[1]), new XAttribute("Position", spl[4]), new XAttribute("PartNumber", spl[5]),
                    new XAttribute("Description", spl[6]), new XAttribute("Count", spl[7]), new XAttribute("Note", ""), new XAttribute("Dxf", ""), new XAttribute("Thickness", ""),
                    new XAttribute("FPLength", ""), new XAttribute("FPWidth", "")));
                count++;
            }
            return count;
        }

        public void drawTable(string name, List<TableRow> rows, int rowStart, int rowEnd, Inventor.Point2d insPt, ref Point2d pt, int count = 2,double lw = 0.05,double offset = 0, bool lastLineW = false)
        {
            try
            {
                ssd = doc.SketchedSymbolDefinitions.Add(name);
            }
            catch
            {
                addToSheet(doc.ActiveSheet, name, ref pt, true);
                doc.SketchedSymbolDefinitions[name].Delete();
                ssd = doc.SketchedSymbolDefinitions.Add(name);
            }
            int countColumns = 7;
            if (cols.Count != countColumns)
            {
                cols.RemoveRange(countColumns, 4);
            }
            count += rowStart;
            ssd.Edit(out ds);

            double colsLen = cols.Sum(e => e.width);
            double rowsLen = 0;
            try
            {
            for (int i = rowStart; i < rowEnd; i++)
            {
                sl = ds.SketchLines.AddByTwoPoints(insPt, tg.CreatePoint2d(insPt.X - colsLen, insPt.Y));
                insPt.Y -= rows[i].height;
                rowsLen += rows[i].height;
                if (i < count) sl.LineWeight = lw;
            }
            sl = ds.SketchLines.AddByTwoPoints(insPt, tg.CreatePoint2d(insPt.X - colsLen, insPt.Y));
            if (lastLineW) sl.LineWeight = lw;

            insPt.X -= colsLen; insPt.Y = 0;
            for (int i = 0; i < cols.Count; i++)
            {
                sl = ds.SketchLines.AddByTwoPoints(insPt, tg.CreatePoint2d(insPt.X, insPt.Y - rowsLen));
                insPt.X += cols[i].width;
                sl.LineWeight = lw;
            }
            sl = ds.SketchLines.AddByTwoPoints(insPt, tg.CreatePoint2d(insPt.X, insPt.Y - rowsLen));
            sl.LineWeight = lw;
            insPt.X -= colsLen;
            for (int i = rowStart; i < rowEnd; i++)
            {
                var tr = rows[i];
                insPt.Y -= tr.height;
                for (int j = 0; j < countColumns; j++)
                {
                    var cell = tr.cells[j];
                    if (cell.value != "") addText(cell, insPt,rows);
                    insPt.X += cols[cell.colInd].width;
                }
                insPt.X -= colsLen;
            }
            ds.SketchPoints.Add(tg.CreatePoint2d(offset/10, 0)).InsertionPoint = true;
            ds.SketchPoints.Add(tg.CreatePoint2d(0, -rowsLen)).ConnectionPoint = true;
            ssd.ExitEdit();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }
        private Inventor.TextBox addText(TableCell cell, Inventor.Point2d insPt, List<TableRow> rows)
        {
            Inventor.Point2d pt1, pt2;
            pt1 = tg.CreatePoint2d(insPt.X + cell.paddingLeft, insPt.Y + cell.paddingBottom);
            pt2 = tg.CreatePoint2d(insPt.X + cols[cell.colInd].width - cell.paddingRight, insPt.Y + rows[cell.rowInd].height - cell.paddingTop);
            //if (cell.rotation != 0) return formatText(cell, tg.CreatePoint2d(pt2.X, pt1.Y), tg.CreatePoint2d(pt1.X, pt2.Y));
            return formatText(cell, pt1, pt2);
        }
        private Inventor.TextBox formatText(TableCell cell, Inventor.Point2d pt1, Inventor.Point2d pt2, TextStyle ts = null)
        {
            string format = "";
            if (cell.underline) format += " Underline ='True' ";
            if (cell.bold) format += " Bold = 'True' ";
            if (!cell.italic) format += " Italic = 'False' ";
            if (cell.font != "GOST Common") format += @" Font = '" + cell.font + @"' ";
            //if (cell.fontSize != 3.5) format += @" FontSize = '" + (cell.fontSize / 10).ToString() + @"' ";
            if (format != "")
                format = @"<StyleOverride" + format + @">" + cell.value + @"</StyleOverride>";
            else format = cell.value;
            if (cell.value.IndexOf('$') != -1) format = formatString(cell, '$');
            tb = ds.TextBoxes.AddByRectangle(pt1, pt2, format);
            if (ts != null) tb.Style = ts;
            if (cell.fontSize != 3.5) tb.Style.FontSize = cell.fontSize / 10;

            tb.SingleLineText = cell.single;
            tb.HorizontalJustification = (Inventor.HorizontalTextAlignmentEnum)cell.horAlign;
            tb.VerticalJustification = (Inventor.VerticalTextAlignmentEnum)cell.vertAlign;
            tb.Rotation = cell.rotation;
            tb.WidthScale = cell.widthScale;
            double l = pt2.X - pt1.X;
            if (cell.rotation != 0) l = pt2.Y - pt1.Y;
            if (tb.FittedTextWidth > l) tb.WidthScale = l / tb.FittedTextWidth * (cell.widthScale - 0.05); 
            return tb;
        }
        private string formatString(TableCell txt, char c)
        {
            string sub1, sub2;
            int ind = txt.value.IndexOf(c);
            sub1 = txt.value.Substring(0, ind);
            sub2 = txt.value.Substring(ind, txt.value.Length - ind);
            if (sub1.LastIndexOf(' ') != -1)
                sub1 = sub1.Insert(sub1.LastIndexOf(' ') + 1, "<Stack FractionalTextScale='1'>");
            if (sub2.IndexOf(' ') != -1)
                sub2 = sub2.Insert(sub2.IndexOf(' '), "</Stack>");
            sub2 = sub2.Replace('$', '/');
            return sub1 + sub2;
        }
        //private void formatText(ref Inventor.TextBox _TB, string align, double widthCol, double textH = 0.35, string font = "GOST Common",
        //        double width = 1, bool underline = false, bool italic = true, bool bold = false)
        //{
        //    string format = ""; string txt = "";

        //    txt = _TB.Text;
        //    _TB.WidthScale = width;
        //    _TB.VerticalJustification = VerticalTextAlignmentEnum.kAlignTextMiddle;
        //    switch (align)
        //    {
        //        case "По центру":
        //            _TB.HorizontalJustification = HorizontalTextAlignmentEnum.kAlignTextCenter;
        //            break;
        //        case "Слева":
        //            _TB.HorizontalJustification = HorizontalTextAlignmentEnum.kAlignTextLeft;
        //            break;
        //        case "Справа":
        //            _TB.HorizontalJustification = HorizontalTextAlignmentEnum.kAlignTextRight;
        //            break;
        //    }
        //    if (underline)
        //    {
        //        format += " Underline ='True' ";
        //    }
        //    if (!italic)
        //    {
        //        format += " Italic = 'False' ";
        //    }
        //    if (bold)
        //    {
        //        format += " Bold = 'True' ";
        //    }
        //    if (font != "GOST Common")
        //    {
        //        format += @" Font = '" + font.ToString() + @"' ";
        //    }

        //    if (Math.Round(textH * 100) != 35)
        //    {
        //        format += @" FontSize = '" + textH.ToString() + @"' ";
        //    }

        //    double w = widthCol / _TB.FittedTextWidth * (width - 0.05);

        //    //if (setting.checkBox3.Checked && _TB.FittedTextWidth > widthCol)
        //    //{
        //    //    _TB.SingleLineText = true;
        //    //    _TB.WidthScale = widthCol / _TB.FittedTextWidth * (width - 0.05);
        //    //}

        //    if (_TB.Text.IndexOf('/') != -1)
        //    {
        //        string sub1, sub2;
        //        int ind = txt.IndexOf('/');
        //        sub1 = txt.Substring(0, ind);
        //        sub2 = txt.Substring(ind, txt.Length - ind);
        //        if (sub1.LastIndexOf(' ') != -1)
        //            sub1 = sub1.Insert(sub1.LastIndexOf(' ') + 1, "<Stack FractionalTextScale='1'>");
        //        if (sub2.IndexOf(' ') != -1)
        //            sub2 = sub2.Insert(sub2.IndexOf(' '), "</Stack>");
        //        txt = sub1 + sub2;
        //        _TB.FormattedText = txt;
        //    }

        //    if (_TB.Text.IndexOf('^') != -1)
        //    {
        //        string sub1, sub2;
        //        int ind = txt.IndexOf('^');
        //        sub1 = txt.Substring(0, ind);
        //        sub2 = txt.Substring(ind, txt.Length - ind);
        //        //if (sub1.LastIndexOf(' ') != -1)
        //        //  sub1 = sub1.Insert(sub1.LastIndexOf(' ') + 1, "<Stack FractionalTextScale='0.7'>");
        //        if (sub2.IndexOf(' ') != -1)
        //            sub2 = sub2.Insert(sub2.IndexOf(' '), "</Stack>");
        //        else sub2 += "</Stack>";
        //        txt = sub1 + "<Stack>" + sub2;
        //        _TB.FormattedText = txt;
        //    }

        //    if (format != "")
        //        _TB.FormattedText = @"<StyleOverride" + format + @">" + txt + @"</StyleOverride>";
        //}
    }
}
