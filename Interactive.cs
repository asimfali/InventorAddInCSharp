using System;
using InvDoc;
using Inventor;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using InterfaceDll;
using System.Xml.Linq;

namespace InvAddIn
{
    class Interactive : System.Windows.Forms.Form
    {
        Document doc, doc1;
        PartComponentDefinition def;
        CommandManager mgr;
        SelectSet ss;
        object selEnt = null;
        UserInterfaceManager uiMgr;
        InteractionEvents evts;
        SelectEvents sEv;
        KeyboardEvents kEv;
        WorkPlane wp;
        bool EvInt = false, term = false;
        List<object> elems = new List<object>();
        composeIM im;
        ApplicationEvents aev;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.Button button3;
        private System.Windows.Forms.Button button4;
        private System.Windows.Forms.Button button5;
        private System.Windows.Forms.Button button6;
        private System.Windows.Forms.Button button1;

        private void InitializeComponent()
        {
            this.button1 = new System.Windows.Forms.Button();
            this.button2 = new System.Windows.Forms.Button();
            this.button3 = new System.Windows.Forms.Button();
            this.button4 = new System.Windows.Forms.Button();
            this.button5 = new System.Windows.Forms.Button();
            this.button6 = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(0, 0);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(75, 23);
            this.button1.TabIndex = 0;
            this.button1.Text = "КП Copy";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(81, 0);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(91, 23);
            this.button2.TabIndex = 1;
            this.button2.Text = "КП в сборке";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // button3
            // 
            this.button3.Location = new System.Drawing.Point(0, 29);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(172, 23);
            this.button3.TabIndex = 2;
            this.button3.Text = "Копировать свойства (C)";
            this.button3.UseVisualStyleBackColor = true;
            this.button3.Click += new System.EventHandler(this.button3_Click_1);
            // 
            // button4
            // 
            this.button4.Location = new System.Drawing.Point(177, 0);
            this.button4.Name = "button4";
            this.button4.Size = new System.Drawing.Size(93, 23);
            this.button4.TabIndex = 3;
            this.button4.Text = "Размеры (D)";
            this.button4.UseVisualStyleBackColor = true;
            this.button4.Click += new System.EventHandler(this.button4_Click);
            // 
            // button5
            // 
            this.button5.Location = new System.Drawing.Point(178, 29);
            this.button5.Name = "button5";
            this.button5.Size = new System.Drawing.Size(92, 23);
            this.button5.TabIndex = 4;
            this.button5.Text = "Овалы (O)";
            this.button5.UseVisualStyleBackColor = true;
            this.button5.Click += new System.EventHandler(this.button5_Click);
            // 
            // button6
            // 
            this.button6.Location = new System.Drawing.Point(178, 58);
            this.button6.Name = "button6";
            this.button6.Size = new System.Drawing.Size(92, 23);
            this.button6.TabIndex = 5;
            this.button6.Text = "Обновить (U)";
            this.button6.UseVisualStyleBackColor = true;
            this.button6.Click += new System.EventHandler(this.button6_Click);
            // 
            // Interactive
            // 
            this.ClientSize = new System.Drawing.Size(276, 261);
            this.Controls.Add(this.button6);
            this.Controls.Add(this.button5);
            this.Controls.Add(this.button4);
            this.Controls.Add(this.button3);
            this.Controls.Add(this.button2);
            this.Controls.Add(this.button1);
            this.Name = "Interactive";
            this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.Interactive_KeyDown);
            this.ResumeLayout(false);

        }

        public Interactive(Document d)
        {
            doc = d;
            if (doc.DocumentType == DocumentTypeEnum.kPartDocumentObject)
                def = ((PartDocument)doc).ComponentDefinition;
            mgr = I.app.CommandManager;
            aev = I.app.ApplicationEvents;
            ss = doc.SelectSet;
            uiMgr = I.app.UserInterfaceManager;
            InitializeComponent();
            KeyPreview = true;
        }

        private void changeDoc(_Document DocumentObject, EventTimingEnum BeforeOrAfter, NameValueMap Context, out HandlingCodeEnum HandlingCode)
        {
            HandlingCode = HandlingCodeEnum.kEventNotHandled;
            if (doc.Equals(DocumentObject) || BeforeOrAfter != EventTimingEnum.kBefore) return;
            doc1 = DocumentObject;
            HandlingCode = HandlingCodeEnum.kEventHandled;
            aev.OnActivateDocument -= changeDoc;
            imate();
            this.Close();
        }

        private void imate()
        {
            doc1.Activate();
            var m = new CopyIMate(doc1, im);
        }

        private void sImate(object o)
        {
            im = new composeIM(o, true);
            I.app.StatusBarText = "Выберите документ";
            aev.OnActivateDocument += changeDoc;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            imateBtn();
        }

        void imateBtn()
        {
            this.Hide();
            if (ss.Count != 1)
            {
                evts = mgr.CreateInteractionEvents();
                sEv = evts.SelectEvents;
                sEv.SingleSelectEnabled = true;
                sEv.OnPreSelect += selEv;
                evts.Start();
                var bn = withBP(" пары");
                bn.Expanded = true;
                System.Windows.Forms.MessageBox.Show("Выберите КП");
                while (!EvInt) uiMgr.DoEvents();
                //System.Windows.Forms.MessageBox.Show("Не была выбрана КП");
                System.Windows.Forms.MessageBox.Show("Перейдите в другой документ");
                this.Close();
                //return; 
            }
            if (ss.Count != 0)
                sImate(ss[1]);
            else if (selEnt != null)
                sImate(selEnt);
        }

        private void withBN(BrowserNode n, Func<BrowserNode, bool> f, ref BrowserNode r)
        {
            foreach (BrowserNode item in n.BrowserNodes)
            {
                if (item.BrowserNodes.Count > 0) withBN(item, f, ref r);

                if (f(item))
                {
                    r = item; return;
                }
            }
        }

        private void selEv(ObjectsEnumerator JustSelectedEntities, SelectionDeviceEnum SelectionDevice,
            Point ModelPosition, Point2d ViewPosition, View View)
        {

        }

        private void keyOp(int KeyASCII)
        {
            switch (KeyASCII)
            {
                case 32:
                    break;
                case 13:
                    EvInt = true;
                    break;
                default:
                    break;
            }
        }

        private BrowserNode withBP(string txt)
        {
            var n = doc.BrowserPanes.ActivePane.TopNode;
            BrowserNode bn = null; withBN(n, e => e.FullPath.EndsWith(txt), ref bn);
            return bn;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            asmBtn();
        }

        void asmBtn()
        {
            iMateDefinition m = null;
            this.Hide();
            if (ss.Count != 1)
            {
                System.Windows.Forms.MessageBox.Show("Выберите КП или часть сборки");
                this.Close();
                return;
            }
            var occ = ss[1] as ComponentOccurrence;
            if (occ != null)
            {
                var bn = withBP(occ._DisplayName);
                withBN(bn, el => el.FullPath.EndsWith(" пары"), ref bn);
                bn.Expanded = true;
                evts = mgr.CreateInteractionEvents();
                sEv = evts.SelectEvents;
                sEv.SingleSelectEnabled = true;
                sEv.OnPreSelect += selEv;
                evts.Start();
                System.Windows.Forms.MessageBox.Show("Выберите КП");
                while (!EvInt) uiMgr.DoEvents();
            }
            if (ss[1] is iMateDefinition) selEnt = ss[1];
            im = new composeIM(selEnt, true);
            doc1 = doc.ActivatedObject as Document;
            var t = new CopyIMate(doc1, im);
            this.Close();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            //this.Hide();
            //var drws = getDocs(DocumentTypeEnum.kDrawingDocumentObject);
            //var prts = getDocs(DocumentTypeEnum.kPartDocumentObject);
            //if (drws == null || prts == null) return;
            //DrViews v = new DrViews(drws.ElementAt(0) as DrawingDocument, prts.ElementAt(0));
            //this.Close();
        }

        private IEnumerable<Document> getDocs(DocumentTypeEnum t)
        {
            return u.gets<Document>(I.visDocs(), f => f.DocumentType == t);
        }

        private void button3_Click_1(object sender, EventArgs e)
        {
            //holesBtn();
            createCopy();
        }
        void holesBtn()
        {
            this.Close();
            while (!doEvent()) ;
        }
        bool doEvent()
        {
            initEvents("Выберите ребро", new SelectionFilterEnum[] {
            SelectionFilterEnum.kPartFeatureFilter});
            bool r = run();
            termEvents();
            wp = null;
            return r;
        }

        private void initEvents(string txt, SelectionFilterEnum[] f)
        {
            if (evts == null) evts = mgr.CreateInteractionEvents();
            evts.InteractionDisabled = false;
            sEv = evts.SelectEvents;
            foreach (var item in f)
            {
                sEv.AddSelectionFilter(item);
            }
            evts.StatusBarText = txt;
            sEv.OnPreSelect += Sel_OnPreSelect;
            sEv.OnSelect += Sel_OnSelect;
            kEv = evts.KeyboardEvents;
            kEv.OnKeyDown += KeyboardEvents_OnKeyDown;
            evts.Start();
        }

        private void KeyboardEvents_OnKeyDown(int Key, ShiftStateEnum ShiftKeys)
        {
            if (Key == 32)
            {
                EvInt = true;
            }
            if (Key == 27)
            {
                EvInt = true;
                term = true;
            }
            if (Key == 77)
            {
                u.changePlane(ref wp, def);
            }
        }

        private void termEvents()
        {
            evts.Stop();
            sEv.OnPreSelect -= Sel_OnPreSelect;
            sEv.OnSelect -= Sel_OnSelect;
            kEv.OnKeyDown -= KeyboardEvents_OnKeyDown;
            elems.Clear();
            sEv = null; kEv = null; evts = null;
        }

        private bool run()
        {
            while (!EvInt) uiMgr.DoEvents();
            if (term) return true;
            foreach (var item in sEv.SelectedEntities)
            {
                elems.Add(item);
            }
            foreach (var item in elems)
            {
                PartFeature pf = null;
                Face f = item as Face;
                if (f != null)
                {
                    pf = f.CreatedByFeature as PartFeature;
                }
                if (pf == null) pf = item as PartFeature;
                if (pf != null)
                {
                    var h = u.addHole(pf.Parent as SheetMetalComponentDefinition, pf);
                    if (wp != null)
                        u.addMirror(doc, wp, new List<PartFeature>() { h as PartFeature });
                }
            }
            elems.Clear();
            EvInt = false;
            return false;
        }

        private void Sel_OnSelect(ObjectsEnumerator JustSelectedEntities, SelectionDeviceEnum SelectionDevice,
            Point ModelPosition, Point2d ViewPosition, View View)
        {

        }

        private void Sel_OnPreSelect(ref object PreSelectEntity, out bool DoHighlight,
            ref ObjectCollection MorePreSelectEntities, SelectionDeviceEnum SelectionDevice,
            Point ModelPosition, Point2d ViewPosition, View View)
        {
            DoHighlight = true;
            if (!(PreSelectEntity is HoleFeature ||
                PreSelectEntity is PunchToolFeature)) DoHighlight = false;
        }

        private void button4_Click(object sender, EventArgs e)
        {
            createDims();
        }

        void createDims()
        {
            bool ret = true;
            while (ret)
            {
                this.Close();
                Document doc = I.aDoc();
                InteractiveDims dims = new InteractiveDims(doc);
                dims.fill();
                ret = dims.run();
                dims.clear();
            }
        }

        void createCopy()
        {
            bool ret = true;
            while (ret)
            {
                this.Close();
                Document doc = I.aDoc();
                CopyProps props = new CopyProps(doc);
                props.fill();
                ret = props.run();
                props.clear();
            }
        }

        public void getDoc(DrawingCurve dc, ref Document doc, ref Edge e, ref SketchEntity se)
        {
            e = dc.ModelGeometry as Edge;
            if (e == null)
            {
                se = dc.ModelGeometry as SketchEntity;
                var fp = se.Parent.Parent as FlatPattern;
                doc = fp.Document as Document;
                return;
            }
            doc = e.Parent.ComponentDefinition.Document as Document;
        }

        void updateSlot()
        {
            this.Close();
            DrawingDocument doc = I.aDoc() as DrawingDocument;
            if (doc == null) return;
            Sheet sh = doc.ActiveSheet;
            foreach (DrawingDimension item in sh.DrawingDimensions)
            {
                var t = item.Text.FormattedText.ToLower();
                if (t.IndexOf("отв") != -1 || t.IndexOf("шт") != -1)
                {
                    if (item.Attached == false) continue;
                    GeometryIntent gi = null;
                    DiameterGeneralDimension dim = item as DiameterGeneralDimension;
                    if (dim != null) gi = dim.Intent;
                    RadiusGeneralDimension r = item as RadiusGeneralDimension;
                    if (r != null) gi = r.Intent;
                    if (gi == null) return;
                    DrawingCurve dc = gi.Geometry as DrawingCurve;
                    string txt = getData(dc);
                    if (r != null)
                    {
                        r.Text.FormattedText = txt;
                    }
                    else if (dim != null)
                    {
                        dim.Text.FormattedText = txt + "<DimensionValue/>";
                    }
                }
            }
        }

        public string getData(DrawingCurve dc)
        {
            Document d = null; Edge e = null; SketchEntity se = null;
            getDoc(dc, ref d, ref e, ref se);
            string txt = null;
            if (se != null)
            {
                SketchModel sm = new SketchModel(d);
                sm.fill(se);
                txt = sm.txt();
            }
            if (dc.CurveType == CurveTypeEnum.kCircleCurve && e != null)
            {
                HoleModel sm = new HoleModel(d);
                sm.fill(e);
                txt = sm.txt();

            }
            else if (dc.CurveType == CurveTypeEnum.kCircularArcCurve && e != null)
            {
                SlotModel sm = new SlotModel(d);
                sm.fill(e);
                txt = sm.txt();
            }
            return txt;
        }

        void addSlot()
        {
            this.Close();
            Document doc = I.aDoc();
            bool ret = true;
            while (ret)
            {
                SlotDraw sl = new SlotDraw(doc);
                sl.fill();
                ret = sl.run();
                if (sl.dcs == null) return;
                var dcs = sl.dcs;
                string txt = getData(dcs.Parent);
                if (dcs.GeometryType == Curve2dTypeEnum.kCircularArcCurve2d)
                {
                    var sh = sl.sh;
                    var pt1 = I.CP2d(sl.p1.X, sl.p1.Y);
                    var gi = sh.CreateGeometryIntent(sl.dcs.Parent, pt1);
                    var pt2 = I.CP2d(sl.p2.X, sl.p2.Y);
                    //sh.DrawingNotes.GeneralNotes.AddFitted(pt, txt);
                    var dim = sh.DrawingDimensions.GeneralDimensions.AddRadius(pt2, gi);
                    dim.HideValue = true;
                    dim.Text.FormattedText = txt;
                }
                else if (dcs.GeometryType == Curve2dTypeEnum.kCircleCurve2d)
                {
                    var sh = sl.sh;
                    var pt1 = I.CP2d(sl.p1.X, sl.p1.Y);
                    var gi = sh.CreateGeometryIntent(sl.dcs.Parent, pt1);
                    var pt2 = I.CP2d(sl.p2.X, sl.p2.Y);
                    //sh.DrawingNotes.GeneralNotes.AddFitted(pt, txt);
                    var dim = sh.DrawingDimensions.GeneralDimensions.AddDiameter(pt2, gi);
                    //dim.HideValue = true;
                    dim.LeaderFromCenter = true;
                    dim.SingleDimensionLine = false;
                    dim.Text.FormattedText = txt + dim.Text.FormattedText;
                }
                sl.clearClient();
                sl.clear();
            }
        }

        private void Interactive_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
        {
            switch (e.KeyCode)
            {
                case System.Windows.Forms.Keys.O:
                    this.addSlot();
                    break;
                case System.Windows.Forms.Keys.U:
                    this.updateSlot();
                    break;
                case System.Windows.Forms.Keys.D:
                    createDims();
                    break;
                case System.Windows.Forms.Keys.C:
                    this.createCopy();
                    break;
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            addSlot();
        }

        private void button6_Click(object sender, EventArgs e)
        {
            updateSlot();
        }

        private void selectMulti()
        {
            //mgr.ClearPrivateEvents();
            if (evts == null)
                evts = mgr.CreateInteractionEvents();
            //evts.InteractionDisabled = false;
            sEv = evts.SelectEvents;
            kEv = evts.KeyboardEvents;
            sEv.OnSelect += selEv;
            kEv.OnKeyPress += keyOp;
            evts.Start();
            while (!EvInt)
            {
                uiMgr.DoEvents();
            }
            evts.Stop();
            kEv.OnKeyPress -= keyOp;
            sEv.OnSelect -= selEv;
            sEv = null; kEv = null; evts = null;
        }

        private void selEv(ref object PreSelectEntity, out bool DoHighlight,
            ref ObjectCollection MorePreSelectEntities, SelectionDeviceEnum SelectionDevice,
            Point ModelPosition, Point2d ViewPosition, View View)
        {
            DoHighlight = true;
            if (PreSelectEntity is CompositeiMateDefinition)
            {
                sEv.OnPreSelect -= selEv;
                evts.Stop();
                sEv = null;
                EvInt = true;
                selEnt = PreSelectEntity;
                //sImate(PreSelectEntity);
            }
        }
    }

    public class DrViews
    {
        DrawingDocument drw, nDrw;
        DrawingViews dvs, ndvs;
        Document doc;
        string path;
        public DrViews(DrawingDocument drw, Document doc)
        {
            this.drw = drw; this.doc = doc;
            path = file.p(doc.FullDocumentName);
            nDrw = I.addDrw(path + u.nameForSave(doc, true));
            addViews();
            I.open(nDrw.FullDocumentName, true);
        }
        public void addViews()
        {
            dvs = drw.Sheets[1].DrawingViews;
            ndvs = nDrw.Sheets[1].DrawingViews;
            foreach (DrawingView view in dvs)
            {
                addView(view);
            }
        }
        public void addView(DrawingView v)
        {
            if (v.ParentView == null)
            {
                object cam = null;
                var o = v.Camera.ViewOrientationType;
                if (o == ViewOrientationTypeEnum.kArbitraryViewOrientation)
                    ndvs.AddBaseView((_Document)doc, v.Center, v.Scale, ViewOrientationTypeEnum.kArbitraryViewOrientation,
                        v.ViewStyle, ArbitraryCamera: v.Camera);
                else
                    ndvs.AddBaseView((_Document)doc, v.Center, v.Scale, v.Camera.ViewOrientationType, v.ViewStyle);
            }
            else if (v.ParentView != null)
                ndvs.AddProjectedView(ndvs[ndvs.Count], v.Center, v.ViewStyle, v.Scale);
        }
    }

    public class EventsDims
    {
        CommandManager mgr;
        public InteractionEvents evts;
        UserInterfaceManager uiMgr;
        SelectEvents sEv;
        KeyboardEvents kEv;
        public MouseEvents mEv;
        public string txt;
        public SelectionFilterEnum[] filter;
        public bool EvInt = false, term = false;
        bool me, pe, se, ke;
        SelectEventsSink_OnPreSelectEventHandler ph;
        KeyboardEventsSink_OnKeyDownEventHandler kh;
        SelectEventsSink_OnSelectEventHandler sh;
        MouseEventsSink_OnMouseDoubleClickEventHandler mh;
        MouseEventsSink_OnMouseClickEventHandler mc;
        public MouseEventsSink_OnMouseMoveEventHandler mm;
        List<object> elems = new List<object>();
        public EventsDims(string prompt, SelectionFilterEnum[] fs)
        {
            mgr = I.app.CommandManager;
            uiMgr = I.app.UserInterfaceManager;
            txt = prompt;
            filter = fs;
            initEvents();
        }

        public void setWindow()
        {
            sEv.WindowSelectEnabled = true;
        }

        public void initEvents()
        {
            if (evts == null) evts = mgr.CreateInteractionEvents();
            evts.InteractionDisabled = false;
            sEv = evts.SelectEvents;
            foreach (var item in filter)
            {
                sEv.AddSelectionFilter(item);
            }
            evts.StatusBarText = txt;
            kEv = evts.KeyboardEvents;
            mEv = evts.MouseEvents;
            kEv.OnKeyDown += KeyboardEvents_OnKeyDown;
        }

        public void termEvents()
        {
            evts.Stop();
            if (ph != null) sEv.OnPreSelect -= ph;
            if (sh != null) sEv.OnSelect -= sh;
            if (kh != null) kEv.OnKeyDown -= kh;
            if (mh != null) mEv.OnMouseDoubleClick -= mh;
            if (mm != null) mEv.OnMouseMove -= mm;
            if (mc != null) mEv.OnMouseClick -= mc;
            kEv.OnKeyDown -= KeyboardEvents_OnKeyDown;
            elems.Clear();
            sEv = null; kEv = null; evts = null;
            ph = null; sh = null; kh = null; mh = null; mm = null; mc = null;
        }
        public bool run(Func<object, bool> func, Action<InteractionEvents> a = null)
        {
            evts.Start();
            if (a != null) a(evts);
            while (!EvInt) uiMgr.DoEvents();
            if (term) return true;
            foreach (var item in sEv.SelectedEntities)
            {
                elems.Add(item);
            }
            foreach (var item in elems)
            {
                func(item);
            }
            elems.Clear();
            EvInt = false;
            return false;
        }
        public void setPreSelectEv(SelectEventsSink_OnPreSelectEventHandler h)
        {
            sEv.OnPreSelect += h; ph = h;
        }
        public void setKeyEv(KeyboardEventsSink_OnKeyDownEventHandler h)
        {
            kEv.OnKeyDown += h; kh = h;
        }
        public void setSelectEv(SelectEventsSink_OnSelectEventHandler h)
        {
            sEv.OnSelect += h; sh = h;
        }
        public void setMouseEv(MouseEventsSink_OnMouseDoubleClickEventHandler h)
        {
            mEv.OnMouseDoubleClick += h; mh = h;
        }

        public void setMouseClick(MouseEventsSink_OnMouseClickEventHandler h)
        {
            mEv.OnMouseClick += h; mc = h;
        }

        public void setMouseMove(MouseEventsSink_OnMouseMoveEventHandler h)
        {
            mEv.OnMouseMove += h; mm = h;
        }

        private void KeyboardEvents_OnKeyDown(int Key, ShiftStateEnum ShiftKeys)
        {
            if (Key == 32)
            {
                EvInt = true;
            }
            if (Key == 27)
            {
                EvInt = true;
                term = true;
            }
        }
    }

    public class MyEventArgs : EventArgs
    {
        public bool ret;
        public MyEventArgs(bool r)
        {
            ret = r;
        }
    }

    public class CopyProps
    {
        DrawingDocument doc;
        Sheet sh;
        SelectSet ss;
        Point2d pt;
        double r = 0;
        int ind = 0;
        static public HighlightSet hs;
        object selEnt = null;
        int num = 0;
        double text_height = 0.7, offset = 0.3;
        Vector2d dir;
        EventsDims cur = null;
        Point2d cenPt = null;
        List<EventsDims> evts = new List<EventsDims>();
        public event EventHandler<MyEventArgs> change;

        List<DrawingCurve> dcs = new List<DrawingCurve>();
        string name = "dim", colName = "dimCol";
        LineTypeEnum lt;
        double lw;
        Color col;
        DrawingCurve dc;
        DrawingView dv;
        bool mouse = false;
        public CopyProps(Document doc)
        {
            if (doc.DocumentType != DocumentTypeEnum.kDrawingDocumentObject) return;
            this.doc = doc as DrawingDocument;
            //sh = this.doc.ActiveSheet;
        }
        public bool run()
        {
            for (int i = 0; i < evts.Count; i++)
            {
                cur = evts[i];
                if (cur.mm != null)
                {
                    cur.mEv.MouseMoveEnabled = true;
                }
                cur.run(el => set(el), a => copy(a));
                cur.termEvents();
                if (cur.term) return false;
                num++;
            }
            clear();
            return true;
        }

        void copy(InteractionEvents i)
        {
            var e = i.SelectEvents;
        }


        public bool set(object o)
        {
            return true;
        }

        public EventsDims add(string txt, SelectionFilterEnum[] f)
        {
            var ev = new EventsDims(txt, f);
            evts.Add(ev);
            return ev;
        }
        public void onChange(bool r)
        {
            MyEventArgs a = new MyEventArgs(r);
            change(this, a);
        }

        public void fill()
        {
            //u.clearClientLine(doc, name, colName);
            var ev = add("Выберите линию", new SelectionFilterEnum[] {
            SelectionFilterEnum.kDrawingDefaultFilter});
            ev.setPreSelectEv(Sel_OnPreSelectDC);
            ev.setSelectEv(Sel_OnSelect);
            ev.setKeyEv(KeyboardEvents_OnKeyDown);

            //ev = add("Выберите линию", new SelectionFilterEnum[]
            //{
            //    SelectionFilterEnum.kDrawingCurveSegmentFilter
            //});
            //ev.setPreSelectEv(Sel_OnPreSelectDC);
            //ev.setSelectEv(Sel_OnSelect);

            ev = add("Выберите куда копировать свойсва", new SelectionFilterEnum[] {
            SelectionFilterEnum.kDrawingDefaultFilter});
            this.cur = ev;
            ev.setWindow();
            ev.setPreSelectEv(Sel_OnPreSelectDC);
            ev.setSelectEv(Sel_OnSelect1);
            ev.setKeyEv(this.KeyboardEvents_OnKeyDown);
        }
        private void Sel_OnSelect(ObjectsEnumerator JustSelectedEntities, SelectionDeviceEnum SelectionDevice,
            Point ModelPosition, Point2d ViewPosition, View View)
        {
            foreach (var o in JustSelectedEntities)
            {
                if (o is DrawingCurveSegment)
                {
                    dc = ((DrawingCurveSegment)o).Parent;
                    lt = dc.Segments[1].Layer.LineType;
                    lw = dc.Segments[1].Layer.LineWeight;
                    col = dc.Segments[1].Layer.Color;
                    if (!dc.Color.Equals(col)) col = dc.Color;
                    if (dc.LineType != lt) lt = dc.LineType;
                    if (dc.LineWeight != lw) lw = dc.LineWeight;
                    evts[num].EvInt = true;
                }
            }
        }
        private void Sel_OnSelect1(ObjectsEnumerator JustSelectedEntities, SelectionDeviceEnum SelectionDevice,
            Point ModelPosition, Point2d ViewPosition, View View)
        {
            foreach (var o in JustSelectedEntities)
            {
                if (o is DrawingCurveSegment)
                {
                    dc = ((DrawingCurveSegment)o).Parent;
                    dcs.Add(dc);
                }
            }
        }
        public void clear()
        {
            evts.Clear();
            if (hs != null) hs.Clear();
        }

        private void Sel_OnPreSelectDC(ref object PreSelectEntity, out bool DoHighlight,
            ref ObjectCollection MorePreSelectEntities, SelectionDeviceEnum SelectionDevice,
            Point ModelPosition, Point2d ViewPosition, View View)
        {
            DoHighlight = true;
            if (!(PreSelectEntity is DrawingCurveSegment)) DoHighlight = false;
        }
        private void KeyboardEvents_OnKeyDown(int Key, ShiftStateEnum ShiftKeys)
        {
            if (Key == 77)
            {
                this.mouse = !this.mouse;
            }
            if (Key == 13)
            {
                if (dcs.Count > 0 && dc != null)
                {
                    apply();
                }
                cur.EvInt = true;
            }
        }
        void apply()
        {
            foreach (DrawingCurve s in dcs)
            {
                s.LineType = lt;
                s.LineWeight = lw;
                s.Color = col;
            }
        }
    }

    public class InteractiveDims
    {
        DrawingDocument doc;
        Sheet sh;
        SelectSet ss;
        Point2d pt;
        double r = 0;
        int ind = 0;
        static public HighlightSet hs;
        object selEnt = null;
        int num = 0;
        double text_height = 0.7, offset = 0.3;
        Vector2d dir;
        EventsDims cur = null;
        Point2d cenPt = null;
        List<EventsDims> evts = new List<EventsDims>();
        public event EventHandler<MyEventArgs> change;
        List<Centermark> cen = new List<Centermark>();
        List<DrawingCurve> dcs;
        string name = "dim", colName = "dimCol";
        List<Vector2d> dirs2d = new List<Vector2d>() { I.CV2d(1, 0), I.CV2d(0, 1) };
        DrawingCurve dc;
        DrawingView dv;
        bool mouse = false;

        public InteractiveDims(Document doc)
        {
            if (doc.DocumentType != DocumentTypeEnum.kDrawingDocumentObject) return;
            this.doc = doc as DrawingDocument;
            sh = this.doc.ActiveSheet;
        }
        public bool run()
        {
            for (int i = 0; i < evts.Count; i++)
            {
                cur = evts[i];
                if (cur.mm != null)
                {
                    cur.mEv.MouseMoveEnabled = true;
                }
                cur.run(el => set(el), a => draw(a.InteractionGraphics));
                cur.termEvents();
                if (cur.term) return false;
                num++;
            }
            clear();
            return true;
        }

        public bool set(object o)
        {
            return true;
        }

        public EventsDims add(string txt, SelectionFilterEnum[] f)
        {
            var ev = new EventsDims(txt, f);
            evts.Add(ev);
            return ev;
        }
        public void onChange(bool r)
        {
            MyEventArgs a = new MyEventArgs(r);
            change(this, a);
        }

        public void fill()
        {
            u.clearClientLine(doc, name, colName);
            var ev = add("Выберите центры", new SelectionFilterEnum[] {
            SelectionFilterEnum.kDrawingCurveSegmentFilter});
            ev.setPreSelectEv(Sel_OnPreSelectCM);
            ev.setSelectEv(Sel_OnSelect);
            ev.setKeyEv(KeyboardEvents_OnKeyDown);

            //ev = add("Выберите линию", new SelectionFilterEnum[]
            //{
            //    SelectionFilterEnum.kDrawingCurveSegmentFilter
            //});
            //ev.setPreSelectEv(Sel_OnPreSelectDC);
            //ev.setSelectEv(Sel_OnSelect);

            ev = add("Укажите положение размера", new SelectionFilterEnum[] {
            SelectionFilterEnum.kDrawingDefaultFilter});
            this.cur = ev;
            ev.setKeyEv(this.KeyboardEvents_OnKeyDown1);
            ev.setMouseClick(MEv_OnMouseClick);
            ev.setMouseMove(MEv_OnMouseMove);
        }

        private void position(InteractionGraphics ig, View View, Point2d sheet)
        {
            //var sheet = View.Camera.ViewToModelSpace(ViewPosition);
            Point2d st = dcs[0].CenterPoint;
            //var pt = I.CP(cen[0].Position.X, sheet.Y);
            var vec = st.VectorTo(dcs[dcs.Count - 1].CenterPoint);
            double w = vec.Length;
            vec.Normalize();
            Vector2d norm = null;
            u.normal(vec, st.VectorTo(I.CP2d(sheet.X, sheet.Y)), out norm);
            var pt2 = st.Copy(); pt2.TranslateBy(norm);
            //var pt3 = cen[1].Position.Copy();
            //var pt4 = pt3.Copy(); pt4.TranslateBy(norm);
            var color = u.addColor(ig.GraphicsDataSets, 15, 255, 0, 0);
            addLines(ig, norm, color);
            norm.Normalize(); norm.ScaleBy(text_height);
            if (sheet.X > st.X || sheet.Y < st.Y) norm.ScaleBy(-1);
            vec.ScaleBy(w);
            var pts = u.getRect(pt2, vec, norm);
            this.cenPt = pt2;
            norm.AddVector(vec); norm.ScaleBy(0.5);
            this.cenPt.TranslateBy(norm);
            //Point2d[] pts2 = new Point2d[] { pt3, pt4 };
            var lg = u.clientLine(ig.OverlayClientGraphics,
                ig.GraphicsDataSets, name, colName, pts, 10, false, true);
            lg.ColorSet = color;
            //var lg2 = u.clientLine(ig.OverlayClientGraphics,
            //    ig.GraphicsDataSets, name, colName, u.getPoints(pts2), 2, false, true);
            //lg2.ColorSet = color;
            //u.addClientLine(lg, I.CP(cen[0].Position.X, cen[0].Position.Y), true);
            u.viewUpdate(true, View, ig);
        }

        private void MEv_OnMouseMove(MouseButtonEnum Button, ShiftStateEnum ShiftKeys,
            Point ModelPosition, Point2d ViewPosition, View View)
        {
            if (!this.mouse) return;
            var sheet = I.CP2d(ModelPosition.X, ModelPosition.Y);
            position(cur.evts.InteractionGraphics, View, sheet);
        }

        void addLines(InteractionGraphics ig, Vector2d n, GraphicsColorSet c)
        {
            int number = 1;
            foreach (DrawingCurve item in dcs)
            {
                var lg = addLine(item.CenterPoint, n, ig, number);
                lg.ColorSet = c;
                number++;
            }
        }

        private LineGraphics addLine(Point2d pt, Vector2d n, InteractionGraphics ig, int num)
        {
            var p2 = pt.Copy(); p2.TranslateBy(n);
            Point2d[] pts = new Point2d[] { pt, p2 };
            var lg = u.clientLine(ig.OverlayClientGraphics,
                ig.GraphicsDataSets, name, colName, u.getPoints(pts), num, false, true);
            return lg;
        }

        private void KeyboardEvents_OnKeyDown1(int Key, ShiftStateEnum ShiftKeys)
        {
            if (Key == 77)
            {
                this.mouse = !this.mouse;
            }
            if (Key == 78)
            {
                offset += text_height / 2;
                draw(cur.evts.InteractionGraphics);
            }
            if (Key == 66)
            {
                offset -= text_height / 2;
                draw(cur.evts.InteractionGraphics);
            }
            if (Key == 13)
            {
                addDims();
                cur.EvInt = true;
            }
        }

        void draw(InteractionGraphics ig)
        {
            if (dc == null) return;
            var pt = dc.CenterPoint;
            this.dir = u.findMinDist(dv, pt, pt.VectorTo(dcs[1].CenterPoint));
            this.dir = add(this.dir, text_height + offset);
            pt.TranslateBy(this.dir);
            //pt = u.getPoint(dv, false);
            position(ig, I.app.ActiveView, pt);
        }

        private void addDims()
        {
            LinearGeneralDimension d = null;
            Point2d origin = dc.CenterPoint;
            dcs.Sort((a, b) => origin.DistanceTo(a.CenterPoint).CompareTo(origin.DistanceTo(b.CenterPoint)));
            DimArray.sh = sh; DimArray.dv = dv;
            DimsArr da = new DimsArr(this.dcs, this.dir);
            da.fillDims(sh);
            da.check();
            da.draw();
        }

        private Vector2d add(Vector2d v, double d)
        {
            var dir = v.Copy();
            dir.Normalize();
            var dist = v.Length + d;
            dir.ScaleBy(dist);
            return dir;
        }

        private void KeyboardEvents_OnKeyDown(int Key, ShiftStateEnum ShiftKeys)
        {
            //if (Key == 76)
            //{
            //    if (cen.Count == 1) addLinked(cen[0]); 
            //}

            if (Key == 65)
            {
                if (r == 0) setRadius();
                else r = 0;
            }

            if (Key == 78)
            {
                select();
            }
        }
        public void setRadius()
        {
            r = u.getRadius(dc);
        }
        public bool select()
        {
            this.dcs = new List<DrawingCurve>() { dc };
            var v = dirs2d[ind];
            ind++;
            if (dirs2d.Count == ind) ind = 0;
            dv = dc.Parent;
            var dcs = u.findAtVector(dv, dc.CenterPoint, v, r);
            this.dcs.AddRange(dcs);
            if (this.dcs.Count() == 1) return false;
            if (hs != null) hs.Clear();
            foreach (DrawingCurve item in dcs)
            {
                highlight(item.Segments[1]);
            }
            return true;
        }
        public void addLinked(Centermark cm)
        {
            foreach (var item in cm.Centerlines)
            {
                highlight(item);
            }
        }
        public void highlight(object o)
        {
            if (hs == null)
            {
                hs = doc.CreateHighlightSet();
                hs.Color = I.app.TransientObjects.CreateColor(255, 0, 0, 0.5);
            }
            hs.AddItem(o);
        }
        public void clear()
        {
            evts.Clear();
            if (hs != null) hs.Clear();
        }
        private void Sel_OnSelect(ObjectsEnumerator JustSelectedEntities, SelectionDeviceEnum SelectionDevice,
            Point ModelPosition, Point2d ViewPosition, View View)
        {
            foreach (var o in JustSelectedEntities)
            {
                if (o is Centermark)
                {
                    cen.Add(o as Centermark);
                    //if (cen.Count > 1) { evts[num].EvInt = true; }
                }
                else if (o is DrawingCurveSegment)
                {
                    dc = ((DrawingCurveSegment)o).Parent;
                    setRadius();
                    if (!select())
                        select();
                    //evts[num].EvInt = true;
                }
            }
        }
        private void MEv_OnMouseClick(MouseButtonEnum Button, ShiftStateEnum ShiftKeys,
            Point ModelPosition, Point2d ViewPosition, View View)
        {
            addDims();
            cur.EvInt = true;
        }

        private void Sel_OnPreSelectCM(ref object PreSelectEntity, out bool DoHighlight,
            ref ObjectCollection MorePreSelectEntities, SelectionDeviceEnum SelectionDevice,
            Point ModelPosition, Point2d ViewPosition, View View)
        {
            DoHighlight = false;
            if (PreSelectEntity is Centermark) DoHighlight = true;
            if (PreSelectEntity is DrawingCurveSegment)
            {
                DrawingCurveSegment dcs = PreSelectEntity as DrawingCurveSegment;
                if (dcs.Parent.CurveType == CurveTypeEnum.kCircleCurve ||
                    dcs.Parent.CurveType == CurveTypeEnum.kCircularArcCurve) DoHighlight = true;
            }
        }
        private void Sel_OnPreSelectDC(ref object PreSelectEntity, out bool DoHighlight,
            ref ObjectCollection MorePreSelectEntities, SelectionDeviceEnum SelectionDevice,
            Point ModelPosition, Point2d ViewPosition, View View)
        {
            DoHighlight = true;
            if (!(PreSelectEntity is DrawingCurveSegment)) DoHighlight = false;
        }
    }

    public class DimsArr
    {
        List<DimArray> arr = new List<DimArray>();
        List<LinearGeneralDimension> dims = new List<LinearGeneralDimension>();
        Vector2d n;
        public DimsArr(List<DrawingCurve> dcs, Vector2d v)
        {
            n = v;
            for (int i = 0; i < dcs.Count - 1; i++)
            {
                DimArray da = new DimArray(dcs[i], dcs[i + 1]);
                arr.Add(da);
            }
        }
        public void fillDims(Sheet sh)
        {
            foreach (var item in sh.DrawingDimensions.GeneralDimensions)
            {
                var d = item as LinearGeneralDimension;
                if (d != null) dims.Add(d);
            }
        }
        public void check()
        {
            DimArray a1 = null, a2 = null;
            for (int i = 0; i < arr.Count - 1; i++)
            {
                a1 = arr[i];
                a2 = arr[i + 1];
                if (a1.check(a2))
                {
                    a2.setVisible(false);
                }
                //while (a1.check(a2) && i < arr.Count - 1)
                //{
                //    a2 = arr[i+1];
                //    a1.count++;
                //    a1.dc2 = a2.dc2;
                //    a2.setVisible(false);
                //    i++;
                //}
                //a1 = a2; a2 = null;
            }
            set();
        }
        public bool check(DimArray d)
        {
            foreach (var item in dims)
            {
                if (item.IntentOne == null || item.IntentTwo == null) continue;
                if (check(item.IntentOne, d.dc1, d.dc2) &&
                    check(item.IntentTwo, d.dc1, d.dc2)) return true;
            }
            return false;
        }
        public bool check(GeometryIntent gi, DrawingCurve dc1, DrawingCurve dc2)
        {
            if (gi.IntentType == IntentTypeEnum.kPointEnumIntent)
            {
                return gi.Geometry.Equals(dc1) || gi.Geometry.Equals(dc2);
            }
            else if (gi.IntentType == IntentTypeEnum.kPoint2dIntent)
            {
                dynamic d = gi.Geometry;
                var g = d.AttachedEntity as FlatPunchResult;
                if (g != null) return false;
                return d.AttachedEntity.Geometry.Equals(dc1) ||
                    d.AttachedEntity.Geometry.Equals(dc2);
            }
            return false;
        }
        public void draw()
        {
            foreach (var item in arr)
            {
                if (check(item)) continue;
                item.draw(n);
            }
        }
        public void set()
        {
            DimArray prev = arr[0];
            int prevInd = 0;
            for (int i = 1; i < arr.Count; i++)
            {
                var cur = arr[i];
                if (!cur.vis)
                {
                    prev.dc2 = arr[i].dc2;

                }
                else
                {
                    prev.count = i - prevInd;
                    prevInd = i;
                    prev = cur;
                }
                if (i == arr.Count - 1)
                {
                    prev.count = i - prevInd + 1;
                }
            }
        }
    }

    public class DimArray
    {
        public DrawingCurve dc1, dc2;
        static public Sheet sh = null;
        static public DrawingView dv = null;
        GeometryIntent i1, i2;
        Point2d p1, p2;
        Point mp1, mp2;
        public bool vis = true;
        string s = "";
        Vector2d n;
        public double d = 0;
        public int count = 1;
        public DimArray(DrawingCurve d1, DrawingCurve d2)
        {
            dc1 = d1; dc2 = d2;
            fill();
        }
        public void fill()
        {
            i1 = sh.CreateGeometryIntent(dc1);
            i2 = sh.CreateGeometryIntent(dc2);
            mp1 = getPoint(dc1);
            mp2 = getPoint(dc2);
            d = u.round(mp1.DistanceTo(mp2));
        }
        public void setVisible(bool v)
        {
            vis = v;
        }
        public bool check(DimArray a)
        {
            return u.eq(a.d, d);
        }
        private Point2d getPoint(GeometryIntent i)
        {
            dynamic d = i.Geometry;
            return d.CenterPoint;
        }
        private Point getPoint(DrawingCurve dc)
        {
            dynamic m = dc.ModelGeometry;
            return m.Geometry.Center;
        }
        public LinearGeneralDimension draw(Vector2d v)
        {
            if (!vis) return null;
            if (count != 1)
            {
                s = $"{this.count}x{(this.d * 10).ToString("0.#")}=";
            }
            fill();
            n = v;
            p1 = getPoint(i1); p2 = getPoint(i2);
            var mp = u.midPt(p1, p2);
            mp.TranslateBy(n);
            var d = sh.DrawingDimensions.GeneralDimensions.AddLinear(mp, i1, i2);
            if (s != "") d.Text.FormattedText = s + d.Text.FormattedText;
            return d;
        }
    }

    public class SlotDraw
    {
        DrawingDocument doc;
        public Sheet sh;
        SelectSet ss;
        public Point2d pt1, pt2;
        public Point p1, p2;
        static public HighlightSet hs;
        object selEnt = null;
        public double w = 2.5, h = 0.7;
        int num = 0;
        EventsDims cur = null;
        List<EventsDims> evts = new List<EventsDims>();
        List<DrawingDimension> dims = new List<DrawingDimension>();
        public event EventHandler<MyEventArgs> change;
        string name = "dim", colName = "dimCol";
        public DrawingCurveSegment dcs;
        DrawingView dv;

        public SlotDraw(Document doc)
        {
            if (doc.DocumentType != DocumentTypeEnum.kDrawingDocumentObject) return;
            this.doc = doc as DrawingDocument;
            sh = this.doc.ActiveSheet;
            fillDims();
        }
        public bool run()
        {
            for (int i = 0; i < evts.Count; i++)
            {
                cur = evts[i];
                if (cur.mm != null)
                {
                    cur.mEv.MouseMoveEnabled = true;
                }
                cur.run(el => set(el));
                cur.termEvents();
                if (cur.term) return false;
                num++;
            }
            clear();
            return true;
        }

        public bool set(object o)
        {
            if (o is DrawingCurveSegment) dcs = o as DrawingCurveSegment;
            if (dcs != null && dcs.GeometryType == Curve2dTypeEnum.kCircularArcCurve2d) w = 1.5;
            return true;
        }

        public void fillDims()
        {
            foreach (DrawingDimension dim in sh.DrawingDimensions)
            {
                var t = dim.Text.FormattedText.ToLower();
                if (t.IndexOf("отв") != -1 || t.IndexOf("шт") != -1) dims.Add(dim);
            }
        }

        public bool checkDim(double r, Curve2dTypeEnum t)
        {
            foreach (var item in dims)
            {
                if (t == Curve2dTypeEnum.kCircularArcCurve2d && u.eq(item.ModelValue, r)) return true;
                if (t == Curve2dTypeEnum.kCircleCurve2d && u.eq(item.ModelValue, r * 2)) return true;
            }
            return false;
        }

        public EventsDims add(string txt, SelectionFilterEnum[] f)
        {
            var ev = new EventsDims(txt, f);
            evts.Add(ev);
            return ev;
        }
        public void onChange(bool r)
        {
            MyEventArgs a = new MyEventArgs(r);
            change(this, a);
        }

        public void fill()
        {
            u.clearClientLine(doc, name, colName);
            var ev = add("Выберите отверстие или овал", new SelectionFilterEnum[] {
            SelectionFilterEnum.kDrawingCurveSegmentFilter});
            ev.setPreSelectEv(Sel_OnPreSelectCM);
            ev.setSelectEv(Sel_OnSelect);
            ev.setKeyEv(KeyboardEvents_OnKeyDown);
            ev = add("Укажите положение размера", new SelectionFilterEnum[] {
            SelectionFilterEnum.kDrawingDefaultFilter});
            ev.setMouseClick(MEv_OnMouseClick);
            ev.setMouseMove(MEv_OnMouseMove);
        }

        private void MEv_OnMouseMove(MouseButtonEnum Button, ShiftStateEnum ShiftKeys,
            Point ModelPosition, Point2d ViewPosition, View View)
        {
            var sheet = View.Camera.ViewToModelSpace(ViewPosition);
            var pt = I.CP(sheet.X, sheet.Y);
            //double w = this.pt.DistanceTo(this.pt);
            Point[] pts, pts1;
            if (pt1 == null) return;
            var tmp = View.Camera.ViewToModelSpace(pt1);
            if (tmp.X <= pt.X)
            {
                pts = u.getRect(pt, w, h);
                if (w == 1.5)
                {
                    pts1 = u.getRect(pt, w, -h);
                    pts = pts.Concat(pts1).ToArray();
                }
            }
            else
            {
                pts = u.getRect(pt, -w, h);
                if (w == 1.5)
                {
                    pts1 = u.getRect(pt, -w, -h);
                    pts = pts.Concat(pts1).ToArray();
                }
            }

            var ig = cur.evts.InteractionGraphics;
            var lg = u.clientLine(ig.OverlayClientGraphics, ig.GraphicsDataSets, name, colName, pts, 1, false);
            u.addClientLine(lg, View.Camera.ViewToModelSpace(this.pt1), true);
            u.viewUpdate(true, View, ig);
        }

        public void clear()
        {
            u.clearClientLine(doc, name, colName);
        }

        private void KeyboardEvents_OnKeyDown(int Key, ShiftStateEnum ShiftKeys)
        {
        }
        public void addLinked(Centermark cm)
        {
            foreach (var item in cm.Centerlines)
            {
                highlight(item);
            }
        }
        public void highlight(object o)
        {
            if (hs == null)
            {
                hs = doc.CreateHighlightSet();
                hs.Color = I.app.TransientObjects.CreateColor(255, 0, 0, 0.5);
            }
            hs.AddItem(o);
        }
        public void clearClient()
        {
            evts.Clear();
            if (hs != null) hs.Clear();
        }
        private void Sel_OnSelect(ObjectsEnumerator JustSelectedEntities, SelectionDeviceEnum SelectionDevice,
            Point ModelPosition, Point2d ViewPosition, View View)
        {
            p1 = ModelPosition;
            pt1 = View.Camera.ModelToViewSpace(p1);
        }
        private void MEv_OnMouseClick(MouseButtonEnum Button, ShiftStateEnum ShiftKeys,
            Point ModelPosition, Point2d ViewPosition, View View)
        {
            p2 = ModelPosition;
            evts[num].EvInt = true;
        }

        private void Sel_OnPreSelectCM(ref object PreSelectEntity, out bool DoHighlight,
            ref ObjectCollection MorePreSelectEntities, SelectionDeviceEnum SelectionDevice,
            Point ModelPosition, Point2d ViewPosition, View View)
        {
            DoHighlight = true;
            DrawingCurveSegment dcs = PreSelectEntity as DrawingCurveSegment;

            if (dcs.GeometryType != Curve2dTypeEnum.kCircularArcCurve2d &&
                dcs.GeometryType != Curve2dTypeEnum.kCircleCurve2d) DoHighlight = false;
            else
            {
                dynamic g = dcs.Geometry;
                double r = g.Radius;
                if (checkDim(r, dcs.GeometryType)) DoHighlight = false;
            }
        }
        private void Sel_OnPreSelectDC(ref object PreSelectEntity, out bool DoHighlight,
            ref ObjectCollection MorePreSelectEntities, SelectionDeviceEnum SelectionDevice,
            Point ModelPosition, Point2d ViewPosition, View View)
        {
            DoHighlight = true;
            if (!(PreSelectEntity is DrawingCurveSegment)) DoHighlight = false;
        }
    }

    public class composeIM
    {
        public List<Im> ims = new List<Im>();
        AssemblyDocument aDoc;
        public string name;
        public object ml;

        public composeIM(object m, bool rev)
        {
            CompositeiMateDefinition cm = m as CompositeiMateDefinition;
            if (cm != null)
            {
                foreach (var item in cm)
                {
                    var e = new Im(item);
                    ims.Add(e);
                }
                name = cm.Name;
                ml = cm.MatchList;
                if (rev)
                {
                    string[] lst = ml as string[];
                    if (lst.Length == 0) return;
                    name = lst[0];
                    ml = new string[] { cm.Name };
                }
            }
            iMateDefinition im = m as iMateDefinition;
            if (im != null)
            {
                name = im.Name;
                ml = im.MatchList;
                ims.Add(new Im(im));
            }
        }
    }
    public enum imType { Insert, Mate, Flush }
    public class Im
    {
        public object ent;
        public object dist;
        public bool ao;
        public string name;
        public object ml;
        public imType t;
        public Im(object m)
        {
            var im = m as InsertiMateDefinition;
            if (im != null)
            {
                ent = im.Entity; dist = im.Distance.Value;
                ao = im.AxesOpposed;
                t = imType.Insert;
                name = im.Name;
                if (im.MatchList != null)
                    ml = im.MatchList;
                return;
            }
            var mm = m as MateiMateDefinition;
            if (mm != null)
            {
                ent = mm.Entity; dist = mm.Offset.Value;
                t = imType.Mate; name = mm.Name;
                if (mm.MatchList != null) ml = mm.MatchList;
                return;
            }
            var fm = m as MateiMateDefinition;
            if (fm != null)
            {
                ent = fm.Entity; dist = fm.Offset.Value;
                t = imType.Flush; name = fm.Name;
                if (fm.MatchList != null) ml = fm.MatchList;
                return;
            }
        }
    }

    public class SlotModel
    {
        Document doc;
        SheetMetalComponentDefinition smcd;
        FlatPattern fp;
        Face f;
        List<EdgeLoop> loop = new List<EdgeLoop>();
        double[] bh = new double[2];
        IEnumerable<Edge> arcs;
        public int count { get { return loop.Count(); } }
        public SlotModel(Document doc)
        {
            this.doc = doc;
            smcd = I.getSMCD(doc);
            fp = I.getFP(doc);
        }
        public void fill(Edge e)
        {
            if (e == null) return;
            bh = set(e.EdgeUses[1].EdgeLoop);
            f = u.get<Face>(e.Faces, el => el.SurfaceType == SurfaceTypeEnum.kPlaneSurface);
            foreach (EdgeLoop el in f.EdgeLoops)
            {
                if (el.Edges.Count != 4) continue;
                if (check(el)) loop.Add(el);
            }
        }
        public bool check(EdgeLoop el)
        {
            var bh1 = set(el);
            if (bh1 == null) return false;
            if (bh1.Length != bh.Length) return false;
            for (int i = 0; i < bh1.Length; i++)
            {
                if (!u.eq(bh1[i], bh[i])) return false;
            }
            return true;
        }
        public double[] set(EdgeLoop el)
        {
            arcs = u.gets<Edge>(el.Edges, fi => fi.GeometryType == CurveTypeEnum.kCircularArcCurve);
            if (arcs.Count() != 2) return null;
            Arc3d a1 = arcs.ElementAt(0).Geometry as Arc3d, a2 = arcs.ElementAt(1).Geometry as Arc3d;
            return new double[] { a1.Radius * 2, a1.Center.VectorTo(a2.Center).Length + a1.Radius * 2 };
        }
        public string txt()
        {
            return $"{bh[0] * 10:0.#}x{bh[1] * 10:0.#};R={bh[0] * 5:0.#}\n{count} отв.";
        }
    }
    public class SketchModel
    {
        Document doc;
        SheetMetalComponentDefinition smcd;
        PlanarSketch ps;
        FlatPattern fp;
        Face f;
        double r;
        List<SketchEntity> loop = new List<SketchEntity>();
        public int count { get { return loop.Count(); } }
        public SketchModel(Document doc)
        {
            this.doc = doc;
            smcd = I.getSMCD(doc);
            fp = I.getFP(doc);
        }
        public void fill(SketchEntity e)
        {
            if (e == null) return;
            ps = e.Parent as PlanarSketch;
            SketchCircle sc = e as SketchCircle;
            if (sc != null) fill(sc);
            SketchArc sa = e as SketchArc;
            if (sa != null) fill(sa);
        }
        public void fill(SketchCircle c)
        {
            r = c.Radius;
            foreach (SketchCircle item in ps.SketchCircles)
            {
                if (u.eq(item.Radius, r)) loop.Add(c as SketchEntity);
            }
        }
        public void fill(SketchArc c)
        {
            r = c.Radius;
            foreach (SketchArc item in ps.SketchArcs)
            {
                if (u.eq(item.Radius, r)) loop.Add(c as SketchEntity);
            }
        }
        public string txt()
        {
            return $"{count} отв. ";
        }
    }

    public class HoleModel
    {
        Document doc;
        SheetMetalComponentDefinition smcd;
        FlatPattern fp;
        Vector vec;
        Face f;
        Edge first;
        Point origin;
        List<EdgeLoop> loop = new List<EdgeLoop>();
        double[] bh = new double[2];
        IEnumerable<Edge> arcs;
        Matrix mtx;
        public int count { get { return loop.Count(); } }
        public int edCount { get { return arcs.Count(); } }
        public HoleModel(Document doc)
        {
            this.doc = doc;
            smcd = I.getSMCD(doc);
            fp = I.getFP(doc);
        }
        public void setMtx(Edge e, Vector dir)
        {
            vec = dir;
            first = e;
            origin = u.getCenter(e);
            mtx = I.getMatrix();
            f = u.get<Face>(e.Faces, el => el.SurfaceType == SurfaceTypeEnum.kPlaneSurface);
            if (!(f.Geometry is Plane pl)) return;
            Vector v2 = pl.Normal.AsVector();
            Vector v3 = dir.CrossProduct(v2);
            mtx.SetToAlignCoordinateSystems(origin, dir, v2, v3, I.CP(), I.CV(1, 0, 0), I.CV(0, 1, 0), I.CV(0, 0, 1));
            transform(origin);
        }
        public void transform(Point pt)
        {
            pt.TransformBy(mtx);
        }
        public void find()
        {
            arcs = u.gets<Edge>(f.Edges, fi => check(fi));
        }
        public bool check(Edge e)
        {
            if (e.Equals(first)) return false;
            if (!(e.GeometryType == CurveTypeEnum.kCircleCurve ||
                e.GeometryType == CurveTypeEnum.kCircularArcCurve)) return false;
            Point pt = u.getCenter(e);
            transform(pt);
            double[] c1 = { }, c2 = { };
            origin.GetPointData(ref c1);
            pt.GetPointData(ref c2);
            var v = origin.VectorTo(pt);
            int count = 0;
            //var dp = v.CrossProduct(vec);
            //if (u.eq(dp.Length, 0)) return false;
            //if (!v.IsParallelTo(vec)) return false;
            for (int i = 1; i < c1.Length; i++)
            {
                if (u.eq(c1[i], c2[i])) count++;
            }
            return count == 2;
            //return u.eq(pt.Y, origin.Y) && u.eq(pt.Z, origin.Z);
        }
        public void fill(Edge e)
        {
            if (e == null) return;
            bh = set(e.EdgeUses[1].EdgeLoop);
            f = u.get<Face>(e.Faces, el => el.SurfaceType == SurfaceTypeEnum.kPlaneSurface);
            foreach (EdgeLoop el in f.EdgeLoops)
            {
                if (el.Edges.Count != 1) continue;
                if (check(el)) loop.Add(el);
            }
        }
        public bool check(EdgeLoop el)
        {
            var bh1 = set(el);
            if (bh1 == null) return false;
            if (bh1.Length != bh.Length) return false;
            for (int i = 0; i < bh1.Length; i++)
            {
                if (!u.eq(bh1[i], bh[i])) return false;
            }
            return true;
        }
        public double[] set(EdgeLoop el)
        {
            arcs = u.gets<Edge>(el.Edges, fi => fi.GeometryType == CurveTypeEnum.kCircleCurve);
            if (arcs.Count() != 1) return null;
            Circle a1 = arcs.ElementAt(0).Geometry as Circle;
            return new double[] { a1.Radius * 2 };
        }
        public string txt()
        {
            return $"{count} отв. ";
        }
    }

    internal class InteractiveBtn : Button
    {
        //public static Drawings m_Drw;
        //public static Drawings getDrw { get { return m_Drw; } }
        public InteractiveBtn(string displayName, string internalName, string clientId, string description, string tooltip,
            System.Drawing.Icon standardIcon, System.Drawing.Icon largeIcon)
            : base(displayName, internalName, clientId, description, tooltip, standardIcon, largeIcon)
        {
        }
        protected override void ButtonDefinition_OnExecute(NameValueMap context)
        {
            var i = new Interactive(I.aDoc());
            System.Windows.Forms.Application.Run(i);
        }
    }
}
