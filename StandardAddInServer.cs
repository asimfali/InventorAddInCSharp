using System;
using System.Runtime.InteropServices;
using Inventor;
using System.Windows.Forms;
using Microsoft.Win32;
using System.Linq;
using System.Globalization;
using System.Drawing;
using InterfaceDll;

namespace Macros
{
    /// <summary>
    /// This is the primary AddIn Server class that implements the ApplicationAddInServer interface
    /// that all Inventor AddIns are required to implement. The communication between Inventor and
    /// the AddIn is via the methods on this interface.
    /// </summary>
    [GuidAttribute("c09e4c46-e69b-478a-a2c1-ad4b52e62879")]
    public class StandardAddInServer : Inventor.ApplicationAddInServer
    {
                                                      
        #region Data Members
        public static Inventor.Application m_inventorApplication { get; set; }
        //public static TransientGeometry m_tg { get; set; }
        //public static TransientObjects m_objs { get; set; }
        public static System.Collections.Generic.HashSet<System.Windows.Forms.Form> forms = new System.Collections.Generic.HashSet<System.Windows.Forms.Form>();
        public static InvAddIn.DataToArray data;
        public InvAddIn.Material material;
        //public InvAddIn.InventorEvents events = new InvAddIn.InventorEvents();
        public Inventor.Document m_Doc;
        public Inventor.ContentCenter cc;
        private InvAddIn. Spec m_SpecButton;
        private InvAddIn.MyBreakOperation m_Break;
        private InvAddIn.MyBreakOperationVert m_BreakVert;
        private InvAddIn.BalloonOp m_Balloon;
        private InvAddIn.CutButton m_Cut;
        //private InvAddIn.SetButton m_set;
        private InvAddIn.ContentBtn m_cc;
        private InvAddIn.PDFButton m_PDF;
        private InvAddIn.TablesButton m_Tbl;
        private InvAddIn.OffsetBtn m_Offset;
        private InvAddIn.StickerBtn m_Sticker;
        private InvAddIn.PassportBtn m_Passport;
        private InvAddIn.IMateBtn m_IMate;
        private InvAddIn.Print m_Print;
        private InvAddIn.FP m_FP;
        private InvAddIn.PropBtn m_PropBtn;
        private InvAddIn.RecoverBtn m_RecBtn;
        private InvAddIn.PartsBtn m_parts;
        private InvAddIn.DrawingBtn m_DrwBtn;
        private InvAddIn.SurfaceBtn m_surfBtn;
        private InvAddIn.ListBtn m_ListBtn;
        private InvAddIn.VarBtn m_VarBtn;
        private InvAddIn.GabButton m_Gab;
        private InvAddIn.SketchInModel m_sketchInModel;
        private InvAddIn.DrwArrBtn m_DrwArrBtn;
        private InvAddIn.SketchBtn m_Sketch;
        private InvAddIn.BreakButton m_xmlBreak;
        private InvAddIn.SplineButton m_spline;
        private InvAddIn.ComboBoxBtn m_ComboBox;
        private InvAddIn.Assembly m_Asm;
        //private InvAddIn.Sketches m_Sketches;
        RibbonPanel m_partSketchRibbonPanel;
        //private DocumentEvents m_DocEvts;
        private Inventor.ApplicationEvents m_appEvts;
        private Inventor.SketchEvents m_SketchEvts;
        private Inventor.UserInputEvents m_InputEvts;
        private Inventor.AssemblyEvents m_asmbllEvts;
        //private Inventor.UserInterfaceEvents m_UIEvnts;
        //private Inventor.DocumentEventsSink_OnSaveEventHandler DocEvts_OnSave;
        //private Inventor.DocumentEventsSink_OnChangeEventHandler DocEvts_OnChange;
        public string prop;
        public static string drwTemplate;
        public double val;
        public static string spec { get; set; }
        public static string BreakHor { get; set; }
        public static string BreakVert { get; set; }
        public static string Alp { get; set; }
        public static string Balloon { get; set; }
        public static string Mat { get; set; }
        public static string OriginPoint { get; set; }
        public static string Arrow { get; set; }
        public static string DXFFromPDF { get; set; }
        public static string StandartPart { get; set; }
        public static string Printflag { get; set; }
        public static string Spike { get; set; }
        public static string dim { get; set; }
        private bool deleteView = false;
        private System.Collections.Generic.List<Inventor.Balloon> balloons;
        static public System.Collections.Generic.List<string> lst;
        static public System.Collections.Generic.List<string> oldval;
        static public System.Collections.Generic.List<Inventor.ComponentOccurrence> lstco;
        static public System.Collections.Generic.List<string> lstSort;
        string[] attr;
        public static MyXML xml;
        //private Inventor.NameValueMap NVM;

        #endregion


        public StandardAddInServer()
        {
        }

        #region ApplicationAddInServer Members

        public void Activate(Inventor.ApplicationAddInSite addInSiteObject, bool firstTime)
        {
            // This method is called by Inventor when it loads the addin.
            // The AddInSiteObject provides access to the Inventor Application object.
            // The FirstTime flag indicates if the addin is loaded for the first time.

            // Initialize AddIn members.
//              try
//              {
            m_inventorApplication = addInSiteObject.Application;
            //m_tg = m_inventorApplication.TransientGeometry;
            //m_objs = m_inventorApplication.TransientObjects;
            InvAddIn.Button.InventorApplication = m_inventorApplication;
            InvAddIn.InvComboBox.InventorApplication = m_inventorApplication;
            cc = m_inventorApplication.ContentCenter;
            settings();
            Type m_type;
            m_type = this.GetType();
            Icon specIcon = InvAddIn.Properties.Resources.Spec;
            Icon breakIcon = InvAddIn.Properties.Resources.BreakHor;
            Icon breakVertIcon = InvAddIn.Properties.Resources.BreakVert;
            Icon BalloonIcon = InvAddIn.Properties.Resources.Balloon;
            Icon CutIcon = InvAddIn.Properties.Resources.Cut;
            Icon SetIcon = InvAddIn.Properties.Resources.Settings;
            Icon PDFIcon = InvAddIn.Properties.Resources.PDF;
            Icon TblIcon = InvAddIn.Properties.Resources.Table;
            Icon OffsetIcon = InvAddIn.Properties.Resources.Offset;
            Icon StickerIcon = InvAddIn.Properties.Resources.Sticker;
            Icon PassportIcon = InvAddIn.Properties.Resources.Passport;
            Icon printIcon = InvAddIn.Properties.Resources.Print;
            Icon IMateIcon = InvAddIn.Properties.Resources.IMate;
            Icon FastenersIcon = InvAddIn.Properties.Resources.Fasteners;
            Icon FPIcon = InvAddIn.Properties.Resources.FP;
            Icon PropIcon = InvAddIn.Properties.Resources.Prop;
            Icon RecoverIcon = InvAddIn.Properties.Resources.Recover;
            Icon PartsIcon = InvAddIn.Properties.Resources.Parts;
            Icon DrwIcon = InvAddIn.Properties.Resources.Drw;
            Icon SketchIcon = InvAddIn.Properties.Resources.Sketch;
            Icon ComboBoxIcon = InvAddIn.Properties.Resources.ComboBox;
            Icon UpdateIcon = InvAddIn.Properties.Resources.update;
            Icon SetIconLarge = new Icon(SetIcon, 32, 32);
            Icon TblIconLarge = new Icon(TblIcon, 32, 32);
            Icon standartbreakIcon = new Icon(breakIcon, 16, 16);
            Icon largebreakIcon = new Icon(breakIcon, 32, 32);

            m_SpecButton = new InvAddIn.Spec("Спецификация", "Spec","{ED6E59F8-E475-4C14-A5BA-1DDBD041CD82}", "Спецификация",
                "Спецификация с переменными данными", specIcon, specIcon);
            m_SpecButton.addShortCut("Alt+S");

            m_Break = new InvAddIn.MyBreakOperation("ГР", "BreakHor", "{77F3E204-1071-4B17-914F-799E51B292FD}", "Добавить горизонтальный разрыв", 
                "Горизонтальный разрыв с определенными параметрами", standartbreakIcon, largebreakIcon);

            m_BreakVert = new InvAddIn.MyBreakOperationVert("ВР", "BreakVert" ,"{57B15465-B68E-42F0-8CEC-B3FBC5B509AB}", "Добавить вертикальный разрыв",
                "Вертикальный разрыв с определенными параметрами", breakVertIcon, breakVertIcon);

            m_Balloon = new InvAddIn.BalloonOp("Позиция", "Balloon", "{E641CEF1-A4CB-4C69-9AA8-DCFA6CBB9D4F}", "Вынести номер позиции", 
                "Групповой номер позиции", BalloonIcon, BalloonIcon);

            m_Cut = new InvAddIn.CutButton("Вырез", "Cut", "{C88EAB8B-7C31-4BF3-A5AA-B6B2073DABF1}", "Вырез",
                "Вырез", CutIcon, CutIcon);

            //m_set = new InvAddIn.SetButton("Настройки", "Autodesk:Macros:Set", CommandTypesEnum.kNonShapeEditCmdType,
            //    "{E7AB7D8E-A972-451B-B6CF-46927D1469B8}", "Используемые модули", "Используемые модули", SetIcon, SetIconLarge, ButtonDisplayEnum.kDisplayTextInLearningMode);

            m_cc = new InvAddIn.ContentBtn("Вставка крепежа", "CC", "{60DBC787-F747-43BD-9F6E-6B4F2B85FCF3}", "Вставка крепежа",
                "Вставка крепежа",FastenersIcon,FastenersIcon);
            InvAddIn.ContentBtn.getCC = cc;

            m_PDF = new InvAddIn.PDFButton("PDF\\DXF", "PDF", "{078FB534-4264-4159-9AA7-1D64A8629C7D}", "Создание PDF\\DXF",
                "Создание PDF\\DXF", PDFIcon, PDFIcon);

            m_Print = new InvAddIn.Print("Печать", "Print",
           "{29A0A185-CAE8-4EC7-94D6-2CE0A2DAB064}", "Печать", "Печать", printIcon, printIcon);

            m_FP = new InvAddIn.FP("Развернуть", "FP",
           "{F51893B9-5E6B-4484-8B0A-72E8B27CEAA4}", "Развернуть", "Развернуть", FPIcon, FPIcon);


            m_Tbl = new InvAddIn.TablesButton("Таблица", "Table",
            "{0483E272-C3EF-479E-A654-F5CF9E165642}", "Создание таблиц", "Создание таблиц", TblIcon, TblIconLarge);
            m_Tbl.addShortCut("Alt+G");


            m_Offset = new InvAddIn.OffsetBtn("Шип", "Offset",
            "{7FF6407E-6193-4D27-9C34-BCA7E2F85435}", "Вырез под шип", "Вырез под шип", OffsetIcon, OffsetIcon);

            m_Sticker = new InvAddIn.StickerBtn("Наклейки", "Sticker",
            "{8D0C4823-6192-4923-9BB5-61D8984AFBE0}", "Наклейки", "Наклейки", StickerIcon, StickerIcon);

            m_Passport = new InvAddIn.PassportBtn("Паспорт", "Passport",
             "{02A3630E-9118-4B5A-B1F2-FC8F3B64D10A}", "Паспорт", "Паспорт", PassportIcon, PassportIcon);

            m_IMate = new InvAddIn.IMateBtn("Конструктивная пара", "IMate",
            "{9845CF95-4D8A-4E24-B27B-79908DF23CF3}", "Конструктивная пара", "Конструктиваная пара", IMateIcon, IMateIcon);
            m_IMate.addShortCut("Alt+Q");


            m_PropBtn = new InvAddIn.PropBtn("Свойства проекта", "Prop",
            "{FF99F993-5248-4834-BD56-EC1AF0023A7C}", "Свойства проекта", "Свойства проекта", PropIcon, PropIcon);
            m_PropBtn.addShortCut("Alt+C");

            m_RecBtn = new InvAddIn.RecoverBtn("Восстановить", "Recover",
            "{168BFABD-D59E-410A-A30A-83EF48A469E4}", "восстановить зависимости", "Восстановить зависимости", RecoverIcon, RecoverIcon);

             m_parts = new InvAddIn.PartsBtn("Создать детали", "Parts",
            "{246BF5BA-872C-4792-9F6A-27FE9FC45C96}", "Создать детали", "Создать детали", PartsIcon, PartsIcon);
            m_parts.addShortCut("Alt+D");

            m_DrwBtn = new InvAddIn.DrawingBtn("Размеры", "Dimensions",
                "{EEF8F8DA-EBE1-417D-BCFC-B3344AF6698C}", "Размеры", "Размеры", DrwIcon, DrwIcon);

            m_ListBtn = new InvAddIn.ListBtn("Лист", "Autodesk:Macros:List", "{D18615A7-F27A-4A15-B18E-FACBAC3BE295}", "Работа с листом чертежа",
                "Работа с листом чертежа");

            m_surfBtn = new InvAddIn.SurfaceBtn("Шероховатость поверхности", "Autodesk:Macros:surf", "{E5E9F634-54BD-48B5-9BF1-D2114433ADA9}",
                "Добавление символов шероховатости необрабатвыемой поверхности", "Добавление символов шероховатости необрабатвыемой поверхности");

            m_VarBtn = new InvAddIn.VarBtn("Вставить параметрический элемент", "Autodesk:Macros:var", "{F6CCEBE3-0409-4083-9BF2-A33C107AA625}",
                "Добавление параметрического элемента по точкам и направлению в эскизе", "Добавление параметрического элемента по точкам и направлению в эскизе");

            m_Gab = new InvAddIn.GabButton("Размеры", "Autodesk:Macros:gab", "{2591E510-DDA0-49CD-B42E-78C55DEFCCBF}",
            "Добавление размеров на вид", "Добавление размеров на вид");

            m_Asm = new InvAddIn.Assembly("Разнести", "Autodesk:Macros:asm", "{B47B41A2-A984-4326-A275-321C7D663C72}", "Изменить расстояние", "Изменить расстояние");
            m_Asm.addShortCut("Alt+R");

            //m_Sketches = new InvAddIn.Sketches("Эскизы", "Autodesk:Macros:sketches", "{0B21CAD3-D1B4-4776-BF07-8BC9F60DAAF1}", "Копировать эскиз", "Копировать эскиз");
            m_DrwArrBtn = new InvAddIn.DrwArrBtn("Массив", "Autodesk:Macros:DrwArr", "{BC16CA31-2D3F-4F2C-93AF-ED4BE5CA0E5A}",
             "Добавление размера массива отверстий", "Добавление размера массива отверстий");

            m_xmlBreak = new InvAddIn.BreakButton("Разрыв", "Autodesk:Macros:xmlBreak", "{13CE307F-73E2-411D-9662-BAEFE0D8F961}",
                "Добавить разрывы", "Добавить разрывы");

            m_spline = new InvAddIn.SplineButton("Вырыв", "Autodesk:Macros:spline", "{394B0B7D-B034-4F76-B02D-C37F31C8028F}", "Добавить вырыв", "добавить вырыв");

            m_Sketch = new InvAddIn.SketchBtn(SketchIcon, SketchIcon);
            m_sketchInModel = new InvAddIn.SketchInModel("Отверстия", "Autodesk:Macros:sim", "{D4BDD51E-B174-487C-857E-FC592089CBCB}",
                "Создать из базового эскиза", "Создать из базового эскиза");
            m_ComboBox = new InvAddIn.ComboBoxBtn(ComboBoxIcon, ComboBoxIcon, "Данные для Массива");

            if (firstTime == true)
            {
                InvAddIn.Panel m_specPanel = new InvAddIn.Panel(m_inventorApplication, m_SpecButton._ButtonDefinition, 
                    "Drawing", "id_TabAnnotateESKD", "Спецификация", "Autodesk:Macros:SpecPanel", "{F7CAE311-990B-4C57-8183-8B59AED1040F}");
                if (BreakHor == "1") m_specPanel.addControl(m_Break._ButtonDefinition);
                if (BreakVert == "1") m_specPanel.addControl(m_BreakVert._ButtonDefinition);
                m_specPanel.addControl(m_PDF._ButtonDefinition);
                m_specPanel.addControl(m_Print._ButtonDefinition);
                m_specPanel.addControl(m_Passport._ButtonDefinition);
                m_specPanel.addControl(m_DrwBtn._ButtonDefinition);
                m_specPanel.addControl(m_RecBtn._ButtonDefinition);
                m_specPanel.addControl(m_parts._ButtonDefinition);
                //m_specPanel.addControl(m_set._ButtonDefinition);
                m_specPanel.CmdControl.UseLargeIcon = true;
                m_specPanel.addControl(m_Tbl._ButtonDefinition);
                m_specPanel.CmdControl.UseLargeIcon = true;
                InvAddIn.Panel m_PartPanel = new InvAddIn.Panel(m_inventorApplication, m_Cut._ButtonDefinition, 
                    "Part", "id_TabSheetMetal", "Элементы", "Autodesk:Macros:SMF", "{BDB0D42F-8E9E-4EF7-A84C-0655C17E0071}");
                 m_PartPanel.addControl(m_Offset._ButtonDefinition);
                m_PartPanel.addControl(m_PDF._ButtonDefinition);
                 m_PartPanel.addControl(m_FP._ButtonDefinition);
                m_PartPanel.addControl(m_IMate._ButtonDefinition);
                m_PartPanel.addControl(m_parts._ButtonDefinition);
                InvAddIn.Panel m_AssemblyPanel = new InvAddIn.Panel(m_inventorApplication, m_cc._ButtonDefinition,
                     "Assembly", "id_TabAssemble", "Крепеж", "Autodesk:Macros:CC", "{9664E6B8-74A6-432D-BAB1-1ADFFDB7CBB9}");
                 m_AssemblyPanel.addControl(m_Sticker._ButtonDefinition);
                m_AssemblyPanel.addControl(m_IMate._ButtonDefinition);
                m_AssemblyPanel.addControl(m_PropBtn._ButtonDefinition);
                m_AssemblyPanel.addControl(m_RecBtn._ButtonDefinition);
                m_AssemblyPanel.addControl(m_parts._ButtonDefinition);
                 m_AssemblyPanel.addControl(m_PDF._ButtonDefinition);
                InvAddIn.Panel m_SketchPanel = new InvAddIn.Panel("Part", "id_TabSketch", "id_PanelP_2DSketchFormat");
                m_SketchPanel.addControl(m_Sketch._ButtonDefinition);
                m_SketchPanel.addControl(m_ComboBox._ComboBoxDef,"SketchCenterlineCmd");
            }

            material = new InvAddIn.Material();                                                                                                                              
            m_appEvts = m_inventorApplication.ApplicationEvents;
            m_asmbllEvts = m_inventorApplication.AssemblyEvents;
            //m_modEvents = m_inventorApplication.ModelingEvents;
            m_SketchEvts = m_inventorApplication.SketchEvents;
            m_InputEvts = m_inventorApplication.CommandManager.UserInputEvents;
            m_appEvts.OnSaveDocument += new ApplicationEventsSink_OnSaveDocumentEventHandler(DocEvts_OnSaveHdlr);
            InvAddIn.InventorEvents events = new InvAddIn.InventorEvents();
            events.uInterfaceEv.OnEnvironmentChange += uInterfaceEv_OnEnvironmentChange;
            //m_appEvts.OnActivateDocument += m_appEvts_OnActivateDocument;
            m_appEvts.OnOpenDocument += new ApplicationEventsSink_OnOpenDocumentEventHandler(m_appEvts_OnOpenDocument);
            m_SketchEvts.OnNewSketch += new SketchEventsSink_OnNewSketchEventHandler(SketchEvts_OnNewSketch);
            m_InputEvts.OnTerminateCommand += new UserInputEventsSink_OnTerminateCommandEventHandler(UserInputEventsSink_OnTerminate);
            m_InputEvts.OnContextMenu += new UserInputEventsSink_OnContextMenuEventHandler(UserInputEventsSink_OnContextMenu);
            m_InputEvts.OnActivateCommand += new UserInputEventsSink_OnActivateCommandEventHandler(UserInputEventsSink_OnActivateCommand);
            //m_modEvents.OnGenerateMember += new ModelingEventsSink_OnGenerateMemberEventHandler(ModelingEventsSink_OnGenerateMember);
            m_asmbllEvts.OnNewOccurrence += new AssemblyEventsSink_OnNewOccurrenceEventHandler(AssemblyEventsSink_OnNewOccurrence);


//             }
// 			catch(Exception e)
// 			{
// 				MessageBox.Show(e.ToString());
// 			}
        }

        void uInterfaceEv_OnEnvironmentChange(Inventor.Environment Environment, EnvironmentStateEnum EnvironmentState, EventTimingEnum BeforeOrAfter, NameValueMap Context, out HandlingCodeEnum HandlingCode)
        {
            HandlingCode = HandlingCodeEnum.kEventHandled;
            if (BeforeOrAfter != EventTimingEnum.kAfter && EnvironmentState != EnvironmentStateEnum.kActivateEnvironmentState) return;
            if (Environment.DisplayName == "2D эскиз")
            {
                PartDocument pDoc = m_inventorApplication.ActiveDocument as PartDocument;
                if (pDoc == null) return;
                data = new InvAddIn.DataToArray(pDoc, m_ComboBox._ComboBoxDef);
            }
        }

//         void m_appEvts_OnActivateDocument(_Document DocumentObject, EventTimingEnum BeforeOrAfter, NameValueMap Context, out HandlingCodeEnum HandlingCode)
//         {
//             HandlingCode = HandlingCodeEnum.kEventHandled;
//             if (DocumentObject.DocumentType == DocumentTypeEnum.kPartDocumentObject && BeforeOrAfter == EventTimingEnum.kAfter)
//             {
//                 data = new InvAddIn.DataToArray(DocumentObject as PartDocument, m_ComboBox._ComboBoxDef);
//             }
//         }

        static public bool activeteForm()
        {
            bool flag = true;
                    foreach (var item in Macros.StandardAddInServer.forms)
                    {
                        if (item != null && item.Created)
                        { flag = false; item.Activate(); }
                    }
                    return flag;
        }
        private void AssemblyEventsSink_OnNewOccurrence(_AssemblyDocument DocumentObject, ComponentOccurrence Occurrence, EventTimingEnum BeforeOrAfter, NameValueMap Context, out HandlingCodeEnum HandlingCode)
        {
            HandlingCode = HandlingCodeEnum.kEventNotHandled;
            try
            {
                if (BeforeOrAfter == EventTimingEnum.kAfter && Occurrence.SurfaceBodies[1].Name == "Половинка1" && Occurrence.SurfaceBodies[2].Name == "Половинка2")
                {
                    PartDocument doc = (PartDocument)Occurrence.Definition.Document; Asset aset = null, aset2 = null;
                    if (Occurrence.Name.ToLower().IndexOf("серый") != -1)
                    {
                        aset = InvDoc.u.createColor(doc, "Gray_", "Серый_", 192, 192, 192); aset2 = aset;
                    }

                    else if (Occurrence.Name.ToLower().IndexOf("синий") != -1)
                    {
                        aset = InvDoc.u.createColor(doc, "Blue_", "Синий_", 0, 0, 255); aset2 = aset;
                    }

                    else if (Occurrence.Name.ToLower().IndexOf("желто-") != -1)
                    {
                        aset = InvDoc.u.createColor(doc, "Yelow_", "Желто_", 255, 255, 0);
                        aset2 = InvDoc.u.createColor(doc, "Green_", "Зеленый_", 0, 255, 0);
                    }
                    if (!doc.ComponentDefinition.SurfaceBodies[1].Appearance.Name.EndsWith("_"))
                    {
                        doc.ComponentDefinition.SurfaceBodies[1].Appearance = aset;
                        doc.ComponentDefinition.SurfaceBodies[2].Appearance = aset2;
                        doc.Save();
                    }
                }
            }
            catch (Exception)
            {
                throw;
            }
            HandlingCode = HandlingCodeEnum.kEventHandled;
        }
        private void UserInputEventsSink_OnContextMenu(SelectionDeviceEnum SelectionDevice, NameValueMap AdditionalInfo, CommandBar CommandBar)
        {
//             try
//             {                                                                        
                SelectSet set = m_inventorApplication.ActiveDocument.SelectSet;

                if (m_inventorApplication.ActiveDocument.DocumentType == DocumentTypeEnum.kAssemblyDocumentObject)
                {
                    CommandBar.Controls.AddButton(m_Asm._ButtonDefinition);
                }
                else if (m_inventorApplication.ActiveDocument.DocumentType == DocumentTypeEnum.kPartDocumentObject)
                {
                    //if (set.Count == 1 && set[1] is PlanarSketch)
                    //CommandBar.Controls.AddButton(m_Sketches._ButtonDefinition);
                }
                else if (m_inventorApplication.ActiveDocument.DocumentType == DocumentTypeEnum.kDrawingDocumentObject)
                {
                    if ((set.Count == 1 && set[1] is DrawingCurveSegment) || (set.Count == 2 && set[1] is DrawingCurveSegment && set[2] is DrawingCurveSegment))
                    {
                        CommandBar.Controls.AddButton(m_surfBtn._ButtonDefinition);
                        CommandBar.Controls.AddButton(m_Gab._ButtonDefinition, 1);
                    }
                    else if (set.Count == 1 && set[1] is DrawingView)
                    {
                        CommandBar.Controls.AddButton(m_Gab._ButtonDefinition, 1);
                        CommandBar.Controls.AddButton(m_spline._ButtonDefinition, 1);
                        CommandBar.Controls.AddButton(m_xmlBreak._ButtonDefinition, 1);
                    }
                    else if (m_inventorApplication.ActiveEditObject is Sheet)
                    {
                        CommandBar.Controls.AddButton(m_ListBtn._ButtonDefinition);
                    }
                }

                if (set.Count > 1) CommandBar.Controls.AddButton(m_xmlBreak._ButtonDefinition, 1);
                if (set.Count == 1 && set[1] is SketchedSymbol /*ObjectTypeEnum.kSketchedSymbolObject*/)
                {
                    SketchedSymbol ss = (SketchedSymbol)m_inventorApplication.ActiveDocument.SelectSet[1];
                    string name = ss.Name;
                    InvAddIn.TablesButton.nameTable = name;
                    CommandBar.Controls.AddButton(m_Tbl._ButtonDefinition);
                }
                else if ((set.Count > 0) && (set[1] is Face || set[1] is FaceProxy || set[1] is HoleFeature))
                {
                    CommandBar.Controls.AddButton(m_sketchInModel._ButtonDefinition);
                }
                else if (set.Count == 1 && set[1] is PlanarSketch)
                {
                    CommandBar.Controls.AddButton(m_VarBtn._ButtonDefinition, 1);
                }
            /*}*/
//             catch 
//             {
// 
//             }
        }

        private void m_appEvts_OnOpenDocument(_Document DocumentObject, string FullDocumentName, EventTimingEnum BeforeOrAfter, NameValueMap Context, out HandlingCodeEnum HandlingCode)
        {
            HandlingCode = HandlingCodeEnum.kEventNotHandled;
            try
            {
                if (DocumentObject != null && DocumentObject.DocumentType == DocumentTypeEnum.kDrawingDocumentObject && 
                    BeforeOrAfter == EventTimingEnum.kAfter)
                {
                    Inventor.DrawingDocument m_Drw = (Inventor.DrawingDocument)DocumentObject;
                    Property p = InvDoc.u.getProp(InvDoc.u.referendedDoc(m_Drw as Document), "Изв");
                    if (p != null)
                    {
                        string date = "";
                        string name = p.Value.ToString();
                        p = InvDoc.u.getProp(InvDoc.u.referendedDoc(m_Drw as Document), "ИзвД");
                        if (p != null) date = p.Value.ToString();
                        name = "ТПМШ." + name + "-" + date.Substring(date.LastIndexOf(".")+1);
                        InvAddIn.Drawings.addIzv(m_Drw, name, date);
                    }
                    if (InvDoc.u.referendedDoc(m_Drw as Document).SubType == "{9C464203-9BAE-11D3-8BAD-0060B0CE6BB4}")
                    {
                        bool find = true;
                        PartDocument pDoc = (PartDocument)InvDoc.u.referendedDoc(m_Drw as Document);
                        string materialDownLine = pDoc.PropertySets[4]["MaterialDownLine"].Value.ToString();
                        string materialUpLine = pDoc.PropertySets[4]["MaterialUpLine"].Value.ToString();
                        string materialLine = pDoc.PropertySets[4]["MaterialLine"].Value.ToString();
                        string materialCenter = pDoc.PropertySets[4]["MaterialCenter"].Value.ToString();
                        foreach (TextStyle txtstl in m_Drw.StylesManager.TextStyles)
                        {
                            if (txtstl.Name == "Material center")
                            {
                                find = false;
                            }
                        }
                        if (find)
                        {
                            Inventor.Application invApp = (Inventor.Application)m_Drw.Parent;
                            drwTemplate = invApp.DesignProjectManager.ActiveDesignProject.TemplatesPath;
                            InvDoc.XML name = new InvDoc.XML(@"C:\ProgramData\Autodesk\Inventor Addins\TemplatePath.xml");
                            System.Collections.Generic.List<string> strs = new System.Collections.Generic.List<string>();
                            strs = name.ReadXML("Template", "DrawingTemplate");
                            drwTemplate = drwTemplate + strs[0];
                            Inventor.DrawingDocument tmpDoc = (Inventor.DrawingDocument)invApp.Documents.Open(drwTemplate, false);
                            TitleBlockDefinition def = tmpDoc.TitleBlockDefinitions[1].CopyTo((Inventor._DrawingDocument)m_Drw, true);
                            tmpDoc.Close(true);
                        }
                        if (materialDownLine.Count() > 30)
                            findStyle(m_Drw.TitleBlockDefinitions[1], m_Drw.ActiveSheet.TitleBlock, "Material rigth", false, 0.65);
                        else findStyle(m_Drw.TitleBlockDefinitions[1], m_Drw.ActiveSheet.TitleBlock, "Material rigth", false, 1);
                        if (materialUpLine.Count() > 30)
                            findStyle(m_Drw.TitleBlockDefinitions[1], m_Drw.ActiveSheet.TitleBlock, "Material up", false, 0.65);
                        else findStyle(m_Drw.TitleBlockDefinitions[1], m_Drw.ActiveSheet.TitleBlock, "Material up", false, 1);
                        if (materialLine.Count() > 4)
                            findStyle(m_Drw.TitleBlockDefinitions[1], m_Drw.ActiveSheet.TitleBlock, "Material left", false, 0.8);
                        else findStyle(m_Drw.TitleBlockDefinitions[1], m_Drw.ActiveSheet.TitleBlock, "Material left", false, 1);

                        InvDoc.u.addProp((Document)m_Drw, "List", "");
                        if (m_Drw.Sheets.Count == 1) m_Drw.PropertySets[4]["List"].Value = "";
                        else m_Drw.PropertySets[4]["List"].Value = "1";
                        try
                        {
                            if (materialUpLine.Count() == 0)
                                m_Drw.StylesManager.TextStyles["Material up"].Underline = false;
                            else m_Drw.StylesManager.TextStyles["Material up"].Underline = true;
                        }
                        catch {}
                    }
                    else
                    {
                        try
                        {
                            m_Drw.StylesManager.TextStyles["Material up"].Underline = false;
                        }
                        catch {}
                    }
               }
            }
            catch {}
            HandlingCode = HandlingCodeEnum.kEventHandled;
        }

        private void findStyle(TitleBlockDefinition tbd, TitleBlock tb, string styleName, bool edit, double scale)
        {
            DrawingSketch sketch;
            sketch = tbd.Sketch;
            bool flag = false;
            if (edit) tbd.Edit(out sketch);
            foreach (Inventor.TextBox text in sketch.TextBoxes)
            {
                string ft = text.FormattedText;
                int indx = ft.IndexOf("MaterialDownLine");
                if (indx != -1 && styleName == "Material rigth")
                {
                    string reztxt = tb.GetResultText(text);
                    if (text.Style.WidthScale != scale)
                    { text.Style.WidthScale = scale; flag = true; }
                }
                indx = ft.IndexOf("MaterialUpLine");
                if (indx != -1 && styleName == "Material up")
                {
                    string reztxt = tb.GetResultText(text);
                    if (text.Style.WidthScale != scale)
                    { text.Style.WidthScale = scale; flag = true; }
                }
                indx = ft.IndexOf("MaterialLine");
                if (indx != -1 && styleName == "Material left")
                {
                    string reztxt = tb.GetResultText(text);
                    if (text.Style.WidthScale != scale)
                        text.Style.WidthScale = scale;
                }
            }
            if (edit) tbd.ExitEdit();
        }

        private void UserInputEventsSink_OnActivateCommand(string CommandName, NameValueMap Context)
        {
            if (Balloon == "1" && CommandName == "DrawingBalloonCmd")
            {
                balloons = new System.Collections.Generic.List<Balloon>();
                m_Doc = m_inventorApplication.ActiveDocument;
                if (m_Doc.DocumentType == DocumentTypeEnum.kDrawingDocumentObject)
                {
                    Inventor.DrawingDocument m_Drw = (Inventor.DrawingDocument)m_Doc;
                    try
                    {
                        foreach (Balloon bal in m_Drw.ActiveSheet.Balloons)
                        {
                            balloons.Add(bal);
                        }
                    }
                    catch (Exception ex) { MessageBox.Show(ex.ToString()); }
                }
            }
            if (Alp == "1" && CommandName == "AppDeleteCmd")
            {
                try
                {
                    Inventor.DrawingDocument m_drw = (Inventor.DrawingDocument)m_inventorApplication.ActiveDocument;
                    SelectSet ss = m_drw.SelectSet;
                    if (ss.Count != 0)
                    {
                        try
                        {
                            DrawingView view = (DrawingView)ss[1];
                            deleteView = true;
                        }
                        catch { }
                    }

                }
                catch (Exception)
                {

                    throw;
                }
            }
        }

        private void UserInputEventsSink_OnTerminate(string CommandName, NameValueMap Context)
        {
            if (Balloon == "1" && CommandName == "DrawingBalloonCmd")
            {
                m_Doc = m_inventorApplication.ActiveDocument;
                if (m_Doc.DocumentType == DocumentTypeEnum.kDrawingDocumentObject &&
                    InvDoc.u.referendedDoc(m_Doc).DocumentType == DocumentTypeEnum.kAssemblyDocumentObject)
                {
                    try
                    {
                        Inventor.DrawingDocument m_Drw = (Inventor.DrawingDocument)m_Doc;
                        Inventor.AssemblyDocument m_AsmDoc = (Inventor.AssemblyDocument)InvDoc.u.referendedDoc(m_Doc);
                        foreach (Balloon bal in m_Drw.ActiveSheet.Balloons)
                        {
                            if (balloons.Find(delegate(Inventor.Balloon b) { return b.Equals(bal); }) == null)
                            {
                                addBalloonValueSet(bal, m_AsmDoc);
                            }
                        }
                        
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.ToString());
                    }
                }
            }
            if (Alp == "1" && CommandName.StartsWith("Drawing") && CommandName.EndsWith("ViewCmd"))
            {
                DrawingLabel((Inventor.DrawingDocument)m_inventorApplication.ActiveDocument);
            }
            if (deleteView)
            {
                DrawingLabel((Inventor.DrawingDocument)m_inventorApplication.ActiveDocument);
                deleteView = false;
            }
        }

        private void UserInputEventsSink_OnStopCommand(CommandIDEnum CommandID)
        {
            if (CommandID == CommandIDEnum.kLineCommand)
            {

            }
        }

        static public void addBalloonValueSet(Inventor.Balloon m_ball, Inventor.AssemblyDocument m_Asm)
        {
            lst = new System.Collections.Generic.List<string>();
            bool flag = false;
            ComponentOccurrence DefaultCO = null;
            lstco = new System.Collections.Generic.List<ComponentOccurrence>();
            InvDoc.XML lstxml = new InvDoc.XML(@"C:\ProgramData\Autodesk\Inventor Addins\StandartElements.xml");
            oldval = new System.Collections.Generic.List<string>();
            if (lst.Count == 0) 
            lst = lstxml.ReadXML("StandartElements", "Value");
            ComponentDefinition def = m_ball.BalloonValueSets[1].ReferencedRow.BOMRow.ComponentDefinitions[1];
            foreach (ComponentOccurrence compOcc in m_Asm.ComponentDefinition.Occurrences)
            {
                if (compOcc.Definition == def){         
                    DefaultCO = compOcc;
                    break;
                }
            }
            findToBalloon(DefaultCO);
            if (lstco.Count >= 1)
                for (int i = lstco.Count-1; i >= 0; i--)
                {
                    m_ball.BalloonValueSets.Add(lstco[i]);
                }
            if (flag)
            {
            }
        }

        //private void findUsingRay(Inventor.AssemblyComponentDefinition acd, ComponentOccurrence co)
        //{
        //    Inventor.Point startPt = co.SurfaceBodies[1].Vertices[1].Point;
        //    Inventor.UnitVector uvec = co.SurfaceBodies[1].Faces[1].Evaluator.
        //   acd.FindUsingRay(startPt,)
        //}

        private string nameFile(string path, bool with)
        {
            int lastInd = path.LastIndexOf('\\');
            path = path.Substring(lastInd+1, path.Length - lastInd-1);
            if (with)
            {
                return path;
            }
            else
            {
                if (path.LastIndexOf('.') != 0)
                {
                    return path.Substring(0, path.Length - 4);
                }
                else return path;
            }
        }

        private void sortBalloon(ref System.Collections.Generic.List<ComponentOccurrence> lstco)
        {
            Inventor.AssemblyComponentDefinition oAsmDoc = (AssemblyComponentDefinition)lstco[1].ContextDefinition;
            Inventor.BOMView bomView = oAsmDoc.BOM.BOMViews["Структурированный"];
            BOMRowsEnumerator bomRows = bomView.BOMRows;
            int i, j;
            lstco.Sort(
                delegate(ComponentOccurrence co1, ComponentOccurrence co2)
            {
                i = 0; j = 0;
                foreach (BOMRow row in bomRows)
                {
                    if(row.ReferencedFileDescriptor.FullFileName == co1.ReferencedDocumentDescriptor.FullDocumentName)
                    i = Convert.ToInt16(row.ItemNumber);
                    if(row.ReferencedFileDescriptor.FullFileName == co2.ReferencedDocumentDescriptor.FullDocumentName)
                    j = Convert.ToInt16(row.ItemNumber);
                    if (i != 0 && j != 0) break;
                }
                if (i>j) return 1;
                    else return -1;
             }
            );
        }

        public static ComponentOccurrence findToBalloon(ComponentOccurrence co)
        {
            if (co != null)
            {
                try
                {
                    foreach (AssemblyConstraint item in co.Constraints)
                    {
                        if (item.AffectedOccurrenceOne != null && item.AffectedOccurrenceOne.Name != co.Name)
                        {
                            ComponentOccurrence tmpOcc = item.AffectedOccurrenceOne;
                            if (!oldval.Contains(tmpOcc.ReferencedDocumentDescriptor.FullDocumentName))
                            {
                                string name1 = ((Document)tmpOcc.Definition.Document).PropertySets[3][14].Value.ToString();
                                if (lst.Contains(name1))
                                {
                                    oldval.Add(tmpOcc.ReferencedDocumentDescriptor.FullDocumentName);
                                    lstco.Add(tmpOcc);
                                    findToBalloon(tmpOcc);
                                }
                            }
                        }
                    }
                }
                catch { }
                return null;
            }
                return null;
        }

        public void Deactivate()
        {
            try
            {

            m_appEvts.OnSaveDocument -= new ApplicationEventsSink_OnSaveDocumentEventHandler(DocEvts_OnSaveHdlr);
            m_SketchEvts.OnNewSketch -= new SketchEventsSink_OnNewSketchEventHandler(SketchEvts_OnNewSketch);
            m_appEvts = null;
            m_SketchEvts = null;
            Marshal.ReleaseComObject(m_inventorApplication);
            m_inventorApplication = null;
            if (m_partSketchRibbonPanel != null)
                m_partSketchRibbonPanel.Delete();

            GC.Collect();
            GC.WaitForPendingFinalizers();

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        public void ExecuteCommand(int commandID)
        {
            // Note:this method is now obsolete, you should use the 
            // ControlDefinition functionality for implementing commands.
        }

        public object Automation
        {

            get
            {
                // TODO: Add ApplicationAddInServer.Automation getter implementation
                return null;
            }
        }
        

        private void DocEvts_OnSaveHdlr(_Document DocumentObject, EventTimingEnum BeforeOrAfter, NameValueMap Context, out HandlingCodeEnum HandlingCode)
        {
            HandlingCode = HandlingCodeEnum.kEventNotHandled;
            try
            {
                m_Doc = DocumentObject;
                if (m_Doc.DocumentType == DocumentTypeEnum.kDrawingDocumentObject)
                {
                    if (BeforeOrAfter == EventTimingEnum.kBefore)
                    {
                        Inventor.DrawingDocument m_DrwDoc = (DrawingDocument)m_Doc;
                        if (Arrow == "1")
                        {
                            Inventor.Application invApp = (Inventor.Application)m_Doc.Parent;
                            invApp.ScreenUpdating = false;
                            foreach (Inventor.HoleThreadNote htn in m_DrwDoc.ActiveSheet.DrawingNotes.HoleThreadNotes)
                            {
                                if (htn.FirstArrowheadType != ArrowheadTypeEnum.kFilledArrowheadType)
                                {
                                    htn.FirstArrowheadType = ArrowheadTypeEnum.kFilledArrowheadType;
                                    htn.SingleDimensionLine = false;
                                    htn.SecondArrowheadType = ArrowheadTypeEnum.kFilledArrowheadType;
                                    htn.LeaderFromCenter = true;
                                }
                            }
                            invApp.ScreenUpdating = true;
                        }
                    }
                }
                m_Doc = null;

            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

            HandlingCode = HandlingCodeEnum.kEventHandled;
        }

        public static void settings()
        {
            System.Collections.Generic.List<string> attr = new System.Collections.Generic.List<string>();
            System.Collections.Generic.List<string> tmplst = new System.Collections.Generic.List<string>();
            InvDoc.XML set = new InvDoc.XML(@"C:\ProgramData\Autodesk\Inventor Addins\Modules.xml");
            tmplst = set.ReadXML("Spec","Spec",ref attr);
            if (tmplst.Count != 0) spec = tmplst[0];
            tmplst = set.ReadXML("BreakHor", "BreakHor", ref attr);
            if (tmplst.Count != 0) BreakHor = tmplst[0];
            tmplst = set.ReadXML("BreakVert", "BreakVert", ref attr);
            if (tmplst.Count != 0) BreakVert = tmplst[0];
            tmplst = set.ReadXML("Alp", "Alp", ref attr);
            if (tmplst.Count != 0) Alp = tmplst[0];
            tmplst = set.ReadXML("Balloon", "Balloon", ref attr);
            if (tmplst.Count != 0) Balloon = tmplst[0];
            tmplst = set.ReadXML("Material", "Material", ref attr);
            if (tmplst.Count != 0) Mat = tmplst[0];
            tmplst = set.ReadXML("OriginPoint", "OriginPoint", ref attr);
            if (tmplst.Count != 0) OriginPoint = tmplst[0];
            tmplst = set.ReadXML("Arrow", "Arrow", ref attr);
            if (tmplst.Count != 0) Arrow = tmplst[0];
            tmplst = set.ReadXML("DXF", "DXF", ref attr);
            if (tmplst.Count != 0) DXFFromPDF = tmplst[0];
            tmplst = set.ReadXML("Custom", "Custom", ref attr);
            if (tmplst.Count != 0) StandartPart = tmplst[0];
            tmplst = set.ReadXML("Print", "Print", ref attr);
            if (tmplst.Count != 0) Printflag = tmplst[0];
            tmplst = set.ReadXML("Dimensions", "Dimensions", ref attr);
            if (tmplst.Count != 0) dim = tmplst[0];
        }

        private void DrawingLabel(DrawingDocument DrwDoc)
        {
            try
            {
                if (lstSort == null)
                {
                    lstSort = new System.Collections.Generic.List<string>();
                    attr = new string[3];
                    attr[0] = "Sort";
                    attr[1] = "FontSize";
                    attr[1] = "FontSizeSmall";
                    InvDoc.XML viewNames = new InvDoc.XML( @"C:\ProgramData\Autodesk\Inventor Addins\ViewNames.xml");
                    lstSort = viewNames.ReadXML("Sort", "Value", ref attr);
                }
                string Alf = attr[0].ToString();
                Double fontSize = Convert.ToDouble(attr[1].ToString()) / 10;
                Double fontSizeSmall = Convert.ToDouble(attr[2].ToString()) / 10;
                System.Collections.Generic.List<DrawingView> SecDV = new System.Collections.Generic.List<DrawingView>();
                System.Collections.Generic.List<DrawingView> StDV = new System.Collections.Generic.List<DrawingView>();
                System.Collections.Generic.List<DrawingView> DetDV = new System.Collections.Generic.List<DrawingView>();
                System.Collections.Generic.List<DrawingView> PrDV = new System.Collections.Generic.List<DrawingView>();
                System.Collections.Generic.List<DrawingView> AuxDV = new System.Collections.Generic.List<DrawingView>();
                double scale = DrwDoc.ActiveSheet.DrawingViews[1].Scale;

                foreach (Inventor.Sheet sheet in DrwDoc.Sheets)
                {
                    foreach (Inventor.DrawingView view in sheet.DrawingViews)
                    {
                        if (view.ViewType == DrawingViewTypeEnum.kSectionDrawingViewType && view.ShowLabel == true &&
                            view.Label.FormattedText.IndexOf("<DrawingViewName/>") != -1)
                            SecDV.Add(view);
                        if (view.ViewType == DrawingViewTypeEnum.kStandardDrawingViewType && view.ShowLabel == true &&
                            view.Label.FormattedText.IndexOf("<DrawingViewName/>") != -1)
                            StDV.Add(view);
                        if (view.ViewType == DrawingViewTypeEnum.kDetailDrawingViewType && view.ShowLabel == true &&
                            view.Label.FormattedText.IndexOf("<DrawingViewName/>") != -1)
                            DetDV.Add(view);
                        if (view.ViewType == DrawingViewTypeEnum.kProjectedDrawingViewType && view.ShowLabel == true &&
                            view.Label.FormattedText.IndexOf("<DrawingViewName/>") != -1)
                            PrDV.Add(view);
                        if (view.ViewType == DrawingViewTypeEnum.kAuxiliaryDrawingViewType && view.ShowLabel == true &&
                            view.Label.FormattedText.IndexOf("<DrawingViewName/>") != -1)
                            AuxDV.Add(view);
                    }
                    foreach (string str in lstSort)
                    {
                        switch (str)
                        {
                            case "Разрезы":
                                addView(SecDV, scale, Alf, fontSize, fontSizeSmall, true);
                                Alf = Alf.Substring(SecDV.Count, Alf.Length - SecDV.Count);
                                break;
                            case "Выносные виды":
                                addView(DetDV, scale, Alf, fontSize, fontSizeSmall, false);
                                Alf = Alf.Substring(DetDV.Count, Alf.Length - DetDV.Count);
                                break;
                            case "Проекционные виды":
                                addView(PrDV, scale, Alf, fontSize, fontSizeSmall, false);
                                Alf = Alf.Substring(PrDV.Count, Alf.Length - PrDV.Count);
                                break;
                            case "Вспомогательные виды":
                                addView(AuxDV, scale, Alf, fontSize, fontSizeSmall, false);
                                Alf = Alf.Substring(AuxDV.Count, Alf.Length - AuxDV.Count);
                                break;
                            case "Стандартные измененные виды":
                                addView(StDV, scale, Alf, fontSize, fontSizeSmall, false);
                                Alf = Alf.Substring(StDV.Count, Alf.Length - StDV.Count);
                                break;
                        }
                    }
                    SecDV.Clear(); StDV.Clear(); DetDV.Clear(); PrDV.Clear(); AuxDV.Clear();
                }
            }
            catch (Exception ex)
            {
                //MessageBox.Show(ex.ToString());
            }

        }

        private void addView(System.Collections.Generic.List<DrawingView> drwviews, double scale, string labels, double FontSize, double FontSizeSmall, bool section)
        {
            if (drwviews.Count == 0) return;
            int i = 0;
            string formatstring;
            foreach (DrawingView view in drwviews)
            {
                int start = view.Label.FormattedText.LastIndexOf('>');
                string str = view.Label.FormattedText.Substring(start+1, view.Label.FormattedText.Length - start-1);
                if (str == " )") str = "";
                if (!section)
                formatstring = styleOverride("","<DrawingViewName/>", "FontSize='" + FontSize + "'");
                else formatstring = styleOverride("", "<DrawingViewName/>-<DrawingViewName/>", "FontSize='" + FontSize + "'");

                if (view.Scale == scale) view.Label.FormattedText = formatstring + str;
                else
                {
                    formatstring = styleOverride(formatstring, "( <DrawingViewScale/> )", "FontSize='" + FontSizeSmall + "'");
                    view.Label.FormattedText = formatstring + str;
                }
                view.Name = labels[i].ToString();
                i++;
            }
        }

        private string styleOverride(string oldstr ,string str, string val)
        {
            return oldstr + "<StyleOverride " + val + ">" + str + "</StyleOverride>";
        }

        private void SketchEvts_OnNewSketch(_Document DocumentObject, Sketch Sketch, EventTimingEnum BeforeOrAfter, NameValueMap Context, out HandlingCodeEnum HandlingCode)
        {
            HandlingCode = HandlingCodeEnum.kEventNotHandled;
            try
            {
                if (OriginPoint == "1")
                {
                    m_Doc = DocumentObject;
                    PlanarSketch pSketch;
                    if (m_Doc.DocumentType == DocumentTypeEnum.kPartDocumentObject)
                    {
                        if (BeforeOrAfter == EventTimingEnum.kAfter)
                        {
                            if (Sketch.Type == ObjectTypeEnum.kPlanarSketchObject)
                            {
                                pSketch = (PlanarSketch)Sketch;
                                if (pSketch.Type != ObjectTypeEnum.kDrawingSketchObject)
                                {
                                    PartDocument m_pDoc;
                                    m_pDoc = (PartDocument)m_Doc;
                                    pSketch.OriginPoint = m_pDoc.ComponentDefinition.WorkPoints[1];
                                    DocumentObject.Update();
                                    m_pDoc = null;
                                }
                                pSketch = null;
                            }
                        }
                    }
                    m_Doc = null;
                }
            }
            catch {}
            HandlingCode = HandlingCodeEnum.kEventHandled;
        }

        #endregion
    }
}
