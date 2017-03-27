using System;
using System.Windows.Forms;
using System.Drawing;
using Inventor;

namespace InvAddIn
{
    internal class MyBreakOperation : Button
    {
        public static BreakOp m_Break;
        public Inventor.Document pDoc { get; set; }
        public static BreakOp getBreak
        {
            get
            {
                return m_Break;
            }
        }

        #region "Methods"
        public MyBreakOperation(string displayName, string internalName, CommandTypesEnum commandType, string clientId, string description, string tooltip, Icon standardIcon, Icon largeIcon, ButtonDisplayEnum buttonDisplayType)
            : base(displayName, internalName, commandType, clientId, description, tooltip, standardIcon, largeIcon, buttonDisplayType)
        {

        }
        public MyBreakOperation(string displayName, string internalName, CommandTypesEnum commandType, string clientId, string description, string tooltip, ButtonDisplayEnum buttonDisplayType)
            : base(displayName, internalName, commandType, clientId, description, tooltip, buttonDisplayType)
        {

        }

        protected override void ButtonDefinition_OnExecute(NameValueMap context)
        {
            //System.Windows.Forms.Application.Run(m_Break = new BreakOp(Inventor.Document oDoc));
            m_Break = new BreakOp(InventorApplication.ActiveDocument, true);
        }

        #endregion
    }

    internal class MyBreakOperationVert : Button
    {
        public static BreakOp m_Break;
        public Inventor.Document pDoc { get; set; }
        public static BreakOp getBreak
        {
            get
            {
                return m_Break;
            }
        }

        #region "Methods"
        public MyBreakOperationVert(string displayName, string internalName, CommandTypesEnum commandType, string clientId, string description, string tooltip, Icon standardIcon, Icon largeIcon, ButtonDisplayEnum buttonDisplayType)
            : base(displayName, internalName, commandType, clientId, description, tooltip, standardIcon, largeIcon, buttonDisplayType)
        {

        }
        public MyBreakOperationVert(string displayName, string internalName, CommandTypesEnum commandType, string clientId, string description, string tooltip, ButtonDisplayEnum buttonDisplayType)
            : base(displayName, internalName, commandType, clientId, description, tooltip, buttonDisplayType)
        {

        }

        protected override void ButtonDefinition_OnExecute(NameValueMap context)
        {
            //InvAddIn.Tables tbl = new Tables(InventorApplication.ActiveDocument, InventorApplication, "fsdf");
            //tbl.addLeader(true);
            m_Break = new BreakOp(InventorApplication.ActiveDocument, false);
        }

        #endregion
    }

    public class BreakOp
    {
        private Inventor.DrawingDocument m_DrwDoc;
        private Inventor.Application invApp;
        private Inventor.Point pt;
        private Inventor.Point2d pt1, pt2;
        private CommandManager cmdMgr;
        private InteractionEvents interEvts;
        private Inventor.MouseEvents mouseEvts;
        private TransientGeometry m_TG;
        private DrawingView m_View;
        private Inventor.BreakOperation oBreak;
        public int num = 10;
        public double gap = 0.2;
        private bool horizont;
        public BreakOp(Inventor.Document oDoc, bool hor)
        {
            if (oDoc.DocumentType == DocumentTypeEnum.kDrawingDocumentObject)
            {
                invApp = (Inventor.Application)oDoc.Parent;
                m_TG = invApp.TransientGeometry;
                cmdMgr = invApp.CommandManager;
                m_DrwDoc = (Inventor.DrawingDocument)oDoc;
                horizont = hor;
                m_View = (DrawingView)cmdMgr.Pick(SelectionFilterEnum.kDrawingViewFilter, "Вид для разрыва:");
                interEvts = cmdMgr.CreateInteractionEvents();
                //interEvts.InteractionDisabled = false;
                mouseEvts = interEvts.MouseEvents;
                mouseEvts.OnMouseClick += new MouseEventsSink_OnMouseClickEventHandler(mouseClick);
                interEvts.StatusBarText = "Первая точка разрыва:";
                
                interEvts.Start();
                
                //m_DrwDoc.ActiveSheet.DrawingViews
                //if  ( m_DrwDoc.SelectSet[1] )
                //pt1 = m_DrwDoc.SelectSet[1];
            }
        }
        private void mouseClick(MouseButtonEnum Button, ShiftStateEnum ShiftKeys, Inventor.Point ModelPosition, Point2d ViewPosition, Inventor.View View)
        {
            if (Button == MouseButtonEnum.kLeftMouseButton)
            {
                if (pt1 == null)
                {
                    //m_View = (DrawingView)View;
                    pt = ModelPosition;
                    pt1 = ViewPosition;
                    pt1.X = pt.X;
                    pt1.Y = pt.Y;
                    interEvts.StatusBarText = "Вторая точка разрыва:";
                }
                else
                {
                    pt = ModelPosition;
                    pt2 = ViewPosition;
                    pt2.X = pt.X;
                    pt2.Y = pt.Y;
                    interEvts.Stop();
                    try
                    {
                        if (horizont)
                       oBreak = m_View.BreakOperations.Add(BreakOrientationEnum.kHorizontalBreakOrientation,
                                            pt1, pt2, BreakStyleEnum.kRectangularBreakStyle, num, gap,1,false);
                        else
                            oBreak = m_View.BreakOperations.Add(BreakOrientationEnum.kVerticalBreakOrientation,
                                                    pt1, pt2, BreakStyleEnum.kRectangularBreakStyle, num, gap, 1, false);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.ToString());
                        mouseEvts.OnMouseClick -= new MouseEventsSink_OnMouseClickEventHandler(mouseClick);
                    }
                }
            }
        }
    }
}