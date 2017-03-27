using System;
using System.Windows.Forms;
using System.Drawing;
using Inventor;

namespace InvAddIn
{
    class BalloonOp: Button
    {
        public static MyBalloon m_Balloon;
        public Inventor.Document pDoc { get; set; }
        public static MyBalloon getBalloon
        {
            get
            {
                return m_Balloon;
            }
        }

        #region "Methods"
        public BalloonOp(string displayName, string internalName, CommandTypesEnum commandType, string clientId, string description, string tooltip, Icon standardIcon, Icon largeIcon, ButtonDisplayEnum buttonDisplayType)
            : base(displayName, internalName, commandType, clientId, description, tooltip, standardIcon, largeIcon, buttonDisplayType)
        {

        }
        public BalloonOp(string displayName, string internalName, CommandTypesEnum commandType, string clientId, string description, string tooltip, ButtonDisplayEnum buttonDisplayType)
            : base(displayName, internalName, commandType, clientId, description, tooltip, buttonDisplayType)
        {

        }

        protected override void ButtonDefinition_OnExecute(NameValueMap context)
        {
            //System.Windows.Forms.Application.Run(m_Break = new BreakOp(Inventor.Document oDoc));
            m_Balloon = new MyBalloon(InventorApplication.ActiveDocument);
        }

        #endregion
    }
    public class MyBalloon
    {
        private Inventor.DrawingDocument m_DrwDoc;
        private Inventor.AssemblyDocument m_AsmDoc;
        private BalloonOp oBalloon;
        private Inventor.Application invApp;
        private Inventor.Point pt;
        private Inventor.Point2d pt1, pt2;
        private CommandManager cmdMgr;
        private InteractionEvents interEvts;
        private Inventor.MouseEvents mouseEvts;
        private Inventor.SelectEvents selEvts;
        private TransientGeometry m_TG;
        private DrawingView m_View;
        private Inventor.DrawingCurveSegment dcs;
        private Inventor.DrawingCurve dc;
        private object obj;
        private Inventor.Balloon ball;
        private Inventor.GeometryIntent gInt;
        private Inventor.ObjectCollection objs;
        private Inventor.ComponentOccurrence compOcc;
        private Inventor.BreakOperation oBreak;

        public MyBalloon(Inventor.Document oDoc)
        {
            if (oDoc.DocumentType == DocumentTypeEnum.kDrawingDocumentObject && InvDoc.u.referendedDoc(oDoc).DocumentType == DocumentTypeEnum.kAssemblyDocumentObject)
            {
                m_DrwDoc = (Inventor.DrawingDocument)oDoc;
                m_AsmDoc = (Inventor.AssemblyDocument)InvDoc.u.referendedDoc(oDoc);
                invApp = (Inventor.Application)oDoc.Parent;
                cmdMgr = invApp.CommandManager;
                interEvts = cmdMgr.CreateInteractionEvents();
                selEvts = interEvts.SelectEvents;
                selEvts.AddSelectionFilter(SelectionFilterEnum.kDrawingCurveSegmentFilter);
                selEvts.SingleSelectEnabled = true;
                selEvts.OnSelect += new SelectEventsSink_OnSelectEventHandler(selectEvts);
                mouseEvts = interEvts.MouseEvents;
                mouseEvts.MouseMoveEnabled = true;
                mouseEvts.OnMouseClick += new MouseEventsSink_OnMouseClickEventHandler(mouseClick);
                mouseEvts.OnMouseMove += new MouseEventsSink_OnMouseMoveEventHandler(mouseMove);
                objs = invApp.TransientObjects.CreateObjectCollection();

                interEvts.StatusBarText = "Первая точка выноски:";
                interEvts.Start();
                //dcs = (Inventor.DrawingCurveSegment)cmdMgr.Pick(SelectionFilterEnum.kDrawingCurveSegmentFilter, "Выберите компонент:");
                
                //m_Type = obj.GetType();
                //compOcc = (ComponentOccurrence)m_Type.GetProperty("ContainingOccurrence");
                
                //interEvts.InteractionGraphics;
                //if (obj.
                //    dcs.Parent.ModelGeometry.GetType() == Inventor.ObjectTypeEnum.kEdgeProxyObject)
                //{

                //}
            }
        }

        private void selectEvts(ObjectsEnumerator JustSelectedEntities, SelectionDeviceEnum SelectionDevice, Inventor.Point ModelPosition, Point2d ViewPosition, Inventor.View View)
        {
            selEvts.AddToSelectedEntities(JustSelectedEntities[1]);
        }

        private void mouseMove(MouseButtonEnum Button, ShiftStateEnum ShiftKeys, Inventor.Point ModelPosition, Point2d ViewPosition, Inventor.View View)
        {
            
        }

        private void mouseClick(MouseButtonEnum Button, ShiftStateEnum ShiftKeys, Inventor.Point ModelPosition, Point2d ViewPosition, Inventor.View View)
        {
            if (Button == MouseButtonEnum.kLeftMouseButton && pt1 == null)
            {
                pt = ModelPosition;
                pt1 = ViewPosition;
                pt1.X = pt.X;
                pt1.Y = pt.Y;
                interEvts.StatusBarText = "Вторая точка выноски:";
                objs.Add(pt1);
                if (selEvts.SelectedEntities.Count != 0)
                    dcs = (Inventor.DrawingCurveSegment)selEvts.SelectedEntities[1];
                dc = dcs.Parent;
            }
            else if (Button == MouseButtonEnum.kLeftMouseButton && pt2 == null)
            {
                try
                {
                    pt2 = ViewPosition;
                    pt2.X = ModelPosition.X; pt2.Y = ModelPosition.Y;
                    objs.Add(pt2);
                    interEvts.Stop();
                    obj = dc.ModelGeometry;
                    try
                    {
                        compOcc = ((Inventor.FaceProxy)obj).ContainingOccurrence;
                    }
                    catch (Exception)
                    {
                        try
                        {
                            compOcc = ((Inventor.EdgeProxy)obj).ContainingOccurrence;
                        }
                        catch (Exception ex1)
                        {
                            MessageBox.Show(ex1.ToString());
                        }
                    }
                    gInt = m_DrwDoc.ActiveSheet.CreateGeometryIntent(dc);
                    objs.Add(gInt);
                    ball = m_DrwDoc.ActiveSheet.Balloons.Add(objs);
                    ball.BalloonValueSets.Add(m_AsmDoc.ComponentDefinition.Occurrences[1]);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.ToString());
                }
            }
        }
    }
}
