using System;
using System.Runtime.InteropServices;
using Inventor;
using System.Windows.Forms;
using Microsoft.Win32;

namespace InvAddIn
{
   public class Utility
    {
        #region Data Members
        String str;
        Inventor.Document m_doc;
        Inventor.DrawingDocument m_Drw;
        Inventor.AssemblyDocument m_Asm;
        Inventor.Style style;
        #endregion

        public Utility(Document newDoc)
        {
            m_doc = newDoc;
            if (m_doc.DocumentType == DocumentTypeEnum.kAssemblyDocumentObject) m_Asm = (AssemblyDocument)m_doc;
            if (m_doc.DocumentType == DocumentTypeEnum.kDrawingDocumentObject) m_Drw = (DrawingDocument)m_doc;
        }
        public Utility(DrawingDocument newDoc)
        {
            m_Drw = newDoc;
        }

        #region Methods

        public TextStyle addTextStyle(TextStyle oldStyle, string name ,double scale = 1, bool italic = true, HorizontalTextAlignmentEnum hor = HorizontalTextAlignmentEnum.kAlignTextCenter,
            VerticalTextAlignmentEnum vert = VerticalTextAlignmentEnum.kAlignTextMiddle,double size = 3.5, string font = "GOST Common", bool bold = true, bool underline = false)
        {
            TextStyle newStyle = (TextStyle)oldStyle.Copy(name);
            newStyle.Bold = bold; newStyle.Underline = underline;
            newStyle.Font = font; newStyle.FontSize = size / 10; newStyle.HorizontalJustification = hor;
            newStyle.VerticalJustification = vert; newStyle.Italic = italic; newStyle.WidthScale = scale;
            return newStyle;
        }

        //public void PlaceOnSheet(string name, Inventor.Point2d insPt)
        //{
        //    try 
        //    {
        //        foreach (SketchedSymbol ss in m_Drw.ActiveSheet.SketchedSymbols)
        //        {
        //            if (ss.Name.StartsWith(name))
        //            {
        //                insPt = ss.Position;
        //                ss.Delete();
        //            }
        //        }
        //        m_Drw.ActiveSheet.SketchedSymbols.Add(m_Drw.SketchedSymbolDefinitions[name], insPt);
        //    }
        //        catch(Exception ex)
        //    {
        //}
        //}

        #endregion
    }

}
