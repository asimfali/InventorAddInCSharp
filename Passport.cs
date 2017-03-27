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
    public partial class Passport : Form
    {
        Inventor.Application invApp;
        DrawingDocument drwDoc;
        LinearGeneralDimension dim;
        Inventor.DimensionStyle stl; 
        Inventor.TextStyle stlTxt;
        string str, format;
        Sheet sheet;

        public Passport(Inventor.Document pDoc, Inventor.Application m_InvApp)
        {
            InitializeComponent();
            invApp = m_InvApp;
            if (pDoc.DocumentType == DocumentTypeEnum.kDrawingDocumentObject)
            {
                drwDoc = (DrawingDocument)pDoc;
                sheet = drwDoc.ActiveSheet;
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                this.Hide();

                foreach (DrawingDimension drwdim in sheet.DrawingDimensions)
                {
                    foreach (Inventor.TextStyle txtStl in drwDoc.StylesManager.TextStyles)
                    {
                        if (txtStl.InUse == true)
                        {
                            if (txtStl.Font != textBox1.Text)
                                txtStl.Font = textBox1.Text;
                            if (txtStl.FontSize != (Convert.ToDouble(textBox2.Text) / 10))
                                txtStl.FontSize = (Convert.ToDouble(textBox2.Text) / 10);
                            if (txtStl.Italic == true)
                                txtStl.Italic = false;
                            if (txtStl.Bold == true)
                                txtStl.Bold = false;
                        }
                    }
                    if (drwdim.Type == ObjectTypeEnum.kLinearGeneralDimensionObject)
                    {
                        dim = (LinearGeneralDimension)drwdim;
                        stl = dim.Style;
                        if (stl.TextStyle.Font != textBox1.Text)
                            stl.TextStyle.Font = textBox1.Text;
                        if (stl.TextStyle.FontSize != (Convert.ToDouble(textBox2.Text)/10))
                            stl.TextStyle.FontSize = (Convert.ToDouble(textBox2.Text)/10);
                        if (stl.TextStyle.Italic == true)
                            stl.TextStyle.Italic = false;
                        if (stl.TextStyle.Bold == true)
                            stl.TextStyle.Bold = false;
                        stl.LinearPrecision = LinearPrecisionEnum.kZeroDecimalPlaceLinearPrecision;
                        stl.AngularPrecision = AngularPrecisionEnum.kZeroDecimalPlaceAngularPrecision;
                        int modelVal = (int)Math.Round(drwdim.ModelValue * 10);
                        str = drwdim.Text.Text; int n;
                        format = "";
                        format = @" Font='" + textBox1.Text + "' ";
                        //format += @" FontSize = '" + textBox2.Text + "' ";
                        //"<DimensionValue/><StyleOverride Italic='True'>A</StyleOverride><StyleOverride Italic='True'><Stack FractionalTextScale='0,700000'>^2</Stack></StyleOverride>"
                        if (!int.TryParse(str, out n) && drwdim.HideValue == true)
                        {
                            format += @" Italic='True' ";
                            format += @" Bold='True' ";
                            if (str.Count() > 1)
                            {
                                drwdim.Text.FormattedText = @"<StyleOverride" + format + @">" + str[0] + @"</StyleOverride>" +
                                    @"<StyleOverride" + format + @">" + "<Stack FractionalTextScale='0,700000'>^"+ str[1] + @"</Stack></StyleOverride>";
                            }
                            else
                            {
                                drwdim.Text.FormattedText = @"<StyleOverride" + format + @">" + str + @"</StyleOverride>";
                            }
                        }
                        if (modelVal >= 3000)
                        {
                            int res = roundDim(modelVal, Convert.ToInt16(textBox4.Text));
                            if (modelVal != res)
                                drwdim.OverrideModelValue = (double)res / 10;
                        }
                    }
                }
                foreach (DrawingView view in sheet.DrawingViews)
                {
                    format = ""; bool flag = true;
                    if (flag)
                    {
                        stlTxt = view.Label.TextStyle;
                       
                        if (stlTxt.Font != textBox1.Text)
                            stlTxt.Font = textBox1.Text;
                        if (stlTxt.FontSize != (Convert.ToDouble(textBox2.Text) / 10))
                            stlTxt.FontSize = (Convert.ToDouble(textBox2.Text) / 10);
                        if (stlTxt.Italic == true)
                            stlTxt.Italic = false;
                        if (stlTxt.Bold == true)
                            stlTxt.Bold = false;
                        flag = false;
                    }
                    str = view.Label.FormattedText;
                    int indx = str.IndexOf("Font='");
                    while (indx != -1)
                    {
                        str = findOverride(str, "Font='", textBox1.Text, textBox1.Text, indx);
                        indx = str.IndexOf("Italic='", indx + 1);
                        str = findOverride(str, "Italic='", "False", "False", indx);
                        indx = str.IndexOf("Font='", indx + 1);
                    }

                    indx = str.IndexOf("<Br/>");
                    if (indx != -1)
                    {
                        string oldstr = str.Substring(0, indx);
                        str = str.Substring(indx, str.Length - indx);
                        indx = str.IndexOf("FontSize='");
                        if (indx != -1)
                        {
                            str = oldstr + findOverride(str, "FontSize='", (Convert.ToDouble(textBox3.Text) / 10).ToString(),"",indx);
                        }
                        view.Label.FormattedText = str;
                    }
                }

                this.Close();
            }
            catch (Exception)
            {
                this.Close();
            }
        }
        private string findOverride(string str,string findstr, string val, string cmpVal, int indx)
        {
            string first = str.Substring(0, indx + findstr.Count());
            string midle = str.Substring(indx + findstr.Count(), str.IndexOf("'", findstr.Count() + indx + 1) - indx - findstr.Count());
            string last = str.Substring(str.IndexOf("'", findstr.Count()+indx+1), str.Length - str.IndexOf("'", findstr.Count()+ indx + 1));
            return (midle != val)? first + val + last : str;
        }

        private string styleOverride(string oldstr, string str, string val)
        {
            return oldstr + "<StyleOverride " + val + ">" + str + "</StyleOverride>";
        }

        public int roundDim(int val, int rnd)
        {
            if (val%rnd != 0)
            {
                val /= rnd;
                if (val % rnd <= (double)rnd/2)
                val = (val) * rnd;
                else
                val = (val + 1) * rnd;
                //l = (int)l; l = (l + 1) * 5;

            }
            return val;
        }
    }
    internal class PassportBtn : Button
    {
        public static Passport m_Passport;
        public Inventor.Document pDoc { get; set; }
        public static Passport getPassport
        {
            get
            {
                return m_Passport;
            }
        }

        #region "Methods"
        public PassportBtn(string displayName, string internalName, CommandTypesEnum commandType, string clientId, string description, string tooltip, Icon standardIcon, Icon largeIcon, ButtonDisplayEnum buttonDisplayType)
            : base(displayName, internalName, commandType, clientId, description, tooltip, standardIcon, largeIcon, buttonDisplayType)
        {

        }
        public PassportBtn(string displayName, string internalName, CommandTypesEnum commandType, string clientId, string description, string tooltip, ButtonDisplayEnum buttonDisplayType)
            : base(displayName, internalName, commandType, clientId, description, tooltip, buttonDisplayType)
        {

        }

        protected override void ButtonDefinition_OnExecute(NameValueMap context)
        {
            //System.Windows.Forms.Application.Run(m_Break = new BreakOp(Inventor.Document oDoc));
            try
            {
                Macros.StandardAddInServer.forms.Add(m_Passport);
                if (Macros.StandardAddInServer.activeteForm()) System.Windows.Forms.Application.Run(m_Passport = new InvAddIn.Passport(InventorApplication.ActiveDocument, InventorApplication));

            }
            catch (Exception)
            {
            }
        }

        #endregion
    }
}
