using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Inventor;


namespace InvAddIn
{
    public partial class Tables : Form
    {
        public Inventor.Application invApp;
        public Inventor.DrawingDocument m_Drw;
        private Inventor.TransientGeometry m_TG;
        private Inventor.SketchedSymbolDefinition m_SketchDef;
        private Inventor.DrawingSketch m_DrwSketch;
        private Inventor.SketchLine m_SketchLine;
        private string[] spl;
        public Setting setting;
        public int cc, cr;
        public string value="";
        public string textheader, columnname, val="";
        public int rowheight, columnwidth;
        private InvDoc.XML tbl;
        private List<string> attrXML, valXML;
        private List<System.Drawing.Point> lstTxt;
        private Dictionary<float, float> textComp, textCompRev;
        private List<string> styleParam; /*номер строки; номер столбца; сжатие;*/
        private char splitChar = ';' ;

        char separator = System.Globalization.CultureInfo.CurrentCulture.NumberFormat.CurrencyDecimalSeparator[0];

        public Tables(Inventor.Document pDoc, Inventor.Application m_InvApp, string name)
        {
            try
            {
            if (pDoc.DocumentType == DocumentTypeEnum.kDrawingDocumentObject)
            {
                m_Drw = (DrawingDocument)pDoc;
                invApp = m_InvApp;
                m_TG = invApp.TransientGeometry;
                EventArgs e = null; Object obj = null;
                textComp = new Dictionary<float, float>();
                textCompRev = new Dictionary<float, float>();
                styleParam = new List<string>();
                textComp.Add(6F, 2F); textComp.Add(8F, 2.5F); textComp.Add(10F, 3F);
                textComp.Add(12F, 3.5F); textComp.Add(14F, 4F); textComp.Add(16F, 4.5F); textComp.Add(18F, 5F); textComp.Add(20F, 5.5F);
                textComp.Add(22F, 6F); textComp.Add(24F, 6.5F); textComp.Add(26F, 7F); textComp.Add(28F, 8F); textComp.Add(30F, 10F);
                foreach (float key in textComp.Keys)
                {
                    textCompRev.Add(textComp[key], key);
                }

            InitializeComponent();
            new DataGridViewRowsReorderBehavior(_dgv);
            setting = new Setting();
            _dgv.DefaultCellStyle.Font = fontStyle(italic: setting.checkBox4.Checked, bold: setting.checkBox5.Checked);
            rowheight = Convert.ToInt16(setting.comboBox8.Text);
            columnwidth = Convert.ToInt16(setting.comboBox9.Text);
            CountColumn_TextChanged(obj, e);
            CountRow_TextChanged(obj, e);
            ColumnDimensions_Leave(obj, e);
            RowDimensions_Leave(obj, e);
            this._dgv.Focus();
            if (name != null)
            {
                this.textBox2.Text = name;
                open(_dgv);
            }
            valXML = new List<string>();
            attrXML = new List<string>();
            if (System.IO.File.Exists(@"C:\ProgramData\Autodesk\Inventor Addins\Tables.xml"))
                {
                    tbl = new InvDoc.XML(@"C:\ProgramData\Autodesk\Inventor Addins\Tables.xml");
                    tbl.ReadXML("Table", ref valXML, ref attrXML);
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
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private void CountColumn_TextChanged(object sender, EventArgs e)
        {
            if (this.CountColumn.Text != "" && this.CountRow.Text != "")
            {

                cc = Convert.ToInt16(this.CountColumn.Text);
                val = "";

                if (this.ColumnDimensions.Text == "")
                {
                    for (int i = 0; i < cc; i++)
                    {
                        val += setting.comboBox9.Text + ';';
                    }
                }
                else val = this.ColumnDimensions.Text;


                if (val != "") spl = val.Split(';');

                if (this._dgv.ColumnCount == 0)
                {
                    this.ColumnDimensions.Text = "";
                }
                else if (this._dgv.ColumnCount < cc)
                {
                    for (int i = this._dgv.Columns.Count; i < cc; i++)
                    {
                        val += setting.comboBox9.Text + ';';
                    }
                }
                else if (this._dgv.ColumnCount > cc)
                {
                    val = "";
                    for (int i = 0; i < cc; i++)
                    {
                        val += spl[i] + ';'; 
                    }
                }
                this.ColumnDimensions.Text = val;

                this._dgv.ColumnCount = Convert.ToInt16(this.CountColumn.Text);
                this._dgv.RowCount = Convert.ToInt16(this.CountRow.Text);
            }

            //if (this.CountColumn.Text != "" && this.CountRow.Text != "")
            //{
            //    int cc, cr; string textheader, columnname;
            //    cc = Convert.ToInt16(this.CountColumn.Text);
            //    cr = Convert.ToInt16(this.CountRow.Text);
            //    if (this._dgv.ColumnCount == 0)
            //    {
            //       this._dgv.ColumnCount = cc; 
            //    }
            //    else if (this._dgv.ColumnCount < cc)
            //    {
            //        for(int i=this._dgv.Columns.Count; i<cc;i++)
            //        {
            //          textheader = "Столбец" + i; columnname = "Column" + i;
            //          this._dgv.Columns.Add(columnname, textheader);
            //        }
            //    }
            //    else if (this._dgv.ColumnCount > cc)
            //    {
            //        for (int i = this._dgv.Columns.Count; i < cc; i++)
            //        {
            //            textheader = "Столбец" + i; columnname = "Column" + i;
            //            this._dgv.Columns.RemoveAt(i);
            //        }
            //    }
            //    else
            //    {
            //    }
            //    //if (this._dgv.RowCount == 0)
            //    //{
            //    //    this._dgv.RowCount = cr; 
            //    //}
            //}
        }

        private void CountRow_TextChanged(object sender, EventArgs e)
        {
             if (this.CountColumn.Text != "" && this.CountRow.Text != "")
            {
                cr = Convert.ToInt16(this.CountRow.Text);

                val = "";

                if (this.RowDimensions.Text == "")
                {
                    for (int i = 0; i < cr; i++)
                    {
                        val += setting.comboBox8.Text + ';';
                    }
                }
                else val = this.RowDimensions.Text;


                if (val != "") spl = val.Split(';');

                if (this._dgv.RowCount == 0)
                {
                    this.RowDimensions.Text = "";
                }
                else if (this._dgv.RowCount < cr)
                {
                    for (int i = this._dgv.Rows.Count; i < cr; i++)
                    {
                        val += setting.comboBox8.Text + ';';
                    }
                }
                else if (this._dgv.RowCount > cr)
                {
                    val = "";
                    for (int i = 0; i < cr; i++)
                    {
                        val += spl[i] + ';'; 
                    }
                }
                this.RowDimensions.Text = val;

                this._dgv.RowCount = Convert.ToInt16(this.CountRow.Text);
                this._dgv.ColumnCount = Convert.ToInt16(this.CountColumn.Text);  
            }
        }

        private void ColumnDimensions_Leave(object sender, EventArgs e)
        {
            try
            {
                if (this.ColumnDimensions.Text == "")
                {
                    for (int i = 0; i < cc; i++)
                    {
                        val += columnwidth.ToString() + ';';
                    }
                }
                else val = this.ColumnDimensions.Text;

                if (val != "") spl = val.Split(';');

                for (int i = 0; i < this._dgv.ColumnCount; i++)
                {
                    if (spl[i] != columnwidth.ToString())
                    {
                        //double fitt = (Convert.ToDouble(spl[i]) / columnwidth);
                        this._dgv.Columns[i].Width = (int)(100 * Convert.ToDouble(spl[i]) / columnwidth);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private void RowDimensions_Leave(object sender, EventArgs e)
        {
            try
            {
                if (this.RowDimensions.Text == "")
                {
                    for (int i = 0; i < cc; i++)
                    {
                        val += rowheight.ToString() + ';';
                    }
                }
                else val = this.RowDimensions.Text;

                if (val != "") spl = val.Split(';');
                //if (this._dgv.RowCount > (spl.Count()-1)) count = this._dgv.RowCount;
                //else count = this.spl.Count();
                for (int i = 0; i < this._dgv.RowCount; i++)
                {
                    if (spl[i] != rowheight.ToString())
                    {
                        this._dgv.Rows[i].Height = (int)(_dgv.RowTemplate.Height * Convert.ToDouble(spl[i]) / rowheight);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private void drawTable(string name, Inventor.Point2d insPt, double[] arrColumns, double[] arrRow,
            List<System.Drawing.Point> lstX, List<System.Drawing.Point> lstY,
            int[] intArr = null, int[] intArr2 = null,
            DataGridView _dgv = null, bool col = false, Point2d insertPt = null)
        {
            try
            {
                invApp.ScreenUpdating = false;
                lstTxt = new List<System.Drawing.Point>();
                //numRow += 1;
                List<Inventor.Point2d> lstpt = new List<Point2d>();
                if (insertPt == null) insertPt = invApp.TransientGeometry.CreatePoint2d(0,0);
                insPt.X -= arrColumns.Sum();
                double sumColumns, sumRows;
                sumColumns = arrColumns.Sum();
                sumRows = arrRow.Sum() - arrRow[arrRow.Count()-1];
                try
                {
                    m_SketchDef = m_Drw.SketchedSymbolDefinitions.Add(name);
                    m_SketchDef.Edit(out m_DrwSketch);
                }
                catch 
                {
                    foreach (SketchedSymbol ss in m_Drw.ActiveSheet.SketchedSymbols)
                    {
                        if (ss.Name.StartsWith(name))
                        {
                            insertPt = ss.Position;
                            ss.Delete();
                        }
                    }
                    m_Drw.SketchedSymbolDefinitions[name].Delete();                   
                    m_SketchDef = m_Drw.SketchedSymbolDefinitions.Add(name);
                    m_SketchDef.Edit(out m_DrwSketch);
                }

                //горизонтальные линии
                for (int i = 0; i < arrRow.Count()-1; i++)
                {
                    lstpt = LinesY(ref lstY, arrColumns, i, invApp.TransientGeometry.CreatePoint2d(insPt.X, insPt.Y));
                    if (lstpt != null)
                    {
                        for (int r = 0; r < lstpt.Count; r += 2)
                        {
                            m_SketchLine = m_DrwSketch.SketchLines.AddByTwoPoints(lstpt[r], lstpt[r + 1]);
                            if (intArr != null && System.Array.IndexOf(intArr, i) != -1)
                            {
                                m_SketchLine.LineWeight = 0.05;
                            }
                        }
                    }
                    else
                    m_SketchLine = m_DrwSketch.SketchLines.AddByTwoPoints(insPt, m_TG.CreatePoint2d(insPt.X + sumColumns, insPt.Y));
                    insPt.Y -= arrRow[i];
                    if (intArr != null && System.Array.IndexOf(intArr, i) != -1)
                    {
                        m_SketchLine.LineWeight = 0.05;
                    }
                }
                lstpt = null;
                m_SketchLine = m_DrwSketch.SketchLines.AddByTwoPoints(insPt, m_TG.CreatePoint2d(insPt.X + sumColumns, insPt.Y));
                m_SketchLine.LineWeight = 0.05;
                insPt.Y += (sumRows);

                //вертикальные линии
                for (int j = 0; j < arrColumns.Length; j++)
                {
                    lstpt = LinesX(ref lstX, arrRow, j, invApp.TransientGeometry.CreatePoint2d(insPt.X,insPt.Y));
                    if (lstpt != null)
                    {
                        for (int r = 0; r < lstpt.Count; r += 2)
                        {
                            m_SketchLine = m_DrwSketch.SketchLines.AddByTwoPoints(lstpt[r], lstpt[r + 1]);
                            if (col || j == 0)
                                m_SketchLine.LineWeight = 0.05;
                        }
                    }
                    else m_SketchLine = m_DrwSketch.SketchLines.AddByTwoPoints(insPt, m_TG.CreatePoint2d(insPt.X, insPt.Y - (sumRows)));
                    if (col || j == 0)
                    m_SketchLine.LineWeight = 0.05;
                    insPt.X += arrColumns[j];
                }
                m_SketchLine = m_DrwSketch.SketchLines.AddByTwoPoints(insPt, m_TG.CreatePoint2d(insPt.X, insPt.Y - (sumRows)));
                m_SketchLine.LineWeight = 0.05;

                bool flag;
                DataGridViewRow _dgvr;
                string str;
                Inventor.TextBox _TB;
                int CC = (_dgv.ColumnCount <= arrColumns.Count()) ? _dgv.ColumnCount : arrColumns.Count();

                for (int j = 0; j < _dgv.Rows.Count - 1; j++)
                {
                    _dgvr = _dgv.Rows[j];
                    insPt.X -= sumColumns;
                    for (int i = 0; i < CC; i++)
                    {
                        if (_dgvr.Cells[i].Value != null && _dgvr.Cells[i].Value.ToString() != "")
                        {
                            //if (i < (_dgvr.Cells.Count - 1) && _dgvr.Cells[i + 1].Value.ToString() == str)
                            //{

                            //}
                            flag = true;
                            for (int num = 0; num < lstTxt.Count; num += 2)
                            {
                                if ((_dgvr.Cells[i].RowIndex >= lstTxt[num].Y && _dgvr.Cells[i].RowIndex <= lstTxt[num + 1].Y) &&
                                    (_dgvr.Cells[i].ColumnIndex >= lstTxt[num].X && _dgvr.Cells[i].ColumnIndex <= lstTxt[num + 1].X))
                                {
                                    flag = false;
                                    break;
                                }
                            }
                            _TB = null;
                            if (flag)
                            {
                               _TB = text(_dgv, _dgvr.Cells[i], ref lstTxt, ref arrColumns, ref arrRow);
                            }

                            str = _dgvr.Cells[i].Value.ToString();

                            if (_TB != null)
                            {
                                DataGridViewContentAlignment al = _dgvr.Cells[i].Style.Alignment;
                                string align = "";
                                switch (al)
                                {
                                    case DataGridViewContentAlignment.MiddleCenter:
                                        align = "По центру";
                                        break;
                                    case DataGridViewContentAlignment.MiddleLeft:
                                        align = "Слева";
                                        break;
                                    case DataGridViewContentAlignment.MiddleRight:
                                        align = "Справа";
                                        break;
                                    default:
                                        align = setting.comboBox7.Text;
                                        break;
                                }

                                Func<string, double> conv = s =>
                                {
                                    char rep = (separator == ',') ? '.' : ',';
                                    return (s.IndexOf(separator) != -1) ? Convert.ToDouble(s) : Convert.ToDouble(s.Replace(rep, separator));
                                };

                                if (_dgvr.Cells[i].Style.Font != null)
                                {
                                    float h = _dgvr.Cells[i].Style.Font.Size;
                                    h = (float)(Math.Round(h * 100)) / 100;
                                    textComp.TryGetValue(h, out h);
                                    if (h == 3.5)
                                    {
                                        h = (float)conv(setting.comboBox1.SelectedItem.ToString());
                                    }
                                    formatText(ref _TB, align, arrColumns[i], width: conv(setting.comboBox2.Text) / 100,
                                      underline: _dgvr.Cells[i].Style.Font.Underline, italic: _dgvr.Cells[i].Style.Font.Italic, bold: _dgvr.Cells[i].Style.Font.Bold,
                                      textH: h / 10, font: _dgv.DefaultCellStyle.Font.Name);
                                }

                                else formatText(ref _TB, align, arrColumns[i], width: conv(setting.comboBox2.Text) / 100,
                                    underline: _dgv.DefaultCellStyle.Font.Underline, italic: _dgv.DefaultCellStyle.Font.Italic, bold: _dgv.DefaultCellStyle.Font.Bold,
                                    textH: conv(setting.comboBox1.Text) / 10, font: _dgv.DefaultCellStyle.Font.Name);

                                foreach (string tmp in styleParam)
                                {
                                    string[] spl = tmp.Split(';');
                                    if (spl[0] == _dgvr.Cells[i].RowIndex.ToString() && spl[1] == _dgvr.Cells[i].ColumnIndex.ToString())
                                    {
                                        _TB.WidthScale = Convert.ToDouble(spl[2]) / 100;
                                        break;
                                    }
                                }
                            }
                        }
                        insPt.X += arrColumns[i];
                    }
                    insPt.Y -= arrRow[j];
                }

                m_DrwSketch.SketchPoints.Add(m_TG.CreatePoint2d(0, 0)).InsertionPoint = true;
                if (setting.checkBox2.Checked)
                {
                    _TB = m_DrwSketch.TextBoxes.AddByRectangle(m_TG.CreatePoint2d(-10, 0.05), m_TG.CreatePoint2d(-0.1, 1.5), setting.textBox2.Text);
                    _TB.HorizontalJustification = HorizontalTextAlignmentEnum.kAlignTextRight;
                    _TB.VerticalJustification = VerticalTextAlignmentEnum.kAlignTextLower;
                }
                m_SketchDef.ExitEdit(true);
                (m_Drw.ActiveSheet.SketchedSymbols.Add(m_SketchDef, insertPt)).Static = true;
                invApp.ScreenUpdating = true;
            }
            catch (Exception e)
            {
                MessageBox.Show(e.ToString());
                invApp.ScreenUpdating = true;
            }
        }

        private void formatText(ref Inventor.TextBox _TB, string align, double widthCol, double textH = 0.35, string font = "GOST Common", 
            double width = 1, bool underline = false, bool italic = true, bool bold = false)
        {
            string format = ""; string txt = "";

            txt = _TB.Text;
            _TB.WidthScale = width;
            _TB.VerticalJustification = VerticalTextAlignmentEnum.kAlignTextMiddle; 
            switch (align)
            {
                case "По центру":
                    _TB.HorizontalJustification = HorizontalTextAlignmentEnum.kAlignTextCenter;
                    break;
                case "Слева":
                    _TB.HorizontalJustification = HorizontalTextAlignmentEnum.kAlignTextLeft;
                    break;
                case "Справа":
                    _TB.HorizontalJustification = HorizontalTextAlignmentEnum.kAlignTextRight;
                    break;
            }
            if (underline)
            {
                format += " Underline ='True' ";
            }
            if (!italic)
            {
                format += " Italic = 'False' ";
            }
            if (bold)
            {
                format += " Bold = 'True' ";
            }
            if (font != "GOST Common")
            {
                format += @" Font = '" + font.ToString() + @"' ";
            }

            if (Math.Round(textH*100) != 35)
            {
                format += @" FontSize = '" + textH.ToString() + @"' ";
            }

            double w = widthCol / _TB.FittedTextWidth * (width - 0.05);

            if (setting.checkBox3.Checked && _TB.FittedTextWidth > widthCol)
            {
                _TB.SingleLineText = true;
                _TB.WidthScale = widthCol / _TB.FittedTextWidth * (width - 0.05);
            }

            if (_TB.Text.IndexOf('/') != -1)
            {
                string sub1, sub2;
                int ind = txt.IndexOf('/');
                sub1 = txt.Substring(0, ind);
                sub2 = txt.Substring(ind, txt.Length - ind);
                if (sub1.LastIndexOf(' ') != -1)
                    sub1 = sub1.Insert(sub1.LastIndexOf(' ')+1, "<Stack FractionalTextScale='1'>");
                if (sub2.IndexOf(' ') != -1)
                    sub2 = sub2.Insert(sub2.IndexOf(' '), "</Stack>");
                txt = sub1 + sub2;
                _TB.FormattedText =  txt;
            }

            if (_TB.Text.IndexOf('^') != -1)
            {
                string sub1, sub2;
                int ind = txt.IndexOf('^');
                sub1 = txt.Substring(0, ind);
                sub2 = txt.Substring(ind, txt.Length - ind);
                //if (sub1.LastIndexOf(' ') != -1)
                  //  sub1 = sub1.Insert(sub1.LastIndexOf(' ') + 1, "<Stack FractionalTextScale='0.7'>");
                if (sub2.IndexOf(' ') != -1)
                    sub2 = sub2.Insert(sub2.IndexOf(' '), "</Stack>");
                else sub2 += "</Stack>";
                txt = sub1 + "<Stack>" + sub2;
                _TB.FormattedText = txt;
            }

            if (format != "")
              _TB.FormattedText = @"<StyleOverride" + format + @">" + txt + @"</StyleOverride>";
        }

        private Inventor.TextBox text(DataGridView _dgv, DataGridViewCell cell, ref List<System.Drawing.Point> lstTxt, ref double[] arrColumns, ref double[] arrRow)
        { 
            string val = cell.Value.ToString();
            int endX = -1; int endY = -1;
            Inventor.TextBox TB;
            for (int j = cell.RowIndex; j < _dgv.RowCount; j++)
            {
                for (int i = cell.ColumnIndex; i < _dgv.ColumnCount; i++)
                {
                    if (i < (_dgv.ColumnCount -1) && _dgv[i+1,j].Value != null && val == _dgv[i+1, j].Value.ToString())
                    {
                        continue;
                    }
                    else
                    {
                        endX = i;
                        break;
                    }
                 }
                if (j < (_dgv.RowCount -1) && _dgv[endX, j+1].Value != null && val == _dgv[endX, j+1].Value.ToString())
                {
                    continue;
                }
                else
                {
                    endY = j;
                    break;
                }
             }
            if (cell.RowIndex != endX || cell.ColumnIndex != endY)
            {
                lstTxt.Add(new System.Drawing.Point(cell.ColumnIndex,cell.RowIndex));
                lstTxt.Add(new System.Drawing.Point(endX, endY));
            }

            Func<string, double> conv = s =>
            {
                char rep = (separator == ',') ? '.' : ',';
                return (s.IndexOf(separator) != -1) ? Convert.ToDouble(s)/10 : Convert.ToDouble(s.Replace(rep, separator))/10;
            };

            double left = conv(setting.comboBox3.Text);
            double rigth = conv(setting.comboBox4.Text);
            double upper = conv(setting.comboBox5.Text);
            double bottom = conv(setting.comboBox6.Text);

            Point2d pt1 = m_TG.CreatePoint2d(-sumArr(arrColumns, cell.ColumnIndex, arrColumns.Count()-1)+left, -sumArr(arrRow, 0 , endY)+ bottom);
            Point2d pt2 = m_TG.CreatePoint2d(-sumArr(arrColumns, endX + 1, arrColumns.Count() - 1) - rigth, -sumArr(arrRow, 0, cell.RowIndex) + (arrRow[cell.RowIndex]) - upper); 
            TB = m_DrwSketch.TextBoxes.AddByRectangle(pt1, pt2, val);
            //TB.Style.FontSize = Convert.ToDouble(setting.comboBox1.SelectedItem.ToString().Replace('.',separator));

            return TB;
        }

        private List<Inventor.Point2d> LinesX(ref List<System.Drawing.Point> lst, double[] arr, int i, Inventor.Point2d pt)
        {
            List<Inventor.Point2d> tmpPts = new List<Point2d>();
            Inventor.Point2d startPt; bool flag = false;
            startPt = invApp.TransientGeometry.CreatePoint2d(pt.X, pt.Y);
            //tmpPts.Add(pt);
            for (int j=0; j<lst.Count; j+=2)
            {
                //if (i > lst[j].X && i <= lst[j + 1].X)
                //{
                //    startPt.Y -= arr[lst[j].Y];
                //    indx = j;
                //    flag = true;
                //}
                //else if (flag)
                //{
                //    tmpPts.Add(invApp.TransientGeometry.CreatePoint2d(startPt.X, startPt.Y - sumArr(arr,indx,j)));
                //    flag = false;
                //}
                if (i > lst[j].X && i <= lst[j + 1].X)
                {
                    if (!flag)
                    {
                        Inventor.Point2d invPt = invApp.TransientGeometry.CreatePoint2d(startPt.X, startPt.Y - sumArr(arr, lst[j].Y - 1));
                        tmpPts.Add(invPt);
                        if (tmpPts.Count > 1 && invPt.Y > tmpPts[tmpPts.Count - 2].Y)
                            tmpPts.RemoveAt(tmpPts.Count - 1);
                        flag = true;
                    }
                    if (flag)
                    {
                        Inventor.Point2d invPt = invApp.TransientGeometry.CreatePoint2d(startPt.X, startPt.Y - sumArr(arr, lst[j].Y));
                        tmpPts.Add(invPt);
                        if (tmpPts.Count > 1 && invPt.Y > tmpPts[tmpPts.Count - 2].Y)
                            tmpPts.RemoveAt(tmpPts.Count - 1);
                        flag = false;
                        //tmpPts.Add(invApp.TransientGeometry.CreatePoint2d(startPt.X, startPt.Y - sumArr(arr, lst[j].Y)));
                    }
                    
                }
            }
            if (tmpPts.Count > 0)
            {
                if (tmpPts[0].Y != pt.Y)
                {
                    tmpPts.Insert(0, pt);
                }
                else tmpPts.RemoveAt(0);
            }
            
                //tmpPts.Insert(0, pt);
            //tmpPts.RemoveRange(tmpPts.Count - 1, 1);
//             if (tmpPts.Count > 1)
//                 return tmpPts;
            if (tmpPts.Count == 0)
            {
                return null;                
            }
            else if ((tmpPts.Count)%2 == 1)
            {
                tmpPts.Add(invApp.TransientGeometry.CreatePoint2d(startPt.X, startPt.Y - sumArr(arr,arr.Count()-2)));
                //tmpPts.Sort(delegate(Inventor.Point2d pt1, Inventor.Point2d pt2) { if (pt1.Y <= pt2.Y) return 1; else return -1; });
                removeDub(ref tmpPts);
                return tmpPts;
            }
                else if((tmpPts.Count)%2 == 0)
            {
                //tmpPts.Sort(delegate(Inventor.Point2d pt1, Inventor.Point2d pt2) { if (pt1.Y <= pt2.Y) return 1; else return -1; });
                removeDub(ref tmpPts);
                return tmpPts;
            }
            return null;
        }

        //private void removeDubLine(ref List<Inventor.Point2d> lstPt)
        //{
        //   for (int i = 0; i < lstPt.Count-1;)
        //       {

        //       }
        //}

        private void removeDub (ref List<Inventor.Point2d> lstPt)
        {
            int count = 2, start;
           for (int i = 0; i < lstPt.Count-1;)
           {
               if (lstPt[i].X == lstPt[i + 1].X && lstPt[i].Y == lstPt[i + 1].Y)
               {
                   start = i;
                   i++;
//                    if (i < (lstPt.Count-1) &&(lstPt[i].X == lstPt[i + 2].X && lstPt[i].Y == lstPt[i + 2].Y))
//                    {
//                        i++;
//                        count++;
//                    }
                   //if (i != (lstPt.Count - 1))
                       lstPt.RemoveRange(start, count);
                   //else lstPt.RemoveRange(start, count - 1);
                   i -= count;
                   count = 2;
               }
               i++;
           }
        }

        private List<Inventor.Point2d> LinesY(ref List<System.Drawing.Point> lst, double[] arr, int i, Inventor.Point2d pt)
        {
            List<Inventor.Point2d> tmpPts = new List<Point2d>();
            Inventor.Point2d startPt; bool flag = false;
            startPt = invApp.TransientGeometry.CreatePoint2d(pt.X, pt.Y);
            //int indx = 0;
            //tmpPts.Add(pt);
            for (int j = 0; j < lst.Count; j += 2)
            {
                //if (i > lst[j].X && i <= lst[j + 1].X)
                //{
                //    startPt.Y -= arr[lst[j].Y];
                //    indx = j;
                //    flag = true;
                //}
                //else if (flag)
                //{
                //    tmpPts.Add(invApp.TransientGeometry.CreatePoint2d(startPt.X, startPt.Y - sumArr(arr,indx,j)));
                //    flag = false;
                //}
                if (i > lst[j].Y && i <= lst[j + 1].Y)
                {
                    if (!flag)
                    { tmpPts.Add(invApp.TransientGeometry.CreatePoint2d(startPt.X + sumArr(arr, lst[j].X-1), startPt.Y)); flag = true; }
                    if (flag)
                    {
                        tmpPts.Add(invApp.TransientGeometry.CreatePoint2d(startPt.X + sumArr(arr, lst[j].X), startPt.Y));
                        flag = false;
                    }
                }
            }

            if (tmpPts.Count > 0)
            {
                if (tmpPts[0].X != pt.X)
                {
                    tmpPts.Insert(0, pt);
                }
                else tmpPts.RemoveAt(0);
            }


            if (tmpPts.Count == 0)
            {
                return null;
            }
            else if ((tmpPts.Count) % 2 == 1)
            {
                tmpPts.Add(invApp.TransientGeometry.CreatePoint2d(startPt.X + sumArr(arr, arr.Count() - 1), startPt.Y));
                removeDub(ref tmpPts);
                return tmpPts;
            }
            else if ((tmpPts.Count) % 2 == 0)
            {
                removeDub(ref tmpPts);
                return tmpPts;
            }
            return null;
        }

        private double sumArr(double[] arr, int end)
        {
            double sum = 0;
            for (int i=0; i <= end; i++)
            {
               sum += arr[i];
            }
            return sum;
        }

        private double sumArr(double[] arr,int start, int end)
        {
            double sum = 0;
            for (int i = start; i <= end; i++)
            {
                sum += arr[i];
            }
            return sum;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (this.textBox2.Text != "")
            {
                int[] intarr = new int[] { 0 }; bool col = false, flag = false; System.Drawing.Point start = new System.Drawing.Point(0,0), 
                    end = new System.Drawing.Point(0,0);
                if (setting.textBox1.Text != "")
                {
                    string[] strArr = setting.textBox1.Text.Split(';');
                    intarr = new int[strArr.Count()];
                    for (int i = 0; i < strArr.Count()-1; i++)
                    {
                        intarr[i+1] = Convert.ToInt32(strArr[i].ToString());
                    } 
                }
                    
                if (setting.checkBox1.Checked) col = true;
                List<System.Drawing.Point> lstX = new List<System.Drawing.Point>();
                List<System.Drawing.Point> lstY = new List<System.Drawing.Point>();
                List<string> lstcells = new List<string>();
                //for (int j = 0; j < this._dgv.ColumnCount; j++)
                //{
                    for (int i = 0; i < this._dgv.RowCount-1; i++)
                    {
                        foreach (DataGridViewCell cell in _dgv.Rows[i].Cells)
                        {
                            if (cell.Value != null)
                                lstcells.Add(cell.Value.ToString());
                            else lstcells.Add("");
                        }
                        lstcells.Add("end");
                        findInRow(lstcells, ref lstX, i);
                        lstcells.Clear();
                    }

                    for (int i = 0; i < this._dgv.ColumnCount; i++)
                    {
                        for (int k = 0; k < _dgv.RowCount; k++ )
                        {
                            if (_dgv[i,k].Value != null)
                                lstcells.Add(_dgv[i, k].Value.ToString());
                            else lstcells.Add("");
                        }
                        lstcells.Add("end");
                        findInColumn(lstcells, ref lstY, i);
                        lstcells.Clear();
                    }

                //}
                //lst.Add(start); lst.Add(end);
                    AttribAdd(this._dgv, this.textBox2.Text);
                    Inventor.Point2d pt = m_TG.CreatePoint2d();
                    if (comboBox1.Text == "Исполнения" || comboBox1.Text == "Исполнения_AC")
                    {
                        pt.X = m_Drw.ActiveSheet.Width - 18.5 - 1;
                        pt.Y = inputArrSum(this.RowDimensions.Text) + 0.5;
                    }
                    else if (comboBox1.Text == "Гибка")
                    {
                        pt.X = m_Drw.ActiveSheet.Width - 0.5;
                        pt.Y = inputArrSum(this.RowDimensions.Text) + 2.5;
                    }
                    drawTable(this.textBox2.Text, m_TG.CreatePoint2d(), inputArr(this.ColumnDimensions.Text),
                        inputArr(this.RowDimensions.Text),lstX,lstY, intArr: intarr, _dgv: this._dgv, col: col,insertPt: pt);
            }
        }

        private void findInRow(List<string> lstcells, ref List<System.Drawing.Point> lst, int y)
        {
            System.Drawing.Point start = new System.Drawing.Point(-1, -1),
                    end = new System.Drawing.Point(0, 0); bool flag = true;
            //List<string> lstcells = new List<string>();

            for (int i = 0; i < lstcells.Count-1; )
            {
                if (lstcells[i] != "")
                {
                    if (lstcells[i] == lstcells[i + 1])
                    {
                        if (flag)
                        { start.X = i; start.Y = y; }
                        flag = false;
                    }
                    else if (start.X != -1)
                    {
                        end.X = i; end.Y = y;
                        flag = true;
                        lst.Add(start);
                        lst.Add(end);
                        start.X = -1; start.Y = -1;
                    }
                }
                i++;
            }
        }

        private void findInColumn(List<string> lstcells, ref List<System.Drawing.Point> lst, int y)
        {
            System.Drawing.Point start = new System.Drawing.Point(-1, -1),
                    end = new System.Drawing.Point(0, 0); bool flag = true;
            //List<string> lstcells = new List<string>();

            for (int i = 0; i < lstcells.Count - 1; )
            {
                if (lstcells[i] != "")
                {
                    if (lstcells[i] == lstcells[i + 1])
                    {
                        if (flag)
                        { start.X = y; start.Y = i; }
                        flag = false;
                    }
                    else if (start.X != -1)
                    {
                        end.X = y; end.Y = i;
                        flag = true;
                        lst.Add(start);
                        lst.Add(end);
                        start.X = -1; start.Y = -1;
                    }
                }
                i++;
            }
        }

        private double inputArrSum(string txt)
        {
            spl = txt.Split(';');
            return spl.Where(s => s != "").Sum(e => double.Parse(e))/10;
        }

        private double[] inputArr(string txt)
        {
            spl = txt.Split(';');
            double[] tmp = new double[spl.Count()-1];
            for(int i=0; i<spl.Count()-1;i++)
            {
               tmp[i] = Convert.ToDouble(spl[i])/10;
            }
            return tmp;
        }

        private void _dgv_RowsAdded(object sender, DataGridViewRowsAddedEventArgs e)
        {
            if (this._dgv.RowCount == Convert.ToInt16(this.CountRow.Text)+1)
            {
                cr = Convert.ToInt16(this.CountRow.Text);
                this.CountRow.Text = this._dgv.RowCount.ToString();

                val = "";

                if (this.RowDimensions.Text == "")
                {
                    for (int i = 0; i < this._dgv.RowCount; i++)
                    {
                        val += setting.comboBox8.Text + ';';
                    }
                }
                else val = this.RowDimensions.Text;


                if (val != "") spl = val.Split(';');

                if (this._dgv.RowCount == 0)
                {
                    this.RowDimensions.Text = "";
                }
                else if (this._dgv.RowCount > cr)
                {
                    for (int i = cr; i < this._dgv.Rows.Count; i++)
                    {
                        val += setting.comboBox8.Text + ';';
                    }
                }
                //else if (this._dgv.RowCount > cr)
                //{
                //    val = "";
                //    for (int i = 0; i < cr; i++)
                //    {
                //        val += spl[i] + ';';
                //    }
                //}
                this.RowDimensions.Text = val;
           }
        }

        private void AttribAdd(DataGridView _dgv, string name)
        {
            //Inventor.AttributeSets attSets;
            AttribDelete(name);
            AttribDelete("height" + name);
            AttribDelete("width" + name);
            AttribDelete("Name" + name);
            AttribDelete("TextHeigth" + name);
            AttribDelete("Padding" + name);
            AttribDelete("Style" + name);
            AttribDelete("MyParam" + name);
            Inventor.AttributeSet attSet;
            Inventor.Attribute attr;
            string rowName; int i = 0;// string cellName; int j = 0;
            string val = "", stl = "";

            foreach (DataGridViewRow row in _dgv.Rows)
            {
                rowName = name + "Row" + i;
                i += 1;
                attSet = m_Drw.AttributeSets.Add(rowName);
                if (i == _dgv.RowCount)
                    break;
                foreach (DataGridViewCell cel in row.Cells)
                {
                    val = val + (string)cel.FormattedValue + ';';
                    if (cel.Style.Font != null)   
                    {
                        stl += cel.ColumnIndex.ToString() + ';' + cel.RowIndex.ToString() + ';'
                            + Convert.ToBoolean(cel.Style.Font.Italic.ToString()) + ';'
                            + Convert.ToBoolean(cel.Style.Font.Bold.ToString()) + ';'
                            + Convert.ToBoolean(cel.Style.Font.Underline.ToString()) + ';'
                            + cel.Style.Font.Size + ';'
                            + cel.Style.Font.Name + ';'
                            + (int)cel.Style.Alignment + ';';
                    }
                    if (cel.Style.Alignment != DataGridViewContentAlignment.NotSet)
                    {
                        stl += cel.ColumnIndex.ToString() + ';' + cel.RowIndex.ToString() + ';'
                            + Convert.ToBoolean(_dgv.DefaultCellStyle.Font.Italic.ToString()) + ';'
                            + Convert.ToBoolean(_dgv.DefaultCellStyle.Font.Bold.ToString()) + ';'
                            + Convert.ToBoolean(_dgv.DefaultCellStyle.Font.Underline.ToString()) + ';'
                            + "" + ';'
                            + "" + ';'
                            + (int)cel.Style.Alignment + ';';
                    }
                }
                attr = attSet.Add("String", ValueTypeEnum.kStringType, val);

                val = "";
            }
            string txt = setting.textBox1.Text.Replace(';', '`');
            string pad = setting.comboBox3.Text + ';' + setting.comboBox4.Text + ';' + setting.comboBox5.Text + ';' + setting.comboBox6.Text + ';' +
                setting.comboBox8.Text + ';' + setting.comboBox9.Text + ';' + 
                setting.checkBox1.Checked.ToString() + ';' + setting.checkBox2.Checked.ToString() + ';' + setting.checkBox3.Checked.ToString() + ';' +
                setting.checkBox4.Checked.ToString() + ';' + setting.checkBox5.Checked.ToString() + ';' + txt + ';' + setting.textBox2.Text + ';' + setting.textBox3.Text;
                attSet = m_Drw.AttributeSets.Add("height" + name);
                attSet.Add("Spl", ValueTypeEnum.kStringType, this.RowDimensions.Text);
                attSet = m_Drw.AttributeSets.Add("width" + name);
                attSet.Add("Spl", ValueTypeEnum.kStringType, this.ColumnDimensions.Text);
                attSet = m_Drw.AttributeSets.Add("Name" + name);
                attSet.Add("Spl", ValueTypeEnum.kStringType, this.textBox2.Text);
                attSet = m_Drw.AttributeSets.Add("TextHeigth" + name);
                attSet.Add("Spl", ValueTypeEnum.kStringType, setting.comboBox1.Text);
                attSet = m_Drw.AttributeSets.Add("Padding" + name);
                attSet.Add("Spl", ValueTypeEnum.kStringType, pad);
            if (stl != "")
            {
                attSet = m_Drw.AttributeSets.Add("Style" + name);
                attSet.Add("Spl", ValueTypeEnum.kStringType, stl);
            }
            if (styleParam.Count != 0)
            {
                string tmp = "";
                foreach (string str in styleParam)
                {
                    tmp = str.Replace(';', '`');
                    val += tmp + ";";
                }
                attSet = m_Drw.AttributeSets.Add("MyParam" + name);
                attSet.Add("Spl", ValueTypeEnum.kStringType, val);
                val = "";
            }
                //j = 0;
        }

        private void удалитьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //Object pr;
            try
            {
                AttribDelete(this.textBox2.Text);
                AttribDelete("height" + this.textBox2.Text);
                AttribDelete("width" + this.textBox2.Text);
                AttribDelete("Name" + this.textBox2.Text);
                AttribDelete("TextHeigth" + this.textBox2.Text);
                AttribDelete("Padding" + this.textBox2.Text);
                AttribDelete("Style" + this.textBox2.Text);
                AttribDelete("MyParam" + this.textBox2.Text);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
            //m_Drw.AttributeManager.PurgeAttributeSets("*", false, out pr);
        }

        private void AttribDelete(string name)
        {
            //Inventor.Attribute attr;
            try
            {
                foreach (Inventor.AttributeSet attSet in m_Drw.AttributeSets)
                {
                    if (attSet.Name.StartsWith(name))
                    {

                        foreach (Inventor.Attribute attr in attSet)
                        {
                            attr.Delete();
                        }
                        attSet.Delete();
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
            //for (int i = 1; i <= m_AsmDoc.AttributeSets[name].Count;i++ )
            //{
            //    m_AsmDoc.AttributeSets[name][i].Delete();
            //}
        }

        private void сохранитьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            DrawingDocument m_DrwOld = m_Drw;
            m_Drw = (DrawingDocument)invApp.ActiveDocument;
            AttribAdd(this._dgv, this.textBox2.Text);
            m_Drw = m_DrwOld;
        }

        private void open(DataGridView _dgv)
        {
            int i = 0; bool flag = true; DataGridViewContentAlignment align;
            try
            {
                //while (_dgv.Rows.Count != 0)
                //    _dgv.Rows.Remove(_dgv.Rows[_dgv.Rows.Count - 1]);
                int oldCountRow = _dgv.RowCount;
                _dgv.RowCount = 1;
                _dgv.Rows.Clear();

                foreach (Inventor.AttributeSet attrset in m_Drw.AttributeSets)
                {
                    if (attrset.Name.StartsWith(this.textBox2.Text) && attrset.Count != 0)
                    {
                        i += 1;
                        if (i == m_Drw.AttributeSets.Count)
                            break;
                        string[] str = attrset[1].Value.ToString().Split(';');
                        if (flag) { flag = false; _dgv.ColumnCount = str.Count() - 1; this.CountColumn.Text = (str.Count() - 1).ToString(); }
                        //Array.Resize(ref str, str.Length - 1);
                        _dgv.Rows.Add(str);
                    }
                    if (attrset.Name.StartsWith("Name" + this.textBox2.Text) && attrset.Count != 0)
                    {
                        this.textBox2.Text = attrset[1].Value.ToString();
                    }
                    if (attrset.Name.StartsWith("height" + this.textBox2.Text) && attrset.Count != 0)
                    {
                        this.RowDimensions.Text = attrset[1].Value.ToString();
                    }
                    if (attrset.Name.StartsWith("width" + this.textBox2.Text) && attrset.Count != 0)
                    {
                        this.ColumnDimensions.Text = attrset[1].Value.ToString();
                    }
                    if (attrset.Name.StartsWith("TextHeigth" + this.textBox2.Text) && attrset.Count != 0)
                    {
                        setting.comboBox1.Text = attrset[1].Value.ToString();
                    }
                    if (attrset.Name.StartsWith("Padding" + this.textBox2.Text) && attrset.Count != 0)
                    {
                        string [] strs =  attrset[1].Value.ToString().Split(';');
                        setting.comboBox3.Text = strs[0];
                        setting.comboBox4.Text = strs[1];
                        setting.comboBox5.Text = strs[2];
                        setting.comboBox6.Text = strs[3];
                        setting.comboBox8.Text = strs[4];
                        setting.comboBox9.Text = strs[5];
                        setting.checkBox1.Checked = Convert.ToBoolean(strs[6]);
                        setting.checkBox2.Checked = Convert.ToBoolean(strs[7]);
                        setting.checkBox3.Checked = Convert.ToBoolean(strs[8]);
                        setting.checkBox4.Checked = Convert.ToBoolean(strs[9]);
                        setting.checkBox5.Checked = Convert.ToBoolean(strs[10]);
                        setting.textBox1.Text = strs[11].Replace('`', ';');
                        if (strs.Count() == 14)
                        {
                            setting.textBox2.Text = strs[12];
                            setting.textBox3.Text = strs[13];
                        }
                    }
                    if (attrset.Name.StartsWith("Style" + this.textBox2.Text) && attrset.Count != 0)
                    {
                        string[] stl = attrset[1].Value.ToString().Split(';');
                        for (int st = 0; st < stl.Count()-1; st += 8)
                        {
                            bool italic = false, underline = false, bold = false; float textH = 12; string nameFont = "Gost Common";
                            int row = Convert.ToInt16(stl[st + 1]), col = Convert.ToInt16(stl[st]);
                            if (stl[st + 2] == "True") italic = true;
                            if (stl[st + 3] == "True") bold = true;
                            if (stl[st + 4] == "True") underline = true;                
                            if (stl[st + 5] != "") textH = Convert.ToInt16(stl[st+5]);
                            if (stl[st + 6] != "") nameFont = stl[st + 6];
                            if (stl[st + 7] != "")
                            {
                                align = (DataGridViewContentAlignment)Convert.ToInt16(stl[st + 7]);
                                _dgv[col, row].Style.Alignment = align;
                            }
                            _dgv[col,row].Style.Font = fontStyle(underline, italic, bold, textH: textH, nameFont: nameFont);
                        }
                    }
                    if (attrset.Name.StartsWith("MyParam" + this.textBox2.Text) && attrset.Count != 0)
                    {
                        string[] stl = attrset[1].Value.ToString().Split(';'); 
                        for (int st = 0; st < stl.Count()-1; st++)
                        {
                            string tmp = stl[st].Replace('`', ';');
                            string tmp2 = styleParam.Find(delegate(string str) { return str == tmp; });
                            if (tmp2 == null) styleParam.Add(tmp);
                        }
                    }
                }
                //_dgv.Rows.RemoveAt(_dgv.Rows.Count-1);
                //string val = "";
                //if (oldCountRow )
                this.CountRow.Text = _dgv.RowCount.ToString();
                EventArgs e = null; Object obj = null;
                ColumnDimensions_Leave(obj, e);
                RowDimensions_Leave(obj, e);
            }
            catch (Exception e)
            {
                MessageBox.Show(e.ToString());
            }
        }

        private void открытьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            open(this._dgv);
        }

        private void _dgv_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Delete && ((DataGridView)sender).SelectedCells.Count != 0)
            {
                foreach (DataGridViewCell c in ((DataGridView)sender).SelectedCells)
                {
                    c.Value = "";
                }
            }
            if (e.Control && e.KeyCode == Keys.S)
            {
                AttribAdd(this._dgv, this.textBox2.Text);
            }
            if (e.Control && e.KeyCode == Keys.C)
            {
                if (((DataGridView)sender).SelectedCells.Count == 1)
                value = ((DataGridView)sender).CurrentCell.Value.ToString();
                else if (((DataGridView)sender).SelectedCells[0].RowIndex == ((DataGridView)sender).SelectedCells[((DataGridView)sender).SelectedCells.Count-1].RowIndex)
                {
                    value = "";
                    for (int i = 0; i < ((DataGridView)sender).SelectedCells.Count; i++)
                    {
                        value += ((DataGridView)sender).SelectedCells[i].Value + ";";
                    }
                }
            }
            if (e.KeyCode == Keys.V && ((DataGridView)sender).SelectedCells.Count != 0)
            {
                if (value.IndexOf(';') == -1)
                {
                    foreach (DataGridViewCell c in ((DataGridView)sender).SelectedCells)
                    {
                        c.Value = value;
                    }
                }
                else
                {
                    string[] spl = value.Split(';');
                    int j = 0;

                    for (int i = ((DataGridView)sender).SelectedCells.Count-1; i >=0 ; i--)
                    {
                        ((DataGridView)sender).SelectedCells[i].Value = spl[j];
                            j++;
                            if (j == (spl.Count() - 1)) j = 0;
                    }
                }
            }
        }

        private void настройкиToolStripMenuItem_Click(object sender, EventArgs e)
        {
            setting.Show();
        }

        void toolStripMenuItem2_TextChanged(object sender, System.EventArgs e)
        {
            ToolStripComboBox tscb = (ToolStripComboBox)sender;
            if (tscb.SelectedItem.ToString() != setting.comboBox1.Text)
            {
                foreach (DataGridViewCell c in this._dgv.SelectedCells)
                {
                    c.Style.Font = new Font(c.Style.Font, FontStyle.Bold);
                }
            }
        }

        private void поЦентруToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (this._dgv.SelectedCells.Count != 0)
            {
                foreach (DataGridViewCell c in this._dgv.SelectedCells)
                {
                    c.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                }
            }
        }

        private void справаToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (this._dgv.SelectedCells.Count != 0)
            {
                foreach (DataGridViewCell c in this._dgv.SelectedCells)
                {
                    c.Style.Alignment = DataGridViewContentAlignment.MiddleRight;
                }
            }
        }

        private void слеваToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (this._dgv.SelectedCells.Count != 0)
            {
                foreach (DataGridViewCell c in this._dgv.SelectedCells)
                {
                    c.Style.Alignment = DataGridViewContentAlignment.MiddleLeft;
                }
            }
        }

        private void даToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                if (this._dgv.SelectedCells.Count != 0)
                {
                    foreach (DataGridViewCell c in this._dgv.SelectedCells)
                    {
                        if (c.Style.Font != null)
                            c.Style.Font = fontStyle( underline: c.Style.Font.Underline, bold: c.Style.Font.Bold);
                        else c.Style.Font = fontStyle();
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private void даToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            if (this._dgv.SelectedCells.Count != 0)
            {
                foreach (DataGridViewCell c in this._dgv.SelectedCells)
                {
                    if (c.Style.Font != null)
                    c.Style.Font = fontStyle(italic: c.Style.Font.Italic, underline: true, bold: c.Style.Font.Bold);
                    else c.Style.Font = fontStyle();
                }
            }
        }

        private void даToolStripMenuItem2_Click(object sender, EventArgs e)
        {
            if (this._dgv.SelectedCells.Count != 0)
            {
                foreach (DataGridViewCell c in this._dgv.SelectedCells)
                {
                    {
                        if (c.Style.Font != null)
                        c.Style.Font = fontStyle(italic: c.Style.Font.Italic, underline: c.Style.Font.Underline, bold: true);
                        else c.Style.Font = fontStyle();
                    }
                }
            }
        }

        private void нетToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            if (this._dgv.SelectedCells.Count != 0)
            {
                foreach (DataGridViewCell c in this._dgv.SelectedCells)
                {
                    if (c.Style.Font != null)
                    c.Style.Font = fontStyle(italic: true, underline: c.Style.Font.Underline, bold: c.Style.Font.Bold);
                    else c.Style.Font = fontStyle();
                }
            }
        }

        private void нетToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (this._dgv.SelectedCells.Count != 0)
            {
                foreach (DataGridViewCell c in this._dgv.SelectedCells)
                {
                    try
                    {
                        if (c.Style.Font != null)
                            c.Style.Font = fontStyle(italic: false, underline: c.Style.Font.Underline, bold: c.Style.Font.Bold);
                        else c.Style.Font = fontStyle();
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.ToString());
                    }
                }
            }
        }

        private Font fontStyle(bool underline = false, bool italic = true, bool bold = false, float textH = 12, string nameFont = "Gost Common")
        {
            if (underline && italic) return new Font(nameFont, textH, FontStyle.Underline | FontStyle.Italic);
            if (underline && bold) return new Font(nameFont, textH, FontStyle.Underline | FontStyle.Bold);
            if (bold && italic) return new Font(nameFont, textH, FontStyle.Bold | FontStyle.Italic);
            if (underline && bold && italic) return new Font(nameFont, textH, FontStyle.Underline | FontStyle.Bold | FontStyle.Italic);
            if (italic) return new Font(nameFont, textH, FontStyle.Italic);
            if (underline) return new Font(nameFont, textH, FontStyle.Underline);
            if (bold) return new Font(nameFont, textH, FontStyle.Bold);
            if (!italic) return new Font(nameFont, textH, FontStyle.Regular);
            return null;
        }

        private void Tables_FormClosed(object sender, FormClosedEventArgs e)
        {
            InvAddIn.TablesButton.nameTable = this.textBox2.Text;
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            ComboBox cb = (ComboBox)sender;
            if (cb.Text != "")
            {
                LoadTable(cb.Text);
            }
        }

        public static bool eq(LinearGeneralDimension dim, DrawingCurve dc)
        {
            DrawingCurve i1 = dim.IntentOne.Geometry as DrawingCurve, i2 = dim.IntentTwo.Geometry as DrawingCurve;
            return (i1.Equals(dc) || i2.Equals(dc)) ? true : false;
        }

        public static bool eq(Edge edgeModel, Edge dc)
        {
            bool start, end;
            start = InvDoc.u.comparePoint(edgeModel.StartVertex, dc.StartVertex, 1000);
            if (!start) start = InvDoc.u.comparePoint(edgeModel.StopVertex, dc.StartVertex, 1000);
            end = InvDoc.u.comparePoint(edgeModel.StopVertex, dc.StopVertex, 1000);
            if (!end) end = InvDoc.u.comparePoint(edgeModel.StartVertex, dc.StopVertex, 1000);
            return (start && end);
        }

        public static bool eq(Edge edgeModel, DrawingCurve dc)
        {
            bool start, end;
            start = InvDoc.u.comparePoint2d(edgeModel.StartVertex, ((Edge)dc.ModelGeometry).StartVertex, 1000);
            if (!start) start = InvDoc.u.comparePoint2d(edgeModel.StopVertex, ((Edge)dc.ModelGeometry).StartVertex, 1000);
            end = InvDoc.u.comparePoint2d(edgeModel.StopVertex, ((Edge)dc.ModelGeometry).StopVertex, 1000);
            if (!end) end = InvDoc.u.comparePoint2d(edgeModel.StartVertex, ((Edge)dc.ModelGeometry).StopVertex, 1000);
            return (start && end);
        }

        static public string findDim(DrawingCurve dc)
        {
            DrawingDocument m_Drw = Macros.StandardAddInServer.m_inventorApplication.ActiveDocument as DrawingDocument;
            try
            {
                char separator = System.Globalization.CultureInfo.CurrentCulture.NumberFormat.CurrencyDecimalSeparator[0];
                foreach (DrawingDimension dim in m_Drw.ActiveSheet.DrawingDimensions)
                {
                    if (dim.Type == ObjectTypeEnum.kLinearGeneralDimensionObject)
                    {
                        LinearGeneralDimension ldim = (LinearGeneralDimension)dim;

                        if (/*eq(ldim, dc)*/ldim.IntentOne.Geometry.Equals(dc) || ldim.IntentTwo.Geometry.Equals(dc))
                        {
                            string str = ldim.Text.Text;
                            if (str.IndexOf('*') != -1) str = str.Remove(str.IndexOf('*'));

                            if ((((DrawingCurve)ldim.IntentOne.Geometry).EdgeType == DrawingEdgeTypeEnum.kBendUpEdge ||
                                ((DrawingCurve)ldim.IntentOne.Geometry).EdgeType == DrawingEdgeTypeEnum.kBendDownEdge) &&
                                (((DrawingCurve)ldim.IntentTwo.Geometry).EdgeType == DrawingEdgeTypeEnum.kBendUpEdge ||
                                ((DrawingCurve)ldim.IntentTwo.Geometry).EdgeType == DrawingEdgeTypeEnum.kBendDownEdge)
                                )
                            {
                                Func<string, double> conv = s =>
                                {
                                    char rep = (separator == ',') ? '.' : ',';
                                    return (s.IndexOf(separator) != -1) ? Convert.ToDouble(s) : Convert.ToDouble(s.Replace(rep, separator));
                                };

                                double val = conv(str) + 0.8;
                                return "Упор=" + val.ToString() + " мм";
                            }

                            return "Упор=" + str + " мм";
                        }
                        foreach (DrawingCurveSegment dcs in dc.Segments)
                        {
                            if (ldim.IntentOne.Geometry.Equals(dcs) || ldim.IntentTwo.Geometry.Equals(dcs))
                            {
                                string str = ldim.Text.Text;
                                if (str.IndexOf('*') != -1) str = str.Remove(str.IndexOf('*'));
                                return "Упор=" + str + " мм";
                            }
                        }
                    }
                }
            }
            catch (Exception)
            {
                return "";
            }
            return "";
        }

        static public void LoadBend(ref List<string> lst, List<DrawingCurve> lstDC)
        {
            string val = "", upor = "", direct = ""; int j = 0; double dval, ang, r, l; BendOrderSourceTypeEnum bendOrder;
            SheetMetalComponentDefinition smcd;
            DrawingDocument m_Drw = Macros.StandardAddInServer.m_inventorApplication.ActiveDocument as DrawingDocument;
            if (InvDoc.u.referendedDoc(m_Drw as Document).SubType == "{9C464203-9BAE-11D3-8BAD-0060B0CE6BB4}")
            {
                smcd = (SheetMetalComponentDefinition)((PartDocument)(InvDoc.u.referendedDoc(m_Drw as Document))).ComponentDefinition;
                if (smcd.HasFlatPattern)
                {
                    FlatPattern fp = smcd.FlatPattern;
                    for (int i = 1; i < fp.FlatBendResults.Count; i++)
                    {
                        if (fp.FlatBendResults[i].IsOnBottomFace)
                        {
                            Edge edge = fp.FlatBendResults[i].Edge;
                            fp.FlatBendResults[i].GetBendOrder(out j, out bendOrder);
                            val += j.ToString() + ';';
                            if (fp.FlatBendResults[i].IsDirectionUp) direct = "Вверх" + ';';
                            else direct = "Вниз" + ';';
                            val += direct;
                            dval = fp.FlatBendResults[i].Angle * 180 / 3.1415926;
                            ang = 180 - Math.Round(dval,1);
                            val += ang.ToString("0.##") + ';';
                            r = Math.Round(fp.FlatBendResults[i].InnerRadius, 2);
                            //if (r == 0.15) val += "1,5" + ';';
                            /*else*/ val += ((r*10).ToString()) + ';';
                            l = InvDoc.u.len(edge.StartVertex, edge.StopVertex);

                            foreach (DrawingCurve dc in lstDC)
                            {
                                if ( eq(edge, dc)
                                    //InvDoc.util.comparePoint2d(edge.StartVertex, ((Edge)dc.ModelGeometry).StartVertex, 1000) &&
                                    //        InvDoc.util.comparePoint2d(edge.StopVertex, ((Edge)dc.ModelGeometry).StopVertex, 1000))
                                    )
                                {
                                    upor = findDim(dc);
                                if (upor != "")
                                    break;
                                }
                            }

                            if (r > 1 && r < 10) val += "1;-;-;-;-;-;-;-;-";
                            else if (r >= 10) val += "10;-;-;-;-;-;-;-;-";
                            else
                            {
                                if (ang >= 90 && ang < 180)
                                {
                                    if (l >= 100 && l < 200) val += "2,16;90;0,8;8;90;0,8;" + l * 10 + ";" + upor + ";";
                                    else if (l >= 200 && l < 250) val += "16;90;0,8;8;90;0,8;" + l * 10 + ";" + upor + ";";
                                    else val += "7,8,9;90;0,8;8;90;0,8;" + l * 10 + ";" + upor + ";";
                                }
                                if (ang < 90 && ang >= 35)
                                {
                                    if (l >= 100 && l < 200) val += "2,16;35;0,8;8;35;0,8;" + l * 10 + ";" + upor + ";";
                                    else if (l >= 200 && l < 250) val += "16;35;0,8;8;35;0,8;" + l * 10 + ";" + upor + ";";
                                    else val += "7,8,9;35;0,8;8;35;0,8;" + l * 10 + ";" + upor + ";";
                                }
                                if (ang == 0)
                                {
                                    val = "";
                                    val += j + ";" + direct  + "45" + ";" + r*10 + ";" + "7,8,9;35;0,8;8;35;0,8;" + l * 10 + ";" + upor + ";";
                                }
                                if (ang < 0)
                                {
                                    val = "";
                                    val += j + ";" + direct + "45" + ";" + r * 10 + ";" + "7,8,9;35;0,8;8;35;0,8;" + l * 10 + ";" + upor + ";";
                                }
                            }
                            lst.Add(val);
                            
                            if (ang == 0)
                            {
                                val = "";
                                val += j + ".1;" + direct  + "0" + ";" + r*10 + ";"+ "8;180;-;-;-;-;-;-;-;";
                                lst.Add(val);
                            }
                            else if (ang < 0)
                            {
                                val = "";
                                val += j + ".1;" + direct  + Math.Round(ang,1) + ";" + r*10 + ";" + "16;180;-;-;-;-;-;-;-;";
                                lst.Add(val);
                            }    
                        }
                        val = "";
                        //if (upor == "") lst.Remove(lst[lst.Count - 1]);
                    }
                }
            }

            lst.Sort(delegate(string str1, string str2)
            {
                try
                {
                    int i1, i2;
                    i1 = Convert.ToInt16(str1.Substring(0, str1.IndexOf(';')));
                    i2 = Convert.ToInt16(str2.Substring(0, str2.IndexOf(';')));
                    return (i1 < i2) ? -1 : 1;
                }
                catch (Exception)
                {
                    string i1, i2;
                    i1 = str1.Substring(0, str1.IndexOf(';'));
                    i2 = str2.Substring(0, str2.IndexOf(';'));
                    return i1.CompareTo(i2);
                }
            });
            lst = removeDublicate(lst);
            lst = replace(lst, "2,16;90;0,8;8;90;0,8;", "2,16;35;0,8;8;35;0,8;");
            lst = replace(lst, "16;90;0,8;8;90;0,8;", "16;35;0,8;8;35;0,8;");
            lst = replace(lst, "7,8,9;90;0,8;8;90;0,8;", "7,8,9;35;0,8;8;35;0,8;");
        }

        public static List<string> replace(List<string> lst, string old, string n)
        {
            string tmp = lst.Find(e => e.IndexOf(n) != -1);
            List<string> tmplst1 = new List<string>();
            if (tmp != null)
            {
                foreach (var item in lst)
                {
                    tmp = item;
                    tmp = item.Replace(old, n);
                    tmplst1.Add(tmp);
                }
            }
            else return lst;
            return tmplst1;
        }

        public static List<string> removeDublicate(List<string> lst)
        {
            return lst.Distinct(new InvDoc.LambdaComparer<string>((val1, val2) =>
                val1.Substring(0, val1.IndexOf(';')) == val2.Substring(0, val2.IndexOf(';')))).ToList();
        }

        public static List<DrawingCurve> addBends()
        {
            Inventor.Application invApp = Macros.StandardAddInServer.m_inventorApplication;
            try
            {
                DrawingDocument m_Drw = invApp.ActiveDocument as DrawingDocument;
                List<Point2d> ptsExc = new List<Point2d>();
                invApp.ScreenUpdating = false;  bool flag = true;
                List<DrawingCurve> lstDC = new List<DrawingCurve>();
                List<int> valI = new List<int>();
                int j = 0, count = 0; BendOrderSourceTypeEnum bendOrder;
                DrawingView fpView = null; DrawingCurve dcStart = null;
                SheetMetalComponentDefinition smcd;
                if (InvDoc.u.referendedDoc(m_Drw as Document).SubType == "{9C464203-9BAE-11D3-8BAD-0060B0CE6BB4}")
                {
                    smcd = (SheetMetalComponentDefinition)((PartDocument)(InvDoc.u.referendedDoc(m_Drw as Document))).ComponentDefinition;
                    if (smcd.HasFlatPattern)
                    {
                        FlatPattern fp = smcd.FlatPattern;
                        foreach (DrawingView view in m_Drw.ActiveSheet.DrawingViews)
                        {
                            if (view.IsFlatPatternView)
                            {
                                fpView = view;
                                break;
                            }
                        }
                        if (fpView != null)
                        {
                            foreach (DrawingCurve dc in fpView.DrawingCurves)
                            {
                                if (dc.EdgeType == DrawingEdgeTypeEnum.kBendUpEdge || dc.EdgeType == DrawingEdgeTypeEnum.kBendDownEdge)
                                {
                                    bool fl = false;
                                    foreach (DrawingCurveSegment segm in dc.Segments)
                                    {
                                        if (segm.StartPoint.DistanceTo(segm.EndPoint) < 0.1)
                                        { fl = true; break; }
                                    }
                                    if (fl) continue;
                                    count++;
                                    if (count == 1) dcStart = dc;
                                    Edge edgeModel = (Edge)dc.ModelGeometry;
                                    lstDC.Add(dc);
                                    foreach (LeaderNote ldr in m_Drw.ActiveSheet.DrawingNotes.LeaderNotes)
                                    {
                                        if (ldr.Text.IndexOf("</Br>") != -1) break;
                                        LeaderNode node = ldr.Leader.RootNode;
                                        if (node.ChildNodes.Count != 0 && node.ChildNodes[1].AttachedEntity != null && node.ChildNodes[1].AttachedEntity.Geometry.Equals(dc))
                                        {
                                            flag = false;
                                            break;
                                        }
                                        //else flag = true;
                                    }
                                     if (flag)
                                    {
                                        foreach (FlatBendResult fbr in fp.FlatBendResults)
                                        {
                                            if (eq(edgeModel, fbr.Edge)) 
//                                                 InvDoc.util.comparePoint(edgeModel.StartVertex, fbr.Edge.StartVertex, 1000) &&
//                                                 InvDoc.util.comparePoint(edgeModel.StopVertex, fbr.Edge.StopVertex, 1000))
                                            {
                                                fbr.GetBendOrder(out j, out bendOrder);
                                                int v = valI.Find(delegate (int val) {return j == val;});
                                                if (v == 0)
                                                {
                                                valI.Add(j);
                                                addLeader(dc, j.ToString() /*+ "<Br/>"*/, true, ref ptsExc, ref fpView);
                                                break;
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                            if (flag) addLeader(dcStart, "ЛС ()", true, ref ptsExc, ref fpView, x: -0.4, y: -0.4);
                            flag = true;
                        }
                    }
                }
                invApp.ScreenUpdating = true;
                ptsExc.Clear();
                return lstDC;
            }
            catch (Exception)
            {
                invApp.ScreenUpdating = true;
                return null;
            }
        }

        static public DrawingCurveSegment findCenter(DrawingCurve dc)
        {
            Point2d center = InvDoc.u.midPt(dc.StartPoint, dc.EndPoint, 0, 0);
            foreach (DrawingCurveSegment seg in dc.Segments)
            {
                UnitVector2d start = center.VectorTo(seg.StartPoint).AsUnitVector(), end = center.VectorTo(seg.EndPoint).AsUnitVector();
                if (!start.IsEqualTo(end)) return seg;
            }
            return dc.Segments[1];
        }

        static public void addLeader(DrawingCurve dc ,string txt ,bool flag, ref List<Point2d> ptsExc, ref DrawingView dv, double x = 0.4, double y = 0.4)
        {
            Inventor.DrawingCurveSegment dcs;
            Inventor.Application invApp = Macros.StandardAddInServer.m_inventorApplication;
            DrawingDocument m_Drw = invApp.ActiveDocument as DrawingDocument;
            Inventor.GeometryIntent gInt;
            Inventor.ObjectCollection objs = invApp.TransientObjects.CreateObjectCollection();
            LeaderNote ln;
            //double [] midPt = new double[2];
            //double [] param = {0.5};
            Inventor.Point2d midPt; 
            try
            {
                 Sheet sh = m_Drw.ActiveSheet;
                    if (dc.Segments.Count == 1)
                    {
                        dcs = dc.Segments[1];
                    }
                    else
                    {
//                         int count;
//                         if (dc.Segments.Count % 2 == 0) count = dc.Segments.Count / 2;
//                         else count = dc.Segments.Count / 2 + 1;
//                         dcs = dc.Segments[count];
                        dcs = findCenter(dc);
                    }
                if (dcs.StartPoint.VectorTo(dcs.EndPoint).X > 0)
                    midPt = invApp.TransientGeometry.CreatePoint2d((dcs.EndPoint.X - dcs.StartPoint.X) / 1.9 + dcs.StartPoint.X, (dcs.EndPoint.Y - dcs.StartPoint.Y) / 1.9 + dcs.StartPoint.Y);
                else midPt = invApp.TransientGeometry.CreatePoint2d((dcs.EndPoint.X - dcs.StartPoint.X) / 2.1 + dcs.StartPoint.X, (dcs.EndPoint.Y - dcs.StartPoint.Y) / 2.1 + dcs.StartPoint.Y);

                    if (dc.EdgeType == DrawingEdgeTypeEnum.kBendDownEdge || dc.EdgeType == DrawingEdgeTypeEnum.kBendUpEdge)
                    {
                        
                        gInt = sh.CreateGeometryIntent(dc, midPt);
                        if (flag)
                        {
                            //Vector2d vec = invApp.TransientGeometry.CreateVector2d(x, y);
                            //if (midPt.X > dv.Position.X || midPt.Y > dv.Position.Y)
                            //{
                            //    vec = invApp.TransientGeometry.CreateVector2d(-x, -y);
                            //}
                            //Point2d pt = midPt.Copy();
                            //    //invApp.TransientGeometry.CreatePoint2d(midPt.X + x, midPt.Y + y);
                            //ObjectsEnumerator ob; double tol = 0.5;
                            //pt.TranslateBy(vec);
                            //ob = sh.FindUsingPoint(pt, tol);
                            //int i = 0; double ofset = 0.2;
                            
                            //while (ob.Count != 0 || ptsExc.Exists(e => InvDoc.util.eq(e,pt)))
                            //{
                            //    i++;
                            //    vec = InvDoc.util.rotate(vec, midPt, Math.PI / 2);
                            //    pt = midPt.Copy();
                            //    pt.TranslateBy(vec);
                            //    //ptsExc.Add(pt);
                            //    ob = sh.FindUsingPoint(pt, tol);
                            //    if (i % 4 == 0)
                            //    {
                            //        vec = invApp.TransientGeometry.CreateVector2d(x+ofset, y+ofset);
                            //        ofset += 0.2;
                            //    }
                            //}
                            //objs.Add(pt);
                            //ptsExc.Add(pt);
                            objs.Add(InvDoc.u.spiralPosition(dv, sh, x, y, midPt, ref ptsExc));
                        }
                        objs.Add(gInt);
                        ln = m_Drw.ActiveSheet.DrawingNotes.LeaderNotes.Add(objs, txt);
                        //ln.AttributeSets.Add("Bend").Add("Bend", ValueTypeEnum.kStringType, txt);

                        //ln.LineSpacingType = TextLineSpacingTypeEnum.kExactLineSpacing;
                        //ln.LineSpacing = 0.2;
                        //ln.DimensionStyle.TextStyle.Rotation = 3.1415926 / 2;
                        //ln.DimensionStyle.LinearMultiLineTextOrientation = MultiLineDimensionTextEnum.kFirstLineCenteredOnLandingLine;
                    }
                }
                catch(Exception ex)
            {
                    MessageBox.Show(ex.ToString());
            }
        }

        public List<string> performanceDinamic(List<string> asmDocs)
        {
            List<string> lst = new List<string>();
            int count = asmDocs.Count;
            Documents docs = Macros.StandardAddInServer.m_inventorApplication.Documents;
            NameValueMap nvmOptions = I.objs.CreateNameValueMap();
            nvmOptions.Add("SkipAllUnresolvedFiles", true); int i = 0;
            foreach (var item in asmDocs)
            {
                Document doc = docs.OpenWithOptions(item, nvmOptions, false);
                string pn = InvDoc.u.getProp(doc, "Part Number").Value.ToString(), desc = InvDoc.u.getProp(doc,"Description").Value.ToString(),
                    note = InvDoc.u.getProp(doc,"Comments").Value.ToString();
                string type = "", decnumber = "", c = count.ToString("00"), EWA = "";
                string tmp = CreateComponent.perf(desc, pn,ref type,ref decnumber, ref note, EWA, ref c);
                Regex regex = new Regex(@"\b(\w*-)(\d*)(\w.*)");
                Match m = regex.Match(type);
//                 if (m.Groups[1].Value != "") m.Groups[1
                tmp = m.Groups[1] + note + m.Groups[3];  
                    //type.Insert(m.Groups[1].Index, note);
                Regex r = new Regex(@"\b(\d.*)\b");
                m = r.Match(desc);
                if (m.Groups[0].Value != null && m.Groups[0].Value != "")
                {
                    tmp = m.Groups[0].Value;
                }
                if (i == 0)
                {
                    lst.Add(pn + "      " + ";" + tmp + ";" + note + ";");
                }
                else
                {
                    lst.Add("-" + i.ToString("00") + ";" + tmp + ";" + note + ";");
                }
                i++;
            }
            return lst;
        }

        public void performance(ref List<string> lst)
        {
            if (InvDoc.u.referendedDoc(m_Drw as Document).DocumentType == DocumentTypeEnum.kAssemblyDocumentObject)
            {
                AssemblyDocument asmDoc = (AssemblyDocument)m_Drw.Sheets[1].DrawingViews[1].ReferencedDocumentDescriptor.ReferencedDocument;
                string n = "", val = ""; string[] spl; Regex regex, r2, r3; Match m1, m2; int isp = 0; List<string> exept = null;
                string name = asmDoc.PropertySets[3][2].Value.ToString();
                regex = new Regex(@"-(\d{2})\b");

                VariableDataForSpec vd = new VariableDataForSpec((Document)asmDoc);
                List<string> vals = new List<string>();
                vals = vd.AttribLoad("Var");
                if (vals.Count != 0)
                {
                spl = vals[vals.Count-1].Split(';');
                isp = int.Parse(spl[0].Substring(4));
                }

                int k = 0, start = 0;
                if (regex.IsMatch(name))
                {
                    val = regex.Match(name).Groups[1].Value;
                    start = int.Parse(val);
                }
                k = start;
                string desc = asmDoc.PropertySets[3][14].Value.ToString();
                Property t = InvDoc.u.getProp((Document)asmDoc, "perf");
                if (name.StartsWith("AC") || name.StartsWith("АС"))
                {
                    regex = new Regex(@"(\d+)", RegexOptions.IgnoreCase);
                    r2 = new Regex(@"(\w+\d+-\d+[E|Е])");
                    m1 = r2.Match(desc);
                    exept = new List<string>(3) { "39", "59", "79" };

                    for (int i = 0; i < regex.Matches(desc).Count; i++)
                    {
                        n = "";
                        string str = regex.Matches(desc)[i].Groups[1].Value;
                        if (str.Length > 2 || exept.Exists(e => e == str)) continue;
                        int ind = regex.Matches(desc)[i].Groups[1].Index;
                        if (m1.NextMatch().Success && m1.NextMatch().Index < ind) n = m1.NextMatch().Value + str;
                        else if (m1.Success && m1.Index < ind) n = m1.Value + str;
                        if (start == 0) name = name + "      ";
                        if (k == start) lst.Add(name + ";" + n + ";" + str + ";" + "208;");
                        else lst.Add("-" + k.ToString("00") + ";" + n + ";" + str + ";" + "208;");
                        k++;
                    }
                    if (k < isp)
                    {
                        for (int i = 0; i < regex.Matches(desc).Count; i++)
                        {
                            n = "";
                            string str = regex.Matches(desc)[i].Groups[1].Value;
                            if (str.Length > 2 || exept.Exists(e => e == str)) continue;
                            int ind = regex.Matches(desc)[i].Groups[1].Index;
                            if (m1.NextMatch().Success && m1.NextMatch().Index < ind) n = m1.NextMatch().Value + str;
                            else if (m1.Success && m1.Index < ind) n = m1.Value + str;
                            if (start == 0) name = name + "      ";
                            if (k == start) lst.Add(name + ";" + n + ";" + str + ";" + "600;");
                            else lst.Add("-" + k.ToString("00") + ";" + n + ";" + str + ";" + "600;");
                            k++;
                        }
                    }
                }
                else if (t != null)
                {
                    string perfstr = t.Value.ToString();

                    foreach (var item in perfstr.Split(';'))
                    {
                        regex = new Regex(@"\d+", RegexOptions.IgnoreCase);
                        n = "";
                        string str = regex.Matches(item)[0].Value;
                        if (str.Length > 3) continue;
                        if (start == 0) name = name + "      ";
                        if (k == start) lst.Add(name + ";" + item + ";" + str + ";");
                        else lst.Add("-" + k.ToString("00") + ";" + item + ";" + str + ";");
                        k++;
                    }
                    for (int i = 0; i < regex.Matches(perfstr).Count; i++)
                    {
                        
                    }
                }
                else
                {
                    regex = new Regex(@"(\d+)", RegexOptions.IgnoreCase);
                    r2 = new Regex(@"(\w+-)");
                    m1 = r2.Match(desc);
                    r3 = new Regex(@"(П\d{4})");
                    m2 = r3.Match(desc);
                    for (int i = 0; i < regex.Matches(desc).Count; i++)
                    {
                        n = "";
                        string str = regex.Matches(desc)[i].Groups[1].Value;
                        if (str.Length > 3) continue;
                        int ind = regex.Matches(desc)[i].Groups[1].Index;
                        if (m1.Success && m1.Index < ind) n = m1.Value + str;
                        else if (m1.NextMatch().Success && m1.NextMatch().Index < ind) n = m1.NextMatch().Value + str;
                        if (m2.Success && m2.Index > ind) n += m2.Value;
                        else if (m2.NextMatch().Success && m2.NextMatch().Index > ind) n += m2.NextMatch().Value;
                        if (start == 0) name = name + "      ";
                        if (k == start) lst.Add(name + ";" + n + ";" + str + ";");
                        else lst.Add("-" + k.ToString("00") + ";" + n + ";" + str + ";");
                        k++;
                    }
                }
            }
        }

        public void LoadTable(string name)
        {
            List<string> tmp = new List<string>();
            List<string> bnds = new List<string>();
            List<string> perf = new List<string>();
            string rowDim = ""; string[] sPL;
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
            //string[] sPL = tmp[0].Split(';');
            string strTmp = tbl.substring(tmp[0], "Name=");
            if (strTmp != "") this.textBox2.Text = strTmp;
            string fn = tbl.substring(tmp[0], "FontName=");
            if (fn != "") setting.textBox3.Text = strTmp;
            strTmp = tbl.substring(tmp[0], "TextHeight=");
            if (strTmp != "") setting.comboBox1.Text = strTmp;

            float he = Convert.ToSingle(strTmp.Replace(',',separator));
            he = textCompRev[he];
            if (strTmp != "12.0" || fn != "GOST Common") _dgv.DefaultCellStyle.Font = fontStyle(italic: true, textH: he, nameFont: fn);
            strTmp = tbl.substring(tmp[0], "TextWidth=");
            if (strTmp != "") setting.comboBox2.Text = strTmp;
            strTmp = tbl.substring(tmp[0], "RowHeight=");
            if (strTmp != "") setting.comboBox8.Text = strTmp;
            strTmp = tbl.substring(tmp[0], "ColumnWidth=");
            if (strTmp != "") setting.comboBox9.Text = strTmp;
            strTmp = tbl.substring(tmp[0], "Align=");
            if (strTmp != "") setting.comboBox7.Text = strTmp;

            strTmp = tbl.substring(tmp[0], "Title=");
            if (strTmp != "") setting.textBox2.Text = strTmp;
            if (strTmp != "") setting.checkBox2.Checked = true;

            strTmp = tbl.substring(tmp[0], "FormatRows=");
            if (strTmp != "") setting.textBox1.Text = strTmp.Replace('`',';');
            strTmp = tbl.substring(tmp[0], "SelectColumns=");
            if (strTmp != "") setting.checkBox1.Checked = Convert.ToBoolean(strTmp);
            strTmp = tbl.substring(tmp[0], "CompressText=");
            if (strTmp != "") setting.checkBox3.Checked = Convert.ToBoolean(strTmp);
            strTmp = tbl.substring(tmp[0], "Italic=");
            if (strTmp != "") setting.checkBox4.Checked = Convert.ToBoolean(strTmp);
            strTmp = tbl.substring(tmp[0], "Bold=");
            if (strTmp != "") setting.checkBox5.Checked = Convert.ToBoolean(strTmp);
            strTmp = tbl.substring(tmp[0], "ColumnsDimension=");

            if (strTmp != "")
            {
                sPL = strTmp.Split('`');
                _dgv.ColumnCount = sPL.Count() - 1;
                this.CountColumn.Text = (sPL.Count() - 1).ToString();
                this.ColumnDimensions.Text = strTmp.Replace('`',';');
            }
            strTmp = tbl.substring(tmp[0], "Padding=");
            if (strTmp != "")
            {
                sPL = strTmp.Split('`');
                setting.comboBox3.Text = sPL[0];
                setting.comboBox4.Text = sPL[1];
                setting.comboBox5.Text = sPL[2];
                setting.comboBox6.Text = sPL[3];
            }
            _dgv.RowCount = 1;
            _dgv.Rows.Clear();
            int j = 0;
            for (int i = 2; i < tmp.Count; i+=2)
            { 
                strTmp = tbl.substring(tmp[i], "RowHeight=");
                if (strTmp != "") rowDim += strTmp + ';';
                else rowDim += tbl.substring(tmp[0], "RowHeight=") + ';';
                sPL = tmp[i+1].Split(';');
                _dgv.Rows.Add(sPL);
                j++;
            }
            if (name == "Гибка")
            {
                List<DrawingCurve> lstDC = new List<DrawingCurve>();
                lstDC = addBends();
                LoadBend(ref bnds, lstDC);
                InvDoc.u.replaceText(Macros.StandardAddInServer.m_inventorApplication.ActiveDocument as DrawingDocument, "ЛС", "ЛС (" + bnds.Count + ")");
                foreach (string str in bnds)
                {
                    sPL = str.Split(';');
                    _dgv.Rows.Add(sPL);
                    rowDim += tbl.substring(tmp[0], "RowHeight=") + ';';
                }
            }
            if (name == "Исполнения" || name == "Исполнения_AC")
            {
                if (InvDoc.u.referendedDoc(m_Drw as Document).DocumentType == DocumentTypeEnum.kAssemblyDocumentObject)
                {
                    List<string> asms = TableInv.getAsms(InvDoc.u.referendedDocDesc(m_Drw as Document).FullDocumentName);
                    if (asms.Count > 1) perf = performanceDinamic(asms);
                }
                else performance(ref perf);
                _dgv.Columns[0].CellTemplate.Style.Alignment = DataGridViewContentAlignment.MiddleRight;
                _dgv[0, 0].Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                foreach (string str in perf)
                {
                    sPL = str.Split(';');
                    _dgv.Rows.Add(sPL);
                    rowDim += tbl.substring(tmp[0], "RowHeight=") + ';';
                }
            }

            //if (name == "Изделие") LoadBend(ref );
            //foreach (string str in bnds)
            //{
            //    sPL = str.Split(';');
            //    _dgv.Rows.Add(sPL);
            //    rowDim += tbl.substring(tmp[0], "RowHeight=") + ';';
            //}
            rowDim += tbl.substring(tmp[0], "RowHeight=") + ';';
            this.CountRow.Text = _dgv.RowCount.ToString();
            this.RowDimensions.Text = rowDim;
            EventArgs e = null; Object obj = null;
            ColumnDimensions_Leave(obj, e);
            RowDimensions_Leave(obj, e);
        }

        private void toolStripMenuItem2_OwnerChanged(object sender, EventArgs e)
        {
            if (this._dgv.SelectedCells.Count != 0)
            {
                ToolStripComboBox tscb = (ToolStripComboBox)sender;

                Func<string, double> conv = s =>
                {
                    char rep = (separator == ',') ? '.' : ',';
                    return (s.IndexOf(separator) != -1) ? Convert.ToDouble(s) : Convert.ToDouble(s.Replace(rep, separator));
                };

                double tmp = conv(tscb.SelectedItem.ToString());
                float h = (float)tmp;
                h = textCompRev[h];
                foreach (DataGridViewCell c in this._dgv.SelectedCells)
                {
                    if (c.Style.Font != null)
                        c.Style.Font = fontStyle(italic: true, underline: c.Style.Font.Underline, bold: c.Style.Font.Bold, textH: h);
                    else c.Style.Font = fontStyle(textH: h);
                }
            }
        }

        private void toolStripMenuItem2_Enter(object sender, EventArgs e)
        {
            
        }

        private void высотаТекстаToolStripMenuItem_Click(object sender, EventArgs e)
        {
            
        }

        private void высотаТекстаToolStripMenuItem_DropDownOpening(object sender, EventArgs e)
        {
            //if (this._dgv.SelectedCells.Count != 0)
            //{
            //    ToolStripComboBox tscb = toolStripMenuItem2;
            //    float txt;
            //    if (_dgv.CurrentCell.Style.Font != null)
            //        txt = _dgv.CurrentCell.Style.Font.Size;
            //    else txt = _dgv.DefaultCellStyle.Font.Size;
            //    txt = textComp[txt];
            //    tscb.SelectedText = txt.ToString();
            //}
        }

        private void toolStripTextBox1_DoubleClick(object sender, EventArgs e)
        {
            try
            {
            System.Windows.Forms.ToolStripTextBox tb = toolStripTextBox1;
            if (this._dgv.SelectedCells.Count != 0 && tb.Text != "GOST Common")
            {
                foreach (DataGridViewCell c in this._dgv.SelectedCells)
                {
                    if (c.Style.Font != null)
                        c.Style.Font = fontStyle(italic: c.Style.Font.Italic, underline: c.Style.Font.Underline, bold: c.Style.Font.Bold,nameFont: tb.Text);
                    else c.Style.Font = fontStyle(nameFont: tb.Text);
                }
            }
            }
            catch(Exception)
            {

            }
        }

        private void toolStripComboBox1_TextChanged(object sender, EventArgs e)
        {
            if (this._dgv.SelectedCells.Count != 0)
            {
               string val = "";
               foreach (DataGridViewCell c in this._dgv.SelectedCells)
               {
                   val += c.RowIndex.ToString() + ';' + c.ColumnIndex.ToString() + ';' +  toolStripComboBox1.Text;
               }
               string tmp = styleParam.Find(delegate(string str) { return str == val; });
               if (tmp == null && toolStripComboBox1.Text != "100") styleParam.Add(val);
            }
            
        }

        private void сжатиеToolStripMenuItem_DropDownOpening(object sender, EventArgs e)
        {
            try
            {
                bool flag = true;
                foreach (string str in styleParam)
                {
                    string[] spl = str.Split(';');
                    if (spl[0] == _dgv.CurrentCell.RowIndex.ToString() && spl[1] == _dgv.CurrentCell.ColumnIndex.ToString())
                    {
                        toolStripComboBox1.Text = spl[2];
                        flag = false;
                        break;
                    }
                }
                if (flag) toolStripComboBox1.Text = "100";
            }
            catch (Exception)
            {
                toolStripComboBox1.Text = "100";
            }
        }

        //private void toolStripMenuItem2_Leave(object sender, EventArgs e)
        //{
        //    
        //}
    }

    internal class TablesButton : Button
    {
        public static Tables m_Table;
        public static string nameTable {get; set;}
        public Inventor.Document pDoc { get; set; }
        public static Tables getTable
        {
            get
            {
                return m_Table;
            }
        }

        #region "Methods"
        public TablesButton(string displayName, string internalName, string clientId, string description, string tooltip, Icon standardIcon, Icon largeIcon)
            : base(displayName, internalName, clientId, description, tooltip, standardIcon, largeIcon)
        {

        }

        protected override void ButtonDefinition_OnExecute(NameValueMap context)
        {
            Macros.StandardAddInServer.forms.Add(m_Table);
            if (Macros.StandardAddInServer.activeteForm()) System.Windows.Forms.Application.Run(m_Table = new Tables(InventorApplication.ActiveDocument, InventorApplication, nameTable));
        }

        #endregion
    }
    public class DataGridViewRowsReorderBehavior
    {
        private DataGridViewRow _RowUnderMouseDown;
        public string val1, val2; public int x, y;
        public DataGridViewRowsReorderBehavior(DataGridView grid)
        {
            grid.AllowDrop = true;
            grid.MouseMove += new MouseEventHandler(Grid_MouseMove);
            grid.MouseDown += new MouseEventHandler(Grid_MouseDown);
            grid.DragOver += new DragEventHandler(Grid_DragOver);
            grid.DragDrop += new DragEventHandler(Grid_DragDrop);
        }

        private void Grid_CellBeginEdit(object sender, DataGridViewCellCancelEventArgs e)
        {
            _RowUnderMouseDown = ((DataGridView)sender).Rows[e.RowIndex];
            if (e.ColumnIndex == 3 || e.ColumnIndex == 4)
            {
                x = e.ColumnIndex; y = e.RowIndex;
                val1 = _RowUnderMouseDown.Cells[3].Value.ToString();
                val2 = _RowUnderMouseDown.Cells[4].Value.ToString();
            }
        }

        private void Grid_DragDrop(object sender, DragEventArgs e)
        {
            if (_RowUnderMouseDown != null)
            {
                DataGridViewRow rowUnderMouse = FindRow(_RowUnderMouseDown.DataGridView, e.X, e.Y, true);
                if (rowUnderMouse != null && e.Effect == DragDropEffects.Move)
                {
                    DataGridView grid = _RowUnderMouseDown.DataGridView;
                    try
                    {
                        //object data = _RowUnderMouseDown.DataBoundItem;
                        //CurrencyManager cm = (CurrencyManager)grid.BindingContext[grid.DataSource, grid.DataMember];
                        int index = rowUnderMouse.Index + 1;

                        //cm.List.Insert(index, data);
                        DataGridViewRow dgr = grid.Rows.SharedRow(_RowUnderMouseDown.Index);
                        int oldIndex = _RowUnderMouseDown.Index;
                        //grid.Rows.InsertCopy(oldIndex, index+1);
                        grid.Rows.Insert(index, 1);
                        for (int i = 0; i < dgr.Cells.Count; i++)
                        {
                            //grid.Rows[oldIndex].Cells[i].Value = grid.Rows[index].Cells[i].Value;
                            grid.Rows[index].Cells[i].Value = dgr.Cells[i].Value;
                        }
                        if (index < oldIndex)
                            grid.Rows.RemoveAt(oldIndex + 1);
                        else
                            grid.Rows.RemoveAt(oldIndex);
                    }
                    catch (System.Exception ee)
                    {
                        MessageBox.Show(ee.ToString());
                    }
                }
            }
        }
        private void Grid_DragOver(object sender, DragEventArgs e)
        {
            e.Effect = DragDropEffects.Move;
        }
        private void Grid_MouseDown(object sender, MouseEventArgs e)
        {
            _RowUnderMouseDown = FindRow(sender as DataGridView, e.X, e.Y, false);
        }
        private void Grid_MouseMove(object sender, MouseEventArgs e)
        {
            if (_RowUnderMouseDown != null && (e.Button & MouseButtons.Left) == MouseButtons.Left)
            {
                Rectangle rect = _RowUnderMouseDown.DataGridView.GetRowDisplayRectangle(_RowUnderMouseDown.Index, true);
                if (rect != Rectangle.Empty && rect.Contains(e.X, e.Y) == false)
                    _RowUnderMouseDown.DataGridView.DoDragDrop(_RowUnderMouseDown, DragDropEffects.Move);
            }
        }

        private static DataGridViewRow FindRow(DataGridView grid, int x, int y, bool convertToLocal)
        {
            if (convertToLocal)
            {
                System.Drawing.Point p = grid.PointToClient(new System.Drawing.Point(x, y));
                x = p.X;
                y = p.Y;
            }
            int index = grid.HitTest(x, y).RowIndex;
            return (index >= 0) ? grid.Rows[index] : null;
        }
    }
}
