using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Inventor;
using TableDll;
using System.Xml;
using System.Xml.Linq;

namespace InvAddIn
{
    public partial class Table : Form
    {
        TableInv tbl;
        public Inventor.AssemblyDocument m_AsmDoc;
        public string partNumberAsm;
        public Inventor.ComponentDefinition oCompDef;
        public Inventor.Document oDoc;
        public Inventor.DrawingDocument m_DrwDoc;
        private Inventor.BOM m_BOM;
        private Inventor.BOMView m_BOMView;
        private Inventor.TransientGeometry m_TG;
        private Inventor.SketchedSymbolDefinition m_SketchDef;
        private Inventor.DrawingSketch m_DrwSketch;
        private Inventor.SketchLine m_SketchLine;
        private VariableDataForSpec varForSpec;
        public Dictionary<string, string> dic;
        private bool m_first;
        private string path, value;
        private string[] strs;
        private List<Inventor.Point2d> pts = new List<Point2d>();
        //BindingSource bs;
        public static System.Drawing.Point point = new System.Drawing.Point();
        public int pos = 0;
        public double scale = 50;
        int start = 0;
        //public int maxWidth, maxHeigth;
        public Rectangle boundsScreen;
                                
        public Table(Inventor.Document pDoc, Inventor.Application m_InvApp)
        {
            boundsScreen = Screen.PrimaryScreen.WorkingArea;
            scale = boundsScreen.Width * 0.465 / 16.5;
            m_TG = m_InvApp.TransientGeometry;
            switch (pDoc.DocumentType)
            {
                case DocumentTypeEnum.kAssemblyDocumentObject:
                    m_AsmDoc = (AssemblyDocument)pDoc;
                    break;
                case DocumentTypeEnum.kDrawingDocumentObject:
                    m_DrwDoc = (DrawingDocument)pDoc;
                    m_AsmDoc = (AssemblyDocument)InvDoc.u.referendedDoc(m_DrwDoc as Document);//m_DrwDoc.ReferencedDocuments.Count];
                    if (m_AsmDoc.FullDocumentName.EndsWith(">"))
                        m_AsmDoc = (AssemblyDocument)((Inventor.Application)m_AsmDoc.Parent).Documents.Open(m_AsmDoc.FullFileName, false);
                    break;                  
            }
            path = pDoc.FullFileName.ToString();
            path = path.Substring(0, path.LastIndexOf('\\'));
            m_BOM = m_AsmDoc.ComponentDefinition.BOM;
            partNumberAsm = m_AsmDoc.PropertySets[3][2].Value.ToString();
            m_BOM.StructuredViewEnabled = true;
              if (!m_BOM.StructuredViewFirstLevelOnly)
                m_BOM.StructuredViewFirstLevelOnly = true;
              //if (m_BOM.StructuredViewFirstLevelOnly)
              //    m_BOM.StructuredViewFirstLevelOnly = false;
            m_BOMView = m_BOM.BOMViews["Структурированный"];
            InitializeComponent();
            this.ClientSize = Screen.PrimaryScreen.WorkingArea.Size;
            this.Format.Width = (int)(0.6 * scale);
            this.Zone.Width = (int)(0.6 * scale);
            this.Position.Width = (int)(0.8 * scale);
            this.PartNumber.Width = (int)(5 * scale);
            this.Description.Width = (int)(6.3 * scale);
            this.Count.Width = (int)(1 * scale);
            this.Note.Width = (int)(2.2 * scale);

            this.FormatVar.Width = this.Format.Width;
            this.ZoneVar.Width = this.Zone.Width;
            this.PositionVar.Width = this.Position.Width;
            this.PartNumberVar.Width = this.PartNumber.Width;
            this.DescriptionVar.Width = this.Description.Width;
            this.CountVar.Width = this.Count.Width;
            this.NoteVar.Width = this.Note.Width;

            int offset = (int)(0.004*boundsScreen.Width);
            int width = dgv.Columns.OfType<DataGridViewColumn>().Sum(e => e.Width);
            width += 22;
            this.dgv.Size = new System.Drawing.Size(width, this.Size.Height-menu.Height-50/*(int)(boundsScreen.Height*0.08)*/);
            this.dgv.Location = new System.Drawing.Point(0, this.menu.Bounds.Bottom);
            this.dgvVar.Size = dgv.Size;
            this.dgvVar.Location = new System.Drawing.Point(dgv.Location.X + dgv.Width, dgv.Location.Y);
            this.dgv.RowHeadersWidth = this.dgvVar.RowHeadersWidth = 20;

            label1.Location = new System.Drawing.Point(dgv.Location.X + width / 2 - label1.Width/2, dgv.Location.Y /*- label1.Width/2*/ - 24);
            label2.Location = new System.Drawing.Point(dgv.Location.X + width + width / 2 - label2.Width/2, dgv.Location.Y /*- label2.Width/2*/ - 24);
            label2.Visible = false;
            dgvVar.Visible = false;

            width *= 2;
            width += (int)(0.001 * boundsScreen.Width);
            int bottomButton = dgv.Bounds.Bottom + (boundsScreen.Height - dgv.Bounds.Bottom)/2 - button1.Bounds.Height/2;
            this.button1.Width = button2.Width = button3.Width = txtBox.Width = /*txtBox3.Width = */(int)(0.045 * boundsScreen.Width);
            this.button1.Location = new System.Drawing.Point(dgvVar.Bounds.Right + offset, this.menu.Bounds.Bottom + offset);
            this.button3.Location = new System.Drawing.Point(button1.Bounds.Left, button1.Bounds.Bottom + offset);
            this.button2.Location = new System.Drawing.Point(button3.Bounds.Left, button3.Bounds.Bottom + offset);
            this.lbl.Location = new System.Drawing.Point(button2.Bounds.Left, button2.Bounds.Bottom + offset);
            this.txtBox.Location = new System.Drawing.Point(button2.Bounds.Left, lbl.Bounds.Bottom + offset);
            //this.lbl3.Location = new System.Drawing.Point(button2.Bounds.Left, txtBox.Bounds.Bottom + offset);
            //this.txtBox3.Location = new System.Drawing.Point(button2.Bounds.Left, lbl3.Bounds.Bottom + offset);

            //initSpec(dgv);

            tbl = tbl ?? new TableInv(m_DrwDoc, @"C:\ProgramData\Autodesk\Inventor Addins\Sequence.xml");
            //if (txtBox.Text != "") tbl.split = txtBox.Text;
            //VariableDataForSpec vd = new VariableDataForSpec((Document)m_AsmDoc);
            //List<string> lst = vd.AttribLoad("Split");
            AttributeSet attSet;
            if (m_AsmDoc.AttributeSets.get_NameIsUsed("Split"))
            {
                attSet = m_AsmDoc.AttributeSets["Split"];

                if (attSet.Count != 0) txtBox.Text = attSet[1].Value.ToString();
            }

            tbl.addTable(m_BOMView);
            if (tbl.countPerf == 0)
                tbl.showTable(dgv,0,100, tbl.group, varData:false);
            else
            {
                tbl.showTable(dgv, 0, 100, tbl.group,varData:false);
                tbl.showTable(dgvVar, 100, 5000, tbl.group,varData:true);
                dgvVar.Visible = true;
            }
            //tbl.asmDoc = m_AsmDoc;
            balloon();
            dgv.RowHeaderMouseClick += dataGridView1_RowHeaderMouseClick;
            dgv.KeyDown += dgv_KeyDown;
        }

        private void balloon()
        {
            string partNumber, desc, note;
            dic = dic ?? new Dictionary<string, string>();
            dic.Clear();
            foreach (Balloon ball in m_DrwDoc.ActiveSheet.Balloons)
            {
            foreach (BalloonValueSet item in ball.BalloonValueSets)
            {
                BOMRow row = item.ReferencedRow.BOMRow;
                dic = tbl.propDoc(row);
                partNumber = dic["PartNumber"]; desc = dic["desc"]; note = dic["note"];
                TableRow tr = tbl.rows.Find(r => r.cells[3].value == partNumber && r.cells[4].value == desc);

                if (tr != null)
                {
                    if (tr.group < 100) dgv[0,tr.dvgRow].Style.BackColor = System.Drawing.Color.Green;
                    else dgvVar[0, tr.dvgRow].Style.BackColor = System.Drawing.Color.Green;
                }
            }

            }
        }

        private void renumber(DataGridView _dgv)
        {
            int i = 1; int val = 0;
            Dictionary<int, int> ins = new Dictionary<int, int>();
            Dictionary<int, int> rep = new Dictionary<int, int>();

            try
            {
                foreach (DataGridViewRow dgvr in _dgv.Rows)
                {
                    if (dgvr.Cells[3].Value != null && dgvr.Cells[4].Value.ToString() == "Переменные данные для исполнений:") break;
                    if (dgvr.Cells[3].Value != null && dgvr.Cells[2].Value.ToString() != "")
                    {
                        val = Convert.ToInt16(dgvr.Cells[2].Value);
                        if (!rep.ContainsKey(val))
                        {
                            rep.Add(val, i);
                            i++;
                        }
                        else dgvr.Cells[2].Value = val + 100;
                    }
                }

                i = 1; val = 0;
                foreach (DataGridViewRow dgvr in _dgv.Rows)
                {
                    if (dgvr.Cells[3].Value != null && dgvr.Cells[2].Value.ToString() != "")
                    {
                        val = Convert.ToInt16(dgvr.Cells[2].Value);
                        if (!ins.ContainsKey(val))
                        {
                            ins.Add(val,i);
                            i++;
                        }
                    }
                }

                foreach (DataGridViewRow dgvr in _dgv.Rows)
                {
                    if (dgvr.Cells[3].Value != null && dgvr.Cells[2].Value.ToString() != "")
                    {
                        val = Convert.ToInt16(dgvr.Cells[2].Value);
                        dgvr.Cells[2].Value = ins[val];
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private void rowDelete(BOMView bomView, DataGridView _dgv)
        {
            bool remove = true;
            foreach (DataGridViewRow dgvr in _dgv.Rows)
            {
                if (dgvr.Cells[3].Value != null && dgvr.Cells[2].Value.ToString() != "") 
                {
                    foreach (Inventor.BOMRow row in m_BOMView.BOMRows)
                    {
                        oCompDef = row.ComponentDefinitions[1];
                        oDoc = (Inventor.Document)oCompDef.Document;
                        string partNumber = (string)oDoc.PropertySets[3][2].Value;
                        string desc = (string)oDoc.PropertySets[3][14].Value;
                        string key = (string)oDoc.PropertySets[1][4].Value.ToString();
                        string pos = (string)row.ItemNumber;
                        partNumber = partNumber.Trim();
                        if (partNumber == dgvr.Cells[3].Value.ToString() && desc == dgvr.Cells[4].Value.ToString())
                        {
                            remove = false;
                            break;
                        }
                    }
                    if (remove && dgvr.Cells[4].Value.ToString().StartsWith("p"))
                    _dgv.Rows.Remove(dgvr);
                remove = true;
            }
            }
        }

        private void copySelectedRowsToClipboard(DataGridView dgv)
        {
            dgv.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableWithoutHeaderText;
            Clipboard.Clear();
            if (dgv.GetClipboardContent() != null)
            {
                Clipboard.SetDataObject(dgv.GetClipboardContent());
                Clipboard.GetData(DataFormats.Text);
                IDataObject dt = Clipboard.GetDataObject();
                if (dt.GetDataPresent(typeof(string)))
                {
                    string tb = (string)(dt.GetData(typeof(string)));
                    Encoding encoding = Encoding.GetEncoding(1251);
                    byte[] dataStr = encoding.GetBytes(tb);
                    Clipboard.SetDataObject(encoding.GetString(dataStr));
                }
            }
            dgv.ClipboardCopyMode = DataGridViewClipboardCopyMode.Disable;
        }

        private void safeToolStripMenuItem_Click(object sender, System.EventArgs e)
        {
            try
            {
                //varForSpec = varForSpec ?? new VariableDataForSpec();
                //varForSpec.doc = (Document)m_AsmDoc;
                //dic = dic ?? new Dictionary<string, string>();
                //if (dic.Count == 0)
                //{
                //    dic.Add("Row", ""); dic.Add("Split", txtBox.Text); dic.Add("Power", txtBox3.Text);
                //}
                //varForSpec.AttribAdd(dgv, dic);
                VariableDataForSpec vd = new VariableDataForSpec((Document)m_AsmDoc);
                vd.AttribDelete("Split");
                if (txtBox.Text != "") vd.AttribAdd<string>("Split", txtBox.Text, ValueTypeEnum.kStringType);
                tbl.saveInAsm();
            }
            catch (Exception ex) { MessageBox.Show(ex.ToString()); }
        }

        private void clearToolStripMenuItem_Click(object sender, System.EventArgs e)
        {
            //string name = "Row";
            //Object pr;
            //try
            //{
            //    varForSpec = varForSpec ?? new VariableDataForSpec((Document)m_AsmDoc);
            //dic = dic ?? new Dictionary<string, string>();
            //if (dic.Count == 0)
            //{
            //    dic.Add("Row", ""); dic.Add("Isp",""); dic.Add("Split", ""); dic.Add("Power", "");
            //}
            //varForSpec.AttribDelete(dic);
            //}
            //catch (Exception ex)
            //{
            //    MessageBox.Show(ex.ToString());
            //}
            //m_AsmDoc.AttributeManager.PurgeAttributeSets("*",false, out pr);
            tbl.clearPerf();
            tbl.saveInAsm();
        }

        private void CutSelectedRowsToClipboard(DataGridView dgv)
        {
            dgv.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableWithoutHeaderText;
            Clipboard.Clear();
            if (dgv.GetClipboardContent() != null)
            {
                Clipboard.SetDataObject(dgv.GetClipboardContent());
                Clipboard.GetData(DataFormats.Text);
                IDataObject dt = Clipboard.GetDataObject();
                if (dt.GetDataPresent(typeof(string)))
                {
                    string tb = (string)(dt.GetData(typeof(string)));
                    Encoding encoding = Encoding.GetEncoding(1251);
                    byte[] dataStr = encoding.GetBytes(tb);
                    Clipboard.SetDataObject(encoding.GetString(dataStr));
                }
            }
            foreach (DataGridViewRow dgvr in dgv.SelectedRows)
            {
                dgv.Rows.Remove(dgvr);
            }
            dgv.ClipboardCopyMode = DataGridViewClipboardCopyMode.Disable;
        }

        private void PasteSelectedRowsFromClipboard(object sender, KeyEventArgs e)
        {
            //if user clicked Shift+Ins or Ctrl+V (paste from clipboard)
                char[] rowSplitter = { '\r', '\n' };
                char[] columnSplitter = { '\t' };

                //get the text from clipboard
                IDataObject dataInClipboard = Clipboard.GetDataObject();
                string stringInClipboard = (string)dataInClipboard.GetData(DataFormats.Text);

                //split it into lines
                string[] rowsInClipboard = stringInClipboard.Split(rowSplitter, StringSplitOptions.RemoveEmptyEntries);

                //get the row and column of selected cell in grid
                int r = dgv.SelectedCells[0].RowIndex;
                int c = dgv.SelectedCells[0].ColumnIndex;

                //add rows into grid to fit clipboard lines
                //if (dataGridView1.Rows.Count < (r + rowsInClipboard.Length))             
                //    dataGridView1.Rows.Add(r + rowsInClipboard.Length - dataGridView1.Rows.Count);
                r += 1;

                dgv.Rows.Insert(r, rowsInClipboard.Length);

                // loop through the lines, split them into cells and place the values in the corresponding cell.
                for (int iRow = 0; iRow < rowsInClipboard.Length; iRow++)
                {
                    //split row into cell values
                    string[] valuesInRow = rowsInClipboard[iRow].Split(columnSplitter);

                    //cycle through cell values
                    for (int iCol = 0; iCol < valuesInRow.Length; iCol++)
                    {
                        //assign cell value, only if it within columns of the grid
                        if (dgv.ColumnCount - 1 >= c + iCol)
                            dgv.Rows[r + iRow].Cells[c + iCol].Value = valuesInRow[iCol];
                    }
                }
        }

        private void dataGridView1_RowHeaderMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            point.X = e.ColumnIndex;
            point.Y = e.RowIndex;
        }

        private DataGridView CopyDataGridView(DataGridView dgv_org, int start, int end)
        {
            DataGridView dgv_copy = new DataGridView();
            try
            {
                if (dgv_copy.Columns.Count == 0)
                {
                    foreach (DataGridViewColumn dgvc in dgv_org.Columns)
                    {
                        dgv_copy.Columns.Add(dgvc.Clone() as DataGridViewColumn);
                    }
                }

                DataGridViewRow row = new DataGridViewRow();

                for (int i = start; i < end; i++)
                {
                    row = (DataGridViewRow)dgv_org.Rows[i].Clone();
                    int intColIndex = 0;
                    foreach (DataGridViewCell cell in dgv_org.Rows[i].Cells)
                    {
                        row.Cells[intColIndex].Value = cell.Value;
                        intColIndex++;
                    }
                    dgv_copy.Rows.Add(row);
                }
                dgv_copy.AllowUserToAddRows = false;
                dgv_copy.Refresh();

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
            return dgv_copy;
        }

        private void addSS(string name, Inventor.Point2d insPt)
        {
            try 
            {
                foreach (SketchedSymbol ss in m_DrwDoc.ActiveSheet.SketchedSymbols)
                {
                    if (ss.Name.StartsWith(name))
                    {
                        return;
                    }
                }
                m_DrwDoc.ActiveSheet.SketchedSymbols.Add(m_DrwDoc.SketchedSymbolDefinitions[name], insPt);
            }
                catch(Exception)
            {

                }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //drawTable("Spec",m_TG.CreatePoint2d(0, 0), new double[] { 0.6, 0.6, 0.8, 7, 6.3, 1, 2.2}, dataGridView1.Rows.Count - 1,
            //    intArr2: new int[] {3,4,6}, _dgv: dataGridView1,heigthRow: 0.8);

            tbl = tbl ?? new TableInv(m_DrwDoc, @"C:\ProgramData\Autodesk\Inventor Addins\Sequence.xml");
            if (tbl.countPerf != 0) tbl.saveInAsm();
            int[] arr = txtBox.Text.Split(';').Select(s => int.Parse(s)).ToArray();
            VariableDataForSpec vd = new VariableDataForSpec((Document)m_AsmDoc);
            vd.AttribDelete("Split");
            if (txtBox.Text != "") vd.AttribAdd<string>("Split", txtBox.Text, ValueTypeEnum.kStringType);
            tbl.addToSheet("Spec",arr);
            //tbl.addTable(dgv);
            //tbl.drawTable("Спецификация", 0, dataGridView1.RowCount, m_TG.CreatePoint2d());
            //string[] tmp = txtBox.Text.Split(new char[]{';'});
            //int[] spl = new int[tmp.Count()/*+1*/]; int sum = 1;
            //foreach (Inventor.SketchedSymbol ss in m_DrwDoc.ActiveSheet.SketchedSymbols)
            //{
            //    if (ss.Name.StartsWith("Spec"))
            //    {
            //        pts.Add(ss.Position);
            //        ss.Delete();
            //    }
            //}
            //for (int i = 0; i < tmp.Count(); i++)
            //{
            //        sum += Convert.ToInt16(tmp[i]);
            //    if (sum > dataGridView1.RowCount) {
            //        spl[i] = dataGridView1.RowCount;
            //        Array.Resize(ref spl, i+1);
            //        break;
            //        }
            //        spl[i] = sum;
            //}
            ////spl[spl.Count()-1] = 100;
            //Inventor.Application invApp = (Inventor.Application)m_DrwDoc.Parent;
            //invApp.ScreenUpdating = false;
            //drawTables(dataGridView1, spl);
            //invApp.ScreenUpdating = true;
        }

        private void Variable_Click(object sender, System.EventArgs e)
        {
            //string num; string var; string name;
            //if (txtBox1.Text != "")
            //{
            //    name = (string)m_AsmDoc.PropertySets[3][2].Value;
            //    CMSD("Переменные данные для исполнений:", "p5.1", true);
            //    for (int i = 0; i < Convert.ToInt16(txtBox1.Text)+1; i++)
            //    {
            //        var = name + i.ToString("-0#");
            //        if (i == 0) var = name;
            //        num = "p5.1." + i;
            //        CMSD(var, num, true);
            //    }
            //    dgv.Sort(dgv.Columns[7], ListSortDirection.Ascending);
            //}
            if (dgvVar.Visible == false) { dgvVar.Visible = true; label2.Visible = true; }


            if (partNumberAsm[partNumberAsm.Length - 3] == '-')
            {
                start = int.Parse(partNumberAsm.Substring(partNumberAsm.Length - 2));
                partNumberAsm = partNumberAsm.Substring(0, partNumberAsm.Length - 3);
            }

            if (tbl.countPerf == 0) 
            {
                if (start == 0)
                {
                    tbl.addPerf(partNumberAsm);
                    tbl.addDGVR(dgvVar, partNumberAsm);
                }
                else
                {
                    tbl.addPerf(partNumberAsm + "-" + String.Format("{0:00}", tbl.countPerf + start));
                    tbl.addDGVR(dgvVar, partNumberAsm + "-" + String.Format("{0:00}", tbl.countPerf + start - 1));
                }
            }
            else
            {
                
                tbl.addPerf(partNumberAsm + "-" + String.Format("{0:00}", tbl.countPerf + start));
                tbl.addDGVR(dgvVar, partNumberAsm + "-" + String.Format("{0:00}", tbl.countPerf + start - 1));
            }
        }

        private void renumber_Click(object sender, System.EventArgs e)
        {
            renumber(dgv);
        }

        private void Add_Doc(string str1,string str2, string str3, string key)
        {
            dgv.Rows.Insert(dgv.CurrentRow.Index, new string[] { "", "", "", str1, str2, str3, "", key });
        }

        private void sort_Click(object sender, System.EventArgs e)
        {
            dgv.Sort(dgv.Columns[7], ListSortDirection.Ascending); 
        }
        private void Add_Click(object sender, System.EventArgs e)
        {
            string str1 = dgv.CurrentRow.Cells[7].Value.ToString();
            string str2 = dgv[7, dgv.CurrentRow.Index + 1].Value.ToString();
            string val = "";
            DataGridViewRow dgvr = dgv.CurrentRow;
            val = compare(str1, str2);
            if (val != "")
            {
                //dataGridView1.Rows.Insert( dgvr.Index +1,dgvr.Cells[0].Value.ToString(), dgvr.Cells[1].Value.ToString(), dgvr.Cells[2].Value.ToString(),
                //    dgvr.Cells[3].Value.ToString(), dgvr.Cells[4].Value.ToString(), dgvr.Cells[5].Value.ToString(), dgvr.Cells[6].Value.ToString(), val);
                dgv.Rows.Insert(dgvr.Index + 1, "","", "", "", "", "", "", val);
            }
        }
        public string compare(string str1, string str2)
        {
            string val = "";
            if (str1 != str2)
            {
                for (int i = 0; i < str1.Count(); i++)
                {
                    if (str1.Count() > str2.Count()) break;
                    if (str1.Count() == 1) { val = str1 + "."; break; }
                    if (str1[i] < str2[i] && !(str1[i] == (str2[i]-1)))
                    {
                        val = str1.Substring(0, i) + (char)(str1[i]+1);
                        break;
                    }
                }
                if (val == "") val = str1 + (char)(str1[str1.Count()-1] + 1);
            }
            return val;
        }
        private void subVariable_Click(object sender, System.EventArgs e)
        {
            DataGridViewRow dgvr;
            dgvr = dgv.CurrentRow;
                tbl.addVarRows(dgvr.Cells[3].Value.ToString(), dgvr.Cells[4].Value.ToString(), tbl.countPerf);
            dgv.Rows.Remove(dgvr);
            tbl.showTable(dgv, 0, 99, tbl.group,varData:false);
            tbl.showTable(dgvVar, 100, 5000, tbl.group,varData:true);
            balloon();
        }

        private void CMSD(string str1, string key, bool even = false)
        {
            if (even)
            {
            dgv.Rows.Insert(dgv.CurrentRow.Index, new string[] { "", "", "", "", str1, "", "", key });
            dgv[4, dgv.CurrentRow.Index - 1].Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgv[4, dgv.CurrentRow.Index - 1].Style.Font = new Font(dgv.Font, FontStyle.Underline);
            }
            else
            {
                dgv.Rows.Insert(0, new string[] { "", "", "", "", str1, "", "", key });
                dgv[4, 0].Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                dgv[4, 0].Style.Font = new Font(dgv.Font, FontStyle.Underline);
            }
        }

        private void dgvVar_CellBeginEdit(object sender, DataGridViewCellCancelEventArgs e)
        {

        }

        private void dgvVar_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            TableRow tr = tbl.rows.First(r => r.dvgRow == e.RowIndex && r.group > 100);
            strs = new string[7];
            DataGridViewCellCollection Cells = dgvVar.Rows[e.RowIndex].Cells;

            for (int i = 0; i < strs.Length; i++)
            {
                //if (tr.cells[i].value != Cells[i].Value.ToString())
                if (Cells[i].Value != null)
                tr.cells[i].value = Cells[i].Value.ToString();
            }
        }

        private void dgvVar_RowsRemoved(object sender, DataGridViewRowsRemovedEventArgs e)
        {
            TableRow tr; 
            if (e.RowCount == 1 && e.RowIndex > 2)
            {   
                tr = tbl.rows.First(r => r.dvgRow == e.RowIndex && r.group > 100);
                tbl.rows.Remove(tr);
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (dgvVar.Visible == true && tbl.countPerf == 0) { dgvVar.Visible = false; label2.Visible = false; }
            if (tbl.countPerf > 0)
            {
                tbl.removePerf();
                tbl.removeRows(tbl.countPerf +1 , tbl.countPerf +2);
                tbl.showTable(dgv, 0, 100, tbl.group,varData:false);
                tbl.showTable(dgvVar, 100, 5000, tbl.group, varData: true);
            }
        }

        private void экспортВXMLToolStripMenuItem_Click(object sender, EventArgs e)
        {
            InvDoc.InvDocument<AssemblyDocument> iDoc = new InvDoc.InvDocument<AssemblyDocument>(m_AsmDoc);
            iDoc.doc = (Document)m_AsmDoc;
            string name = iDoc.path + iDoc.getProp("Part Number").Value + "(variable).xml";
            tbl.saveDataInXML(new XDocument(),"Variable",name,true,false);   
        }

        private void импортИзXMLToolStripMenuItem_Click(object sender, EventArgs e)
        {
            InvDoc.InvDocument<AssemblyDocument> iDoc = new InvDoc.InvDocument<AssemblyDocument>(m_AsmDoc);
            string n = "";
            tbl.loadDataFromXML(iDoc.path,true, ref n);
            tbl.saveInAsm();
            //dgvVar.Rows.Clear();
            //tbl.showTable(dgvVar, 100, 5000, tbl.group, varData: true);
        }

        private void dgv_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            TableRow tr = tbl.rows.First(r => r.dvgRow == e.RowIndex && r.group < 100);
            strs = new string[7];
            DataGridViewCellCollection Cells = dgv.Rows[e.RowIndex].Cells;

            for (int i = 0; i < strs.Length; i++)
            {
                //if (tr.cells[i].value != Cells[i].Value.ToString())
                if (Cells[i].Value != null)
                tr.cells[i].value = Cells[i].Value.ToString();
            }
        }

        private void dgv_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Delete && ((DataGridView)sender).SelectedCells.Count != 0)
            {
                foreach (DataGridViewCell c in ((DataGridView)sender).SelectedCells)
                {
                    c.Value = "";
                }
            }
            //if (e.Control && e.KeyCode == Keys.S)
            //{
            //    AttribAdd(this._dgv, this.textBox2.Text);
            //}
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
                Clipboard.SetText(value);
            }
            if (e.KeyCode == Keys.V && ((DataGridView)sender).SelectedCells.Count != 0)
            {
                value = Clipboard.GetText();
                if (value != ""/*.IndexOf(';') == -1*/)
                {
                    foreach (DataGridViewCell c in ((DataGridView)sender).SelectedCells)
                    {
                        c.Value = Clipboard.GetText();
                    }
                }
                //else
                //{
                //    string[] spl = value.Split(';');
                //    int j = 0;

                //    for (int i = ((DataGridView)sender).SelectedCells.Count-1; i >=0 ; i--)
                //    {
                //        ((DataGridView)sender).SelectedCells[i].Value = spl[j];
                //            j++;
                //            if (j == (spl.Count() - 1)) j = 0;
                //    }
                //}
            }
        }

        private void сохранитьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            saveData(m_AsmDoc);
        }

        public static void saveData(AssemblyDocument m_AsmDoc)
        {
            XDocument xd = new XDocument();
            InvDoc.InvDocument<AssemblyDocument> iDoc = new InvDoc.InvDocument<AssemblyDocument>(m_AsmDoc);
            iDoc.doc = (Document)m_AsmDoc;
            string name = iDoc.path + iDoc.getProp("Part Number").Value + ".xml";
            //tbl.saveDataInXML(xd,"Data",name, false);
            BOM m_BOM = m_AsmDoc.ComponentDefinition.BOM;
            string partNumberAsm = m_AsmDoc.PropertySets[3][2].Value.ToString();
            m_BOM.StructuredViewEnabled = true;
            // if (!m_BOM.StructuredViewFirstLevelOnly)
            // m_BOM.StructuredViewFirstLevelOnly = true;
            if (m_BOM.StructuredViewFirstLevelOnly)
                m_BOM.StructuredViewFirstLevelOnly = false;
            BOMView m_BOMView = m_BOM.BOMViews["Структурированный"];
            TableInv tbl = null;
            tbl = new TableInv(m_AsmDoc, @"C:\ProgramData\Autodesk\Inventor Addins\Sequence.xml");
            tbl.addTable(m_BOMView);
            tbl.saveDataInXML(xd, "Data", name, false);
        }

        private void открытьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            InvDoc.InvDocument<AssemblyDocument> iDoc = new InvDoc.InvDocument<AssemblyDocument>(m_AsmDoc);
            string n = "";
            tbl.loadDataFromXML(iDoc.path,false, ref n);
        }

        private void загрузитьДанныеToolStripMenuItem_Click_1(object sender, EventArgs e)
        {
            DataGridViewRow dgvr;
            InvDoc.InvDocument<AssemblyDocument> iDoc = new InvDoc.InvDocument<AssemblyDocument>(m_AsmDoc);
            iDoc.doc = (Document)m_AsmDoc;
            dgvr = dgvVar.CurrentRow;
            TableRow tr = tbl.rows.OfType<TableRow>().FirstOrDefault(t => t.group > 100 && t.dvgRow == dgvr.Index);
            if (tr == null) return;
            //OpenFileDialog ofd = new OpenFileDialog();
            //ofd.InitialDirectory = iDoc.path;
            //ofd.Filter = "(*.iam, *.ipt)|*.iam;*.ipt";
            //ofd.ShowDialog();
            string fileName = InvDoc.u.OFD(iDoc.path, "(*.iam, *.ipt)|*.iam;*.ipt");
            //string fileName = ofd.FileName;
            Document doc = Macros.StandardAddInServer.m_inventorApplication.Documents.Open(fileName, false);
            Dictionary<string, string> dic = new Dictionary<string, string>();
            dic = TableInv.propDoc(doc);
            try
            {
            tr.cells[3].value = dic["PartNumber"];
            tr.cells[4].value = dic["desc"];
            tr.cells[6].value = dic["note"];
            tr.cells[7].value = dic["dxf"];
            tr.cells[8].value = dic["thickness"];
            tr.cells[9].value = dic["FPLength"];
            tr.cells[10].value = dic["FPWidth"];
            }
            catch (Exception)
            {
            }
            doc.Close();
            tbl.showTable(dgvVar, 100, 5000, tbl.group, varData: true);
            //tbl.addVarRows(dgvr.Cells[3].Value.ToString(), dgvr.Cells[4].Value.ToString(), tbl.countPerf);
            //dgv.Rows.Remove(dgvr);
            //tbl.showTable(dgv, 0, 99, tbl.group, varData: false);
            //tbl.showTable(dgvVar, 100, 5000, tbl.groupVar, varData: true);
            //balloon();
        }
    }
}
