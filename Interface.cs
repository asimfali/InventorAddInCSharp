using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Xml.Linq;
using System.Drawing;

namespace InterfaceDll
{
    public class MyLabel
    {
        Label lbl;
        public Label Lbl
        {
            get { return lbl; }
            set { lbl = value; }
        }
        public List<Label> lbls = new List<Label>(10);
        string name = "Label";
        int count = 1;

        public Label addLabel(string txt, System.Drawing.Point loc, int width, int heigth)
        {
            this.lbl = new System.Windows.Forms.Label();
            this.lbl.AutoSize = true;
            this.lbl.Location = loc;
            this.lbl.Name = name + count; count++;
            this.lbl.Size = new System.Drawing.Size(width, heigth);
            this.lbl.TabIndex = 0;
            this.lbl.Text = txt;
            return lbl;
        }
    }
    public class MyCheckBox
    {
        CheckBox checkbox;
        public CheckBox Checkbox
        {
            get { return checkbox; }
            set { checkbox = value; }
        }
        
        public List<CheckBox> lbls = new List<CheckBox>(10);
        string name = "CheckBox";
        int count = 1;

        public CheckBox addCheckBox(string txt, System.Drawing.Point loc, int width, int heigth)
        {
            this.checkbox = new CheckBox();
            this.checkbox.AutoSize = true;
            this.checkbox.Location = loc;
            this.checkbox.Name = name + count; count++;
            this.checkbox.Size = new System.Drawing.Size(width, heigth);
            this.checkbox.TabIndex = 0;
            this.checkbox.Text = txt;
            return checkbox;
        }
    }
    public class MyDGV : IDisposable
    {
        DataGridView dgv;
        public DataGridView Dgv
        {
            get { return dgv; }
            set { dgv = value; }
        }
        DataGridViewTextBoxColumn dgvCol;
        DataGridViewComboBoxColumn dgvComboBoxCol;

        public DataGridViewTextBoxColumn DgvCol
        {
            get { return dgvCol; }
            set { dgvCol = value; }
        }
        public XElement node;
        public List<XElement> nodes = new List<XElement>(10);
        public List<DataGridView> dgvs = new List<DataGridView>(10);
        public List<DataGridViewTextBoxColumn> dgvCols = new List<DataGridViewTextBoxColumn>(20);
        public List<DataGridViewComboBoxColumn> dgvCBCols = new List<DataGridViewComboBoxColumn>(20);
        public changeData change;
        public List<changeData> changes = new List<changeData>();
        string name = "DGV";
        string[] names;
        int count = 1, countDGVC = 1;

        public void Dispose()
        {
            foreach (var item in dgvCols)
            {
                item.Dispose();
            }
            foreach (var item in dgvs)
            {
                item.Dispose();
            }
        }
        public MyDGV() { }
        public DataGridView addDGV(System.Drawing.Point loc, int width, int heigth, XElement node, Dictionary<string,string> colHeaders,float[] weigth, string type = "TB")
        {
            dgv = new DataGridView();

            this.dgv.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            //this.dgv.Columns.AddRange(new System.Windows.Forms.DataGridViewTextBoxColumn[] {
            //addDGVCol("Название",(float)0.3),addDGVCol("Id",(float)0.1),addDGVCol("Категория",(float)0.6)});
            int i=0;
            names = new string[colHeaders.Count];
            if (type == "TB")
            foreach (var item in colHeaders)
            {
                this.dgv.Columns.Add(addDGVCol(item.Key, item.Value, weigth[i]));
                names[i] = item.Key;
                i++;
            }
            else if (type == "CB")
            {
                foreach (var item in colHeaders)
                {
                    this.dgv.Columns.Add(addDGVCBCol(item.Key, item.Value, weigth[i]));
                    names[i] = item.Key;
                    i++;
                }
            }
            this.dgv.Location = loc;
            this.dgv.Name = name + count; count++;
            this.dgv.Size = new System.Drawing.Size(width, heigth);
            this.dgv.TabIndex = 1;
            this.node = node;
            dgvs.Add(dgv); nodes.Add(node);
            //this.dgv.RowsRemoved += new DataGridViewRowsRemovedEventHandler(deleteRow);
            this.dgv.CellBeginEdit += new DataGridViewCellCancelEventHandler(changeCellBefore);
            this.dgv.CellValueChanged += new DataGridViewCellEventHandler(changeCell);
            this.dgv.KeyDown += new KeyEventHandler(keyDown);
            i = 0;
            if (node != null)
            foreach (var value in node.Elements())
            {
                //value.Add(new XAttribute("Category", node.FirstAttribute.Value));
                //dgv.Rows.Add(new string[] { value.Attribute(co).Value, value.Attribute(attId).Value, value.Attribute("Category").Value });
                addRow(value, colHeaders);
            }
            return dgv;
         }

        public void deleteRow(object sender, DataGridViewRowsRemovedEventArgs e)
        {
            DataGridView tmpDgv = (DataGridView)sender;
            cell(tmpDgv, -1, e.RowIndex, true);
        }

        public void addRow(XElement el, Dictionary<string, string> colHeaders)
        {
            string [] vals = new string[colHeaders.Count];
            int i=0;
            foreach (var item in colHeaders)
	            {
                   /* if (el.Attribute(item.Key).Value == null) vals[i] = "";*/
		           vals[i] = el.Attribute(item.Key).Value;
                   i++;
	            }
            dgv.Rows.Add(vals);
        }

        public void changeVal(XElement el,string name, string val)
        {
            el.Attribute(name).Value = val;
        }

        public void changeVal(XElement el, string[] names, string[] vals)
        {
            for (int i = 0; i < vals.Length; i++)
            {
                if (vals[i] == null) continue;
                changeVal(el, names[i], vals[i]);
            }
        }
        public XElement findElem(XElement el, string[] names, string[] vals)
        {
            bool flag = false;
            XElement newEl = el.Descendants().FirstOrDefault(n => n.Attribute(names[0]).Value == vals[0]);
            if (newEl != null)
            {
                for (int i = 1; i < names.Length; i++)
                {
                    XAttribute attr = newEl.Attribute(names[i]);
                    if (attr == null) { flag = false; break; }
                    if (attr.Value == vals[i]) flag = true;
                }
            }
            if (flag) return newEl;
            else return null;
        }
        public void addElem(XElement el, string[] names, string[] vals)
        {
            XElement elem = new XElement("Value");
            for (int i = 0; i < vals.Length; i++)
            {
                //if (vals[i] == null) elem.Add(new XAttribute(names[i], ""));
                //else 
                    elem.Add(new XAttribute(names[i], vals[i])); 
            }
            el.Add(elem);
        }

        public void safe()
        {
            foreach (var ch in changes)
            {
                var ie = findElem(ch.el, names, ch.oldVal);
                if (ie != null)
                {
                    changeVal(ie, names, ch.val);
                }
                else
                {
                    addElem(ch.el, names, ch.val);
                }
            }
        }

        public void changeCellBefore(object sender, EventArgs e)
        {
            dgv = (DataGridView)sender;
            DataGridViewCellCancelEventArgs arg = (DataGridViewCellCancelEventArgs)e;
            cell(dgv, arg.ColumnIndex, arg.RowIndex, true);
        }

        public void cell(DataGridView dgv, int colInd, int rowInd, bool flag)
        {
            var ie = changes.FirstOrDefault(ch => ch.rowInd == rowInd);
            if (ie != null)
            {
                change = ie;
                return;
            }
            else
            {
                change = new changeData(dgv.ColumnCount);
                changes.Add(change);
            }
            change.colInd = colInd; change.rowInd = rowInd;
            if (colInd == -1)
            {
                int i = 0;
                foreach (DataGridViewCell item in dgv.Rows[rowInd].Cells)
                {
                    if (item.Value == null) change.oldVal[i] = "";
                    else change.oldVal[i] = item.Value.ToString();
                    i++;
                }
                return;
            }
            if (dgv[colInd, rowInd].Value != null)
            {
                int i = 0;
                foreach (DataGridViewCell item in dgv.Rows[rowInd].Cells)
                {
                    change.oldVal[i] = item.Value.ToString();
                    i++;
                }
            }
        }

        public void changeCell(object sender, EventArgs e)
        {
            dgv = (DataGridView)sender;
            DataGridViewCellEventArgs arg = (DataGridViewCellEventArgs)e;
            if (change == null) return;
            int ind = dgvs.FindIndex(dg => dg.Equals(dgv));
            change.el = nodes[ind];
            if (dgv[arg.ColumnIndex, arg.RowIndex].Value != null)
                change.val[arg.ColumnIndex] = dgv[arg.ColumnIndex, arg.RowIndex].Value.ToString();
            else change.val[arg.ColumnIndex] = "";
        }

        public void keyDown(object sender, KeyEventArgs e)
        {
            DataGridView tmpDgv = (DataGridView)sender;
            string txt = "";
            if (e.Control && e.KeyCode == Keys.C)
            {
                foreach (DataGridViewTextBoxCell item in tmpDgv.SelectedCells)
	            {
		            txt += item.Value.ToString();
	            }
                Clipboard.SetText(txt);
            }
            if (e.Control && e.KeyCode == Keys.S)
            {
               
            }
            if (e.KeyCode == Keys.Delete)
            {
                foreach (DataGridViewTextBoxCell item in tmpDgv.SelectedCells)
                {
                    item.Value = "";
                }
            }
            if (e.Control && e.KeyCode == Keys.V)
            {
                char[] rowSplitter = { '\r', '\n' };
                char[] columnSplitter = { '\t' };

                //get the text from clipboard
                IDataObject dataInClipboard = Clipboard.GetDataObject();
                string stringInClipboard = (string)dataInClipboard.GetData(DataFormats.Text);

                //split it into lines
                string[] rowsInClipboard = stringInClipboard.Split(rowSplitter, StringSplitOptions.RemoveEmptyEntries);

                //get the row and column of selected cell in grid
                int r = tmpDgv.SelectedCells[0].RowIndex;
                int c = tmpDgv.SelectedCells[0].ColumnIndex;

                //add rows into grid to fit clipboard lines
                if (tmpDgv.Rows.Count < (r + rowsInClipboard.Length))
                    tmpDgv.Rows.Add(r + rowsInClipboard.Length - tmpDgv.Rows.Count);

                // loop through the lines, split them into cells and place the values in the corresponding cell.
                for (int iRow = 0; iRow < rowsInClipboard.Length; iRow++)
                {
                    //split row into cell values
                    string[] valuesInRow = rowsInClipboard[iRow].Split(columnSplitter);

                    //cycle through cell values
                    for (int iCol = 0; iCol < valuesInRow.Length; iCol++)
                    {
                        //assign cell value, only if it within columns of the grid
                        if (tmpDgv.ColumnCount - 1 >= c + iCol)
                        {
                            cell(tmpDgv, c + iCol, r + iRow, true);
                            tmpDgv.Rows[r + iRow].Cells[c + iCol].Value = valuesInRow[iCol];
                        }
                    }
                }
            }
        }

        public DataGridViewTextBoxColumn addDGVCol(string name,string header, float fillWidth)
        {
            dgvCol = new DataGridViewTextBoxColumn();
            //dgvCol.Width = width;
            dgvCol.Name = name + countDGVC; countDGVC++;
            dgvCol.HeaderText = header;
            dgvCol.AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            dgvCol.FillWeight = fillWidth;
            dgvCols.Add(dgvCol);
            return dgvCol;
        }

        public DataGridViewComboBoxColumn addDGVCBCol(string name, string header, float fillWidth)
        {
            dgvComboBoxCol = new DataGridViewComboBoxColumn();
            //dgvCol.Width = width;
            dgvComboBoxCol.Name = name + countDGVC; countDGVC++;
            dgvComboBoxCol.HeaderText = header;
            dgvComboBoxCol.AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            dgvComboBoxCol.FlatStyle = FlatStyle.Flat;
            dgvComboBoxCol.FillWeight = fillWidth;
            dgvCBCols.Add(dgvComboBoxCol);
            return dgvComboBoxCol;
        }

        public class changeData
        {
            public string[] val;
            public string[] oldVal;
            public XElement el;
            public int rowInd, colInd;
            public changeData(int bound) { val = new string[bound]; oldVal = new string[bound]; }
        }
    }
    public class MyComboBox
    {
        ComboBox cb;

        public ComboBox Cb
        {
            get { return cb; }
            set { cb = value; }
        }
        public List<ComboBox> cbs = new List<ComboBox>(10);
        string name = "ComboBox";
        int count = 1;

        public MyComboBox() { }
        public ComboBox addComboBox(string name, System.Drawing.Point loc, int width, int heigth, XElement node)
        {
            if (cbs.Exists(co => co.Location == loc))
            {
                cb = cbs.First(lo => lo.Location == loc);
                cb.Items.Clear();
            }
            else
            {
                this.cb = new System.Windows.Forms.ComboBox();
            }
            this.cb.FormattingEnabled = true;
            this.cb.Location = loc;
            this.cb.Name = name + count; count++;
            this.cb.Size = new System.Drawing.Size(width, heigth);
            this.cb.TabIndex = 0;
            //this.cb.SelectedIndexChanged += new EventHandler(this.cb_TextUpdate);
            var col = node.Elements().Select(e => e.FirstAttribute.Value).ToArray();
            cb.Items.AddRange(col);
            cbs.Add(this.cb);
            return cb;
        }
    }

    public class MyTextBox
    {
        TextBox txtBox;

        public TextBox TxtBox
        {
            get { return txtBox; }
            set { txtBox = value; }
        }
        public List<TextBox> txtBoxes = new List<TextBox>(10);
        string name = "TextBox";
        int count = 1;

        public MyTextBox() { }

        public TextBox addTextBox(string text, System.Drawing.Point loc, int width, int heigth)
        {
            this.txtBox = new System.Windows.Forms.TextBox();
            this.txtBox.Location = loc;
            //this.txtBox.Multiline = true;
            this.txtBox.Name = name + count; count++;
            this.txtBox.Size = new System.Drawing.Size(width, heigth);
            this.txtBox.Text = text;
            txtBoxes.Add(TxtBox);
            this.txtBox.TabIndex = 1;
            return txtBox;
        }
    }
    public class brushes
    {
        public int ind;
        public System.Windows.Forms.ListBox lbInd;
        public Brush br;
        public brushes(int ind, System.Windows.Forms.ListBox lbInd, Brush br) { this.ind = ind; this.lbInd = lbInd; this.br = br; }
    }

    public class MyListBox
    {
        ListBox lb;

        public ListBox Lb
        {
            get { return lb; }
            set { lb = value; }
        }
        public List<ListBox> lbs = new List<ListBox>(10);
        string name = "ListBox";
        int count = 1;
        Brush myBrush = Brushes.Black;
        //string elemName;
        //Dictionary<int, Brush> brushes = new Dictionary<int, Brush>(20);
        List<brushes> lstBruches = new List<brushes>(20);

        public MyListBox() { }

        public ListBox addLstBox(System.Drawing.Point loc, int width, int heigth, XElement node, string attName, string attId)
        {
            lb = new System.Windows.Forms.ListBox();
            lb.DrawMode = DrawMode.OwnerDrawVariable;
            lb.DrawItem += new DrawItemEventHandler(drawItem);
            lb.BackColor = Color.Cornsilk;
            this.lb.FormattingEnabled = true;
            this.lb.Location = loc;
            this.lb.Name = name + count; count++;
            this.lb.Size = new System.Drawing.Size(width, heigth);
            this.lb.TabIndex = 1;
            this.lb.SelectionMode = SelectionMode.MultiExtended;
            //lb.KeyDown += new KeyEventHandler(keyDown);
            //lb.ContextMenuStrip = this.contextMenuStrip1;
            foreach (var value in node.Elements())
            {
                switch (value.Name.ToString())
                {
                    case "Category":
                        myBrush = Brushes.Green;

                        break;
                    case "Element":
                        myBrush = Brushes.Red;
                        break;
                    default:
                        break;
                }
                int ind;
                if (value.Attribute("Articul") == null || value.Attribute("Articul").Value == "")
                    ind = lb.Items.Add(value.Attribute(attName).Value);
                else ind = lb.Items.Add(value.Attribute(attName).Value + "   " + value.Attribute("Articul").Value);
                lstBruches.Add(new brushes(ind, lb, myBrush));
            }
            //brushes.Clear();
            this.lbs.Add(lb);
            //this.lstBruches.Add(brushes);
            return lb;
        }
        private void drawItem(object sender, System.Windows.Forms.DrawItemEventArgs e)
        {
            e.DrawBackground();
            var br = lstBruches.Find(l => l.lbInd.Equals(sender) && l.ind == e.Index);
            if (br != null)
            e.Graphics.DrawString(((ListBox)sender).Items[e.Index].ToString(), e.Font, br.br, e.Bounds, StringFormat.GenericDefault);
            e.DrawFocusRectangle();
        }
    }
    public class MyToolStrip
    {
        ToolStripMenuItem tsmi;
        public List<ToolStripMenuItem> tsmis = new List<ToolStripMenuItem>(10);
        string name = "ListBox";
        int count = 1;

        public ToolStripMenuItem Tsmi
        {
            get { return tsmi; }
            set { tsmi = value; }
        }
        public ToolStripMenuItem addToolStrip(string text)
        {
            tsmi = new ToolStripMenuItem();
            tsmi.Name = name + count; count++;
            tsmi.AutoSize = true;
            tsmi.Text = text;
            tsmis.Add(tsmi);
            //tsmi.Click += new EventHandler(addElem);
            return tsmi;
        }
    }
    public class MyButton
    {
        System.Windows.Forms.Button btn;

        internal System.Windows.Forms.Button Btn
        {
            get { return btn; }
            set { btn = value; }
        }
        public List<System.Windows.Forms.Button> btns = new List<System.Windows.Forms.Button>(10);
        string name = "Button";
        int count = 1;
        public System.Windows.Forms.Button addButton(string txt, System.Drawing.Point loc, int width, int height)
        {
            btn = new System.Windows.Forms.Button();
            this.btn.Location = loc;
            this.btn.Name = name + count; count++;
            this.btn.Size = new System.Drawing.Size(width, height);
            this.btn.TabIndex = 1;
            this.btn.Text = txt;
            this.btn.UseVisualStyleBackColor = true;
            btns.Add(btn);
            return btn;
        }
    }
}
