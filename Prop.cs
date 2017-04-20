using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Xml.Linq;
using System.Text;
using System.Windows.Forms;
using Inventor;
using TableDll;
using InterfaceDll;
using System.Text.RegularExpressions;
using ut = InvDoc.u;
using InvDoc;

namespace InvAddIn
{
    public partial class Prop : Form
    {
        public Inventor.AssemblyDocument m_AsmDoc;
        public Inventor.Application m_InvApp;
        public Inventor.ComponentDefinition oCompDef;
        public Inventor.Document oDoc;
        public Inventor.DrawingDocument m_DrwDoc;
        private Inventor.BOM m_BOM;
        private Inventor.BOMRowsEnumerator m_BOMRowEnum;
        private Inventor.BOMRow m_BomRow;
        private Inventor.BOMView m_BOMView;
        private Inventor.TransientGeometry m_TG;
        private Inventor.SketchedSymbolDefinition m_SketchDef;
        private Inventor.DrawingSketch m_DrwSketch;
        private Inventor.SketchLine m_SketchLine;
        private MyDGV myDGV = new MyDGV();
        private bool m_first;
        private int num;
        private int offset = 0;
        private string path = "", value = "", filePath;
        public FormTreeView ftv;
        InvDoc.XML props, ex, sort;
        //private List<Inventor.Point2d> pts = new List<Point2d>();
        List<Inventor.Document> objs = new List<Document>();
        List<DataGridViewCell> cellsChanged = new List<DataGridViewCell>();
        System.Collections.Generic.List<int> lst = new List<int>();
        System.Collections.Generic.List<string> match = new System.Collections.Generic.List<string>();
        System.Collections.Generic.List<string> result = new System.Collections.Generic.List<string>();
        System.Collections.Generic.List<string> ex_val = new System.Collections.Generic.List<string>();
        System.Collections.Generic.List<string> ex_attr = new System.Collections.Generic.List<string>();
        public System.Collections.Generic.List<string> sort_val = new System.Collections.Generic.List<string>();
        public System.Collections.Generic.List<string> sort_attr = new System.Collections.Generic.List<string>();
        //public System.Collections.Generic.List<string> itemNumber = new System.Collections.Generic.List<string>();
        //public System.Collections.Generic.List<string> sortNumber = new System.Collections.Generic.List<string>();
        public System.Collections.Generic.List<string> count = new System.Collections.Generic.List<string>();
        //public System.Collections.Generic.List<string> finding = new System.Collections.Generic.List<string>();

        //BindingSource bs;
        public static System.Drawing.Point point = new System.Drawing.Point();
        public int pos = 0;

        public Prop(Inventor.Document pDoc)
        {
            m_InvApp = (Inventor.Application)pDoc.Parent;
            //m_BOMView.Sort("Default BOM Structure", true, "Component Type", true, "Description", true);
            m_first = true;
            InitializeComponent();
            Rectangle bns = Screen.PrimaryScreen.Bounds;
            this.Bounds = bns;
            this.WindowState = FormWindowState.Maximized;
            bns.Height = bns.Height - bns.Height / 12 * 2;
            bns.Y = menuStrip1.Size.Height;
            dataGridView1.Bounds = bns;
            //dataGridView1.Visible = false;
            myDGV.Dgv = dataGridView1;
            part.Location = new System.Drawing.Point(10, dataGridView1.Bounds.Y + dataGridView1.Bounds.Height + 10);
            dec.Location = new System.Drawing.Point(part.Bounds.X + part.Bounds.Width + 10, dataGridView1.Bounds.Y + dataGridView1.Bounds.Height + 10);
            bom.Location = new System.Drawing.Point(dec.Bounds.X + dec.Bounds.Width + 10, dataGridView1.Bounds.Y + dataGridView1.Bounds.Height + 10);
            //checkBox3.Location = new System.Drawing.Point(checkBox2.Bounds.X + checkBox2.Bounds.Width + 10, dataGridView1.Bounds.Y + dataGridView1.Bounds.Height + 10);
            //label1.Location = new System.Drawing.Point(checkBox3.Bounds.X + checkBox3.Bounds.Width + 10, dataGridView1.Bounds.Y + dataGridView1.Bounds.Height + 10);
            //textBox1.Width = 200;
            //textBox1.Location = new System.Drawing.Point(label1.Bounds.X + label1.Bounds.Width + 1, dataGridView1.Bounds.Y + dataGridView1.Bounds.Height + 10);
            //button1.Location = new System.Drawing.Point(textBox1.Bounds.X + textBox1.Bounds.Width + 1, dataGridView1.Bounds.Y + dataGridView1.Bounds.Height + 10);

            //bns.Y = 0;
            //bns.Height = menuStrip1.Size.Height;
            //menuStrip1.Bounds = bns;
        }

        private void initializeProp(string filePath)
        {
            m_BOM.StructuredViewEnabled = true;
            if (m_BOM.StructuredViewFirstLevelOnly)
                m_BOM.StructuredViewFirstLevelOnly = false;
            m_BOMView = m_BOM.BOMViews["Структурированный"];
            if (m_first)
            {
                props = new InvDoc.XML(filePath);
                ex = new InvDoc.XML(@"C:\ProgramData\Autodesk\Inventor Addins\Exceptions.xml");
                props.ReadXML("Properties", ref match, ref result);
                ex.ReadXML("Exceptions", ref ex_val, ref ex_attr);
                sort = new InvDoc.XML(@"C:\ProgramData\Autodesk\Inventor Addins\Sequence.xml");
                sort.ReadXML("Sequence", ref sort_val, ref sort_attr);
                //Property p;
                num = dataGridView1.Columns.Add("Имя файла", "Имя файла");
                dataGridView1.Columns[num].Width = 300;
                for (int i = 0; i < result.Count; i++)
                {
                    string ss = props.substring(result[i], "name=");
                    string cn = props.substring(result[i], "columnName="); // cn - columnName
                    int w = Convert.ToInt16(props.substring(result[i], "width="));
                    string f = props.substring(result[i], "format=");
                    num = dataGridView1.Columns.Add(ss, cn);
                    if (f != "") dataGridView1.Columns[num].DefaultCellStyle.Format = f;
                    dataGridView1.Columns[num].Width = w;
                }
                num = dataGridView1.Columns.Add("fileName", "Файл");
                dataGridView1.Columns[num].Visible = false;
            }
            List<int> ii = new List<int>();
            num = addRow((Inventor.Document)m_AsmDoc, ref ii);
            addCells(ii, num);
            addFromBOM(m_BOMView.BOMRows);
            foreach (DataGridViewCell c in cellsChanged)
            {
                c.Style.BackColor = System.Drawing.Color.LightGray;
            }
            //myDGV.setColors();
        }
        private void initializeBOM(string filePath)
        {
            m_BOM.StructuredViewEnabled = true;
            if (m_BOM.StructuredViewFirstLevelOnly)
                m_BOM.StructuredViewFirstLevelOnly = false;
            m_BOMView = m_BOM.BOMViews["Структурированный"];
            try
            {
                m_BOMView.Sort("Keywords", true);
            }
            catch { }
            if (m_first)
            {
                props = new InvDoc.XML(filePath);
                ex = new InvDoc.XML(@"C:\ProgramData\Autodesk\Inventor Addins\Exceptions.xml");
                props.ReadXML("Properties", ref match, ref result);
                ex.ReadXML("Exceptions", ref ex_val, ref ex_attr);
                sort = new InvDoc.XML(@"C:\ProgramData\Autodesk\Inventor Addins\Sequence.xml");
                sort.ReadXML("Sequence", ref sort_val, ref sort_attr);
                //Property p;
                num = dataGridView1.Columns.Add("Poz", "Поз.");
                dataGridView1.Columns[num].Width = 100;
                for (int i = 0; i < result.Count; i++)
                {
                    string ss = props.substring(result[i], "name=");
                    string cn = props.substring(result[i], "columnName="); // cn - columnName
                    int w = Convert.ToInt16(props.substring(result[i], "width="));
                    string f = props.substring(result[i], "format=");
                    num = dataGridView1.Columns.Add(ss, cn);
                    if (f != "") dataGridView1.Columns[num].DefaultCellStyle.Format = f;
                    dataGridView1.Columns[num].Width = w;
                }
                num = dataGridView1.Columns.Add("fileName", "Файл");
                dataGridView1.Columns[num].Width = 100;
                dataGridView1.Columns[num].Visible = false;
            }
            List<int> ii = new List<int>();
            num = addRow((Inventor.Document)m_AsmDoc, ref ii);
            addCells(ii, num);
            addFromBOM(m_BOMView.BOMRows);
            //dataGridView1[0, 0].Value = "0";
            //for (int i = 0; i < itemNumber.Count; i++)
            //{
            //    dataGridView1[0, i+1].Value = itemNumber[i];
            //}
            if (m_first)
            {
                DataGridViewTextBoxColumn col = new DataGridViewTextBoxColumn();
                col.HeaderText = "Кол.";
                col.Name = "Count";
                col.Width = 25;
                dataGridView1.Columns.Insert(3, col);
            }

            int k = 0;
            if (m_first)
                dataGridView1[3, 0].Value = "1";
            else
            {
                k = num;
                dataGridView1[3, k].Value = "1";
            }
            for (int i = 0; i < count.Count; i++)
            {
                dataGridView1[3, k + i + 1].Value = count[i];
            }

            foreach (DataGridViewCell c in cellsChanged)
            {
                c.Style.BackColor = System.Drawing.Color.LightGray;
            }
            count.Clear();

        }

        private void addCells(List<int> ii, int num)
        {
            try
            {
                foreach (int i in ii)
                {
                    cellsChanged.Add(dataGridView1[i, num]);
                }
            }
            catch { }
        }

        private object addProp(Inventor.Document doc, string name, string val = "")
        {
            Property p;
            try
            {
                p = doc.PropertySets[1][name];
            }
            catch
            {
                try
                {
                    p = doc.PropertySets[3][name];
                }
                catch
                {
                    try
                    {
                        p = doc.PropertySets[4][name];
                    }
                    catch
                    {
                        p = doc.PropertySets[4].Add(val, name);
                    }
                }
            }
            //if (val != "")   
            //{
            try
            {
                if (val != p.Value.ToString())
                    p.Value = val;
            }
            catch { };
            //}
            return p.Value;
        }

        private object addProp(Inventor.ApprenticeServerDocument doc, string name, string val = "")
        {
            Property p;
            try
            {
                p = doc.PropertySets[3][name];
            }
            catch
            {
                try
                {
                    p = doc.PropertySets[4][name];
                }
                catch
                {
                    p = doc.PropertySets[4].Add(val, name);
                }
            }
            //if (val != "")   
            //{
            try
            {
                if (val != p.Value.ToString())
                    p.Value = val;
            }
            catch { };
            //}
            return p.Value;
        }

        private object getProp(Inventor.Document doc, string name)
        {
            Property p;
            try
            {
                try
                {
                    p = doc.PropertySets[1][name];
                }
                catch
                { p = doc.PropertySets[3][name]; }
            }
            catch
            {
                try
                {
                    p = doc.PropertySets[4][name];
                }
                catch
                {
                    p = doc.PropertySets[4].Add("", name);
                }
            }
            return p.Value;
        }

        private object getProp(Inventor.ApprenticeServerDocument doc, string name)
        {
            Property p;
            try
            {
                try
                {
                    p = doc.PropertySets[1][name];
                }
                catch
                { p = doc.PropertySets[3][name]; }
            }
            catch
            {
                try
                {
                    p = doc.PropertySets[4][name];
                }
                catch
                {
                    p = doc.PropertySets[4].Add("", name);
                }
            }
            return p.Value;
        }

        private int recursive(BOMRowsEnumerator rows, int num)
        {
            int count;
            foreach (BOMRow row in rows)
            {
                lst.Add(num);
                try
                {
                    count = row.ChildRows.Count;
                    recursive(row.ChildRows, num + 1);
                }
                catch
                {

                }
            }
            return 0;
        }

        private bool findInDgv(string name)
        {
            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                try
                {
                    if (row.Cells[dataGridView1.ColumnCount - 1].Value.ToString() == name)
                        return true;
                }
                catch
                {

                }
            }
            return false;
        }

        private bool findInExep(string name)
        {
            string val = ex_val.Find(delegate(string n) { return name.ToUpper().IndexOf(n.ToUpper()) != -1; });
            return val != null ? true : false;
        }

        private void addFromBOM(BOMRowsEnumerator rows)
        {
            try
            {
                foreach (Inventor.BOMRow row in rows)
                {
                    if (part.Checked == true)
                    { if (row.ReferencedFileDescriptor.FullFileName.IndexOf(path) == -1) goto nex; }

                    int n; string pad = "";
                    string[] tmp = row.ItemNumber.Split('.');
                    for (int i = 0; i < tmp.Count(); i++)
                    {
                        pad += "    ";
                    }
                    oCompDef = row.ComponentDefinitions[1];
                    oDoc = (Inventor.Document)oCompDef.Document;
                    if (bom.Checked == true && oCompDef.BOMStructure == BOMStructureEnum.kPurchasedBOMStructure) continue;
                    string name = oDoc.FullDocumentName;
                    name = name.Substring(name.LastIndexOf('\\') + 1, name.Length - 1 - name.LastIndexOf('\\'));
                    //if (part.Checked == false)
                    //{
                    if (findInExep(name)) goto nex;
                    //}
                    List<int> ii = new List<int>();
                    n = addRow(oDoc, ref ii, pad);
                    addCells(ii, n);
                    count.Add(row.ItemQuantity.ToString());
                    //string itemn = row.ItemNumber;
                    //string[] tmpstr = itemn.Split('.');
                    //for (int i = 0; i<tmpstr.Length; i++)
                    //{

                    //}
                    //itemNumber.Add(row.ItemNumber);
                    //sortNumber.Add(row.ItemNumber);
                    ii = null;

                nex:
                    if (row.ChildRows != null)
                    {
                        addFromBOM(row.ChildRows);
                    }
                }
            }
            catch
            {
                //MessageBox.Show(e.ToString());
            }
        }

        private int addRow(Inventor.Document oDoc, ref List<int> ii, string pad = "")
        {
            object[] strs;
            strs = new object[result.Count + 2];
            objs.Add(oDoc);
            if (dataGridView1.Columns["Count"] == null)
            {
                for (int i = 0; i < result.Count; i++)
                {
                    string ss = props.substring(result[i], "name=");
                    string val = props.substring(result[i], "value=");
                    object valProp = getProp(oDoc, ss);
                    if (valProp.GetType() == typeof(System.DateTime)) valProp = ((System.DateTime)valProp).ToString("dd.MM.yyyy");
                    if (val != "" && val != valProp.ToString())
                        ii.Add(i + 1);
                    strs[i + 1] = (val == "") ? getProp(oDoc, ss) : addProp(oDoc, ss, val);
                }
            }
            else
            {
                int n = 1;
                for (int i = 0; i < result.Count; i++)
                {
                    string ss = props.substring(result[i], "name=");
                    string val = props.substring(result[i], "value=");
                    object valProp = getProp(oDoc, ss);
                    if (valProp.GetType() == typeof(System.DateTime)) valProp = ((System.DateTime)valProp).ToString("dd.MM.yyyy");
                    if (val != "" && val != valProp.ToString())
                        ii.Add(i + n);
                    strs[i + n] = (val == "") ? getProp(oDoc, ss) : addProp(oDoc, ss, val);
                    if (i == 1) n++;
                }
            }
            string name = oDoc.FullFileName;
            if (!part.Checked)
            {
                if (findInDgv(name)) return 0;
            }
            else
            {
                if (findInDgv(name))
                {
                    name = name + ":" + offset.ToString();
                    offset++;
                }
            }
            strs[result.Count + 1] = name;

            name = pad + name.Substring(name.LastIndexOf('\\') + 1, name.Length - 1 - name.LastIndexOf('\\'));
            strs[0] = name;
            return dataGridView1.Rows.Add(strs);
        }

        private void открытьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //OpenFileDialog ofd = new OpenFileDialog();
            //ofd.Filter = "iam |*.iam";
            //ofd.Title = "Выберите файл добавляемой сборки";
            if (path == "") path = m_InvApp.DesignProjectManager.ActiveDesignProject.WorkspacePath;
            //ofd.InitialDirectory = path;
            //if (m_InvApp != null)
            //{
            //    foreach (ProjectPath pp in m_InvApp.DesignProjectManager.ActiveDesignProject.LibraryPaths)
            //        ofd.CustomPlaces.Add(pp.Path);
            //}
            //ofd.Multiselect = true;
            //ofd.ShowDialog();
            //string filename = ofd.FileName;
            string[] names = InvDoc.u.OFD(path, "Inventor assemly |*.iam", true).Split('|');
            addData(names);
        }

        private void addData(string[] names)
        {
            foreach (string filename in names)
            {
                if (filename == "") continue;
                NameValueMap nvm = I.objs.CreateNameValueMap();
                nvm.Add("SkipAllUnresolvedFiles", true);
                m_AsmDoc = (Inventor.AssemblyDocument)m_InvApp.Documents.OpenWithOptions(filename, nvm,false);

                if (m_first)
                {
                    m_InvApp = (Inventor.Application)m_AsmDoc.Parent;
                    //path = m_AsmDoc.FullFileName.ToString();
                    //path = path.Substring(0, path.LastIndexOf('\\'));
                }
                else count.Clear();
                path = m_AsmDoc.FullFileName.ToString();
                path = path.Substring(0, path.LastIndexOf('\\'));

                m_BOM = m_AsmDoc.ComponentDefinition.BOM;

                //m_BOMView.Sort("Default BOM Structure", true, "Component Type", true, "Description", true);

                filePath = (System.IO.File.Exists(path + "\\" + "Properties.xml")) ? path + "\\" + "Properties.xml" : @"C:\ProgramData\Autodesk\Inventor Addins\Properties.xml";
                initializeProp(filePath);
                //m_AsmDoc.Close(true);
                m_first = false;
            }
            myDGV.setColors();
            //myDGV.setColors("Designer", designer);
        }

        public static void addSortProp(string[] props, Document doc, XElement el, XMLDoc xml)
        {
            foreach (Document item in doc.ReferencedDocuments)
	        {
                List<Property> lst = InvDoc.u.getProps(item, props); 
                if (lst[0].Value.ToString() == "") continue;
                XElement row = new XElement("row");
                foreach (Property pr in lst)
                {
                    string name = pr.Name, val = pr.Value.ToString();
                    name = name.Replace(" ", "");
                    row.Add(new XAttribute(name, val));
                }
                if (xml.exist("PartNumber", lst[0].Value.ToString())) continue;
		        if (item.DocumentType == DocumentTypeEnum.kAssemblyDocumentObject && item.ReferencedDocuments.Count != 0)
                    addSortProp(props, item, row, xml);
                el.Add(row);
	        }
        }

        private void dataGridView1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Delete && ((DataGridView)sender).SelectedCells.Count != 0)
            {
                foreach (DataGridViewCell c in ((DataGridView)sender).SelectedCells)
                {
                    c.Value = "";
                }
            }
            if (e.Control && e.KeyCode == Keys.C)
            {
                if (((DataGridView)sender).SelectedCells.Count == 1)
                    value = ((DataGridView)sender).CurrentCell.Value.GetType() != typeof(System.DateTime) ? ((DataGridView)sender).CurrentCell.Value.ToString() :
                        ((System.DateTime)((DataGridView)sender).CurrentCell.Value).ToString("dd.MM.yyyy");
                else if (((DataGridView)sender).SelectedCells[0].RowIndex == ((DataGridView)sender).SelectedCells[((DataGridView)sender).SelectedCells.Count - 1].RowIndex)
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

                    for (int i = ((DataGridView)sender).SelectedCells.Count - 1; i >= 0; i--)
                    {
                        ((DataGridView)sender).SelectedCells[i].Value = spl[j];
                        j++;
                        if (j == (spl.Count() - 1)) j = 0;
                    }
                }
            }
        }

        private void сохранитьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            string filename = ""; //bool flag = false;
            List<int> rows = new List<int>();
            //            Inventor.ApprenticeServerComponent appComp;
            //             Type typ = Type.GetTypeFromProgID("Inventor.ApprenticeServer");
            //             appComp = (Inventor.ApprenticeServerComponent)Activator.CreateInstance(typ);
            foreach (DataGridViewCell c in cellsChanged)
            {
                int ii = rows.Find(delegate(int i)
                {
                    return i == c.RowIndex;
                });
                if (ii == 0) { rows.Add(c.RowIndex); }
            }
            //            List<int> tmpi = new List<int>();
            //            foreach (int i in rows)
            //            {
            //                if (i == 0) flag = true;
            //                if (flag) tmpi.Add(i);
            //            }

            //tmpi.Clear();
            int col = rows.RemoveAll(delegate(int i) { return i == 0 ? true : false; });
            if (col != 0) rows.Add(0);

            foreach (int i in rows)
            {
                Inventor.Document tmpDoc;

                string[] tmp = dataGridView1[dataGridView1.Columns["fileName"].Index, i].Value.ToString().Split(':');
                if (tmp.Length < 2) continue;
                filename = tmp[0] + ":" + tmp[1];
                //filename = row.Cells[dataGridView1.Columns["fileName"].Index].Value.ToString();

                try
                {
                    tmpDoc = objs.Find(delegate(Inventor.Document doc)
                                                  {
                                                      return doc.FullFileName == filename;
                                                  });
                    if (tmpDoc == null) tmpDoc = m_InvApp.Documents.Open(filename, false);
                    //tmpDoc = appComp.Open(filename);

                    List<DataGridViewCell> cells = cellsChanged.FindAll(delegate(DataGridViewCell cell)
                    {
                        return cell.RowIndex == i;
                    });

                    foreach (DataGridViewCell c in cells)
                    {
                        if (dataGridView1.Columns[c.ColumnIndex].Name == "forSort") continue;
                        if (dataGridView1.Columns[c.ColumnIndex].Name == "fileName") continue;
                        object ob = addProp(tmpDoc, dataGridView1.Columns[c.ColumnIndex].Name, c.Value.ToString());
                    }
                    if (cells.Count != 0)
                        //tmpDoc.PropertySets.FlushToFile();
                        tmpDoc.Save2();
                }
                catch
                {
                    //MessageBox.Show(ex.ToString());
                    //appComp.Close();
                }
                //tmpDoc.Close();
                //appComp.Close();
            }

            foreach (DataGridViewCell c in cellsChanged)
            {
                c.Style.BackColor = System.Drawing.Color.Empty;
            }
            cellsChanged.Clear();

            //foreach (DataGridViewRow row in dataGridView1.Rows)
            //{
            //    string name = row.Cells[dataGridView1.ColumnCount-1].Value.ToString();
            //    Inventor.Document oDoc = m_InvApp.Documents.Open(name,false);
            //    for (int i = 0; i < result.Count; i++)
            //    {
            //        string ss = props.substring(result[i], "name=");
            //        object val = addProp(oDoc, ss, row.Cells[i+1].Value.ToString());
            //    }
            //    oDoc.Save();
            //    oDoc.Close(true);
            //}
        }

        private void dataGridView1_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            cellsChanged.Add(dataGridView1[e.ColumnIndex, e.RowIndex]);
            if ((dataGridView1["Part Number", e.RowIndex].Value == null || dataGridView1["Part Number", e.RowIndex].Value.ToString() == "")
                && (dataGridView1["Type", e.RowIndex].Value.ToString() != "" && dataGridView1["DecNumber", e.RowIndex].Value.ToString() != ""))
            {
                dataGridView1["Part Number", e.RowIndex].Value = "=<Type>.<DecNumber>";
                dataGridView1["Part Number", e.RowIndex].Style.BackColor = System.Drawing.Color.LightGray;
                cellsChanged.Add(dataGridView1["Part Number", e.RowIndex]);
            }
            dataGridView1[e.ColumnIndex, e.RowIndex].Style.BackColor = System.Drawing.Color.LightGray;
        }

        private void децимальныеНомераToolStripMenuItem_Click_1(object sender, EventArgs e)
        {
            List<string> lst1 = new List<string>();
            //lst1.Add("1"); lst1.Add("    2"); lst1.Add("        3"); lst1.Add("    4"); lst1.Add("5");
            for (int i = 0; i < dataGridView1.RowCount - 1; i++)
            {
                if (!part.Checked && dec.Checked && dataGridView1[dataGridView1.Columns["DecNumber"].Index, i].Value.ToString() != "")
                    lst1.Add(dataGridView1[0, i].Value.ToString() + ":");
                else
                    lst1.Add(dataGridView1[0, i].Value.ToString());
            }
            ftv = new FormTreeView(ref lst1);
            ftv.Show();
        }

        public DataGridView GetDGV()
        {
            return dataGridView1;
        }

        public void exportExcel(DataGridView dgv, string fileName, string fileExtension, string filePath)
        {
            try
            {
                string myFile = filePath + "\\" + fileName + fileExtension;
                System.IO.StreamWriter fs = new System.IO.StreamWriter(myFile, false);
                fs.WriteLine(@"<?xml version=""1.0""?>");      /* encoding=""WINDOWS-1251"" */
                fs.WriteLine(@"<?mso-application progid=""Excel.Sheet""?>");
                fs.WriteLine(@"<ss:Workbook xmlns:ss=""urn:schemas-microsoft-com:office:spreadsheet"">");
                // Создаём стили для таблицы
                fs.WriteLine(@"  <ss:Styles>");
                // Стиль для заголовков колонок
                fs.WriteLine(@"    <ss:Style ss:ID=""1"">");
                fs.WriteLine(@"      <ss:Font ss:Bold=""1""/>");
                fs.WriteLine(@"      <ss:Alignment ss:Horizontal=""Center"" ss:Vertical=""Center"" ss:WrapText=""1""/>");
                fs.WriteLine(@"     <ss:Borders>");
                fs.WriteLine(@"         <ss:Border ss:Position=""Bottom"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>");
                fs.WriteLine(@"         <ss:Border ss:Position=""Left"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>");
                fs.WriteLine(@"         <ss:Border ss:Position=""Right"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>");
                fs.WriteLine(@"         <ss:Border ss:Position=""Top"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>");
                fs.WriteLine(@"     </ss:Borders>");
                fs.WriteLine(@"    </ss:Style>");
                // Стиль для информации в колонках
                fs.WriteLine(@"    <ss:Style ss:ID=""2"">");
                fs.WriteLine(@"      <ss:Alignment ss:Vertical=""Center"" ss:WrapText=""1""/>");
                fs.WriteLine(@"     <ss:Borders>");
                fs.WriteLine(@"         <ss:Border ss:Position=""Bottom"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>");
                fs.WriteLine(@"         <ss:Border ss:Position=""Left"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>");
                fs.WriteLine(@"         <ss:Border ss:Position=""Right"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>");
                fs.WriteLine(@"         <ss:Border ss:Position=""Top"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>");
                fs.WriteLine(@"     </ss:Borders>");
                fs.WriteLine(@"    </ss:Style>");
                fs.WriteLine(@"    <ss:Style ss:ID=""3"">");
                fs.WriteLine(@"      <ss:Alignment ss:Horizontal=""Center"" ss:Vertical=""Center"" ss:WrapText=""1""/>");
                fs.WriteLine(@"     <ss:Borders>");
                fs.WriteLine(@"         <ss:Border ss:Position=""Bottom"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>");
                fs.WriteLine(@"         <ss:Border ss:Position=""Left"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>");
                fs.WriteLine(@"         <ss:Border ss:Position=""Right"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>");
                fs.WriteLine(@"         <ss:Border ss:Position=""Top"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>");
                fs.WriteLine(@"     </ss:Borders>");
                fs.WriteLine(@"    </ss:Style>");
                fs.WriteLine(@"  </ss:Styles>");
                // Записываем содержимое таблицы
                fs.WriteLine(@"<ss:Worksheet ss:Name=""Sheet1"">");
                fs.WriteLine(@"  <ss:Table>");
                for (int i = 0; i < dgv.ColumnCount - 1; i++)
                {
                    if (dgv.Columns[i].Visible == true)
                        fs.WriteLine(String.Format(@"    <ss:Column ss:Width=""{0}""/>", dgv.Columns[i].Width));
                }
                fs.WriteLine(@"    <ss:Row>");
                for (int i = 0; i < dgv.ColumnCount - 1; i++)
                {
                    if (dgv.Columns[i].Visible == true)
                        fs.WriteLine(String.Format(@"      <ss:Cell ss:StyleID=""1"">""<ss:Data ss:Type=""String"">{0}</ss:Data></ss:Cell>", dgv.Columns[i].HeaderText));
                }
                fs.WriteLine(@"    </ss:Row>");

                // В процессе добавления проверяем пустые строки
                int subtractBy; string cellText;
                if (dgv.AllowUserToAddRows == true) subtractBy = 1;
                else subtractBy = 1;
                // Записываем содержимое каждой ячейки

                for (int i = 0; i < dgv.RowCount - subtractBy; i++)
                {
                    //fs.WriteLine(String.Format(@"    <ss:Row ss:Height=""{0}"">",dgv.Rows[i].Height));
                    fs.WriteLine(@"    <ss:Row>");
                    for (int intCol = 0; intCol < dgv.ColumnCount - 2; intCol++)
                    {
                        if (dgv.Columns[intCol].Visible == true)
                        {
                            if (dgv[intCol, i].Value == null) continue;
                            cellText = dgv[intCol, i].Value.ToString();
                            //Type type = dgv[intCol, i].ValueType;
                            if (cellText == null) cellText = "";
                            if (intCol != 3)
                                fs.WriteLine(String.Format(@"      <ss:Cell ss:StyleID=""2"">""<ss:Data ss:Type=""String"">{0}</ss:Data></ss:Cell>", cellText));
                            else fs.WriteLine(String.Format(@"      <ss:Cell ss:StyleID=""3"">""<ss:Data ss:Type=""Number"">{0}</ss:Data></ss:Cell>", cellText));
                        }
                    }
                    fs.WriteLine(@"    </ss:Row>");
                }
                // Закрываем документ
                fs.WriteLine(@"  </ss:Table>");
                fs.WriteLine(@"</ss:Worksheet>");
                fs.WriteLine(@"</ss:Workbook>");
                fs.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
                throw;
            }

            // Открываем файл в Microsoft Excel
            // 10 = SW_SHOWDEFAULT
            //ShellEx(Me.Handle, "Open", myFile, "", "", 10)
        }

        //private void экспортВExcelToolStripMenuItem_Click(object sender, EventArgs e)
        //{
        //    string name = "Завесы";
        //    if (textBox1.Text != "") name = textBox1.Text;
        //    exportExcel(dataGridView1, name, ".xls", path);
        //}

        private void Prop_FormClosing(object sender, FormClosingEventArgs e)
        {

            foreach (Inventor.Document doc in objs)
            {
                try
                { doc.ReleaseReference(); }
                catch { }
            }
        }


        private void findInDGWR(DataGridViewRow row, DataGridView _dgv, ref int summ, ref List<int> rows)
        {
            int f;

            for (int i = row.Index + 1; i < _dgv.Rows.Count - 1; i++)
            {
                f = rows.Find(delegate(int num) { return i == num; });
                if (f == 0 && row.Cells[2].Value.ToString() == _dgv[2, i].Value.ToString() && row.Cells[1].Value.ToString() == _dgv[1, i].Value.ToString())
                {
                    rows.Add(i);
                    summ += Convert.ToInt16(_dgv[3, i].Value);
                }
            }
        }


        private void groupStandart(DataGridView _dgv)
        {
            List<DataGridViewRow> col = new List<DataGridViewRow>();
            List<int> rows = new List<int>();
            List<int> summs = new List<int>();
            int summ, val;
            for (int i = 0; i < _dgv.Rows.Count - 1; i++)
            {
                val = Convert.ToInt16(_dgv[3, i].Value);
                summ = val;
                if (_dgv[2, i].Value.ToString() != "")
                    findInDGWR(_dgv.Rows[i], _dgv, ref summ, ref rows);
                if (summ != val)
                {
                    col.Add(_dgv.Rows[i]);
                    summs.Add(summ);
                }
                summ = 0;
            }
            if (col.Count != 0)
            {
                _dgv.Rows.Add(new[] { "", "", "Общее количество крепежа" });
                for (int i = 0; i < col.Count; i++)
                {
                    try
                    {
                        val = _dgv.Rows.Add();
                        for (int i1 = 0; i1 < col[i].Cells.Count; i1++)
                        {
                            _dgv[i1, val].Value = col[i].Cells[i1].Value;
                        }
                        _dgv[3, val].Value = summs[i];
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.ToString());
                        throw;
                    }

                }
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            groupStandart(dataGridView1);
        }

        private void экспортВXMLToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //XElement el = new XElement();
        }

        private void сегодняToolStripMenuItem_Click(object sender, EventArgs e)
        {
            foreach (DataGridViewCell item in dataGridView1.SelectedCells)
            {
                item.Value = System.DateTime.Now.ToString("dd.MM.yyyy");
            }
        }

        private void гребенюкToolStripMenuItem_Click(object sender, EventArgs e)
        {
            foreach (DataGridViewCell item in dataGridView1.SelectedCells)
            {
                item.Value = ((ToolStripMenuItem)sender).Text;
            }
        }

        private void федотовToolStripMenuItem_Click(object sender, EventArgs e)
        {
            foreach (DataGridViewCell item in dataGridView1.SelectedCells)
            {
                item.Value = ((ToolStripMenuItem)sender).Name;
            }
        }

        private void сидоровToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void калягинаToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void голубевToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void маррToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void создатьСборкуСвойствToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (InvAddIn.PropBtn.prjPr == null)
            {
                string path = InvDoc.u.pathUtil(I.aDoc());
                XMLDoc xdoc = new XMLDoc(path + "\\Pathes.xml", "row");
                string name = ""; bool first = true; string nameforsave = "";
                while ((name = InvDoc.u.OFD(path, filter: "Assembly files(*.iam)|*.iam|Part files(*.ipt)|*.ipt")) != "")
                {
                    if (first)
                    {
                        nameforsave = System.IO.Path.GetFileNameWithoutExtension(name);
                        first = false;
                    }
                    xdoc.Doc.Root.Add(new XElement("row", new XAttribute("ffn", name)));
                }
                xdoc.Name = System.IO.Path.Combine(System.IO.Path.GetDirectoryName(xdoc.Name), nameforsave + ".xml");
                xdoc.save();
            }
            else
            {
                projectProperties pr = InvAddIn.PropBtn.prjPr;
                List<string> nmes = new List<string>(){"ffn"};
                XMLDoc xdoc = pr.projectPr.copy();

                pr.removeAtt(new List<string> { "ffn" }, pr.projectPr);
//                 foreach (var item in xdoc.Doc.Root.Descendants("row"))
//                 {
//                     XElement el = new XElement("row");
//                     foreach (var name in nmes)
//                     {
//                         el.Add(new XAttribute(name, item.Attribute(name).Value));
//                     }
//                     item.ReplaceAttributes(el.Attributes());
// //                     foreach (var att in item.Attributes())
// //                     {
// //                         if (!nmes.Contains(att.Name.ToString())) 
// //                             att.Remove(); 
// //                     }
//                 }
//                 xdoc.save();
                pr.backColorBlock();
                pr.projectPr.save();
                pr.fillProps(pr.projectPr);
            }
        }

        private void content(Document doc, XMLDoc xdoc, string[] props)
        {
            List<Property> lst = InvDoc.u.getProps(doc, props);
            XElement n = new XElement("row");
            foreach (Property pr in lst)
            {
                string name = pr.Name, val = pr.Value.ToString();
                name = name.Replace(" ", "");
                n.Add(new XAttribute(name, val));
            }
            xdoc.El.Add(n);
            addSortProp(props, doc, n, xdoc);
        }

        private void открытьСборкуСвойствToolStripMenuItem_Click(object sender, EventArgs e)
        {
            projectProperties pr = new projectProperties(new string[] { "Part Number", "Description" });
//             string path = InvDoc.util.OFD(InvDoc.util.pathUtil(Macros.StandardAddInServer.m_inventorApplication.ActiveDocument), "XML files(*.xml)|*.xml");
//             XMLDoc xdocPath = new XMLDoc(path, "row");
//             string[] names = xdocPath.Doc.Descendants("Value").Select(el => el.Value).ToArray();
//             //addData(names);
//             Document doc = null;/* = Macros.StandardAddInServer.m_inventorApplication.ActiveDocument;*/
//             string pathContent = System.IO.Path.GetDirectoryName(names[0]) + "\\Содержание.xml";
//             XMLDoc xdoc = new XMLDoc(pathContent, "row");
//             foreach (string filename in names)
//             {
//                 if (filename == "") continue;
//                 NameValueMap nvm = I.objs.CreateNameValueMap();
//                 nvm.Add("SkipAllUnresolvedFiles", true);
//                 doc = m_InvApp.Documents.OpenWithOptions(filename, nvm, false);
//                 content(doc, xdoc, new string[] { "Part Number", "Description" });
//             }
//             foreach (XElement item in xdoc.El.Elements())
//             {
//                 xdoc.sortContent(item, "PartNumber", "");
//             }
//             xdoc.save();
        }

        private void переименоватьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            string path = InvDoc.u.OFD(file.p(I.aDoc().FullFileName), "XML files(*.xml)|*.xml");
            XMLDoc xdocPath = new XMLDoc(path, "row");
            string fileName = path.Replace(".xml", ".txt");
            if (System.IO.File.Exists(fileName)) System.IO.File.Delete(fileName);
            using (System.IO.StreamWriter f = new System.IO.StreamWriter(fileName, false, System.Text.Encoding.Unicode))
            {
               foreach (XElement el in xdocPath.El.Descendants("row"))
	            {
                    string txt = "";
                    for (int i = 0; i < el.Ancestors().Count(); i++)
                    {
                        txt += "\t";
                    }
                    txt += el.Attribute("PartNumber").Value + "(" + el.Attribute("Description").Value + ")";
                    f.WriteLine(txt);
                }
            }
        }

        private void открытьПроектToolStripMenuItem_Click(object sender, EventArgs e)
        {
            dataGridView1.Visible = false; dec.Visible = false; bom.Visible = false; part.Visible = false;
            InvAddIn.PropBtn.prjPr = new projectProperties(this);
            InvAddIn.PropBtn.prjPr.show();
//             Form f = (sender as ToolStripMenuItem).Owner.Parent as Form;
//             f.TopMost = true;
//             MyDGV mydgv = new MyDGV();
//             dataGridView1.Visible = false;
//              XElement node = null;
//              this.WindowState = FormWindowState.Maximized;
//              System.Drawing.Rectangle bnds = Screen.PrimaryScreen.WorkingArea;
//              System.Drawing.Point pt = new System.Drawing.Point(bnds.Left, bnds.Top+50);
//             Dictionary<string,string> dic = new Dictionary<string,string>() {{"name","Название файла"},{"description","наименование"},{"decNumber", "децимальный номер"},{"base","название файла для наследования"},{"files","Файлы для замены"},{"replace", "Замена"},{"count", "Кол-во"}};
//             float[] weigth = {0.2f,0.2f,0.1f,0.2f,0.1f,0.2f,0.1f};
//             DataGridView dgv = mydgv.addDGV(pt, bnds.Width, bnds.Height - 150, node, dic, weigth);
//             this.Controls.Add(dgv);
        }

        private void сохранитьToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            projectProperties pr = InvAddIn.PropBtn.prjPr;
            IEnumerable<IGrouping<int, MyDGV.changeData>> gr = pr.mydgv.changes.GroupBy(c => c.rowInd);
            foreach (IGrouping<int, MyDGV.changeData> g in gr)
            {
                MyDGV.changeData cd = g.ElementAt(0);
                //if (cd.oldColor.Equals(System.Drawing.Color.Indigo)) continue;
                //XAttribute att = cd.el.Attribute("ffn");
                string fn = cd.oldVal[cd.oldVal.Length-1];
                InventorPRoperties prop = pr.properties[fn];
                for (int i = 0; i < cd.val.Length; i++)
			    {
                    if (cd.val[i] != null)
                    {
                        if (prop[i] == null)
                        {
                            prop.add<string>(prop.names[i], "");
                        }
                        prop[i].Value = cd.val[i];
                        cd.val[i] = prop[i].Value.ToString();
                        pr.mydgv.Dgv[i, cd.rowInd].Style.BackColor = cd.oldColor;
                    }
			    }
                prop.Doc.Save2(false);
//                 for (int i = 0; i < cd.val.Length; i++)
//                 {
//                     if (cd.val[i] != null)
//                     {
//                         cd.val[i] = prop[i].Value.ToString();
//                         pr.mydgv.Dgv[i, cd.rowInd].Style.BackColor = cd.oldColor;
//                     }
//                 } 
            }
        }

        private void обновитьDXFToolStripMenuItem_Click(object sender, EventArgs e)
        {
            projectProperties pr = InvAddIn.PropBtn.prjPr;
            DataGridView dgv = pr.mydgv.Dgv;
            List<Document> docs = new List<Document>();
            I.silent(true);
            foreach (DataGridViewRow row in dgv.SelectedRows)
            {
                string n = dgv[dgv.Columns[dgv.ColumnCount - 1].Index, row.Index].Value.ToString();
                Document doc = pr.properties[n].Doc;
                if (doc.Dirty) doc.Update2();
                docs.Add(doc);
            }
            IEnumerable<IGrouping<string, Document>> gr = docs.GroupBy(d => System.IO.Path.GetDirectoryName(d.FullFileName));
            List<Document> drws = new List<Document>();
            foreach (IGrouping<string, Document> g in gr)
            {
                IEnumerable<string> names = System.IO.Directory.GetFiles(g.Key, "*.idw").Where(ind => ind.IndexOf("OldVersions") == -1);
                foreach (Document item in g)
	            {
                    string v = item.PropertySets[3][2].Value.ToString();
		            string fn = names.FirstOrDefault(el => el.IndexOf(v) != -1);
                    if (fn != null)
                    drws.Add(pr.open(fn));
	            }
            }
            docs.AddRange(drws);
            
            InvAddIn.PDFOp pdf = new PDFOp(docs);
            I.silent(false);
            foreach (DataGridViewRow row in dgv.SelectedRows)
            {
                string n = dgv[dgv.Columns[dgv.ColumnCount - 1].Index, row.Index].Value.ToString();
                InventorPRoperties p = pr.properties[n];
                foreach (DataGridViewColumn col in dgv.Columns)
                {
                   string str = col.Name.Remove(col.Name.IndexOf(":"));
                   if (str == "ffn") continue;
                   string txt = p[str].Value.ToString();
                   if (txt.Length > 1 && txt[txt.Length - 3] == ':')
                       txt = txt.Remove(txt.IndexOf(" "));
                   if (dgv[col.Index, row.Index].Value.ToString() != txt)
                       dgv[col.Index, row.Index].Value = txt;
                }
            }
        }

        private void обновитьPDFToolStripMenuItem_Click(object sender, EventArgs e)
        {
            projectProperties pr = InvAddIn.PropBtn.prjPr;
            DataGridView dgv = pr.mydgv.Dgv;
            string path = System.IO.Path.GetDirectoryName(pr.projectPr.Name);
            path += "\\Документация\\";
            XElement elem = pr.projectPr.El.FirstNode as XElement;
            string ffn; Document doc; AssemblyComponentDefinition acd;
            
            if (dgv.SelectedRows.Count == 0)
            {
                foreach (var el in pr.projectPr.El.Elements())
                {
                    ffn = el.FirstAttribute.Value;
                    doc = I.open(ffn);

                    acd = I.getACD(doc);
                    pr.mKart(el, acd.BOM.BOMViews[1]);
                }

                pr.copyFiles(pr.projectPr, path);
            }
            else
            {
                XMLDoc xdoc = pr.addSelectedRow(pr.mydgv.Dgv);
                xdoc.Name = pr.projectPr.Name;

                foreach (var el in pr.projectPr.El.Elements())
                {
                    ffn = el.FirstAttribute.Value;
                    doc = I.open(ffn);

                    acd = I.getACD(doc);
                    pr.mKart(el, acd.BOM.BOMViews[1]);
                }
                pr.copyFiles(xdoc, path);
            }
            Excel.InvExcel exc = new Excel.InvExcel();
            string title = elem.Attribute("PartNumber").Value;
            exc.add(pr.projectPr.El, title);
            exc.save(path + title + ".xls");
            exc.close();
        }

        private void открытьФайлыToolStripMenuItem_Click(object sender, EventArgs e)
        {
            projectProperties pr = InvAddIn.PropBtn.prjPr;
            XMLDoc xdoc = pr.addSelectedRow(pr.mydgv.Dgv);
            foreach (var item in xdoc.El.Descendants("row"))
            {
                pr.open(item.Attribute("ffn").Value, true);
            }
        }

        private void децимальныеНомераToolStripMenuItem_Click(object sender, EventArgs e)
        {
            projectProperties pr = InvAddIn.PropBtn.prjPr;
            ToolStripMenuItem tsmi = sender as ToolStripMenuItem;
            XMLDoc d = pr.projectPr;
            InterfaceDll.MyDGV.DataGridViewRowsReorderBehavior edit = null;
            if (pr.countDecNumberClick == 0)
            {
                pr.mydgv.removeEvents();
                edit = new InterfaceDll.MyDGV.DataGridViewRowsReorderBehavior(pr.mydgv.Dgv);
                tsmi.Text = "Создать децимальные номера";
            }
            else if (pr.countDecNumberClick == 1)
            {
                pr.removeAtt(new List<string> { "ffn" }, d);
                d.addDecNumbers();
                d.save();
                tsmi.Text = "Закончить редактирование";
            }
            else if (pr.countDecNumberClick == 2)
            {
                edit.removeEvents();
                pr.mydgv.addEvents();
                tsmi.Text = "Перезапустите функцию";
                tsmi.Enabled = false;
            }
            else { }
            pr.countDecNumberClick++;
        }

        private void проверитьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            projectProperties pr = InvAddIn.PropBtn.prjPr;
            DataGridView dgv = pr.mydgv.Dgv;

            foreach (DataGridViewRow row in dgv.Rows)
            {
                if (row.Cells[0].Value == null) continue;
                string fn = row.Cells[pr.ffnIndex].Value.ToString();
                if (pr.check(fn))
                {
                    row.Cells[0].Style.ForeColor = System.Drawing.Color.Green;
                }
                else
                {
                    row.Cells[0].Style.ForeColor = System.Drawing.Color.Red;
                }
            }
        }     
    }

    public class projectProperties
    {
        public MyDGV mydgv;
        public string designer;
        public int countDecNumberClick = 0;
        //DataGridView dgv;
        public XMLDoc propNames, projectPr;
        string oldEl = "", nEl = "";
        int indNew, indOld;
        string filePath, path;
        Document doc;
        InventorPRoperties pr;
        Documents docs;
        NameValueMap nvm;
        public Dictionary<string,InventorPRoperties> properties, mProperties;
        Dictionary<string, string> dic;
        Dictionary<int, string> dgvLink;
        float[] weigth;
        string[] namesProps;
        public int ffnIndex;
        Form f;
        InterfaceDll.MyDGV.DataGridViewRowsReorderBehavior behavior;
        XElement node = null;
        System.Drawing.Rectangle bnds;
        System.Drawing.Point pt;
        public projectProperties(Form form)
        {
            mydgv = new MyDGV();
            f = form;
            docs = I.app.Documents;
            nvm = I.objs.CreateNameValueMap();
            nvm.Add("SkipAllUnresolvedFiles", true);
            properties = new Dictionary<string, InventorPRoperties>();
            bnds = Screen.PrimaryScreen.WorkingArea;
            Control ms = f.Controls["menuStrip1"];
            pt = new System.Drawing.Point(bnds.Left, ms.Bounds.Bottom);
            doc = I.aDoc();
            path = System.IO.Path.GetDirectoryName(doc.FullFileName);
            filePath = (System.IO.File.Exists(path + "\\" + "Properties.xml")) ? path + "\\" + "Properties.xml" : @"C:\ProgramData\Autodesk\Inventor Addins\Properties.xml";
            propNames = new XMLDoc(filePath, "head");
            fillDic();
        }

        void Dgv_RowsRemoved(object sender, DataGridViewRowsRemovedEventArgs e)
        {
            indOld = e.RowIndex;
            nEl = mydgv.Dgv[ffnIndex, indNew].Value.ToString();
            oldEl = mydgv.Dgv[ffnIndex, indOld].Value.ToString();
            if (nEl != "" && oldEl != "")
            {
                XElement n = projectPr.find("ffn", nEl), old = projectPr.find("ffn", oldEl);
                if (n != null && old != null)
                {
                    int count = n.Descendants().Count();
                    projectPr.insert(old, n);
                    DataGridViewRow [] rows = new DataGridViewRow[count];
                    for (int i = 0; i < count; i++)
			        {
			            rows[i] = mydgv.Dgv.Rows.SharedRow(i+indOld);
			        }
                    mydgv.Dgv.Rows.InsertRange(indNew, rows);
                    for (int j = 0; j < count; j++)
                    {
                        mydgv.Dgv.Rows.RemoveAt(j + indNew);
                    }
                }
                nEl = "";
            }
        }

        void Dgv_RowsAdded(object sender, DataGridViewRowsAddedEventArgs e)
        {
            indNew = e.RowIndex;
        }
        public projectProperties(string[] props)
        {
            path = InvDoc.u.OFD(file.p(I.aDoc().FullDocumentName), "XML files(*.xml)|*.xml");
            projectPr = new XMLDoc(path, "row");
            properties = new Dictionary<string, InventorPRoperties>();
            namesProps = props;
            if (!projectPr.El.HasAttributes) addStructure(projectPr);
            else loadStructure(projectPr);
            fillProps(projectPr);
            projectPr.Name = System.IO.Path.Combine(System.IO.Path.GetDirectoryName(projectPr.Name), "Содержание.xml");
            projectPr.save();
        }
        public void addCont(string path, XMLDoc xdoc, string n)
        {
            if (projectPr != null)
            {
                removeAtt(new List<string> { "PDF", "PartNumber", "Description" },xdoc);
                string oldPath = xdoc.Name;
                xdoc.Name = path + n;
                xdoc.El.SetAttributeValue("Path", "");
                xdoc.El.SetAttributeValue("Docum", "");
                xdoc.save();
                xdoc.Name = oldPath;
            }
        }
        public Document open(string fname, bool vis = false)
        {
            return docs.OpenWithOptions(fname, nvm, vis);
        }
        public XMLDoc addSelectedRow(DataGridView dgv)
        {
            XMLDoc xdoc = new XMLDoc("C:\\tmp.xml", "row");
            foreach (DataGridViewRow row in dgv.SelectedRows)
            {
                XElement el = projectPr.getXElement(dgv[ffnIndex, row.Index].Value.ToString(), "ffn", "row");
                xdoc.El.Add(el);
            }
            return xdoc;
        }
        public void fillDic()
        {
            var elems = propNames.El.Element("Properties").Elements();
            weigth = new float[elems.Count()+1];
            namesProps = new string[elems.Count()];
            int i = 0;    
            double sum = elems.Sum(e => InvDoc.u.convToDouble(e.Attribute("width").Value));
            dic = new Dictionary<string, string>();
            foreach (var item in elems)
            {
                dic.Add(item.Attribute("name").Value, item.Attribute("columnName").Value);
                weigth[i] = (float)(InvDoc.u.convToDouble(item.Attribute("width").Value) / sum);
                namesProps[i] = item.Attribute("name").Value;
                i++;
            }
            weigth[i] = 0.01f;
            dic.Add("ffn", "Путь"); 
        }
        public void fillProps(XMLDoc xdoc)
        {
            //string[] names = xdocPath.Doc.Descendants("row").Select(el => el.Attribute("ffn").Value).ToArray();
            //string pathContent = System.IO.Path.GetDirectoryName(names[0]) + "\\Содержание.xml";
            bool f = true;
            foreach (XElement filename in xdoc.El.Descendants("row").Where(a => a.Attribute("ffn") != null))
            {
                addPropToXML(properties[filename.Attribute("ffn").Value], filename);
                if (f && filename.Attribute("Designer") != null)
                {
                    designer = filename.Attribute("Designer").Value; 
                    f = false;
                }
            }
            foreach (XElement item in xdoc.El.Elements())
            {
                xdoc.sortContent(item, "PartNumber", "");
            }
            //projectPr.save();
        }
        public void addStructure(XMLDoc xdoc)
        {
            string[] names = xdoc.El.Elements().Select(el => el.Attribute("ffn").Value).ToArray();
            foreach (var n in names)
            {
                XElement filename = xdoc.El.Elements().FirstOrDefault(c => c.Attribute("ffn").Value == n);
                if (filename.Attribute("ffn").Value == "") continue;
                Document doc = open(filename.Attribute("ffn").Value);
                content(doc, xdoc, filename, namesProps);
            }
            xdoc.El.Add(new XAttribute("first", "no"));
            //xdoc.save();
        }
        public void loadStructure(XMLDoc xdoc)
        {
            foreach (var filename in xdoc.El.Descendants("row").Where(el => el.Attribute("ffn") != null))
            {
//                 NameValueMap nvm = I.objs.CreateNameValueMap();   
//                 nvm.Add("SkipAllUnresolvedFiles", true);
//                 Documents docs = Macros.StandardAddInServer.m_inventorApplication.Documents;
                doc = I.open(filename.Attribute("ffn").Value); //docs.OpenWithOptions(filename.Attribute("ffn").Value, nvm, false);
                pr = new InventorPRoperties(doc, namesProps);
                if (!properties.ContainsKey(doc.FullFileName))
                    properties.Add(doc.FullFileName, pr);
            }
        }
        private void content(Document doc, XMLDoc xdoc, XElement el, string[] props)
        {
            pr = new InventorPRoperties(doc, props);
            if (!properties.ContainsKey(doc.FullFileName))
            {
                properties.Add(doc.FullFileName, pr);
                //InvDoc.util.getProps(doc, props);
                XElement n = el/*new XElement("row")*/;
                if (n.Attribute("ffn") == null) n.Add(new XAttribute("ffn", doc.FullFileName));
                //addPropToXML(pr, n);
                //xdoc.El.Add(n);
                addSortProp(props, doc, n, xdoc);
            }
        }
        public void addPropToXML(InventorPRoperties pr, XElement n)
        {
            foreach (Property p in pr.props)
            {
                if (p == null)
                    continue;
                string name = p.Name, val = p.Value.ToString();
                if (val.Length > 2 && val[val.Length - 3] == ':')
                    val = val.Remove(val.IndexOf(" "));
                name = name.Replace(" ", "");
                n.SetAttributeValue(name, val);
                //n.Add(new XAttribute(name, val));
            }
        }
        public void addToMKart(Document doc, XElement el, BOMView bView)
        {
            SheetMetalComponentDefinition smcd = I.getSMCD(doc);
            if (smcd != null)
            {
                FlatPattern fp = smcd.FlatPattern;
                if (fp == null) return;
                el.SetAttributeValue("w", Math.Round(fp.Width*10, 1));
                el.SetAttributeValue("l", Math.Round(fp.Length*10, 1));
                el.SetAttributeValue("t", Math.Round((double)smcd.Thickness.Value * 10, 1));
            }
            BOMRow row = ut.get<BOMRow>(bView.BOMRows, f => f.ReferencedFileDescriptor.FullFileName == doc.FullDocumentName);
            if (row == null) return;
            el.SetAttributeValue("Count", row.TotalQuantity);
        }

        public void backColorBlock()
        {
            foreach (DataGridViewRow row in mydgv.Dgv.Rows)
            {
                if (row.Index == mydgv.Dgv.RowCount - 1)
                    break;
                string name = row.Cells[mydgv.Dgv.ColumnCount - 1].Value.ToString();
                    XElement el = projectPr.getXElement(name, "ffn", "row");
                if (row.Cells[1].Style.BackColor == System.Drawing.Color.Silver)
                {
                    el.SetAttributeValue("block", 1);
                }
                else if(el.Attribute("block") != null)
                {
                    foreach (DataGridViewCell c in row.Cells)
                    {
                        c.Style.BackColor = System.Drawing.Color.Silver; 
                    }
                }
            }
        }
        public void removeAtt(List<string> nmes, XMLDoc xdoc)
        {
            //List<string> nmes = new List<string>() { "ffn" };

            foreach (var item in xdoc.El.Descendants("row"))
            {
                XElement el = new XElement("row");
                foreach (var name in nmes)
                {
                    el.Add(new XAttribute(name, item.Attribute(name).Value));
                }
                item.ReplaceAttributes(el.Attributes());
            }
        }
        public void createDir(string name)
        {
            if (!System.IO.Directory.Exists(name))
                System.IO.Directory.CreateDirectory(name);
        }

        public void copyLibraryPDF(string pathNew, string pathOld)
        {
            Regex regex = new Regex(@"_\d*\w*\d*\w_\d*_\d*_", RegexOptions.IgnoreCase);
            string name = System.IO.Path.GetFileName(pathOld);
            Match m = regex.Match(name);
        }

        public bool check(string fn)
        {
            bool r = false;
            r = exist(fn);
//             if (fn.EndsWith(".ipt"))
//             {
//                 
//             }
            return r;
        }

        public bool exist(string name)
        {
            string path = System.IO.Path.ChangeExtension(name, "idw");
            return System.IO.File.Exists(path);
        }

        public void mKart(XElement el, BOMView bView = null)
        {
            if (el.Attribute("ffn") != null && el.Attribute("ffn").Value != "")
            {
                string ffn = el.Attribute("ffn").Value;
                Document doc = I.open(ffn);
                foreach (var item in el.Elements("row"))
                {
                    //doc = I.open(item.Attribute("ffn").Value);
                    if (item.HasElements) 
                    {

                        AssemblyComponentDefinition acd = I.getACD(I.open(item.Attribute("ffn").Value)); BOMView view = acd.BOM.BOMViews[1];
                        mKart(item, view);
                    }
                    addToMKart(I.open(item.Attribute("ffn").Value), item, bView);
                }
                addToMKart(doc, el, bView);
            }
        }

        public void copyFiles(XMLDoc xdoc, string path)
        {

            string pathPDF = path + "PDF\\",
                pathDXF = path + "DXF\\";
            createDir(path);
            createDir(pathPDF); createDir(pathDXF);
            foreach (var el in xdoc.El.Descendants("row"))
            {
                string[] data = XMLDoc.getXAttributesValues(el, new string[] { "Литера1", "Литера2", "PartNumber", "CreationTime", "RevisionNumber", "ffn", "Vendor" });
                PDFOp pdf = new PDFOp();
                string tmp;// = pdf.translit(data[2] ?? "");
                string namePDF = (data[0] ?? "") + (data[1] ?? "") + "_EKD_" + (data[2] ?? "");
                if (namePDF.EndsWith("0") || (namePDF[namePDF.Length - 3] == '_' && namePDF[namePDF.Length - 4] == '0')) namePDF += "_SB";
                tmp = data[3] ?? "";
                if (tmp != "")
                    namePDF += "_" + pdf.forDxf(tmp);
                namePDF = pdf.translit(namePDF);
                namePDF += ".pdf";
                namePDF = namePDF.TrimStart(new char[] { '_' });
                if ((data[6] ?? "") != "") namePDF = data[6];
                el.SetAttributeValue("PDF", namePDF);
                bool fPDF = copy(System.IO.Path.GetDirectoryName(data[5]) + "\\PDF\\", namePDF, pathPDF), fDXF;
                if (!fPDF)
                    mydgv.setColors("ffn", data[5], System.Drawing.Color.Red, false);
                //if (!f) copyLibraryPDF(path, data[5]);
                //if (namePDF.StartsWith("_")) namePDF = namePDF.Substring(1, namePDF.Length - 1);
                string nameDXF = (data[4] ?? "").Trim();
                if (nameDXF.IndexOf(" ") != -1) nameDXF = nameDXF.Remove(nameDXF.IndexOf(" "));
                if (nameDXF.ToUpper().EndsWith(".dxf".ToUpper())) 
                {
                    fDXF = copy(System.IO.Path.GetDirectoryName(data[5]) + "\\DXF\\", nameDXF, pathDXF);
                    if (!fDXF) mydgv.setColors("ffn", data[5], System.Drawing.Color.Blue, false);
                    if (!(fDXF && fPDF)) mydgv.setColors("ffn", data[5], System.Drawing.Color.Violet, false);
                }
            }
            string nCont = "Содержание.xml";
            if (!System.IO.File.Exists(pathPDF + nCont))
            addCont(pathPDF, xdoc, nCont);
            else
            {
                System.IO.FileInfo fi = new System.IO.FileInfo(pathPDF + nCont);
                if (fi.LastWriteTimeUtc == fi.CreationTimeUtc) addCont(pathPDF, xdoc, nCont);
            }
        }
        public void addSortProp(string[] props, Document doc, XElement el, XMLDoc xml)
        {
            foreach (Document item in doc.ReferencedDocuments)
            {
                string pn = item.PropertySets[3][2].Value.ToString().Trim();
                if (pn == "") continue;
                pr = new InventorPRoperties(item, props);
                XElement row = new XElement("row");
                if (row.Attribute("ffn") == null) row.Add(new XAttribute("ffn", item.FullFileName));
                //addPropToXML(pr, row);
//                 Property prop = pr["Part Number"];
//                 if (prop != null && xml.exist("PartNumber", prop.Value.ToString())) continue;
                if (!properties.ContainsKey(item.FullFileName))
                {
                    properties.Add(item.FullFileName, pr);
                    if (item.DocumentType == DocumentTypeEnum.kAssemblyDocumentObject && item.ReferencedDocuments.Count != 0)
                        addSortProp(props, item, row, xml);
                    el.Add(row);
                }
            }
        }
        public void show()
        {
            path = InvDoc.u.OFD(InvDoc.u.pathUtil(I.aDoc()), "files(*.xml)|*.xml;*.iam");
            if (path.EndsWith(".iam"))
            {
                string path1 = InvDoc.u.pathUtil(I.aDoc());
                projectPr = new XMLDoc(path1 + "\\Pathes.xml", "row");
                //string name = ""; bool first = true; string nameforsave = "";
                projectPr.Doc.Root.Add(new XElement("row", new XAttribute("ffn", path)));
            }
            else
            projectPr = new XMLDoc(path, "row");
            if (!projectPr.El.HasAttributes) addStructure(projectPr);
            else loadStructure(projectPr);
            fillProps(projectPr);
            mydgv.addDGV(pt, bnds.Width, bnds.Height - 50, projectPr.El, dic, weigth);
            f.Controls.Add(mydgv.Dgv);
            mydgv.xmlEv += mydgv_xmlEv;
            mydgv.addEvents();
            ffnIndex = mydgv.Dgv.Columns.OfType<DataGridViewColumn>().FirstOrDefault(c => c.Name.StartsWith("ffn")).Index;
            filePath = (System.IO.File.Exists(path + "\\" + "dgv.xml")) ? path + "\\" + "Properties.xml" : @"C:\ProgramData\Autodesk\Inventor Addins\dgv.xml";
            XMLDoc xmd = new XMLDoc(filePath,"head");
            MyToolStripMenuItem tsmi = new MyToolStripMenuItem("Меню");
            tsmi.add(xmd.El, null);
            MyContextMenuStrip cms = new MyContextMenuStrip();
            cms.el = xmd.El;
            cms.add("Меню", mydgv.Dgv);
            mydgv.Dgv.RowDividerDoubleClick += Dgv_RowDividerDoubleClick;
            cms.add(tsmi.tsmis);
            cms.itemClicked(cms.cms_ItemClicked);
            //f.Controls.Add(cms.Cms);
            foreach (DataGridViewCell item in mydgv.Dgv.SelectedCells)
            {
                item.Selected = false;
            }
            mydgv.Dgv.Columns[mydgv.Dgv.ColumnCount-1].Visible = false;
            mydgv.setColors("Part Number");
            mydgv.setColors("Designer", designer, System.Drawing.Color.Silver, true);
            mydgv.setColorsInv("Литера1", "А");
            backColorBlock();
            //behavior = new MyDGV.DataGridViewRowsReorderBehavior(mydgv.Dgv);
            //mydgv.Dgv.CellValueChanged -= mydgv.changeCell;
            //mydgv.Dgv.RowsAdded += Dgv_RowsAdded;
            //mydgv.Dgv.RowsRemoved += Dgv_RowsRemoved;
        }

        void mydgv_xmlEv(object sender, xmlEventArgs e)
        {
            if (e.Vals[0] == "delete")
            {
                projectPr.delete("ffn", e.Vals[1]);
            }
        }

        public bool copy(string inputPath, string name, string outputPath)
        {
            string fn = System.IO.Path.Combine(inputPath, name),
                fon = System.IO.Path.Combine(outputPath, name);
            if (!System.IO.File.Exists(fn)) return false;
            System.IO.File.Copy(fn, fon, true);
            return true;
        }

        void Dgv_RowDividerDoubleClick(object sender, DataGridViewRowDividerDoubleClickEventArgs e)
        {
            string path = InvDoc.u.pathUtil(I.aDoc());
            string name = InvDoc.u.OFD(path, filter: "Assembly files(*.iam)|*.iam|Part files(*.ipt)|*.ipt");
            XMLDoc xdoc = new XMLDoc("C:\\xdoc.xml", "row");
            xdoc.El.Add(new XElement("row", new XAttribute("ffn", name)));
            addStructure(xdoc);
            fillProps(xdoc);
            foreach (XElement item in xdoc.El.Elements())
            {
                xdoc.sortContent(item, "PartNumber", "");
            }
            name = mydgv.Dgv[mydgv.Dgv.Columns[mydgv.Dgv.Columns.Count - 1].Index, e.RowIndex].Value.ToString();
            XElement el = projectPr.El.Descendants("row").Where(a => a.Attribute("ffn").Value == name).FirstOrDefault();
            if (el != null)
                el.AddAfterSelf(xdoc.El.Element("row"));
            mydgv.insertRow(xdoc.El.Element("row"), dic, e.RowIndex);
            mydgv.setColors("Part Number");
            mydgv.setColors("Designer", designer, System.Drawing.Color.Silver,true);
        }
    }
        internal class PropBtn : Button
        {
            public static Prop m_Prop;
            public static projectProperties prjPr;
            public Inventor.Document pDoc { get; set; }
            public static Prop getProp
            {
                get
                {
                    return m_Prop;
                }
            }

            #region "Methods"
            public PropBtn(string displayName, string internalName,string clientId, string description, string tooltip, Icon standardIcon, Icon largeIcon)
                : base(displayName, internalName, clientId, description, tooltip, standardIcon, largeIcon)
            {

            }

            protected override void ButtonDefinition_OnExecute(NameValueMap context)
            {
                Macros.StandardAddInServer.forms.Add(m_Prop);
                if (Macros.StandardAddInServer.activeteForm()) System.Windows.Forms.Application.Run(m_Prop = new InvAddIn.Prop(InventorApplication.ActiveDocument));
            }

            #endregion
        }
}
