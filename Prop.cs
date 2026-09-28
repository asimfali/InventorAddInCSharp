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
        public static string spath = "";
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
            pathCB.Width = this.Width / 3;
            pathCB.Location = new System.Drawing.Point(10, dataGridView1.Bounds.Y + dataGridView1.Bounds.Height + 10);
            bom.Location = new System.Drawing.Point(pathCB.Bounds.X + pathCB.Bounds.Width + 10, dataGridView1.Bounds.Y + dataGridView1.Bounds.Height + 10);
            Litera.Location = new System.Drawing.Point(bom.Bounds.X + bom.Bounds.Width + 10, dataGridView1.Bounds.Y + dataGridView1.Bounds.Height + 10);
            Default.Location = new System.Drawing.Point(Litera.Bounds.X + Litera.Bounds.Width + 10, dataGridView1.Bounds.Y + dataGridView1.Bounds.Height + 10);
            decCB.Location = new System.Drawing.Point(Default.Bounds.X + Default.Bounds.Width + 10, dataGridView1.Bounds.Y + dataGridView1.Bounds.Height + 10);
            model.Location = new System.Drawing.Point(decCB.Bounds.X + decCB.Bounds.Width + 10, dataGridView1.Bounds.Y + dataGridView1.Bounds.Height + 10);
            sortCB.Location = new System.Drawing.Point(model.Bounds.X + model.Bounds.Width + 10, dataGridView1.Bounds.Y + dataGridView1.Bounds.Height + 10);  
            //checkBox3.Location = new System.Drawing.Point(checkBox2.Bounds.X + checkBox2.Bounds.Width + 10, dataGridView1.Bounds.Y + dataGridView1.Bounds.Height + 10);
            //label1.Location = new System.Drawing.Point(checkBox3.Bounds.X + checkBox3.Bounds.Width + 10, dataGridView1.Bounds.Y + dataGridView1.Bounds.Height + 10);
            //textBox1.Width = 200;
            //textBox1.Location = new System.Drawing.Point(label1.Bounds.X + label1.Bounds.Width + 1, dataGridView1.Bounds.Y + dataGridView1.Bounds.Height + 10);
            //button1.Location = new System.Drawing.Point(textBox1.Bounds.X + textBox1.Bounds.Width + 1, dataGridView1.Bounds.Y + dataGridView1.Bounds.Height + 10);

            //bns.Y = 0;
            //bns.Height = menuStrip1.Size.Height;
            //menuStrip1.Bounds = bns; 
            pathCB.DropDownHeight = this.Height * 2 / 3;
            string _path = I.app.DesignProjectManager.ActiveDesignProject.WorkspacePath;
            pathCB.Text = _path;
            pathCB.Items.AddRange(dirs(_path));
            Prop.spath = I.app.DesignProjectManager.ActiveDesignProject.WorkspacePath;
            this.KeyPreview = true;
            this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.Prop_KeyPress);
            this.FormClosing += Prop_FormClosing;
        }

        public static string[] dirs(string p)
        {
            MyXML exc = new MyXML("PathFilter.xml");
            var ie = file.getDirs(p, "", exc.elem.Element("Filter"),false);
            return ie.ToArray();
        }

        void Prop_KeyPress(object sender, KeyEventArgs e)
        {
            if (e.Control && e.KeyCode == Keys.O)
            {
                open();
            }
            else if (e.Control && e.KeyCode == Keys.W)
            {
                this.Close();
            }
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
                ex = new InvDoc.XML(I.p() + @"\Exceptions.xml");
                props.ReadXML("Properties", ref match, ref result);
                ex.ReadXML("Exceptions", ref ex_val, ref ex_attr);
                sort = new InvDoc.XML(I.p() + @"\Sequence.xml");
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
                ex = new InvDoc.XML(I.p() + @"\Exceptions.xml");
                props.ReadXML("Properties", ref match, ref result);
                ex.ReadXML("Exceptions", ref ex_val, ref ex_attr);
                sort = new InvDoc.XML(I.p() + @"\Sequence.xml");
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
//                     if (part.Checked == true)
//                     { if (row.ReferencedFileDescriptor.FullFileName.IndexOf(path) == -1) goto nex; }

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
                    var spl = val.Split(':');
                    object valProp = getProp(oDoc, ss);
                    if (valProp.GetType() == typeof(System.DateTime)) valProp = ((System.DateTime)valProp).ToString("dd.MM.yyyy");
                    if (spl.Count() == 1)
                    {
                        if (val != "" && val != valProp.ToString())
                            ii.Add(i + 1);
                        strs[i + 1] = (val == "") ? getProp(oDoc, ss) : addProp(oDoc, ss, val);
                    }
                    else if (spl.Count() == 2)
                    {
                        string tmp = valProp.ToString();
                        strs[i + 1] = (tmp == spl[0]) ? getProp(oDoc, ss) : addProp(oDoc, ss, spl[1]);
                    }
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
                    var spl = val.Split(':');
                    if (valProp.GetType() == typeof(System.DateTime)) valProp = ((System.DateTime)valProp).ToString("dd.MM.yyyy");
                    if (spl.Count() == 1)
                    {
                        if (val != "" && val != valProp.ToString())
                            ii.Add(i + n);
                        strs[i + n] = (val == "") ? getProp(oDoc, ss) : addProp(oDoc, ss, val);
                    }
                    else if (spl.Count() == 2)
                    {
                        string tmp = valProp.ToString();
                        strs[i + n] = (tmp == spl[0]) ? getProp(oDoc, ss) : addProp(oDoc, ss, spl[1]);
                    }
                    if (i == 1) n++;
                }
            }
            string name = oDoc.FullFileName;
//             if (!part.Checked)
//             {
//                 if (findInDgv(name)) return 0;
//             }
//             else
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

                filePath = (System.IO.File.Exists(path + "\\" + "Properties.xml")) ? path + "\\" + "Properties.xml" : I.p() + @"\Properties.xml";
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
//             List<string> lst1 = new List<string>();
//             //lst1.Add("1"); lst1.Add("    2"); lst1.Add("        3"); lst1.Add("    4"); lst1.Add("5");
//             for (int i = 0; i < dataGridView1.RowCount - 1; i++)
//             {
//                 if (!part.Checked && dec.Checked && dataGridView1[dataGridView1.Columns["DecNumber"].Index, i].Value.ToString() != "")
//                     lst1.Add(dataGridView1[0, i].Value.ToString() + ":");
//                 else
//                     lst1.Add(dataGridView1[0, i].Value.ToString());
//             }
//             ftv = new FormTreeView(ref lst1);
//             ftv.Show();
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
            InvAddIn.PropBtn.lastPath = null;
            InvAddIn.PropBtn.m_Prop = null;
            InvAddIn.PropBtn.prjPr = null;
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
                pr.lit = this.Litera.Checked;
                pr.def = this.Default.Checked;
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
            pr.setParent(this);
            pr.lit = Litera.Checked;
            pr.def = Default.Checked;
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
            Prop.spath = this.pathCB.Text;
            open();
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

        public static void open(bool add = false)
        {
            InvAddIn.PropBtn.m_Prop.dataGridView1.Visible = false;
//             InvAddIn.PropBtn.m_Prop.dec.Visible = false;
//             InvAddIn.PropBtn.m_Prop.bom.Visible = false;
//             InvAddIn.PropBtn.m_Prop.part.Visible = false;
//             InvAddIn.PropBtn.m_Prop.Litera.Visible = false;
//             InvAddIn.PropBtn.m_Prop.Default.Visible = false;
            foreach (Control item in InvAddIn.PropBtn.m_Prop.Controls)
            {
                if (item is CheckBox) item.Visible = false;
                if (item is ComboBox) item.Visible = false;
            }
            InvAddIn.PropBtn.prjPr = new projectProperties(InvAddIn.PropBtn.m_Prop);
            InvAddIn.PropBtn.prjPr.lit = InvAddIn.PropBtn.m_Prop.Litera.Checked;
            InvAddIn.PropBtn.prjPr.def = InvAddIn.PropBtn.m_Prop.Default.Checked;
            InvAddIn.PropBtn.prjPr.sort = InvAddIn.PropBtn.m_Prop.sortCB.Text;
            InvAddIn.PropBtn.prjPr.model = InvAddIn.PropBtn.m_Prop.model.Text;
            if (InvAddIn.PropBtn.m_Prop.decCB.Text != "") InvAddIn.PropBtn.prjPr.decNum = InvAddIn.PropBtn.m_Prop.decCB.Text;

            InvAddIn.PropBtn.prjPr.show(add);
        }

        private void сохранитьToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            save();
        }

        public static void save()
        {
            projectProperties pr = InvAddIn.PropBtn.prjPr;
            IEnumerable<IGrouping<int, MyDGV.changeData>> gr = pr.mydgv.changes.GroupBy(c => c.rowInd);
            foreach (IGrouping<int, MyDGV.changeData> g in gr)
            {
                MyDGV.changeData cd = g.ElementAt(0);
                string fn = cd.oldVal[cd.oldVal.Length - 1];
                InventorPRoperties prop = pr.properties[fn];
                for (int i = 0; i < cd.val.Length; i++)
                {
                    if (cd.val[i] != null)
                    {
                        try
                        {
                            if (prop[prop.names[i]] == null)
                            {
                                prop.add<string>(prop.names[i], "");
                            }
                            if (!prop.Doc.IsModifiable) continue;
                            prop[prop.names[i]].Value = cd.val[i];
                            cd.val[i] = prop[prop.names[i]].Value.ToString();
                            pr.mydgv.Dgv[i, cd.rowInd].Style.BackColor = cd.oldColor;
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show(ex.ToString());
                        }
                    }
                }
                prop.Doc.Save2(false);
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
                    drws.Add(pr.open(fn,false,true));
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
            string path = file.p(I.aDoc().FullDocumentName) + "Документация\\";
            XElement elem = pr.projectPr.El.FirstNode as XElement;
            string ffn; Document doc; AssemblyComponentDefinition acd;
            MyForm F = new MyForm("DocumInterface.xml", "Название");
            F.f.ShowDialog();
            pr.docum = F.cbs[0].Text;

            if (dgv.SelectedRows.Count == 0)
            {
//                 foreach (var el in pr.projectPr.El.Elements())
//                 {
//                     ffn = el.FirstAttribute.Value;
//                     doc = I.open(ffn);
// 
//                     acd = I.getACD(doc);
//                     projectProperties.mKart(el, acd.BOM.BOMViews[1]);
//                 }
//                 pr.createDir(path);
//                 excelRun(pr.projectPr.El, path, elem.Attribute("PartNumber").Value);
                pr.copyFiles(pr.projectPr, path);
            }
            else
            {
                XMLDoc xdoc = pr.addSelectedRow(pr.mydgv.Dgv);
                xdoc.Name = pr.projectPr.Name;

//                 foreach (var el in pr.projectPr.El.Elements())
//                 {
//                     ffn = el.FirstAttribute.Value;
//                     doc = I.open(ffn);
// 
//                     acd = I.getACD(doc);
//                     projectProperties.mKart(el, acd.BOM.BOMViews[1]);
//                 }
//                 pr.createDir(path);
//                 excelRun(pr.projectPr.El, path, elem.Attribute("PartNumber").Value);
                pr.copyFiles(xdoc, path);
            }
        }

        public static void excelRun(XElement el, string path, string title)
        {
            Excel.InvExcel exc = new Excel.InvExcel($"{path}{title}.xlsx");
            exc.add(el);
            //exc.add(el, title);
            //exc.save(path + title + ".xls");
            //exc.close();
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

        private void дляТехнологовToolStripMenuItem_Click(object sender, EventArgs e)
        {
            projectProperties.techMKarts();
        }

        private void Prop_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.Control)
            {
                switch (e.KeyCode)
                {
                    case Keys.B:
                        openCurrent();
                        break;
                    default:
                        break;
                }
            }
        }

        public void openCurrent()
        {
            InvAddIn.PropBtn.m_Prop.dataGridView1.Visible = false;
            //             InvAddIn.PropBtn.m_Prop.dec.Visible = false;
            //             InvAddIn.PropBtn.m_Prop.bom.Visible = false;
            //             InvAddIn.PropBtn.m_Prop.part.Visible = false;
            //             InvAddIn.PropBtn.m_Prop.Litera.Visible = false;
            //             InvAddIn.PropBtn.m_Prop.Default.Visible = false;
            foreach (Control item in InvAddIn.PropBtn.m_Prop.Controls)
            {
                if (item is CheckBox) item.Visible = false;
                if (item is ComboBox) item.Visible = false;
            }
            InvAddIn.PropBtn.prjPr = new projectProperties(InvAddIn.PropBtn.m_Prop);
            InvAddIn.PropBtn.prjPr.lit = InvAddIn.PropBtn.m_Prop.Litera.Checked;
            InvAddIn.PropBtn.prjPr.def = InvAddIn.PropBtn.m_Prop.Default.Checked;
            InvAddIn.PropBtn.prjPr.sort = InvAddIn.PropBtn.m_Prop.sortCB.Text;
            InvAddIn.PropBtn.prjPr.model = InvAddIn.PropBtn.m_Prop.model.Text;
            if (InvAddIn.PropBtn.m_Prop.decCB.Text != "") InvAddIn.PropBtn.prjPr.decNum = InvAddIn.PropBtn.m_Prop.decCB.Text;

            var p = I.aDoc().FullDocumentName;

            InvAddIn.PropBtn.prjPr.show(false, p);
        }

        private void toolStripTextBox1_Click(object sender, EventArgs e)
        {

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

        private void заменитьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            MyForm F = new MyForm("ReplaceInterface.xml", "Заменить");
            //F.bnts[0].Click += fillet_Click;
            F.f.ShowDialog();
            string find = "", repl = "";
            find = F.cbs[0].Text;
            repl = F.cbs[1].Text;
            System.Windows.Forms.Control.ControlCollection cls = InvAddIn.PropBtn.m_Prop.Controls;
            DataGridView dgv = cls[cls.Count-1] as DataGridView;
            if (dgv == null) return;
            foreach (DataGridViewCell c in dgv.SelectedCells)
            {
                c.Value = Regex.Replace(c.Value as string, find, repl);  
            }
            //this.Close();
        }

        private void маршруткиToolStripMenuItem_Click(object sender, EventArgs e)
        {
            projectProperties.minMKarts(decCB.Text);
            //InvAddIn.PropBtn.prjPr = new projectProperties(InvAddIn.PropBtn.m_Prop);
            //if (decCB.Text != "") InvAddIn.PropBtn.prjPr.decNum = decCB.Text;
            
            //InvAddIn.PropBtn.prjPr.model = InvAddIn.PropBtn.m_Prop.model.Text;
            //Prop.spath = this.pathCB.Text;
            //InvAddIn.PropBtn.prjPr.createMKarts();
        }
    }

    public class projectProperties
    {
        Prop parent = null;
        //public static string PrPath = null;
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
        List<string> columns;
        float[] weigth;
        string[] namesProps;
        public int ffnIndex;
        Form f;
        public bool lit = false;
        public bool def = false;
        public string model = "";
        public string sort = "";
        public string decNum = "", docum;
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
            filePath = (System.IO.File.Exists(path + "\\" + "Properties.xml")) ? path + "\\" + "Properties.xml" : I.p() + @"\Properties.xml";
            propNames = new XMLDoc(filePath, "head");
            string w = XMLDoc.getXAttributeValue(propNames.El.Element("Properties"), "width");
            if (w != null)
            bnds.Width = (int)(u.convToDouble(w)/100*bnds.Width);
            fillDic();
        }

        public void setParent(Prop p)
        {
            parent = p;
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
                if (docum != null)
                {
                    xdoc.El.SetAttributeValue("Docum", docum);
                }
                xdoc.save();
                xdoc.Name = oldPath;
            }
        }
        public Document open(string fname, bool vis = false, bool upd = false)
        {
            Document doc = docs.OpenWithOptions(fname, nvm, vis);
            if (upd)
            {
               CreateComponent.update(doc);
            }
            return doc;
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
                MyXML exc = new MyXML("Except.xml");
                addSortProp(props, doc, n, xdoc, exc.elem.Element("Exceptions"));
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
        public static void addToMKart(Document doc, XElement el, BOMView bView)
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
            BOMRow row = null;
            row = ut.get<BOMRow>(bView.BOMRows, f => f.ReferencedFileDescriptor != null && f.ReferencedFileDescriptor.FullFileName == doc.FullDocumentName);
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

        public static XElement mKart(XElement el, BOMView bView = null)
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
            return el;
        }

        public void copyFiles(XMLDoc xdoc, string path)
        {
            bool final = false;
            string pathPDF = path + "PDF\\",
                pathDXF = path + "DXF\\",
                pathFinal = "Не менять\\";
            createDir(path);
            createDir(pathPDF); createDir(pathDXF);
            file.removeFiles(pathPDF, ".xml");
            file.removeFiles(pathPDF, ".");
            file.removeFiles(pathDXF, ".");
            //if (System.IO.Directory.Exists(pathDXF + pathFinal)) final = true;
            foreach (var el in xdoc.El.Descendants("row"))
            {
                string[] data = XMLDoc.getXAttributesValues(el, new string[] { "Литера1", "Литера2", "PartNumber", "CreationTime", "RevisionNumber", "ffn", "Vendor" });
                PDFOp pdf = new PDFOp();
                string tmp;// = pdf.translit(data[2] ?? "");
                string namePDF = (data[0] ?? "") + (data[1] ?? "") + "_EKD_" + (data[2] ?? "");
                if (namePDF.EndsWith("0") || (namePDF[namePDF.Length - 3] == '-' && namePDF[namePDF.Length - 4] == '0')) namePDF += "_SB";
                tmp = data[3] ?? "";
                if (tmp != "")
                    namePDF += "_" + pdf.forDxf(tmp);
                namePDF = pdf.translit(namePDF);
                namePDF += ".pdf";
                namePDF = namePDF.TrimStart(new char[] { '_' });
                if ((data[6] ?? "") != "") namePDF = data[6];
                //string namePDF = pdf.getNameIzv(I.open(data[5]), data[3] ?? "");
                el.SetAttributeValue("PDF", namePDF);
                //u.regex(ref data[5], @"\^\d\d", "");
                bool fPDF = copy(System.IO.Path.GetDirectoryName(data[5]) + "\\PDF\\", namePDF, pathPDF), fDXF;
                if (!fPDF)
                    mydgv.setColors("ffn", data[5], System.Drawing.Color.Red, false);
                //if (!f) copyLibraryPDF(path, data[5]);
                //if (namePDF.StartsWith("_")) namePDF = namePDF.Substring(1, namePDF.Length - 1);
                string nameDXF = (data[4] ?? "").Trim();
                if (nameDXF.IndexOf(" ") != -1) nameDXF = nameDXF.Remove(nameDXF.IndexOf(" "));
                if (nameDXF.ToUpper().EndsWith(".dxf".ToUpper())) 
                {
                    fDXF = copy(System.IO.Path.GetDirectoryName(data[5]) + "\\DXF\\", nameDXF, pathDXF, pathFinal);
//                     if (final) 
//                         if (copy(System.IO.Path.GetDirectoryName(data[5]) + "\\" + pathFinal, nameDXF, pathDXF) && !fDXF) 
//                             fDXF = false;
                    if (!fDXF) mydgv.setColors("ffn", data[5], System.Drawing.Color.Blue, false);
                    else if (!(fDXF && fPDF)) mydgv.setColors("ffn", data[5], System.Drawing.Color.Violet, false);
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
        public void addSortProp(string[] props, Document doc, XElement el, XMLDoc xml, XElement exc)
        {

            foreach (Document item in doc.ReferencedDocuments)
            {
                string pn = item.PropertySets[3][2].Value.ToString().Trim();
                if (item.DocumentType == DocumentTypeEnum.kPartDocumentObject)
                {
                    PartComponentDefinition cd = (item as PartDocument).ComponentDefinition;
                    if (cd.SurfaceBodies.Count == 0)
                        continue;
                }
                if (except(exc, item))
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
                        addSortProp(props, item, row, xml, exc);
                    el.Add(row);
                }
            }
        }
        public static bool except(XElement el, Document doc)
        {
            string fn = doc.FullDocumentName;
            foreach (var item in el.Elements())
            {
                if (fn.IndexOf(item.Value) != -1)
                    return false; 
            }
            return true;
        }
        public static XMLDoc addPatches(XMLDoc projectPr, string path, bool add)
        {
            if (add && InvAddIn.PropBtn.lastPath != null) path = InvAddIn.PropBtn.lastPath + "|" + path;
            InvAddIn.PropBtn.lastPath = path;
            var pathes = u.getSpl(path, '|');
            if (path.EndsWith(".iam") || path.EndsWith(".ipt"))
            {
                string path1 = InvDoc.u.pathUtil(I.aDoc());
                projectPr = new XMLDoc(path1 + "\\Pathes.xml", "row");
                //string name = ""; bool first = true; string nameforsave = "";
                foreach (var item in pathes)
                {
                    projectPr.Doc.Root.Add(new XElement("row", new XAttribute("ffn", item)));
                }
            }
            else
                projectPr = new XMLDoc(path, "row");
            return projectPr;
        }
        public class compare : IComparer<string>
        {
            string[] s1, s2;
            string model = "";
            string dn = null;
            string f = "", n = "", l = "00", val = "";
            public compare(string str, string model)
            {
                if (str != "")
                {
                    var spl = str.Split('|');
                    if (spl.Length != 2) return;
                    s1 = spl[0].Split(';'); s2 = spl[1].Split(';');
                    this.model = model;
                    dn = getReg();
                }
            }
            public int Compare(string x, string y)
            {
                string vx = get(x), vy = get(y);
                return String.Compare(vx, vy);
            }
            public string get(string v)
            {
                string fn = file.name(v);
                if (dn == null) return "";
                Regex r = new Regex(dn);
                Match m = r.Match(fn);
                if (m.Groups.Count == 4)
                {
                    f = m.Groups[2].Value;
                    n = m.Groups[3].Value;
                }
                r = new Regex(@"\^(\d*)");
                m = r.Match(fn);
                if (m.Groups.Count == 2)
                {
                    l = m.Groups[1].Value;
                }
                else l = "00";
                val = sort(f, s1) + sort(n, s2) + l;
                return val;
            }
            string sort(string v, string[] spl)
            {
                v = conv(v);
                for (int i = 0; i < spl.Length; i++)
                {
                    if (spl[i] == v) return i.ToString(); 
                }
                return "";
            }
            public string getReg()
            {
                if (model == "3") return @".*(\d)\d\d(\d)(\D).*";
                else if (model == "4") return @".*(\d)\d(\d)\d(\D).*";
                return null;
            }
            public string conv(string v)
            {
                return v.ToUpper().Replace("Е", "E").Replace("А", "A");
            }
        }
        public List<string> clear(List<string> lst, string v, string model)
        {
            if (v == null || v == "") return lst;
            if (v.IndexOf('|') == -1) return lst;
            var spl1 = v.Split('|');
            if (spl1[0] == null) return lst;
            var spl2 = spl1[0].Split(';');
            string pat = null;
            if (model == "3") pat = @".*\D\d\d\d(\d)\D.*";
            else if (model == "4") pat = @".*\D\d\d(\d)\d\D.*";
            Regex reg = new Regex(pat);
            List<string> ret = new List<string>();
            foreach (var item in lst)
            {
                Match m = reg.Match(item);
                if (m.Groups.Count != 2) continue;
                if (spl2.Contains(m.Groups[1].Value)) ret.Add(item);
            }
            return ret;
        }
        public void show(bool add, string pathO = "")
        {
            //string pathO = "";
            var fn = doc.FullDocumentName;
            if (pathO != "")
            {

            }
            else
            {
                if (sort == "" && decNum == "")
                    pathO = InvDoc.u.OFD(InvDoc.u.pathUtil(I.aDoc()), "files(*.xml)|*.xml;*.iam|Inventor Part(*.ipt)|*.ipt", true, fn);
                else
                {
                    MyXML exc = new MyXML("PathFilter.xml");
                    exc.elem.Element("Filter").Add(new XElement("Value", ".ipt"));
                    exc.elem.Element("Filter").Add(new XElement("Value", ".xls"));
                    exc.remove("Value", "^");
                    //MessageBox.Show("путь: " + prPath + "\n" + decNum);
                    // return;
                    var spl = file.getFiles(Prop.spath, decNum, exc.elem.Element("Filter"));
                    if (sort != "" && model != "") spl.Sort(new compare(sort, model));
                    spl = clear(spl, sort, model);
                    pathO = file.join(spl);
                }
            }
            path = pathO;
            projectPr = addPatches(projectPr, path, add);
            if (!projectPr.El.HasAttributes) addStructure(projectPr);
            else loadStructure(projectPr);
            fillProps(projectPr);   
            if (f.Controls[f.Controls.Count - 1] is DataGridView)
            {
//                 DataGridView tmp = f.Controls[f.Controls.Count - 1] as DataGridView;
//                 tmp.Rows.Clear();
//                 tmp.Update();
                f.Controls.RemoveAt(f.Controls.Count - 1);
            }
            if (def)
            {
                changeProps();
            }
            mydgv.addDGV(pt, bnds.Width, bnds.Height - 50, projectPr.El, dic, weigth);
            //if (!(f.Controls[f.Controls.Count - 1] is DataGridView))
                f.Controls.Add(mydgv.Dgv);
            mydgv.xmlEv -= mydgv_xmlEv;
            mydgv.xmlEv += mydgv_xmlEv;
            mydgv.removeEvents();
            mydgv.addEvents();
            ffnIndex = mydgv.Dgv.Columns.OfType<DataGridViewColumn>().FirstOrDefault(c => c.Name.StartsWith("ffn")).Index;
            filePath = (System.IO.File.Exists(path + "\\" + "dgv.xml")) ? path + "\\" + "Properties.xml" : I.p() + @"\dgv.xml";
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
            mydgv.Dgv.Columns[mydgv.Dgv.ColumnCount - 1].Visible = false;
            mydgv.setColors("Part Number");
            if (!lit)
            {
                mydgv.setColors("Designer", designer, System.Drawing.Color.Silver, true);
                mydgv.setColors("Vendor", "", System.Drawing.Color.Silver, true);
                mydgv.setColorsInv("Литера1", "А");
            }
            backColorBlock();
            mydgv.keyEv -= mydgv_keyEv;
            mydgv.keyEv += mydgv_keyEv;
            if (columns != null)
                mydgv.setColors(columns, System.Drawing.Color.LightGray);
            //behavior = new MyDGV.DataGridViewRowsReorderBehavior(mydgv.Dgv);
            //mydgv.Dgv.CellValueChanged -= mydgv.changeCell;
            //mydgv.Dgv.RowsAdded += Dgv_RowsAdded;
            //mydgv.Dgv.RowsRemoved += Dgv_RowsRemoved;
        }

        void changeProps()
        {
            columns = new List<string>();
            XElement el = XMLDoc.getXElement(propNames.El, "Properties");
            foreach (var item in el.Elements())
            {
                string fn = XMLDoc.getAttributeValue(item, "name");
                string v = XMLDoc.getAttributeValue(item, "value");
                if (v == null) continue;
                if (v == "сегодня") 
                    v = System.DateTime.Now.ToString("dd.MM.yyyy");
                columns.Add(fn);
                if (fn.IndexOf(" ") != -1)
                    fn = fn.Replace(" ", "");
                foreach (var e in projectPr.El.Descendants())
                {
                    if (v.IndexOf(":") == -1)
                    XMLDoc.attr(e, fn, v);
                    else
                    {
                        var spl = v.Split(':');
                        var str = XMLDoc.getAttributeValue(e, fn);
                        str = str == spl[0] ? spl[1] : str;
                        XMLDoc.attr(e, fn, str);
                    }
                }
            }
        }

        public static string[] getPathes(string name)
        {
            List<string> p = new List<string>();
            MyXML xml = new MyXML(name);
            foreach (var el in xml.elem.Elements())
            {
                var val = MyXML.getAtt(el, "ffn");
                p.Add(val);
            }
            return p.ToArray();
        }

        public static void minMKarts(string decNum = "")
        {
            int i = 1;
            XElement bel = null;
            string p = file.p(I.aDoc().FullDocumentName) + "Документация\\Маршрутки\\";
            file.dir(p);
            string pathO = "";
            if (decNum == "")
                pathO = InvDoc.u.OFD(InvDoc.u.pathUtil(I.aDoc()), "files(*.xml)|*.xml;*.iam|Inventor Part(*.ipt)|*.ipt", true);
            else
            {
                //string prPath = I.app.DesignProjectManager.ActiveDesignProject.WorkspacePath;
                MyXML exc = new MyXML("PathFilter.xml");
                exc.elem.Element("Filter").Add(new XElement("Value", ".ipt"));
                var spl = file.getFiles(Prop.spath, decNum, exc.elem.Element("Filter"));
                pathO = file.join(spl);
            }
            var pathes = u.getSpl(pathO, '|');
            if (pathes.Length == 1)
            {
                var n = pathes[0];
                if (n.EndsWith(".xml"))
                {
                    pathes = getPathes(n);
                }
            }
            AssemblyComponentDefinition acd = null;
            Dictionary<string, string> props = new Dictionary<string, string>()
            {
                {"PartNumber", "Part Number" }, {"Description", "Description"}, {"RevisionNumber", "Revision Number"}
            };
            
            foreach (var path in pathes)
            {
                XElement el = new XElement("MKart");
                AssemblyDocument adoc = I.open(path, drw: false) as AssemblyDocument;
                acd = adoc.ComponentDefinition;
                var bom = acd.BOM;
                bom.PartsOnlyViewEnabled = true;
                //foreach (BOMView item in bom.BOMViews)
                //{
                //    var te = item;
                //}
                var view = bom.BOMViews[3];
                minMKart(view, props, el);
                Excel.InvExcel exc = new Excel.InvExcel($"{p}{file.name(path)}.xlsx");
                exc.add(el);
            }
        }

        public static void techMKarts(string decNum = "")
        {
            int i = 1;
            XElement bel = null;
            string p = file.p(I.aDoc().FullDocumentName) + "Документация\\Данные\\";
            file.dir(p);
            string pathO = "";
            if (decNum == "")
                pathO = InvDoc.u.OFD(InvDoc.u.pathUtil(I.aDoc()), "files(*.xml)|*.xml;*.iam|Inventor Part(*.ipt)|*.ipt", true);
            else
            {
                //string prPath = I.app.DesignProjectManager.ActiveDesignProject.WorkspacePath;
                MyXML exc = new MyXML("PathFilter.xml");
                exc.elem.Element("Filter").Add(new XElement("Value", ".ipt"));
                var spl = file.getFiles(Prop.spath, decNum, exc.elem.Element("Filter"));
                pathO = file.join(spl);
            }
            var pathes = u.getSpl(pathO, '|');
            AssemblyComponentDefinition acd = null;
            Dictionary<string, string> props = new Dictionary<string, string>()
            {
                 {"RevisionNumber", "Revision Number"}
            };

            foreach (var path in pathes)
            {
                XElement el = new XElement("MKart");
                AssemblyDocument adoc = I.open(path, drw: false) as AssemblyDocument;
                acd = adoc.ComponentDefinition;
                var bom = acd.BOM;
                bom.PartsOnlyViewEnabled = true;
                //foreach (BOMView item in bom.BOMViews)
                //{
                //    var te = item;
                //}
                var view = bom.BOMViews[3];
                techKart(view, props, el);
                Excel.InvExcel exc = new Excel.InvExcel($"{p}{file.name(path)}.xlsx");
                exc.addTech(el);
            }
        }

        public static void minMKart(BOMView view, Dictionary<string, string> props, XElement par)
        {
            foreach (BOMRow r in view.BOMRows)
            {
                PartComponentDefinition def = r.ComponentDefinitions[1] as PartComponentDefinition;
                if (def == null) continue;
                var dics = u.getProps((Document)def.Document, props);
                if (dics.Count != 3) continue;
                var el = MyXML.addXElement("row", dics);
                SheetMetalComponentDefinition smcd = def as SheetMetalComponentDefinition;
                if (smcd != null)
                {
                    FlatPattern fp = smcd.FlatPattern;
                    if (fp == null)
                    {
                        MessageBox.Show($"Нет развертки на деталь: {((Document)smcd.Document).FullDocumentName}");
                        return;
                    }
                    el.SetAttributeValue("w", Math.Round(fp.Width * 10, 1));
                    el.SetAttributeValue("l", Math.Round(fp.Length * 10, 1));
                    el.SetAttributeValue("t", Math.Round((double)smcd.Thickness.Value * 10, 1));
                }
                el.SetAttributeValue("Count", r.ItemQuantity);
                par.Add(el);
            }
        }

        public static void techKart(BOMView view, Dictionary<string, string> props, XElement par)
        {
            foreach (BOMRow r in view.BOMRows)
            {
                PartComponentDefinition def = r.ComponentDefinitions[1] as PartComponentDefinition;
                if (def == null) continue;
                var dics = u.getProps((Document)def.Document, props);
                if (dics.Count != 1) continue;
                var el = MyXML.addXElement("row", dics);
                SheetMetalComponentDefinition smcd = def as SheetMetalComponentDefinition;

                if (smcd != null)
                {
                    FlatPattern fp = smcd.FlatPattern;
                    if (fp == null) return;
                    var st = smcd.ActiveSheetMetalStyle;
                    var mat = st.Material;
                    SheetMetalFeatures smf = smcd.Features as SheetMetalFeatures;
                    el.SetAttributeValue("mass", Math.Round(smcd.MassProperties.Mass, 2));
                    el.SetAttributeValue("area", Math.Round(fp.TopFace.Evaluator.Area*2, 2));
                    el.SetAttributeValue("w", Math.Round(fp.Width * 10, 1));
                    el.SetAttributeValue("l", Math.Round(fp.Length * 10, 1));
                    el.SetAttributeValue("t", Math.Round((double)smcd.Thickness.Value * 10, 1));
                    el.SetAttributeValue("density", Math.Round(mat.Density, 3));
                    el.SetAttributeValue("name", file.name(((Document)smcd.Document).FullDocumentName));
                    
                    var cf = u.get<CutFeature>(smf.CutFeatures, f => f.Name.ToLower() == "mark");
                    if (cf != null)
                    {
                        var count = cf.Definition.Profile.Count / 2;
                        el.SetAttributeValue("bend", count);
                    }
                    else
                    {
                        el.SetAttributeValue("bend", smcd.Bends.Count);
                    }
                }
                el.SetAttributeValue("Count", r.ItemQuantity);
                par.Add(el);
            }
        }

        public void createMKarts()
        {
            int i = 1;
            XElement vars = null, bel = null;
            string p = file.p(I.aDoc().FullDocumentName) + "Документация\\Маршрутки\\";
            createDir(p);
            string pathO = "";
            if (decNum == "")
            pathO = InvDoc.u.OFD(InvDoc.u.pathUtil(I.aDoc()), "files(*.xml)|*.xml;*.iam|Inventor Part(*.ipt)|*.ipt", true);
            else
            {
                //string prPath = I.app.DesignProjectManager.ActiveDesignProject.WorkspacePath;
                MyXML exc = new MyXML("PathFilter.xml");
                exc.elem.Element("Filter").Add(new XElement("Value", ".ipt"));
                var spl = file.getFiles(Prop.spath, decNum, exc.elem.Element("Filter"));
                pathO = file.join(spl);
            }
            var pathes = u.getSpl(pathO, '|');
            AssemblyComponentDefinition acd = null;
            properties.Clear();
            foreach (var path in pathes)
            {
                string lPath = "";
                if (path.ToLower().IndexOf("e") != 1 || path.ToLower().IndexOf("е") != -1)
                {
                    var names = TableInv.getAsms(path);
                    lPath = file.join(names);
                }
                if (lPath == "") lPath = path;
                projectPr = addPatches(projectPr, lPath, false);
                if (!projectPr.El.HasAttributes) addStructure(projectPr);
                else loadStructure(projectPr);
                fillProps(projectPr);
                string type = "", dn = @".*(\d)\d\d(\d).*\.(\d*)\.(\d*)";
                if (model != "")
                {
                    switch (model)
                    {
                        case "3":
                            dn = @".*(\d)\d(\d)\d.*\.(\d*)\.(\d*)";
                            break;
                        case "4":
                            break;
                        default:
                            break;
                    }
                }
                XElement elem = projectPr.El.FirstNode as XElement;
                XElement elo = null;
                foreach (var el in projectPr.El.Elements())
                {
                    string ffn = el.FirstAttribute.Value;
                    doc = I.open(ffn);
                    acd = I.getACD(doc);
                    string pn = u.getPropValue(doc, "Part Number");
                    if (pn != "")
                    {
                        Regex reg = new Regex(dn);
                        Match m = reg.Match(pn);
                        if (m.Groups.Count == 5)
                        {
                            type = m.Groups[1].Value + m.Groups[2].Value;
                        }
                    }
                    if (elo == null)
                    {
                        bel = new XElement("row");
                        elo = projectProperties.mKart(el, acd.BOM.BOMViews[1]);
                    }
                    else
                    {
                        var mk = projectProperties.mKart(el, acd.BOM.BOMViews[1]);
                        if (vars == null) vars = MyXML.addXElement("row", new Dictionary<string, string>() { { "Description", "Исполнения: "} });
                        if (mk != null)
                        {
                            XElement xl = MyXML.addXElement("row", new Dictionary<string, string>() { { "Description", "Исполнение " + i.ToString("00") } });
                            xl.Add(mk);
                            vars.Add(xl);
                            vars.Add(new XElement("row"));
                            i++;
                        }
                    }
                }
                if (elo == null) continue;
                sortXML(elo, type, dn, model);
                if (vars != null) 
                {
                    bel.Add(elo); bel.Add(vars);
                    elo = bel;
                }
                changeCountXML(elo);
                Prop.excelRun(elo, p, elem.Attribute("PartNumber").Value);
                projectPr = new XMLDoc(path, "row");
                properties.Clear();
            }
        }

        public void changeCountXML(XElement el)
        {
            foreach (var item in el.Elements())
            {
                if (item.HasElements) changeCountXML(item);
                string v = MyXML.getAtt(item, "Count"), p = MyXML.getAtt(el, "Count");
                if (p != "" && v != "" && p != v)
                {
                    int i = int.Parse(p) * int.Parse(v);
                    MyXML.changeAtt(item, "Count", i.ToString());
                } 
            }
        }

        public void sortXML(XElement el, string type, string dn, string model)
        {
            string bpath = "";
            if (model == "")
            {
                bpath = MyXML.getAtt(el, "ffn");
            }
            MyXML.forElems(el, "row", a => addSort(a, type, dn, bpath));
            XMLDoc.sortRec(el, "sort", "decNum");
        }

        public void addSort(XElement el, string type, string reg, string bpath)
        {

            var pn = MyXML.getAtt(el, "PartNumber");
            string s = "5";
            string decNum = "", t = "";
            if (pn == "") s = "9";
            else if (bpath != "")
            {
                string p = file.p(bpath), ffn = MyXML.getAtt(el, "ffn");
                bool flag = ffn.StartsWith(p);
                Regex re = new Regex(@"\.(\d\d)(\.)(\d\d\d)");
                Match match = re.Match(pn);
                if (match.Groups.Count == 4)
                {
                    decNum = match.Groups[1].Value + match.Groups[3].Value;
                    if (decNum[4] != '0' && flag) s = "6";
                    else if (decNum[4] == '0' && !flag) s = "7";
                    else if (decNum[4] != '0' && !flag) s = "8";
                }
            }
            else
            {
                Regex r = new Regex(reg);
                Match m = r.Match(pn);
                if (m.Groups.Count == 5)
                {
                    t = m.Groups[1].Value + m.Groups[2].Value;
                    if (t == type) s = "0";
                    decNum = m.Groups[3].Value + m.Groups[4].Value;
                    if (decNum[0] == '0' && decNum[4] != '0' && type == t)
                        s = "6";
                }
            }
            MyXML.addAtt(el, "sort", s);
            MyXML.addAtt(el, "decNum", decNum);
        }

        void mydgv_keyEv(object sender, myKeyEventArgs e)
        {
            switch (e.Vals[0])
            {
                case "save":
                    Prop.save();
                    break;
                case "close":
                    f.Close();
                    break;
                case "open":
                    Prop.open();
                    break;
                case "add":
                    Prop.open(true);
                    break;
                default:
                    break;
            }
        }

        void mydgv_xmlEv(object sender, xmlEventArgs e)
        {
            if (e.Vals[0] == "delete")
            {
                projectPr.delete("ffn", e.Vals[1]);
            }
        }
        public bool copy(string inputPath, string name, string outputPath, string addPath)
        {
            if (copy(inputPath + addPath, name, outputPath)) return true;
            else return copy(inputPath, name, outputPath);
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
            public static string lastPath;
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
