using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Inventor;
using InvDoc;
using InterfaceDll;
using System.Xml.Linq;
using Color = System.Drawing.Color;

namespace InvAddIn
{
    public partial class EditProp : Form
    {
        public HashSet<Document> docs = new HashSet<Document>();
        public HashSet<String> names = new HashSet<string>();
        Rectangle bnds;
        MyDGV mydgv;

        public EditProp()
        {
            InitializeComponent();
            MyXML xml = new MyXML("EditorProps.xml", "head");
            if (xml != null && xml.getRoot().Name == "head" )
            {
                foreach (var item in xml.getRoot().Elements())
                {
                    comboBox1.Items.Add(MyXML.getAtt(item, "name"));
                }
            }
            this.WindowState = FormWindowState.Maximized;
            bnds = Screen.PrimaryScreen.WorkingArea;
            mydgv = new MyDGV();
            mydgv.changePasteCopyColor = true;
            mydgv.form = this;
            var pt = new System.Drawing.Point(dataGridView1.Left, dataGridView1.Top);
            dataGridView1.Visible = false;
            mydgv.addDGV(pt, bnds.Width, bnds.Height - 50, null,
                new Dictionary<string, string> { { "ffn", "Имя файла"} }, new float[] { 0.1f});
            mydgv.removeEvents();
            mydgv.addEvents();
            mydgv.Dgv.CellValueChanged -= new DataGridViewCellEventHandler(mydgv.changeCell);
            mydgv.Dgv.CellEndEdit += new DataGridViewCellEventHandler(endCellEdit);
        }

        public void endCellEdit(object sender, EventArgs e)
        {
            DataGridView dgv = (DataGridView)sender;
            DataGridViewCellEventArgs arg = (DataGridViewCellEventArgs)e;
            dgv[arg.ColumnIndex, arg.RowIndex].Style.BackColor = Color.LightYellow;
        }

        private void открытьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            var pathO = InvDoc.u.OFD(InvDoc.u.pathUtil(I.aDoc()), "Inventor files(*.ipt, *.iam)|*.ipt;*.iam", true);
            var spl = pathO.Split('|');
            
            foreach (var item in spl)
            {
                addRow(item);
            }
        }

        public void addRow(string ffn)
        {
            var doc = I.open(ffn, false, false);
            docs.Add(doc);
            var name = file.nameWithExt(ffn);
            mydgv.addRow(new string[] { name });
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string name = comboBox1.Text;
            addColumn(name);
        }
        private void addColumn(string name)
        {
            if (names.Contains(name)) return;
            names.Add(name);
            var col = mydgv.addDGVCol(name, name, 0.1f);
            mydgv.Dgv.Columns.Add(col);
            foreach (DataGridViewRow item in mydgv.Dgv.Rows)
            {
                var tmp = item.Cells[0].Value;
                if (tmp == null) continue;
                string n = tmp.ToString();
                var val = getProp(n, name);
                if (val == "") continue;
                mydgv.Dgv[col.Index, item.Index].Value = val;
            }
        }
        public Document getDoc(string name)
        {
            foreach (Document item in docs)
            {
                var n = file.nameWithExt(item.FullDocumentName);
                if (n == name) return item;
            }
            return null;
        }
        public string getProp(string fn, string name)
        {
            Document doc = getDoc(fn);
            if (doc == null) return "";
            return u.getPropValue(doc, name);
        }
        public void setProp(string fn, string name, string value)
        {
            if (fn == "") return;
            Document doc = getDoc(fn);
            if (doc == null) return;
            u.addProp(doc, name, value);
        }

        private void сохранитьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            string name = "";
            foreach (DataGridViewRow row in mydgv.Dgv.Rows)
            {
                if (row.Cells[0].Value == null) continue;
                foreach (DataGridViewColumn col in mydgv.Dgv.Columns)
                {
                    string val = mydgv.Dgv[col.Index, row.Index].Value.ToString();
                    string colName = col.Name.Split(':')[0];
                    if (colName == "ffn") name = val;
                    else setProp(name, colName, val);
                }
            }
            foreach (var item in docs)
            {
                item.Save2(false);
            }
        }

        private string saveName()
        {
            string name = ""; bool first = true; string nameforsave = "";
            SaveFileDialog sfd = new SaveFileDialog();
            sfd.InitialDirectory = InvDoc.u.pathUtil(I.aDoc());
            sfd.Filter = "xml(*.xml)|*.xml";
            sfd.FilterIndex = 1;
            if (sfd.ShowDialog() == DialogResult.OK)
            {
                return sfd.FileName;
            }
            return "";
        }

        private void addDocsXML(XElement el)
        {
            foreach (var item in docs)
            {
                el.Add(new XElement("doc", new XAttribute("path", item.FullDocumentName)));
            }
        }

        public void addPropsNames(XElement el)
        {
            foreach (var item in names)
            {
                el.Add(new XElement("prop", new XAttribute("name", item)));
            }
        }

        private void fillXML(XElement el)
        {
            XElement xmlDocs = new XElement("docs");
            addDocsXML(xmlDocs);
            XElement xmlProps = new XElement("props");
            addPropsNames(xmlProps);
            el.Add(xmlDocs);
            el.Add(xmlProps);
        }

        private void сохранитьШаблонToolStripMenuItem_Click(object sender, EventArgs e)
        {
            XDocument xdoc = new XDocument();
            XElement el = new XElement("head");
            fillXML(el);
            string fn = saveName();
            if (fn == "") return;
            xdoc.Add(el);
            xdoc.Save(fn);
        }

        private void addColumns(XElement el)
        {
            foreach (var item in el.Elements())
            {
                string name = XMLDoc.getAttributeValue(item, "name", "");
                if (name != "")
                {
                    addColumn(name);
                }
            }
        }
        public void addRows(XElement el)
        {
            foreach (var item in el.Elements())
            {
                string name = XMLDoc.getAttributeValue(item, "path", "");
                if (name != "")
                {
                    addRow(name);
                }
            }
        }

        private void загрузитьИзШаблонаToolStripMenuItem_Click(object sender, EventArgs e)
        {
            var fn = u.OFD(InvDoc.u.pathUtil(I.aDoc()), "xml(*.xml)|*.xml", false);
            if (fn == "") return;
            XMLDoc xml = new XMLDoc(fn, "head");
            XElement el = XMLDoc.getXElement(xml.El, "docs");
            addRows(el);
            el = XMLDoc.getXElement(xml.El, "props");
            addColumns(el);
        }
    }
    internal class EditPropBtn : Button
    {
        public static EditProp edit_Prop;
        public static string lastPath;
        public static projectProperties prjPr;
        public Inventor.Document pDoc { get; set; }
        public static EditProp getProp
        {
            get
            {
                return edit_Prop;
            }
        }


        #region "Methods"
        public EditPropBtn(string displayName, string internalName, string clientId, string description, string tooltip,
            ButtonDisplayEnum buttonDisplayType = ButtonDisplayEnum.kDisplayTextInLearningMode, CommandTypesEnum commandType = CommandTypesEnum.kNonShapeEditCmdType)
            : base(displayName, internalName, commandType, clientId, description, tooltip, buttonDisplayType) { }

        protected override void ButtonDefinition_OnExecute(NameValueMap context)
        {
            Macros.StandardAddInServer.forms.Add(edit_Prop);
            if (Macros.StandardAddInServer.activeteForm()) System.Windows.Forms.Application.Run(edit_Prop = new InvAddIn.EditProp());
        }

        #endregion
    }
}
