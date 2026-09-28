using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Inventor;
using System.Windows.Forms;
using InvDoc;
using System.Xml.Linq;

namespace InvAddIn
{
    class ContenForm: Form
    {
        private ListView listView1;
        private ListView listView2;
        XMLDoc xdoc;
        string gpath = "";
        List<string> pathes = null;
    
        public ContenForm()
        {
            string p = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
            xdoc = new XMLDoc(p + "\\ContentLibrary.xml", "Content");
            gpath = XMLDoc.getAttributeValue(xdoc.El, "path", "");
            pathes = new List<string>();
            InitializeComponent();
            init(listView1, xdoc.El, "name");
        }

        public void init(ListView lv, XElement el, string name)
        {
            foreach (var item in el.Elements())
            {
                string v = XMLDoc.getAttributeValue(item, name);
                pathes.Add(XMLDoc.getAttributeValue(item, "path", ""));
                lv.Items.Add(v);
            }
        }

        public void run(string fn)
        {
            AssemblyComponentDefinition acd = I.getACD(I.aDoc());
            CommandManager cmd = I.app.CommandManager;
            cmd.PostPrivateEvent(PrivateEventTypeEnum.kFileNameEvent, fn);
            ControlDefinition control = cmd.ControlDefinitions["AssemblyPlaceComponentCmd"];
            control.Execute2(true);
            this.Close();
        }

        private void InitializeComponent()
        {
            this.listView1 = new System.Windows.Forms.ListView();
            this.listView2 = new System.Windows.Forms.ListView();
            this.SuspendLayout();
            // 
            // listView1
            // 
            this.listView1.Location = new System.Drawing.Point(3, 3);
            this.listView1.Name = "listView1";
            this.listView1.Size = new System.Drawing.Size(513, 661);
            this.listView1.TabIndex = 0;
            this.listView1.UseCompatibleStateImageBehavior = false;
            this.listView1.View = System.Windows.Forms.View.List;
            this.listView1.Click += new System.EventHandler(this.listView1_Click);
            // 
            // listView2
            // 
            this.listView2.Location = new System.Drawing.Point(522, 3);
            this.listView2.Name = "listView2";
            this.listView2.Size = new System.Drawing.Size(681, 661);
            this.listView2.TabIndex = 1;
            this.listView2.UseCompatibleStateImageBehavior = false;
            this.listView2.View = System.Windows.Forms.View.List;
            this.listView2.Click += new System.EventHandler(this.listView2_Click);
            // 
            // ContenForm
            // 
            this.ClientSize = new System.Drawing.Size(1206, 666);
            this.Controls.Add(this.listView2);
            this.Controls.Add(this.listView1);
            this.Name = "ContenForm";
            this.ResumeLayout(false);

        }

        private void listView1_Click(object sender, EventArgs e)
        {
            ListView lv = (ListView)sender;
            if (lv.SelectedItems.Count == 0) return;
            listView2.Clear();
            string txt = lv.SelectedItems[0].Text;
            XElement el = XMLDoc.find("name", txt, xdoc.El);
            init(listView2, el, "Description");
        }

        private void listView2_Click(object sender, EventArgs e)
        {
            ListView lv = (ListView)sender;
            if (lv.SelectedItems.Count == 0) return;
            string txt = lv.SelectedItems[0].Text;
            int i = lv.SelectedItems[0].Index;
            XElement el = XMLDoc.find("Description", txt, xdoc.El);
            string fn = gpath + pathes[i] + XMLDoc.getAttributeValue(el, "fn", "");
            if (!System.IO.File.Exists(fn)) return;
            run(fn);
        }

    }

    internal class ContenBtn : Button
    {
        public static ContenForm m_ccForm;
        public static string name = "", typ = "", offset = "";
        public Inventor.Document pDoc { get; set; }
        public static ContenForm getIMate
        {
            get
            {
                return m_ccForm;
            }
        }

        #region "Methods"
        public ContenBtn(string displayName, string internalName, string clientId, string description, string tooltip,
            System.Drawing.Icon standardIcon, System.Drawing.Icon largeIcon)
            : base(displayName, internalName, clientId, description, tooltip, standardIcon, largeIcon)
        {
        }

        protected override void ButtonDefinition_OnExecute(NameValueMap context)
        {
            m_ccForm = new InvAddIn.ContenForm();
            Macros.StandardAddInServer.forms.Add(m_ccForm);
            //InterfaceDll.MyEvents ev = new InterfaceDll.MyEvents(m_IMate);
            //ev.addKeyEvent();
            if (Macros.StandardAddInServer.activeteForm()) System.Windows.Forms.Application.Run(m_ccForm);
        }

        #endregion
    }
}
