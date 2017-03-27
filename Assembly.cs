using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using u = InvDoc.u;
using Inventor;
using InterfaceDll;
using System.Windows.Forms;

namespace InvAddIn
{
    class Assembly: Button
    {
        public AssemblyDocument doc;
        InsertConstraint ins;
        SelectSet ss;
        

        public Assembly(string displayName, string internalName, string clientId, string description, string tooltip, 
        ButtonDisplayEnum buttonDisplayType = ButtonDisplayEnum.kDisplayTextInLearningMode, CommandTypesEnum commandType = CommandTypesEnum.kNonShapeEditCmdType)
        : base(displayName, internalName, commandType, clientId, description, tooltip, buttonDisplayType) { }

        protected override void ButtonDefinition_OnExecute(NameValueMap context)
        {
            Document doc = Macros.StandardAddInServer.m_inventorApplication.ActiveDocument;
            if (doc.DocumentType == DocumentTypeEnum.kAssemblyDocumentObject)
            {
                this.doc = doc as AssemblyDocument;
                ss = this.doc.SelectSet;
                if (ss.Count == 0) return;
                getForm();
            }
        }

        public void offset(ComponentOccurrence occ, double d = 20)
        {
            ins = occ.Constraints[1] as InsertConstraint;
            if (ins == null) return;
            str(ins.Distance, d);
        }

        private string last(Parameter p)
        {
            string exp = p.Expression;
            int ind = exp.LastIndexOfAny(new char[] { '+', '-' });
            return ind < 0 ? "0": exp.Substring(ind);
        }

        private void str(Parameter p, double d)
        {
            string exp = p.Expression;
            string sing = "+";
            if (d < 0) sing = "-";
            int ind = exp.LastIndexOfAny(new char[]{'+','-'});
            if (ind > 0)
            {
                exp = exp.Remove(ind) + sing + Math.Abs(d).ToString();
                p.Expression = exp;
            }
            else
                p.Expression = exp + sing + Math.Abs(d).ToString();
        }

        public void getForm()
        {
            int offsetY = 30, offsetX = 10;
            Form f = new Form();
            f.Height = 150; f.Width = 400; f.WindowState = FormWindowState.Normal; f.Text = "Смещение"; f.StartPosition = FormStartPosition.CenterScreen;
            f.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink; f.AutoSize = true;
            System.Drawing.Point insPt = new System.Drawing.Point(5, 5);
            InterfaceDll.Lbl lbs = new Lbl(offsetX, offsetY, 100, 15, insPt, f, "Дистанция");
            InterfaceDll.CB cbs = new CB(offsetX, offsetY, 200, 15, insPt, f);
            cbs.position(lbs.last(), true);
            //cbs.last().Items.AddRange(new[] { "2.65", "2.8", "3", "4.5", "6" });
            ComponentOccurrence occ = ss[1] as ComponentOccurrence;
            if (occ == null) return;
            ins = occ.Constraints[1] as InsertConstraint;
            if (ins == null) return;


            cbs.last().Text = last(ins.Distance);
            InterfaceDll.Btn btns = new Btn(offsetX, offsetY, 100, 20, insPt, f, click, "Добавить");
            btns.center(cbs.last(), offsetY + 5);
            f.Show();
        }

        private void click(object sender, EventArgs e)
        {
            Form f = (Form)((System.Windows.Forms.Button)sender).Parent;
            ComboBox cb = f.Controls.OfType<ComboBox>().First();
            double d = u.convToDouble(cb.Text);
            u.transactStart(doc as Document, "Разнести");
            foreach (ComponentOccurrence occ in ss)
            {
                offset(occ, d);
            }
            doc.Update2();
            u.transactEnd();
            f.Close();
        }

    }
}
