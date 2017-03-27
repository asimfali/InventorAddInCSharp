using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Inventor;
using System.Xml;
using System.Xml.Linq;
using InvDoc;

namespace InvAddIn
{
    public partial class sheet : Form
    {
        DrawingDocument drw;
        InvDoc.InvDocument<Document> invDoc;
        XMLDoc xmlDoc;
        public sheet()
        {
            drw = (DrawingDocument)Macros.StandardAddInServer.m_inventorApplication.ActiveDocument;
            invDoc = new InvDoc.InvDocument<Document>((Document)drw);
            invDoc.doc = (Document)drw;
            InitializeComponent();
            if (System.IO.File.Exists(@"C:\ProgramData\Autodesk\Inventor Addins\sheet.xml"))
            {
                xmlDoc = new XMLDoc(@"C:\ProgramData\Autodesk\Inventor Addins\sheet.xml", "head");
                foreach (var item in xmlDoc.Doc.Descendants("w"))
                {
                    comboBox1.Items.Add(item.Value);
                }
                if (comboBox1.Items.Count != 0)
                {
                    comboBox1.Text = comboBox1.Items[0].ToString();
                }
                foreach (var item in xmlDoc.Doc.Descendants("h"))
                {
                    comboBox2.Items.Add(item.Value);
                }
                if (comboBox2.Items.Count != 0)
                {
                    comboBox2.Text = comboBox2.Items[0].ToString();
                }
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            invDoc.addSheet(double.Parse(comboBox1.Text), double.Parse(comboBox2.Text));
        }

        private void button2_Click(object sender, EventArgs e)
        {
            
            invDoc.changeSheet(drw.ActiveSheet,double.Parse(comboBox1.Text), double.Parse(comboBox2.Text));
        }
    }
}
