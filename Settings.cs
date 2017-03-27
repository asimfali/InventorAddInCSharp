using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Inventor;
using InvDoc;


namespace InvDoc
{
    public class Settings : Form
    {
        private Container components;
        private CheckedListBox lstbox;
        private System.Windows.Forms.Button btn;
        private Label lbl;
        System.Collections.Generic.List<string> strs;
        List<string> attr;
        InvDoc.XML reader;

        public Settings()
        {
            try
            {
                InitializeComponent();
                reader = new XML(@"C:\ProgramData\Autodesk\Inventor Addins\Modules.xml");
                strs = new System.Collections.Generic.List<string>();
                attr = new List<string>();
                strs = reader.ReadXML("Spec", ref attr);
                string[] names = attr.ToArray();
                lstbox.Items.AddRange(names);
                lstbox.CheckOnClick = true;
                bool check;
                for (int i = 0; i < lstbox.Items.Count; i++)
                {
                    if (strs[i] == "1") check = true;
                    else check = false;
                    lstbox.SetItemChecked(i, check);
                }
                reader.Save();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                if (components != null)
                {
                    components.Dispose();
                }
            }
            base.Dispose(disposing);
        }


        private void InitializeComponent()
        {
            this.lbl = new System.Windows.Forms.Label();
            this.lstbox = new System.Windows.Forms.CheckedListBox();
            this.btn = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // lbl
            // 
            this.lbl.Location = new System.Drawing.Point(10, 10);
            this.lbl.Name = "lbl";
            this.lbl.Size = new System.Drawing.Size(250, 20);
            this.lbl.TabIndex = 1;
            this.lbl.Text = "Выберите нужные модули:";
            // 
            // lstbox
            // 
            this.lstbox.Location = new System.Drawing.Point(10, 30);
            this.lstbox.Name = "lstbox";
            this.lstbox.Size = new System.Drawing.Size(300, 289);
            this.lstbox.TabIndex = 0;
            // 
            // btn
            // 
            this.btn.Location = new System.Drawing.Point(134, 334);
            this.btn.Name = "btn";
            this.btn.Size = new System.Drawing.Size(50, 30);
            this.btn.TabIndex = 1;
            this.btn.Text = "OK";
            this.btn.Click += new System.EventHandler(this.btn_Click);
            // 
            // Settings
            // 
            this.ClientSize = new System.Drawing.Size(320, 376);
            this.Controls.Add(this.lstbox);
            this.Controls.Add(this.lbl);
            this.Controls.Add(this.btn);
            this.Name = "Settings";
            this.ResumeLayout(false);

        }

        private void btn_Click(object sender, EventArgs e)
        {
            try
            {
                bool check;
                for (int i = 0; i < lstbox.Items.Count; i++)
                {
                    check = lstbox.GetItemChecked(i);
                    if (check) strs[i] = "1";
                    else strs[i] = "0";
                }
                reader.WriteXML(strs);
                reader.Save();
                Macros.StandardAddInServer.settings();
                this.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

    }

//     internal class SetButton : Button
//     {
//         public static Settings m_settings;
//         public Inventor.Document pDoc { get; set; }
//         public static Settings getSettings
//         {
//             get
//             {
//                 return m_settings;
//             }
//         }
// 
//         #region "Methods"
//         public SetButton(string displayName, string internalName, CommandTypesEnum commandType, string clientId, string description, string tooltip, Icon standardIcon, Icon largeIcon, ButtonDisplayEnum buttonDisplayType)
//             : base(displayName, internalName, commandType, clientId, description, tooltip, standardIcon, largeIcon, buttonDisplayType)
//         {
// 
//         }
//         public SetButton(string displayName, string internalName, CommandTypesEnum commandType, string clientId, string description, string tooltip, ButtonDisplayEnum buttonDisplayType)
//             : base(displayName, internalName, commandType, clientId, description, tooltip, buttonDisplayType)
//         {
// 
//         }
// 
//         protected override void ButtonDefinition_OnExecute(NameValueMap context)
//         {
//             Macros.StandardAddInServer.forms.Add(m_settings);
//             if (Macros.StandardAddInServer.activeteForm()) System.Windows.Forms.Application.Run(m_settings = new Settings());
//         }
// 
//         #endregion
//     }    
}
