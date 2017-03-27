using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace InvAddIn
{
    public partial class Setting : Form
    {
        public Setting()
        {
            InitializeComponent();
        }

        private void Setting_FormClosed(object sender, FormClosedEventArgs e)
        {
            //this.Hide();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            InvAddIn.TablesButton.m_Table.rowheight = Convert.ToInt16(this.comboBox8.Text);
            InvAddIn.TablesButton.m_Table.columnwidth = Convert.ToInt16(this.comboBox9.Text);
            this.Hide();
        }
    }
}
