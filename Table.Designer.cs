using System.Drawing;
using System.Windows.Forms;
using Inventor;
namespace InvAddIn
{
    partial class Table
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;
        

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            this.dgv = new System.Windows.Forms.DataGridView();
            this.Format = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Zone = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Position = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.PartNumber = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Description = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Count = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Note = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.contextMenuStrip1 = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.VariableToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.subVariableToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.button1 = new System.Windows.Forms.Button();
            this.button3 = new System.Windows.Forms.Button();
            this.txtBox = new System.Windows.Forms.TextBox();
            this.lbl = new System.Windows.Forms.Label();
            this.menu = new System.Windows.Forms.MenuStrip();
            this.fileToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.safeToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.clearToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.переменныеДанныеToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.экспортВXMLToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.импортИзXMLToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.dgvVar = new System.Windows.Forms.DataGridView();
            this.FormatVar = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ZoneVar = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.PositionVar = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.PartNumberVar = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.DescriptionVar = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.CountVar = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.NoteVar = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.contextMenuStrip2 = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.загрузитьДанныеToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.KitsToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.button2 = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.dgv)).BeginInit();
            this.contextMenuStrip1.SuspendLayout();
            this.menu.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgvVar)).BeginInit();
            this.contextMenuStrip2.SuspendLayout();
            this.SuspendLayout();
            // 
            // dgv
            // 
            this.dgv.BackgroundColor = System.Drawing.SystemColors.Control;
            this.dgv.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgv.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.Format,
            this.Zone,
            this.Position,
            this.PartNumber,
            this.Description,
            this.Count,
            this.Note});
            this.dgv.ContextMenuStrip = this.contextMenuStrip1;
            this.dgv.Location = new System.Drawing.Point(16, 26);
            this.dgv.Name = "dgv";
            this.dgv.Size = new System.Drawing.Size(870, 557);
            this.dgv.TabIndex = 0;
            this.dgv.CellEndEdit += new System.Windows.Forms.DataGridViewCellEventHandler(this.dgv_CellEndEdit);
            this.dgv.KeyDown += new System.Windows.Forms.KeyEventHandler(this.dgv_KeyDown);
            // 
            // Format
            // 
            this.Format.HeaderText = "Формат";
            this.Format.Name = "Format";
            // 
            // Zone
            // 
            this.Zone.HeaderText = "Зона";
            this.Zone.Name = "Zone";
            // 
            // Position
            // 
            this.Position.HeaderText = "Поз.";
            this.Position.Name = "Position";
            // 
            // PartNumber
            // 
            this.PartNumber.HeaderText = "Обозначение";
            this.PartNumber.Name = "PartNumber";
            // 
            // Description
            // 
            this.Description.HeaderText = "Наименование";
            this.Description.Name = "Description";
            // 
            // Count
            // 
            this.Count.HeaderText = "Кол.";
            this.Count.Name = "Count";
            // 
            // Note
            // 
            this.Note.HeaderText = "Прим.";
            this.Note.Name = "Note";
            // 
            // contextMenuStrip1
            // 
            this.contextMenuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.VariableToolStripMenuItem});
            this.contextMenuStrip1.Name = "contextMenuStrip1";
            this.contextMenuStrip1.Size = new System.Drawing.Size(241, 26);
            // 
            // VariableToolStripMenuItem
            // 
            this.VariableToolStripMenuItem.Name = "VariableToolStripMenuItem";
            this.VariableToolStripMenuItem.Size = new System.Drawing.Size(240, 22);
            this.VariableToolStripMenuItem.Text = "Добавить в переменные данные";
            this.VariableToolStripMenuItem.Click += new System.EventHandler(this.subVariable_Click);
            // 
            // subVariableToolStripMenuItem
            // 
            this.subVariableToolStripMenuItem.Name = "subVariableToolStripMenuItem";
            this.subVariableToolStripMenuItem.Size = new System.Drawing.Size(224, 22);
            this.subVariableToolStripMenuItem.Text = "Добавить в переменные данные";
            this.subVariableToolStripMenuItem.Click += new System.EventHandler(this.subVariable_Click);
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(1164, 44);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(89, 38);
            this.button1.TabIndex = 1;
            this.button1.Text = "Разместить на чертеже";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // button3
            // 
            this.button3.Location = new System.Drawing.Point(960, 50);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(89, 38);
            this.button3.TabIndex = 1;
            this.button3.Text = "Добавить исполнение";
            this.button3.UseVisualStyleBackColor = true;
            this.button3.Click += new System.EventHandler(this.Variable_Click);
            // 
            // txtBox
            // 
            this.txtBox.Location = new System.Drawing.Point(0, 0);
            this.txtBox.Name = "txtBox";
            this.txtBox.Size = new System.Drawing.Size(100, 20);
            this.txtBox.TabIndex = 1;
            this.txtBox.Text = "28;36;40";
            // 
            // lbl
            // 
            this.lbl.AutoSize = true;
            this.lbl.Location = new System.Drawing.Point(0, 0);
            this.lbl.Name = "lbl";
            this.lbl.Size = new System.Drawing.Size(65, 13);
            this.lbl.TabIndex = 2;
            this.lbl.Text = "Разбиение:";
            // 
            // menu
            // 
            this.menu.BackColor = System.Drawing.SystemColors.Control;
            this.menu.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.fileToolStripMenuItem,
            this.переменныеДанныеToolStripMenuItem});
            this.menu.Location = new System.Drawing.Point(0, 0);
            this.menu.Name = "menu";
            this.menu.Size = new System.Drawing.Size(1912, 24);
            this.menu.TabIndex = 1;
            this.menu.Text = "menuStrip1";
            // 
            // fileToolStripMenuItem
            // 
            this.fileToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.safeToolStripMenuItem,
            this.clearToolStripMenuItem});
            this.fileToolStripMenuItem.Name = "fileToolStripMenuItem";
            this.fileToolStripMenuItem.Size = new System.Drawing.Size(45, 20);
            this.fileToolStripMenuItem.Text = "Файл";
            // 
            // safeToolStripMenuItem
            // 
            this.safeToolStripMenuItem.Name = "safeToolStripMenuItem";
            this.safeToolStripMenuItem.Size = new System.Drawing.Size(178, 22);
            this.safeToolStripMenuItem.Text = "Сохранить в модель";
            this.safeToolStripMenuItem.Click += new System.EventHandler(this.safeToolStripMenuItem_Click);
            // 
            // clearToolStripMenuItem
            // 
            this.clearToolStripMenuItem.Name = "clearToolStripMenuItem";
            this.clearToolStripMenuItem.Size = new System.Drawing.Size(178, 22);
            this.clearToolStripMenuItem.Text = "Очистить модель";
            this.clearToolStripMenuItem.Click += new System.EventHandler(this.clearToolStripMenuItem_Click);
            // 
            // переменныеДанныеToolStripMenuItem
            // 
            this.переменныеДанныеToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.экспортВXMLToolStripMenuItem,
            this.импортИзXMLToolStripMenuItem});
            this.переменныеДанныеToolStripMenuItem.Name = "переменныеДанныеToolStripMenuItem";
            this.переменныеДанныеToolStripMenuItem.Size = new System.Drawing.Size(124, 20);
            this.переменныеДанныеToolStripMenuItem.Text = "Переменные данные";
            // 
            // экспортВXMLToolStripMenuItem
            // 
            this.экспортВXMLToolStripMenuItem.Name = "экспортВXMLToolStripMenuItem";
            this.экспортВXMLToolStripMenuItem.Size = new System.Drawing.Size(147, 22);
            this.экспортВXMLToolStripMenuItem.Text = "Экспорт в XML";
            this.экспортВXMLToolStripMenuItem.Click += new System.EventHandler(this.экспортВXMLToolStripMenuItem_Click);
            // 
            // импортИзXMLToolStripMenuItem
            // 
            this.импортИзXMLToolStripMenuItem.Name = "импортИзXMLToolStripMenuItem";
            this.импортИзXMLToolStripMenuItem.Size = new System.Drawing.Size(147, 22);
            this.импортИзXMLToolStripMenuItem.Text = "Импорт из XML";
            this.импортИзXMLToolStripMenuItem.Click += new System.EventHandler(this.импортИзXMLToolStripMenuItem_Click);
            // 
            // dgvVar
            // 
            this.dgvVar.BackgroundColor = System.Drawing.SystemColors.Control;
            this.dgvVar.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgvVar.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.FormatVar,
            this.ZoneVar,
            this.PositionVar,
            this.PartNumberVar,
            this.DescriptionVar,
            this.CountVar,
            this.NoteVar});
            this.dgvVar.ContextMenuStrip = this.contextMenuStrip2;
            this.dgvVar.Location = new System.Drawing.Point(1314, 82);
            this.dgvVar.Name = "dgvVar";
            this.dgvVar.Size = new System.Drawing.Size(240, 150);
            this.dgvVar.TabIndex = 3;
            this.dgvVar.CellBeginEdit += new System.Windows.Forms.DataGridViewCellCancelEventHandler(this.dgvVar_CellBeginEdit);
            this.dgvVar.CellEndEdit += new System.Windows.Forms.DataGridViewCellEventHandler(this.dgvVar_CellEndEdit);
            this.dgvVar.RowsRemoved += new System.Windows.Forms.DataGridViewRowsRemovedEventHandler(this.dgvVar_RowsRemoved);
            this.dgvVar.KeyDown += new System.Windows.Forms.KeyEventHandler(this.dgv_KeyDown);
            // 
            // FormatVar
            // 
            this.FormatVar.HeaderText = "Формат";
            this.FormatVar.Name = "FormatVar";
            // 
            // ZoneVar
            // 
            this.ZoneVar.HeaderText = "Зона";
            this.ZoneVar.Name = "ZoneVar";
            // 
            // PositionVar
            // 
            this.PositionVar.HeaderText = "Позиция";
            this.PositionVar.Name = "PositionVar";
            // 
            // PartNumberVar
            // 
            this.PartNumberVar.HeaderText = "Обозначение";
            this.PartNumberVar.Name = "PartNumberVar";
            // 
            // DescriptionVar
            // 
            this.DescriptionVar.HeaderText = "Наименование";
            this.DescriptionVar.Name = "DescriptionVar";
            // 
            // CountVar
            // 
            this.CountVar.HeaderText = "Кол.";
            this.CountVar.Name = "CountVar";
            // 
            // NoteVar
            // 
            this.NoteVar.HeaderText = "Прим.";
            this.NoteVar.Name = "NoteVar";
            // 
            // contextMenuStrip2
            // 
            this.contextMenuStrip2.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.загрузитьДанныеToolStripMenuItem});
            this.contextMenuStrip2.Name = "contextMenuStrip2";
            this.contextMenuStrip2.Size = new System.Drawing.Size(169, 26);
            // 
            // загрузитьДанныеToolStripMenuItem
            // 
            this.загрузитьДанныеToolStripMenuItem.Name = "загрузитьДанныеToolStripMenuItem";
            this.загрузитьДанныеToolStripMenuItem.Size = new System.Drawing.Size(168, 22);
            this.загрузитьДанныеToolStripMenuItem.Text = "Загрузить данные";
            this.загрузитьДанныеToolStripMenuItem.Click += new System.EventHandler(this.загрузитьДанныеToolStripMenuItem_Click_1);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.BackColor = System.Drawing.SystemColors.Control;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.label1.Location = new System.Drawing.Point(892, 269);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(140, 20);
            this.label1.TabIndex = 4;
            this.label1.Text = "Общие данные:";
            this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.BackColor = System.Drawing.SystemColors.Control;
            this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.label2.Location = new System.Drawing.Point(1076, 286);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(190, 20);
            this.label2.TabIndex = 4;
            this.label2.Text = "Переменные данные:";
            this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // KitsToolStripMenuItem
            // 
            this.KitsToolStripMenuItem.Name = "KitsToolStripMenuItem";
            this.KitsToolStripMenuItem.Size = new System.Drawing.Size(32, 19);
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(991, 162);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(89, 38);
            this.button2.TabIndex = 1;
            this.button2.Text = "Удалить исполнение";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // Table
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1912, 500);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.dgvVar);
            this.Controls.Add(this.menu);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.button2);
            this.Controls.Add(this.button3);
            this.Controls.Add(this.dgv);
            this.Controls.Add(this.lbl);
            this.Controls.Add(this.txtBox);
            this.MainMenuStrip = this.menu;
            this.Name = "Table";
            this.Text = "Спецификация";
            this.WindowState = System.Windows.Forms.FormWindowState.Maximized;
            ((System.ComponentModel.ISupportInitialize)(this.dgv)).EndInit();
            this.contextMenuStrip1.ResumeLayout(false);
            this.menu.ResumeLayout(false);
            this.menu.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgvVar)).EndInit();
            this.contextMenuStrip2.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        
        //private void dgv_dragDrop(object sender, System.Windows.Forms.DragEventArgs e)
        //{
        //    if (_RowUnderMouseDown != null)
        //    {
        //        DataGridViewRow rowUnderMouse = FindRow(_RowUnderMouseDown.DataGridView, e.X, e.Y, true);
        //        if (rowUnderMouse != null && e.Effect == DragDropEffects.Move)
        //        {
        //            DataGridView grid = _RowUnderMouseDown.DataGridView;
        //            object data = _RowUnderMouseDown.DataBoundItem;
        //            CurrencyManager cm = (CurrencyManager)grid.BindingContext[grid.DataSource, grid.DataMember];
        //            int index = rowUnderMouse.Index;
        //            grid.Rows.Remove(_RowUnderMouseDown);
        //            cm.List.Insert(index, data);
        //        }
        //    }
        //}

        //private void dgv_dragOver(object sender, System.Windows.Forms.DragEventArgs e)
        //{
        //    throw new System.NotImplementedException();
        //}

        //private void dgv_mouseMove(object sender, System.Windows.Forms.MouseEventArgs e)
        //{
        //    throw new System.NotImplementedException();
        //}

        //private void dgv_mouseDown(object sender, System.Windows.Forms.MouseEventArgs e)
        //{
        //    throw new System.NotImplementedException();
        //}
        
        public class DataGridViewRowsReorderBehavior
        {
            private DataGridViewRow _RowUnderMouseDown;
            public string val1, val2; public int x, y;
            public DataGridViewRowsReorderBehavior(DataGridView grid)
            {
                grid.AllowDrop = true;
                grid.MouseMove += new MouseEventHandler(Grid_MouseMove);
                grid.MouseDown += new MouseEventHandler(Grid_MouseDown);
                grid.DragOver += new DragEventHandler(Grid_DragOver);
                grid.DragDrop += new DragEventHandler(Grid_DragDrop);
                grid.CellEndEdit += new DataGridViewCellEventHandler(Grid_CellEndEdit);
                grid.CellBeginEdit += new DataGridViewCellCancelEventHandler(Grid_CellBeginEdit);
            }

            private void Grid_CellBeginEdit(object sender, DataGridViewCellCancelEventArgs e)
            {
                _RowUnderMouseDown = ((DataGridView)sender).Rows[e.RowIndex];
                if (e.ColumnIndex == 3 || e.ColumnIndex == 4)
                {
                    x = e.ColumnIndex; y = e.RowIndex;
                    val1 = _RowUnderMouseDown.Cells[3].Value.ToString();
                    val2 = _RowUnderMouseDown.Cells[4].Value.ToString();
                }
            }

            private void Grid_CellEndEdit(object sender, DataGridViewCellEventArgs e)
            {
                if  (e.ColumnIndex == 3 || e.ColumnIndex == 4)
                {
                    Inventor.Application invApp = Macros.StandardAddInServer.m_inventorApplication;
                    Inventor.Document oDoc = invApp.ActiveDocument;
                    Inventor.AssemblyDocument m_AsmDoc = (AssemblyDocument)InvDoc.u.referendedDoc(oDoc as Document);
                    Inventor.BOM m_BOM = m_AsmDoc.ComponentDefinition.BOM;
                    m_BOM.StructuredViewEnabled = true;
                    if (!m_BOM.StructuredViewFirstLevelOnly)
                        m_BOM.StructuredViewFirstLevelOnly = true;
                    Inventor.BOMView m_BOMView = m_BOM.BOMViews["Структурированный"];
                    foreach (Inventor.BOMRow row in m_BOMView.BOMRows)
                    {
                        Inventor.ComponentDefinition oCompDef = row.ComponentDefinitions[1];
                        oDoc = (Inventor.Document)oCompDef.Document;
                        string partNumber = (string)oDoc.PropertySets[3][2].Value;
                        string desc = (string)oDoc.PropertySets[3][14].Value;
                        if (partNumber == val1 && desc == val2)
                        {
                            switch (x)
                            {
                                case 3:
                            oDoc.PropertySets[3][2].Value = ((DataGridView)sender)[x,y].Value.ToString();
                                    break;
                                case 4:
                            oDoc.PropertySets[3][14].Value = ((DataGridView)sender)[x,y].Value.ToString();
                                    break;
                            }
                            break;
                        }
                    }
                }
            }

            private void Grid_DragDrop(object sender, DragEventArgs e)
            {
                if (_RowUnderMouseDown != null)
                {
                    DataGridViewRow rowUnderMouse = FindRow(_RowUnderMouseDown.DataGridView, e.X, e.Y, true);
                    if (rowUnderMouse != null && e.Effect == DragDropEffects.Move)
                    {
                        DataGridView grid = _RowUnderMouseDown.DataGridView;
                        try {
                        //object data = _RowUnderMouseDown.DataBoundItem;
                        //CurrencyManager cm = (CurrencyManager)grid.BindingContext[grid.DataSource, grid.DataMember];
                        int index = rowUnderMouse.Index+1;
                        
                        //cm.List.Insert(index, data);
                        DataGridViewRow dgr = grid.Rows.SharedRow(_RowUnderMouseDown.Index);
                        int oldIndex = _RowUnderMouseDown.Index;
                        //grid.Rows.InsertCopy(oldIndex, index+1);
                        grid.Rows.Insert(index, 1);
                        for (int i = 0; i < dgr.Cells.Count; i++)
                        {
                            //grid.Rows[oldIndex].Cells[i].Value = grid.Rows[index].Cells[i].Value;
                            grid.Rows[index].Cells[i].Value = dgr.Cells[i].Value;
                        }
                        string val = "";
                        string str1 = grid[7, grid.Rows[index - 1].Index].Value.ToString();
                            string str2 =  grid[7, grid.Rows[index + 1].Index].Value.ToString();
                            val = InvAddIn.Spec.m_Tbl.compare(str1, str2);
                            grid[7, grid.Rows[index].Index].Value = val;
                        //grid.Rows.Insert(index, 2);
                            //grid.Rows.
                            if (index < oldIndex)
                        grid.Rows.RemoveAt(oldIndex+1);
                            else
                        grid.Rows.RemoveAt(oldIndex);
                             }                     
                        catch (System.Exception ee)
                        {
                            MessageBox.Show(ee.ToString());
                        }
                    }
                }
            }
            private void Grid_DragOver(object sender, DragEventArgs e)
            {
                e.Effect = DragDropEffects.Move;
            }
            private void Grid_MouseDown(object sender, MouseEventArgs e)
            {
                _RowUnderMouseDown = FindRow(sender as DataGridView, e.X, e.Y, false);
            }
            private void Grid_MouseMove(object sender, MouseEventArgs e)
            {
                if (_RowUnderMouseDown != null && (e.Button & MouseButtons.Left) == MouseButtons.Left)
                {
                    Rectangle rect = _RowUnderMouseDown.DataGridView.GetRowDisplayRectangle(_RowUnderMouseDown.Index, true);
                    if (rect != Rectangle.Empty && rect.Contains(e.X, e.Y) == false)
                        _RowUnderMouseDown.DataGridView.DoDragDrop(_RowUnderMouseDown, DragDropEffects.Move);
                }
            }

            private static DataGridViewRow FindRow(DataGridView grid, int x, int y, bool convertToLocal)
            {
                if (convertToLocal)
                {
                    System.Drawing.Point p = grid.PointToClient(new System.Drawing.Point(x, y));
                    x = p.X;
                    y = p.Y;
                }
                int index = grid.HitTest(x, y).RowIndex;
                return (index >= 0) ? grid.Rows[index] : null;
            }
        }

        #endregion




        private System.Windows.Forms.DataGridView dgv;
        private System.Windows.Forms.ContextMenuStrip contextMenuStrip1;
        private System.Windows.Forms.ToolStripMenuItem VariableToolStripMenuItem;
        private ToolStripMenuItem subVariableToolStripMenuItem;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Button button3;
        private System.Windows.Forms.TextBox txtBox;
        private Label lbl;
        //public System.Windows.Forms.TextBox txtBox3;
        //private Label lbl3;
        private MenuStrip menu;
        private ToolStripMenuItem fileToolStripMenuItem;
        private ToolStripMenuItem safeToolStripMenuItem;
        private ToolStripMenuItem clearToolStripMenuItem;
        private DataGridViewTextBoxColumn Format;
        private DataGridViewTextBoxColumn Zone;
        private DataGridViewTextBoxColumn Position;
        private DataGridViewTextBoxColumn PartNumber;
        private DataGridViewTextBoxColumn Description;
        private DataGridViewTextBoxColumn Count;
        private DataGridViewTextBoxColumn Note;
        private DataGridView dgvVar;
        private DataGridViewTextBoxColumn FormatVar;
        private DataGridViewTextBoxColumn ZoneVar;
        private DataGridViewTextBoxColumn PositionVar;
        private DataGridViewTextBoxColumn PartNumberVar;
        private DataGridViewTextBoxColumn DescriptionVar;
        private DataGridViewTextBoxColumn CountVar;
        private DataGridViewTextBoxColumn NoteVar;
        private Label label1;
        private Label label2;
        private ToolStripMenuItem переменныеДанныеToolStripMenuItem;
        private ToolStripMenuItem экспортВXMLToolStripMenuItem;
        private ToolStripMenuItem импортИзXMLToolStripMenuItem;
        private ToolStripMenuItem KitsToolStripMenuItem;
        private System.Windows.Forms.Button button2;
        private ContextMenuStrip contextMenuStrip2;
        private ToolStripMenuItem загрузитьДанныеToolStripMenuItem;
    }

}