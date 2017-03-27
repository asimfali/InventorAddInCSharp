namespace InvAddIn
{
    partial class Tables
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
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle1 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle2 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle3 = new System.Windows.Forms.DataGridViewCellStyle();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.CountColumn = new System.Windows.Forms.TextBox();
            this.CountRow = new System.Windows.Forms.TextBox();
            this.ColumnDimensions = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.RowDimensions = new System.Windows.Forms.TextBox();
            this._dgv = new System.Windows.Forms.DataGridView();
            this.contextMenuStrip1 = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.имяШрифтаToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripTextBox1 = new System.Windows.Forms.ToolStripTextBox();
            this.высотаТекстаToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripMenuItem2 = new System.Windows.Forms.ToolStripComboBox();
            this.курсивToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.даToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.нетToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.подчеркиваниеToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.даToolStripMenuItem1 = new System.Windows.Forms.ToolStripMenuItem();
            this.нетToolStripMenuItem1 = new System.Windows.Forms.ToolStripMenuItem();
            this.жирныйToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.даToolStripMenuItem2 = new System.Windows.Forms.ToolStripMenuItem();
            this.нетToolStripMenuItem2 = new System.Windows.Forms.ToolStripMenuItem();
            this.выравниваниеToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.поЦентруToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.справаToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.слеваToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.сжатиеToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripComboBox1 = new System.Windows.Forms.ToolStripComboBox();
            this.button1 = new System.Windows.Forms.Button();
            this.comboBox1 = new System.Windows.Forms.ComboBox();
            this.label5 = new System.Windows.Forms.Label();
            this.label7 = new System.Windows.Forms.Label();
            this.textBox2 = new System.Windows.Forms.TextBox();
            this.menuStrip1 = new System.Windows.Forms.MenuStrip();
            this.файлToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.сохранитьToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.открытьToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.удалитьToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.настройкиToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            ((System.ComponentModel.ISupportInitialize)(this._dgv)).BeginInit();
            this.contextMenuStrip1.SuspendLayout();
            this.menuStrip1.SuspendLayout();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(22, 45);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(119, 13);
            this.label1.TabIndex = 0;
            this.label1.Text = "Количество столбцов:\r\n";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(203, 45);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(101, 13);
            this.label2.TabIndex = 0;
            this.label2.Text = "Количество строк:";
            // 
            // CountColumn
            // 
            this.CountColumn.Location = new System.Drawing.Point(147, 42);
            this.CountColumn.Name = "CountColumn";
            this.CountColumn.Size = new System.Drawing.Size(50, 20);
            this.CountColumn.TabIndex = 1;
            this.CountColumn.Text = "1";
            this.CountColumn.Leave += new System.EventHandler(this.CountColumn_TextChanged);
            // 
            // CountRow
            // 
            this.CountRow.Location = new System.Drawing.Point(310, 42);
            this.CountRow.Name = "CountRow";
            this.CountRow.Size = new System.Drawing.Size(50, 20);
            this.CountRow.TabIndex = 1;
            this.CountRow.Text = "1";
            this.CountRow.Leave += new System.EventHandler(this.CountRow_TextChanged);
            // 
            // ColumnDimensions
            // 
            this.ColumnDimensions.Location = new System.Drawing.Point(489, 42);
            this.ColumnDimensions.Name = "ColumnDimensions";
            this.ColumnDimensions.Size = new System.Drawing.Size(340, 20);
            this.ColumnDimensions.TabIndex = 2;
            this.ColumnDimensions.Leave += new System.EventHandler(this.ColumnDimensions_Leave);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(376, 45);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(107, 13);
            this.label3.TabIndex = 0;
            this.label3.Text = "Размеры столбцов:";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(839, 45);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(89, 13);
            this.label4.TabIndex = 0;
            this.label4.Text = "Размеры строк:";
            // 
            // RowDimensions
            // 
            this.RowDimensions.Location = new System.Drawing.Point(934, 42);
            this.RowDimensions.Name = "RowDimensions";
            this.RowDimensions.Size = new System.Drawing.Size(340, 20);
            this.RowDimensions.TabIndex = 2;
            this.RowDimensions.Leave += new System.EventHandler(this.RowDimensions_Leave);
            // 
            // _dgv
            // 
            this._dgv.BackgroundColor = System.Drawing.SystemColors.ControlLightLight;
            this._dgv.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing;
            this._dgv.ContextMenuStrip = this.contextMenuStrip1;
            dataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter;
            dataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.Window;
            dataGridViewCellStyle1.Font = new System.Drawing.Font("GOST Common", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            dataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.ControlText;
            dataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.False;
            this._dgv.DefaultCellStyle = dataGridViewCellStyle1;
            this._dgv.GridColor = System.Drawing.SystemColors.ControlText;
            this._dgv.Location = new System.Drawing.Point(9, 150);
            this._dgv.Margin = new System.Windows.Forms.Padding(0);
            this._dgv.Name = "_dgv";
            dataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter;
            dataGridViewCellStyle2.BackColor = System.Drawing.SystemColors.Window;
            dataGridViewCellStyle2.Font = new System.Drawing.Font("GOST Common", 3.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Millimeter, ((byte)(204)));
            dataGridViewCellStyle2.ForeColor = System.Drawing.SystemColors.WindowText;
            dataGridViewCellStyle2.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle2.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this._dgv.RowHeadersDefaultCellStyle = dataGridViewCellStyle2;
            dataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter;
            dataGridViewCellStyle3.Font = new System.Drawing.Font("GOST Common", 9.749999F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this._dgv.RowsDefaultCellStyle = dataGridViewCellStyle3;
            this._dgv.RowTemplate.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter;
            this._dgv.RowTemplate.DefaultCellStyle.Font = new System.Drawing.Font("GOST Common", 12F, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this._dgv.RowTemplate.Height = 40;
            this._dgv.Size = new System.Drawing.Size(1266, 714);
            this._dgv.TabIndex = 3;
            this._dgv.RowsAdded += new System.Windows.Forms.DataGridViewRowsAddedEventHandler(this._dgv_RowsAdded);
            this._dgv.KeyDown += new System.Windows.Forms.KeyEventHandler(this._dgv_KeyDown);
            // 
            // contextMenuStrip1
            // 
            this.contextMenuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.имяШрифтаToolStripMenuItem,
            this.высотаТекстаToolStripMenuItem,
            this.курсивToolStripMenuItem,
            this.подчеркиваниеToolStripMenuItem,
            this.жирныйToolStripMenuItem,
            this.выравниваниеToolStripMenuItem,
            this.сжатиеToolStripMenuItem});
            this.contextMenuStrip1.Name = "contextMenuStrip1";
            this.contextMenuStrip1.Size = new System.Drawing.Size(155, 158);
            // 
            // имяШрифтаToolStripMenuItem
            // 
            this.имяШрифтаToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.toolStripTextBox1});
            this.имяШрифтаToolStripMenuItem.Name = "имяШрифтаToolStripMenuItem";
            this.имяШрифтаToolStripMenuItem.Size = new System.Drawing.Size(154, 22);
            this.имяШрифтаToolStripMenuItem.Text = "Имя шрифта";
            // 
            // toolStripTextBox1
            // 
            this.toolStripTextBox1.Name = "toolStripTextBox1";
            this.toolStripTextBox1.Size = new System.Drawing.Size(100, 21);
            this.toolStripTextBox1.Text = "GOST Common";
            this.toolStripTextBox1.DoubleClick += new System.EventHandler(this.toolStripTextBox1_DoubleClick);
            // 
            // высотаТекстаToolStripMenuItem
            // 
            this.высотаТекстаToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.toolStripMenuItem2});
            this.высотаТекстаToolStripMenuItem.Name = "высотаТекстаToolStripMenuItem";
            this.высотаТекстаToolStripMenuItem.Size = new System.Drawing.Size(154, 22);
            this.высотаТекстаToolStripMenuItem.Text = "Высота текста";
            this.высотаТекстаToolStripMenuItem.DropDownOpening += new System.EventHandler(this.высотаТекстаToolStripMenuItem_DropDownOpening);
            this.высотаТекстаToolStripMenuItem.Click += new System.EventHandler(this.высотаТекстаToolStripMenuItem_Click);
            // 
            // toolStripMenuItem2
            // 
            this.toolStripMenuItem2.Items.AddRange(new object[] {
            "2",
            "2.5",
            "3",
            "3.5",
            "4",
            "4.5",
            "5",
            "5.5",
            "6",
            "6.5",
            "7",
            "8",
            "10"});
            this.toolStripMenuItem2.Name = "toolStripMenuItem2";
            this.toolStripMenuItem2.Size = new System.Drawing.Size(152, 21);
            this.toolStripMenuItem2.Text = "3,5";
            this.toolStripMenuItem2.Enter += new System.EventHandler(this.toolStripMenuItem2_Enter);
            this.toolStripMenuItem2.TextChanged += new System.EventHandler(this.toolStripMenuItem2_OwnerChanged);
            // 
            // курсивToolStripMenuItem
            // 
            this.курсивToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.даToolStripMenuItem,
            this.нетToolStripMenuItem});
            this.курсивToolStripMenuItem.Name = "курсивToolStripMenuItem";
            this.курсивToolStripMenuItem.Size = new System.Drawing.Size(154, 22);
            this.курсивToolStripMenuItem.Text = "Курсив";
            // 
            // даToolStripMenuItem
            // 
            this.даToolStripMenuItem.Name = "даToolStripMenuItem";
            this.даToolStripMenuItem.Size = new System.Drawing.Size(93, 22);
            this.даToolStripMenuItem.Text = "Да";
            this.даToolStripMenuItem.Click += new System.EventHandler(this.даToolStripMenuItem_Click);
            // 
            // нетToolStripMenuItem
            // 
            this.нетToolStripMenuItem.Name = "нетToolStripMenuItem";
            this.нетToolStripMenuItem.Size = new System.Drawing.Size(93, 22);
            this.нетToolStripMenuItem.Text = "Нет";
            this.нетToolStripMenuItem.Click += new System.EventHandler(this.нетToolStripMenuItem_Click);
            // 
            // подчеркиваниеToolStripMenuItem
            // 
            this.подчеркиваниеToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.даToolStripMenuItem1,
            this.нетToolStripMenuItem1});
            this.подчеркиваниеToolStripMenuItem.Name = "подчеркиваниеToolStripMenuItem";
            this.подчеркиваниеToolStripMenuItem.Size = new System.Drawing.Size(154, 22);
            this.подчеркиваниеToolStripMenuItem.Text = "Подчеркивание";
            // 
            // даToolStripMenuItem1
            // 
            this.даToolStripMenuItem1.Name = "даToolStripMenuItem1";
            this.даToolStripMenuItem1.Size = new System.Drawing.Size(93, 22);
            this.даToolStripMenuItem1.Text = "Да";
            this.даToolStripMenuItem1.Click += new System.EventHandler(this.даToolStripMenuItem1_Click);
            // 
            // нетToolStripMenuItem1
            // 
            this.нетToolStripMenuItem1.Name = "нетToolStripMenuItem1";
            this.нетToolStripMenuItem1.Size = new System.Drawing.Size(93, 22);
            this.нетToolStripMenuItem1.Text = "Нет";
            this.нетToolStripMenuItem1.Click += new System.EventHandler(this.нетToolStripMenuItem1_Click);
            // 
            // жирныйToolStripMenuItem
            // 
            this.жирныйToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.даToolStripMenuItem2,
            this.нетToolStripMenuItem2});
            this.жирныйToolStripMenuItem.Name = "жирныйToolStripMenuItem";
            this.жирныйToolStripMenuItem.Size = new System.Drawing.Size(154, 22);
            this.жирныйToolStripMenuItem.Text = "Жирный";
            // 
            // даToolStripMenuItem2
            // 
            this.даToolStripMenuItem2.Name = "даToolStripMenuItem2";
            this.даToolStripMenuItem2.Size = new System.Drawing.Size(93, 22);
            this.даToolStripMenuItem2.Text = "Да";
            this.даToolStripMenuItem2.Click += new System.EventHandler(this.даToolStripMenuItem2_Click);
            // 
            // нетToolStripMenuItem2
            // 
            this.нетToolStripMenuItem2.Name = "нетToolStripMenuItem2";
            this.нетToolStripMenuItem2.Size = new System.Drawing.Size(93, 22);
            this.нетToolStripMenuItem2.Text = "Нет";
            // 
            // выравниваниеToolStripMenuItem
            // 
            this.выравниваниеToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.поЦентруToolStripMenuItem,
            this.справаToolStripMenuItem,
            this.слеваToolStripMenuItem});
            this.выравниваниеToolStripMenuItem.Name = "выравниваниеToolStripMenuItem";
            this.выравниваниеToolStripMenuItem.Size = new System.Drawing.Size(154, 22);
            this.выравниваниеToolStripMenuItem.Text = "Выравнивание";
            // 
            // поЦентруToolStripMenuItem
            // 
            this.поЦентруToolStripMenuItem.Name = "поЦентруToolStripMenuItem";
            this.поЦентруToolStripMenuItem.Size = new System.Drawing.Size(126, 22);
            this.поЦентруToolStripMenuItem.Text = "По центру";
            this.поЦентруToolStripMenuItem.Click += new System.EventHandler(this.поЦентруToolStripMenuItem_Click);
            // 
            // справаToolStripMenuItem
            // 
            this.справаToolStripMenuItem.Name = "справаToolStripMenuItem";
            this.справаToolStripMenuItem.Size = new System.Drawing.Size(126, 22);
            this.справаToolStripMenuItem.Text = "Справа";
            this.справаToolStripMenuItem.Click += new System.EventHandler(this.справаToolStripMenuItem_Click);
            // 
            // слеваToolStripMenuItem
            // 
            this.слеваToolStripMenuItem.Name = "слеваToolStripMenuItem";
            this.слеваToolStripMenuItem.Size = new System.Drawing.Size(126, 22);
            this.слеваToolStripMenuItem.Text = "Слева";
            this.слеваToolStripMenuItem.Click += new System.EventHandler(this.слеваToolStripMenuItem_Click);
            // 
            // сжатиеToolStripMenuItem
            // 
            this.сжатиеToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.toolStripComboBox1});
            this.сжатиеToolStripMenuItem.Name = "сжатиеToolStripMenuItem";
            this.сжатиеToolStripMenuItem.Size = new System.Drawing.Size(154, 22);
            this.сжатиеToolStripMenuItem.Text = "Сжатие";
            this.сжатиеToolStripMenuItem.DropDownOpening += new System.EventHandler(this.сжатиеToolStripMenuItem_DropDownOpening);
            // 
            // toolStripComboBox1
            // 
            this.toolStripComboBox1.Items.AddRange(new object[] {
            "100",
            "95",
            "90",
            "85",
            "80",
            "75",
            "70",
            "65",
            "60",
            "55",
            "50"});
            this.toolStripComboBox1.Name = "toolStripComboBox1";
            this.toolStripComboBox1.Size = new System.Drawing.Size(121, 21);
            this.toolStripComboBox1.Text = "100";
            this.toolStripComboBox1.TextChanged += new System.EventHandler(this.toolStripComboBox1_TextChanged);
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(1163, 83);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(109, 46);
            this.button1.TabIndex = 4;
            this.button1.Text = "Разместить на чертеже";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // comboBox1
            // 
            this.comboBox1.FormattingEnabled = true;
            this.comboBox1.Location = new System.Drawing.Point(130, 80);
            this.comboBox1.Name = "comboBox1";
            this.comboBox1.Size = new System.Drawing.Size(374, 21);
            this.comboBox1.TabIndex = 5;
            this.comboBox1.SelectedIndexChanged += new System.EventHandler(this.comboBox1_SelectedIndexChanged);
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(22, 82);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(102, 13);
            this.label5.TabIndex = 0;
            this.label5.Text = "Готовые шаблоны:";
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(560, 83);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(78, 13);
            this.label7.TabIndex = 0;
            this.label7.Text = "Имя таблицы:";
            // 
            // textBox2
            // 
            this.textBox2.Location = new System.Drawing.Point(644, 79);
            this.textBox2.Name = "textBox2";
            this.textBox2.Size = new System.Drawing.Size(241, 20);
            this.textBox2.TabIndex = 1;
            this.textBox2.Text = "Пользовательская";
            this.textBox2.Leave += new System.EventHandler(this.CountColumn_TextChanged);
            // 
            // menuStrip1
            // 
            this.menuStrip1.BackColor = System.Drawing.SystemColors.ControlLight;
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.файлToolStripMenuItem,
            this.настройкиToolStripMenuItem});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Size = new System.Drawing.Size(1284, 24);
            this.menuStrip1.TabIndex = 6;
            this.menuStrip1.Text = "menuStrip1";
            // 
            // файлToolStripMenuItem
            // 
            this.файлToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.сохранитьToolStripMenuItem,
            this.открытьToolStripMenuItem,
            this.удалитьToolStripMenuItem});
            this.файлToolStripMenuItem.Name = "файлToolStripMenuItem";
            this.файлToolStripMenuItem.Size = new System.Drawing.Size(45, 20);
            this.файлToolStripMenuItem.Text = "Файл";
            // 
            // сохранитьToolStripMenuItem
            // 
            this.сохранитьToolStripMenuItem.Name = "сохранитьToolStripMenuItem";
            this.сохранитьToolStripMenuItem.Size = new System.Drawing.Size(129, 22);
            this.сохранитьToolStripMenuItem.Text = "Сохранить";
            this.сохранитьToolStripMenuItem.Click += new System.EventHandler(this.сохранитьToolStripMenuItem_Click);
            // 
            // открытьToolStripMenuItem
            // 
            this.открытьToolStripMenuItem.Name = "открытьToolStripMenuItem";
            this.открытьToolStripMenuItem.Size = new System.Drawing.Size(129, 22);
            this.открытьToolStripMenuItem.Text = "Открыть";
            this.открытьToolStripMenuItem.Click += new System.EventHandler(this.открытьToolStripMenuItem_Click);
            // 
            // удалитьToolStripMenuItem
            // 
            this.удалитьToolStripMenuItem.Name = "удалитьToolStripMenuItem";
            this.удалитьToolStripMenuItem.Size = new System.Drawing.Size(129, 22);
            this.удалитьToolStripMenuItem.Text = "Удалить";
            this.удалитьToolStripMenuItem.Click += new System.EventHandler(this.удалитьToolStripMenuItem_Click);
            // 
            // настройкиToolStripMenuItem
            // 
            this.настройкиToolStripMenuItem.Name = "настройкиToolStripMenuItem";
            this.настройкиToolStripMenuItem.Size = new System.Drawing.Size(73, 20);
            this.настройкиToolStripMenuItem.Text = "Настройки";
            this.настройкиToolStripMenuItem.Click += new System.EventHandler(this.настройкиToolStripMenuItem_Click);
            // 
            // Tables
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1284, 873);
            this.Controls.Add(this.comboBox1);
            this.Controls.Add(this.button1);
            this.Controls.Add(this._dgv);
            this.Controls.Add(this.RowDimensions);
            this.Controls.Add(this.ColumnDimensions);
            this.Controls.Add(this.CountRow);
            this.Controls.Add(this.textBox2);
            this.Controls.Add(this.CountColumn);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.label7);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.menuStrip1);
            this.MainMenuStrip = this.menuStrip1;
            this.Name = "Tables";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Tables";
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.Tables_FormClosed);
            ((System.ComponentModel.ISupportInitialize)(this._dgv)).EndInit();
            this.contextMenuStrip1.ResumeLayout(false);
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox CountColumn;
        private System.Windows.Forms.TextBox CountRow;
        private System.Windows.Forms.TextBox ColumnDimensions;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.TextBox RowDimensions;
        private System.Windows.Forms.DataGridView _dgv;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.ComboBox comboBox1;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.TextBox textBox2;
        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.ToolStripMenuItem файлToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem сохранитьToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem открытьToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem удалитьToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem настройкиToolStripMenuItem;
        private System.Windows.Forms.ContextMenuStrip contextMenuStrip1;
        private System.Windows.Forms.ToolStripMenuItem выравниваниеToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem поЦентруToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem справаToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem слеваToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem высотаТекстаToolStripMenuItem;
        private System.Windows.Forms.ToolStripComboBox toolStripMenuItem2;
        private System.Windows.Forms.ToolStripMenuItem курсивToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem даToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem нетToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem подчеркиваниеToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem даToolStripMenuItem1;
        private System.Windows.Forms.ToolStripMenuItem нетToolStripMenuItem1;
        private System.Windows.Forms.ToolStripMenuItem жирныйToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem даToolStripMenuItem2;
        private System.Windows.Forms.ToolStripMenuItem нетToolStripMenuItem2;
        private System.Windows.Forms.ToolStripMenuItem имяШрифтаToolStripMenuItem;
        private System.Windows.Forms.ToolStripTextBox toolStripTextBox1;
        private System.Windows.Forms.ToolStripMenuItem сжатиеToolStripMenuItem;
        private System.Windows.Forms.ToolStripComboBox toolStripComboBox1;
    }
}