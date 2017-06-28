using System;
using System.Collections.Generic;
using System.Windows.Forms;
using System.Linq;
using System.Text;
using Inventor;
using InvDoc;
using ExtensionMethods;
using System.Xml;
using System.Xml.Linq;
using System.Text.RegularExpressions;
using InterfaceDll;
using ut = InvDoc.u;

namespace InvAddIn
{

    public class CreateComponent : Form
    {
        private MenuStrip menuStrip1;
        private ToolStripMenuItem создатьФайлыToolStripMenuItem;
        private ToolStripMenuItem перенаправитьToolStripMenuItem;
        public Parts m_Parts;
        private ToolStripMenuItem разместитьВсеДеталиToolStripMenuItem;
        private MyDGV mydgv;
        DataGridView dgv;
        public System.Windows.Forms.TextBox txtbox1, txt2;
        private ToolStripMenuItem крепежToolStripMenuItem;
        public XMLDoc desc;
        public System.Windows.Forms.TextBox txtbox2;
        CheckBox checkbox;
        ComboBox cBox;
        Inventor.Application invApp = Macros.StandardAddInServer.m_inventorApplication;
        private ToolStripMenuItem чертежToolStripMenuItem;
        private ToolStripMenuItem текущийДокументToolStripMenuItem;
        private ToolStripMenuItem переназватьToolStripMenuItem;
        private ToolStripMenuItem разместитьВсеДеталиToolStripMenuItem1;
        private ToolStripMenuItem обновитьContentCenterToolStripMenuItem;
        private ToolStripMenuItem редкоеToolStripMenuItem;
        private ToolStripMenuItem добавитьСвойствоToolStripMenuItem1;
        private ToolStripMenuItem комплектФайловToolStripMenuItem1;
        private ToolStripMenuItem stepToolStripMenuItem;
        private ToolStripMenuItem шипыИПазыToolStripMenuItem;
        private ToolStripMenuItem шипToolStripMenuItem1;
        private ToolStripMenuItem пазToolStripMenuItem1;
        private ToolStripMenuItem выдавитьToolStripMenuItem1;
        private ToolStripMenuItem добавитьВСборкуToolStripMenuItem1;
        private ToolStripMenuItem центрыОтверстийToolStripMenuItem;
        private ToolStripMenuItem сплайнToolStripMenuItem;
        private ToolStripMenuItem создатьСборкуToolStripMenuItem;
        private ToolStripMenuItem сортировкаToolStripMenuItem;
        private ToolStripMenuItem артикулToolStripMenuItem;
        private ToolStripMenuItem пересобратьСборкуToolStripMenuItem;
        private ToolStripMenuItem тестToolStripMenuItem;
        private ToolStripMenuItem параметрыToolStripMenuItem;
        private ToolStripMenuItem добавитьПараметрыToolStripMenuItem1;
        private ToolStripMenuItem массивToolStripMenuItem1;
        private ToolStripMenuItem данныеДляМассиваToolStripMenuItem;
        private ToolStripMenuItem копироватьВXMLToolStripMenuItem;
        private ToolStripMenuItem структураВXMLToolStripMenuItem;
        private ToolStripMenuItem создатьДеталиToolStripMenuItem;
        private ToolStripMenuItem загрузитьXMLToolStripMenuItem;
        private ToolStripMenuItem добавитьИсполнениеToolStripMenuItem;
        private ToolStripMenuItem комплектФайловСЧертежамиToolStripMenuItem;
        private ToolStripMenuItem переименоватьСборкуСЧертежамиToolStripMenuItem;
        private ToolStripMenuItem общиеРазмерыToolStripMenuItem;
        private ToolStripMenuItem размерыДоГибовToolStripMenuItem;
        private ToolStripMenuItem позицииToolStripMenuItem;
        private ToolStripMenuItem крепежToolStripMenuItem2;
        Document doc;
        private ToolStripMenuItem заменитьИмяКПToolStripMenuItem;
        private ToolStripMenuItem детальToolStripMenuItem;
        private ToolStripMenuItem сборкаToolStripMenuItem;
        private ToolStripMenuItem addTRToolStripMenuItem1;
        private ToolStripMenuItem copyAttrToolStripMenuItem1;
        private ToolStripMenuItem кластерToolStripMenuItem;
        private ToolStripMenuItem спецToolStripMenuItem;
        private ToolStripMenuItem крепежToolStripMenuItem1;
        private ToolStripMenuItem регионToolStripMenuItem;
        private ToolStripMenuItem обновитьПозицииToolStripMenuItem;
        private ToolStripMenuItem обновитьToolStripMenuItem;
        private ToolStripMenuItem обновитьКрепежToolStripMenuItem;
        private ToolStripMenuItem получитьТекстToolStripMenuItem;
        private ToolStripMenuItem перевестиТекстToolStripMenuItem;
        private ToolStripMenuItem крепежToolStripMenuItem3;
        private ToolStripMenuItem обновитьСтилиToolStripMenuItem;
        private ToolStripMenuItem скругленияToolStripMenuItem;
        private ToolStripMenuItem ручнаяГибкаToolStripMenuItem;
        private ToolStripMenuItem переименоватьСборкуToolStripMenuItem;
        private ToolStripMenuItem габаритыToolStripMenuItem;
        private ToolStripMenuItem открытьПапкуToolStripMenuItem;
        private ToolStripMenuItem удалитьЭлементToolStripMenuItem;
        static HashSet<object> del = new HashSet<object>();

        private void InitializeComponent()
        {
            this.menuStrip1 = new System.Windows.Forms.MenuStrip();
            this.создатьФайлыToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.общиеРазмерыToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.размерыДоГибовToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.позицииToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.крепежToolStripMenuItem2 = new System.Windows.Forms.ToolStripMenuItem();
            this.чертежToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.текущийДокументToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.переназватьToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.пересобратьСборкуToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.переименоватьСборкуСЧертежамиToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.параметрыToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.добавитьПараметрыToolStripMenuItem1 = new System.Windows.Forms.ToolStripMenuItem();
            this.данныеДляМассиваToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.загрузитьXMLToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.массивToolStripMenuItem1 = new System.Windows.Forms.ToolStripMenuItem();
            this.копироватьВXMLToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.структураВXMLToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.создатьДеталиToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.разместитьВсеДеталиToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.разместитьВсеДеталиToolStripMenuItem1 = new System.Windows.Forms.ToolStripMenuItem();
            this.обновитьContentCenterToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.сортировкаToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.тестToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.шипыИПазыToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.шипToolStripMenuItem1 = new System.Windows.Forms.ToolStripMenuItem();
            this.пазToolStripMenuItem1 = new System.Windows.Forms.ToolStripMenuItem();
            this.выдавитьToolStripMenuItem1 = new System.Windows.Forms.ToolStripMenuItem();
            this.добавитьВСборкуToolStripMenuItem1 = new System.Windows.Forms.ToolStripMenuItem();
            this.крепежToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.перенаправитьToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.редкоеToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.добавитьСвойствоToolStripMenuItem1 = new System.Windows.Forms.ToolStripMenuItem();
            this.добавитьИсполнениеToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.комплектФайловToolStripMenuItem1 = new System.Windows.Forms.ToolStripMenuItem();
            this.комплектФайловСЧертежамиToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.stepToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.центрыОтверстийToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.сплайнToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.создатьСборкуToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.артикулToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.заменитьИмяКПToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.детальToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.сборкаToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.copyAttrToolStripMenuItem1 = new System.Windows.Forms.ToolStripMenuItem();
            this.addTRToolStripMenuItem1 = new System.Windows.Forms.ToolStripMenuItem();
            this.кластерToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.обновитьПозицииToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.регионToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.обновитьToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.спецToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.крепежToolStripMenuItem1 = new System.Windows.Forms.ToolStripMenuItem();
            this.получитьТекстToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.перевестиТекстToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.обновитьКрепежToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.обновитьСтилиToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.крепежToolStripMenuItem3 = new System.Windows.Forms.ToolStripMenuItem();
            this.скругленияToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.ручнаяГибкаToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.удалитьЭлементToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.переименоватьСборкуToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.габаритыToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.открытьПапкуToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.menuStrip1.SuspendLayout();
            this.SuspendLayout();
            // 
            // menuStrip1
            // 
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.детальToolStripMenuItem,
            this.сборкаToolStripMenuItem,
            this.чертежToolStripMenuItem,
            this.параметрыToolStripMenuItem,
            this.разместитьВсеДеталиToolStripMenuItem,
            this.редкоеToolStripMenuItem,
            this.шипыИПазыToolStripMenuItem,
            this.создатьФайлыToolStripMenuItem,
            this.перенаправитьToolStripMenuItem,
            this.крепежToolStripMenuItem});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Size = new System.Drawing.Size(675, 24);
            this.menuStrip1.TabIndex = 0;
            this.menuStrip1.Text = "menuStrip1";
            this.menuStrip1.ItemClicked += new System.Windows.Forms.ToolStripItemClickedEventHandler(this.menuStrip1_ItemClicked);
            // 
            // создатьФайлыToolStripMenuItem
            // 
            this.создатьФайлыToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.удалитьЭлементToolStripMenuItem,
            this.общиеРазмерыToolStripMenuItem,
            this.размерыДоГибовToolStripMenuItem,
            this.позицииToolStripMenuItem,
            this.крепежToolStripMenuItem2});
            this.создатьФайлыToolStripMenuItem.Name = "создатьФайлыToolStripMenuItem";
            this.создатьФайлыToolStripMenuItem.Size = new System.Drawing.Size(63, 20);
            this.создатьФайлыToolStripMenuItem.Text = "Удалить";
            this.создатьФайлыToolStripMenuItem.Click += new System.EventHandler(this.создатьФайлыToolStripMenuItem_Click);
            // 
            // общиеРазмерыToolStripMenuItem
            // 
            this.общиеРазмерыToolStripMenuItem.Name = "общиеРазмерыToolStripMenuItem";
            this.общиеРазмерыToolStripMenuItem.Size = new System.Drawing.Size(165, 22);
            this.общиеРазмерыToolStripMenuItem.Text = "Общие размеры";
            this.общиеРазмерыToolStripMenuItem.Click += new System.EventHandler(this.общиеРазмерыToolStripMenuItem_Click);
            // 
            // размерыДоГибовToolStripMenuItem
            // 
            this.размерыДоГибовToolStripMenuItem.Name = "размерыДоГибовToolStripMenuItem";
            this.размерыДоГибовToolStripMenuItem.Size = new System.Drawing.Size(165, 22);
            this.размерыДоГибовToolStripMenuItem.Text = "Размеры до гибов";
            this.размерыДоГибовToolStripMenuItem.Click += new System.EventHandler(this.размерыДоГибовToolStripMenuItem_Click);
            // 
            // позицииToolStripMenuItem
            // 
            this.позицииToolStripMenuItem.Name = "позицииToolStripMenuItem";
            this.позицииToolStripMenuItem.Size = new System.Drawing.Size(165, 22);
            this.позицииToolStripMenuItem.Text = "Позиции";
            this.позицииToolStripMenuItem.Click += new System.EventHandler(this.позицииToolStripMenuItem_Click);
            // 
            // крепежToolStripMenuItem2
            // 
            this.крепежToolStripMenuItem2.Name = "крепежToolStripMenuItem2";
            this.крепежToolStripMenuItem2.Size = new System.Drawing.Size(165, 22);
            this.крепежToolStripMenuItem2.Text = "Крепеж";
            this.крепежToolStripMenuItem2.Click += new System.EventHandler(this.крепежToolStripMenuItem2_Click);
            // 
            // чертежToolStripMenuItem
            // 
            this.чертежToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.текущийДокументToolStripMenuItem,
            this.переименоватьСборкуСЧертежамиToolStripMenuItem,
            this.регионToolStripMenuItem,
            this.обновитьПозицииToolStripMenuItem,
            this.addTRToolStripMenuItem1,
            this.copyAttrToolStripMenuItem1,
            this.получитьТекстToolStripMenuItem,
            this.перевестиТекстToolStripMenuItem,
            this.обновитьСтилиToolStripMenuItem,
            this.пересобратьСборкуToolStripMenuItem,
            this.переназватьToolStripMenuItem});
            this.чертежToolStripMenuItem.Name = "чертежToolStripMenuItem";
            this.чертежToolStripMenuItem.Size = new System.Drawing.Size(58, 20);
            this.чертежToolStripMenuItem.Text = "Чертеж";
            // 
            // текущийДокументToolStripMenuItem
            // 
            this.текущийДокументToolStripMenuItem.Name = "текущийДокументToolStripMenuItem";
            this.текущийДокументToolStripMenuItem.Size = new System.Drawing.Size(266, 22);
            this.текущийДокументToolStripMenuItem.Text = "Создать чертежи";
            this.текущийДокументToolStripMenuItem.Click += new System.EventHandler(this.текущийДокументToolStripMenuItem_Click);
            // 
            // переназватьToolStripMenuItem
            // 
            this.переназватьToolStripMenuItem.Name = "переназватьToolStripMenuItem";
            this.переназватьToolStripMenuItem.Size = new System.Drawing.Size(266, 22);
            this.переназватьToolStripMenuItem.Text = "Переименовать";
            this.переназватьToolStripMenuItem.Click += new System.EventHandler(this.переназватьToolStripMenuItem_Click);
            // 
            // пересобратьСборкуToolStripMenuItem
            // 
            this.пересобратьСборкуToolStripMenuItem.Name = "пересобратьСборкуToolStripMenuItem";
            this.пересобратьСборкуToolStripMenuItem.Size = new System.Drawing.Size(266, 22);
            this.пересобратьСборкуToolStripMenuItem.Text = "Пересобрать сборку";
            this.пересобратьСборкуToolStripMenuItem.Click += new System.EventHandler(this.пересобратьСборкуToolStripMenuItem_Click);
            // 
            // переименоватьСборкуСЧертежамиToolStripMenuItem
            // 
            this.переименоватьСборкуСЧертежамиToolStripMenuItem.Name = "переименоватьСборкуСЧертежамиToolStripMenuItem";
            this.переименоватьСборкуСЧертежамиToolStripMenuItem.Size = new System.Drawing.Size(266, 22);
            this.переименоватьСборкуСЧертежамиToolStripMenuItem.Text = "Переименовать сборку с чертежами";
            this.переименоватьСборкуСЧертежамиToolStripMenuItem.Click += new System.EventHandler(this.переименоватьСборкуСЧертежамиToolStripMenuItem_Click);
            // 
            // параметрыToolStripMenuItem
            // 
            this.параметрыToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.добавитьПараметрыToolStripMenuItem1,
            this.загрузитьXMLToolStripMenuItem,
            this.данныеДляМассиваToolStripMenuItem,
            this.массивToolStripMenuItem1,
            this.копироватьВXMLToolStripMenuItem,
            this.структураВXMLToolStripMenuItem,
            this.создатьДеталиToolStripMenuItem});
            this.параметрыToolStripMenuItem.Name = "параметрыToolStripMenuItem";
            this.параметрыToolStripMenuItem.Size = new System.Drawing.Size(38, 20);
            this.параметрыToolStripMenuItem.Text = "XML";
            // 
            // добавитьПараметрыToolStripMenuItem1
            // 
            this.добавитьПараметрыToolStripMenuItem1.Name = "добавитьПараметрыToolStripMenuItem1";
            this.добавитьПараметрыToolStripMenuItem1.Size = new System.Drawing.Size(183, 22);
            this.добавитьПараметрыToolStripMenuItem1.Text = "Добавить параметры";
            this.добавитьПараметрыToolStripMenuItem1.Click += new System.EventHandler(this.добавитьПараметрыToolStripMenuItem_Click);
            // 
            // данныеДляМассиваToolStripMenuItem
            // 
            this.данныеДляМассиваToolStripMenuItem.Name = "данныеДляМассиваToolStripMenuItem";
            this.данныеДляМассиваToolStripMenuItem.Size = new System.Drawing.Size(183, 22);
            this.данныеДляМассиваToolStripMenuItem.Text = "Данные для массива";
            this.данныеДляМассиваToolStripMenuItem.Click += new System.EventHandler(this.добавитьДанныеДляМассиваToolStripMenuItem_Click);
            // 
            // загрузитьXMLToolStripMenuItem
            // 
            this.загрузитьXMLToolStripMenuItem.Name = "загрузитьXMLToolStripMenuItem";
            this.загрузитьXMLToolStripMenuItem.Size = new System.Drawing.Size(183, 22);
            this.загрузитьXMLToolStripMenuItem.Text = "Загрузить XML";
            this.загрузитьXMLToolStripMenuItem.Click += new System.EventHandler(this.загрузитьXMLToolStripMenuItem_Click);
            // 
            // массивToolStripMenuItem1
            // 
            this.массивToolStripMenuItem1.Name = "массивToolStripMenuItem1";
            this.массивToolStripMenuItem1.Size = new System.Drawing.Size(183, 22);
            this.массивToolStripMenuItem1.Text = "Массив";
            this.массивToolStripMenuItem1.Click += new System.EventHandler(this.массивToolStripMenuItem_Click);
            // 
            // копироватьВXMLToolStripMenuItem
            // 
            this.копироватьВXMLToolStripMenuItem.Name = "копироватьВXMLToolStripMenuItem";
            this.копироватьВXMLToolStripMenuItem.Size = new System.Drawing.Size(183, 22);
            this.копироватьВXMLToolStripMenuItem.Text = "Копировать в XML";
            this.копироватьВXMLToolStripMenuItem.Click += new System.EventHandler(this.копироватьВXMLToolStripMenuItem_Click);
            // 
            // структураВXMLToolStripMenuItem
            // 
            this.структураВXMLToolStripMenuItem.Name = "структураВXMLToolStripMenuItem";
            this.структураВXMLToolStripMenuItem.Size = new System.Drawing.Size(183, 22);
            this.структураВXMLToolStripMenuItem.Text = "Структура в XML";
            this.структураВXMLToolStripMenuItem.Click += new System.EventHandler(this.структураВXMLToolStripMenuItem_Click);
            // 
            // создатьДеталиToolStripMenuItem
            // 
            this.создатьДеталиToolStripMenuItem.Name = "создатьДеталиToolStripMenuItem";
            this.создатьДеталиToolStripMenuItem.Size = new System.Drawing.Size(183, 22);
            this.создатьДеталиToolStripMenuItem.Text = "Создать детали";
            this.создатьДеталиToolStripMenuItem.Click += new System.EventHandler(this.создатьДеталиToolStripMenuItem_Click);
            // 
            // разместитьВсеДеталиToolStripMenuItem
            // 
            this.разместитьВсеДеталиToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.разместитьВсеДеталиToolStripMenuItem1,
            this.обновитьContentCenterToolStripMenuItem,
            this.сортировкаToolStripMenuItem,
            this.тестToolStripMenuItem});
            this.разместитьВсеДеталиToolStripMenuItem.Name = "разместитьВсеДеталиToolStripMenuItem";
            this.разместитьВсеДеталиToolStripMenuItem.Size = new System.Drawing.Size(79, 20);
            this.разместитьВсеДеталиToolStripMenuItem.Text = "Библиотека";
            this.разместитьВсеДеталиToolStripMenuItem.Click += new System.EventHandler(this.создатьДеревоToolStripMenuItem_Click);
            // 
            // разместитьВсеДеталиToolStripMenuItem1
            // 
            this.разместитьВсеДеталиToolStripMenuItem1.Name = "разместитьВсеДеталиToolStripMenuItem1";
            this.разместитьВсеДеталиToolStripMenuItem1.Size = new System.Drawing.Size(199, 22);
            this.разместитьВсеДеталиToolStripMenuItem1.Text = "Разместить все детали";
            this.разместитьВсеДеталиToolStripMenuItem1.Click += new System.EventHandler(this.разместитьВсеДеталиToolStripMenuItem1_Click);
            // 
            // обновитьContentCenterToolStripMenuItem
            // 
            this.обновитьContentCenterToolStripMenuItem.Name = "обновитьContentCenterToolStripMenuItem";
            this.обновитьContentCenterToolStripMenuItem.Size = new System.Drawing.Size(199, 22);
            this.обновитьContentCenterToolStripMenuItem.Text = "Обновить ContentCenter";
            this.обновитьContentCenterToolStripMenuItem.Click += new System.EventHandler(this.обновитьContentCenterToolStripMenuItem_Click);
            // 
            // сортировкаToolStripMenuItem
            // 
            this.сортировкаToolStripMenuItem.Name = "сортировкаToolStripMenuItem";
            this.сортировкаToolStripMenuItem.Size = new System.Drawing.Size(199, 22);
            this.сортировкаToolStripMenuItem.Text = "Сортировка";
            this.сортировкаToolStripMenuItem.Click += new System.EventHandler(this.сортировкаToolStripMenuItem_Click);
            // 
            // тестToolStripMenuItem
            // 
            this.тестToolStripMenuItem.Name = "тестToolStripMenuItem";
            this.тестToolStripMenuItem.Size = new System.Drawing.Size(199, 22);
            this.тестToolStripMenuItem.Text = "Тест";
            this.тестToolStripMenuItem.Click += new System.EventHandler(this.тестToolStripMenuItem_Click);
            // 
            // шипыИПазыToolStripMenuItem
            // 
            this.шипыИПазыToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.шипToolStripMenuItem1,
            this.выдавитьToolStripMenuItem1,
            this.добавитьВСборкуToolStripMenuItem1,
            this.пазToolStripMenuItem1});
            this.шипыИПазыToolStripMenuItem.Name = "шипыИПазыToolStripMenuItem";
            this.шипыИПазыToolStripMenuItem.Size = new System.Drawing.Size(86, 20);
            this.шипыИПазыToolStripMenuItem.Text = "Шипы и пазы";
            // 
            // шипToolStripMenuItem1
            // 
            this.шипToolStripMenuItem1.Name = "шипToolStripMenuItem1";
            this.шипToolStripMenuItem1.Size = new System.Drawing.Size(171, 22);
            this.шипToolStripMenuItem1.Text = "Шип";
            this.шипToolStripMenuItem1.Click += new System.EventHandler(this.шипToolStripMenuItem1_Click);
            // 
            // пазToolStripMenuItem1
            // 
            this.пазToolStripMenuItem1.Name = "пазToolStripMenuItem1";
            this.пазToolStripMenuItem1.Size = new System.Drawing.Size(171, 22);
            this.пазToolStripMenuItem1.Text = "Паз";
            this.пазToolStripMenuItem1.Click += new System.EventHandler(this.пазToolStripMenuItem1_Click);
            // 
            // выдавитьToolStripMenuItem1
            // 
            this.выдавитьToolStripMenuItem1.Name = "выдавитьToolStripMenuItem1";
            this.выдавитьToolStripMenuItem1.Size = new System.Drawing.Size(171, 22);
            this.выдавитьToolStripMenuItem1.Text = "Выдавить";
            this.выдавитьToolStripMenuItem1.Click += new System.EventHandler(this.выдавитьToolStripMenuItem1_Click);
            // 
            // добавитьВСборкуToolStripMenuItem1
            // 
            this.добавитьВСборкуToolStripMenuItem1.Name = "добавитьВСборкуToolStripMenuItem1";
            this.добавитьВСборкуToolStripMenuItem1.Size = new System.Drawing.Size(171, 22);
            this.добавитьВСборкуToolStripMenuItem1.Text = "Добавить в сборку";
            this.добавитьВСборкуToolStripMenuItem1.Click += new System.EventHandler(this.добавитьВСборкуToolStripMenuItem1_Click);
            // 
            // крепежToolStripMenuItem
            // 
            this.крепежToolStripMenuItem.Name = "крепежToolStripMenuItem";
            this.крепежToolStripMenuItem.Size = new System.Drawing.Size(58, 20);
            this.крепежToolStripMenuItem.Text = "Крепеж";
            this.крепежToolStripMenuItem.Click += new System.EventHandler(this.крепежToolStripMenuItem_Click);
            // 
            // перенаправитьToolStripMenuItem
            // 
            this.перенаправитьToolStripMenuItem.Name = "перенаправитьToolStripMenuItem";
            this.перенаправитьToolStripMenuItem.Size = new System.Drawing.Size(98, 20);
            this.перенаправитьToolStripMenuItem.Text = "Перенаправить";
            this.перенаправитьToolStripMenuItem.Click += new System.EventHandler(this.перенаправитьToolStripMenuItem_Click);
            // 
            // редкоеToolStripMenuItem
            // 
            this.редкоеToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.stepToolStripMenuItem,
            this.комплектФайловToolStripMenuItem1,
            this.комплектФайловСЧертежамиToolStripMenuItem,
            this.обновитьToolStripMenuItem,
            this.добавитьСвойствоToolStripMenuItem1,
            this.добавитьИсполнениеToolStripMenuItem,
            this.сплайнToolStripMenuItem,
            this.артикулToolStripMenuItem,
            this.центрыОтверстийToolStripMenuItem,
            this.создатьСборкуToolStripMenuItem,
            this.заменитьИмяКПToolStripMenuItem,
            this.габаритыToolStripMenuItem,
            this.открытьПапкуToolStripMenuItem});
            this.редкоеToolStripMenuItem.Name = "редкоеToolStripMenuItem";
            this.редкоеToolStripMenuItem.Size = new System.Drawing.Size(54, 20);
            this.редкоеToolStripMenuItem.Text = "Общее";
            // 
            // добавитьСвойствоToolStripMenuItem1
            // 
            this.добавитьСвойствоToolStripMenuItem1.Name = "добавитьСвойствоToolStripMenuItem1";
            this.добавитьСвойствоToolStripMenuItem1.Size = new System.Drawing.Size(231, 22);
            this.добавитьСвойствоToolStripMenuItem1.Text = "Добавить свойство";
            this.добавитьСвойствоToolStripMenuItem1.Click += new System.EventHandler(this.добавитьСвойствоToolStripMenuItem1_Click);
            // 
            // добавитьИсполнениеToolStripMenuItem
            // 
            this.добавитьИсполнениеToolStripMenuItem.Name = "добавитьИсполнениеToolStripMenuItem";
            this.добавитьИсполнениеToolStripMenuItem.Size = new System.Drawing.Size(231, 22);
            this.добавитьИсполнениеToolStripMenuItem.Text = "Добавить исполнение";
            this.добавитьИсполнениеToolStripMenuItem.Click += new System.EventHandler(this.добавитьИсполнениеToolStripMenuItem_Click);
            // 
            // комплектФайловToolStripMenuItem1
            // 
            this.комплектФайловToolStripMenuItem1.Name = "комплектФайловToolStripMenuItem1";
            this.комплектФайловToolStripMenuItem1.Size = new System.Drawing.Size(231, 22);
            this.комплектФайловToolStripMenuItem1.Text = "Комплект файлов";
            this.комплектФайловToolStripMenuItem1.Click += new System.EventHandler(this.комплектФайловToolStripMenuItem1_Click);
            // 
            // комплектФайловСЧертежамиToolStripMenuItem
            // 
            this.комплектФайловСЧертежамиToolStripMenuItem.Name = "комплектФайловСЧертежамиToolStripMenuItem";
            this.комплектФайловСЧертежамиToolStripMenuItem.Size = new System.Drawing.Size(231, 22);
            this.комплектФайловСЧертежамиToolStripMenuItem.Text = "Комплект файлов с чертежами";
            this.комплектФайловСЧертежамиToolStripMenuItem.Click += new System.EventHandler(this.комплектФайловСЧертежамиToolStripMenuItem_Click);
            // 
            // stepToolStripMenuItem
            // 
            this.stepToolStripMenuItem.Name = "stepToolStripMenuItem";
            this.stepToolStripMenuItem.Size = new System.Drawing.Size(231, 22);
            this.stepToolStripMenuItem.Text = "3DPDF";
            this.stepToolStripMenuItem.Click += new System.EventHandler(this.stepToolStripMenuItem_Click);
            // 
            // центрыОтверстийToolStripMenuItem
            // 
            this.центрыОтверстийToolStripMenuItem.Name = "центрыОтверстийToolStripMenuItem";
            this.центрыОтверстийToolStripMenuItem.Size = new System.Drawing.Size(231, 22);
            this.центрыОтверстийToolStripMenuItem.Text = "Центры отверстий";
            this.центрыОтверстийToolStripMenuItem.Click += new System.EventHandler(this.центрыОтверстийToolStripMenuItem_Click);
            // 
            // сплайнToolStripMenuItem
            // 
            this.сплайнToolStripMenuItem.Name = "сплайнToolStripMenuItem";
            this.сплайнToolStripMenuItem.Size = new System.Drawing.Size(231, 22);
            this.сплайнToolStripMenuItem.Text = "Сплайн";
            this.сплайнToolStripMenuItem.Click += new System.EventHandler(this.сплайнToolStripMenuItem_Click);
            // 
            // создатьСборкуToolStripMenuItem
            // 
            this.создатьСборкуToolStripMenuItem.Name = "создатьСборкуToolStripMenuItem";
            this.создатьСборкуToolStripMenuItem.Size = new System.Drawing.Size(231, 22);
            this.создатьСборкуToolStripMenuItem.Text = "Создать сборку";
            this.создатьСборкуToolStripMenuItem.Click += new System.EventHandler(this.создатьСборкуToolStripMenuItem_Click);
            // 
            // артикулToolStripMenuItem
            // 
            this.артикулToolStripMenuItem.Name = "артикулToolStripMenuItem";
            this.артикулToolStripMenuItem.Size = new System.Drawing.Size(231, 22);
            this.артикулToolStripMenuItem.Text = "Артикул";
            this.артикулToolStripMenuItem.Click += new System.EventHandler(this.артикулToolStripMenuItem_Click);
            // 
            // заменитьИмяКПToolStripMenuItem
            // 
            this.заменитьИмяКПToolStripMenuItem.Name = "заменитьИмяКПToolStripMenuItem";
            this.заменитьИмяКПToolStripMenuItem.Size = new System.Drawing.Size(231, 22);
            this.заменитьИмяКПToolStripMenuItem.Text = "Заменить имя КП";
            this.заменитьИмяКПToolStripMenuItem.Click += new System.EventHandler(this.заменитьИмяКПToolStripMenuItem_Click);
            // 
            // детальToolStripMenuItem
            // 
            this.детальToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.кластерToolStripMenuItem,
            this.скругленияToolStripMenuItem,
            this.ручнаяГибкаToolStripMenuItem});
            this.детальToolStripMenuItem.Name = "детальToolStripMenuItem";
            this.детальToolStripMenuItem.Size = new System.Drawing.Size(57, 20);
            this.детальToolStripMenuItem.Text = "Деталь";
            // 
            // сборкаToolStripMenuItem
            // 
            this.сборкаToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.спецToolStripMenuItem,
            this.крепежToolStripMenuItem1,
            this.обновитьКрепежToolStripMenuItem,
            this.крепежToolStripMenuItem3,
            this.переименоватьСборкуToolStripMenuItem});
            this.сборкаToolStripMenuItem.Name = "сборкаToolStripMenuItem";
            this.сборкаToolStripMenuItem.Size = new System.Drawing.Size(56, 20);
            this.сборкаToolStripMenuItem.Text = "Сборка";
            // 
            // copyAttrToolStripMenuItem1
            // 
            this.copyAttrToolStripMenuItem1.Name = "copyAttrToolStripMenuItem1";
            this.copyAttrToolStripMenuItem1.Size = new System.Drawing.Size(266, 22);
            this.copyAttrToolStripMenuItem1.Text = "Копировать технические требования";
            // 
            // addTRToolStripMenuItem1
            // 
            this.addTRToolStripMenuItem1.Name = "addTRToolStripMenuItem1";
            this.addTRToolStripMenuItem1.Size = new System.Drawing.Size(266, 22);
            this.addTRToolStripMenuItem1.Text = "Добавить технические требования";
            // 
            // кластерToolStripMenuItem
            // 
            this.кластерToolStripMenuItem.Name = "кластерToolStripMenuItem";
            this.кластерToolStripMenuItem.Size = new System.Drawing.Size(152, 22);
            this.кластерToolStripMenuItem.Text = "Кластер";
            // 
            // обновитьПозицииToolStripMenuItem
            // 
            this.обновитьПозицииToolStripMenuItem.Name = "обновитьПозицииToolStripMenuItem";
            this.обновитьПозицииToolStripMenuItem.Size = new System.Drawing.Size(266, 22);
            this.обновитьПозицииToolStripMenuItem.Text = "Обновить позиции";
            // 
            // регионToolStripMenuItem
            // 
            this.регионToolStripMenuItem.Name = "регионToolStripMenuItem";
            this.регионToolStripMenuItem.Size = new System.Drawing.Size(266, 22);
            this.регионToolStripMenuItem.Text = "Выровнять позиции";
            // 
            // обновитьToolStripMenuItem
            // 
            this.обновитьToolStripMenuItem.Name = "обновитьToolStripMenuItem";
            this.обновитьToolStripMenuItem.Size = new System.Drawing.Size(231, 22);
            this.обновитьToolStripMenuItem.Text = "Обновить";
            // 
            // спецToolStripMenuItem
            // 
            this.спецToolStripMenuItem.Name = "спецToolStripMenuItem";
            this.спецToolStripMenuItem.Size = new System.Drawing.Size(248, 22);
            this.спецToolStripMenuItem.Text = "Данные для автокрепежа";
            // 
            // крепежToolStripMenuItem1
            // 
            this.крепежToolStripMenuItem1.Name = "крепежToolStripMenuItem1";
            this.крепежToolStripMenuItem1.Size = new System.Drawing.Size(248, 22);
            this.крепежToolStripMenuItem1.Text = "Автокрепеж";
            // 
            // получитьТекстToolStripMenuItem
            // 
            this.получитьТекстToolStripMenuItem.Name = "получитьТекстToolStripMenuItem";
            this.получитьТекстToolStripMenuItem.Size = new System.Drawing.Size(266, 22);
            this.получитьТекстToolStripMenuItem.Text = "Получить текст";
            // 
            // перевестиТекстToolStripMenuItem
            // 
            this.перевестиТекстToolStripMenuItem.Name = "перевестиТекстToolStripMenuItem";
            this.перевестиТекстToolStripMenuItem.Size = new System.Drawing.Size(266, 22);
            this.перевестиТекстToolStripMenuItem.Text = "Перевести текст";
            // 
            // обновитьКрепежToolStripMenuItem
            // 
            this.обновитьКрепежToolStripMenuItem.Name = "обновитьКрепежToolStripMenuItem";
            this.обновитьКрепежToolStripMenuItem.Size = new System.Drawing.Size(248, 22);
            this.обновитьКрепежToolStripMenuItem.Text = "Обновить крепеж (Все открытые)";
            // 
            // обновитьСтилиToolStripMenuItem
            // 
            this.обновитьСтилиToolStripMenuItem.Name = "обновитьСтилиToolStripMenuItem";
            this.обновитьСтилиToolStripMenuItem.Size = new System.Drawing.Size(266, 22);
            this.обновитьСтилиToolStripMenuItem.Text = "Обновить стили";
            // 
            // крепежToolStripMenuItem3
            // 
            this.крепежToolStripMenuItem3.Name = "крепежToolStripMenuItem3";
            this.крепежToolStripMenuItem3.Size = new System.Drawing.Size(248, 22);
            this.крепежToolStripMenuItem3.Text = "Крепеж (Все открытые)";
            // 
            // скругленияToolStripMenuItem
            // 
            this.скругленияToolStripMenuItem.Name = "скругленияToolStripMenuItem";
            this.скругленияToolStripMenuItem.Size = new System.Drawing.Size(152, 22);
            this.скругленияToolStripMenuItem.Text = "Скругления";
            // 
            // ручнаяГибкаToolStripMenuItem
            // 
            this.ручнаяГибкаToolStripMenuItem.Name = "ручнаяГибкаToolStripMenuItem";
            this.ручнаяГибкаToolStripMenuItem.Size = new System.Drawing.Size(152, 22);
            this.ручнаяГибкаToolStripMenuItem.Text = "Ручная гибка";
            // 
            // удалитьЭлементToolStripMenuItem
            // 
            this.удалитьЭлементToolStripMenuItem.Name = "удалитьЭлементToolStripMenuItem";
            this.удалитьЭлементToolStripMenuItem.Size = new System.Drawing.Size(165, 22);
            this.удалитьЭлементToolStripMenuItem.Text = "Удалить элемент";
            // 
            // переименоватьСборкуToolStripMenuItem
            // 
            this.переименоватьСборкуToolStripMenuItem.Name = "переименоватьСборкуToolStripMenuItem";
            this.переименоватьСборкуToolStripMenuItem.Size = new System.Drawing.Size(248, 22);
            this.переименоватьСборкуToolStripMenuItem.Text = "Переименовать сборку";
            // 
            // габаритыToolStripMenuItem
            // 
            this.габаритыToolStripMenuItem.Name = "габаритыToolStripMenuItem";
            this.габаритыToolStripMenuItem.Size = new System.Drawing.Size(231, 22);
            this.габаритыToolStripMenuItem.Text = "Габариты";
            // 
            // открытьПапкуToolStripMenuItem
            // 
            this.открытьПапкуToolStripMenuItem.Name = "открытьПапкуToolStripMenuItem";
            this.открытьПапкуToolStripMenuItem.Size = new System.Drawing.Size(231, 22);
            this.открытьПапкуToolStripMenuItem.Text = "Открыть папку";
            // 
            // CreateComponent
            // 
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Inherit;
            this.ClientSize = new System.Drawing.Size(675, 24);
            this.Controls.Add(this.menuStrip1);
            this.Location = new System.Drawing.Point(300, 130);
            this.MainMenuStrip = this.menuStrip1;
            this.Name = "CreateComponent";
            this.StartPosition = System.Windows.Forms.FormStartPosition.Manual;
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        public CreateComponent(Inventor.Document doc)
        {
            m_Parts = new Parts(Macros.StandardAddInServer.m_inventorApplication.ActiveDocument);
            desc = new XMLDoc(@"C:\ProgramData\Autodesk\Inventor Addins\Description.xml", "Description");
            InitializeComponent();
            InterfaceDll.Data.EventHandler = new Data.MyEvent(func);
            this.ShowDialog();
        }

        private void перенаправитьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Document doc = Macros.StandardAddInServer.m_inventorApplication.ActiveDocument;
            string newName = InvDoc.u.OFD(System.IO.Path.GetDirectoryName(doc.FullDocumentName));
            foreach (Document docum in Macros.StandardAddInServer.m_inventorApplication.Documents.VisibleDocuments)
            {
                if (docum.DocumentType == DocumentTypeEnum.kPartDocumentObject)
                {
                    Parts.replaceFullFileName((PartDocument)docum, newName);
                }
//                 else if (docum.DocumentType == DocumentTypeEnum.kAssemblyDocumentObject)
//                 {
//                     Parts.replaceFullFileName((AssemblyDocument)docum, newName);
//                 }
                else if (docum.DocumentType == DocumentTypeEnum.kDrawingDocumentObject)
                {
                    Parts.replaceFullFileName((DrawingDocument)docum, newName);
                }
            }
        }

        private void создатьФайлыToolStripMenuItem_Click(object sender, EventArgs e)
        {
//             this.Hide();
//             m_Parts.create();
//             this.Show();
        }

        private void создатьДеревоToolStripMenuItem_Click(object sender, EventArgs e)
        {
        }

        public void createDescription(DataGridView dgv, string name, string path)
        {
            XMLDoc xmldoc = new XMLDoc(path + "\\" + name, "Head");
            XElement based = new XElement("Parts");
            List<string> add = new List<string>();
            List<KeyValuePair<string,string>> replace = new List<KeyValuePair<string,string>>();
            List<string> remove = new List<string>();
            DataGridViewColumn colName = null, colBase = null, colReplace = null, colLib = null, colDesc = null, colPN = null, colCount = null;
            foreach (DataGridViewColumn item in dgv.Columns)
            {
                if (item.HeaderText == "Название файла") colName = item;
                else if (item.HeaderText == "название файла для наследования") colBase = item;
                else if (item.HeaderText == "Файлы для замены") colReplace = item;
                else if (item.HeaderText == "Замена") colLib = item;
                else if (item.HeaderText == "наименование") colDesc = item;
                else if (item.HeaderText == "децимальный номер") colPN = item;
                else if (item.HeaderText == "Кол-во") colCount = item;
            }
            xmldoc.Doc.Root.Add(based);
            for (int i = 0; i < dgv.Rows.Count - 1; i++)
            {
                DataGridViewRow dgvr = dgv.Rows[i];
                if (dgvr.Cells[0] != null || dgvr.Cells[0].Value.ToString() != "")
                {
                    string val = dgvr.Cells[0].Value.ToString();
                    if (val.EndsWith(".ipt"))
                    {
                        XElement elem = new XElement("Part");
                        based.Add(elem);
                        if (dgv[0, i].Value != null) elem.Value = dgv[0, i].Value.ToString();
                        if (dgv[colBase.Index, i].Value != null) XMLDoc.addXAttribute(elem, "base", dgv[colBase.Index, i].Value.ToString());
                        //elem.SetAttributeValue("base", dgv[colBase.Index, i].Value.ToString());
                        if (dgv[colPN.Index, i].Value != null) elem.SetAttributeValue("DecNumber", dgv[colPN.Index, i].Value.ToString());
                        if (this.txtbox1.Text != "") XMLDoc.addXAttribute(elem, "Type", this.txtbox1.Text);
                            //elem.SetAttributeValue("Type", this.txtbox1.Text);
                        if (dgv[colDesc.Index, i].Value != null) elem.SetAttributeValue("D",dgv[colDesc.Index, i].Value.ToString());
                    }
                    else
                    {
                        XElement elem = new XElement("Asm");
                        based.Add(elem);
                        string asmName = "";
                        if (dgv[0, i].Value != null) { elem.Value = dgv[0, i].Value.ToString(); asmName = dgv[0, i].Value.ToString(); };
                        if (dgv[colBase.Index, i].Value != null) elem.SetAttributeValue("base", dgv[colBase.Index, i].Value.ToString());
                        if (dgv[colPN.Index, i].Value != null) elem.SetAttributeValue("DecNumber", dgv[colPN.Index, i].Value.ToString());
                        if (this.txtbox1.Text != "") XMLDoc.addXAttribute(elem, "Type", this.txtbox1.Text);
                        //elem.SetAttributeValue("Type", this.txtbox1.Text);
                        if (dgv[colDesc.Index, i].Value != null) elem.SetAttributeValue("D", dgv[colDesc.Index, i].Value.ToString());
//                         if (dgv.Rows[i+1].Cells[0].Value == null || dgv.Rows[i+1].Cells[0].Value == "")
//                         {
                        int k = 0;
                            do
                            {   
                                string lib, rep;
                                if (dgv[colReplace.Index, i+k].Value == null) rep = "";
                                else rep = dgv[colReplace.Index, i+k].Value.ToString();

                                if (dgv[colLib.Index, i+k].Value == null) lib = "";
                                else lib = dgv[colLib.Index, i+k].Value.ToString();

                                if (rep != "" && lib != "") replace.Add(new KeyValuePair<string, string>(rep, lib));
                                else if (rep == "" && lib != "") 
                                {
                                    add.Add(lib);
                                    if (dgv[colCount.Index, i + k].Value == null)
                                        add.Add("");
                                    else
                                    add.Add(dgv[colCount.Index, i + k].Value.ToString()); 
                                }
                                else if (lib == "" && rep != "") remove.Add(rep);
                                k++;
                            }
                            while (i + k < dgv.RowCount - 1 && (dgv.Rows[i + k].Cells[0].Value == null || dgv.Rows[i + k].Cells[0].Value.ToString() == "" 
                                || asmName == dgv[0, i + k].Value.ToString())) ;
                        //}
                            
                        if (replace.Count != 0) elem.Add(new XAttribute("replace", replaceUnion(replace)));
                        if (remove.Count != 0) elem.Add(new XAttribute("remove", union(remove)));
                        if (add.Count != 0) elem.Add(new XAttribute("files", union(add)));
                        replace.Clear(); remove.Clear(); add.Clear(); i += k - 1; k = 0;
                    }
                }
            }
          xmldoc.save();
        }

        public string replaceUnion(List<KeyValuePair<string,string>> lst)
        {
           string ret = "";
           foreach (var item in lst)
	        {
                ret += item.Key + "$" + item.Value + "$"; 
	        }
           ret = ret.Remove(ret.Length - 1, 1);
           return ret;
        }

        private string union(List<string> lst)
        {
            string ret = "";
            foreach (var item in lst)
            {
                ret += item + "$"; 
            }
            ret = ret.Remove(ret.Length - 1, 1);
            return ret;
        }

        private void создатьОписаниеToolStripMenuItem_Click(object sender, EventArgs e)
        {
            doc = invApp.ActiveDocument;
            m_Parts = m_Parts ?? new Parts(doc);
            if (m_Parts.libraryPath.Count == 0) m_Parts.libraryPath = null;
            this.WindowState = FormWindowState.Maximized;
            System.Drawing.Rectangle bnds = Screen.PrimaryScreen.WorkingArea;

            MyLabel lbl = new MyLabel();
            MyTextBox txtBOX = new MyTextBox();
            int offset = 10;
            System.Drawing.Point pt = new System.Drawing.Point(0, 40);

            MyCheckBox chk = new MyCheckBox();
            checkbox = chk.addCheckBox("Сборка", new System.Drawing.Point(offset, pt.Y), 100, 20);
            checkbox.CheckedChanged += new EventHandler(chkdChanged);
            this.Controls.Add(checkbox);

            Label lbl1 = lbl.addLabel("Модель", MyLabel.position(checkbox, offset, 2), 50, 20);
            this.Controls.Add(lbl1);
            txtbox1 = txtBOX.addTextBox("", MyLabel.position(lbl1, offset, -2), 200, 20);
            this.Controls.Add(txtbox1);
            autoComplete(txtbox1);

            Label lbl2 = lbl.addLabel("Вид", MyLabel.position(txtbox1, offset, 2), 20, 20);
            this.Controls.Add(lbl2);

            MyComboBox cmbBox = new MyComboBox();
            cBox = cmbBox.addComboBox("Filter", MyLabel.position(lbl2, offset, -2), 200, 20);
            this.Controls.Add(cBox);
            autoComplete(cBox);

            m_Parts.doc = doc;
            Property pr = m_Parts.getProp("Type");
            if (pr != null) txtbox1.Text = pr.Value.ToString();

            lbl1 = lbl.addLabel("Название файла описания", MyLabel.position(cBox, offset, 2), 100, 20);
            this.Controls.Add(lbl1);
            txt2 = txtBOX.addTextBox("Сборки.xml", MyLabel.position(lbl1, offset, -2), 200, 20);
            if (checkbox.Checked == false) txt2.Text = "Детали.xml";
            this.Controls.Add(txt2);
            MyButton btn = new MyButton();

            System.Windows.Forms.Button btn1 = btn.addButton("Создать", MyLabel.position(txt2, offset, 0), 200, 20);
            btn1.Click += new EventHandler(btnClick);
            this.Controls.Add(btn1);

            System.Windows.Forms.Button btn2 = btn.addButton("Загрузить данные", MyLabel.position(btn1, offset, 0), 200, 20);
            btn2.Click += new EventHandler(LoadClick);
            this.Controls.Add(btn2);

            pt.Y += 30;
            mydgv = new MyDGV();
            XElement node = null;
            Dictionary<string,string> dic = new Dictionary<string,string>() {{"name","Название файла"},{"description","наименование"},{"decNumber", "децимальный номер"},{"base","название файла для наследования"},{"files","Файлы для замены"},{"replace", "Замена"},{"count", "Кол-во"}};
            float[] weigth = {0.2f,0.2f,0.1f,0.2f,0.1f,0.2f,0.1f};
            dgv = mydgv.addDGV(pt, bnds.Width, bnds.Height - 150, node, dic, weigth);
            this.Controls.Add(dgv);
            dgv.DataError += new DataGridViewDataErrorEventHandler(dataError);
            dgv.CellEndEdit += new DataGridViewCellEventHandler(cellChanged);  
            dgv.EditingControlShowing += new DataGridViewEditingControlShowingEventHandler(editingControlShowing);
        }

        private void chkdChanged(object sender, EventArgs e)
        {
            if (((CheckBox)sender).Checked == true) txt2.Text = "Сборки.xml";
            else txt2.Text = "Детали.xml";
        }

        private void cellChanged(object sender, DataGridViewCellEventArgs e)
        {
            DataGridView dgv = (DataGridView)sender;
            if (e.ColumnIndex == 1)
            {
             desc = desc ?? new XMLDoc(@"C:\ProgramData\Autodesk\Inventor Addins\Description.xml", "Description");
             string n = dgv[e.ColumnIndex, e.RowIndex].Value.ToString();
             if (n.IndexOf(this.txtbox1.Text) != -1)
             {
                 n = n.Replace(" " + this.txtbox1.Text, "");
             }
             XElement el = desc.Doc.Root.Descendants().FirstOrDefault(eel => eel.Attribute("name").Value.ToLower() == n.ToLower());
             if (el.HasElements) { checkbox.Checked = true; txt2.Text = "Сборки.xml"; }
             if (el != null)
                 dgv[e.ColumnIndex + 1, e.RowIndex].Value = el.Attribute("decNumber").Value; 
             if ((dgv[0, e.RowIndex].Value == null || dgv[0, e.RowIndex].Value.ToString() == "") && this.txtbox1.Text != "")
             {
                 if (checkbox.Checked)
                     dgv[0, e.RowIndex].Value = n + " (" + txtbox1.Text + "." + dgv[e.ColumnIndex + 1, e.RowIndex].Value + ").iam";
                 else
                 { 
                     dgv[3, e.RowIndex].Value = Macros.StandardAddInServer.m_inventorApplication.ActiveDocument.name() + ".ipt";
                     dgv[0, e.RowIndex].Value = n + " (" + txtbox1.Text + "." + dgv[e.ColumnIndex + 1, e.RowIndex].Value + ").ipt"; 
                 }
             }
            }
        }

        private void dataError(object sender, DataGridViewDataErrorEventArgs anError)
        {
            string err = "";
            err = anError.Context.ToString();
        }

        private void autoComplete(System.Windows.Forms.TextBox tb)
        {
            tb.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
            tb.AutoCompleteSource = AutoCompleteSource.CustomSource;
            AutoCompleteStringCollection col = new AutoCompleteStringCollection();
            XMLDoc xmldoc = new XMLDoc(@"C:\ProgramData\Autodesk\Inventor Addins\Description.xml", "Description");
            addItems(xmldoc.Doc, col, "Type", "name");
            tb.AutoCompleteCustomSource = col;
        }

        private void autoComplete(System.Windows.Forms.ComboBox cb)
        {
            XMLDoc xmldoc = new XMLDoc(@"C:\ProgramData\Autodesk\Inventor Addins\Description.xml", "Description");
            HashSet<string> hs = new HashSet<string>();
            foreach (var item in xmldoc.El.Descendants("Value"))
            {
               if (item.Attribute("type") != null) hs.Add(item.Attribute("type").Value);
            }
            cb.Items.AddRange(hs.ToArray());
        }

        public void addItems(XDocument doc, AutoCompleteStringCollection col,string name ,string attName)
        {
            foreach (var item in doc.Root.Descendants(name))
	        {
                if (item.Attribute(attName).Value != null)
                {
                    col.Add(item.Attribute(attName).Value);
                }
	        }
        }

        private void editingControlShowing(object sender, DataGridViewEditingControlShowingEventArgs e)
        {
            try
            {
            DataGridView dgv = (DataGridView)sender;
            AssemblyDocument asmDoc; 
            List<string> exept = new List<string>() { "OldVersions", ".ipj", ".dwg", ".xml", ".idw", ".lck", ".pdf", ".dxf", ".png" };
            string filter = "*.*";
            DataGridViewColumn colName = null, colBase = null, colReplace = null, colLib = null, colDesc = null, colPN = null;
            foreach (DataGridViewColumn item in dgv.Columns)
            {
                if (item.HeaderText == "Название файла") colName = item;
                else if (item.HeaderText == "название файла для наследования") colBase = item;
                else if (item.HeaderText == "Файлы для замены") colReplace = item;
                else if (item.HeaderText == "Замена") colLib = item;
                else if (item.HeaderText == "наименование") colDesc = item;
                else if (item.HeaderText == "децимальный номер") colPN = item;
            }
            System.Windows.Forms.TextBox autoText;
            autoText = e.Control as System.Windows.Forms.TextBox;
            autoText.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
            autoText.AutoCompleteSource = AutoCompleteSource.CustomSource;
            AutoCompleteStringCollection DataCollection = new AutoCompleteStringCollection();
            System.IO.SearchOption opt = System.IO.SearchOption.TopDirectoryOnly;
            DataGridViewCell curCell = dgv.CurrentCell;
            string titleText = dgv.Columns[curCell.ColumnIndex].HeaderText;
            desc = desc ?? new XMLDoc(@"C:\ProgramData\Autodesk\Inventor Addins\Description.xml", "Description");
            switch (titleText)
            {   
               case "Название файла":
                    if (autoText != null)
                    {
                        addItems(DataCollection, filter, exept, opt);
                        autoText.AutoCompleteCustomSource = DataCollection;
                    }
                    break;
                case "наименование":
                    if (autoText != null)
                    {
                        addItems(desc.Doc,"name","Value",DataCollection);
                        autoText.AutoCompleteCustomSource = DataCollection;
                    }
                    break;
                case "децимальный номер":
                    break;
                case "название файла для наследования":
                    if (autoText != null)
                    {
                        addItems(DataCollection, filter, exept, m_Parts.libraryPath);
                        autoText.AutoCompleteCustomSource = DataCollection;
                    }
                    break;
                case "Файлы для замены":
                    string asmname;
                    if (dgv[colName.Index, curCell.RowIndex].Value == null) asmname = "";
                    else asmname = dgv[colName.Index, curCell.RowIndex].Value.ToString();
                    int i = 0;
                        while (asmname == "")
                        {
                            if (dgv[colName.Index, curCell.RowIndex - i].Value == null) asmname = "";
                            else asmname = dgv[colName.Index, curCell.RowIndex - i].Value.ToString();
                            i++;
                        }
                    if (autoText != null && (asmname.IndexOf(".iam") != -1))
                    {
                        asmDoc = (AssemblyDocument)invApp.Documents.Open(doc.path() + "\\" + asmname,false);
                        addItems(asmDoc, DataCollection);
                        autoText.AutoCompleteCustomSource = DataCollection;
                    }
                    break;
                case "Замена":
                    string asmname1;
                    if (dgv[colName.Index, curCell.RowIndex].Value == null) asmname1 = "";
                    else asmname1 = dgv[colName.Index, curCell.RowIndex].Value.ToString();
                    int j = 0;
                        while (asmname1 == "")
                        {
                            if (dgv[colName.Index, curCell.RowIndex - j].Value == null) asmname1 = "";
                            else asmname1 = dgv[colName.Index, curCell.RowIndex - j].Value.ToString();
                            j++;
                        }
                    if (autoText != null && (asmname1.IndexOf(".iam") != -1))
                    {
                        asmDoc = (AssemblyDocument)invApp.Documents.Open(doc.path() + "\\" + asmname1, false);
                        if (dgv[colReplace.Index, curCell.RowIndex].Value != null && dgv[colReplace.Index, curCell.RowIndex].Value.ToString() != "")
                        {
                            string rep = dgv[colReplace.Index, curCell.RowIndex].Value.ToString();
                            if (rep.EndsWith(".ipt"))
                            {
                                bool isContent = false;
                                fullName(asmDoc, rep, ref isContent);
                                //pDoc = (Inventor.PartDocument)invApp.Documents.Open(fullName(asmDoc, rep, ref isContent), false);
                                if (isContent)
                                {
                                    XMLDoc lib = new XMLDoc(@"C:\ProgramData\Autodesk\Inventor Addins\ContentCenter.xml", "Content");
                                    addItems(lib.Doc, DataCollection);
                                }
                            }
                        }
                        XMLDoc lib1 = new XMLDoc(@"C:\ProgramData\Autodesk\Inventor Addins\ContentCenter.xml", "Content");
                        addItems(lib1.Doc, DataCollection);

                        addItems(DataCollection, filter, exept, m_Parts.libraryPath);
                        autoText.AutoCompleteCustomSource = DataCollection;
                    }
                    break;
                default:
                    break;
            }
            }
            catch (Exception)
            {

                throw;
            }
        }

        public void addItems(AssemblyDocument doc, AutoCompleteStringCollection col)
        {
            AssemblyComponentDefinition compDef = doc.ComponentDefinition;
            HashSet<string> set = new HashSet<string>();
            foreach (ComponentOccurrence occ in compDef.Occurrences)
            {
                if (occ.Name.IndexOf(':') == -1) continue;
                string name = occ.Name.Substring(0, occ.Name.IndexOf(':'));
                //set.Add(occ.Name.Substring(0,occ.Name.IndexOf(':')));
                Document docum = null;
                try
                {
                   docum = (Document)occ.Definition.Document;
                }
                catch
                {
                    continue;
                }

                if (docum.DocumentType == DocumentTypeEnum.kPartDocumentObject)
                {
                    set.Add(name + ".ipt");
                }
                else
                {
                    set.Add(name + ".iam");
                }
            }
            if (set.Count != 0)
            {
                col.AddRange(set.ToArray());
            }
        }

        public string fullName (AssemblyDocument doc, string name, ref bool isContent)
        {
            name = name.Substring(0, name.Length - 4);
            AssemblyComponentDefinition compDef = doc.ComponentDefinition;
            foreach (ComponentOccurrence occ in compDef.Occurrences)
            {
                if (occ.Name.ToLower().IndexOf(name.ToLower()) != -1)
                {
                    if (occ.DefinitionDocumentType == DocumentTypeEnum.kPartDocumentObject)
                        isContent = (((PartComponentDefinition)occ.Definition).IsContentMember) ? true: false;
                    return ((Document)occ.Definition.Document).FullDocumentName;
                }
            }
            return "";
        }

        public void addItems(XDocument lib, AutoCompleteStringCollection col)
        {
            HashSet<string> set = new HashSet<string>();
            foreach (var item in lib.Root.Descendants("TableRow"))
            {
                if (item.Attribute("Extra") == null)     
                set.Add("#" + item.Attribute("Description").Value);
                else set.Add("#" + item.Attribute("Description").Value + "#" + item.Attribute("Extra").Value);
            }
            if (set.Count != 0)
            {
                col.AddRange(set.ToArray());
            }
        }

        public void addItems(XDocument desc, string nameAtt, string nameNode, AutoCompleteStringCollection col)
        {
            HashSet<string> set = new HashSet<string>();
            string filter = cBox.Text; 
            foreach (var item in desc.Root.Descendants(nameNode))
            {
                if (filter != "" && item.Attribute("type") != null &&item.Attribute("type").Value != filter) continue;  
                string val = item.Attribute(nameAtt).Value;
                if (val.IndexOf('$') != -1)
                {
                    string repl = val.Substring(val.IndexOf('$'), val.LastIndexOf('$') - val.IndexOf('$') + 1);
                    val = val.Replace(repl, this.txtbox1.Text);
                }
                set.Add(val);
            }
            if (set.Count != 0)
            {
                col.AddRange(set.ToArray());
            }
        }

        public void addItems(AutoCompleteStringCollection col, string filter, List<string> exept, System.IO.SearchOption opt)
        {
            string name = "";
            foreach (var item in System.IO.Directory.EnumerateFiles(doc.path(), filter, opt))
            {
                if (!exept.Exists(e => item.IndexOf(e) != -1))
                {
                   name = item.Substring(item.LastIndexOf('\\')+1, item.Length - 1 - item.LastIndexOf('\\'));
                   col.Add(name); 
                }
            }
        }

        public void addItems(AutoCompleteStringCollection col, string filter, List<string> exept, List<string> libPath)
        {
            string name = "";
            HashSet<string> hs = new HashSet<string>();
            foreach (string path in libPath)
            {
                foreach (var item in System.IO.Directory.EnumerateFiles(path, filter, System.IO.SearchOption.TopDirectoryOnly))
                {
                    if (!exept.Exists(e => item.IndexOf(e) != -1))
                    {
                        name = item.Substring(item.LastIndexOf('\\') + 1, item.Length - 1 - item.LastIndexOf('\\'));
                        hs.Add(name);
                    }
                }
            }
            col.AddRange(hs.ToArray());
        }

        private void разместитьВсеДеталиToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Hide();
            m_Parts.placeAllFamilyContent();
            this.Show();
        }

        private void добавитьАттрибутыToolStripMenuItem_Click(object sender, EventArgs e)
        {
            MyTextBox txtBOX = new MyTextBox();
            MyLabel lbl = new MyLabel();
            int offset = 10;
            System.Drawing.Point pt = new System.Drawing.Point(200, 200);
            Label lbl1 = lbl.addLabel("Название", pt, 100, 20);
            this.Controls.Add(lbl1);
            txtbox1 = txtBOX.addTextBox("введите название", new System.Drawing.Point(pt.X + lbl1.Width + offset, pt.Y), 200, 20);
            this.Controls.Add(txtbox1);
            pt.Y += 30;
            lbl1 = lbl.addLabel("ID", pt, 100, 20);
            this.Controls.Add(lbl1);
            txtbox2 = txtBOX.addTextBox("введите ID", new System.Drawing.Point(pt.X + lbl1.Width + offset, pt.Y), 200, 20);
            this.Controls.Add(txtbox2);
            pt.Y += 30;
            MyButton btn = new MyButton();
            System.Windows.Forms.Button btn1 = btn.addButton("Выполнить", pt, 200, 20);
            btn1.Click += new EventHandler(btnClick);
            this.Controls.Add(btn1);
        }
        private void btnClick(object sender, EventArgs e)
        {
            this.Hide();
            string name = "";
            if (System.IO.File.Exists(doc.path() + '\\' + this.txt2.Text))
            {
                name = doc.path() + '\\' + this.txt2.Text;
                if (System.IO.File.Exists(name + ".1")) System.IO.File.Delete(name + ".1");
                System.IO.File.Move(name, name + ".1");
            }
            createDescription(dgv, this.txt2.Text, doc.path());
            //this.Dispose();
            this.Show();
        }

        private void LoadClick(object sender, EventArgs e)
        {
            
        }

        private void loadtree(XElement el, string filter)
        {
            
        }

        private void крепежToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Hide();
            ContentOp c = new ContentOp();
            Transaction tr = Macros.StandardAddInServer.m_inventorApplication.TransactionManager.StartTransaction(Macros.StandardAddInServer.m_inventorApplication.ActiveDocument, "Крепеж");
            addFasteners(I.app.ActiveDocument as AssemblyDocument, c);
            tr.End();
            //this.Close();
            //this.Dispose();
        }

        static public void addFasteners(AssemblyDocument doc, ContentOp c)
        {
            if (doc == null) return;
            c.programmAdd(doc);
        }

        private void вставитьПарамЭлToolStripMenuItem_Click(object sender, EventArgs e)
        {
            
        }

        private void insertParam()
        {
            this.Hide();
            Document doc = Macros.StandardAddInServer.m_inventorApplication.ActiveDocument;
            if (doc.DocumentType == DocumentTypeEnum.kAssemblyDocumentObject)
            {
                try
                {
                    m_Parts.insertIFeature((PartDocument)doc.ActivatedObject, (AssemblyDocument)doc);
                }
                catch (Exception)
                {
                }
            }
            else if (doc.DocumentType == DocumentTypeEnum.kPartDocumentObject)
                m_Parts.insertIFeature((PartDocument)doc);
            this.Show();
        }

        private void удалитьЭлементToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Hide();
            Document doc = Macros.StandardAddInServer.m_inventorApplication.ActiveDocument;
            CommandManager cmd = Macros.StandardAddInServer.m_inventorApplication.CommandManager;
            if (doc.DocumentType == DocumentTypeEnum.kAssemblyDocumentObject)
            {
            ComponentOccurrence occ = (ComponentOccurrence)cmd.Pick(SelectionFilterEnum.kAssemblyOccurrenceFilter, "Выберите компонент для удаления");
            AssemblyComponentDefinition acd = ((AssemblyDocument)doc).ComponentDefinition;
            Parts.removeOcc(acd, occ);
            }
            this.Show();
        }

        private void addToAsm(AssemblyComponentDefinition compDef, string name)
        {
            invApp.SilentOperation = true;
            Matrix mtx = I.tg.CreateMatrix();
            ComponentOccurrence occ = compDef.Occurrences.Add(name, mtx);
            object wp1 = null, wp2 = null, wp3 = null;
            getPlanes(occ, ref wp1, ref wp2, ref wp3);
            WorkPlane awp = compDef.WorkPlanes[2];
            FlushConstraint fc = compDef.Constraints.AddFlushConstraint((WorkPlaneProxy)wp2, awp, 0);
            awp = compDef.WorkPlanes[3];
            fc = compDef.Constraints.AddFlushConstraint((WorkPlaneProxy)wp3, awp, 0);
            awp = compDef.WorkPlanes[1];
            fc = compDef.Constraints.AddFlushConstraint((WorkPlaneProxy)wp1, awp, 0);
            invApp.SilentOperation = false;
        }

        private void addToAsm()
        {
            if (Macros.StandardAddInServer.m_inventorApplication.ActiveDocument.DocumentType != DocumentTypeEnum.kAssemblyDocumentObject) return;
            AssemblyDocument asm_Doc = (AssemblyDocument)Macros.StandardAddInServer.m_inventorApplication.ActiveDocument;
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Filter = "Детали|*.ipt";
            ofd.Title = "Выберите файлы";
            ofd.Multiselect = true;

            string path = asm_Doc.FullFileName.ToString();
            path = path.Substring(0, path.LastIndexOf('\\'));
            path += "\\";

            ofd.InitialDirectory = path;
            ofd.ShowDialog();
            invApp.SilentOperation = true;
            Matrix mtx = I.tg.CreateMatrix();
            bool flag = false;
            foreach (string fn in ofd.SafeFileNames.OrderByDescending(n => n))
            {
                ComponentOccurrence occ = asm_Doc.ComponentDefinition.Occurrences.Add(path + fn, mtx);
                PartDocument pDoc = (PartDocument)occ.Definition.Document;
                SheetMetalFeatures smf = (pDoc.ComponentDefinition as SheetMetalComponentDefinition).Features as SheetMetalFeatures;
                if (smf.ContourFlangeFeatures.Count != 0)
                {
                    occ.Grounded = true;
                    flag = true;
                    continue;
                }
                if (smf.ContourFlangeFeatures.Count != 0 && flag)
                {
                    occ.Grounded = false;
                    //pDoc.ModelingSettings.AdaptivelyUsedInAssembly = true;
                    PartComponentDefinition compDef = (PartComponentDefinition)occ.Definition;
                    object wp1 = null, wp2 = null, wp3 = null;
                    getPlanes(occ, ref wp1, ref wp2, ref wp3);
                    WorkPlane awp = asm_Doc.ComponentDefinition.WorkPlanes[2];
                    FlushConstraint fc = asm_Doc.ComponentDefinition.Constraints.AddFlushConstraint((WorkPlaneProxy)wp2, awp, 0);
                    awp = asm_Doc.ComponentDefinition.WorkPlanes[3];
                    fc = asm_Doc.ComponentDefinition.Constraints.AddFlushConstraint((WorkPlaneProxy)wp3, awp, 0);
                    awp = asm_Doc.ComponentDefinition.WorkPlanes[1];
                    fc = asm_Doc.ComponentDefinition.Constraints.AddFlushConstraint((WorkPlaneProxy)wp1, awp, 0);
                    //occ.Adaptive = true;
                    continue;
                }
                if ( pDoc.ComponentDefinition.ReferenceComponents.DerivedPartComponents.Count != 0)
                {
                    DerivedPartDefinition def = pDoc.ComponentDefinition.ReferenceComponents.DerivedPartComponents[1].Definition;
                if (smf.FaceFeatures.Count != 0 && (def as DerivedPartUniformScaleDef).Mirror != true)
                {
                    occ.Grounded = false;
                    //ComponentOccurrence occ2 = asm_Doc.ComponentDefinition.Occurrences[1];
                    if (pDoc.ModelingSettings.AdaptivelyUsedInAssembly) pDoc.ModelingSettings.AdaptivelyUsedInAssembly = false;
                    //pDoc.ModelingSettings.AdaptivelyUsedInAssembly = true;
                    PartComponentDefinition compDef = (PartComponentDefinition)occ.Definition;
                    object wp1 = null, wp2 = null, wp3 = null;
                    getPlanes(occ, ref wp1, ref wp2, ref wp3);
                    WorkPlane awp = asm_Doc.ComponentDefinition.WorkPlanes[2];
                    FlushConstraint fc = asm_Doc.ComponentDefinition.Constraints.AddFlushConstraint((WorkPlaneProxy)wp2, awp, 0);
                    awp = asm_Doc.ComponentDefinition.WorkPlanes[3];
                    fc = asm_Doc.ComponentDefinition.Constraints.AddFlushConstraint((WorkPlaneProxy)wp3, awp, 0);
                    string[] param = { "БВ_длина" };
                    if (addLinkParam((Document)asm_Doc, (Document)asm_Doc.ComponentDefinition.Occurrences[1].Definition.Document, param))
                    {
                        awp = asm_Doc.ComponentDefinition.WorkPlanes[1];
                        //occ2.CreateGeometryProxy(((PartComponentDefinition)occ2.Definition).WorkPlanes["Шип_справа"], out wp1);
                        //Plane pla = compDef.Sketches["Шип_проекция"].PlanarEntity as Plane;
                        //occ.CreateGeometryProxy(pla, out wp2);
                        Face fa = occ.SurfaceBodies[1].Faces.OfType<Face>().Where(f => f.CreatedByFeature is FaceFeature).OrderByDescending(o => o.Evaluator.Area).ElementAt(1);
                        asm_Doc.ComponentDefinition.Constraints.AddMateConstraint(awp, fa, "БВ_длина/2" /*+ " + (thick * 10).ToString()*/ /*"БВ_длина/2"*/);
                    }
                    occ.Adaptive = true;
                    continue;
                }

                if ((def as DerivedPartUniformScaleDef).Mirror == true)
                {
                    occ.Grounded = false;
                    PartComponentDefinition compDef = (PartComponentDefinition)occ.Definition;
                    object wp1 = null, wp2 = null, wp3 = null;
                    getPlanes(occ, ref wp1, ref wp2, ref wp3);
                    WorkPlane awp = asm_Doc.ComponentDefinition.WorkPlanes[2];
                    FlushConstraint fc = asm_Doc.ComponentDefinition.Constraints.AddFlushConstraint((WorkPlaneProxy)wp2, awp, 0);
                    awp = asm_Doc.ComponentDefinition.WorkPlanes[3];
                    fc = asm_Doc.ComponentDefinition.Constraints.AddFlushConstraint((WorkPlaneProxy)wp3, awp, 0);
                    string[] param = { "БВ_длина" };
                    if (addLinkParam((Document)asm_Doc, (Document)asm_Doc.ComponentDefinition.Occurrences[1].Definition.Document, param))
                    {
                        awp = asm_Doc.ComponentDefinition.WorkPlanes[1];
                        Face fa = occ.SurfaceBodies[1].Faces.OfType<Face>().OrderByDescending(o => o.Evaluator.Area).ElementAt(0);
                        //double thick = double.Parse(((SheetMetalComponentDefinition)compDef).Thickness.Value.ToString());
                        asm_Doc.ComponentDefinition.Constraints.AddFlushConstraint(awp, fa, "-БВ_длина/2 ");
                    }
                }
                }

//                 if (pDoc.PropertySets[3][14].Value.ToString().ToLower().IndexOf("фланец левый") != -1)
//                 {
//                     occ.Grounded = false;
//                     //ComponentOccurrence occ2 = asm_Doc.ComponentDefinition.Occurrences[1];
//                     //if (pDoc.ModelingSettings.AdaptivelyUsedInAssembly) pDoc.ModelingSettings.AdaptivelyUsedInAssembly = false;
//                     //pDoc.ModelingSettings.AdaptivelyUsedInAssembly = true;
//                     PartComponentDefinition compDef = (PartComponentDefinition)occ.Definition;
//                     object wp1 = null, wp2 = null, wp3 = null;
//                     getPlanes(occ, ref wp1, ref wp2, ref wp3);
//                     WorkPlane awp = asm_Doc.ComponentDefinition.WorkPlanes[2];
//                     FlushConstraint fc = asm_Doc.ComponentDefinition.Constraints.AddFlushConstraint((WorkPlaneProxy)wp2, awp, 0);
//                     awp = asm_Doc.ComponentDefinition.WorkPlanes[3];
//                     fc = asm_Doc.ComponentDefinition.Constraints.AddFlushConstraint((WorkPlaneProxy)wp3, awp, 0);
//                     string[] param = { "БВ_длина" };
//                     if (addLinkParam((Document)asm_Doc, (Document)pDoc, param))
//                     {
//                         awp = asm_Doc.ComponentDefinition.WorkPlanes[1];
//                         //occ2.CreateGeometryProxy(((PartComponentDefinition)occ2.Definition).WorkPlanes["Шип_слева"], out wp1);
//                         //occ.CreateGeometryProxy(compDef.ReferenceComponents.DerivedPartComponents[1].Sketches["Шип_проекция"]., out wp2);
//                         //double thick = double.Parse(((SheetMetalComponentDefinition)compDef).Thickness.Value.ToString());
//                         fc = asm_Doc.ComponentDefinition.Constraints.AddFlushConstraint(awp, (WorkPlaneProxy)wp1, /*0.1 - thick*/ "-БВ_длина/2");
//                     }
//                     //occ.Adaptive = true;
//                     continue;
//                 }
            }
        }

        private void добавитьВСборкуToolStripMenuItem_Click(object sender, EventArgs e)
        {
            
        }
        public static void addLinkParam(string param, Document doc)
        {
            Parameter p = getParameter(doc, param);
            if (p != null) return;
            if (doc.SubType != "{9C464203-9BAE-11D3-8BAD-0060B0CE6BB4}") return;
            SheetMetalComponentDefinition smcd = (doc as PartDocument).ComponentDefinition as SheetMetalComponentDefinition;
            if (smcd.ReferenceComponents.DerivedPartComponents.Count != 1) return;
            DerivedPartComponent dpc = smcd.ReferenceComponents.DerivedPartComponents[1];
            DerivedPartUniformScaleDef def = dpc.Definition as DerivedPartUniformScaleDef;
            foreach (DerivedPartEntity ent in def.Parameters)
            {
                UserParameter rp = ent.ReferencedEntity as UserParameter;
                if (rp == null) continue;
                string n = rp.Name;
                if (n == param)
                {
                    ent.IncludeEntity = true;
                    rp.ExposedAsProperty = true;
                    smcd.ReferenceComponents.DerivedPartComponents[1].Definition = (DerivedPartDefinition)def;
                    //smcd.ReferenceComponents.DerivedPartComponents.Add((DerivedPartDefinition)def);
                    //exposed(dpc.ReferencedDocumentDescriptor.ReferencedDocument as PartDocument, param);
                    return;
                }
            }
        }
        public static void exposed(PartDocument doc, string s)
        {
            foreach (UserParameter item in doc.ComponentDefinition.Parameters.UserParameters)
            {
                if (item.Name == s)
                   { item.ExposedAsProperty = true; return; }
            }
        }
        public static bool addLinkParam(Document to, Document from, string [] param)
        {
            ObjectCollection col = Macros.StandardAddInServer.m_inventorApplication.TransientObjects.CreateObjectCollection();
            bool flag = true;
            foreach (var item in param)
            {
                Parameter p = getParameter(from, item);
                Parameter p1 = getParameter(to, item);
                if (p != null) 
                {  
                    if (p1 == null)
                    {
                        p.ExposedAsProperty = true;
                        col.Add(p);
                    }
                }
                else flag = false;
            }
            if (to.DocumentType == DocumentTypeEnum.kPartDocumentObject)
            {
                PartComponentDefinition compDef = ((PartDocument)to).ComponentDefinition;
                foreach (DerivedParameterTable dpt in compDef.Parameters.DerivedParameterTables)
                {
//                     if (dpt.ReferencedDocumentDescriptor.FullDocumentName == from.FullDocumentName)
//                     {
// //                         foreach (DerivedParameter t in dpt.DerivedParameters)
// // 	                    {
// // 		                    col.Add(t);
// // 	                    }
//                         dpt.LinkedParameters = col;
//                         return flag;
//                     }
                }
                if (col.Count != 0)
                compDef.Parameters.DerivedParameterTables.Add2(from.FullDocumentName, col);
            }
            else if (to.DocumentType == DocumentTypeEnum.kAssemblyDocumentObject)
            {
                AssemblyComponentDefinition compDef = ((AssemblyDocument)to).ComponentDefinition;
                if (col.Count != 0)
                compDef.Parameters.DerivedParameterTables.Add2(from.FullDocumentName, col);
            }
            return flag;
        }
        static public Parameter getParameter(Document doc, string name)
        {
            Parameter p = null;
            try
            {
                if (doc.DocumentType == DocumentTypeEnum.kPartDocumentObject)
                {
                    PartComponentDefinition compDef = (PartComponentDefinition)((PartDocument)doc).ComponentDefinition;
//                     foreach (DerivedParameterTable dpt in compDef.Parameters.DerivedParameterTables)
//                     {
//                         dpt.DerivedParameters[] 
//                     }
                    p = compDef.Parameters[name];
                }
                else if (doc.DocumentType == DocumentTypeEnum.kAssemblyDocumentObject)
                {
                    AssemblyComponentDefinition compDef = (AssemblyComponentDefinition)((AssemblyDocument)doc).ComponentDefinition;
                    p = compDef.Parameters[name];
                }
                return p;
            }
            catch { return null; }
        }
        public void getPlanes(ComponentOccurrence occ, ref object wp1, ref object wp2, ref object wp3)
        {
            if (((Document)occ.Definition.Document).DocumentType == DocumentTypeEnum.kPartDocumentObject)
            {
                PartComponentDefinition compDef = (PartComponentDefinition)occ.Definition;
                occ.CreateGeometryProxy(compDef.WorkPlanes[1], out wp1);
                occ.CreateGeometryProxy(compDef.WorkPlanes[2], out wp2);
                occ.CreateGeometryProxy(compDef.WorkPlanes[3], out wp3);
            }
        }

        private void spikes()
        {
            double R = 3.0 / 10, L = 5.0 / 10, H = 7.5 / 10;
            spike sp = new spike(Macros.StandardAddInServer.m_inventorApplication);
            sp.smcd = (SheetMetalComponentDefinition)((PartDocument)sp.invApp.ActiveDocument).ComponentDefinition;
            sp.R = R; sp.L = L; sp.H = H;
            sp.addSpikeDef(sp.smcd, R, R * 2, H);
            sp.addSpikeDef(sp.smcd, R, L, H, "Капелька");
                sp.addSketch(sp.smcd, H, R, L);
        }

        private void шипToolStripMenuItem_Click(object sender, EventArgs e)
        {
            
        }

        public void addCuts(AssemblyDocument aDoc ,AssemblyComponentDefinition acd, string [] nameIn, string nameOut, String nameSketch, string nameCut)
        {
            SheetMetalComponentDefinition smcd = null;
            PlanarSketch psp = Offset.projectAcrosParts(ref acd, nameIn, nameOut, nameSketch, ref smcd);
            try
            {
                Offset.offsetAdaptive(smcd, psp.Name, 0.11, 0.8, 0.63);
                psp.Visible = false;
            }
            catch (Exception)
            {
            }
        }

        private void cuts()
        {
            if (Macros.StandardAddInServer.m_inventorApplication.ActiveDocument.DocumentType == DocumentTypeEnum.kAssemblyDocumentObject)
            {
                AssemblyDocument aDoc = (AssemblyDocument)Macros.StandardAddInServer.m_inventorApplication.ActiveDocument;
                AssemblyComponentDefinition acd = aDoc.ComponentDefinition;

                //Offset.projectAcrosParts(ref acd, new string[] { "обтекатель", "язык" }, "фланец правый", "Шип_проекция", ref smcd);

                addCuts(aDoc, acd, new string[] { "обтекатель", "язык", "экран блока" }, "фланец правый", "Шип_проекция", "Пазы");
            }
            //else if (Macros.StandardAddInServer.m_inventorApplication.ActiveDocument.DocumentType == DocumentTypeEnum.kPartDocumentObject)
            //{
            //    PartDocument pDoc = (PartDocument)Macros.StandardAddInServer.m_inventorApplication.ActiveDocument;
            //    SheetMetalComponentDefinition smcd = (SheetMetalComponentDefinition)pDoc.ComponentDefinition;
            //    Offset.offsetAdaptive(smcd, "Шип_проекция", 0.11, 0.8, 0.63);
            //}
        }

        private void пазToolStripMenuItem_Click(object sender, EventArgs e)
        {
            
        }

        private void menuStrip1_ItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {

        }

        private void путиToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            InvDocument<Document> invDoc = new InvDocument<Document>(Macros.StandardAddInServer.m_inventorApplication.ActiveDocument);
            if (invDoc.pathes.Count == 0)
                invDoc.pathes = null;
            //string path = invDoc.path + System.DateTime.Now.ToString("dd:mm:yyyy") + "\\";
            //if (!System.IO.Directory.Exists(path)) System.IO.Directory.CreateDirectory(path);
            string path = invDoc.path + invDoc.getType + "(" +  System.DateTime.Now.ToString("dd.MM.yyyy") + ").xml";
            XMLDoc xmlDoc = new XMLDoc(path, "Path");
            xmlDoc.Doc.Root.Add(new XElement("Name", invDoc.getType + "(" +  System.DateTime.Now.ToString("dd.MM.yyyy") + ")"));
            foreach (var item in invDoc.pathes)
            {
                xmlDoc.Doc.Root.Add(new XElement("Value", item));
            }
            xmlDoc.save();
        }

        private void сохранитьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (Macros.StandardAddInServer.m_inventorApplication.ActiveDocument.DocumentType == DocumentTypeEnum.kAssemblyDocumentObject)
            Table.saveData((AssemblyDocument)Macros.StandardAddInServer.m_inventorApplication.ActiveDocument);
        }

        private void отверстияToolStripMenuItem_Click(object sender, EventArgs e)
        {
            
        }

        private void extrude()
        {
            Document doc = Macros.StandardAddInServer.m_inventorApplication.ActiveDocument;
            if (doc.SubType == "{9C464203-9BAE-11D3-8BAD-0060B0CE6BB4}")
            {
                SheetMetalComponentDefinition smcd = (SheetMetalComponentDefinition)((PartDocument)doc).ComponentDefinition;
                Parameter param = smcd.Parameters.OfType<Parameter>().FirstOrDefault(p => p.Name.ToLower().IndexOf("длина") != -1);
                if (param != null)
                    Parts.addContourFlange(smcd, param.Name);
            }
        }

        private void выдавитьToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void текущийДокументToolStripMenuItem_Click(object sender, EventArgs e)
        {
            MyXML xml = new MyXML("AddDrawingInterface.xml");
            MyForm f = new MyForm(xml, "Чертеж"/*, null,*/ );
            DialogResult dr = f.f.ShowDialog();
            if (dr != DialogResult.OK) return;
            List<string> values = new List<string>();
            for (int i = 0; i < f.cbs.Count(); i++)
            {
                values.Add(f.cbs[i].Text);
            }
            values.Add(f.chks[0].Checked.ToString()); values.Add(f.chks[1].Checked.ToString());
            drawing(I.app.ActiveDocument, values);
            //PartsBtn.xDocSheet = new XMLDoc(@"C:\ProgramData\Autodesk\Inventor Addins\sheet.xml","head");
            //formaForSheet();

            //drawing(Macros.StandardAddInServer.m_inventorApplication.ActiveDocument, true);
            //this.Close();
        }
        public void addEl(XElement par, string nameAtt, string attVal, string align)
        {
            par.Add(new XElement("view1", new XAttribute(nameAtt, attVal), new XAttribute("align", align)));
        }
        public void drawing(Document doc, List<string> values, bool open = false)
        {
            InvDoc.InvDocument<Document> invDoc = new InvDoc.InvDocument<Document>(doc);
            XMLDoc xmlDoc = new XMLDoc(@"C:\ProgramData\Autodesk\Inventor Addins\sheet.xml", "head");
            List<double> scales = new List<double>();
            foreach (var item in xmlDoc.El.Element("Scales").Elements())
            {
                scales.Add(u.convToDouble(item.Value));
            }
            string pn, desc; double offset = 1.5;
            invDoc.doc = doc;
            DrawingView dv = null;
            //pn = invDoc.getProp("Part Number").Value.ToString();
            desc = invDoc.getProp("Description").Value.ToString();
            DrawingDocument drw = invDoc.addDrwDoc();
            try
            {
                string name = invDoc.path + u.nameForSave(doc, true); /*invDoc.path + pn + " (" + desc + ").idw";*/
                drw.SaveAs(name, false);
            if (doc.DocumentType == DocumentTypeEnum.kAssemblyDocumentObject || doc.DocumentType == DocumentTypeEnum.kPartDocumentObject)
            {
            //    AssemblyDocument asmDoc = (AssemblyDocument)doc;
            //}
            //else if (doc.DocumentType == DocumentTypeEnum.kPartDocumentObject)
            //{
            //    PartDocument pDoc = (PartDocument)doc;
                SheetMetalComponentDefinition smcd = invDoc.getSheetMetalCompDef;
                string val = desc.ToLower();
                invDoc.doc = (Document)drw;
                //var ie = xmlDoc.Doc.Root.Descendants("sheet");
                //Regex reg = new Regex(val);
                XElement el = xmlDoc.El.Elements("view").FirstOrDefault(elem => elem.Attribute("name").Value == values[0]);
                double w = 420, h = 297;
                switch (values[2])
                {
                    case "A4":
                        w = 210;
                        break;
                    case "A3":
                        w = 420;
                        break;
                    default:
                        w = u.convToDouble(values[2]);
                        break;
                }
                if (values[4] != "") addEl(el, "offsetX", values[4], values[6]);
                if (values[5] != "") addEl(el, "offsetY", values[5], values[7]);
                if (el != null)
                {
                   if (values[1] == "0")
                   {
                       dv = InvDocument<PartDocument>.addView(drw.Sheets[1], doc, el, "view1", 0, scales);
                   }
                   else
                   {
                       dv = InvDocument<PartDocument>.addView(drw.Sheets[1], doc, el, "view1", u.convToDouble(values[1]), scales);
                   }
                   if (drw.PropertySets[6][8].Value.ToString() == "") drw.PropertySets[6][8].Value = dv.ScaleString;
                   XMLDoc xdoc = null; string tmp = null;
                   if (el.Attribute("TR") != null)
                   {
                       xdoc = new XMLDoc(@"C:\ProgramData\Autodesk\Inventor Addins\" + el.Attribute("TR").Value, "TR");
                       addTR(drw.Sheets[1], xdoc);
                   }
                   else
                   {
                       if (doc.DocumentType == DocumentTypeEnum.kPartDocumentObject)
                           tmp = xmlDoc.El.Element("TR").Elements().ElementAt(0).Attribute("value").Value;
                       else if (doc.DocumentType == DocumentTypeEnum.kAssemblyDocumentObject)
                           tmp = xmlDoc.El.Element("TR").Elements().ElementAt(1).Attribute("value").Value;
                       xdoc = new XMLDoc(@"C:\ProgramData\Autodesk\Inventor Addins\" + tmp, "TR");
                       addTR(drw.Sheets[1], xdoc);
                   }
                }
                invDoc.changeSheet(drw.Sheets[1], w, h);
                Vector2d vec = I.tg.CreateVector2d(0, dv.Height / 2 + offset);
                if (smcd != null)
                {
                    if (!smcd.HasFlatPattern) smcd = null;
                    else if (smcd.Bends.Count == 0) smcd = null;
                    if (smcd == null)
                    {
                        //drw.SelectSet.Select(dv);
                        InvDocument<PartDocument>.addSketchedSymbol(drw.Sheets[1], "АР", new string[] { "" }, dv.Position, vec);
                    }
                }
                if (smcd != null)
                {
                    if (el != null)
                    {
                        switch (values[3])
                        {
                            case "A4":
                                w = 210;
                                break;
                            case "A3":
                                w = 420;
                                break;
                            default:
                                w = u.convToDouble(values[2]);
                                break;
                        }
                    }
                    invDoc.addSheet(w, h);
                    if (el != null)
                    {
                        if (values[1] == "0")
                        {
                            dv = InvDocument<PartDocument>.addView(drw.Sheets[2], doc, el, "view2", 0, scales);
                            //dv.Scale = InvDocument<string>.autoScale(dv, scales);
                        }
                        else
                        {
                            dv = InvDocument<PartDocument>.addView(drw.Sheets[2], doc, el, "view2", u.convToDouble(values[1]), scales);
                        }
                        if (el.Attribute("Dim") != null)
                        {
                            Dimensions dimens = new Dimensions(dv);
                            //Tables.addBends();
                        }
                        //drw.SelectSet.Select(dv);
                        DrawingViewLabel dvl = dv.Label;
                        dvl.FormattedText = @"<StyleOverride Font='AIGDT' Italic='false'>/</StyleOverride> (<DrawingViewScale/>) Раскрой методом АР";
                        //dvl.Position = I.CP2d(dvl.Position.X, dvl.Position.Y - 2); 
                        dv.ShowLabel = true;
                        //InvDocument<PartDocument>.addSketchedSymbol(drw.Sheets[2], "Развертка 1:N АР", new string[] { (1/dv.Scale).ToString("#.#")},
                        //    dv.Position, vec);
                    }
                    drw.Sheets[1].Activate();
                }
            }
            drw.Save2();
            if (open)
            {
                invDoc.openDrwDoc(drw.FullFileName, true);
            }
            else
                drw.Close();

            }
            catch (Exception)
            {
                if (drw != null) drw.Close();
            }
        }

        private void добавитьЛистToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (Macros.StandardAddInServer.m_inventorApplication.ActiveDocument.DocumentType == DocumentTypeEnum.kDrawingDocumentObject)
            {
                sheet sh = new sheet();
                sh.ShowDialog();
            }
        }

        private void всеОткрытыеToolStripMenuItem_Click(object sender, EventArgs e)
        {
            foreach (Document doc in Macros.StandardAddInServer.m_inventorApplication.Documents.VisibleDocuments)
            {
                //drawing(doc); 
            }
        }

        private void переназватьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            InvDocument<Document> invDoc = new InvDocument<Document>(Macros.StandardAddInServer.m_inventorApplication.ActiveDocument);
            Inventor.Application app = Macros.StandardAddInServer.m_inventorApplication;
            List<string> files = new List<string>();
            Document doc; string pn = "", desc = "", name = "", ext = "";
            doc = Macros.StandardAddInServer.m_inventorApplication.ActiveDocument;          
            if (doc.DocumentType == DocumentTypeEnum.kDrawingDocumentObject)
            {
                files = invDoc.openFiles("*.idw", true);
            }
            else
            {
                files = invDoc.openFiles("*.ipt", true);
                files.AddRange(invDoc.openFiles("*.iam", true));
                files.AddRange(invDoc.openFiles("*.idw", true));
            }
                app.SilentOperation = true;

            if (ext == ".idw") invDoc.nvmOptions.Add("DeferUpdates", true);
                Inventor.ProgressBar pb = app.CreateProgressBar(true, files.Count, "Переименование файлов");
                pb.Message = "Подождите...";

                XMLDoc xmlDoc = new XMLDoc(invDoc.path + "compare.xml", "head");
                foreach (var f in files)
                {
                    try
                    {
                        doc = invDoc.openDoc(f);
                        
                        if (doc.DocumentType != DocumentTypeEnum.kDrawingDocumentObject)
                        {
                            pn = doc.PropertySets[3][2].Value.ToString();
                            if (pn == "") continue;
                            desc = doc.PropertySets[3][14].Value.ToString();
                            if (desc == "") continue;
                            ext = System.IO.Path.GetExtension(f);
                            if (doc.RequiresUpdate) doc.Update2(false);
                            doc.Close();
                        }
                        else
                        {
                        if (doc.ReferencedDocuments.Count == 0) 
                        {
                            doc.Close(); continue; 
                        }
                        pn = InvDoc.u.referendedDoc(doc).PropertySets[3][2].Value.ToString();
                        desc = InvDoc.u.referendedDoc(doc).PropertySets[3][14].Value.ToString();
                        ext = System.IO.Path.GetExtension(f);
                        
                        doc.Close();
                        }
                        name = pn + " (" + desc + ")" + ext;
                        int ind = f.LastIndexOf('\\');
                        string newName = f.Remove(ind + 1) + name;
                        if (f != newName)
                        {
                            xmlDoc.El.Add(new XElement("name", new XAttribute("old", f), new XAttribute("new", newName)));
                            System.IO.File.Move(f, newName);
                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.ToString());
                    }
                    pb.UpdateProgress();
                }
                pb.Close();
                xmlDoc.save();
                
                app.SilentOperation = false;
        }

        public static XMLDoc xml(XElement el)
        {
            XMLDoc xdoc = new XMLDoc(@"C:\ProgramData\Autodesk\Inventor Addins\rename.xml", "head");
            foreach (var item in el.Descendants())
            {
                if (!item.HasAttributes) continue;
                if (item.Attribute("old").Value == item.Attribute("new").Value) continue;
                XElement fel = xdoc.find(item.FirstAttribute.Name.ToString(),item.FirstAttribute.Value.ToString());
                if (fel != null) fel.Remove();
                //xdoc.addXElement(item);
                XElement nel = xdoc.addXElement("row");
                XMLDoc.addXAttributes(nel, new Dictionary<string, string>() { { "old", item.Attribute("old").Value }, { "new", item.Attribute("new").Value } });
                xdoc.El.Add(nel);
            }
            xdoc.save();
            return xdoc;
        }

        public static void renameAsm(Document doc, bool drw = false)
        {
            XElement bel = new XElement("row");
            addName(doc, bel);
            XElement oldNames = names(doc, bel, null);

            IEnumerable<Document> drws = null;
            XMLDoc xdoc = new XMLDoc(@"C:\ProgramData\Autodesk\Inventor Addins\rename1.xml", "head");
            xml(oldNames);
            xdoc.El.Add(oldNames);
            if (drw) drws = openDrw(doc);
            foreach (Document item in drws)
            {
                repair(item, xdoc);
            }
            
//             doc.Save2();
//             doc.Close();
//             if (bel.Attribute("old") != null && bel.Attribute("old").Value != bel.Attribute("new").Value)
//             {
//                 System.IO.File.Move(bel.Attribute("old").Value, bel.Attribute("new").Value);
//             }
        }

        private void габаритныеРазмерыToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Document doc = Macros.StandardAddInServer.m_inventorApplication.ActiveDocument;
            renameAsm(doc);
        }

        public static void openAsms(Document doc)
        {
            Inventor.Application app = doc.Parent as Inventor.Application;
            app.SilentOperation = true;
            NameValueMap nvmOptions = app.TransientObjects.CreateNameValueMap();
            nvmOptions.Add("SkipAllUnresolvedFiles", true);
            //nvmOptions.Add("DeferUpdates", true);
            string path = app.DesignProjectManager.ActiveDesignProject.WorkspacePath;
            foreach (var ffn in System.IO.Directory.EnumerateFiles(path, "*.iam", System.IO.SearchOption.AllDirectories))
            {
                if (ffn.EndsWith(".iam") && ffn.IndexOf("OldVersions") == -1)
                    app.Documents.OpenWithOptions(ffn, nvmOptions, false);
            }
            app.SilentOperation = false;
        }

        public static IEnumerable<Document> openDrw(Document doc)
        {
            I.silent(true);
            var files = I.getFiles<Document>(file.p(I.aDoc().FullFileName), "idw");
            I.silent(false);
            return files;

//             List<Document> docs = new List<Document>();
//             Inventor.Application app = doc.Parent as Inventor.Application;
//             app.SilentOperation = true;
//             NameValueMap nvmOptions = app.TransientObjects.CreateNameValueMap();
//             nvmOptions.Add("SkipAllUnresolvedFiles", true);
//             //nvmOptions.Add("DeferUpdates", true);
//             string path = System.IO.Path.GetDirectoryName(doc.FullFileName); //app.DesignProjectManager.ActiveDesignProject.WorkspacePath;
//             foreach (var ffn in System.IO.Directory.EnumerateFiles(path, "*.idw", System.IO.SearchOption.AllDirectories))
//             {
//                 if (ffn.EndsWith(".idw") && ffn.IndexOf("OldVersions") == -1)
//                 docs.Add(app.Documents.OpenWithOptions(ffn, nvmOptions, false)); 
//             }
//             app.SilentOperation = false;
//             return docs;
        }

        public static bool findInVisibleDoc(Document doc)
        {
            foreach (Document item in Macros.StandardAddInServer.m_inventorApplication.Documents.VisibleDocuments)
            {
                if (item.Equals(doc)) return true;
            }
            return false;
        }

        public static void move(string o, string n)
        {
            if (System.IO.File.Exists(n))
            {
                string p = System.IO.Path.GetDirectoryName(n), name = System.IO.Path.GetFileName(n);
                p += @"\old\";
                if (!System.IO.Directory.Exists(p)) System.IO.Directory.CreateDirectory(p);
                System.IO.File.Move(n, p + name);
            }
            System.IO.File.Move(o, n);
        }

        public static void repair(Document doc, XMLDoc xdoc)
        {
            string p = System.IO.Path.GetDirectoryName(doc.FullDocumentName);
            string n = "", o = ""; bool flag = true;
            //string p = System.IO.Path.GetDirectoryName(doc.FullDocumentName);
            if (xdoc == null) return;
            foreach (DocumentDescriptor dd in doc.ReferencedDocumentDescriptors)
            {
                string fdn = dd.FullDocumentName;
                XElement el = xdoc.find("old", fdn);
                if (el == null)
                {
                    fdn = System.IO.Path.Combine(p, System.IO.Path.GetFileName(fdn));
                    el = xdoc.find("old", fdn);
                }
                if (el == null) continue;
                if (flag) {flag = false; n = System.IO.Path.GetFileNameWithoutExtension(el.Attribute("new").Value) + ".idw";}
                if (el.Attribute("new").Value.ToString().Trim() == el.Attribute("old").Value.ToString().Trim()) continue;
                if (System.IO.File.Exists(el.Attribute("new").Value) && dd.FullDocumentName != el.Attribute("new").Value)
                    dd.ReferencedFileDescriptor.ReplaceReference(el.Attribute("new").Value);
            }
            o = doc.FullDocumentName;
            
            if (o != "" && System.IO.File.Exists(o) && n != "" && !System.IO.File.Exists(n) && o != n)
            {
                update(doc as Document);
                doc.Close();
                move(o, n);
            }
        }

        public static bool rename(Document doc, XElement el)
        {
            if (el.Attribute("old") != null && el.Attribute("old").Value != el.Attribute("new").Value)
            {
                Document drw = null, asm = null;

                if (System.IO.File.Exists(el.Attribute("old").Value))

                    //System.IO.File.Move(el.Attribute("old").Value, el.Attribute("new").Value);
                {
                    move(el.Attribute("old").Value, el.Attribute("new").Value);
                    //doc = I.app.Documents.Open(el.Attribute("new").Value);
                }
                if (doc != null && doc.ReferencingDocuments != null && doc.ReferencingDocuments.Count != 0)     
                {
                    
                    foreach (Document item in doc.ReferencingDocuments)
                    {
                        drw = null; asm = null;
                        if (item.DocumentType == DocumentTypeEnum.kDrawingDocumentObject)
                            drw = item;
                        else if (item.DocumentType == DocumentTypeEnum.kAssemblyDocumentObject)
                            asm = item;
                        if (drw != null)
                        {
                            foreach (DocumentDescriptor dd in drw.ReferencedDocumentDescriptors)
                            {
                                if (dd.FullDocumentName == el.Attribute("old").Value)
                                    dd.ReferencedFileDescriptor.ReplaceReference(el.Attribute("new").Value);
                            }
                            update(drw as Document);
                            //drw.Save2();
                            string oldName = drw.FullFileName;
                            string path = System.IO.Path.GetDirectoryName(oldName), ext = System.IO.Path.GetExtension(oldName), newName = System.IO.Path.GetFileNameWithoutExtension(el.Attribute("new").Value);
                            if (oldName != path + "\\" + newName + ext)
                            {
                                drw.ReleaseReference();
                                System.IO.File.Move(oldName, path + "\\" + newName + ext);
                                NameValueMap nvmOptions = I.objs.CreateNameValueMap();
                                nvmOptions.Add("SkipAllUnresolvedFiles", true);
                                Macros.StandardAddInServer.m_inventorApplication.Documents.OpenWithOptions(path + "\\" + newName + ext, nvmOptions, false);
                            }

                        }
                        if (asm != null)
                        {
                            foreach (DocumentDescriptor dd in asm.ReferencedDocumentDescriptors)
                            {
                                if (dd.FullDocumentName == el.Attribute("old").Value && System.IO.File.Exists(el.Attribute("new").Value))
                                    dd.ReferencedFileDescriptor.ReplaceReference(el.Attribute("new").Value);
                            }
                            update(asm as Document);
                            //asm.Save2();
                        }
                    }
                }
                //doc.Close();
                //doc.Save2();
                return true;
            }
            return false;
        }

        public static XElement names(Document asm, XElement bel, DocumentDescriptor dd)
        {
            XElement el = null;
            //System.Diagnostics.Process[] pr = System.Diagnostics.Process.GetProcessesByName("Inventor");
            foreach (DocumentDescriptor d in asm.ReferencedDocumentDescriptors)
            {
                if (d.ReferencedFileDescriptor.LibraryName != null) continue;
                Document doc = d.ReferencedDocument as Document;   
                el = new XElement("row");
                addName(doc, el);
                if (doc.DocumentType == DocumentTypeEnum.kAssemblyDocumentObject)
                {
                    el = names(doc, el, d);
                }
                if (rename(doc,el))
                {
                    d.ReferencedFileDescriptor.ReplaceReference(el.Attribute("new").Value);
                    //doc.Save2();
                }
                bel.Add(el);
            }
            XElement elem = new XElement("row");
            addName(asm, elem); 
            if (rename(asm, elem) && dd != null)
            {
                dd.ReferencedFileDescriptor.ReplaceReference(elem.Attribute("new").Value);
                //asm.Save2();
            }
            elem.Add(bel);
            return elem;
        }

        public static bool addName(Document doc, XElement el)
        {
            string ffn = doc.FullFileName; string path = System.IO.Path.GetDirectoryName(ffn), ext = System.IO.Path.GetExtension(ffn);
            string name = System.IO.Path.GetFileNameWithoutExtension(ffn);
            if (name[name.Length - 3] == '^') name = name.Substring(name.Length - 3, 3);
            else name = "";
            if (ffn.IndexOf("Content Center") != -1) { el = null; return false; }
            string pn = InvDoc.u.getProp(doc, "Part Number").Value.ToString().Trim(), desc = InvDoc.u.getProp(doc, "Description").Value.ToString();
            if (pn == "") { el = null ;return false; }
            string n = path + "\\" + pn + " (" + desc + ")" + name + ext;
            el.Add(new XAttribute("old", ffn), new XAttribute("new", n));
            return true;
        }

        void btn_Click(object sender, EventArgs e)
        {
            Form f = (Form)((System.Windows.Forms.Button)sender).Parent;
            foreach (Document doc in Macros.StandardAddInServer.m_inventorApplication.Documents.VisibleDocuments)
            {
                ut.addProp(doc, f.Controls[2].Text, f.Controls[3].Text); 
            }
        }

        private void разместитьВсеДеталиToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            m_Parts.placeAllFamilyContent();
        }

        private void обновитьContentCenterToolStripMenuItem_Click(object sender, EventArgs e)
        {
            m_Parts = new Parts(Macros.StandardAddInServer.m_inventorApplication.ActiveDocument);
            m_Parts.createContentXML();
        }

        private void copyAttrToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            InvDoc.InvDocument<DrawingDocument> doc = new InvDocument<DrawingDocument>((DrawingDocument)Macros.StandardAddInServer.m_inventorApplication.ActiveDocument);
            Sheet sh = doc.GetDoc.Sheets[1];
            XMLDoc xdoc = new XMLDoc(doc.path + "TR.xml", "TR");
            copyAttr(sh, xdoc);
        }

        public static void copyAttr(Sheet sh, XMLDoc xdoc)
        {
            XElement elem = null;
            if (sh.Sketches[1].Name == "Технические требования")
            {
                AttributeSet attset = sh.Sketches[1].AttributeSets[1];
                xdoc.Doc.Root.Add(elem = new XElement(attset.Name));
                foreach (Inventor.Attribute item in attset)
                {
                    elem.Add(new XElement(item.Name, item.Value));
                }
                xdoc.Doc.Root.Add(elem = new XElement("Data"));
                foreach (Inventor.TextBox item in sh.Sketches[1].TextBoxes)
                {
                    elem.Add(new XElement("Text", new XAttribute("OriginX", item.Origin.X), new XAttribute("OriginY", item.Origin.Y), item.FormattedText));
                }
            }
            xdoc.save();
        }

        private void addTRToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            InvDoc.InvDocument<DrawingDocument> doc = new InvDocument<DrawingDocument>((DrawingDocument)Macros.StandardAddInServer.m_inventorApplication.ActiveDocument);
            Sheet sh = doc.GetDoc.Sheets[1];
            string name = ut.OFD(doc.path, "XML files(*.xml)|*.xml");
            XMLDoc xdoc = new XMLDoc(name, "TR");
            addTR(sh, xdoc);
            this.Close();
        }

        public static void addTR(Sheet sh, XMLDoc xdoc)
        {
            if (sh.Sketches.Count == 0)
            {
                DrawingSketch ds = sh.Sketches.Add();
                ds.Name = "Технические требования";
                Macros.StandardAddInServer.m_inventorApplication.SilentOperation = !Macros.StandardAddInServer.m_inventorApplication.SilentOperation;
                Macros.StandardAddInServer.m_inventorApplication.ScreenUpdating = !Macros.StandardAddInServer.m_inventorApplication.ScreenUpdating;
                ds.Edit();
                DrawingStylesManager mgr = ((DrawingDocument)sh.Parent).StylesManager;
                DrawingStandardStyle oldStl = mgr.ActiveStandardStyle;
                mgr.ActiveStandardStyle = mgr.StandardStyles["ГОСТ"]; 
                foreach (var item in xdoc.Doc.Root.Element("Data").Elements())
                {
                    Inventor.TextBox tb = ds.TextBoxes.AddFitted(I.tg.CreatePoint2d(Double.Parse(item.Attribute("OriginX").Value.Replace('.', ',')), Double.Parse(item.Attribute("OriginY").Value.Replace('.', ','))), item.Value);
                    tb.SingleLineText = true;
                    tb.VerticalJustification = VerticalTextAlignmentEnum.kAlignTextBaseline;
                    if (tb.Style.Bold)
                    tb.Style.Bold = false;
                }
                mgr.ActiveStandardStyle = oldStl;
                ds.ExitEdit();
                AttributeSet attSet = ds.AttributeSets.Add("com_autodesk_MSD_AIS_Gost");
                int i = 1;
                foreach (var item in xdoc.Doc.Root.Element("com_autodesk_MSD_AIS_Gost").Elements())
                {
                    if (i == 2)
                    {
                        attSet.Add(item.Name.ToString(), ValueTypeEnum.kIntegerType, item.Value);
                    }
                    else if (i == 7 || i == 8)
                    {
                        attSet.Add(item.Name.ToString(), ValueTypeEnum.kDoubleType, item.Value.Replace('.', ','));
                    }
                    else
                    {
                        attSet.Add(item.Name.ToString(), ValueTypeEnum.kStringType, item.Value);
                    }
                    i++;
                }
                Macros.StandardAddInServer.m_inventorApplication.SilentOperation = !Macros.StandardAddInServer.m_inventorApplication.SilentOperation;
                Macros.StandardAddInServer.m_inventorApplication.ScreenUpdating = !Macros.StandardAddInServer.m_inventorApplication.ScreenUpdating;
            }
        }

        private void добавитьСвойствоToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            Form f = new Form();
            InterfaceDll.MyLabel mylbl1 = new MyLabel(); int offsetY = 20;
            InterfaceDll.MyTextBox mytxt1 = new MyTextBox();
            InterfaceDll.MyComboBox mycb = new MyComboBox();
            f.Height = 100; f.Width = 400; f.WindowState = FormWindowState.Normal; f.Text = "Добавить свойство"; f.StartPosition = FormStartPosition.CenterScreen;
            System.Drawing.Point insPt = new System.Drawing.Point(5, 5);
            Label lbl1 = mylbl1.addLabel("Имя свойства", insPt, 100, 15);
            f.Controls.Add(lbl1);
            insPt.Y += offsetY;
            lbl1 = mylbl1.addLabel("Значение", insPt, 100, 15);
            f.Controls.Add(lbl1);
            insPt.X += 150; insPt.Y -= offsetY;
            System.Windows.Forms.ComboBox cb = mycb.addComboBox("", insPt, 200, 15, new string[] { "Изв", "ИзвД", "CountDXF", "thick", "perf"});   
            insPt.Y += offsetY;
            System.Windows.Forms.TextBox txt2 = mytxt1.addTextBox("", insPt, 200, 15);
            f.Controls.Add(cb); f.Controls.Add(txt2);
            MyButton myBtn = new MyButton();
            insPt.Y += offsetY; insPt.X = 200 - 50;
            System.Windows.Forms.Button btn = myBtn.addButton("Добавить", insPt, 100, 20);
            f.Controls.Add(btn);
            btn.Click += btn_Click;
            f.Show();
        }

        public static void pack(Document doc)
        {
            PackAndGoLib.PackAndGoComponent packAndGoComp = new PackAndGoLib.PackAndGoComponent();
            string locName;
            //Document doc = Macros.StandardAddInServer.m_inventorApplication.ActiveDocument;
            locName = doc.path() + "\\";
            locName += doc.name() + "\\";
            if (!System.IO.Directory.Exists(locName)) System.IO.Directory.CreateDirectory(locName);
            PackAndGoLib.PackAndGo packAndGo = packAndGoComp.CreatePackAndGo(Macros.StandardAddInServer.m_inventorApplication.ActiveDocument.FullDocumentName, locName);

            string[] refFiles = new string[] { };
            string[] refecening = new string[] { };
            object refMissFiles = new object();

            // Set the options
            packAndGo.SkipLibraries = false;
            packAndGo.SkipStyles = true;
            packAndGo.SkipTemplates = true;
            packAndGo.CollectWorkgroups = false;
            packAndGo.KeepFolderHierarchy = false;
            packAndGo.IncludeLinkedFiles = false;

            HashSet<string> names = new HashSet<string>();
            if (doc.DocumentType == DocumentTypeEnum.kAssemblyDocumentObject)
            {
                AssemblyDocument asmDoc = doc as AssemblyDocument;
                BOM bom = asmDoc.ComponentDefinition.BOM;
                names.Add(doc.FullFileName);
                packFromBOM(bom.BOMViews[1].BOMRows, ref names);
            }

            if (names.Count != 0)
                packAndGo.AddFilesToPackage(names.ToArray());
            // Start the pack and go to create the package
            packAndGo.CreatePackage(true);
        }

        private void комплектФайловToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            Document doc = Macros.StandardAddInServer.m_inventorApplication.ActiveDocument;
            pack(doc);
        }

        public static void packFromBOM(BOMRowsEnumerator rows, ref HashSet<string> names)
        {
            foreach (BOMRow item in rows)
            {
                if (item.ChildRows != null) packFromBOM(item.ChildRows, ref names);
                if (item.ComponentDefinitions[1] is VirtualComponentDefinition) continue;
                names.Add(item.ReferencedFileDescriptor.FullFileName);
                if (item.ReferencedFileDescriptor.ReferencedFile.ReferencedFiles.Count != 0)
                    foreach (File f in item.ReferencedFileDescriptor.ReferencedFile.ReferencedFiles)
                    {
                        names.Add(f.FullFileName);
                    }
            }
        }

        private void iPartToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            string iPartName = ut.OFD(Macros.StandardAddInServer.m_inventorApplication.DesignProjectManager.ActiveDesignProject.WorkspacePath);
            if (iPartName != "")
            {
                PartDocument doc = (PartDocument)Macros.StandardAddInServer.m_inventorApplication.Documents.Open(iPartName, false);
                iPartFactory fact = doc.ComponentDefinition.iPartFactory;
                Matrix mtx = I.tg.CreateMatrix();
                if (Macros.StandardAddInServer.m_inventorApplication.ActiveDocument.DocumentType == DocumentTypeEnum.kAssemblyDocumentObject)
                {
                    AssemblyDocument asm = (AssemblyDocument)Macros.StandardAddInServer.m_inventorApplication.ActiveDocument;
                    ComponentOccurrence occ = null;
                    Box rbox = asm.ComponentDefinition.RangeBox;
                    //                 mtx.SetCoordinateSystem(rbox.MaxPoint, I.tg.CreateVector(1, 0, 0), I.tg.CreateVector(0, 1, 0),
                    //                     I.tg.CreateVector(0, 0, 1));
                    mtx.SetTranslation(rbox.MinPoint.VectorTo(rbox.MaxPoint));
                    for (int i = 1; i < fact.TableRows.Count + 1; i++)
                    {
                        occ = asm.ComponentDefinition.Occurrences.AddiPartMember(iPartName, mtx, i);
                        mtx.Cell[1, 4] = mtx.Cell[1, 4] + occ.RangeBox.MaxPoint.DistanceTo(occ.RangeBox.MinPoint);
                    }
                }
                else if (Macros.StandardAddInServer.m_inventorApplication.ActiveDocument.DocumentType == DocumentTypeEnum.kDrawingDocumentObject)
                {
                    Matrix2d mtx2d = I.tg.CreateMatrix2d();
                    NameValueMap nvm = I.objs.CreateNameValueMap();
                    DrawingDocument drw = (DrawingDocument)Macros.StandardAddInServer.m_inventorApplication.ActiveDocument;
                    Point2d pt = I.tg.CreatePoint2d();
                    mtx2d.Cell[1, 3] = 12; mtx2d.Cell[2, 3] = 26;
                    pt.TransformBy(mtx2d);
                    double offsetY = 1;
                    for (int i = 1; i < fact.TableRows.Count + 1; i++)
                    {
                        nvm.Clear();
                        nvm.Add("MemberName", fact.TableRows[i].MemberName);
                        DrawingView dv = drw.ActiveSheet.DrawingViews.AddBaseView((_Document)doc, pt, 1, ViewOrientationTypeEnum.kDefaultViewOrientation, DrawingViewStyleEnum.kHiddenLineRemovedDrawingViewStyle, AdditionalOptions: nvm);
                        //mtx2d.Cell[1,3] = mtx2d.Cell[1,3] + dv.Position.X + offsetX + dv.Width/2;
                        mtx2d.Cell[2, 3] = dv.Position.Y - offsetY - dv.Height;
                        pt.X = 0; pt.Y = 0;
                        pt.TransformBy(mtx2d);
                    }
                }
            }
        }

        private void stepToolStripMenuItem_Click(object sender, EventArgs e)
        {
            TranslatorAddIn step = (TranslatorAddIn)Macros.StandardAddInServer.m_inventorApplication.ApplicationAddIns.ItemById["{90AF7F40-0C01-11D5-8E83-0010B541CD80}"];
            TranslationContext context = I.objs.CreateTranslationContext();
            NameValueMap nvm = I.objs.CreateNameValueMap();
            nvm.Add("ApplicationProtocolType", 3);
            nvm.Add("IncludeSketches", false);
            context.Type = IOMechanismEnum.kFileBrowseIOMechanism;
            DataMedium dm = I.objs.CreateDataMedium();
            Document doc = Macros.StandardAddInServer.m_inventorApplication.ActiveDocument;
            if (doc.DocumentType == DocumentTypeEnum.kAssemblyDocumentObject)
            {
                AssemblyDocument asm = doc as AssemblyDocument;
                asm.ComponentDefinition.RepresentationsManager.DesignViewRepresentations[1].Activate();
                asm.ObjectVisibility.AllWorkFeatures = false;
                asm.ObjectVisibility.Sketches = false;
                asm.ObjectVisibility.Sketches3D = false;
            }
            string name = doc.path() + "\\" + doc.name() + ".stp";
            dm.FileName = name;
            step.SaveCopyAs(doc, context, nvm, dm);

            System.Diagnostics.ProcessStartInfo start = new System.Diagnostics.ProcessStartInfo();
            start.FileName = @"C:\Program Files (x86)\Adobe\Acrobat 10.0\Acrobat\Acrobat.exe";
            start.Arguments = name;
            start.WindowStyle = System.Diagnostics.ProcessWindowStyle.Hidden;
            var process = System.Diagnostics.Process.Start(start);
            process.WaitForExit();
            System.IO.File.Delete(name);
            this.Close();
        }

        private void сортироватьСборкуToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Inventor.Application app = Macros.StandardAddInServer.m_inventorApplication;
            if (app.ActiveDocument.DocumentType == DocumentTypeEnum.kAssemblyDocumentObject)
            {
                AssemblyDocument asm = app.ActiveDocument as AssemblyDocument;
                BOM bom = asm.ComponentDefinition.BOM;
                if (!bom.PartsOnlyViewEnabled) bom.PartsOnlyViewEnabled = true;
                BOMView bView = bom.BOMViews["Только детали"];
                sortBOM(bView, asm);
            }
        }

        public static void sortBOM(BOMView bView, AssemblyDocument asm)
        {
            TableInv tbl = null;
            tbl = tbl ?? new TableInv(asm, @"C:\ProgramData\Autodesk\Inventor Addins\Sequence.xml");
            tbl.addTable(bView);
            List<int> except = new List<int>();
            tbl.exc(ref except);
            tbl.rows.Sort((e1, e2) => e1.CompareTo(e2));
            tbl.reNumber(tbl.rows);
            tbl.renumberBom(tbl.rows, bView);
        }

        public static void replace(AssemblyDocument asm_Doc, XMLDoc xdoc)
        {
            AssemblyComponentDefinition acd = asm_Doc.ComponentDefinition;
            
            //int max = 10;
            String name = "";
            string newName = "";
            foreach (var item in xdoc.Doc.Root.Elements())
            {
                name = item.Attribute("find").Value;
                newName = item.Value;
                foreach (ComponentOccurrence occ in acd.Occurrences)
                {
                    Document doc = occ.ReferencedDocumentDescriptor.ReferencedDocument as Document;
                    if (doc != null && doc.DocumentType == DocumentTypeEnum.kAssemblyDocumentObject)
                    {
                        replace(doc as AssemblyDocument, xdoc);
                    }

                    if (occ.ReferencedDocumentDescriptor.FullDocumentName.ToString().IndexOf(name) != -1)
                    {
                        try
                        {
                        string member = ContentOp.memberForPlace(newName);
                        occ.Replace(member, true);
                        break;
                        }
                        catch{}
                    }
                }
            }
        }

        private void заменаКрепежаToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Inventor.Application app = Macros.StandardAddInServer.m_inventorApplication;
            if (app.ActiveDocument.DocumentType == DocumentTypeEnum.kAssemblyDocumentObject)
            {
                AssemblyDocument asm_Doc = app.ActiveDocument as AssemblyDocument;
                XMLDoc xdoc = new XMLDoc(asm_Doc.path() + "\\Replace.xml", "Replace");
                replace(asm_Doc, xdoc);
            }
        }

        private void перевестиВЛитеруToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Inventor.Application app = Macros.StandardAddInServer.m_inventorApplication;
            if (app.ActiveDocument.DocumentType == DocumentTypeEnum.kAssemblyDocumentObject)
            {
                List<string> lit = new List<string> { "М", "О", "О1", "О2", "А" };
                AssemblyDocument asm = app.ActiveDocument as AssemblyDocument;
                BOM bom = asm.ComponentDefinition.BOM;
                BOMView bView = bom.BOMViews[1];
                string type = (ut.getProp((Document)asm, "Type")).Value.ToString();
                DialogResult dr = MessageBox.Show("Перевести в следующую литеру?", "Перевод в следующую литеру", MessageBoxButtons.YesNo);
                if (dr == System.Windows.Forms.DialogResult.Yes)
                {
                    string f = litera_((Document)asm, ref lit, type, "");
                    litera(bView.BOMRows, ref lit, type, f);
                }
            }
        }

        public static void litera(BOMRowsEnumerator rows, ref List<string> lit, string typ, string l)
        {
            foreach (BOMRow row in rows)
            {
                if (row.ChildRows != null) litera(row.ChildRows, ref lit, typ, l);
                Document doc = row.ComponentDefinitions[1].Document as Document;
                litera_(doc, ref lit, typ, l);
            }
        }

        public static string litera_(Document doc, ref List<string> lit, string typ, string l)
        {
            Property type, lit1, lit2, date1, date2;
            type = ut.getProp(doc, "Type");
            if (type != null && typ != type.Value.ToString()) return "";
            lit1 = ut.getProp(doc, "Литера1");
            lit2 = ut.getProp(doc, "Литера2");
            if (lit1 != null && lit2 != null)
            {
                string f = lit1.Value.ToString() + lit2.Value.ToString();
                if (l != "") f = l;
                int ind = lit.FindIndex(e => e == f);
                if (ind != -1 && ind < lit.Count - 1)
                {
                    f = lit[ind + 1];
                    if (f.Count() == 1) { lit1.Value = f; lit2.Value = ""; }
                    else
                    {
                        lit1.Value = f[0];
                        lit2.Value = f[1];
                    }
                    date1 = ut.getProp(doc, "Creation Time");
                    date1.Value = System.DateTime.Now.ToString("dd.MM.yyyy");
                    date2 = ut.getProp(doc, "Engr Date Approved");
                    date2.Value = System.DateTime.Now.ToString("dd.MM.yyyy");
                }
                return f;
            }
            return "";
        }

        private void визуальнаяПробивкаToolStripMenuItem_Click(object sender, EventArgs e)
        {
            PartDocument doc = Macros.StandardAddInServer.m_inventorApplication.ActiveDocument as PartDocument;
            PartComponentDefinition compDef = doc.ComponentDefinition;
            TransientBRep brep = Macros.StandardAddInServer.m_inventorApplication.TransientBRep;
            SurfaceBody baseBody = compDef.SurfaceBodies[1], toolBody = compDef.SurfaceBodies[2], 
                copyBaseBody = brep.Copy(baseBody), copyToolBody = brep.Copy(toolBody);
            Matrix mtx = I.tg.CreateMatrix();
            Point lastPosition = I.tg.CreatePoint();
            //Cylinder cyl = (Cylinder)copyToolBody.Faces.OfType<Face>().FirstOrDefault(f => f.SurfaceType == SurfaceTypeEnum.kCylinderSurface);
            //cyl.
            mtx.Cell[1, 4] = compDef.WorkPoints[2].Point.X;
            mtx.Cell[2, 4] = compDef.WorkPoints[2].Point.Y;
            //mtx.Cell[3, 4] = compDef.WorkPoints[2].Point.Z;
            brep.Transform(copyToolBody, mtx);
            //mtx = I.tg.CreateMatrix();
            bool first = true;

            //for (int i = 0; i < 80; i++)
            //{
            //    mtx.Cell[1, 4] = compDef.WorkPoints[1].Point.X + 2.5 * i;
            //    for (int j = 0; j < 80; j++)
            //    {
            //        mtx.Cell[2, 4] = compDef.WorkPoints[1].Point.Y - 2.5*j;

            //    } 
            //}
            foreach (WorkPoint wp in compDef.WorkPoints)
            {
                if (first) { first = false; continue; }
                mtx.Cell[1, 4] = wp.Point.X - lastPosition.X;
                mtx.Cell[2, 4] = wp.Point.Y - lastPosition.Y;
                brep.Transform(copyToolBody, mtx);
                brep.DoBoolean(copyBaseBody, copyToolBody, BooleanTypeEnum.kBooleanTypeDifference);
                lastPosition = wp.Point;
            }
            NonParametricBaseFeatureDefinition npbf = compDef.Features.NonParametricBaseFeatures.CreateDefinition();
            ObjectCollection col = I.objs.CreateObjectCollection();
            col.Add(copyBaseBody);
            npbf.BRepEntities = col;
            npbf.OutputType = BaseFeatureOutputTypeEnum.kSolidOutputType;
            compDef.Features.NonParametricBaseFeatures.AddByDefinition(npbf);
            baseBody.Visible = false;
            toolBody.Visible = false;
            Macros.StandardAddInServer.m_inventorApplication.ActiveView.Update();
        }

        private void добавитьПараметрыToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Document doc = Macros.StandardAddInServer.m_inventorApplication.ActiveDocument;
            string name = ut.OFD(doc.path(),"XML(*.xml)|*.xml", false);
            if (name == null) return;
            XMLDoc xdoc = new XMLDoc(name, "head");
            foreach (var el in xdoc.El.Elements("Parameter"))
            {
                ut.addParameter(doc, el);
            }
            foreach (var el in xdoc.El.Elements("Plane"))
            {
                string Name = XMLDoc.getXAttributeValue(el, "Name"), BasePlane = XMLDoc.getXAttributeValue(el, "BasePlane"), Offset = XMLDoc.getXAttributeValue(el, "Offset"), rev = XMLDoc.getXAttributeValue(el,"reverse");
                if (InvDoc.Reflect.exist<WorkPlane>((doc as PartDocument).ComponentDefinition.WorkPlanes, Name)) continue;
                ut.addPlane(((PartDocument)doc).ComponentDefinition, Name, BasePlane, Offset, rev);
            }
            foreach (var el in xdoc.El.Elements("Sketch"))
            {
                string Name = XMLDoc.getXAttributeValue(el, "Name"), BasePlane = XMLDoc.getXAttributeValue(el, "BasePlane");
                if (InvDoc.Reflect.exist<PlanarSketch>((doc as PartDocument).ComponentDefinition.Sketches, Name)) continue;
                ut.addSketch(((PartDocument)doc).ComponentDefinition, Name, BasePlane);
            }
        }

        private void шипToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            spikes();
        }

        private void пазToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            cuts();
        }

        private void выдавитьToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            extrude();
        }

        private void добавитьВСборкуToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            addToAsm();
        }

        private void вставитьПараметрическийЭлементToolStripMenuItem_Click(object sender, EventArgs e)
        {
            insertParam();
        }

        private void центрыОтверстийToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (Macros.StandardAddInServer.m_inventorApplication.ActiveDocument.DocumentType == DocumentTypeEnum.kDrawingDocumentObject)
            { Drawings drws = new Drawings((DrawingDocument)Macros.StandardAddInServer.m_inventorApplication.ActiveDocument, true); }
        }

        private void сплайнToolStripMenuItem_Click(object sender, EventArgs e)
        {
            PartDocument doc = Macros.StandardAddInServer.m_inventorApplication.ActiveDocument as PartDocument;
            string name = ut.OFD(doc.path(), "xml|*.xml", false);
            XMLDoc xdoc = new XMLDoc(name, "head"); double scale = 1;
            ObjectCollection col = I.objs.CreateObjectCollection();
            WorkPlane wp = null;

            if (xdoc != null)
            {
                foreach (var surf in xdoc.El.Elements("surface"))
                {
                    doc = Macros.StandardAddInServer.m_inventorApplication.Documents.Add(DocumentTypeEnum.kPartDocumentObject, CreateVisible: false) as PartDocument;
                    PartComponentDefinition compDef = doc.ComponentDefinition;
                    double x = ut.convToDouble(surf.Attribute("GridX").Value);
                    double y = ut.convToDouble(surf.Attribute("GridY").Value);
                    double startY = 0; double startX = 0;
                    foreach (var row in surf.Elements())
                    {
                        if (surf.Attribute("Scale") != null)
                            scale = ut.convToDouble(surf.Attribute("Scale").Value);
                        if (startY == 0) wp = compDef.WorkPlanes[1];
                        else
                        {
                            wp = compDef.WorkPlanes.AddByPlaneAndOffset(wp, y);
                            wp.Visible = false;
                        }
                        col.Add(addSpline(compDef, wp, startX, x, row.Value, scale)); 
                        //i++;
                        startY += y;
                    }
                    LoftDefinition loftDef = compDef.Features.LoftFeatures.CreateLoftDefinition(col, PartFeatureOperationEnum.kSurfaceOperation);
                    LoftFeature loft = compDef.Features.LoftFeatures.Add(loftDef);
                    if (surf.Attribute("Name") != null)
                    loft.Name = surf.Attribute("Name").Value;
                    col.Clear();
                    doc.SaveAs(name.Replace(".xml",".ipt"), false);
                    doc.Close();
                }
            }

            //List<Point2d> pts = new List<Point2d>();
            //Point2d pt = I.tg.CreatePoint2d();
            //ObjectCollection col = I.objs.CreateObjectCollection();
            //pts.Add(pt); col.Add(pt);
            //Point2d pt2 = pt.Copy(); pt2.X = 10; pt2.Y = 1; pts.Add(pt2);
            //pt2 = pt.Copy(); pt2.X = 20; pt2.Y = 5; pts.Add(pt2); col.Add(pt2);
            //pt2 = pt.Copy(); pt2.X = 30; pt2.Y = 0; pts.Add(pt2); col.Add(pt2);
            //pt2 = pt.Copy(); pt2.X = 50; pt2.Y = -40; pts.Add(pt2); col.Add(pt2);
            //PartDocument doc = Macros.StandardAddInServer.m_inventorApplication.ActiveDocument as PartDocument;
            //PlanarSketch sketch = doc.ComponentDefinition.Sketches.Add(doc.ComponentDefinition.WorkPlanes[1]);
            ////SketchSpline spl = sketch.SketchSplines.Add(col, SplineFitMethodEnum.kSweetSplineFit);
            //SketchControlPointSpline scpSpl = sketch.SketchControlPointSplines.Add(col);
        }

        public static Profile addSpline(PartComponentDefinition compDef ,WorkPlane wp, double start, double step, string vals, double scale)
        {
            ObjectCollection col = I.objs.CreateObjectCollection();
            string[] arr = vals.Split(';');
            foreach (var item in arr)
            {
                Point2d pt = ut.addPt2d(start, item);
                col.Add(pt);
                start += step; 
            }
            PlanarSketch sk = compDef.Sketches.Add(wp);
            //SketchControlPointSpline scpSpl = sk.SketchControlPointSplines.Add(col);
            SketchSpline spl = sk.SketchSplines.Add(col, SplineFitMethodEnum.kSweetSplineFit);
            return sk.Profiles.AddForSurface();
        }

        private void создатьСборкуToolStripMenuItem_Click(object sender, EventArgs e)
        {
            string names = ut.OFD(Macros.StandardAddInServer.m_inventorApplication.ActiveDocument.path(), multi: true);

            if (names == null) return;
            AssemblyDocument doc = Macros.StandardAddInServer.m_inventorApplication.Documents.Add(DocumentTypeEnum.kAssemblyDocumentObject,CreateVisible: true) as AssemblyDocument;
            foreach (var name in names.Split('|'))
            {
                addToAsm(doc.ComponentDefinition, name);
                //if (nameforSave == "") nameforSave = name.Replace(".ipt", ".iam");
            }
            doc.ComponentDefinition.RepresentationsManager.DesignViewRepresentations.Add("1");
            doc.ComponentDefinition.RepresentationsManager.DesignViewRepresentations["1"].Activate();
            doc.ComponentDefinition.WorkPlanes[3].Visible = true;
            //doc.SaveAs(nameforSave, false);
            this.Close();
        }

        static public string nextName(string name)
        {
            string pat = "_(\\d+)$";
            int d = 1;
            Regex regex = new Regex(pat);
            foreach (Match item in regex.Matches(name))
            {
                d += int.Parse(item.Groups[1].Value);
                return regex.Replace(name, "_" + d.ToString());  
            }
            return name + "_" + d;
        }

        private RectangularPatternFeature pattern(PartComponentDefinition compDef, string count, string dist, ObjectCollection col, object dir, string name, string iMate,
            string offset, string fastener, string offsetFastener, bool midPlane = false, string count2 = null, string dist2 = null, int numEdge = 1, bool imCheck = false)
        {
            bool rev = false; int c = 0;
            if (name.ToLower().EndsWith("отв")) rev = true;
            if (iMate == null && imCheck)
            {
                iMate = (name.ToLower().EndsWith("отв"))?name.Remove(name.Length-3):name;
            }
            name = name + "Массив";
            while (InvDoc.Reflect.exist<RectangularPatternFeature>(compDef.Features.RectangularPatternFeatures, name))
            {
                name = nextName(name);
            }
            RectangularPatternFeature p = null;
            object hf = col[1]; Cylinder cyl = null; 
            if (hf is HoleFeature) cyl = ((HoleFeature)hf).SideFaces[1].Geometry as Cylinder;
            else if (hf is MirrorFeature) cyl = ((MirrorFeature)hf).Faces[1].Geometry as Cylinder;
            
            if (count2 == null)
                {p = compDef.Features.RectangularPatternFeatures.Add(col, dir, true, count, dist, ComputeType: PatternComputeTypeEnum.kAdjustToModelCompute); c = int.Parse(p.XCount.Value.ToString()) - 1;}
            else if (cyl != null)
            {
                UnitVector vec = cyl.AxisVector, vec2 = null;
                if (dir is Edge)
                {
                    Edge ed = dir as Edge;
                    vec2 = ed.StartVertex.Point.VectorTo(ed.StopVertex.Point).AsUnitVector();
                }
                else if (dir is WorkAxis)
                {
                    WorkAxis wa = dir as WorkAxis;                       
                    vec2 = wa.Line.Direction;     
                }
                object dir2 = ut.axisFromEdge(compDef, vec.CrossProduct(vec2));
                p = compDef.Features.RectangularPatternFeatures.Add(col, dir, true, count, dist, YDirectionEntity: dir2, YCount: count2, YSpacing: dist2, 
                    ComputeType: PatternComputeTypeEnum.kAdjustToModelCompute);
                c = (int.Parse(p.XCount.Value.ToString())*int.Parse(p.YCount.Value.ToString()))-1;
            }
            p.Name = name;
            if (midPlane) { p.XDirectionMidPlanePattern = true; }
            if (midPlane && count2 != null) p.YDirectionMidPlanePattern = true;
            ((PartDocument)compDef.Document).Update();
            if (p.Faces.Count != c) p.NaturalXDirection = !p.NaturalXDirection;
            if (p.Faces.Count != c && count2 != null) p.NaturalYDirection = !p.NaturalYDirection;
            if (p.Faces.Count != c) p.NaturalXDirection = !p.NaturalXDirection;
            IMate im = new IMate(); double off = 0;
            ObjectCollection objs = I.objs.CreateObjectCollection();

            if (iMate != null && hf is HoleFeature)
            {
                if (offset != null) off = ut.convToDouble(offset);
                InsertiMateDefinition insIMateDef1 = im.iMate_((hf as HoleFeature).Faces[1].Edges[numEdge], compDef, off);
                int numPat = (p.YCount == null)?2:int.Parse(p.XCount.Value.ToString())+2;
                if (rev && p.PatternElements.Count > numPat + 1) numPat++;
                InsertiMateDefinition insIMateDef2 = im.iMate_(p.PatternElements[numPat].Faces[1].Edges[numEdge], compDef, off);
                swap(ref objs, rev, insIMateDef1, insIMateDef2);
                //if (comboBox3.Text != "")
                CompositeiMateDefinition compos;
                im.addName(compos = compDef.iMateDefinitions.AddCompositeiMateDefinition(objs), iMate);
                objs.Clear();
            }
            if (fastener != null && hf is HoleFeature)
            {
                int num = (numEdge == 2) ? 1: 2;
                if (offsetFastener != null) off = ut.convToDouble(offsetFastener);
                InsertiMateDefinition insIMateDef = im.iMate_((hf as HoleFeature).Faces[1].Edges[num], compDef, off);
                im.addName(insIMateDef, fastener);
//                 for (int i = 2; i <= p.PatternElements.Count; i++)
//                 {
//                     insIMateDef = im.iMate_(p.PatternElements[i].Faces[1].Edges[2], compDef, off);
//                     im.addName(insIMateDef, fastener);
//                 }
            }
            return p;
        }

        private void swap(ref ObjectCollection col, bool rev,InsertiMateDefinition im1, InsertiMateDefinition im2)
        {
            if (rev)
            {
                col.Add(im2); col.Add(im1);
            }
            else
            {
                col.Add(im1); col.Add(im2);
            }
        }

        public static HoleFeature hole(PartComponentDefinition compDef, string diam, PlanarSketch sk)
        {
            ObjectCollection col = I.objs.CreateObjectCollection();
            foreach (SketchPoint pt in sk.SketchPoints)
	        {
		       if (pt.HoleCenter == true) col.Add(pt);
	        }
            SketchHolePlacementDefinition def = compDef.Features.HoleFeatures.CreateSketchPlacementDefinition(col);
            HoleFeature hf;
            hf = compDef.Features.HoleFeatures.AddDrilledByDistanceExtent(def, diam, compDef.Parameters["Толщина"].Name, PartFeatureExtentDirectionEnum.kPositiveExtentDirection);
            
            ((Document)compDef.Document).Update();
            if (hf.HealthStatus == HealthStatusEnum.kDriverLostHealth) ((DistanceExtent)hf.Extent).Direction = PartFeatureExtentDirectionEnum.kNegativeExtentDirection;
//             HoleFeature hfname = compDef.Features.OfType<HoleFeature>().LastOrDefault(el => el.Name.IndexOf(diam) != -1);
//             if (hfname != null)
//             {
//                 diam = nextName(hfname.Name);
//             }
            string name = diam;
            if (diam.ToLower().EndsWith("отв")) name += "Отв";
            
            while (InvDoc.Reflect.exist<HoleFeature>(compDef.Features.HoleFeatures, name))
            {
                name = nextName(name);
            }
            hf.Name = name;
            return hf;
        }

        private void hole(object ob, string diam, PartComponentDefinition compDef)
        {
            HoleFeature hf = ob as HoleFeature;
            if (hf.HealthStatus == HealthStatusEnum.kDriverLostHealth) ((DistanceExtent)hf.Extent).Direction = PartFeatureExtentDirectionEnum.kNegativeExtentDirection;
            hf.HoleDiameter.Expression = diam;
//             HoleFeature hfname = compDef.Features.OfType<HoleFeature>().LastOrDefault();
//             if (hfname != null)
//             {
//                 diam = nextName(hfname.Name);
//             }
            //hf.Name = diam;
        }

        private void pattern(object ob, string count, string dist, string diam, string name, PartComponentDefinition compDef)
        {
            RectangularPatternFeature rpf = ob as RectangularPatternFeature;
            rpf.Name = name;
            rpf.Parameters[1].Expression = dist;
            rpf.Parameters[2].Expression = count;
            object hf = rpf.ParentFeatures[1];
            if (hf is HoleFeature) hole(hf, diam, compDef);
        }

//         private void formaForSheet()
//         {
//             InterfaceDll.MyForm f = new MyForm("Создание листов чертежей");
//             XMLDoc xDoc = new XMLDoc(@"C:\ProgramData\Autodesk\Inventor Addins\sheet.xml", "head");
//             int offsetX = 50, offsetY = 30;
//             f.form.AutoSize = true;
//             f.offsetX = 0; f.offsetY = offsetY; f.w = 100; f.h = 10;
//             f.addLabel("Выберите базовый вид", 10, 10);
//             f.setPosition(f.myLbl.lbls[0], offsetX, 0);
//             f.addComboBox("1",f.myLbl.lbls[0],offsetX, 0);
//             f.myCb.fill(xDoc.El, "view", "name");
//             f.addLabel("Масштаб"); f.addLabel("Формат первого листа"); f.addLabel("Формат второго листа");
//             f.addComboBox("2"); f.myCb.fill(xDoc.El.Element("Scales"), "value");
//             f.addComboBox("3"); f.myCb.fill(xDoc.El.Element("Format"), "value");
//             f.addComboBox("4"); f.myCb.fill(xDoc.El.Element("Format"), "value");
//             f.addLabel("Горизонтальный вид смещение");
//             f.addLabel("Вертикальный вид смещение");
//             f.addComboBox("5");
//             f.addComboBox("6");
//             f.addCheck("Убрать выравнивание", f.myCb.cbs[f.myCb.cbs.Count - 2], 10, 0);
//             f.addCheck("Убрать выравнивание");
// 
//             f.w = 100; f.h = 25;
//             f.addButton("Создать", f.myCb.cbs[f.myCb.cbs.Count - 1], 0, 40);
//             f.getValuesComboBox();
//             List<string> values = new List<string>();
//         }

        void func(List<string> param)
        {
            drawing(Macros.StandardAddInServer.m_inventorApplication.ActiveDocument, param, true);
            this.Close();
        }

        private void formaForPattern()
        {
            Form f = new Form();
            InterfaceDll.MyLabel mylbl1 = new MyLabel(); int offsetY = 30, offsetX = 10;
            InterfaceDll.MyComboBox mycb = new MyComboBox();
            InterfaceDll.MyCheckBox mychk = new MyCheckBox();
            f.Height = 150+50; f.Width = 400; f.WindowState = FormWindowState.Normal; f.Text = "Массив"; f.StartPosition = FormStartPosition.CenterScreen;
            System.Drawing.Point insPt = new System.Drawing.Point(5, 5);
            CheckBox chk = mychk.addCheckBox("+ зеркальное оторажение", insPt, 100, 15);
            f.Controls.Add(chk);
            chk = mychk.addCheckBox("Перпендикулярно главному", InterfaceDll.MyLabel.position(chk, offsetX, 0), 100, 15);
            f.Controls.Add(chk);
            insPt.Y += offsetY;
            Label lbl1 = mylbl1.addLabel("Выбрать массив", insPt, 100, 15);
            f.Controls.Add(lbl1);
            System.Windows.Forms.ComboBox cb = mycb.addComboBox("Cb", InterfaceDll.MyLabel.position(lbl1, offsetX, 0) /*new System.Drawing.Point(lbl1.Right + offsetY,insPt.Y)*/, 200, 15);
            f.Controls.Add(cb);
            insPt.Y += offsetY;
            Label lbl2 = mylbl1.addLabel("Тип элементов", InterfaceDll.MyLabel.position(lbl1, -lbl1.Width, offsetY), 100, 15);
            f.Controls.Add(lbl2);
            System.Windows.Forms.ComboBox cb1 = mycb.addComboBox("Cb1", InterfaceDll.MyLabel.position(cb, -cb.Width, offsetY) /*new System.Drawing.Point(lbl1.Right + offsetY,insPt.Y)*/, 200, 15);
            f.Controls.Add(cb1);
            insPt.Y += offsetY;
            chk = mychk.addCheckBox("Сменить сторону", insPt, 100, 15);
            f.Controls.Add(chk);
            chk = mychk.addCheckBox("Конструктивная пара", InterfaceDll.MyLabel.position(chk, offsetX, 0), 100, 15);
            f.Controls.Add(chk);
            MyButton myBtn = new MyButton();
            insPt.Y += offsetY; insPt.X = 200 - 50;
            System.Windows.Forms.Button btn = myBtn.addButton("Добавить", insPt, 100, 20);
            f.Controls.Add(btn);
            HashSet<string> set = new HashSet<string>();
            foreach (var item in PartsBtn.xDoc.El.Elements("Pattern"))
            {
                if (item.Attribute("Name") != null) set.Add(item.Attribute("Name").Value); 
            }
            cb.Items.AddRange(set.ToArray());
            cb1.Items.AddRange(new string[] { "Отверстие", "Овал", "Другой элемент"});
            cb1.Text = "Добавить";
            btn.Click += forma_Click;
            f.Show();
        }

        void forma_Click(object sender, EventArgs e)
        {
            PartDocument doc = Macros.StandardAddInServer.m_inventorApplication.ActiveDocument as PartDocument;
            Form f = (Form)((System.Windows.Forms.Button)sender).Parent;
            ComboBox cb = f.Controls.OfType<ComboBox>().First();
            ComboBox cb1 = f.Controls.OfType<ComboBox>().Last();
            CheckBox mir = f.Controls.OfType<CheckBox>().First(), dir = f.Controls.OfType<CheckBox>().ElementAt(1),
                side = f.Controls.OfType<CheckBox>().ElementAt(2), imCheck = f.Controls.OfType<CheckBox>().ElementAt(3);
            if (cb.Text == "") return;
            XElement el = PartsBtn.xDoc.El.Elements("Pattern").FirstOrDefault(ell => ell.Attribute("Name").Value == cb.Text);
            string typ;
            typ = cb1.Text;
            PartDocument d1 = doc;
            XElement nod = el.PreviousNode as XElement;
            if (nod.Name != "Parameter") nod = nod.PreviousNode as XElement; 
            if (doc.ReferencedDocumentDescriptors.Count == 1) d1 = InvDoc.u.referendedDocDesc(doc as Document).ReferencedDocument as PartDocument;
            System.Collections.Generic.Stack<XElement> nods = new Stack<XElement>();
            while (nod != null && nod.Name == "Parameter")
	        {
               nods.Push(nod);
	           nod = nod.PreviousNode as XElement;
	        }
            while (nods.Count != 0)
            {
                ut.addParameter(d1 as Document, nods.Pop()); 
            }

            if (typ == "Добавить")
            {
                this.Close();
                PartsBtn.cc.Activate();
                return;
            }
            f.Hide();
            this.Hide();
            if (typ == "")
                patterns(doc, el, mir: mir.Checked, perp: dir.Checked, side: side.Checked, im: imCheck.Checked);
            else patterns(doc, el, typ, mir: mir.Checked, perp: dir.Checked, side: side.Checked, im: imCheck.Checked);
            this.Close();
            /*PartsBtn.cc.Show();*/
            PartsBtn.cc.Activate();
        }

        private object addMirror(PartDocument doc, ObjectCollection col, object dir, bool rot = false, string name = "")
        {
            PartComponentDefinition compDef = doc.ComponentDefinition;
            UnitVector vec = null;
            if (dir is WorkAxis)
            {
                WorkAxis wa = dir as WorkAxis;
                vec = wa.Line.Direction;
            }
            else if (dir is Edge)
            {
                Edge ed = dir as Edge;
                vec = ed.StartVertex.Point.VectorTo(ed.StopVertex.Point).AsUnitVector();
            }
            else if (dir is Path)
            {
                Path p = dir as Path;
                SketchLine sl = p[1].SketchEntity as SketchLine;
                vec = sl.Geometry3d.Direction;
            }
            
            PartFeature pat = col[1] as PartFeature;
            if (rot && pat.Faces[1].SurfaceType == SurfaceTypeEnum.kCylinderSurface)
            {
                Cylinder cyl = pat.Faces[1].Geometry as Cylinder;
                vec = vec.CrossProduct(cyl.AxisVector);
            }
            WorkPlane pl = ut.planeFromAxis(compDef, vec);
            addMirror(doc, col, pl, name);
            //object dirMir = ut.mirrorDir(doc.ComponentDefinition, vec, pl);
            return pl;
        }

        private void addMirror(PartDocument doc, RectangularPatternFeature pat, object dir, bool rot = false, string name = "")
        {
            PartComponentDefinition compDef = doc.ComponentDefinition;  
            UnitVector vec = null;
            if (dir is WorkAxis)
            {
                WorkAxis wa = dir as WorkAxis;
                vec = wa.Line.Direction;
            }
            else if (dir is Edge)
            {
                Edge ed = dir as Edge;
                vec = ed.StartVertex.Point.VectorTo(ed.StopVertex.Point).AsUnitVector();
            }
            if (rot && pat.Faces[1].SurfaceType == SurfaceTypeEnum.kCylinderSurface)
            {
                Cylinder cyl = pat.Faces[1].Geometry as Cylinder;
                vec = vec.CrossProduct(cyl.AxisVector);
            }
            WorkPlane pl = ut.planeFromAxis(compDef,vec);
            addMirror(doc, pat, pl, name);
        }

        private void addMirror(PartDocument doc, RectangularPatternFeature pat, WorkPlane pl, string name = "")
        {
            ObjectCollection col = I.objs.CreateObjectCollection();
            col.Add(pat);
            MirrorFeature mf = doc.ComponentDefinition.Features.MirrorFeatures.Add(col, pl, false, PatternComputeTypeEnum.kAdjustToModelCompute);
            if (name != "") name = name + "Зерк";
            InvDoc.Reflect.setProp<MirrorFeature, string>(mf, "Name", name);
            //ut.addNameToFeature<PartFeature>(mf as PartFeature, name);
            VariableData vd = new VariableData(doc as Document);
            vd.AttribAdd<RectangularPatternFeature, string>(pat, "Mirror", mf.Name, ValueTypeEnum.kStringType);
        }

        private void addMirror(PartDocument doc, ObjectCollection col, WorkPlane pl, string name = "")
        {
            MirrorFeature mf = doc.ComponentDefinition.Features.MirrorFeatures.Add(col, pl, false, PatternComputeTypeEnum.kAdjustToModelCompute);
            if (name != "") name = name + "Зерк";
            InvDoc.Reflect.setProp<MirrorFeature, string>(mf, "Name", name);
            //ut.addNameToFeature<PartFeature>(mf as PartFeature, name);
            VariableData vd = new VariableData(doc as Document);
            //vd.AttribAdd<string>(pat, "Mirror", mf.Name, ValueTypeEnum.kStringType);
        }

        private void patterns(PartDocument doc, XElement el, string typ = "Отверстие", bool mir = false, bool perp = true, bool side = false, bool im = false)
        {
            string diam = "", count = "", countY = "", step = "", stepY = "", iMate = "", offset = "", fastener = "", offsetFastener = "", a = "", b = "";
            WorkPlane wp = null;
            diam = XMLDoc.getXAttributeValue(el, "DiameterName");
            count = XMLDoc.getXAttributeValue(el, "CountName");
            countY = XMLDoc.getXAttributeValue(el, "CountNameY");
            step = XMLDoc.getXAttributeValue(el, "StepName");
            stepY = XMLDoc.getXAttributeValue(el, "StepNameY");
            iMate = XMLDoc.getXAttributeValue(el, "iMate");
            offset = XMLDoc.getXAttributeValue(el, "offset");
            fastener = XMLDoc.getXAttributeValue(el, "Fastener");
            offsetFastener = XMLDoc.getXAttributeValue(el, "offsetFastener");
            a = XMLDoc.getXAttributeValue(el, "a");
            a = a ?? "5";
            b = XMLDoc.getXAttributeValue(el, "b");
            b = b ?? "8";
            if (diam == null && count == null && step == null) return;
            HashSet<string> names = new HashSet<string>();
            names.Add(diam);
            ut.getParametersFromGroup(doc as Document, count.Remove(count.Length - 3), ref names);
            names.Add(count); names.Add(step);
            if (stepY != "" && countY != "")
            {
                names.Add(stepY); names.Add(countY);
            }
            ut.findParameter(doc as Document, names);
            //doc.Update();
            int direct = int.Parse(el.Attribute("Direct").Value);
            CommandManager cmd = Macros.StandardAddInServer.m_inventorApplication.CommandManager;
            SketchPoint sp = m_Parts.Sp;
            if (sp == null) sp = cmd.Pick(SelectionFilterEnum.kSketchPointFilter, "Выберите исходную точку") as SketchPoint;
            object sk = null;
            if (sp != null /*&& sp.Reference == true*/) sk = addSketchFromDerSketch(doc.ComponentDefinition, sp, el.Attribute("Name").Value, side) as object;
            if (sk == null)
            {
                if (el.Attribute("SketchName") != null)
                {
                    sk = doc.ComponentDefinition.Sketches.OfType<PlanarSketch>().FirstOrDefault(s => s.Name == el.Attribute("SketchName").Value);
                    if (el.Attribute("SketchName").Value == "Last") sk = doc.ComponentDefinition.Sketches.OfType<PlanarSketch>().Where(s => s.Consumed == false).LastOrDefault();
                }
                if (sk == null && typ == "Отверстие") sk = cmd.Pick(SelectionFilterEnum.kSketchObjectFilter, "Выберите эскиз");
                else if (sk == null && typ == "Другой элемент") sk = cmd.Pick(SelectionFilterEnum.kPartFeatureFilter, "Выберите элемент для массива");
            }
            object hf = null;
            if (sk is PlanarSketch && typ == "Отверстие") hf = hole(doc.ComponentDefinition, diam, sk as PlanarSketch);
            else if (sk is PlanarSketch && typ == "Овал")
            {
                SheetMetalFeatures smf = doc.ComponentDefinition.Features as SheetMetalFeatures;
                hf = FPOp.flatCut(smf, sk as PlanarSketch, typ, a, b);
            }
            else hf = sk;
            object dir = null;// = cmd.Pick(SelectionFilterEnum.kPartEdgeLinearFilter, "Направление");

            Point ptCen;
            if (hf is HoleFeature)
            {
                ptCen = ((Circle)(hf as HoleFeature).Faces[1].Edges[1].Geometry).Center;
                hole(hf, diam, doc.ComponentDefinition);
                dir = ut.findAtPoint(doc.ComponentDefinition, ptCen, new SelectionFilterEnum[] { SelectionFilterEnum.kPartEdgeLinearFilter }, axis: true);
            }
            else if (hf is RectangularPatternFeature)
            {
                pattern(hf, count, step, diam, el.Attribute("Name").Value, doc.ComponentDefinition);
                return;
            }
            VariableData vd = new VariableData(doc as Document);
            if (sp.Reference) sp = sp.ReferencedEntity as SketchPoint;
            AttributeSet attSet = vd.getAttrSet(sp, "dir");
            if (attSet != null)
            {
                dir = I.tg.CreateUnitVector((double)attSet["X"].Value, (double)attSet["Y"].Value, (double)attSet["Z"].Value) as object;
            }
            string val = null;
            SketchLine slDir = null;
            Inventor.Attribute att = vd.getAttrib<SketchPoint>(sp, "SL");
            if (att != null) 
            val = att.Value.ToString();
            if (val != null)
            {
                slDir = ut.bindRefkey((sp.Parent.Parent as ComponentDefinition).Document as Document, val) as SketchLine;
            }
            if (slDir != null)
            {
                dir = ut.createPath(doc, slDir);
            }
            else
                dir = ut.axisFromEdge(doc.ComponentDefinition, dir);
            if (dir == null) dir = cmd.Pick(SelectionFilterEnum.kPartEdgeLinearFilter, "Направление");
            object dirMir = null; UnitVector ve = null;
            ObjectCollection col = I.objs.CreateObjectCollection(); col.Add(hf);
            if (mir)
            {
                wp = m_Parts.Wp;
                
                Edge e = dir as Edge;
                if (e != null && e.GeometryType == CurveTypeEnum.kLineSegmentCurve) ve = (e.Geometry as LineSegment).Direction;
                else if (dir is WorkAxis) ve = (dir as WorkAxis).Line.Direction;
                else if (dir is UnitVector) ve = dir as UnitVector;
                if (wp != null) 
                {
                    addMirror(doc, col, wp, el.Attribute("Name").Value);

                    if (ve != null)
                    {
                        dir = ut.axisFromEdge(doc.ComponentDefinition, ve);
                    }
                }
                else if (perp)
                    wp = addMirror(doc, col, dir, perp, el.Attribute("Name").Value) as WorkPlane;
                else wp = addMirror(doc, col, dir, false, el.Attribute("Name").Value) as WorkPlane;
            }
            RectangularPatternFeature rpf = null;
            if (stepY == "" && countY == "")
            {
                if (direct == 2)
                    rpf = pattern(doc.ComponentDefinition, count, step, col, dir, el.Attribute("Name").Value, iMate, offset, fastener, offsetFastener, true, imCheck: im);
                else rpf = pattern(doc.ComponentDefinition, count, step, col, dir, el.Attribute("Name").Value, iMate, offset, fastener, offsetFastener, imCheck: im);
            }
            else
            {
                if (direct == 2)
                    rpf = pattern(doc.ComponentDefinition, count, step, col, dir, el.Attribute("Name").Value, iMate, offset, fastener, offsetFastener, true, countY, stepY, imCheck: im);
                else rpf = pattern(doc.ComponentDefinition, count, step, col, dir, el.Attribute("Name").Value, iMate, offset, fastener, offsetFastener, false, countY, stepY, imCheck: im);
            }
            if (ve != null)
            {
                col.Clear();
                MirrorFeature mf = doc.ComponentDefinition.Features.MirrorFeatures[doc.ComponentDefinition.Features.MirrorFeatures.Count];
                col.Add(mf);
                dirMir = ut.mirrorDir(doc.ComponentDefinition, ve, wp);
                if (stepY == "" && countY == "")
                {
                    if (direct == 2)
                        rpf = pattern(doc.ComponentDefinition, count, step, col, dirMir, el.Attribute("Name").Value, iMate, offset, fastener, offsetFastener, true, imCheck: im);
                    else rpf = pattern(doc.ComponentDefinition, count, step, col, dirMir, el.Attribute("Name").Value, iMate, offset, fastener, offsetFastener, imCheck: im);
                }
                else
                {
                    if (direct == 2)
                        rpf = pattern(doc.ComponentDefinition, count, step, col, dirMir, el.Attribute("Name").Value, iMate, offset, fastener, offsetFastener, true, countY, stepY, imCheck: im);
                    else rpf = pattern(doc.ComponentDefinition, count, step, col, dirMir, el.Attribute("Name").Value, iMate, offset, fastener, offsetFastener, false, countY, stepY, imCheck: im);
                }
            }
        }

        private PlanarSketch addSketchFromDerSketch(PartComponentDefinition compDef,SketchPoint sp, string name = "", bool side = false)
        {
            UnitVector dir = (sp.Parent as PlanarSketch).PlanarEntityGeometry.Normal;
            Vector vec = dir.AsVector(); vec.ScaleBy(-1);
            Point pt = sp.Geometry3d;
            pt.TranslateBy(vec);
            PlanarSketch ps = null;
            int i = (side)?1:0;
            Face f = ut.findAtRay<Face>(compDef, pt, dir, new SelectionFilterEnum[] { SelectionFilterEnum.kPartFacePlanarFilter }, ind: i);
            if (f == null)
            {
                Vector v = dir.AsVector(); v.ScaleBy(-1);
                f = ut.findAtRay<Face>(compDef, pt, v.AsUnitVector(), new SelectionFilterEnum[] { SelectionFilterEnum.kPartFacePlanarFilter });
            }
            if (f != null)
            {
               ps = compDef.Sketches.Add(f);
               name = name + "Эскиз";
               while (InvDoc.Reflect.exist<PlanarSketch>(compDef.Sketches, name))
               {
                   name = nextName(name);
               }
               if (name != "" && ps.Name != name) ps.Name = name;
               ps.AddByProjectingEntity(sp);
            }
            return ps;
        }

        private void массивToolStripMenuItem_Click(object sender, EventArgs e)
        {
            PartDocument doc = Macros.StandardAddInServer.m_inventorApplication.ActiveDocument as PartDocument;
            m_Parts.Ss = doc.SelectSet;

            string name = "";
            Transaction tr = ut.transAct(doc, "Массив");
            try
            {
                if (PartsBtn.xDoc == null) return;
                PartsBtn.xDoc = PartsBtn.xDoc ?? new XMLDoc(name, "head");
                formaForPattern();
            }

            finally
            {
                tr.End();
            }
            //this.Close();
        }

        private void крепежToolStripMenuItem1_Click(object sender, EventArgs e)
        {
//             XMLDoc xDoc = new XMLDoc(@"C:\ProgramData\Autodesk\Inventor Addins\AutoFastener.xml", "Fasteners");
//             List<string> except = new List<string> { "name", "Direction", "Offset"};
//             foreach (var el in xDoc.El.Descendants())
//             {
//                 foreach (var at in el.Attributes())
//                 {
//                     if (!except.Exists(et => et == at.Name))
//                     {
//                         at.Remove();
//                     }
//                 }
//             }
//             xDoc.save();

             Fastener f = new Fastener(I.aDoc());
             f.add();
        }

        private void добавитьДанныеДляМассиваToolStripMenuItem_Click(object sender, EventArgs e)
        {
            PartDocument doc = Macros.StandardAddInServer.m_inventorApplication.ActiveDocument as PartDocument;
            PartsBtn.Doc = doc;
            if (doc.ReferencedDocumentDescriptors.Count == 1) PartsBtn.Doc = InvDoc.u.referendedDocDesc(doc as Document).ReferencedDocument as PartDocument;
            string name = "";
            Transaction tr = ut.transAct(doc, "Массив");
            try
            {
                if (PartsBtn.xDoc == null)
                {
                    name = ut.OFD(doc.path(), "XML(*.xml)|*.xml", false);
                    if (name == null) return;
                }
                PartsBtn.xDoc = PartsBtn.xDoc ?? new XMLDoc(name, "head");
                formaForPatternData();
            }
            finally
            {
                tr.End();
            }
        }

        private void formaForPatternData()
        {
            Form f = new Form();
            InterfaceDll.MyLabel mylbl1 = new MyLabel(); int offsetY = 30, offsetX = 10, w1 = 200, w2 = 50, w3 = 150;
            System.Drawing.Point poz, poz2;
            InterfaceDll.MyTextBox myTxt = new MyTextBox(); MyButton myBtn = new MyButton();
            f.Height = 150+50; f.Width = 400; f.WindowState = FormWindowState.Normal; f.Text = "Массив данные"; f.StartPosition = FormStartPosition.CenterScreen;
            System.Drawing.Point insPt = new System.Drawing.Point(5, 5);
            Label lbl = mylbl1.addLabel("Название", insPt, 100, 15);
            f.Controls.Add(lbl);
            System.Windows.Forms.TextBox txtbox = myTxt.addTextBox("", MyLabel.position(lbl, offsetX, 0), w1, 15);
            f.Controls.Add(txtbox);
            poz = InterfaceDll.MyLabel.position(txtbox, offsetX, -7);
            //ut.autoComplete<PartDocument>(txtbox, PartsBtn.Doc, "Кол");
            lbl = mylbl1.addLabel("Кол-во X", InterfaceDll.MyLabel.position(lbl, -lbl.Width, offsetY), 100, 15);
            f.Controls.Add(lbl);
            lbl = mylbl1.addLabel("Шаг X", InterfaceDll.MyLabel.position(lbl, -lbl.Width, offsetY), 100, 15);
            f.Controls.Add(lbl);
            txtbox = myTxt.addTextBox("", MyLabel.position(txtbox, -txtbox.Width, offsetY), w2, 15);
            f.Controls.Add(txtbox);
            poz2 = InterfaceDll.MyLabel.position(txtbox, offsetX+30, 0);
            txtbox = myTxt.addTextBox("", MyLabel.position(txtbox, -txtbox.Width, offsetY), w2, 15);
            f.Controls.Add(txtbox);
            lbl = mylbl1.addLabel("Отверстие", InterfaceDll.MyLabel.position(lbl, -lbl.Width, offsetY), 100, 15);
            f.Controls.Add(lbl);
            txtbox = myTxt.addTextBox("", InterfaceDll.MyLabel.position(txtbox,-txtbox.Width, offsetY), w3, 15);
            f.Controls.Add(txtbox);
            if (PartsBtn.param != null)
            ut.autoComplete(txtbox, PartsBtn.param, "Parameter", "Name");
            txtbox.Leave += txtbox_TextChanged;
            lbl = mylbl1.addLabel("Диаметр", InterfaceDll.MyLabel.position(txtbox, offsetX, 0), 100, 15);
            f.Controls.Add(lbl);
            txtbox = myTxt.addTextBox("", MyLabel.position(lbl, offsetX, 0), w2, 15);
            f.Controls.Add(txtbox);

            lbl = mylbl1.addLabel("Кол-во Y", poz2, 100, 15);
            f.Controls.Add(lbl);
            txtbox = myTxt.addTextBox("", MyLabel.position(lbl, offsetX, 0), w2, 15);
            f.Controls.Add(txtbox);
            lbl = mylbl1.addLabel("Шаг Y", InterfaceDll.MyLabel.position(lbl, -lbl.Width, offsetY), 100, 15);
            f.Controls.Add(lbl);
            //poz = MyLabel.position(txtbox, offsetX, -7);
            txtbox = myTxt.addTextBox("", MyLabel.position(txtbox, -txtbox.Width, offsetY), w2, 15);
            f.Controls.Add(txtbox);
            System.Windows.Forms.Button btn = myBtn.addButton("Добавить", poz, 100, 30);
            f.Controls.Add(btn);
            insPt.Y += 4 * offsetY;
            lbl = mylbl1.addLabel("Ответ отв", insPt, 100, 15);
            f.Controls.Add(lbl);
            txtbox = myTxt.addTextBox("", MyLabel.position(lbl, offsetX, 0), w3, 15);
            f.Controls.Add(txtbox);
            if (PartsBtn.param != null)
                ut.autoComplete(txtbox, PartsBtn.param, "Parameter", "Name");
            txtbox.Leave += txtbox_TextChanged_1;
            //insPt.X = poz2.X;
            lbl = mylbl1.addLabel("Диаметр", MyLabel.position(txtbox, offsetX, 0), 100, 15);
            f.Controls.Add(lbl);
            txtbox = myTxt.addTextBox("", MyLabel.position(lbl, offsetX, 0), w2, 15);
            f.Controls.Add(txtbox);

            btn.Click += forma_Data_Click;
            f.Show();
        }

        void txtbox_TextChanged(object sender, EventArgs e)
        {
            System.Windows.Forms.TextBox tb = sender as System.Windows.Forms.TextBox;     
            Form f = tb.Parent as Form;
            System.Windows.Forms.TextBox tbVal = f.Controls.OfType<System.Windows.Forms.TextBox>().ElementAt(4);
            if (PartsBtn.param != null)
            {
               XElement el = PartsBtn.param.getXElement(tb.Text, "Name", "Parameter");
               if (el != null) tbVal.Text = el.Attribute("Value").Value;
            } 
        }

        void txtbox_TextChanged_1(object sender, EventArgs e)
        {
            System.Windows.Forms.TextBox tb = sender as System.Windows.Forms.TextBox;
            Form f = tb.Parent as Form;
            System.Windows.Forms.TextBox tbVal = f.Controls.OfType<System.Windows.Forms.TextBox>().ElementAt(8);
            if (PartsBtn.param != null)
            {
                XElement el = PartsBtn.param.getXElement(tb.Text, "Name", "Parameter");
                if (el != null) tbVal.Text = el.Attribute("Value").Value;
            }
        }

        private void forma_Data_Click(object sender, EventArgs e)
        {
            PartDocument doc = Macros.StandardAddInServer.m_inventorApplication.ActiveDocument as PartDocument;
            Form f = (Form)((System.Windows.Forms.Button)sender).Parent;
            System.Windows.Forms.TextBox tbName = f.Controls.OfType<System.Windows.Forms.TextBox>().ElementAt(0),
            tbCount = f.Controls.OfType<System.Windows.Forms.TextBox>().ElementAt(1),
            tbStep = f.Controls.OfType<System.Windows.Forms.TextBox>().ElementAt(2),
            tbDiamName = f.Controls.OfType<System.Windows.Forms.TextBox>().ElementAt(3),
            tbDiam = f.Controls.OfType<System.Windows.Forms.TextBox>().ElementAt(4),
            tbCountY = f.Controls.OfType<System.Windows.Forms.TextBox>().ElementAt(5),
            tbStepY = f.Controls.OfType<System.Windows.Forms.TextBox>().ElementAt(6),
            tbDiamCounterName = f.Controls.OfType<System.Windows.Forms.TextBox>().ElementAt(7),
            tbDiamCounter = f.Controls.OfType<System.Windows.Forms.TextBox>().ElementAt(8);
            addElem add = new addElem(tbName.Text, "Шаг", "Кол", "Смещение");

            PartsBtn.xDoc.El.Add(add.addParameter(tbDiamName.Text, tbDiam.Text));
            add.addParameters(PartsBtn.xDoc.El, new string[] {tbStep.Text, tbCount.Text });
            if (tbCountY.Text != "" && tbStepY.Text != "") add.addParameters(PartsBtn.xDoc.El, new string[] { tbStepY.Text, tbCountY.Text }, "Y");
            string suff = (tbDiamCounterName.Text.ToLower().EndsWith("отв"))?"":"Отв";
            if (tbDiamCounterName.Text != "") PartsBtn.xDoc.El.Add(add.addParameter(tbDiamCounterName.Text, tbDiamCounter.Text, suff));
            XElement el = add.addPattern(tbDiamName.Text);
            if (tbCountY.Text != "" && tbStepY.Text != "") add.addPattern(el);
            PartsBtn.xDoc.El.Add(el);
            if (tbDiamCounterName.Text != "")
            {
                el = add.addPattern(tbDiamCounterName.Text, "Отв");
                if (tbCountY.Text != "" && tbStepY.Text != "") add.addPattern(el);
                PartsBtn.xDoc.El.Add(el);
            }

//             XElement el = new XElement("Parameter", new XAttribute("Name", tbDiamName.Text), new XAttribute("Value", tbDiam.Text), new XAttribute("Type", "mm"), new XAttribute("Group", tbName.Text));
//             PartsBtn.xDoc.El.Add(el);
//             el = new XElement("Parameter", new XAttribute("Name", tbName.Text + "Шаг"), new XAttribute("Value", tbStep.Text), new XAttribute("Type", "mm"), new XAttribute("Group", tbName.Text));
//             PartsBtn.xDoc.El.Add(el);
//             el = new XElement("Parameter", new XAttribute("Name", tbName.Text + "Кол"), new XAttribute("Value", tbCount.Text), new XAttribute("Type", "ul"), new XAttribute("Group", tbName.Text));
//             PartsBtn.xDoc.El.Add(el);
//             el = new XElement("Parameter", new XAttribute("Name", tbName.Text + "Смещение"),
//                 new XAttribute("Value", tbName.Text + "Шаг / 2 бр * ( ( " + tbName.Text + "Кол - 1 бр ) % 2 бр )"), new XAttribute("Type", "mm"), new XAttribute("Formula", "1"), new XAttribute("Group", tbName.Text));
//             PartsBtn.xDoc.El.Add(el);
//             if (tbCountY.Text != "" && tbStepY.Text != "")
//             {
//                 el = new XElement("Parameter", new XAttribute("Name", tbName.Text + "ШагY"), new XAttribute("Value", tbStepY.Text), new XAttribute("Type", "mm"), new XAttribute("Group", tbName.Text));
//                 PartsBtn.xDoc.El.Add(el);
//                 el = new XElement("Parameter", new XAttribute("Name", tbName.Text + "КолY"), new XAttribute("Value", tbCountY.Text), new XAttribute("Type", "ul"), new XAttribute("Group", tbName.Text));
//                 PartsBtn.xDoc.El.Add(el);
//                 el = new XElement("Parameter", new XAttribute("Name", tbName.Text + "СмещениеY"),
//                 new XAttribute("Value", tbName.Text + "ШагY / 2 бр * ( ( " + tbName.Text + "КолY - 1 бр ) % 2 бр )"), new XAttribute("Type", "mm"), new XAttribute("Formula", "1"), new XAttribute("Group", tbName.Text));
//                 PartsBtn.xDoc.El.Add(el);
//                 el = new XElement("Pattern", new XAttribute("Name", tbName.Text), new XAttribute("SketchName", "Last"), new XAttribute("DiameterName", tbDiamName.Text),
//                                 new XAttribute("CountName", tbName.Text + "Кол"), new XAttribute("StepName", tbName.Text + "Шаг"),
//                                 new XAttribute("CountNameY", tbName.Text + "КолY"), new XAttribute("StepNameY", tbName.Text + "ШагY"),
//                                 new XAttribute("Direct", "2"));
//             }
//             else
//             {
//                 el = new XElement("Pattern", new XAttribute("Name", tbName.Text), new XAttribute("SketchName", "Last"), new XAttribute("DiameterName", tbDiamName.Text),
//                              new XAttribute("CountName", tbName.Text + "Кол"), new XAttribute("StepName", tbName.Text + "Шаг"), new XAttribute("Direct", "2"));
//             }
//             PartsBtn.xDoc.El.Add(el);
            PartsBtn.xDoc.save();
        }

        public class addElem
        {
            string name, group, suff1, suff2, suff3;
            public addElem(string name, string s1, string s2, string s3)
            {
                this.name = name; group = name; suff1 = s1; suff2 = s2; suff3 = s3; 
            }
            public XElement addParameter(string n, string val,string suff = "", string typ = "mm")
            {
                return new XElement("Parameter", new XAttribute("Name", n + suff), new XAttribute("Value", val),
                    new XAttribute("Type", typ), new XAttribute("Group", group));
            }
            public void addParameters(XElement el, string [] vals, string y = "")
            {
                el.Add(addParameter(name,vals[0],suff1 + y));
                el.Add(addParameter(name,vals[1],suff2 + y,"ul"));
                string val = name + suff1 + y + " / 2 бр * ( ( " + name + suff2 + y + " - 1 бр ) % 2 бр )";
                XElement par = addParameter(name, val, suff3 + y);
                par.Add(new XAttribute("Formula", 1));
                el.Add(par);
            }
            public XElement addPattern(string name1)
            {
                return new XElement("Pattern", new XAttribute("Name", name), new XAttribute("SketchName", "Last"), new XAttribute("DiameterName", name1),
                    new XAttribute("CountName", name + suff2), new XAttribute("StepName", name + suff1), new XAttribute("Direct", "2"));
            }
            public XElement addPattern(string name1, string suff = "")
            {
                return new XElement("Pattern", new XAttribute("Name", name + suff), new XAttribute("SketchName", "Last"), new XAttribute("DiameterName", name1),
                    new XAttribute("CountName", name + suff2), new XAttribute("StepName", name + suff1), new XAttribute("Direct", "2"));
            }
            public void addPattern(XElement el, string y = "Y")
            {
                el.Add(new XAttribute("CountName" + y, name + suff2 + y), new XAttribute("StepName" + y, name + suff1 + y));
            }
        }

        private void сортировкаToolStripMenuItem_Click(object sender, EventArgs e)
        {
            XMLDoc im = new XMLDoc(@"C:\ProgramData\Autodesk\Inventor Addins\Imate.xml", "head"),
                af = new XMLDoc(@"C:\ProgramData\Autodesk\Inventor Addins\AutoFastener.xml", "Fasteners");
            XElement el = im.El.Element("Imates");
            foreach (var item in af.El.Descendants("Fastener"))
            {
               el.Add(new XElement("Value", item.Attribute("name").Value)); 
            }
            foreach (var item in af.El.Descendants("Composite"))
            {
               //item.Attribute("name").Value = "_" + item.Attribute("name").Value;
               el.Add(new XElement("Value", item.Attribute("name").Value)); 
            }
            im.sortAlph();
            //af.save();
            im.save();
//             XMLDoc xdoc = ut.XOFD(@"C:\ProgramData\Autodesk\Inventor Addins");
//             xdoc.sortAlph();
//             xdoc.save();
        }

        private void артикулToolStripMenuItem_Click(object sender, EventArgs e)
        {
            DrawingDocument drw = Macros.StandardAddInServer.m_inventorApplication.ActiveDocument as DrawingDocument;
            Sheet sh = drw.ActiveSheet;
            foreach (Balloon bal in sh.Balloons)
            {
                BalloonValueSet vs = bal.BalloonValueSets[1];
                Document doc = vs.ReferencedRow.BOMRow.ComponentDefinitions[1].Document as Document;
                string val = InvDoc.u.getProp(doc, "Catalog Web Link").Value.ToString();
                vs.OverrideValue = val;
            }
        }

        private void перенаправитьToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            Document doc = Macros.StandardAddInServer.m_inventorApplication.ActiveDocument;
            string path = doc.path();
//             if (doc.DocumentType == DocumentTypeEnum.kAssemblyDocumentObject)
//             {
//                 refReplace(doc, path);
//             }
            InvDocument<Document> invDoc = new InvDocument<Document>(doc);
            List<string> files = new List<string>();
            files.AddRange(invDoc.openFiles("*.idw", true));
            foreach (var item in files)
            {
                refReplace(invDoc.openDoc(item), path);
            }
        }

        private void refReplace(Document doc, string path)
        {
            foreach (DocumentDescriptor desc in doc.ReferencedDocumentDescriptors)
            {
//                 if (desc.ReferencedDocumentType == DocumentTypeEnum.kAssemblyDocumentObject)
//                 {
//                     refReplace(desc.ReferencedDocument as Document, path);
//                 }
                string newName = path + "\\" + System.IO.Path.GetFileName(desc.FullDocumentName);
                if (System.IO.File.Exists(newName))
                {
                    desc.ReferencedFileDescriptor.ReplaceReference(newName);
                }
            }
            if (doc.Dirty) doc.Save2();
        }

        private void пересобратьСборкуToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Document doc;
            InvDocument<Document> invDoc = new InvDocument<Document>(doc = Macros.StandardAddInServer.m_inventorApplication.ActiveDocument);
            XMLDoc xmlDoc = new XMLDoc(doc.path() + "\\" + "compare.xml", "head");
            List<string> files = new List<string>();
            Macros.StandardAddInServer.m_inventorApplication.SilentOperation = true;
            files = invDoc.openFiles("*.ipt", true);
            files.AddRange(invDoc.openFiles("*.iam", true));
            files.AddRange(invDoc.openFiles("*.idw", true));
            foreach (var item in files)
            {
//                 if (item.EndsWith("iam"))
//                 {
//                     Parts.replaceReference(invDoc.openAsmDoc(item), xmlDoc);
//                 }
//                 else if (item.EndsWith("ipt")) Parts.replaceReference(invDoc.openPrtDoc(item), xmlDoc);
// 
//                 else if (item.EndsWith("idw"))
//                 {
//                     //if(invDoc.nvmOptions.Count == 1) invDoc.nvmOptions.Add("DeferUpdates", true);
//                     Parts.replaceReference(invDoc.openDrwDoc(item), xmlDoc);
//                 }
                Parts.replaceReference(invDoc.openDoc(item), xmlDoc);
            }
            Macros.StandardAddInServer.m_inventorApplication.SilentOperation = false;
        }

        private void скругленияToolStripMenuItem_Click(object sender, EventArgs e)
        {
            MyForm F = new MyForm("KPSInterface.xml", "Тест");
            //F.bnts[0].Click += fillet_Click;
            F.f.ShowDialog();
            foreach (Document doc in I.app.Documents.VisibleDocuments)
            {
                if (doc.DocumentType == DocumentTypeEnum.kPartDocumentObject)
                {
                    fillet(doc as PartDocument, ut.convToDouble(F.cbs[0].Text), ut.convToDouble(F.cbs[1].Text));
                }
            }
            //F.f.Close();
            this.Close();
//             int offsetY = 30, offsetX = 10;
//             Form f = new Form();
//             f.Height = 150; f.Width = 400; f.WindowState = FormWindowState.Normal; f.Text = "Добавить скругления"; f.StartPosition = FormStartPosition.CenterScreen;
//             f.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink; f.AutoSize = true;
//             System.Drawing.Point insPt = new System.Drawing.Point(5, 5);
//             InterfaceDll.Lbl lbs = new Lbl(offsetX, offsetY, 100, 15, insPt, f, "Внутренний радиус");
//             InterfaceDll.CB cbs = new CB(offsetX, offsetY, 200, 15, insPt, f); 
//             cbs.position(lbs.last(), true);
//             cbs.last().Items.AddRange(new[] { "2.65", "2.8", "3", "4.5", "6" } );
//             cbs.last().Text = "3";
//             lbs.add("Внешний радиус", Y: true);
//             cbs.add("", new[] { "5" }, Y: true);
//             cbs.last().Text = "5";
//             InterfaceDll.Btn btns = new Btn(offsetX, offsetY, 100, 20, insPt, f, fillet_Click, "Добавить");
//             btns.center(cbs.last(), offsetY + 5);
//             f.Show();

            //InterfaceDll.MyLabel mylbl1 = new MyLabel();
            //InterfaceDll.MyTextBox mytxt1 = new MyTextBox();
            //InterfaceDll.MyComboBox mycb = new MyComboBox();
            //Label lbl1 = mylbl1.addLabel("Внешний радиус", insPt, 100, 15);
            //f.Controls.Add(lbl1);
            //insPt.Y += offsetY;
            //lbl1 = mylbl1.addLabel("Внутренний радиус", insPt, 100, 15);
            //f.Controls.Add(lbl1);
            //insPt.X += 150; insPt.Y -= offsetY;
            //System.Windows.Forms.ComboBox cb = mycb.addComboBox("2", insPt, 200, 15, new string[] { "5"});
            //cb.Text = "2";
            //insPt.Y += offsetY;
            //System.Windows.Forms.ComboBox cb2 = mycb.addComboBox("2,5", insPt, 200, 15, new string[] { "2.65", "2.8", "3", "4.5", "6" });
            //cb2.Text = "2.5";
            //f.Controls.Add(cb); f.Controls.Add(cb2);
            //lbl1 = mylbl1.addLabel("Не учитывать длину менее", InterfaceDll.MyLabel.position(lbl1, 0, offsetY), 100, 15);
            //f.Controls.Add(lbl1);
            //MyButton myBtn = new MyButton();
            //insPt.Y += offsetY; insPt.X = 200 - 50;
            //System.Windows.Forms.Button btn = myBtn.addButton("Добавить", insPt, 100, 20);
            //f.Controls.Add(btn);
            //btn.Click += fillet_Click;
            //f.Show();
        }

        void fillet_Click(object sender, EventArgs e)
        {
            Form f = (Form)((System.Windows.Forms.Button)sender).Parent;
            foreach (Document doc in Macros.StandardAddInServer.m_inventorApplication.Documents.VisibleDocuments)
            {
                if (doc.DocumentType == DocumentTypeEnum.kPartDocumentObject)
                {
                    var ie = f.Controls.OfType<ComboBox>();
                    fillet(doc as PartDocument,ut.convToDouble(ie.ElementAt(1).Text),ut.convToDouble(ie.ElementAt(0).Text));
                }
            }
            Control par = f.Parent;
            f.Close();
            PartsBtn.cc.Close();
        }

        private bool condition(Edge ed, double r, double t)
        {
            if (!ut.eq(ut.getLenght(ed), t)) return false;
            Vertex v = ed.StartVertex;
            if (!ed.EdgeUses[1].EdgeLoop.IsOuterEdgeLoop) return false;
            IEnumerable<Edge> eds = ut.gets<Edge>(v.Edges, f => !f.Equals(ed));
            if (eds.Count() != 2) return false;
            Edge e1 = eds.ElementAt(0), e2 = eds.ElementAt(1);
            LineSegment l1 = e1.Geometry as LineSegment, l2 = e2.Geometry as LineSegment;
            if (l1 == null || l2 == null) return false;
            if (!ut.eq(l1.Direction.AngleTo(l2.Direction), Math.PI / 2)) return false;
            if (ut.getLenght(e1) < r || ut.getLenght(e2) < r) return false;
//             foreach (Face item in ed.Faces)
//             {
//                 if (!condition(item, r, t)) return false; 
//             }
            return true;
        }

        private bool condition(Face f, double r, double t)
        {
            foreach (Edge item in f.Edges)
            {
                if (item.GeometryType == CurveTypeEnum.kCircularArcCurve) 
                    return false;
                double l = InvDoc.u.getLenght(item);
                if (!InvDoc.u.eq(l, t) && l < r) return false;
            }
            return true;
        }

        private void fillet(PartDocument doc,double r1, double r2)
        {
            r1 /= 10; r2 /= 10;
            MeasureTools mt = Macros.StandardAddInServer.m_inventorApplication.MeasureTools;
            EdgeCollection col = Macros.StandardAddInServer.m_inventorApplication.TransientObjects.CreateEdgeCollection(), 
                inner = Macros.StandardAddInServer.m_inventorApplication.TransientObjects.CreateEdgeCollection(),
                outer = Macros.StandardAddInServer.m_inventorApplication.TransientObjects.CreateEdgeCollection();
            List<Point> innerVertx = new List<Point>();
            if (!(doc.SubType == "{9C464203-9BAE-11D3-8BAD-0060B0CE6BB4}")) return;
            SheetMetalComponentDefinition compDef = doc.ComponentDefinition as SheetMetalComponentDefinition;
            double t = InvDoc.u.convToDouble(compDef.Thickness.Value.ToString());
            SheetMetalFeatures feat = compDef.Features as SheetMetalFeatures;
            //HighlightSet hs = doc.CreateHighlightSet();
            foreach (Edge ed in compDef.SurfaceBodies[1].ConcaveEdges)
            {
                if (condition(ed,r2*2,t))
                    /*if (ut.inBody(compDef, ed, r2 * 2))*/ inner.Add(ed);
            }
            if (inner.Count != 0)
            {
                add(inner, "Внутренние скругления", r2, feat);
//                 CornerRoundDefinition indef = feat.CornerRoundFeatures.CreateCornerRoundDefinition(inner, r2);
//                 foreach (CornerRoundFeature item in feat.CornerRoundFeatures)
//                 {
//                     if (item.Name == "Внутренние скругления " + r2 * 10) item.Delete();
//                 }
//                 doc.Update2();
//                 feat.CornerRoundFeatures.Add(indef).Name = "Внутренние скругления " + r2 * 10;
            }        
            foreach (Edge item in compDef.SurfaceBodies[1].Edges)
	            {
                    if (condition(item, r1 * 2, t))
                        /*if (ut.inBody(compDef, item, r1 * 2))*/ 
                            outer.Add(item);
	            }      
            if (outer.Count != 0)
            {
                add(outer, "Внешние скругления", r1, feat);
//                 CornerRoundDefinition outdef = feat.CornerRoundFeatures.CreateCornerRoundDefinition(outer, r1);
//                 foreach (CornerRoundFeature item in feat.CornerRoundFeatures)
//                 {
//                     if (item.Name == "Внешние скругления " + r1 * 10) item.Delete();
//                 }
//                 doc.Update2();
//                 feat.CornerRoundFeatures.Add(outdef).Name = "Внешние скругления " + r1 * 10;
            }
        }

        public void add(EdgeCollection col, string name, double r, SheetMetalFeatures f)
        {
            name = name + " " + r*10;
            CornerRoundFeature crf = ut.get<CornerRoundFeature>(f.CornerRoundFeatures, el => el.Name == name);
            CornerRoundDefinition crd;
            if (crf != null)
            {
                name += "_1";
            }
            crd = f.CornerRoundFeatures.CreateCornerRoundDefinition(col, r);
            crf = f.CornerRoundFeatures.Add(crd);
            crf.Name = name;
        }

        private void тестToolStripMenuItem_Click(object sender, EventArgs e)
        {
            
            //Elements els = new Elements(I.aDoc());
//             Sheet sh = I.getSheet();
//             ClientGraphicsCollection cgc = sh.ClientGraphicsCollection;
//             foreach (ClientGraphics item in cgc)
//             {
//                 int c = item.Count;
//             }

            //file f = new file(I.aDoc().FullFileName);
            //string p = f.p(), name = f.name(), ext = f.ext(), ffn = p + name + ext;

//             Sheet sh = I.getSheet();
//             DataIO io = sh.DataIO;  
//             string [] f = new string []{};
//             StorageTypeEnum [] ste = new StorageTypeEnum []{};
//             io.GetOutputFormats(ref f, ref ste);
//             foreach (var item in f)
//             {
//                 io.WriteDataToFile(item, @"c:\WORK\Завесы\Ангары\t.dwf");
//             }


//             PresentationDocument doc = I.app.ActiveDocument as PresentationDocument;
//             PresentationExplodedView view = doc.ActiveExplodedView;
//             foreach (Trail t in view.Trails)
//             {
//                 foreach (TrailSegment ts in t)
//                 {
//                     LineSegment geom = ts.Geometry as LineSegment;
//                     Point pt = ut.midPt(geom.StartPoint, geom.EndPoint);
//                     geom.EndPoint = pt;
//                 }
//             }
//             DrawingDocument drw = I.aDoc as DrawingDocument;
//             DrawingDocument from = I.app.Documents.Open(@"C:\ProgramData\Autodesk\Inventor Addins\Cluster.idw", false) as DrawingDocument;
//             InvDocument<DrawingDocument>.copySketchSymbolDefinition(from, drw, "Массив");
//             Sheet sh = null;
//             if (drw.Sheets.Count < 3)
//             {
//                 sh = drw.Sheets[1];
//                 InvDocument<DrawingDocument>.addSketchedSymbol(sh, "Массив", new string[] { }, I.CP2d());
//             }
            //DrawingCurveSegment dc1 = I.ss[1] as DrawingCurveSegment, dc2 = I.ss[2] as DrawingCurveSegment;
            
//            util.addDim<LinearGeneralDimension>(dc1.Parent.Parent, dc1.Parent, dc2.Parent, 1.5, DimensionTypeEnum.kHorizontalDimensionType);
//             OrderDims dims = new OrderDims(I.ss);
//             dims.add();


//             PartDocument doc = Macros.StandardAddInServer.m_inventorApplication.ActiveDocument as PartDocument;
//             Reflect.getProp(doc, "Name");

//             DrawingDocument doc = Macros.StandardAddInServer.m_inventorApplication.ActiveDocument as DrawingDocument;
//             foreach (ReferencedOLEFileDescriptor item in doc.ReferencedOLEFileDescriptors)
//             {
//                 item.Delete();
//             }


//             PartDocument doc = Macros.StandardAddInServer.m_inventorApplication.ActiveDocument as PartDocument;
//             FlatPattern fp = (doc.ComponentDefinition as SheetMetalComponentDefinition).FlatPattern;
//             DataIO di;
//             di = fp.DataIO;
//             string sOut = "FLAT PATTERN DXF?AcadVersion=2000&OuterProfileLayer=0&InteriorProfilesLayer=0" +
//                     "&FeatureProfileLayer=1" + "&UnconsumedSketchesLayer=0" +
//                     "&SimplifySplines=True&SplineTolerance=0,1&RebaseGeometry=True&MergeProfilesIntoPolyline=False" +
//                     "&InvisibleLayers=IV_TANGENT;IV_BEND;IV_BEND_DOWN;IV_TOOL_CENTER;IV_TOOL_CENTER_DOWN;IV_ARC_CENTERS;IV_FEATURE_PROFILES_DOWN";
//             string [] formats = new string[6];
//             StorageTypeEnum[] en = new StorageTypeEnum[6];
//             di.GetOutputFormats(ref formats, ref en);
//             di.WriteDataToFile("ACIS SAT", @"C:\tmp.sat");
//             fp.Edit();
//             Edge ed = Macros.StandardAddInServer.m_inventorApplication.CommandManager.Pick(SelectionFilterEnum.kPartEdgeLinearFilter, "Edge") as Edge;
//             ut.drawDirection(fp, doc.Assets[6],ed, "n");

//             PartComponentDefinition compDef = doc.ComponentDefinition;
//             ClientGraphics gs = compDef.ClientGraphicsCollection.Add("Test");
//             GraphicsDataSets data = doc.GraphicsDataSetsCollection.Add("Test");
//             GraphicsNode node = gs.AddNode(1);
//             LineGraphics line = node.AddLineGraphics();
//             node.Appearance = InvDoc.util.createColor(doc, "Black_", "Черный_", 0, 255, 0);
//             GraphicsCoordinateSet coord = data.CreateCoordinateSet(1);
//             coord.Add(1, I.tg.CreatePoint());
//             coord.Add(2, I.tg.CreatePoint(20));
//             line.CoordinateSet = coord;
//             //LineStripGraphics lstrip = gs.AddNode(2).AddLineStripGraphics();
//             Macros.StandardAddInServer.m_inventorApplication.ActiveView.Update();
        }

        private void копироватьВXMLToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (Macros.StandardAddInServer.m_inventorApplication.ActiveDocument.DocumentType == DocumentTypeEnum.kPartDocumentObject)
            {
                PartDocument doc = Macros.StandardAddInServer.m_inventorApplication.ActiveDocument as PartDocument;
                XMLDoc xmlDoc = new XMLDoc(System.IO.Path.GetFileNameWithoutExtension(doc.FullFileName) + ".xml", "head");
                BrowserNode bn = doc.BrowserPanes["Модель"].TopNode.BrowserNodes[1];
                foreach (Parameter par in doc.ComponentDefinition.Parameters)
	                {
		                if (par.ParameterType == ParameterTypeEnum.kModelParameter && !par.Name.StartsWith("d")) ut.parameterToXML(xmlDoc, par);
                        if (par.ParameterType == ParameterTypeEnum.kUserParameter) ut.parameterToXML(xmlDoc,par);
	                }
                foreach (BrowserNode item in bn.BrowserNodes)
                {
                    if (item.NativeObject is WorkPlane)
                    {
                        ut.planeToXML(xmlDoc, item.NativeObject as WorkPlane);
                    }
                    if (item.NativeObject is PlanarSketch) ut.sketchToXML(xmlDoc, item.NativeObject as PlanarSketch);
                }
                xmlDoc.save();
            }
        }

        class CompX : IComparer<Balloon>
        {
            double r;

            public CompX(double r)
            {
                this.r = r;
            }
            public int Compare(Balloon b1, Balloon b2)
            {
                double d = b1.Position.X-b2.Position.X;
                if (InvDoc.u.eq(d, 0))
                {
                    return 0;
                }
                else if (d > 0)
                    return 1;
                else
                    return -1;
            }
        }

        class CompY : IComparer<Balloon>
        {
            double r;

            public CompY()
            {
            }
            public int Compare(Balloon b1, Balloon b2)
            {
                double d = b1.Position.Y - b2.Position.Y;
                if (InvDoc.u.eq(d, 0))
                {
                    return 0;
                }
                else if (d < 0)
                    return 1;
                else
                    return -1;
            }
        }

        public static void alignX(double r, List<Balloon> b)
        {
            int i = 1;
            Balloon bb = b[0];
            List<Balloon> bs = new List<Balloon>();
            bs.Add(bb);
            while (i < b.Count)
            {
                double d = Math.Abs(bb.Position.X - b[i].Position.X);
                if (d < r)
                {
                    bs.Add(b[i]); 
                }
                else
                {
                    bb = b[i];
                    alignXCenter(r, bs);
                    bs.Clear();
                    bs.Add(bb);
                }
                i++;
            }
            if (bs.Count != 0) alignXCenter(r, bs);
            //b.RemoveRange(0, i - 1);
        }

        public static void alignY(double r, List<Balloon> b)
        {
            int i = 1;
            Balloon bb = b[0];
            List<Balloon> bs = new List<Balloon>();
            bs.Add(bb);
            while (i < b.Count)
            {
                double d = Math.Abs(bb.Position.Y - b[i].Position.Y);
                if (d < r)
                {
                    bs.Add(b[i]);
                }
                else
                {
                    bb = b[i];
                    alignYCenter(r, bs);
                    bs.Clear();
                    bs.Add(bb);
                }
                i++;
            }
            if (bs.Count != 0) alignYCenter(r, bs);
            //b.RemoveRange(0, i - 1);
        }

        public static void alignYCenter(double r, List<Balloon> b)
        {
            double y = 0;
            y = b.Average(e => e.Position.Y);
            foreach (Balloon item in b)
            {
                Point2d pt = I.tg.CreatePoint2d(item.Position.X, y);
                item.Position = pt; 
            }
        }

        public static void alignXCenter(double r, List<Balloon> b)
        {
            double x = 0, y = 0, offset = 0;
            b.Sort(new CompY());
            x = b.Average(e => e.Position.X);
            Point2d pt = I.tg.CreatePoint2d(x, b[0].Position.Y);
            b[0].Position = pt;
            for (int i = 1; i < b.Count; i++)
            {
                y = b[i].Position.Y;
                offset = b[i - 1].Position.Y - (b[i - 1].BalloonValueSets.Count) * r;
                if (Math.Abs(y - offset) < r)
                {
                    if (b[i - 1].BalloonValueSets.Count == 1) offset -= 0.25*r;
                    y = offset + r / 2;
                }
                pt = I.tg.CreatePoint2d(x, y);
                b[i].Position = pt;
            }
        }

        private void регионToolStripMenuItem_Click(object sender, EventArgs e)
        {
            DrawingDocument drw = Macros.StandardAddInServer.m_inventorApplication.ActiveDocument as DrawingDocument;
            Sheet sh = drw.ActiveSheet;
            Balloon bal = sh.Balloons[1];
            BalloonStyle bs = bal.Style;
            double r = bs.BalloonDiameter;
            List<Balloon> b = sh.Balloons.OfType<Balloon>().ToList();
            b.Sort(new CompY());
            alignY(r, b);
            b.Sort(new CompX(r));
            alignX(r, b);
        }

        private void структураВXMLToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (Macros.StandardAddInServer.m_inventorApplication.ActiveDocument.DocumentType == DocumentTypeEnum.kAssemblyDocumentObject)
            {
                AssemblyDocument doc = Macros.StandardAddInServer.m_inventorApplication.ActiveDocument as AssemblyDocument;
                XMLDoc xmlDoc = new XMLDoc(System.IO.Path.GetFileNameWithoutExtension(doc.FullFileName) + ".xml", "head");
                XElement el = dataFromDoc(doc as Document, "1");
                BOM bom = doc.ComponentDefinition.BOM;
                //if (bom.StructuredViewEnabled == false) bom.StructuredViewEnabled = true;
                if (bom.StructuredViewFirstLevelOnly == true) bom.StructuredViewFirstLevelOnly = false;
                BOMView bomView = doc.ComponentDefinition.BOM.BOMViews[1];
                //sortBOM(bomView, doc);
                //bomView.Sort("");
                BOMToXML(el, bomView.BOMRows);
                xmlDoc.El.Add(el);
                xmlDoc.remove("DecNumber", "");
                XMLDoc.sortRec(xmlDoc.El.Element("Assembly"), "DecNumber", "Name");
                xmlDoc.save();
            }
        }

        public XElement BOMToXML(XElement el, BOMRowsEnumerator rows)
        {
            foreach (BOMRow item in rows)
            {
                Document doc = item.ComponentDefinitions[1].Document as Document;
                XElement elem = dataFromDoc(doc);
                if (elem != null) el.Add(elem); 
                if (item.ChildRows != null && elem != null)
                {
                    BOMToXML(elem, item.ChildRows);
                }
            }
            return el;
        }

        public XElement dataFromDoc(Document doc, string typ = "")
        {
            PartComponentDefinition pCompDef;
            AssemblyComponentDefinition aCompDef;
            XElement data = null; Property prop;
            string name = "", decNumber = "", sketch = "", basePart = "";
            prop = u.getProp(doc, "Description");
            if (prop != null) name = prop.Value.ToString();
            prop = u.getProp(doc, "DecNumber");
            if (prop != null) decNumber = prop.Value.ToString();
            if (typ != "")
            {
                prop = ut.getProp(doc, "Type");
                if (prop != null) typ = prop.Value.ToString();
            }
            if (doc.DocumentType == DocumentTypeEnum.kPartDocumentObject)
            {
                prop = u.getProp(doc, "Part Number");
                if (prop.Value.ToString() == "") return null;
                pCompDef = (doc as PartDocument).ComponentDefinition;
                IEnumerable<char> nameSketch = pCompDef.Sketches.OfType<PlanarSketch>().Where(s => s.HasReferenceComponent == true).SelectMany(sm => sm.Name + ";");
                sketch = new string(nameSketch.ToArray());
                data = new XElement("Part", new XAttribute("Name", name), new XAttribute("DecNumber", decNumber), new XAttribute("Sketch", sketch));
                nameSketch = pCompDef.Parameters.OfType<UserParameter>().Where(p => p.InUse).SelectMany(sm => sm.Name + ";");
                sketch = new string(nameSketch.ToArray());
                if (sketch != "") data.Add(new XAttribute("P", sketch));
                if (pCompDef is SheetMetalComponentDefinition)
                {
                    data.Add(new XAttribute("SheetMetalStyle", (pCompDef as SheetMetalComponentDefinition).ActiveSheetMetalStyle.Name));
                }
                if (doc.ReferencedDocumentDescriptors.Count == 1)
                {
                    basePart = System.IO.Path.GetFileName(InvDoc.u.referendedDocDesc(doc).FullDocumentName);
                    data.Add(new XAttribute("BasePart", basePart));
                }
            }
            else if (doc.DocumentType == DocumentTypeEnum.kAssemblyDocumentObject)
            {
                aCompDef = (doc as AssemblyDocument).ComponentDefinition;
                data = new XElement("Assembly", new XAttribute("Name", name), new XAttribute("DecNumber", decNumber));
                if (typ != "") data.Add(new XAttribute("Type", typ));
            }
            return data;
        }

        private void создатьДеталиToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Document doc = Macros.StandardAddInServer.m_inventorApplication.ActiveDocument;
            string name = ut.OFD(System.IO.Path.GetDirectoryName(doc.FullDocumentName), "XML files(*.xml)|*.xml");
            XMLDoc xdoc = new XMLDoc(name, "head");
            XElement el = xdoc.Doc.Descendants("Assembly").FirstOrDefault();
            if (el != null)
            {
                string typ = el.Attribute("Type").Value;
                Macros.StandardAddInServer.m_inventorApplication.SilentOperation = true;
                create(el, typ);
                HashSet<string> docs = new HashSet<string>();
                foreach (var item in el.Elements())
                {
                    docs.Add(m_Parts.path + ut.nameForSave(item, typ));
                }
                createAssembly(el, typ, docs);
                foreach (_Document item in Macros.StandardAddInServer.m_inventorApplication.Documents)
                {
                    if (item.Open == false)
                    {
                        item.ReleaseReference();
                        item.Close();
                    }
                }
                Macros.StandardAddInServer.m_inventorApplication.SilentOperation = false;
                //Type t = typeof(sender); System.Reflection.FieldInfo fi = t.GetField("Name"); System.Reflection.MemberInfo mi = t.GetProperty
            }
        }

        public HashSet<string> create(XElement el, string typ)
        {
            HashSet<string> docs = new HashSet<string>();
            foreach (var item in el.Elements())
            {
                if (item.HasElements) docs = create(item, typ);
                if (item.Name == "Part") docs.Add(createPart(item, typ));
                if (item.Name ==  "Assembly") docs.Add(createAssembly(item, typ, docs));
            }
            return docs;
        }

        public DerivedPartMirrorPlaneEnum getMirror(XElement el)
        {
            string name = "Mirror";
            if (el.Attribute(name) != null && el.Attribute(name).Value != "")
            {
                name = el.Attribute(name).Value;
                switch (name)
                {
                    case "XY":
                        return DerivedPartMirrorPlaneEnum.kDerivedPartMirrorPlaneXY;
                    case "XZ":
                        return DerivedPartMirrorPlaneEnum.kDerivedPartMirrorPlaneXZ;
                    case "YZ":
                        return DerivedPartMirrorPlaneEnum.kDerivedPartMirrorPlaneYZ;
                    default:
                        return DerivedPartMirrorPlaneEnum.kDerivedPartNoMirrorPlane;
                }
            }
            else return DerivedPartMirrorPlaneEnum.kDerivedPartNoMirrorPlane;
        }

        public string createPart(XElement el, string typ)
        {
            string fileName = ut.nameForSave(el, typ);
            if (System.IO.File.Exists(m_Parts.path + fileName)) return m_Parts.path + fileName;
            PartDocument doc = m_Parts.addPrtDoc();
            string name = el.Attribute("Name").Value, decNumber = el.Attribute("DecNumber").Value, bp = "", skeches = "";
            ut.partNumber(doc as Document, decNumber, typ);
            ut.addProp(doc as Document, "Description", name);
            if (el.Attribute("Sketch") != null) skeches = el.Attribute("Sketch").Value;
            if (el.Attribute("BasePart") != null && (bp = el.Attribute("BasePart").Value) != "")
            {
                string namePar = "";
                if (el.Attribute("P") != null)
                {
                    namePar = el.Attribute("P").Value;
                    namePar = (namePar.IndexOf(";") == -1) ? namePar : namePar.Remove(namePar.IndexOf(';'));
                }
                DerivedPartMirrorPlaneEnum pln = getMirror(el);
                doc = m_Parts.derivedDoc(m_Parts.path + bp, skeches, doc, namePar, pln);
                if (doc.ComponentDefinition is SheetMetalComponentDefinition && el.Attribute("SheetMetalStyle") != null)
                {
                    SheetMetalComponentDefinition smcd = doc.ComponentDefinition as SheetMetalComponentDefinition;
                    Point pt = I.tg.CreatePoint(0,0,ut.convToDouble(smcd.Thickness.Value.ToString()));
                    object ob = Parts.addBaseFeature(smcd, name);
                    string cond = "";
                    if (el.Attribute("Cond") != null) cond = el.Attribute("Cond").Value;
                    if (ob != null && namePar != "" && cond != "")
                        Parts.addFlange(smcd, namePar, cond);
                    else if (ob == null && namePar != "")
                    {
                        Parts.addContourFlange(smcd, namePar, name);
                    }
                    SheetMetalStyle sms = smcd.SheetMetalStyles.OfType<SheetMetalStyle>().FirstOrDefault(e => e.Name == el.Attribute("SheetMetalStyle").Value);
                    if (sms != null) sms.Activate();
                }
                foreach (PlanarSketch ps in doc.ComponentDefinition.Sketches)
                {
                    ps.Visible = false;
                }
            }
            
            doc.SaveAs(m_Parts.path + fileName, false);
            return m_Parts.path + fileName;
            //doc.ReleaseReference();
            //doc.Close();
        }

        public string createAssembly(XElement el, string typ, HashSet<string> docs)
        {
            string fileName = ut.nameForSave(el, typ);
            if (System.IO.File.Exists(m_Parts.path + fileName)) return m_Parts.path + fileName;
            AssemblyDocument doc = m_Parts.addAsmDoc();
            string name = el.Attribute("Name").Value, decNumber = el.Attribute("DecNumber").Value;
            ut.partNumber(doc as Document, decNumber, typ);
            ut.addProp(doc as Document, "Description", name);
            foreach (var item in docs)
            {
                doc.ComponentDefinition.Occurrences.AddUsingiMates(item); 
            }
            //string fileName = ut.nameForSave(doc as Document);
            doc.SaveAs(m_Parts.path + fileName, false);
            doc.ReleaseReference();
            doc.Close();
            return m_Parts.path + fileName;
        }

        private void загрузитьXMLToolStripMenuItem_Click(object sender, EventArgs e)
        {
            string name = "";
            this.Hide();
                name = ut.OFD(Macros.StandardAddInServer.m_inventorApplication.ActiveDocument.path(), "XML(*.xml)|*.xml", false);
                this.Show();
                if (name == null) return;
                PartsBtn.xDoc = new XMLDoc(name, "head");
        }

        private void ручнаяГибкаToolStripMenuItem_Click(object sender, EventArgs e)
        {
            int offsetY = 30, offsetX = 10; Label lbl;
            Form f = new Form();
            f.Height = 150; f.Width = 400; f.WindowState = FormWindowState.Normal; f.Text = "Ручная гибка"; f.StartPosition = FormStartPosition.CenterScreen;
            f.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink; f.AutoSize = true;
            System.Drawing.Point insPt = new System.Drawing.Point(5, 5);
            InterfaceDll.Lbl lbs = new Lbl(offsetX, offsetY, 100, 15, insPt, f, "Ширина перемычки  ");
            InterfaceDll.CB cbs = new CB(offsetX, offsetY, 200, 15, insPt, f);
            cbs.position(lbs.last(), true);
            cbs.last().Items.AddRange(new[] { "1", "2", "3", "4" });
            lbs.add("Отступ", Y: true);
            cbs.add("", new[] { "2" }, Y: true).Text = "2";
            lbs.add("Кол-во перемычек", Y: true);
            cbs.add("", new[] { "3" }, Y: true);
            InterfaceDll.Btn btns = new Btn(offsetX, offsetY, 100, 20, insPt, f, hand_bend, "Добавить");
            btns.center(cbs.last(), offsetY + 5);
            f.Show();
            f.FormClosing += f_FormClosing;
        }

        void f_FormClosing(object sender, FormClosingEventArgs e)
        {
            Form f = sender as Form;
            Macros.StandardAddInServer.m_inventorApplication.ActiveDocument.Update();
            PartsBtn.cc.Close();
        }

        void hand_bend(object sender, EventArgs e)
        {
            Form f = (Form)((System.Windows.Forms.Button)sender).Parent;
            f.Hide();
            PartDocument doc = Macros.StandardAddInServer.m_inventorApplication.ActiveDocument as PartDocument;
            if (doc.ActivatedObject is FlatPattern)
            {
                FlatPattern fp = doc.ActivatedObject as FlatPattern;
                PlanarSketch ps = fp.Sketches.Add(fp.TopFace);
                CommandManager cmd = Macros.StandardAddInServer.m_inventorApplication.CommandManager;
                Edge ed = cmd.Pick(SelectionFilterEnum.kPartEdgeLinearFilter, "Выберите линию сгиба") as Edge;
                SketchPoint pt1 = ps.AddByProjectingEntity(ed.StartVertex) as SketchPoint,
                    pt2 = ps.AddByProjectingEntity(ed.StopVertex) as SketchPoint;
                var ie = f.Controls.OfType<ComboBox>();
                int count = (int)ut.convToDouble(ie.ElementAt(2).Text);
                double lstart = ut.convToDouble(ie.ElementAt(1).Text)/10;
                double l = ut.convToDouble(ie.ElementAt(0).Text)/10;
                Parameter pStart = null, p = null;
                SketchLine sl = addLine(pt1, pt1.Geometry.VectorTo(pt2.Geometry), lstart, l, ref pStart);
                for (int i = 1; i < count; i++)
                {
                    sl = addLine(sl, l, ref p); 
                }
                ps.DimensionConstraints.AddTwoPointDistance(pt2,sl.EndSketchPoint, DimensionOrientationEnum.kAlignedDim, ut.midPt(sl.EndSketchPoint.Geometry, pt2.Geometry, 0, 2)).Parameter.Expression = pStart.Name;
            }
            f.Show();
        }

        public SketchLine addLine(SketchPoint sp, Vector2d vec, double l1, double l2, ref Parameter param)
        {
            PlanarSketch ps = sp.Parent as PlanarSketch;
            Point2d start, end;
            vec.Normalize();
            vec.ScaleBy(l1);
            start = sp.Geometry;
            start.TranslateBy(vec);
            end = start.Copy();
            vec.Normalize(); vec.ScaleBy(l2);
            end.TranslateBy(vec);
            SketchLine sl = ps.SketchLines.AddByTwoPoints(start, end);
            ps.GeometricConstraints.AddCoincident((SketchEntity)sp, (SketchEntity)sl);
            param = ps.DimensionConstraints.AddTwoPointDistance(sp, sl.StartSketchPoint, DimensionOrientationEnum.kAlignedDim, ut.midPt(sp.Geometry, sl.StartSketchPoint.Geometry, 0, 2)).Parameter;
            return sl;
        }

        public SketchLine addLine(SketchLine startsl, double l1, ref Parameter param)
        {
            PlanarSketch ps = startsl.Parent as PlanarSketch;
            Point2d start, end;
            Vector2d vec = startsl.StartSketchPoint.Geometry.VectorTo(startsl.EndSketchPoint.Geometry);
            vec.Normalize();
            vec.ScaleBy(l1);
            start = startsl.EndSketchPoint.Geometry;
            start.TranslateBy(vec);
            end = start.Copy();
            vec.Normalize(); vec.ScaleBy(l1);
            end.TranslateBy(vec);
            SketchLine sl = ps.SketchLines.AddByTwoPoints(start, end);
            TwoPointDistanceDimConstraint constr = ps.DimensionConstraints.AddTwoPointDistance(startsl.EndSketchPoint, sl.StartSketchPoint, 
                DimensionOrientationEnum.kAlignedDim, ut.midPt(startsl.EndSketchPoint.Geometry, sl.StartSketchPoint.Geometry, 0, 2));
            if (param == null)
                param = constr.Parameter;
            else
                constr.Parameter.Expression = param.Name;
            ps.GeometricConstraints.AddEqualLength(startsl, sl);
            ps.GeometricConstraints.AddCollinear((SketchEntity)startsl, (SketchEntity)sl);
            return sl;
        }

        private void заменитьToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void спецToolStripMenuItem_Click(object sender, EventArgs e)
        {
            string fn = u.WOFD(@"C:\ProgramData\Autodesk\Inventor Addins\", "XML (*.xml)|*.xml", new string[]{file.p(I.aDoc().FullFileName)});
            if (fn == "") { Macros.StandardAddInServer.xml = null; return; }
            Macros.StandardAddInServer.xml = new MyXML(fn);
        }

        private void добавитьИсполнениеToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Document doc = Macros.StandardAddInServer.m_inventorApplication.ActiveDocument;
            List<string> asms = TableInv.getAsms(doc.FullFileName); asms.Sort();
            int isp = 1;
            string name = asms[asms.Count-1]; int l = name.Length;
            if (name[l-7] == '^')
            {
                isp = int.Parse(name.Substring(l-6, 2))+1;
            }
            Form f = new Form();
            InterfaceDll.MyLabel mylbl1 = new MyLabel(); int offsetY = 20;
            InterfaceDll.MyTextBox mytxt1 = new MyTextBox();
            InterfaceDll.MyComboBox mycb = new MyComboBox();
            f.Height = 115; f.Width = 400; f.WindowState = FormWindowState.Normal; f.Text = "Добавить исполнение"; f.StartPosition = FormStartPosition.CenterScreen;
            System.Drawing.Point insPt = new System.Drawing.Point(5, 5);
            Label lbl1 = mylbl1.addLabel("Примечание", insPt, 100, 15);
            f.Controls.Add(lbl1);
            insPt.Y += offsetY;
            lbl1 = mylbl1.addLabel("Номер исполнения", insPt, 100, 15);
            insPt.Y += offsetY;
            f.Controls.Add(lbl1);
            lbl1 = mylbl1.addLabel("Тип", insPt, 100, 15);
            f.Controls.Add(lbl1);
            insPt.X += 150; insPt.Y -= offsetY*2;
            System.Windows.Forms.ComboBox cb = mycb.addComboBox("", insPt, 200, 15, new string[] { "1 кВт", "3 кВт", "4 кВт", "6 кВт", "10 кВт", "12 кВт", "18 кВт", "24 кВт", 
                "27 кВт", "36 кВт", "54 кВт"});
            f.Controls.Add(cb);
            insPt.Y += offsetY;
            System.Windows.Forms.TextBox txt2 = mytxt1.addTextBox(isp.ToString("00"), insPt, 200, 15);
            insPt.Y += offsetY;
            cb = mycb.addComboBox("", insPt, 200, 15, new string[] { "A", "E", "W"});
            f.Controls.Add(cb); f.Controls.Add(txt2);
            MyButton myBtn = new MyButton();
            insPt.Y += offsetY; insPt.X = 200 - 50;
            System.Windows.Forms.Button btn = myBtn.addButton("Добавить", insPt, 100, 20);
            f.Controls.Add(btn);
            btn.Click += btn_ClickIsp;
            f.Show();
        }

        static public string perf(string desc, string pn, ref string type, ref string decnumber, ref string note, string EWA, ref string count)
        {
            if (pn != "")
            {
                var spl = pn.Split('.');
                if (spl.Count() == 3)
                {
                    type = spl[0];
                    decnumber = spl[1] + "." + spl[2];
                }
            }
            else return desc;
            Regex regex = new Regex(@"\b(\d+)"), regex1 = new Regex(@"(КЭВ-)(\d*)(\w)(\d)(\d)(\d*)(\w)");
            Match match = regex.Match(note), match1 = regex1.Match(type);
            string emblem, t, decu = "";
            if (match1 == null) return desc;
            t = match1.Groups[3].Value; emblem = match1.Groups[1].Value;
            if (t == "П" || t == "С" || t == "C")
            {
                string EWA1 = match1.Groups[7].Value;
                if (EWA != "" && EWA1 != EWA)
                {
                    count = "";
                }
                else if (EWA == "") EWA = EWA1;
                if (EWA == "W" || EWA == "A")
                {
                    decu = match1.Groups[4].Value + 1 + match1.Groups[6].Value; 
                }
                else
                {
                    decu = match1.Groups[4].Value + 0 + match1.Groups[6].Value;
                }
                type = emblem + match1.Groups[2].Value + t + decu + EWA;
            }
            emblem = match1.Groups[1].Value;
            note = match.Groups[1].Value;
            if (desc.StartsWith("Завеса КЭВ-"))
            {
                desc = "Завеса " + emblem + note + t + decu + EWA;
            }
            else if (desc.StartsWith("Тепловентилятор"))
            {
                desc = "Тепловентилятор " + emblem + note + t + decu + EWA;
            }
            return desc;
        }

        void btn_ClickIsp(object sender, EventArgs e)
        {
            Form f = (Form)((System.Windows.Forms.Button)sender).Parent;
            Macros.StandardAddInServer.m_inventorApplication.SilentOperation = true;
            Document doc = Macros.StandardAddInServer.m_inventorApplication.ActiveDocument;
            string pn = ut.getProp(doc, "Part Number").Value.ToString(), desc = ut.getProp(doc, "Description").Value.ToString();
            string type = "", decnumber = "";
            string note = f.Controls.OfType<ComboBox>().ElementAt(0).Text;
//             Regex r = new Regex(@"^\d*$");
//             if (r.IsMatch(note))
//                 note += " кВт";
            string EWA = f.Controls.OfType<ComboBox>().ElementAt(1).Text;
            string count = f.Controls.OfType<System.Windows.Forms.TextBox>().ElementAt(0).Text;
            desc = perf(desc, pn, ref type, ref decnumber,ref note, EWA, ref count);
            count = f.Controls.OfType<System.Windows.Forms.TextBox>().ElementAt(0).Text;
            string n = doc.FullFileName;
            if (count != "")
                n = System.IO.Path.Combine(System.IO.Path.GetDirectoryName(n), type + "." + decnumber + "-" + count + " (" + desc + ")" + "^" + count
                 + System.IO.Path.GetExtension(n));
            else
            {
                n = System.IO.Path.Combine(System.IO.Path.GetDirectoryName(n), type + "." + decnumber + " (" + desc + ")" 
                 + System.IO.Path.GetExtension(n));
            }
            if (!System.IO.File.Exists(n)) doc.SaveAs(n, true);
            NameValueMap nvmOptions = I.objs.CreateNameValueMap();
            nvmOptions.Add("SkipAllUnresolvedFiles", true);
            Document docAdd = Macros.StandardAddInServer.m_inventorApplication.Documents.OpenWithOptions(n,nvmOptions,false);
            if (docAdd.DocumentType == DocumentTypeEnum.kAssemblyDocumentObject)
            {
                AssemblyComponentDefinition compDef = (docAdd as AssemblyDocument).ComponentDefinition;
                compDef.RepresentationsManager.DesignViewRepresentations[2].Activate();
            }
            ut.addProp(docAdd, "Description", desc); ut.addProp(docAdd, "Type", type);
            if (count != "")
            ut.addProp(docAdd, "DecNumber", decnumber + "-" + count);
            else ut.addProp(docAdd, "DecNumber", decnumber);
            ut.addProp(docAdd, "Comments", note + " кВт"); ut.addProp(docAdd, "Part Number", "=<Type>.<DecNumber>");
            if (EWA == "W" && System.IO.File.Exists(n = System.IO.Path.Combine(System.IO.Path.GetDirectoryName(n) + "\\W.xml")))
            {
                Parts.removeReplace(docAdd, n);
            }
            else if(EWA == "A" && System.IO.File.Exists(n = System.IO.Path.Combine(System.IO.Path.GetDirectoryName(n) + "\\A.xml")))
            {
                Parts.removeReplace(docAdd, n);
            }
            docAdd.Save2();
            docAdd.ReleaseReference();
            Macros.StandardAddInServer.m_inventorApplication.SilentOperation = false;
            f.Close();
        }

        private void обновитьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //Document doc = .ActiveDocument;
            foreach (Document doc in Macros.StandardAddInServer.m_inventorApplication.Documents.VisibleDocuments)
            {
                update(doc);
            }
        }

        public static void update(Document doc/*, bool update = false*/)
        {
            if (!doc.Dirty) return;/* bool f = false;*/
            foreach (Document item in doc.ReferencedFiles)
            {
                if (item.RequiresUpdate)
                {
                    item.Update2();
                    item.Save2();
                    //if (!update) f = true;
                }
            }
            if (doc.DocumentType == DocumentTypeEnum.kDrawingDocumentObject)
            {
                foreach (Sheet sh in (doc as DrawingDocument).Sheets)
                {
                    if (sh.Status == DrawingSheetStatusBits.kUpToDateDrawingSheet)
                    {
                        sh.Update();
                    }

//                     foreach (DrawingView dv in sh.DrawingViews)
//                     {
//                         if (dv.UpToDate == true) sh.Update();
//                     }
                }
            }
            doc.Update2();  
            doc.Save2();
        }

        private void поискТестToolStripMenuItem_Click(object sender, EventArgs e)
        {
            AssemblyDocument asm = Macros.StandardAddInServer.m_inventorApplication.ActiveDocument as AssemblyDocument;
            ObjectsEnumerator objs; object loc;
            SelectionFilterEnum[] enu = new SelectionFilterEnum[]{SelectionFilterEnum.kAllCircularEntities};
            foreach (ComponentOccurrence occ in asm.ComponentDefinition.Occurrences)
            {
                if (occ.iMateDefinitions.Count != 0 )
                {
                    InsertiMateDefinitionProxy i = occ.iMateDefinitions[1] as InsertiMateDefinitionProxy;
                    EdgeProxy ep = i.Entity as EdgeProxy;
                    Circle c = ep.Geometry as Circle;
                    double r = c.Radius;
                    foreach (Face f in ep.Faces)
                    {
                        if (f.SurfaceType == SurfaceTypeEnum.kPlaneSurface)
                        {
                            Plane pl = f.Geometry as Plane;
                            UnitVector v = pl.Normal;       
                            //asm.ComponentDefinition.FindUsingRay(ep.PointOnEdge, v, r*2, out objs, out loc);
                            objs = asm.ComponentDefinition.FindUsingVector(ep.PointOnEdge, v, enu, true, r, true, out loc);
//                             foreach (Face item in objs)
//                             {
//                                 if (item.GetType() == typeof(FaceProxy))
//                                 {
//                                     string n = "1";
//                                 }
//                             }
                            var col = objs.OfType<EdgeProxy>();
                            foreach (EdgeProxy item in col)
                            {
                                Plane pf = item.Geometry as Plane;
                                if (pf.Normal == v)
                                {
                                    pf = pl;
                                }
                            }
                        }
                    }
                }
            }
        }

        private void удалитьСтандартныеToolStripMenuItem_Click(object sender, EventArgs e)
        {
            AssemblyDocument asm = Macros.StandardAddInServer.m_inventorApplication.ActiveDocument as AssemblyDocument;
            ObjectCollection objs = I.objs.CreateObjectCollection();
            foreach (ComponentOccurrence occ in asm.ComponentDefinition.Occurrences)
            {
                if (occ.ReferencedDocumentDescriptor.FullDocumentName.IndexOf(@"\Content Center Files\") != -1)
                {
                    objs.Add(occ);         
                }
            }
            foreach (ComponentOccurrence item in objs)
            {
                item.Delete(); 
            }
        }

        private void комплектФайловСЧертежамиToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Inventor.Application app = Macros.StandardAddInServer.m_inventorApplication;
            Document doc = app.ActiveDocument;
            app.SilentOperation = true;
            PackAndGoLib.PackAndGoComponent packAndGoComp = new PackAndGoLib.PackAndGoComponent();
            string locName;
            //Document doc = Macros.StandardAddInServer.m_inventorApplication.ActiveDocument;
            locName = doc.path() + "\\";
            locName += doc.name() + "\\";
            if (!System.IO.Directory.Exists(locName)) System.IO.Directory.CreateDirectory(locName);
            PackAndGoLib.PackAndGo packAndGo = packAndGoComp.CreatePackAndGo(app.ActiveDocument.FullDocumentName, locName);

            string[] refFiles = new string[] { };
            string[] refecening = new string[] { };
            string path = System.IO.Path.GetDirectoryName(doc.FullFileName);//app.DesignProjectManager.ActiveDesignProject.WorkspacePath;
            NameValueMap nvm = I.objs.CreateNameValueMap();
            nvm.Add("SkipAllUnresolvedFiles", true);
            foreach (var f in System.IO.Directory.EnumerateFiles(path, "*.idw", System.IO.SearchOption.AllDirectories))
            {
                if (f.EndsWith(".idw") && f.IndexOf("OldVersions") == -1)
                {
                    Macros.StandardAddInServer.m_inventorApplication.Documents.OpenWithOptions(f, nvm, false);
                }
            }
            object refMissFiles = new object();

            // Set the options
            packAndGo.SkipLibraries = true;
            packAndGo.SkipStyles = true;
            packAndGo.SkipTemplates = true;
            packAndGo.CollectWorkgroups = false;
            packAndGo.KeepFolderHierarchy = false;
            packAndGo.IncludeLinkedFiles = true;
                                                                       
            //packAndGo.SearchForReferencedFiles(out refFiles,out refMissFiles);
            //packAndGo.AddFilesToPackage(ref refFiles);
            //packAndGo.SearchForReferencingFiles(ref searchFiles, out refecening, true);
            //packAndGo.AddFilesToPackage(ref refecening);
            HashSet<string> names = new HashSet<string>();
            packFromBOMWithDrw(doc, ref names);
            packAndGo.AddFilesToPackage(names.ToArray());
            packAndGo.CreatePackage(true);
            app.SilentOperation = false;
        }

        public static void packFromBOMWithDrw(Document doc, ref HashSet<string> names)
        {
            foreach (DocumentDescriptor rd in doc.ReferencedDocumentDescriptors)
            {
                if (rd.ReferencedFileDescriptor.LibraryName != null) continue;
                if (rd.ReferencedDocumentType == DocumentTypeEnum.kAssemblyDocumentObject)
                {
                    packFromBOMWithDrw(rd.ReferencedDocument as Document, ref names);
                }

                addwithDrw(rd.ReferencedDocument as Document,ref names);
            }

            addwithDrw(doc,ref names);
        }

        public static void addwithDrw(Document doc, ref HashSet<string> names)
        {
            names.Add(doc.FullFileName);
            foreach (Document item in doc.ReferencingDocuments)
            {
                names.Add(item.FullFileName);
            }
        }

        private void переименоватьСборкуСЧертежамиToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Document doc = I.aDoc();
            renameAsm(doc, true);
        }

        private void переименоватьЧертежиToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Document doc = Macros.StandardAddInServer.m_inventorApplication.ActiveDocument;
            openDrw(doc);
            renameDrw(Macros.StandardAddInServer.m_inventorApplication.Documents.OfType<Document>(), false);
            renameDrw(Macros.StandardAddInServer.m_inventorApplication.Documents.VisibleDocuments.OfType<Document>(), true);
        }
        
        public static void renameDrw(IEnumerable<Document> docs, bool visible)
        {
            XElement el;
            foreach (Document item in docs)
            {
                if (item.DocumentType != DocumentTypeEnum.kDrawingDocumentObject) continue;
                el = new XElement("row");
                if (item.ReferencedDocuments.Count != 0)
                addName(InvDoc.u.referendedDoc(item), el);
                string o = System.IO.Path.GetFileNameWithoutExtension(item.FullDocumentName);
                if (el.Attribute("new") != null && o != System.IO.Path.GetFileNameWithoutExtension(el.Attribute("new").Value))
                {
                    o = item.FullDocumentName;
                    if (visible) 
                    { 
                        item.Save2(); 
                        item.Close(); 
                    }
                    else item.ReleaseReference();
                    string n = System.IO.Path.GetDirectoryName(item.FullDocumentName) + "\\" +
                        System.IO.Path.GetFileNameWithoutExtension(el.Attribute("new").Value) + System.IO.Path.GetExtension(item.FullDocumentName);
                    System.IO.File.Move(o, n);
                }
                if (item.Open) item.Close();
            }
        }

        private void обновитьПозицииToolStripMenuItem_Click(object sender, EventArgs e)
        {
            DrawingDocument m_Drw = Macros.StandardAddInServer.m_inventorApplication.ActiveDocument as DrawingDocument;
            AssemblyDocument m_AsmDoc = InvDoc.u.referendedDoc(m_Drw as Document) as AssemblyDocument;
            foreach (Balloon bal in m_Drw.ActiveSheet.Balloons)
            {
                if (bal.BalloonValueSets.Count == 1)
                Macros.StandardAddInServer.addBalloonValueSet(bal, m_AsmDoc);
            }
        }

        private void текстToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Regex regex = new Regex(@"(\w*\d*), (\w*)");
            DrawingDocument doc = Macros.StandardAddInServer.m_inventorApplication.ActiveDocument as DrawingDocument;
            Sheet sh = doc.ActiveSheet;
            foreach (SketchedSymbol ss in sh.SketchedSymbols)
            {
                if (ss.Name != "Пользовательская") continue;
                SketchedSymbolDefinition d = ss.Definition;
                DrawingSketch s;
                d.Edit(out s);
                foreach (Inventor.TextBox tb in s.TextBoxes)
                {
                    Match m = regex.Match(tb.Text);
                    if (m.Groups[1].Value != "" && m.Groups[2].Value != "")
                    {
                        string format = "<StyleOverride" + @" Bold='true'" + @">" + m.Groups[1].Value + @"</StyleOverride>" + ","
                            + "<StyleOverride" + @" Italic='true'" + @">" + m.Groups[2].Value + @"</StyleOverride>";
                        tb.FormattedText = format; 
                    }
                }
                d.ExitEdit();
            }
        }

        private void dwgToolStripMenuItem_Click(object sender, EventArgs e)
        {
            foreach (DrawingDocument dwg in Macros.StandardAddInServer.m_inventorApplication.Documents.VisibleDocuments)
            {
                dwg.SaveAsInventorDWG(dwg.FullDocumentName.Replace(".idw", ".dwg"),true); 
            }
        }

        private void получитьТекстToolStripMenuItem_Click(object sender, EventArgs e)
        {
            InvDoc.InvDocument<AssemblyDocument> doc = new InvDoc.InvDocument<AssemblyDocument>(Macros.StandardAddInServer.m_inventorApplication.ActiveDocument as AssemblyDocument);
            BOMView view = doc.getBOMView(false,false,"Только детали",true,false);
            string n = u.WOFD(doc.path);
            XMLDoc xml = (n == "") ? new XMLDoc(doc.path + "English.xml", "head") : new XMLDoc(n, "head");
            InvDoc.GetDocs docs = new InvDoc.GetDocs();
            XElement root =  xml.Doc.Root;
            foreach (Document item in docs.docs)
            {
                Property p = u.getProp(item,"Description");
                if (p == null) continue;
                xml.addXElement(root, new string[] { "name", p.Value.ToString(), "value", "" });
            }
//             foreach (BOMRow row in view.BOMRows)
//             {
//                 Property p = InvDoc.util.getProp(row.ComponentDefinitions[1].Document as Document, "Description");
//                 XElement el = xml.Doc.Root;
//                 xml.addXElement(el, new string[] { "name", p.Value.ToString(), "value", "" });
//             }
            xml.save();
        }

        private void обновитьСтилиToolStripMenuItem_Click(object sender, EventArgs e)
        {
            foreach (DrawingDocument dwg in Macros.StandardAddInServer.m_inventorApplication.Documents.VisibleDocuments)
            {
                foreach (Style ds in dwg.StylesManager.Styles)
	            {
                    if (!ds.UpToDate) ds.UpdateFromGlobal();
	            }
                foreach (Sheet s in dwg.Sheets)
                {
                    foreach (HoleThreadNote htn in s.DrawingNotes.HoleThreadNotes)
                    {
                        htn.UsePartUnits = false;
                    }
                    string name = "Таблица сгибов";
                    CustomTable t = s.CustomTables.OfType<CustomTable>().FirstOrDefault(tb => tb.Title == name);

                    if (t != null)
                    {
                        t = u.addCustomTable(t, name);

                        foreach (Row r in t.Rows)
                        {
                            switch (r[2].Value)
                            {
                                case "ВНИЗ":
                                    r[2].Value = "DOWN";
                                    break;
                                case "ВВЕРХ":
                                    r[2].Value = "UP";
                                    break;
                                default:
                                    break;
                            }
                            if (u.convToDouble(r[4].Value) > 0.6)
                            {
                                r[4].Value = u.convert(r[4].Value, 25.4, 3);
                            }
                        }
                    }
                }
            }
        }

        private void перевестиТекстToolStripMenuItem_Click(object sender, EventArgs e)
        {
            string n = u.OFD(Macros.StandardAddInServer.m_inventorApplication.ActiveDocument.path(),"xml|*.xml");
            if (n == "")
            {   
                return;
            }
            XMLDoc xml = new XMLDoc(n, "head");
            Dictionary<string, string> dic = xml.getDictionary(xml.El, "name", "value");
            InvDoc.GetDocs docs = new InvDoc.GetDocs();
            foreach (Document doc in docs.docs)
            {
                 u.translit(doc, "Description", dic);
            }
            docs.rename(dic);
        }

        private void осьСимметрииToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void общиеРазмерыToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ut.action<LinearGeneralDimension>(ut.gets<LinearGeneralDimension>(I.getSheet().DrawingDimensions, d => d.Text.FormattedText.IndexOf("**") == -1),
                a => a.Delete());            
        }

        private void размерыДоГибовToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ut.action<LinearGeneralDimension>(ut.gets<LinearGeneralDimension>(I.getSheet().DrawingDimensions, d => d.Text.FormattedText.IndexOf("**") != -1),
                a => a.Delete()); 
        }

        private void позицииToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ut.action<Balloon>(ut.gets<Balloon>(I.getSheet().Balloons, b => true),
                a => a.Delete()); 
        }
                                                                                            
        private void крепежToolStripMenuItem2_Click(object sender, EventArgs e)
        {
            ut.action<Document>(I.app.Documents.VisibleDocuments, r => removeFasteners(r), f => f.DocumentType == DocumentTypeEnum.kAssemblyDocumentObject);
        }

        static public void removeFasteners(Document doc)
        {
            AssemblyComponentDefinition acd = I.getACD(doc);
            ut.action<FeatureBasedOccurrencePattern>(acd.OccurrencePatterns, p => deletePattern(p));
            ut.action<RectangularOccurrencePattern>(acd.OccurrencePatterns, p => deletePattern(p));
            ut.action<ComponentOccurrence>(acd.Occurrences, oc => oc.Delete(), f => check(f));
        }

        static public bool check(ComponentOccurrence co)
        {
            return co.ReferencedDocumentDescriptor.FullDocumentName.IndexOf(@"\Content Center Files\") != -1 && ut.getLenght(co.RangeBox) < 2.5;
        }

        static public void delete<T>(T f)
        {
            if (f == null) return;
            InvDoc.Reflect.runMethod<T, object>(f, "Delete", null);
        }

        static public void deletePattern(FeatureBasedOccurrencePattern pat)
        {
            ObjectCollection colNew = I.objs.CreateObjectCollection();
            foreach (ComponentOccurrence co in pat.ParentComponents)
            {
                string name = ut.getName(co.Name, ':', 0);
                if (check(co) && co.Name.IndexOf(name) == -1) colNew.Add(co);
            }
            if (colNew.Count != 0) pat.ParentComponents = colNew;
            else pat.Delete();
        }

        static public void deletePattern(RectangularOccurrencePattern pat)
        {
            ObjectCollection colNew = I.objs.CreateObjectCollection();
            foreach (ComponentOccurrence co in pat.ParentComponents)
            {
                string name = ut.getName(co.Name, ':', 0);
                if (check(co) && co.Name.IndexOf(name) == -1) colNew.Add(co);
            }
            if (colNew.Count == 1) pat.ParentComponents = colNew;
            else del.Add(pat);
        }

        private void переименоватьСборкуToolStripMenuItem_Click(object sender, EventArgs e)
        {
            renameAsm(I.app.ActiveDocument);
        }

        private void заменитьИмяКПToolStripMenuItem_Click(object sender, EventArgs e)
        {
            XMLDoc xdoc = new XMLDoc(@"C:\ProgramData\Autodesk\Inventor Addins\rimate.xml", "head");
            Document doc = I.app.ActiveDocument;
            ut.action<Document>(I.app.Documents, f => replaceImate(xdoc, f));
            //replaceImate(xdoc, doc);
            xdoc.save();
        }

        public static void replaceImate(XMLDoc xdoc, Document doc)
        {
            SheetMetalComponentDefinition compDef = I.getSMCD(doc);
            if (compDef != null)
            {
                ut.action<iMateDefinition>(compDef.iMateDefinitions, f => check(xdoc,f));
            }
            AssemblyComponentDefinition acompDef = I.getACD(doc);
            if (compDef != null)
            {
                ut.action<iMateDefinition>(compDef.iMateDefinitions, f => check(xdoc, f));
            }
        }
        public static void check(XMLDoc xdoc, iMateDefinition i)
        {
            if (i.Name.StartsWith("i")) return;
            XElement el = xdoc.find("old",i.Name);
            if (el == null)
                xdoc.addXElement("imate", new Dictionary<string, string>() { { "old", i.Name }, { "new", "" } });
            else if (el.Attribute("new") != null && el.Attribute("new").Value != "")
                i.Name = el.Attribute("new").Value;

        }

        private void обновитьКрепежToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ut.action<AssemblyDocument>(I.app.Documents.VisibleDocuments, d => updateFasteners(d), f => f.DocumentType == DocumentTypeEnum.kAssemblyDocumentObject);
        }

        static public void updateFasteners(AssemblyDocument asm)
        {
            //ContentOp c = new ContentOp();
            removeFasteners(asm as Document);
            addFasteners(asm, new ContentOp());
        }

        private void крепежToolStripMenuItem3_Click(object sender, EventArgs e)
        {
            ut.action<AssemblyDocument>(I.app.Documents.VisibleDocuments, d => addFasteners(d, new ContentOp()), f => f.DocumentType == DocumentTypeEnum.kAssemblyDocumentObject);
        }

        private void открытьПапкуToolStripMenuItem_Click(object sender, EventArgs e)
        {
            string p = I.aDoc().FullFileName;
            file.open(p);
            this.Close();
        }

        private void кластерToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Cluster cl = new Cluster();
            cl.draw();
            cl.txt();
        }

        private void габаритыToolStripMenuItem_Click(object sender, EventArgs e)
        {
            string fn = I.aDoc().FullFileName;
            fn = fn.Replace(".ipt", ".txt");
            fn = fn.Replace(".iam", ".txt");
            using (System.IO.StreamWriter sw = System.IO.File.CreateText(fn))
            {
                Document doc = I.aDoc();
                Box box = null;
                if (doc.DocumentType == DocumentTypeEnum.kPartDocumentObject)
                {
                    SheetMetalComponentDefinition smcd = I.getSMCD();
                    box = smcd.RangeBox;
                }
                else if (doc.DocumentType == DocumentTypeEnum.kAssemblyDocumentObject)
                {
                    AssemblyComponentDefinition acd = I.getACD(doc);
                    box = acd.RangeBox;
                }
                Vector v = box.MinPoint.VectorTo(box.MaxPoint);
                sw.Write("Габариты: \t" + (v.X*10).ToString("##.##") + " мм\n\t\t\t" + (v.Y*10).ToString("##.##") + " мм\n\t\t\t" + (v.Z*10).ToString("##.##") + " мм.\n");
            }
        }

        //private void добавитьШероховатостьToolStripMenuItem_Click(object sender, EventArgs e)
        //{
        //    DrawingDocument drw = (DrawingDocument)Macros.StandardAddInServer.m_inventorApplication.ActiveDocument;
        //    this.Hide();
        //    DrawingView dv = (DrawingView)Macros.StandardAddInServer.m_inventorApplication.CommandManager.Pick(SelectionFilterEnum.kDrawingViewFilter, "Выберите вид");
        //    if (dv != null)
        //    {
        //        DrawingCurve dc1 = null, dc2 = null;
        //        DrawingCurve dc = Drawings.surfCurve(dv, ref dc1, ref dc2);
        //        LinearGeneralDimension dim = Drawings.addDim(dc, -15);
        //        dim.Text.FormattedText = dim.Text.FormattedText + "*";
        //        double val = 0.5; Vector2d vec;
        //        if (dc1 != null)
        //        {
        //            vec = dc2.MidPoint.VectorTo(dc1.MidPoint); vec.Normalize(); vec.ScaleBy(0.1);
        //            Drawings.addSurfaceTextureSymbol(dv, dc1, val,vec);
        //        }
        //        if (dc2 != null)
        //        {
        //            vec = dc1.MidPoint.VectorTo(dc2.MidPoint); vec.Normalize(); vec.ScaleBy(0.1);
        //            Drawings.addSurfaceTextureSymbol(dv, dc2, val, vec);
        //        }
        //        //Drawings.addSurfaceTextureSymbol(dim, 0.3, true);
        //        //Drawings.addSurfaceTextureSymbol(dim, 0.3, false);
        //    }
        //    this.Close();
        //}
    }

    internal class PartsBtn : Button
    {
        public static Parts m_Parts;
        public static XMLDoc xDoc;
        public static XMLDoc param;
        public static XMLDoc xDocSheet;
        public static CreateComponent cc;
        public static PartDocument Doc;
        public Inventor.Document pDoc { get; set; }
        public static Parts getParts
        {
            get
            {
                return m_Parts;
            }
        }

        #region "Methods"
        public PartsBtn(string displayName, string internalName, string clientId, string description, string tooltip, System.Drawing.Icon standardIcon, System.Drawing.Icon largeIcon)
            : base(displayName, internalName, clientId, description, tooltip, standardIcon, largeIcon)
        {
            string file = @"C:\ProgramData\Autodesk\Inventor Addins\Parameters.xml";
            if (System.IO.File.Exists(file))
            param = new XMLDoc(file, "head");
        }

        protected override void ButtonDefinition_OnExecute(NameValueMap context)
        {
            
            try {
               Macros.StandardAddInServer.forms.Add(cc);
               if (Macros.StandardAddInServer.activeteForm()) new CreateComponent(I.aDoc());
                   /*System.Windows.Forms.Application.Run(cc = new CreateComponent(InventorApplication.ActiveDocument));*/
            }
            //catch (InvalidOperationException ex)
            //{
            //    cc.Show();
            //}
            catch(Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.ToString());
            }
    }


        #endregion
    }

    public class Parts : InvDocument<Document>
    {
        //Inventor.Application InvApp;
        InvDoc.XML n;
        new string str;
        string name = "";
        bool openFlag;
        CreateComponent cc;
        PlanarSketch ps;
        SketchPoint sp;

        public SketchPoint Sp
        {
            get { return sp; }
            set { sp = value; }
        }
        WorkPlane wp;

        public WorkPlane Wp
        {
            get { return wp; }
            set { wp = value; }
        }
        SelectSet ss;       

        public SelectSet Ss
        {
            get { return ss; }
            set
            {
                if (value.Count == 1 && ss[1] is SketchPoint) sp = value[1] as SketchPoint;
                else if (value.Count == 2 && value[1] is SketchPoint)
                {
                    sp = value[1] as SketchPoint;
                    if (value[2] is WorkPlane) wp = value[2] as WorkPlane;
                }
            }
        }

        public Parts(Inventor.Document doc): base (doc)
        {
        }

        public void create()
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.InitialDirectory = path;
            ofd.Filter = "Описание создаваемых файлов (.xml)|*.xml";
            ofd.ShowDialog();
            string filePath = ofd.FileName;
            if (filePath == "")
                filePath = (System.IO.File.Exists(path + "\\" + "Parts.xml")) ? path + "\\" + "Parts.xml" : @"C:\ProgramData\Autodesk\Inventor Addins\Parts.xml";
            n = new InvDoc.XML(filePath);

            List<XMLData> data = new List<XMLData>();
            data = n.ReadXML("Parts");
            if (libraryPath.Count == 0)
                libraryPath = null;
            withElem<XMLData>(data, addP);
        }

        private string union(string[] lst)
        {
            string ret = "";
            foreach (var item in lst)
            {
                ret += item + "$";
            }
            ret = ret.Remove(ret.Length - 1, 1);
            return ret;
        }

        public string replaceLib(string data)
        {
            int i = 0;
            string[] tmpstr = data.Split('$');
            foreach (var item in tmpstr)
            {
                if (item.StartsWith("#"))
                {
                    XMLDoc lib = new XMLDoc(@"C:\ProgramData\Autodesk\Inventor Addins\ContentCenter.xml", "Content");
                    string[] tmpspl = item.Split('#');
                    XElement el = null;
                    if (tmpspl.Length == 2)
                    {
                        el = lib.getXElement(tmpspl[1], "Description", "TableRow");
                    }
                    else if (tmpspl.Length == 3)
                    {
                        el = lib.getXElement(tmpspl[1], "Description", tmpspl[2], "Extra", "TableRow");
                    }
                    tmpstr.SetValue(el.Attribute("Id").Value, i);
                }
                i++;
            }
            return union(tmpstr);
        }

        private bool fileExist(string n)
        {
            foreach (string lp in libraryPath)
            {
                if (System.IO.File.Exists(lp + n))
                {
                    name = lp + n;
                    openFlag = true;
                    return true;
                }
            }
            openFlag = false;
            return false;
        }

        public static void replaceReference(DrawingDocument doc, XMLDoc xDoc)
        {
            FileDescriptor fd;                     

            foreach (DocumentDescriptor dd in doc.ReferencedDocumentDescriptors)
            {
                fd = dd.ReferencedFileDescriptor;
                string newName = xDoc.find(xDoc.El, "old", "new", fd.FullFileName);
                if (newName != "" && System.IO.File.Exists(newName) && newName != fd.FullFileName)
                {
                    fd.ReplaceReference(newName);
                }
            }
            doc.Update2(false);
            doc.Save2();
        }

        public static void replaceReference(PartDocument doc, XMLDoc xDoc)
        {
            PartComponentDefinition compDef = doc.ComponentDefinition;
            if (doc.ComponentDefinition.ReferenceComponents.DerivedPartComponents.Count != 0)
            {
                FileDescriptor fd = compDef.ReferenceComponents.DerivedPartComponents[1].ReferencedDocumentDescriptor.ReferencedFileDescriptor;
                string newName = xDoc.find(xDoc.El, "old", "new", fd.FullFileName);
                if (newName != "" && System.IO.File.Exists(newName) && newName != fd.FullFileName)
                {
                    fd.ReplaceReference(newName);
                    doc.Update2(false);
                    doc.Save2();
                }
            }
        }

        public static void replaceReference(AssemblyDocument doc, XMLDoc xDoc)
        {
            foreach (DocumentDescriptor item in doc.ReferencedDocumentDescriptors)
            {
                if (item.ReferenceMissing == true)
                {
                    string newName = xDoc.find(xDoc.El, "old", "new", item.FullDocumentName);
                    if (newName != "" && System.IO.File.Exists(newName) && newName != item.FullDocumentName)
                    {
                        item.ReferencedFileDescriptor.ReplaceReference(newName);
                        
                    }
                }
            }
            doc.Update2(false);
            doc.Save2();
        }

        public static void replaceReference(Document doc, XMLDoc xDoc)
        {
            foreach (DocumentDescriptor item in doc.ReferencedDocumentDescriptors)
            {
                 //if (doc.DocumentType == DocumentTypeEnum.kDrawingDocumentObject)
                 //{
                    string newName = xDoc.find(xDoc.El, "old", "new", ut.nameUtil(item.FullDocumentName, true));
                    if (newName != "" && System.IO.File.Exists(newName) && newName != item.FullDocumentName)
                    {
                        item.ReferencedFileDescriptor.ReplaceReference(newName);
                    }
                //}
            }
            if (doc.Dirty)
            {
                doc.Update2(false);
                doc.Save2();
            }
            doc.Close();
        }

        public static void replaceFullFileName(PartDocument doc, string newName)
        {
            PartComponentDefinition compDef;
            compDef = doc.ComponentDefinition;
            FileDescriptor fd;
            if (compDef.ReferenceComponents.DerivedPartComponents.Count != 0)
            {
                fd = compDef.Parameters.DerivedParameterTables[1].ReferencedDocumentDescriptor.ReferencedFileDescriptor;
                if (newName != "")
                    fd.ReplaceReference(newName);
                doc.Update2(false);
            }
        }

        public static void replaceFullFileName(AssemblyDocument doc, string newName)
        {
            AssemblyComponentDefinition compDef;
            compDef = doc.ComponentDefinition;
            FileDescriptor fd;
            if (compDef.Parameters.DerivedParameterTables.Count != 0)
            {
                fd = compDef.Parameters.DerivedParameterTables[1].ReferencedDocumentDescriptor.ReferencedFileDescriptor;
                if (newName != "")
                    fd.ReplaceReference(newName);
                doc.Update2(false);
            }
        }                                                                                                                                                   

        public static void replaceFullFileName(DrawingDocument doc, string newName)
        {
            FileDescriptor fd;
            fd = doc.ReferencedDocumentDescriptors[1].ReferencedFileDescriptor;// doc.ActiveSheet.DrawingViews[1].ReferencedDocumentDescriptor.ReferencedFileDescriptor;
                if (newName != "")
                    fd.ReplaceReference(newName);
                if (doc.ReferencedDocumentDescriptors.Count > 1)
                {
                    for (int i = 1; i < doc.ReferencedDocumentDescriptors.Count; i++)
                    {
                        newName = InvDoc.u.OFD(System.IO.Path.GetDirectoryName(doc.FullDocumentName));
                        doc.ReferencedDocumentDescriptors[i + 1].ReferencedFileDescriptor.ReplaceReference(newName);
                    }
                }
                doc.Update2(false);
        }

        public object insertIFeature(PartDocument doc, AssemblyDocument asm = null, PlanarSketch ps1 = null)
        {
            string oldName = "", newName = "";
            Inventor.Application invApp = (Inventor.Application)doc.Parent;
            TransientGeometry tg = invApp.TransientGeometry;
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Filter = "inventor definition files(*.ide)|*.ide";
            ofd.InitialDirectory = invApp.iFeatureOptions.RootPath;
            //ofd.ShowDialog();
            //if (ofd.FileName != "") oldName = ofd.FileName;
            ofd.Title = "Выберите параметрический элемент";
            ofd.ShowDialog();
            if (ofd.FileName != "") newName = ofd.FileName;
            if (newName != "")
            {
               oldName = newName.Substring(newName.LastIndexOf('\\')+1, newName.Length - newName.LastIndexOf('\\')-5);
               PartComponentDefinition compDef = doc.ComponentDefinition;
               SketchLine sl = null;
               List<WorkPoint> wps = selectFromSketch(doc, ref sl, ps1);
               XMLDoc xmldoc = null;
               string n = doc.path() + "\\" + oldName + ".xml";
               if (System.IO.File.Exists(n))
                   xmldoc = xmldoc ?? new XMLDoc(n, "Property");
               foreach (WorkPoint item in wps)
               {
               iFeatureDefinition ifd = compDef.Features.iFeatures.CreateiFeatureDefinition(newName);
               foreach (iFeatureInput input in ifd.iFeatureInputs)
               {
                   switch (input.Type)
                   {
                       case ObjectTypeEnum.kiFeatureSketchPlaneInputObject:
                           ((iFeatureSketchPlaneInput)input).PlaneInput = ps.PlanarEntity;
                           break;
                       default:
                           break;
                   }
                   if (xmldoc != null)
                   {
                       foreach (var at in xmldoc.Doc.Root.Attributes())
                       {
                           if (input.Name.IndexOf(at.Name.ToString()) != -1)
                           {
                               ((iFeatureParameterInput)input).Value = double.Parse( at.Value)/10;
                           }
                       }
                   }
               }

               Edge edge1 = null; Edge edge2 = null;
               iFeature feature = compDef.Features.iFeatures.Add(ifd);
                   PlanarSketch sketchplanar = (PlanarSketch)feature.Sketches[1];
                   sketchplanar.OriginPoint = item;
                   Plane pl = sketchplanar.PlanarEntityGeometry;
                   if (sl != null) sketchplanar.AxisEntity = sl;
                   //sketchplanar.RotateSketchObjects(sketchplanar.SketchEntities, tg.CreatePoint2d(), 3.1415926 / 2);
                   var fac = feature.Faces.OfType<Face>().Where(f => f.SurfaceType == SurfaceTypeEnum.kCylinderSurface);
                   try
                   {
                   if (fac.Count() == 2)
                   {
                       edge1 = fac.First().Edges[1];
                       if (pl.DistanceTo(((Circle)edge1.Geometry).Center) != 0) edge1 = fac.First().Edges[2];
                       edge2 = fac.Last().Edges[1];
                       if (pl.DistanceTo(((Circle)edge2.Geometry).Center) != 0) edge2 = fac.Last().Edges[2];
                   }
                   else getImateEdge(oldName,ref edge1, ref edge2, feature);
                       if (asm == null)
                   addImate(oldName, edge1, edge2,(Document)doc, feature);
                       else
                       {
                           ComponentOccurrence o = asm.ComponentDefinition.Occurrences.OfType<ComponentOccurrence>().FirstOrDefault(occ => occ.Definition.Equals(doc.ComponentDefinition));
                           if (o != null)
                           {
                               object ep1, ep2;
                               o.CreateGeometryProxy(edge1, out ep1);
                               o.CreateGeometryProxy(edge2, out ep2);
                               addImate(oldName, ep1, ep2, (Document)asm, feature);
                           }
                       }
                   }
                   catch 
                   {
                   }
                   if (feature != null) return feature;
               }
            }
            return null;
         }

        public void addImate(string name, Edge edge1, Edge edge2, Document doc, iFeature feature)
        {
            IMate im = new IMate(doc);
            XMLDoc xmldoc = new XMLDoc(@"C:\ProgramData\Autodesk\Inventor Addins\AutoFastener.xml", "Fasteners");
            var item = xmldoc.Doc.Descendants("IMate").FirstOrDefault(e => e.Value == name);
            if (item != null)
            {
                Double d = double.Parse(item.Attribute("Offset").Value.Replace('.', ','));
                im.imate(item.Attribute("name").Value, name, d, doc, edge1, edge2, true);
                //if (item.Attribute("Offset2").Value != null && item.Attribute("name2").Value != null)
                //{
                //    foreach (Edge ed in getImateEdges(feature, doc))
                //    {
                //       im.imate(item.Attribute("name2").Value, name, d, doc, ed, null, true); 
                //    }
                //}
            }
        }
        public void addImate(string name, object ed1, object ed2, Document doc, iFeature feature)
        {
            EdgeProxy edge1 = (EdgeProxy)ed1;
            EdgeProxy edge2 = (EdgeProxy)ed2;
            IMate im = new IMate(doc);
            XMLDoc xmldoc = new XMLDoc(@"C:\ProgramData\Autodesk\Inventor Addins\AutoFastener.xml", "Fasteners");
            var item = xmldoc.Doc.Descendants("IMate").FirstOrDefault(e => e.Value == name);
            if (item != null)
            {
                Double d = double.Parse(item.Attribute("Offset").Value.Replace('.', ','));
                im.imate(item.Attribute("name").Value, name, d, doc, edge1, edge2, true);
            }
        }
        public List<Edge> getImateEdges(iFeature feature, Document doc)
        {
            int dist2 = 0;
            List<Edge> edges = new List<Edge>(10);
            Plane pl = ((PlanarSketch)feature.Sketches[1]).PlanarEntityGeometry;
            Edge tmpEdge = null;
            Circle circle = null;
            foreach (Face f in feature.Faces)
            {
                try
                {
                    circle = (Circle)f.Edges[1].Geometry;
                }
                catch
                {
                    continue;
                }
                tmpEdge = f.Edges[1];
                    dist2 = (int)(pl.DistanceTo(circle.Center) * 1000);
                    if (dist2 != 0)
                        tmpEdge = f.Edges[2];
                    edges.Add(tmpEdge);    
            }
            return edges;
        }

        public void getImateEdge(string name,ref Edge edge1, ref Edge edge2, iFeature feature)
        {
            Point ptMax = feature.RangeBox.MaxPoint; Point ptMin = feature.RangeBox.MinPoint;
            Plane pl = ((PlanarSketch)feature.Sketches[1]).PlanarEntityGeometry;
            int dist;
            foreach (Face f in feature.Faces)
            {
                switch (name)
                {
                    case "_ТЭН-резистор":
                        Cylinder cyl = (Cylinder)f.Geometry;
                        if ((int)(cyl.Radius*1000) == 300) edge1 = f.Edges[1];
                        else edge2 = f.Edges[1];
                        break;
                    case "_ТЭН":
                        cyl = (Cylinder)f.Geometry;
                        if ((int)(cyl.Radius * 1000) == 300) edge1 = f.Edges[1];
                        else edge2 = f.Edges[1];
                        break;
                    default:
                if (/*f.Evaluator.RangeBox.MaxPoint.IsEqualTo(ptMax) || */f.Evaluator.RangeBox.MinPoint.IsEqualTo(ptMin))
                {
                    edge1 = f.Edges[1];
                    dist = (int)(pl.DistanceTo(((Circle)edge1.Geometry).Center)*1000);
                    if (dist != 0) 
                        edge1 = f.Edges[2];
                }

                if (/*f.Evaluator.RangeBox.MaxPoint.IsEqualTo(ptMin) || */f.Evaluator.RangeBox.MaxPoint.IsEqualTo(ptMax))
                {
                    edge2 = f.Edges[1];
                    if (edge1 != null)
                    {
                    dist = (int)(pl.DistanceTo(((Circle)edge1.Geometry).Center) * 1000);
                    if (dist != 0) 
                        edge2 = f.Edges[2];
                    }
                }
                break;
                }
            }
            if (edge1 == null)
            {
                getMinDist(ref edge1, feature, ptMin, pl);
            }
            if (edge2 == null)
            {
                getMinDist(ref edge2, feature, ptMax, pl);
            }
        }

        public void getMinDist(ref Edge edge, iFeature feature, Inventor.Point pt, Plane pl)
        {
            double dist = 1000000; int dist2 = 0;
            Edge tmpEdge = null;
            Circle circle = null;
            foreach (Face f in feature.Faces)
            {
                try
                {
                    circle = (Circle)f.Edges[1].Geometry;
                }
                catch
                {
                    continue;
                }
                    tmpEdge = f.Edges[1];
                    if (circle.Center.DistanceTo(pt) < dist)
                    {
                        dist = circle.Center.DistanceTo(pt);
                        edge = tmpEdge;
                        dist2 = (int)(pl.DistanceTo(circle.Center) * 1000);
                        if (dist2 != 0)
                            edge = f.Edges[2];
                    }

            }
        }

        private List<WorkPoint> selectFromSketch(PartDocument m_PartDoc, ref SketchLine line, PlanarSketch ps1 = null)
        {
            CommandManager cmdMgr = ((Inventor.Application)m_PartDoc.Parent).CommandManager;
            List<WorkPoint> wps = new List<WorkPoint>(4);
            if (ps1 != null) ps = ps1;
            else if (m_PartDoc.SelectSet.Count != 0 && m_PartDoc.SelectSet[1] is PlanarSketch)
                ps = (PlanarSketch)m_PartDoc.SelectSet[1];
            else
                ps = (PlanarSketch)cmdMgr.Pick(SelectionFilterEnum.kAllPlanarEntities, "Выберите эскиз:");
            foreach (SketchPoint sp in ps.SketchPoints)
            {
               if (sp.HoleCenter == true)
               {
                   WorkPoint wp = m_PartDoc.ComponentDefinition.WorkPoints.AddByPoint(sp);
                   wp.Visible = false;
                   wps.Add(wp);
               }
            }
            ps.Visible = false;
            var sl = ps.SketchLines.OfType<SketchLine>().FirstOrDefault(l => l.Construction == true);
            if (sl != null) line = sl;
            return wps;
        }

        private void insertOcc(AssemblyComponentDefinition acd, string names)
        {
             string[] strs = names.Split('$');
            for (int i=0;i<strs.Length;i+=2)
            {
                int count = 0;
                if (strs[i + 1] == "") count = 1;
                else count = int.Parse(strs[i + 1]);
                bool flag = fileExist(strs[i]);
                openFlag = true;
                if (flag && acd.Occurrences.Count == 0)
                {
                    Matrix matrix = invApp.TransientGeometry.CreateMatrix();

                    matrix.SetTranslation(invApp.TransientGeometry.CreateVector());
                    for (int j = 0; j < count; j++)
                    {
                        ComponentOccurrence co = acd.Occurrences.Add(name, matrix);
                    }
                    openFlag = false;
                }
                else if (flag)
                    for (int j = 0; j < count; j++)
                    {
                        acd.Occurrences.AddUsingiMates(name);
                    }
                else if (strs[i].StartsWith("v3#"))
                {
                    string n = findInContentCenter(strs[i]);
                    for (int j = 0; j < count; j++)
                    {
                        acd.Occurrences.AddUsingiMates(n);
                    }
                }
            }
        }

        private void replaceOcc(AssemblyComponentDefinition acd, string n)
        {
            string[] names;
            names = n.Split('$');
            for (int i = 0; i < names.Length; i+=2)
            {
                ComponentOccurrence occ = findOcc(names[i], acd.Occurrences);
                Parts.removeOcc(acd, occ, false);
                if (occ != null && fileExist(names[i+1]))
                occ.Replace(name, true);
                else if (occ != null && names[i + 1].StartsWith("v3#"))
                {
                    occ.Replace(findInContentCenter(names[i + 1]),true);
                }
            }
        }

        public static List<ComponentOccurrence> findOccs(AssemblyComponentDefinition acd, string ffn)
        {
            List<ComponentOccurrence> occs = new List<ComponentOccurrence>();
            foreach (ComponentOccurrence occ in acd.Occurrences)
            {
                if (occ.ReferencedDocumentDescriptor.FullDocumentName == ffn) occs.Add(occ); 
            }
            return occs;
        }

        public static void removeReplace(AssemblyComponentDefinition acd, XElement el)
        {
            foreach (var row in el.Elements())
            {
                string name = row.Attribute("name").Value;
                List<ComponentOccurrence> occs = findOccs(acd, name);
                if (el.Name == "delete")
                {
                    
                    foreach (ComponentOccurrence occ in occs)
                    {
                        if (row.Attribute("r") != null)
                            Parts.removeOcc(acd, occ, true);
                        else
                            occ.Delete();
                    }  
                }
                else if (el.Name == "replace")
                {
                    string nameReplace = row.Attribute("nameReplace").Value;
                    string flag = "";
                    if (row.Attribute("r") != null) flag = row.Attribute("r").Value.ToString();
                    if (flag != "")
                    {
                        foreach (ComponentOccurrence occ in occs)
                        {
                            Parts.removeOcc(acd, occ, false);
                        }
                    } 
                    if (occs.Count != 0 && System.IO.File.Exists(nameReplace))
                        occs[0].Replace(nameReplace, true);
                    else if (occs.Count != 0 && nameReplace.StartsWith("v3#"))
                        occs[0].Replace(findInContentCenter(nameReplace), true);
                }
            }
        }

        public static void removeReplace(Document doc, string xmlName)
        {
            XDocument xmlDoc = XDocument.Load(xmlName);
            AssemblyComponentDefinition acd = (doc as AssemblyDocument).ComponentDefinition;
            if (xmlDoc != null)
            {
                foreach (var el in xmlDoc.Root.Elements())
                {
                    removeReplace(acd, el);
                }
            }
        }

        private static string findInContentCenter(string id)
        {
            string memberfilename = "";
            string failuremessage;
            MemberManagerErrorsEnum err;
            ContentCenter c = invApp.ContentCenter;
            ContentFamily fam = (ContentFamily)c.GetContentObject(id.Substring(0, id.LastIndexOf('#') + 1));
            ContentTableRow row = (ContentTableRow)c.GetContentObject(id);
            memberfilename = fam.CreateMember(row, out err, out failuremessage);
            return memberfilename;
        }

        public void createContentXML()
        {
            ContentCenter cc = invApp.ContentCenter;
            XMLDoc xmlDoc = new XMLDoc(@"C:\ProgramData\Autodesk\Inventor Addins\ContentCenter.xml", "Content");
            createNodeContent(cc.TreeViewTopNode, xmlDoc);
            xmlDoc.save();
        }

        private void createNodeContent(ContentTreeViewNode basenode, XMLDoc xmlDoc)
        {
            Dictionary<string, string> vals = new Dictionary<string,string>(5);
            foreach (ContentTreeViewNode cn in basenode.ChildNodes)
            {
                if (cn.ChildNodes.Count != 0)
                {
                    createFamilyContent(cn, xmlDoc);
                    createNodeContent(cn, xmlDoc);
                }
                else
                {
                    createFamilyContent(cn, xmlDoc);
                }
            }
        }

        public void addAttributes(string name, string value)
        {
            ContentOp contOp = new ContentOp();
            List<Edge> edgeCmp = new List<Edge>();
            List<Edge> edgeCmp1 = new List<Edge>();
            contOp.selOp(ref edgeCmp, ref edgeCmp1);
            try                                                                                   
            {
                foreach (Edge ed in edgeCmp)
                {
                    ed.AttributeSets.Add("AutoFastener");
                    ed.AttributeSets["AutoFastener"].Add(name, ValueTypeEnum.kStringType, value);
                }
            }
            catch 
            {
            }
        }

        public void placeAllFamilyContent()
        {
            XMLDoc xmlDoc = new XMLDoc(@"C:\ProgramData\Autodesk\Inventor Addins\Детали.xml", "Head");
            string memberfilename = "";
            string failuremessage;
            MemberManagerErrorsEnum err;
            ContentCenter cc = invApp.ContentCenter;
            foreach (var item in xmlDoc.Doc.Root.Descendants())
            {
                ContentFamily fam = (ContentFamily)cc.GetContentObject(item.Value);
                foreach (ContentTableRow row in fam.TableRows)
                {
                    memberfilename = fam.CreateMember(row, out err, out failuremessage, ContentMemberRefreshEnum.kRefreshOutOfDateParts);
                }
            }
        }

        private void createFamilyContent(ContentTreeViewNode basenode, XMLDoc xmlDoc)
        {
            Dictionary<string, string> vals = new Dictionary<string, string>(5);

            if (basenode.Families.Count != 0)
            {
                XElement elem = xmlDoc.El;
                foreach (ContentFamily fam in basenode.Families)
                {
                    //vals.Add("Description", fam.Description);
                    vals.Add("Id", fam.ContentIdentifier);
                    vals.Add("name", fam.DisplayName);
                    createTableRowContent(fam,xmlDoc.addXElement(elem, "family", vals));
                    vals.Clear();
                    //xmlDoc.El.Add(elem);
                }
                xmlDoc.El = xmlDoc.El.Parent;
            }
        }

        private void createTableRowContent(ContentFamily fam, XElement elem)
        {
            Dictionary<string, string> vals = new Dictionary<string, string>(5);
            List<string> adder = new List<string> { "Description[Project]", "t", "Упрощенный_вид [Custom]" };
            ContentTableColumn c = null; ContentTableColumn a = null;
            string propId, propSetId;
            foreach (ContentTableColumn col in fam.TableColumns.OfType<ContentTableColumn>().Where(e => e.HasPropertyMap))
	        {
                try
                {
                    col.GetPropertyMap(out propSetId, out propId);
                    if (propId == "29") { c = col; break; }
                    //col.GetPropertyMap(out propSetId, out propId);
                }
                catch
                {
                }
	        }
             
            foreach (var item in adder)
	        {
                a = fam.TableColumns.OfType<ContentTableColumn>().FirstOrDefault(e => e.DisplayHeading == item);
            }

            foreach (ContentTableRow item in fam.TableRows)
            {
                //XElement el;
                if (c != null && a != null)
                    elem.Add(new XElement("TableRow",new XAttribute("Id",item.ContentIdentifier), new XAttribute("Description",item[c].Value), new XAttribute("Extra",a.DisplayHeading + "=" + item[a].Value)));
                else if (c != null)
                elem.Add(new XElement("TableRow",new XAttribute("Id",item.ContentIdentifier), new XAttribute("Description",item[c].Value)));
            }

        }

        public void removeOcc(AssemblyComponentDefinition acd, string n)
        {
            string[] names;
            names = n.Split('$');
            for (int i = 0; i < names.Length; i++)
            {
                ComponentOccurrence occ = null;
                while ((occ = findOcc(names[i], acd.Occurrences)) != null)
                {
                    removeOcc(acd, occ); 
                }
                //if (occ != null)
                //{
                //    foreach (AssemblyConstraint constr in occ.Constraints /*acd.Constraints*/)
                //    {
                //        if (constr.OccurrenceOne != null) constr.OccurrenceOne.Delete(); 
                //    }
                // occ.Delete();
                //}
            }
        }

        public static void removeOcc(AssemblyComponentDefinition acd, ComponentOccurrence occ, bool occDelete = true)
        {
            if (occ != null)
            {
                HashSet<ComponentOccurrence> occs = new HashSet<ComponentOccurrence>();
                foreach (AssemblyConstraint constr in occ.Constraints)
                {
                    if (constr.OccurrenceOne != null && constr.OccurrenceTwo != null && constr.OccurrenceTwo.Equals(occ) &&
                        constr.OccurrenceOne.ReferencedDocumentDescriptor.FullDocumentName.IndexOf(@"Content Center Files\") != -1) 
                        occs.Add(constr.OccurrenceOne);
                }
                foreach (ComponentOccurrence item in occs)
                {
                    removeOcc(acd, item, false);
                    item.Delete();
                }
                if (occDelete)
                occ.Delete();
            }
        }

        private static ComponentOccurrence findOcc(string name, ComponentOccurrences co)
        {
            foreach (ComponentOccurrence cop in co)
            {
                string copname = cop.Name.Substring(0,cop.Name.IndexOf(':'));
                if (copname == name.Substring(0,name.Length-4)) return cop; 
            }
            return null;
        }

        public void addP(XMLData data)
        {
            doc = null; 
            string baseDoc = "";
            switch (data.name.ToLower())
            {
                case "part":
                    if (data.attr.ContainsKey("base")){ fileExist(data.attr["base"]); baseDoc = name;};
                        doc = (!fileExist(data.val))?
                        (data.attr.ContainsKey("base")) ? (Document)derivedDoc(baseDoc, data.attr["d"]) : (Document)addPrtDoc():
                        (Document)openPrtDoc(name);
                    break;
                case "asm":
                  if (data.attr.ContainsKey("base")){ fileExist(data.attr["base"]); baseDoc = name;};
                  doc = (!fileExist(data.val))? 
                      (data.attr.ContainsKey("base")) ? (Document)copyAsm(path + data.val,baseDoc): (Document)addAsmDoc():
                        (Document)openAsmDoc(name);
                    break;
            }

            if (doc != null)
            {
                invApp.SilentOperation = true;
                if (data.name.ToLower() == "asm" && data.attr.ContainsKey("base")) openFlag = true;
                bool oldOF = false;
                foreach (var item in data.attr)
                {
                    switch (item.Key)
                    {
                        case "pn":
                            addProp("Part Number", item.Value);
                            break;
                        case "d":
                            addProp("Description", item.Value);
                            break;
                        case "type":
                            addProp(doc,"Type", item.Value);
                            break;
                        case "decnumber":
                            addProp(doc,"DecNumber", item.Value);
                            break;
                        case "files":
                            insertOcc(((AssemblyDocument)doc).ComponentDefinition, replaceLib(item.Value));
                            break;
                        case "replace":
                            replaceOcc(((AssemblyDocument)doc).ComponentDefinition, replaceLib(item.Value));
                            break;
                        case "remove":
                            removeOcc(((AssemblyDocument)doc).ComponentDefinition, item.Value);
                            break;
                        case "base":
                            break;
                        default:
                            addProp(item.Key, item.Value);
                            break;
                    }
                    oldOF = openFlag;
                }
                openFlag = oldOF;
                
                if (data.attr.ContainsKey("decnumber") && data.attr.ContainsKey("type"))
                {
                    addProp("DecNumber", data.attr["decnumber"]);
                    addProp("Type", data.attr["type"]);
                    addProp("Part Number", "=<Type>.<DecNumber>");
                };
                if (openFlag)
                    doc.Save();
                else
                  {
                    doc.SaveAs(path + data.val, false);
                    if (data.attr.ContainsKey("decnumber") && data.attr.ContainsKey("type"))
                        {
                            addProp("Part Number", "=<Type>.<DecNumber>");
                            doc.Save();
                        }
                  }
                if (data.name.ToLower() == "asm" && openFlag)
                {
                    ContentOp coo = new ContentOp();
                    coo.programmAdd((AssemblyDocument)doc);
                    doc.Save();
                }
                doc.Close();
                invApp.SilentOperation = false;    
            }
        }

        public void addP(XElement baseEl)
        {
            doc = null; List<string> names = new List<string> { "pn", "d", "type", "decnumber", "files", "replace", "remove" };
            string baseDoc = ""; string val = "";
            foreach (var data in baseEl.Elements())
            {
                if (data.HasElements) addP(data);
            switch (data.Name.ToString().ToLower())
            {
                case "part":
                    if (data.Attribute("base") != null) { val = data.Attribute("base").Value;}
                    else if (data.Parent.Attribute("base") != null) { val = data.Parent.Attribute("base").Value; }
                    if (val != "") { fileExist(val); baseDoc = name; };
                    doc = (!fileExist(data.Value)) ?
                    (val != "") ? (Document)derivedDoc(baseDoc, data.Attribute("d").Value) : (Document)addPrtDoc() :
                    (Document)openPrtDoc(name);
                    break;
                case "asm":
                    if (data.Attribute("base") != null) { val = data.Attribute("base").Value;}
                    else if (data.Parent.Attribute("base") != null) { val = data.Parent.Attribute("base").Value; }
                    if (val != "") { fileExist(val); baseDoc = name; };
                    doc = (!fileExist(data.Value)) ?
                        (val != "") ? (Document)copyAsm(path + data.Value, baseDoc) : (Document)addAsmDoc() :
                          (Document)openAsmDoc(name);
                    break;
            }

            if (doc != null)
            {
                invApp.SilentOperation = true;
                if (data.Name.ToString().ToLower() == "asm" && data.Attribute("base") != null) openFlag = true;
                bool oldOF = false;
                foreach (var item in data.Attributes())
                {
                    
                    if (!names.Exists(e => e == item.Name.ToString())) addProp(item.Name.ToString(), item.Value);
                    oldOF = openFlag;
                }
                val = "";
                foreach (var item in names)
                {
                    XMLDoc.getXAttributeValue(data, item, ref val);
                    if (val != "")
                    {
                        switch (item)
                        {
                            case "pn":
                                addProp("Part Number", val);
                                break;
                            case "d":
                                addProp("Description", val);
                                break;
                            case "type":
                                addProp(doc, "Type", val);
                                break;
                            case "decnumber":
                                addProp(doc, "DecNumber", val);
                                break;
                            case "files":
                                insertOcc(((AssemblyDocument)doc).ComponentDefinition, replaceLib(val));
                                break;
                            case "replace":
                                replaceOcc(((AssemblyDocument)doc).ComponentDefinition, replaceLib(val));
                                break;
                            case "remove":
                                removeOcc(((AssemblyDocument)doc).ComponentDefinition, val);
                                break;
                            case "base":
                                break;
                            default:
                                break;
                        }
                    }
                    val = "";
                }
                openFlag = oldOF;

                if (data.Attribute("decnumber") != null && data.Attribute("type") != null)
                {
                    addProp("DecNumber", data.Attribute("decnumber").Value);
                    addProp("Type", data.Attribute("type").Value);
                    addProp("Part Number", "=<Type>.<DecNumber>");
                };
                if (openFlag)
                    doc.Save();
                else
                {
                    doc.SaveAs(path + data.Value, false);
                    if (data.Attribute("decnumber") != null && data.Attribute("type") != null)
                    {
                        addProp("Part Number", "=<Type>.<DecNumber>");
                        doc.Save();
                    }
                }
                if (data.Name.ToString().ToLower() == "asm" && openFlag)
                {
                    ContentOp coo = new ContentOp();
                    coo.programmAdd((AssemblyDocument)doc);
                    doc.Save();
                }
                doc.Close();
                invApp.SilentOperation = false;
            }

            }
        }

        static public ContourFlangeFeature addContourFlange(SheetMetalComponentDefinition smcd, string name, string name2 = "")
        {
            try
            {
                PlanarSketch ps;
                if (name2 != "") ps = smcd.Sketches.OfType<PlanarSketch>().FirstOrDefault(s => s.Name.IndexOf(name2) != -1);
                else if (smcd.ReferenceComponents.DerivedPartComponents[1].Sketches.Count != 0) ps = (PlanarSketch)smcd.ReferenceComponents.DerivedPartComponents[1].Sketches[1];
                else if (smcd.Sketches.Count != 0) ps = (PlanarSketch)smcd.Sketches[1];
                else return null;
                if (ps == null) return null;
                SheetMetalFeatures smf = (SheetMetalFeatures)smcd.Features;
                ContourFlangeFeature cff;
                SketchLine sl = ps.SketchLines.OfType<SketchLine>().FirstOrDefault(l => l.Construction == false);
                Path p = smf.CreatePath(sl);
                ContourFlangeDefinition cfd = smf.ContourFlangeFeatures.CreateContourFlangeDefinition(p);
                Parameter param = CreateComponent.getParameter((Document)smcd.Document, name),
                param1 = CreateComponent.getParameter((Document)smcd.Document, "БВ_шип");
                if (param != null)
                {
                    string n = name;
                    if (param.Name.ToLower() == "бв_длина" && param1 != null) n = name + " + БВ_шип*2";
                    else if (param.Name.ToLower() == "бв_длина") n = name + " + 15";
                    cfd.SetDistanceExtent(n, PartFeatureExtentDirectionEnum.kSymmetricExtentDirection);
                    cff = smf.ContourFlangeFeatures.Add(cfd);
                    ps.Visible = false;
                    return cff;
                }
                return null;
            }
            catch
            {
                return null;
            }
        }
        static public FaceFeature addBaseFeature(SheetMetalComponentDefinition compDef, string name)
        {
            try
            {
                PlanarSketch ps;
//                 if (compDef.ReferenceComponents.DerivedPartComponents[1].Sketches.Count != 0) ps = (PlanarSketch)compDef.ReferenceComponents.DerivedPartComponents[1].Sketches[1];
//                 else if (compDef.Sketches.Count != 0) ps = (PlanarSketch)compDef.Sketches[1];
//                 else return null;
                ps = compDef.Sketches.OfType<PlanarSketch>().FirstOrDefault(s => s.Name.IndexOf(name) != -1);
                if (ps == null) return null;
                Profile pr = ps.Profiles.AddForSolid();
                SheetMetalFeatures smf = compDef.Features as SheetMetalFeatures;
                FaceFeatureDefinition ffd = smf.FaceFeatures.CreateFaceFeatureDefinition(pr);
                return smf.FaceFeatures.Add(ffd);
            }
            catch
            {
                return null;
            }
        }

        static public FlangeFeature addFlange(SheetMetalComponentDefinition compDef, string namePar, string cond)
        {
            try
            {
                SheetMetalFeatures smf = compDef.Features as SheetMetalFeatures;
                Face f = smf.FaceFeatures[1].Faces[smf.FaceFeatures[1].Faces.Count - 1];
                EdgeCollection col = I.objs.CreateEdgeCollection();
                double val = ut.convToDouble(cond.TrimStart(new char[] { '<', '>' }));
                IEnumerable<Edge> eds;
                eds = (cond.StartsWith("<"))?
                    f.Edges.OfType<Edge>().Where(e => ut.getLenght(e) < val):
                    f.Edges.OfType<Edge>().Where(e => ut.getLenght(e) > val);
                foreach (var item in eds)
                {
                    col.Add(item); 
                }
                if (col.Count == 0) return null;
                FlangeDefinition fd = smf.FlangeFeatures.CreateFlangeDefinition(col, "90", namePar);
                fd.ApplyAutoMitering = false;
                fd.CornerOptions.CornerReliefShape = CornerReliefShapeEnum.kNoReplacementCornerReliefShape;
                FlangeFeature ff = smf.FlangeFeatures.Add(fd);
                return ff;
            }
            catch
            {
                return null;
            }
        }
    }

    internal class VarBtn : Button
    {
        //public static Drawings m_Drw;
        //public static Drawings getDrw { get { return m_Drw; } }
        public VarBtn(string displayName, string internalName, string clientId, string description, string tooltip,
            ButtonDisplayEnum buttonDisplayType = ButtonDisplayEnum.kDisplayTextInLearningMode, CommandTypesEnum commandType = CommandTypesEnum.kNonShapeEditCmdType)
            : base(displayName, internalName, commandType, clientId, description, tooltip, buttonDisplayType) { }
        protected override void ButtonDefinition_OnExecute(NameValueMap context)
        {
            Document doc = Macros.StandardAddInServer.m_inventorApplication.ActiveDocument;
            Parts m_Parts = new Parts(doc);
            if (doc.DocumentType == DocumentTypeEnum.kAssemblyDocumentObject)
            {
                try
                {
                    m_Parts.insertIFeature((PartDocument)doc.ActivatedObject, (AssemblyDocument)doc);
                }
                catch (Exception)
                {
                }
            }
            else if (doc.DocumentType == DocumentTypeEnum.kPartDocumentObject)
                m_Parts.insertIFeature((PartDocument)doc);
        }
    }
}

namespace ExtensionMethods
{
    public static class MyPath
    {
        public static string path(this Inventor.Document doc)
        {
            return doc.FullFileName.Substring(0,doc.FullFileName.LastIndexOf('\\'));
        }
        public static string path(this Inventor.PartDocument doc)
        {
            return doc.FullFileName.Substring(0, doc.FullFileName.LastIndexOf('\\'));
        }
        public static string path(this Inventor.AssemblyDocument doc)
        {
            return doc.FullFileName.Substring(0, doc.FullFileName.LastIndexOf('\\'));
        }
        public static string path(this Inventor.DrawingDocument doc)
        {
            return doc.FullFileName.Substring(0, doc.FullFileName.LastIndexOf('\\'));
        }
        public static string name(this Inventor.Document doc)
        {
            return doc.FullFileName.Substring(doc.FullFileName.LastIndexOf('\\') + 1, doc.FullFileName.Length - 1 - doc.FullFileName.LastIndexOf('\\')-4);         
        }
    }
}
