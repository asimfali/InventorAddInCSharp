//#define INV14

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
        private ToolStripMenuItem инициализацияToolStripMenuItem;
        private ToolStripMenuItem шипИнверсияToolStripMenuItem;
        private ToolStripMenuItem экспортВSATToolStripMenuItem;
        private ToolStripMenuItem кПToolStripMenuItem;
        private ToolStripMenuItem создатьЭскизыToolStripMenuItem;
        private ToolStripMenuItem выравниваниеОсейToolStripMenuItem;
        private ToolStripMenuItem перенестиВЦентрToolStripMenuItem;
        private ToolStripMenuItem совместитьКПToolStripMenuItem;
        private ToolStripMenuItem копироватьToolStripMenuItem;
        private ToolStripMenuItem восстановитьToolStripMenuItem;
        private ToolStripMenuItem копироватьВПапкуToolStripMenuItem;
        private ToolStripMenuItem dPDFToolStripMenuItem;
        private ToolStripMenuItem подпозицииToolStripMenuItem;
        private ToolStripMenuItem ширинаНадписиToolStripMenuItem;
        private ToolStripMenuItem позицииПоШаблонуToolStripMenuItem;
        private ToolStripMenuItem видыПоШаблонуToolStripMenuItem;
        private ToolStripMenuItem подготовитьШаблонToolStripMenuItem;
        private ToolStripMenuItem комплектНесколькихФайловToolStripMenuItem;
        private ToolStripMenuItem размерыПоШаблонуToolStripMenuItem;
        private ToolStripMenuItem осевыеПоШаблонуToolStripMenuItem;
        private ToolStripMenuItem видимостьToolStripMenuItem;
        private ToolStripMenuItem названиеКПToolStripMenuItem;
        private ToolStripMenuItem связатьВPDFToolStripMenuItem;
        private ToolStripMenuItem удалитьПробелыToolStripMenuItem;
        private ToolStripMenuItem стильРазмеровСборкиToolStripMenuItem;
        private ToolStripMenuItem переназватьToolStripMenuItem1;
        private ToolStripMenuItem получитьНазванияToolStripMenuItem;
        private ToolStripMenuItem свойствоПоШаблонуToolStripMenuItem;
        private ToolStripMenuItem открытьПоШаблонуToolStripMenuItem;
        private ToolStripMenuItem экспортВOBJToolStripMenuItem;
        private ToolStripMenuItem экспортВGLBОдноТелоToolStripMenuItem;
        private ToolStripMenuItem экспортВGLBСборкаToolStripMenuItem;
        private ToolStripMenuItem восстановитьToolStripMenuItem1;
        private ToolStripMenuItem извToolStripMenuItem;
        private ToolStripMenuItem добавитьToolStripMenuItem;
        private ToolStripMenuItem заполнитьToolStripMenuItem;
        private ToolStripMenuItem сформироватьКомплектToolStripMenuItem;
        private ToolStripMenuItem создатьДеталиToolStripMenuItem1;
        private ToolStripMenuItem оЦToolStripMenuItem;
        private ToolStripMenuItem разрывВидаToolStripMenuItem;
        private ToolStripMenuItem найтиЭлементToolStripMenuItem;
        private ToolStripMenuItem переименоватьУзлыToolStripMenuItem;
        private ToolStripMenuItem разверткаToolStripMenuItem;
        private ToolStripMenuItem убратьВинтToolStripMenuItem;
        private ToolStripMenuItem совместитьПлоскостиToolStripMenuItem;
        private ToolStripMenuItem адаптивныеПлоскостиToolStripMenuItem;
        private ToolStripMenuItem связатьКПToolStripMenuItem;
        private ToolStripMenuItem зеркльныеКПToolStripMenuItem;
        private ToolStripMenuItem копироватьПроектToolStripMenuItem;
        private ToolStripMenuItem перенестиКПВСборкуToolStripMenuItem;
        private ToolStripMenuItem копироватьКПToolStripMenuItem;
        private ToolStripMenuItem перенаправитьToolStripMenuItem1;
        private ToolStripMenuItem сайтToolStripMenuItem;
        private ToolStripMenuItem названиеМоделиToolStripMenuItem;
        private ToolStripMenuItem спецификацияToolStripMenuItem;
        private ToolStripMenuItem тест2ToolStripMenuItem;
        private ToolStripMenuItem поверхностьПодОвалToolStripMenuItem;
        private ToolStripMenuItem весьПроектToolStripMenuItem;
        private ToolStripMenuItem создатьСписокФайловToolStripMenuItem;
        private ToolStripMenuItem gltfToolStripMenuItem;
        private ToolStripMenuItem моделиИзСпискаToolStripMenuItem;
        private ToolStripMenuItem dИзСпискаФайловToolStripMenuItem;
        private ToolStripMenuItem данныеИзTitleToolStripMenuItem;
        private ToolStripMenuItem сохранитьDWGToolStripMenuItem;
        private ToolStripMenuItem вставкаПоСценариюToolStripMenuItem;
        private ToolStripMenuItem фаскиToolStripMenuItem;
        private ToolStripMenuItem граниВDXFToolStripMenuItem;
        private ToolStripMenuItem тестSplToolStripMenuItem;
        private ToolStripMenuItem заменитьМодельZToolStripMenuItem;
        private ToolStripMenuItem создатьСборкуToolStripMenuItem1;
        private ToolStripMenuItem автоЭскизыToolStripMenuItem;
        private ToolStripMenuItem автоКПToolStripMenuItem;
        private ToolStripMenuItem фланецОтступToolStripMenuItem;
        private ToolStripMenuItem оптБрToolStripMenuItem;
        private ToolStripMenuItem пазМультидетальToolStripMenuItem;
        private ToolStripMenuItem скрытьToolStripMenuItem;
        private ToolStripMenuItem автоБлокRToolStripMenuItem;
        private ToolStripMenuItem копироватьБлокToolStripMenuItem;
        private ToolStripMenuItem вставитьБлокToolStripMenuItem;
        private ToolStripMenuItem экспортВSTPToolStripMenuItem;
        private ToolStripMenuItem pDFЗеркальныеДеталиToolStripMenuItem;
        private ToolStripMenuItem объединитьToolStripMenuItem;
        private ToolStripMenuItem меткиToolStripMenuItem;
        private ToolStripMenuItem добавитьИсполнениеToolStripMenuItem1;
        private ToolStripMenuItem связатьКрепежToolStripMenuItem;
        private ToolStripMenuItem исполненияToolStripMenuItem;
        private ToolStripMenuItem открепленныеРазмерыToolStripMenuItem;
        private ToolStripMenuItem объединитьГибыToolStripMenuItem;
        private ToolStripMenuItem проекцияToolStripMenuItem;
        private ToolStripMenuItem расширитьИсполнениеToolStripMenuItem;
        private ToolStripMenuItem масштабToolStripMenuItem;
        private ToolStripMenuItem подавитьТелоSToolStripMenuItem;
        private ToolStripMenuItem сменитьToolStripMenuItem;
        private ToolStripMenuItem добавитьКПToolStripMenuItem;
        private ToolStripMenuItem добавитьУдалитьЗаменитьToolStripMenuItem;
        private ToolStripMenuItem выноскиГибовToolStripMenuItem;
        private ToolStripMenuItem изменитьОтверстияToolStripMenuItem;
        static HashSet<object> del = new HashSet<object>();
        public event Macros.MyHandler myEvent;

        private void InitializeComponent()
        {
            this.menuStrip1 = new System.Windows.Forms.MenuStrip();
            this.детальToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.кластерToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.скругленияToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.фаскиToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.ручнаяГибкаToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.создатьДеталиToolStripMenuItem1 = new System.Windows.Forms.ToolStripMenuItem();
            this.создатьСборкуToolStripMenuItem1 = new System.Windows.Forms.ToolStripMenuItem();
            this.связатьКПToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.зеркльныеКПToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.копироватьКПToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.добавитьКПToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.найтиЭлементToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.переименоватьУзлыToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.разверткаToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.объединитьГибыToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.убратьВинтToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.фланецОтступToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.копироватьПроектToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.поверхностьПодОвалToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.автоЭскизыToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.автоКПToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.скрытьToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.подавитьТелоSToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.автоБлокRToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.копироватьБлокToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.вставитьБлокToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.объединитьToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.меткиToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.добавитьИсполнениеToolStripMenuItem1 = new System.Windows.Forms.ToolStripMenuItem();
            this.расширитьИсполнениеToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.сменитьToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.масштабToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.сборкаToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.спецToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.крепежToolStripMenuItem1 = new System.Windows.Forms.ToolStripMenuItem();
            this.обновитьКрепежToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.крепежToolStripMenuItem3 = new System.Windows.Forms.ToolStripMenuItem();
            this.переименоватьСборкуToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.кПToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.вставкаПоСценариюToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.совместитьКПToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.названиеКПToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.совместитьПлоскостиToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.адаптивныеПлоскостиToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.перенестиКПВСборкуToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.исполненияToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.добавитьУдалитьЗаменитьToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.чертежToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.текущийДокументToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.регионToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.обновитьПозицииToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.подготовитьШаблонToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.видыПоШаблонуToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.позицииПоШаблонуToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.размерыПоШаблонуToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.осевыеПоШаблонуToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.addTRToolStripMenuItem1 = new System.Windows.Forms.ToolStripMenuItem();
            this.copyAttrToolStripMenuItem1 = new System.Windows.Forms.ToolStripMenuItem();
            this.стильРазмеровСборкиToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.получитьТекстToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.перевестиТекстToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.обновитьСтилиToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.инициализацияToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.переназватьToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.пересобратьСборкуToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.разрывВидаToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.перенаправитьToolStripMenuItem1 = new System.Windows.Forms.ToolStripMenuItem();
            this.заменитьМодельZToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.связатьКрепежToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.параметрыToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.добавитьПараметрыToolStripMenuItem1 = new System.Windows.Forms.ToolStripMenuItem();
            this.загрузитьXMLToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.данныеДляМассиваToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.массивToolStripMenuItem1 = new System.Windows.Forms.ToolStripMenuItem();
            this.копироватьВXMLToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.структураВXMLToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.создатьЭскизыToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.создатьДеталиToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.dPDFToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.разместитьВсеДеталиToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.разместитьВсеДеталиToolStripMenuItem1 = new System.Windows.Forms.ToolStripMenuItem();
            this.обновитьContentCenterToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.сортировкаToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.тестToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.тест2ToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.копироватьToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.восстановитьToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.данныеИзTitleToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.сохранитьDWGToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.граниВDXFToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.тестSplToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.редкоеToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.stepToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.комплектФайловToolStripMenuItem1 = new System.Windows.Forms.ToolStripMenuItem();
            this.комплектФайловСЧертежамиToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.комплектНесколькихФайловToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.обновитьToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.добавитьСвойствоToolStripMenuItem1 = new System.Windows.Forms.ToolStripMenuItem();
            this.свойствоПоШаблонуToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.добавитьИсполнениеToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.сплайнToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.артикулToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.центрыОтверстийToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.создатьСборкуToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.заменитьИмяКПToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.копироватьВПапкуToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.экспортВSATToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.экспортВSTPToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.габаритыToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.открытьПапкуToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.выравниваниеОсейToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.перенестиВЦентрToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.ширинаНадписиToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.видимостьToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.связатьВPDFToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.pDFЗеркальныеДеталиToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.удалитьПробелыToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.получитьНазванияToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.переназватьToolStripMenuItem1 = new System.Windows.Forms.ToolStripMenuItem();
            this.открытьПоШаблонуToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.экспортВOBJToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.экспортВGLBОдноТелоToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.экспортВGLBСборкаToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.восстановитьToolStripMenuItem1 = new System.Windows.Forms.ToolStripMenuItem();
            this.оЦToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.шипыИПазыToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.шипToolStripMenuItem1 = new System.Windows.Forms.ToolStripMenuItem();
            this.шипИнверсияToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.выдавитьToolStripMenuItem1 = new System.Windows.Forms.ToolStripMenuItem();
            this.добавитьВСборкуToolStripMenuItem1 = new System.Windows.Forms.ToolStripMenuItem();
            this.пазToolStripMenuItem1 = new System.Windows.Forms.ToolStripMenuItem();
            this.пазМультидетальToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.создатьФайлыToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.удалитьЭлементToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.общиеРазмерыToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.размерыДоГибовToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.позицииToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.крепежToolStripMenuItem2 = new System.Windows.Forms.ToolStripMenuItem();
            this.подпозицииToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.открепленныеРазмерыToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.проекцияToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.выноскиГибовToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.извToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.заполнитьToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.добавитьToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.сформироватьКомплектToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.перенаправитьToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.крепежToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.сайтToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.названиеМоделиToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.спецификацияToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.создатьСписокФайловToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.моделиИзСпискаToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.весьПроектToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.gltfToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.dИзСпискаФайловToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.оптБрToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.изменитьОтверстияToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
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
            this.извToolStripMenuItem,
            this.перенаправитьToolStripMenuItem,
            this.крепежToolStripMenuItem,
            this.сайтToolStripMenuItem});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Size = new System.Drawing.Size(827, 24);
            this.menuStrip1.TabIndex = 0;
            this.menuStrip1.Text = "menuStrip1";
            this.menuStrip1.ItemClicked += new System.Windows.Forms.ToolStripItemClickedEventHandler(this.menuStrip1_ItemClicked);
            // 
            // детальToolStripMenuItem
            // 
            this.детальToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.кластерToolStripMenuItem,
            this.скругленияToolStripMenuItem,
            this.фаскиToolStripMenuItem,
            this.ручнаяГибкаToolStripMenuItem,
            this.создатьДеталиToolStripMenuItem1,
            this.создатьСборкуToolStripMenuItem1,
            this.связатьКПToolStripMenuItem,
            this.зеркльныеКПToolStripMenuItem,
            this.копироватьКПToolStripMenuItem,
            this.добавитьКПToolStripMenuItem,
            this.найтиЭлементToolStripMenuItem,
            this.переименоватьУзлыToolStripMenuItem,
            this.разверткаToolStripMenuItem,
            this.объединитьГибыToolStripMenuItem,
            this.убратьВинтToolStripMenuItem,
            this.фланецОтступToolStripMenuItem,
            this.копироватьПроектToolStripMenuItem,
            this.поверхностьПодОвалToolStripMenuItem,
            this.автоЭскизыToolStripMenuItem,
            this.автоКПToolStripMenuItem,
            this.скрытьToolStripMenuItem,
            this.подавитьТелоSToolStripMenuItem,
            this.автоБлокRToolStripMenuItem,
            this.копироватьБлокToolStripMenuItem,
            this.вставитьБлокToolStripMenuItem,
            this.объединитьToolStripMenuItem,
            this.меткиToolStripMenuItem,
            this.добавитьИсполнениеToolStripMenuItem1,
            this.расширитьИсполнениеToolStripMenuItem,
            this.сменитьToolStripMenuItem,
            this.масштабToolStripMenuItem});
            this.детальToolStripMenuItem.Name = "детальToolStripMenuItem";
            this.детальToolStripMenuItem.Size = new System.Drawing.Size(57, 20);
            this.детальToolStripMenuItem.Text = "Деталь";
            // 
            // кластерToolStripMenuItem
            // 
            this.кластерToolStripMenuItem.Name = "кластерToolStripMenuItem";
            this.кластерToolStripMenuItem.Size = new System.Drawing.Size(206, 22);
            this.кластерToolStripMenuItem.Text = "Кластер";
            this.кластерToolStripMenuItem.Click += new System.EventHandler(this.кластерToolStripMenuItem_Click_1);
            // 
            // скругленияToolStripMenuItem
            // 
            this.скругленияToolStripMenuItem.Name = "скругленияToolStripMenuItem";
            this.скругленияToolStripMenuItem.Size = new System.Drawing.Size(206, 22);
            this.скругленияToolStripMenuItem.Text = "Скругления";
            this.скругленияToolStripMenuItem.Click += new System.EventHandler(this.скругленияToolStripMenuItem_Click_1);
            // 
            // фаскиToolStripMenuItem
            // 
            this.фаскиToolStripMenuItem.Name = "фаскиToolStripMenuItem";
            this.фаскиToolStripMenuItem.Size = new System.Drawing.Size(206, 22);
            this.фаскиToolStripMenuItem.Text = "Фаски";
            this.фаскиToolStripMenuItem.Click += new System.EventHandler(this.фаскиToolStripMenuItem_Click);
            // 
            // ручнаяГибкаToolStripMenuItem
            // 
            this.ручнаяГибкаToolStripMenuItem.Name = "ручнаяГибкаToolStripMenuItem";
            this.ручнаяГибкаToolStripMenuItem.Size = new System.Drawing.Size(206, 22);
            this.ручнаяГибкаToolStripMenuItem.Text = "Ручная гибка";
            this.ручнаяГибкаToolStripMenuItem.Click += new System.EventHandler(this.ручнаяГибкаToolStripMenuItem_Click_1);
            // 
            // создатьДеталиToolStripMenuItem1
            // 
            this.создатьДеталиToolStripMenuItem1.Name = "создатьДеталиToolStripMenuItem1";
            this.создатьДеталиToolStripMenuItem1.Size = new System.Drawing.Size(206, 22);
            this.создатьДеталиToolStripMenuItem1.Text = "Создать детали";
            this.создатьДеталиToolStripMenuItem1.Click += new System.EventHandler(this.создатьДеталиToolStripMenuItem1_Click);
            // 
            // создатьСборкуToolStripMenuItem1
            // 
            this.создатьСборкуToolStripMenuItem1.Name = "создатьСборкуToolStripMenuItem1";
            this.создатьСборкуToolStripMenuItem1.Size = new System.Drawing.Size(206, 22);
            this.создатьСборкуToolStripMenuItem1.Text = "Создать сборку";
            this.создатьСборкуToolStripMenuItem1.Click += new System.EventHandler(this.создатьСборкуToolStripMenuItem1_Click);
            // 
            // связатьКПToolStripMenuItem
            // 
            this.связатьКПToolStripMenuItem.Name = "связатьКПToolStripMenuItem";
            this.связатьКПToolStripMenuItem.Size = new System.Drawing.Size(206, 22);
            this.связатьКПToolStripMenuItem.Text = "Связать КП";
            this.связатьКПToolStripMenuItem.Click += new System.EventHandler(this.связатьКПToolStripMenuItem_Click);
            // 
            // зеркльныеКПToolStripMenuItem
            // 
            this.зеркльныеКПToolStripMenuItem.Name = "зеркльныеКПToolStripMenuItem";
            this.зеркльныеКПToolStripMenuItem.Size = new System.Drawing.Size(206, 22);
            this.зеркльныеКПToolStripMenuItem.Text = "Зеркльные КП";
            this.зеркльныеКПToolStripMenuItem.Click += new System.EventHandler(this.зеркльныеКПToolStripMenuItem_Click);
            // 
            // копироватьКПToolStripMenuItem
            // 
            this.копироватьКПToolStripMenuItem.Name = "копироватьКПToolStripMenuItem";
            this.копироватьКПToolStripMenuItem.Size = new System.Drawing.Size(206, 22);
            this.копироватьКПToolStripMenuItem.Text = "Копировать КП";
            this.копироватьКПToolStripMenuItem.Click += new System.EventHandler(this.копироватьКПToolStripMenuItem_Click);
            // 
            // добавитьКПToolStripMenuItem
            // 
            this.добавитьКПToolStripMenuItem.Name = "добавитьКПToolStripMenuItem";
            this.добавитьКПToolStripMenuItem.Size = new System.Drawing.Size(206, 22);
            this.добавитьКПToolStripMenuItem.Text = "Добавить КП";
            this.добавитьКПToolStripMenuItem.Click += new System.EventHandler(this.добавитьКПToolStripMenuItem_Click);
            // 
            // найтиЭлементToolStripMenuItem
            // 
            this.найтиЭлементToolStripMenuItem.Name = "найтиЭлементToolStripMenuItem";
            this.найтиЭлементToolStripMenuItem.Size = new System.Drawing.Size(206, 22);
            this.найтиЭлементToolStripMenuItem.Text = "Найти элемент";
            this.найтиЭлементToolStripMenuItem.Click += new System.EventHandler(this.найтиЭлементToolStripMenuItem_Click);
            // 
            // переименоватьУзлыToolStripMenuItem
            // 
            this.переименоватьУзлыToolStripMenuItem.Name = "переименоватьУзлыToolStripMenuItem";
            this.переименоватьУзлыToolStripMenuItem.Size = new System.Drawing.Size(206, 22);
            this.переименоватьУзлыToolStripMenuItem.Text = "Переименовать узлы";
            this.переименоватьУзлыToolStripMenuItem.Click += new System.EventHandler(this.переименоватьУзлыToolStripMenuItem_Click);
            // 
            // разверткаToolStripMenuItem
            // 
            this.разверткаToolStripMenuItem.Name = "разверткаToolStripMenuItem";
            this.разверткаToolStripMenuItem.Size = new System.Drawing.Size(206, 22);
            this.разверткаToolStripMenuItem.Text = "Развертка";
            this.разверткаToolStripMenuItem.Click += new System.EventHandler(this.разверткаToolStripMenuItem_Click);
            // 
            // объединитьГибыToolStripMenuItem
            // 
            this.объединитьГибыToolStripMenuItem.Name = "объединитьГибыToolStripMenuItem";
            this.объединитьГибыToolStripMenuItem.Size = new System.Drawing.Size(206, 22);
            this.объединитьГибыToolStripMenuItem.Text = "Объединить гибы";
            this.объединитьГибыToolStripMenuItem.Click += new System.EventHandler(this.объединитьГибыToolStripMenuItem_Click);
            // 
            // убратьВинтToolStripMenuItem
            // 
            this.убратьВинтToolStripMenuItem.Name = "убратьВинтToolStripMenuItem";
            this.убратьВинтToolStripMenuItem.Size = new System.Drawing.Size(206, 22);
            this.убратьВинтToolStripMenuItem.Text = "Правка косого среза";
            this.убратьВинтToolStripMenuItem.Click += new System.EventHandler(this.убратьВинтToolStripMenuItem_Click);
            // 
            // фланецОтступToolStripMenuItem
            // 
            this.фланецОтступToolStripMenuItem.Name = "фланецОтступToolStripMenuItem";
            this.фланецОтступToolStripMenuItem.Size = new System.Drawing.Size(206, 22);
            this.фланецОтступToolStripMenuItem.Text = "Фланец отступ";
            this.фланецОтступToolStripMenuItem.Click += new System.EventHandler(this.фланецОтступToolStripMenuItem_Click);
            // 
            // копироватьПроектToolStripMenuItem
            // 
            this.копироватьПроектToolStripMenuItem.Name = "копироватьПроектToolStripMenuItem";
            this.копироватьПроектToolStripMenuItem.Size = new System.Drawing.Size(206, 22);
            this.копироватьПроектToolStripMenuItem.Text = "Копировать проект";
            this.копироватьПроектToolStripMenuItem.Click += new System.EventHandler(this.копироватьПроектToolStripMenuItem_Click);
            // 
            // поверхностьПодОвалToolStripMenuItem
            // 
            this.поверхностьПодОвалToolStripMenuItem.Name = "поверхностьПодОвалToolStripMenuItem";
            this.поверхностьПодОвалToolStripMenuItem.Size = new System.Drawing.Size(206, 22);
            this.поверхностьПодОвалToolStripMenuItem.Text = "Поверхность под овал";
            this.поверхностьПодОвалToolStripMenuItem.Click += new System.EventHandler(this.поверхностьПодОвалToolStripMenuItem_Click);
            // 
            // автоЭскизыToolStripMenuItem
            // 
            this.автоЭскизыToolStripMenuItem.Name = "автоЭскизыToolStripMenuItem";
            this.автоЭскизыToolStripMenuItem.Size = new System.Drawing.Size(206, 22);
            this.автоЭскизыToolStripMenuItem.Text = "Эскизы по сценарию";
            this.автоЭскизыToolStripMenuItem.Click += new System.EventHandler(this.автоЭскизыToolStripMenuItem_Click);
            // 
            // автоКПToolStripMenuItem
            // 
            this.автоКПToolStripMenuItem.Name = "автоКПToolStripMenuItem";
            this.автоКПToolStripMenuItem.Size = new System.Drawing.Size(206, 22);
            this.автоКПToolStripMenuItem.Text = "Авто КП";
            this.автоКПToolStripMenuItem.Click += new System.EventHandler(this.автоКПToolStripMenuItem_Click);
            // 
            // скрытьToolStripMenuItem
            // 
            this.скрытьToolStripMenuItem.Name = "скрытьToolStripMenuItem";
            this.скрытьToolStripMenuItem.Size = new System.Drawing.Size(206, 22);
            this.скрытьToolStripMenuItem.Text = "Скрыть (B)";
            this.скрытьToolStripMenuItem.Click += new System.EventHandler(this.скрытьToolStripMenuItem_Click);
            // 
            // подавитьТелоSToolStripMenuItem
            // 
            this.подавитьТелоSToolStripMenuItem.Name = "подавитьТелоSToolStripMenuItem";
            this.подавитьТелоSToolStripMenuItem.Size = new System.Drawing.Size(206, 22);
            this.подавитьТелоSToolStripMenuItem.Text = "Подавить тело (S)";
            this.подавитьТелоSToolStripMenuItem.Click += new System.EventHandler(this.подавитьТелоSToolStripMenuItem_Click);
            // 
            // автоБлокRToolStripMenuItem
            // 
            this.автоБлокRToolStripMenuItem.Name = "автоБлокRToolStripMenuItem";
            this.автоБлокRToolStripMenuItem.Size = new System.Drawing.Size(206, 22);
            this.автоБлокRToolStripMenuItem.Text = "Авто Блок (R)";
            this.автоБлокRToolStripMenuItem.Click += new System.EventHandler(this.автоБлокRToolStripMenuItem_Click);
            // 
            // копироватьБлокToolStripMenuItem
            // 
            this.копироватьБлокToolStripMenuItem.Name = "копироватьБлокToolStripMenuItem";
            this.копироватьБлокToolStripMenuItem.Size = new System.Drawing.Size(206, 22);
            this.копироватьБлокToolStripMenuItem.Text = "Копировать блок";
            this.копироватьБлокToolStripMenuItem.Click += new System.EventHandler(this.копироватьБлокToolStripMenuItem_Click);
            // 
            // вставитьБлокToolStripMenuItem
            // 
            this.вставитьБлокToolStripMenuItem.Name = "вставитьБлокToolStripMenuItem";
            this.вставитьБлокToolStripMenuItem.Size = new System.Drawing.Size(206, 22);
            this.вставитьБлокToolStripMenuItem.Text = "Вставить блок (T)";
            this.вставитьБлокToolStripMenuItem.Click += new System.EventHandler(this.вставитьБлокToolStripMenuItem_Click);
            // 
            // объединитьToolStripMenuItem
            // 
            this.объединитьToolStripMenuItem.Name = "объединитьToolStripMenuItem";
            this.объединитьToolStripMenuItem.Size = new System.Drawing.Size(206, 22);
            this.объединитьToolStripMenuItem.Text = "Объединить";
            this.объединитьToolStripMenuItem.Click += new System.EventHandler(this.объединитьToolStripMenuItem_Click);
            // 
            // меткиToolStripMenuItem
            // 
            this.меткиToolStripMenuItem.Name = "меткиToolStripMenuItem";
            this.меткиToolStripMenuItem.Size = new System.Drawing.Size(206, 22);
            this.меткиToolStripMenuItem.Text = "Метки (M)";
            this.меткиToolStripMenuItem.Click += new System.EventHandler(this.меткиToolStripMenuItem_Click);
            // 
            // добавитьИсполнениеToolStripMenuItem1
            // 
            this.добавитьИсполнениеToolStripMenuItem1.Name = "добавитьИсполнениеToolStripMenuItem1";
            this.добавитьИсполнениеToolStripMenuItem1.Size = new System.Drawing.Size(206, 22);
            this.добавитьИсполнениеToolStripMenuItem1.Text = "Добавить исполнение";
            this.добавитьИсполнениеToolStripMenuItem1.Click += new System.EventHandler(this.добавитьИсполнениеToolStripMenuItem1_Click);
            // 
            // расширитьИсполнениеToolStripMenuItem
            // 
            this.расширитьИсполнениеToolStripMenuItem.Name = "расширитьИсполнениеToolStripMenuItem";
            this.расширитьИсполнениеToolStripMenuItem.Size = new System.Drawing.Size(206, 22);
            this.расширитьИсполнениеToolStripMenuItem.Text = "Расширить исполнение";
            this.расширитьИсполнениеToolStripMenuItem.Click += new System.EventHandler(this.расширитьИсполнениеToolStripMenuItem_Click);
            // 
            // сменитьToolStripMenuItem
            // 
            this.сменитьToolStripMenuItem.Name = "сменитьToolStripMenuItem";
            this.сменитьToolStripMenuItem.Size = new System.Drawing.Size(206, 22);
            this.сменитьToolStripMenuItem.Text = "Сменить тело (H)";
            this.сменитьToolStripMenuItem.Click += new System.EventHandler(this.сменитьToolStripMenuItem_Click);
            // 
            // масштабToolStripMenuItem
            // 
            this.масштабToolStripMenuItem.Name = "масштабToolStripMenuItem";
            this.масштабToolStripMenuItem.Size = new System.Drawing.Size(206, 22);
            this.масштабToolStripMenuItem.Text = "Масштаб";
            this.масштабToolStripMenuItem.Click += new System.EventHandler(this.масштабToolStripMenuItem_Click);
            // 
            // сборкаToolStripMenuItem
            // 
            this.сборкаToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.спецToolStripMenuItem,
            this.крепежToolStripMenuItem1,
            this.обновитьКрепежToolStripMenuItem,
            this.крепежToolStripMenuItem3,
            this.переименоватьСборкуToolStripMenuItem,
            this.кПToolStripMenuItem,
            this.вставкаПоСценариюToolStripMenuItem,
            this.совместитьКПToolStripMenuItem,
            this.названиеКПToolStripMenuItem,
            this.совместитьПлоскостиToolStripMenuItem,
            this.адаптивныеПлоскостиToolStripMenuItem,
            this.перенестиКПВСборкуToolStripMenuItem,
            this.исполненияToolStripMenuItem,
            this.добавитьУдалитьЗаменитьToolStripMenuItem});
            this.сборкаToolStripMenuItem.Name = "сборкаToolStripMenuItem";
            this.сборкаToolStripMenuItem.Size = new System.Drawing.Size(60, 20);
            this.сборкаToolStripMenuItem.Text = "Сборка";
            // 
            // спецToolStripMenuItem
            // 
            this.спецToolStripMenuItem.Name = "спецToolStripMenuItem";
            this.спецToolStripMenuItem.Size = new System.Drawing.Size(259, 22);
            this.спецToolStripMenuItem.Text = "Данные для автокрепежа";
            this.спецToolStripMenuItem.Click += new System.EventHandler(this.спецToolStripMenuItem_Click_1);
            // 
            // крепежToolStripMenuItem1
            // 
            this.крепежToolStripMenuItem1.Name = "крепежToolStripMenuItem1";
            this.крепежToolStripMenuItem1.Size = new System.Drawing.Size(259, 22);
            this.крепежToolStripMenuItem1.Text = "Автокрепеж";
            this.крепежToolStripMenuItem1.Click += new System.EventHandler(this.крепежToolStripMenuItem1_Click_1);
            // 
            // обновитьКрепежToolStripMenuItem
            // 
            this.обновитьКрепежToolStripMenuItem.Name = "обновитьКрепежToolStripMenuItem";
            this.обновитьКрепежToolStripMenuItem.Size = new System.Drawing.Size(259, 22);
            this.обновитьКрепежToolStripMenuItem.Text = "Обновить крепеж (Все открытые)";
            this.обновитьКрепежToolStripMenuItem.Click += new System.EventHandler(this.обновитьКрепежToolStripMenuItem_Click_1);
            // 
            // крепежToolStripMenuItem3
            // 
            this.крепежToolStripMenuItem3.Name = "крепежToolStripMenuItem3";
            this.крепежToolStripMenuItem3.Size = new System.Drawing.Size(259, 22);
            this.крепежToolStripMenuItem3.Text = "Крепеж (Все открытые)";
            this.крепежToolStripMenuItem3.Click += new System.EventHandler(this.крепежToolStripMenuItem3_Click_1);
            // 
            // переименоватьСборкуToolStripMenuItem
            // 
            this.переименоватьСборкуToolStripMenuItem.Name = "переименоватьСборкуToolStripMenuItem";
            this.переименоватьСборкуToolStripMenuItem.Size = new System.Drawing.Size(259, 22);
            this.переименоватьСборкуToolStripMenuItem.Text = "Переименовать сборку";
            this.переименоватьСборкуToolStripMenuItem.Click += new System.EventHandler(this.переименоватьСборкуToolStripMenuItem_Click_1);
            // 
            // кПToolStripMenuItem
            // 
            this.кПToolStripMenuItem.Name = "кПToolStripMenuItem";
            this.кПToolStripMenuItem.Size = new System.Drawing.Size(259, 22);
            this.кПToolStripMenuItem.Text = "КП";
            this.кПToolStripMenuItem.Click += new System.EventHandler(this.кПToolStripMenuItem_Click);
            // 
            // вставкаПоСценариюToolStripMenuItem
            // 
            this.вставкаПоСценариюToolStripMenuItem.Name = "вставкаПоСценариюToolStripMenuItem";
            this.вставкаПоСценариюToolStripMenuItem.Size = new System.Drawing.Size(259, 22);
            this.вставкаПоСценариюToolStripMenuItem.Text = "Вставка по сценарию (I)";
            this.вставкаПоСценариюToolStripMenuItem.Click += new System.EventHandler(this.вставкаПоСценариюToolStripMenuItem_Click);
            // 
            // совместитьКПToolStripMenuItem
            // 
            this.совместитьКПToolStripMenuItem.Name = "совместитьКПToolStripMenuItem";
            this.совместитьКПToolStripMenuItem.Size = new System.Drawing.Size(259, 22);
            this.совместитьКПToolStripMenuItem.Text = "Совместить КП (Q)";
            this.совместитьКПToolStripMenuItem.Click += new System.EventHandler(this.совместитьКПToolStripMenuItem_Click);
            // 
            // названиеКПToolStripMenuItem
            // 
            this.названиеКПToolStripMenuItem.Name = "названиеКПToolStripMenuItem";
            this.названиеКПToolStripMenuItem.Size = new System.Drawing.Size(259, 22);
            this.названиеКПToolStripMenuItem.Text = "Название КП";
            this.названиеКПToolStripMenuItem.Click += new System.EventHandler(this.названиеКПToolStripMenuItem_Click);
            // 
            // совместитьПлоскостиToolStripMenuItem
            // 
            this.совместитьПлоскостиToolStripMenuItem.Name = "совместитьПлоскостиToolStripMenuItem";
            this.совместитьПлоскостиToolStripMenuItem.Size = new System.Drawing.Size(259, 22);
            this.совместитьПлоскостиToolStripMenuItem.Text = "Совместить плоскости";
            this.совместитьПлоскостиToolStripMenuItem.Click += new System.EventHandler(this.совместитьПлоскостиToolStripMenuItem_Click);
            // 
            // адаптивныеПлоскостиToolStripMenuItem
            // 
            this.адаптивныеПлоскостиToolStripMenuItem.Name = "адаптивныеПлоскостиToolStripMenuItem";
            this.адаптивныеПлоскостиToolStripMenuItem.Size = new System.Drawing.Size(259, 22);
            this.адаптивныеПлоскостиToolStripMenuItem.Text = "Адаптивные плоскости";
            this.адаптивныеПлоскостиToolStripMenuItem.Click += new System.EventHandler(this.адаптивныеПлоскостиToolStripMenuItem_Click);
            // 
            // перенестиКПВСборкуToolStripMenuItem
            // 
            this.перенестиКПВСборкуToolStripMenuItem.Name = "перенестиКПВСборкуToolStripMenuItem";
            this.перенестиКПВСборкуToolStripMenuItem.Size = new System.Drawing.Size(259, 22);
            this.перенестиКПВСборкуToolStripMenuItem.Text = "Перенести КП в сборку";
            this.перенестиКПВСборкуToolStripMenuItem.Click += new System.EventHandler(this.перенестиКПВСборкуToolStripMenuItem_Click);
            // 
            // исполненияToolStripMenuItem
            // 
            this.исполненияToolStripMenuItem.Name = "исполненияToolStripMenuItem";
            this.исполненияToolStripMenuItem.Size = new System.Drawing.Size(259, 22);
            this.исполненияToolStripMenuItem.Text = "Исполнения";
            this.исполненияToolStripMenuItem.Click += new System.EventHandler(this.исполненияToolStripMenuItem_Click);
            // 
            // добавитьУдалитьЗаменитьToolStripMenuItem
            // 
            this.добавитьУдалитьЗаменитьToolStripMenuItem.Name = "добавитьУдалитьЗаменитьToolStripMenuItem";
            this.добавитьУдалитьЗаменитьToolStripMenuItem.Size = new System.Drawing.Size(259, 22);
            this.добавитьУдалитьЗаменитьToolStripMenuItem.Text = "ДобавитьУдалитьЗаменить";
            this.добавитьУдалитьЗаменитьToolStripMenuItem.Click += new System.EventHandler(this.добавитьУдалитьЗаменитьToolStripMenuItem_Click);
            // 
            // чертежToolStripMenuItem
            // 
            this.чертежToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.текущийДокументToolStripMenuItem,
            this.регионToolStripMenuItem,
            this.обновитьПозицииToolStripMenuItem,
            this.подготовитьШаблонToolStripMenuItem,
            this.видыПоШаблонуToolStripMenuItem,
            this.позицииПоШаблонуToolStripMenuItem,
            this.размерыПоШаблонуToolStripMenuItem,
            this.осевыеПоШаблонуToolStripMenuItem,
            this.addTRToolStripMenuItem1,
            this.copyAttrToolStripMenuItem1,
            this.стильРазмеровСборкиToolStripMenuItem,
            this.получитьТекстToolStripMenuItem,
            this.перевестиТекстToolStripMenuItem,
            this.обновитьСтилиToolStripMenuItem,
            this.инициализацияToolStripMenuItem,
            this.переназватьToolStripMenuItem,
            this.пересобратьСборкуToolStripMenuItem,
            this.разрывВидаToolStripMenuItem,
            this.перенаправитьToolStripMenuItem1,
            this.заменитьМодельZToolStripMenuItem,
            this.связатьКрепежToolStripMenuItem});
            this.чертежToolStripMenuItem.Name = "чертежToolStripMenuItem";
            this.чертежToolStripMenuItem.Size = new System.Drawing.Size(60, 20);
            this.чертежToolStripMenuItem.Text = "Чертеж";
            // 
            // текущийДокументToolStripMenuItem
            // 
            this.текущийДокументToolStripMenuItem.Name = "текущийДокументToolStripMenuItem";
            this.текущийДокументToolStripMenuItem.Size = new System.Drawing.Size(278, 22);
            this.текущийДокументToolStripMenuItem.Text = "Создать чертежи (D)";
            this.текущийДокументToolStripMenuItem.Click += new System.EventHandler(this.текущийДокументToolStripMenuItem_Click);
            // 
            // регионToolStripMenuItem
            // 
            this.регионToolStripMenuItem.Name = "регионToolStripMenuItem";
            this.регионToolStripMenuItem.Size = new System.Drawing.Size(278, 22);
            this.регионToolStripMenuItem.Text = "Выровнять позиции";
            this.регионToolStripMenuItem.Click += new System.EventHandler(this.регионToolStripMenuItem_Click_1);
            // 
            // обновитьПозицииToolStripMenuItem
            // 
            this.обновитьПозицииToolStripMenuItem.Name = "обновитьПозицииToolStripMenuItem";
            this.обновитьПозицииToolStripMenuItem.Size = new System.Drawing.Size(278, 22);
            this.обновитьПозицииToolStripMenuItem.Text = "Обновить позиции";
            this.обновитьПозицииToolStripMenuItem.Click += new System.EventHandler(this.обновитьПозицииToolStripMenuItem_Click_1);
            // 
            // подготовитьШаблонToolStripMenuItem
            // 
            this.подготовитьШаблонToolStripMenuItem.Name = "подготовитьШаблонToolStripMenuItem";
            this.подготовитьШаблонToolStripMenuItem.Size = new System.Drawing.Size(278, 22);
            this.подготовитьШаблонToolStripMenuItem.Text = "Подготовить шаблон";
            this.подготовитьШаблонToolStripMenuItem.Click += new System.EventHandler(this.подготовитьШаблонToolStripMenuItem_Click);
            // 
            // видыПоШаблонуToolStripMenuItem
            // 
            this.видыПоШаблонуToolStripMenuItem.Name = "видыПоШаблонуToolStripMenuItem";
            this.видыПоШаблонуToolStripMenuItem.Size = new System.Drawing.Size(278, 22);
            this.видыПоШаблонуToolStripMenuItem.Text = "Виды по шаблону";
            this.видыПоШаблонуToolStripMenuItem.Click += new System.EventHandler(this.видыПоШаблонуToolStripMenuItem_Click);
            // 
            // позицииПоШаблонуToolStripMenuItem
            // 
            this.позицииПоШаблонуToolStripMenuItem.Name = "позицииПоШаблонуToolStripMenuItem";
            this.позицииПоШаблонуToolStripMenuItem.Size = new System.Drawing.Size(278, 22);
            this.позицииПоШаблонуToolStripMenuItem.Text = "Позиции по шаблону";
            this.позицииПоШаблонуToolStripMenuItem.Click += new System.EventHandler(this.позицииПоШаблонуToolStripMenuItem_Click);
            // 
            // размерыПоШаблонуToolStripMenuItem
            // 
            this.размерыПоШаблонуToolStripMenuItem.Name = "размерыПоШаблонуToolStripMenuItem";
            this.размерыПоШаблонуToolStripMenuItem.Size = new System.Drawing.Size(278, 22);
            this.размерыПоШаблонуToolStripMenuItem.Text = "Размеры по шаблону";
            this.размерыПоШаблонуToolStripMenuItem.Click += new System.EventHandler(this.размерыПоШаблонуToolStripMenuItem_Click);
            // 
            // осевыеПоШаблонуToolStripMenuItem
            // 
            this.осевыеПоШаблонуToolStripMenuItem.Name = "осевыеПоШаблонуToolStripMenuItem";
            this.осевыеПоШаблонуToolStripMenuItem.Size = new System.Drawing.Size(278, 22);
            this.осевыеПоШаблонуToolStripMenuItem.Text = "Осевые по шаблону";
            this.осевыеПоШаблонуToolStripMenuItem.Click += new System.EventHandler(this.осевыеПоШаблонуToolStripMenuItem_Click);
            // 
            // addTRToolStripMenuItem1
            // 
            this.addTRToolStripMenuItem1.Name = "addTRToolStripMenuItem1";
            this.addTRToolStripMenuItem1.Size = new System.Drawing.Size(278, 22);
            this.addTRToolStripMenuItem1.Text = "Добавить технические требования";
            this.addTRToolStripMenuItem1.Click += new System.EventHandler(this.addTRToolStripMenuItem1_Click_1);
            // 
            // copyAttrToolStripMenuItem1
            // 
            this.copyAttrToolStripMenuItem1.Name = "copyAttrToolStripMenuItem1";
            this.copyAttrToolStripMenuItem1.Size = new System.Drawing.Size(278, 22);
            this.copyAttrToolStripMenuItem1.Text = "Копировать технические требования";
            this.copyAttrToolStripMenuItem1.Click += new System.EventHandler(this.copyAttrToolStripMenuItem1_Click_1);
            // 
            // стильРазмеровСборкиToolStripMenuItem
            // 
            this.стильРазмеровСборкиToolStripMenuItem.Name = "стильРазмеровСборкиToolStripMenuItem";
            this.стильРазмеровСборкиToolStripMenuItem.Size = new System.Drawing.Size(278, 22);
            this.стильРазмеровСборкиToolStripMenuItem.Text = "Стиль размеров сборки";
            this.стильРазмеровСборкиToolStripMenuItem.Click += new System.EventHandler(this.стильРазмеровСборкиToolStripMenuItem_Click);
            // 
            // получитьТекстToolStripMenuItem
            // 
            this.получитьТекстToolStripMenuItem.Name = "получитьТекстToolStripMenuItem";
            this.получитьТекстToolStripMenuItem.Size = new System.Drawing.Size(278, 22);
            this.получитьТекстToolStripMenuItem.Text = "Получить текст";
            this.получитьТекстToolStripMenuItem.Click += new System.EventHandler(this.получитьТекстToolStripMenuItem_Click_1);
            // 
            // перевестиТекстToolStripMenuItem
            // 
            this.перевестиТекстToolStripMenuItem.Name = "перевестиТекстToolStripMenuItem";
            this.перевестиТекстToolStripMenuItem.Size = new System.Drawing.Size(278, 22);
            this.перевестиТекстToolStripMenuItem.Text = "Перевести текст";
            this.перевестиТекстToolStripMenuItem.Click += new System.EventHandler(this.перевестиТекстToolStripMenuItem_Click_1);
            // 
            // обновитьСтилиToolStripMenuItem
            // 
            this.обновитьСтилиToolStripMenuItem.Name = "обновитьСтилиToolStripMenuItem";
            this.обновитьСтилиToolStripMenuItem.Size = new System.Drawing.Size(278, 22);
            this.обновитьСтилиToolStripMenuItem.Text = "Обновить стили";
            this.обновитьСтилиToolStripMenuItem.Click += new System.EventHandler(this.обновитьСтилиToolStripMenuItem_Click_1);
            // 
            // инициализацияToolStripMenuItem
            // 
            this.инициализацияToolStripMenuItem.Name = "инициализацияToolStripMenuItem";
            this.инициализацияToolStripMenuItem.Size = new System.Drawing.Size(278, 22);
            this.инициализацияToolStripMenuItem.Text = "Инициализация";
            this.инициализацияToolStripMenuItem.Click += new System.EventHandler(this.инициализацияToolStripMenuItem_Click);
            // 
            // переназватьToolStripMenuItem
            // 
            this.переназватьToolStripMenuItem.Name = "переназватьToolStripMenuItem";
            this.переназватьToolStripMenuItem.Size = new System.Drawing.Size(278, 22);
            this.переназватьToolStripMenuItem.Text = "Переименовать";
            this.переназватьToolStripMenuItem.Click += new System.EventHandler(this.переназватьToolStripMenuItem_Click);
            // 
            // пересобратьСборкуToolStripMenuItem
            // 
            this.пересобратьСборкуToolStripMenuItem.Name = "пересобратьСборкуToolStripMenuItem";
            this.пересобратьСборкуToolStripMenuItem.Size = new System.Drawing.Size(278, 22);
            this.пересобратьСборкуToolStripMenuItem.Text = "Пересобрать сборку";
            this.пересобратьСборкуToolStripMenuItem.Click += new System.EventHandler(this.пересобратьСборкуToolStripMenuItem_Click);
            // 
            // разрывВидаToolStripMenuItem
            // 
            this.разрывВидаToolStripMenuItem.Name = "разрывВидаToolStripMenuItem";
            this.разрывВидаToolStripMenuItem.Size = new System.Drawing.Size(278, 22);
            this.разрывВидаToolStripMenuItem.Text = "Разрыв вида (B)";
            this.разрывВидаToolStripMenuItem.Click += new System.EventHandler(this.разрывВидаToolStripMenuItem_Click);
            // 
            // перенаправитьToolStripMenuItem1
            // 
            this.перенаправитьToolStripMenuItem1.Name = "перенаправитьToolStripMenuItem1";
            this.перенаправитьToolStripMenuItem1.Size = new System.Drawing.Size(278, 22);
            this.перенаправитьToolStripMenuItem1.Text = "Перенаправить";
            this.перенаправитьToolStripMenuItem1.Click += new System.EventHandler(this.перенаправитьToolStripMenuItem1_Click_1);
            // 
            // заменитьМодельZToolStripMenuItem
            // 
            this.заменитьМодельZToolStripMenuItem.Name = "заменитьМодельZToolStripMenuItem";
            this.заменитьМодельZToolStripMenuItem.Size = new System.Drawing.Size(278, 22);
            this.заменитьМодельZToolStripMenuItem.Text = "Заменить модель (Z)";
            this.заменитьМодельZToolStripMenuItem.Click += new System.EventHandler(this.заменитьМодельZToolStripMenuItem_Click);
            // 
            // связатьКрепежToolStripMenuItem
            // 
            this.связатьКрепежToolStripMenuItem.Name = "связатьКрепежToolStripMenuItem";
            this.связатьКрепежToolStripMenuItem.Size = new System.Drawing.Size(278, 22);
            this.связатьКрепежToolStripMenuItem.Text = "Связать крепеж";
            this.связатьКрепежToolStripMenuItem.Click += new System.EventHandler(this.связатьКрепежToolStripMenuItem_Click);
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
            this.создатьЭскизыToolStripMenuItem,
            this.создатьДеталиToolStripMenuItem,
            this.dPDFToolStripMenuItem});
            this.параметрыToolStripMenuItem.Name = "параметрыToolStripMenuItem";
            this.параметрыToolStripMenuItem.Size = new System.Drawing.Size(43, 20);
            this.параметрыToolStripMenuItem.Text = "XML";
            // 
            // добавитьПараметрыToolStripMenuItem1
            // 
            this.добавитьПараметрыToolStripMenuItem1.Name = "добавитьПараметрыToolStripMenuItem1";
            this.добавитьПараметрыToolStripMenuItem1.Size = new System.Drawing.Size(191, 22);
            this.добавитьПараметрыToolStripMenuItem1.Text = "Добавить параметры";
            this.добавитьПараметрыToolStripMenuItem1.Click += new System.EventHandler(this.добавитьПараметрыToolStripMenuItem_Click);
            // 
            // загрузитьXMLToolStripMenuItem
            // 
            this.загрузитьXMLToolStripMenuItem.Name = "загрузитьXMLToolStripMenuItem";
            this.загрузитьXMLToolStripMenuItem.Size = new System.Drawing.Size(191, 22);
            this.загрузитьXMLToolStripMenuItem.Text = "Загрузить XML";
            this.загрузитьXMLToolStripMenuItem.Click += new System.EventHandler(this.загрузитьXMLToolStripMenuItem_Click);
            // 
            // данныеДляМассиваToolStripMenuItem
            // 
            this.данныеДляМассиваToolStripMenuItem.Name = "данныеДляМассиваToolStripMenuItem";
            this.данныеДляМассиваToolStripMenuItem.Size = new System.Drawing.Size(191, 22);
            this.данныеДляМассиваToolStripMenuItem.Text = "Данные для массива";
            this.данныеДляМассиваToolStripMenuItem.Click += new System.EventHandler(this.добавитьДанныеДляМассиваToolStripMenuItem_Click);
            // 
            // массивToolStripMenuItem1
            // 
            this.массивToolStripMenuItem1.Name = "массивToolStripMenuItem1";
            this.массивToolStripMenuItem1.Size = new System.Drawing.Size(191, 22);
            this.массивToolStripMenuItem1.Text = "Массив";
            this.массивToolStripMenuItem1.Click += new System.EventHandler(this.массивToolStripMenuItem_Click);
            // 
            // копироватьВXMLToolStripMenuItem
            // 
            this.копироватьВXMLToolStripMenuItem.Name = "копироватьВXMLToolStripMenuItem";
            this.копироватьВXMLToolStripMenuItem.Size = new System.Drawing.Size(191, 22);
            this.копироватьВXMLToolStripMenuItem.Text = "Копировать в XML";
            this.копироватьВXMLToolStripMenuItem.Click += new System.EventHandler(this.копироватьВXMLToolStripMenuItem_Click);
            // 
            // структураВXMLToolStripMenuItem
            // 
            this.структураВXMLToolStripMenuItem.Name = "структураВXMLToolStripMenuItem";
            this.структураВXMLToolStripMenuItem.Size = new System.Drawing.Size(191, 22);
            this.структураВXMLToolStripMenuItem.Text = "Структура в XML";
            this.структураВXMLToolStripMenuItem.Click += new System.EventHandler(this.структураВXMLToolStripMenuItem_Click);
            // 
            // создатьЭскизыToolStripMenuItem
            // 
            this.создатьЭскизыToolStripMenuItem.Name = "создатьЭскизыToolStripMenuItem";
            this.создатьЭскизыToolStripMenuItem.Size = new System.Drawing.Size(191, 22);
            this.создатьЭскизыToolStripMenuItem.Text = "Создать эскизы";
            this.создатьЭскизыToolStripMenuItem.Click += new System.EventHandler(this.создатьЭскизыToolStripMenuItem_Click);
            // 
            // создатьДеталиToolStripMenuItem
            // 
            this.создатьДеталиToolStripMenuItem.Name = "создатьДеталиToolStripMenuItem";
            this.создатьДеталиToolStripMenuItem.Size = new System.Drawing.Size(191, 22);
            this.создатьДеталиToolStripMenuItem.Text = "Создать детали";
            this.создатьДеталиToolStripMenuItem.Click += new System.EventHandler(this.создатьДеталиToolStripMenuItem_Click);
            // 
            // dPDFToolStripMenuItem
            // 
            this.dPDFToolStripMenuItem.Name = "dPDFToolStripMenuItem";
            this.dPDFToolStripMenuItem.Size = new System.Drawing.Size(191, 22);
            this.dPDFToolStripMenuItem.Text = "3DPDF";
            this.dPDFToolStripMenuItem.Click += new System.EventHandler(this.dPDFToolStripMenuItem_Click);
            // 
            // разместитьВсеДеталиToolStripMenuItem
            // 
            this.разместитьВсеДеталиToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.разместитьВсеДеталиToolStripMenuItem1,
            this.обновитьContentCenterToolStripMenuItem,
            this.сортировкаToolStripMenuItem,
            this.тестToolStripMenuItem,
            this.тест2ToolStripMenuItem,
            this.копироватьToolStripMenuItem,
            this.восстановитьToolStripMenuItem,
            this.данныеИзTitleToolStripMenuItem,
            this.сохранитьDWGToolStripMenuItem,
            this.граниВDXFToolStripMenuItem,
            this.тестSplToolStripMenuItem});
            this.разместитьВсеДеталиToolStripMenuItem.Name = "разместитьВсеДеталиToolStripMenuItem";
            this.разместитьВсеДеталиToolStripMenuItem.Size = new System.Drawing.Size(84, 20);
            this.разместитьВсеДеталиToolStripMenuItem.Text = "Библиотека";
            this.разместитьВсеДеталиToolStripMenuItem.Click += new System.EventHandler(this.создатьДеревоToolStripMenuItem_Click);
            // 
            // разместитьВсеДеталиToolStripMenuItem1
            // 
            this.разместитьВсеДеталиToolStripMenuItem1.Name = "разместитьВсеДеталиToolStripMenuItem1";
            this.разместитьВсеДеталиToolStripMenuItem1.Size = new System.Drawing.Size(209, 22);
            this.разместитьВсеДеталиToolStripMenuItem1.Text = "Разместить все детали";
            this.разместитьВсеДеталиToolStripMenuItem1.Click += new System.EventHandler(this.разместитьВсеДеталиToolStripMenuItem1_Click);
            // 
            // обновитьContentCenterToolStripMenuItem
            // 
            this.обновитьContentCenterToolStripMenuItem.Name = "обновитьContentCenterToolStripMenuItem";
            this.обновитьContentCenterToolStripMenuItem.Size = new System.Drawing.Size(209, 22);
            this.обновитьContentCenterToolStripMenuItem.Text = "Обновить ContentCenter";
            this.обновитьContentCenterToolStripMenuItem.Click += new System.EventHandler(this.обновитьContentCenterToolStripMenuItem_Click);
            // 
            // сортировкаToolStripMenuItem
            // 
            this.сортировкаToolStripMenuItem.Name = "сортировкаToolStripMenuItem";
            this.сортировкаToolStripMenuItem.Size = new System.Drawing.Size(209, 22);
            this.сортировкаToolStripMenuItem.Text = "Сортировка";
            this.сортировкаToolStripMenuItem.Click += new System.EventHandler(this.сортировкаToolStripMenuItem_Click);
            // 
            // тестToolStripMenuItem
            // 
            this.тестToolStripMenuItem.Name = "тестToolStripMenuItem";
            this.тестToolStripMenuItem.Size = new System.Drawing.Size(209, 22);
            this.тестToolStripMenuItem.Text = "Тест";
            this.тестToolStripMenuItem.Click += new System.EventHandler(this.тестToolStripMenuItem_Click);
            // 
            // тест2ToolStripMenuItem
            // 
            this.тест2ToolStripMenuItem.Name = "тест2ToolStripMenuItem";
            this.тест2ToolStripMenuItem.Size = new System.Drawing.Size(209, 22);
            this.тест2ToolStripMenuItem.Text = "Тест2";
            this.тест2ToolStripMenuItem.Click += new System.EventHandler(this.тест2ToolStripMenuItem_Click);
            // 
            // копироватьToolStripMenuItem
            // 
            this.копироватьToolStripMenuItem.Name = "копироватьToolStripMenuItem";
            this.копироватьToolStripMenuItem.Size = new System.Drawing.Size(209, 22);
            this.копироватьToolStripMenuItem.Text = "Копировать";
            this.копироватьToolStripMenuItem.Click += new System.EventHandler(this.копироватьToolStripMenuItem_Click);
            // 
            // восстановитьToolStripMenuItem
            // 
            this.восстановитьToolStripMenuItem.Name = "восстановитьToolStripMenuItem";
            this.восстановитьToolStripMenuItem.Size = new System.Drawing.Size(209, 22);
            this.восстановитьToolStripMenuItem.Text = "Восстановить";
            this.восстановитьToolStripMenuItem.Click += new System.EventHandler(this.восстановитьToolStripMenuItem_Click);
            // 
            // данныеИзTitleToolStripMenuItem
            // 
            this.данныеИзTitleToolStripMenuItem.Name = "данныеИзTitleToolStripMenuItem";
            this.данныеИзTitleToolStripMenuItem.Size = new System.Drawing.Size(209, 22);
            this.данныеИзTitleToolStripMenuItem.Text = "Данные из Title";
            this.данныеИзTitleToolStripMenuItem.Click += new System.EventHandler(this.данныеИзTitleToolStripMenuItem_Click);
            // 
            // сохранитьDWGToolStripMenuItem
            // 
            this.сохранитьDWGToolStripMenuItem.Name = "сохранитьDWGToolStripMenuItem";
            this.сохранитьDWGToolStripMenuItem.Size = new System.Drawing.Size(209, 22);
            this.сохранитьDWGToolStripMenuItem.Text = "Сохранить DWG";
            this.сохранитьDWGToolStripMenuItem.Click += new System.EventHandler(this.сохранитьDWGToolStripMenuItem_Click);
            // 
            // граниВDXFToolStripMenuItem
            // 
            this.граниВDXFToolStripMenuItem.Name = "граниВDXFToolStripMenuItem";
            this.граниВDXFToolStripMenuItem.Size = new System.Drawing.Size(209, 22);
            this.граниВDXFToolStripMenuItem.Text = "Грани в DXF";
            this.граниВDXFToolStripMenuItem.Click += new System.EventHandler(this.граниВDXFToolStripMenuItem_Click);
            // 
            // тестSplToolStripMenuItem
            // 
            this.тестSplToolStripMenuItem.Name = "тестSplToolStripMenuItem";
            this.тестSplToolStripMenuItem.Size = new System.Drawing.Size(209, 22);
            this.тестSplToolStripMenuItem.Text = "тест spl";
            this.тестSplToolStripMenuItem.Click += new System.EventHandler(this.тестSplToolStripMenuItem_Click);
            // 
            // редкоеToolStripMenuItem
            // 
            this.редкоеToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.stepToolStripMenuItem,
            this.комплектФайловToolStripMenuItem1,
            this.комплектФайловСЧертежамиToolStripMenuItem,
            this.комплектНесколькихФайловToolStripMenuItem,
            this.обновитьToolStripMenuItem,
            this.добавитьСвойствоToolStripMenuItem1,
            this.свойствоПоШаблонуToolStripMenuItem,
            this.добавитьИсполнениеToolStripMenuItem,
            this.сплайнToolStripMenuItem,
            this.артикулToolStripMenuItem,
            this.центрыОтверстийToolStripMenuItem,
            this.создатьСборкуToolStripMenuItem,
            this.заменитьИмяКПToolStripMenuItem,
            this.копироватьВПапкуToolStripMenuItem,
            this.экспортВSATToolStripMenuItem,
            this.экспортВSTPToolStripMenuItem,
            this.габаритыToolStripMenuItem,
            this.открытьПапкуToolStripMenuItem,
            this.выравниваниеОсейToolStripMenuItem,
            this.перенестиВЦентрToolStripMenuItem,
            this.ширинаНадписиToolStripMenuItem,
            this.видимостьToolStripMenuItem,
            this.связатьВPDFToolStripMenuItem,
            this.pDFЗеркальныеДеталиToolStripMenuItem,
            this.удалитьПробелыToolStripMenuItem,
            this.получитьНазванияToolStripMenuItem,
            this.переназватьToolStripMenuItem1,
            this.открытьПоШаблонуToolStripMenuItem,
            this.экспортВOBJToolStripMenuItem,
            this.экспортВGLBОдноТелоToolStripMenuItem,
            this.экспортВGLBСборкаToolStripMenuItem,
            this.восстановитьToolStripMenuItem1,
            this.оЦToolStripMenuItem,
            this.изменитьОтверстияToolStripMenuItem});
            this.редкоеToolStripMenuItem.Name = "редкоеToolStripMenuItem";
            this.редкоеToolStripMenuItem.Size = new System.Drawing.Size(58, 20);
            this.редкоеToolStripMenuItem.Text = "Общее";
            // 
            // stepToolStripMenuItem
            // 
            this.stepToolStripMenuItem.Name = "stepToolStripMenuItem";
            this.stepToolStripMenuItem.Size = new System.Drawing.Size(247, 22);
            this.stepToolStripMenuItem.Text = "3DPDF";
            this.stepToolStripMenuItem.Click += new System.EventHandler(this.stepToolStripMenuItem_Click);
            // 
            // комплектФайловToolStripMenuItem1
            // 
            this.комплектФайловToolStripMenuItem1.Name = "комплектФайловToolStripMenuItem1";
            this.комплектФайловToolStripMenuItem1.Size = new System.Drawing.Size(247, 22);
            this.комплектФайловToolStripMenuItem1.Text = "Комплект файлов";
            this.комплектФайловToolStripMenuItem1.Click += new System.EventHandler(this.комплектФайловToolStripMenuItem1_Click);
            // 
            // комплектФайловСЧертежамиToolStripMenuItem
            // 
            this.комплектФайловСЧертежамиToolStripMenuItem.Name = "комплектФайловСЧертежамиToolStripMenuItem";
            this.комплектФайловСЧертежамиToolStripMenuItem.Size = new System.Drawing.Size(247, 22);
            this.комплектФайловСЧертежамиToolStripMenuItem.Text = "Комплект файлов с чертежами";
            this.комплектФайловСЧертежамиToolStripMenuItem.Click += new System.EventHandler(this.комплектФайловСЧертежамиToolStripMenuItem_Click);
            // 
            // комплектНесколькихФайловToolStripMenuItem
            // 
            this.комплектНесколькихФайловToolStripMenuItem.Name = "комплектНесколькихФайловToolStripMenuItem";
            this.комплектНесколькихФайловToolStripMenuItem.Size = new System.Drawing.Size(247, 22);
            this.комплектНесколькихФайловToolStripMenuItem.Text = "Комплект нескольких файлов";
            this.комплектНесколькихФайловToolStripMenuItem.Click += new System.EventHandler(this.комплектНесколькихФайловToolStripMenuItem_Click);
            // 
            // обновитьToolStripMenuItem
            // 
            this.обновитьToolStripMenuItem.Name = "обновитьToolStripMenuItem";
            this.обновитьToolStripMenuItem.Size = new System.Drawing.Size(247, 22);
            this.обновитьToolStripMenuItem.Text = "Обновить";
            this.обновитьToolStripMenuItem.Click += new System.EventHandler(this.обновитьToolStripMenuItem_Click_1);
            // 
            // добавитьСвойствоToolStripMenuItem1
            // 
            this.добавитьСвойствоToolStripMenuItem1.Name = "добавитьСвойствоToolStripMenuItem1";
            this.добавитьСвойствоToolStripMenuItem1.Size = new System.Drawing.Size(247, 22);
            this.добавитьСвойствоToolStripMenuItem1.Text = "Добавить свойство";
            this.добавитьСвойствоToolStripMenuItem1.Click += new System.EventHandler(this.добавитьСвойствоToolStripMenuItem1_Click);
            // 
            // свойствоПоШаблонуToolStripMenuItem
            // 
            this.свойствоПоШаблонуToolStripMenuItem.Name = "свойствоПоШаблонуToolStripMenuItem";
            this.свойствоПоШаблонуToolStripMenuItem.Size = new System.Drawing.Size(247, 22);
            this.свойствоПоШаблонуToolStripMenuItem.Text = "Свойство по шаблону";
            this.свойствоПоШаблонуToolStripMenuItem.Click += new System.EventHandler(this.свойствоПоШаблонуToolStripMenuItem_Click);
            // 
            // добавитьИсполнениеToolStripMenuItem
            // 
            this.добавитьИсполнениеToolStripMenuItem.Name = "добавитьИсполнениеToolStripMenuItem";
            this.добавитьИсполнениеToolStripMenuItem.Size = new System.Drawing.Size(247, 22);
            this.добавитьИсполнениеToolStripMenuItem.Text = "Добавить исполнение";
            this.добавитьИсполнениеToolStripMenuItem.Click += new System.EventHandler(this.добавитьИсполнениеToolStripMenuItem_Click);
            // 
            // сплайнToolStripMenuItem
            // 
            this.сплайнToolStripMenuItem.Name = "сплайнToolStripMenuItem";
            this.сплайнToolStripMenuItem.Size = new System.Drawing.Size(247, 22);
            this.сплайнToolStripMenuItem.Text = "Сплайн";
            this.сплайнToolStripMenuItem.Click += new System.EventHandler(this.сплайнToolStripMenuItem_Click);
            // 
            // артикулToolStripMenuItem
            // 
            this.артикулToolStripMenuItem.Name = "артикулToolStripMenuItem";
            this.артикулToolStripMenuItem.Size = new System.Drawing.Size(247, 22);
            this.артикулToolStripMenuItem.Text = "Артикул";
            this.артикулToolStripMenuItem.Click += new System.EventHandler(this.артикулToolStripMenuItem_Click);
            // 
            // центрыОтверстийToolStripMenuItem
            // 
            this.центрыОтверстийToolStripMenuItem.Name = "центрыОтверстийToolStripMenuItem";
            this.центрыОтверстийToolStripMenuItem.Size = new System.Drawing.Size(247, 22);
            this.центрыОтверстийToolStripMenuItem.Text = "Центры отверстий";
            this.центрыОтверстийToolStripMenuItem.Click += new System.EventHandler(this.центрыОтверстийToolStripMenuItem_Click);
            // 
            // создатьСборкуToolStripMenuItem
            // 
            this.создатьСборкуToolStripMenuItem.Name = "создатьСборкуToolStripMenuItem";
            this.создатьСборкуToolStripMenuItem.Size = new System.Drawing.Size(247, 22);
            this.создатьСборкуToolStripMenuItem.Text = "Создать сборку";
            this.создатьСборкуToolStripMenuItem.Click += new System.EventHandler(this.создатьСборкуToolStripMenuItem_Click);
            // 
            // заменитьИмяКПToolStripMenuItem
            // 
            this.заменитьИмяКПToolStripMenuItem.Name = "заменитьИмяКПToolStripMenuItem";
            this.заменитьИмяКПToolStripMenuItem.Size = new System.Drawing.Size(247, 22);
            this.заменитьИмяКПToolStripMenuItem.Text = "Заменить имя КП";
            this.заменитьИмяКПToolStripMenuItem.Click += new System.EventHandler(this.заменитьИмяКПToolStripMenuItem_Click);
            // 
            // копироватьВПапкуToolStripMenuItem
            // 
            this.копироватьВПапкуToolStripMenuItem.Name = "копироватьВПапкуToolStripMenuItem";
            this.копироватьВПапкуToolStripMenuItem.Size = new System.Drawing.Size(247, 22);
            this.копироватьВПапкуToolStripMenuItem.Text = "Копировать в папку";
            this.копироватьВПапкуToolStripMenuItem.Click += new System.EventHandler(this.копироватьВПапкуToolStripMenuItem_Click);
            // 
            // экспортВSATToolStripMenuItem
            // 
            this.экспортВSATToolStripMenuItem.Name = "экспортВSATToolStripMenuItem";
            this.экспортВSATToolStripMenuItem.Size = new System.Drawing.Size(247, 22);
            this.экспортВSATToolStripMenuItem.Text = "Экспорт в SAT";
            this.экспортВSATToolStripMenuItem.Click += new System.EventHandler(this.экспортВSATToolStripMenuItem_Click);
            // 
            // экспортВSTPToolStripMenuItem
            // 
            this.экспортВSTPToolStripMenuItem.Name = "экспортВSTPToolStripMenuItem";
            this.экспортВSTPToolStripMenuItem.Size = new System.Drawing.Size(247, 22);
            this.экспортВSTPToolStripMenuItem.Text = "Экспорт в STP";
            this.экспортВSTPToolStripMenuItem.Click += new System.EventHandler(this.экспортВSTPToolStripMenuItem_Click);
            // 
            // габаритыToolStripMenuItem
            // 
            this.габаритыToolStripMenuItem.Name = "габаритыToolStripMenuItem";
            this.габаритыToolStripMenuItem.Size = new System.Drawing.Size(247, 22);
            this.габаритыToolStripMenuItem.Text = "Габариты";
            this.габаритыToolStripMenuItem.Click += new System.EventHandler(this.габаритыToolStripMenuItem_Click_1);
            // 
            // открытьПапкуToolStripMenuItem
            // 
            this.открытьПапкуToolStripMenuItem.Name = "открытьПапкуToolStripMenuItem";
            this.открытьПапкуToolStripMenuItem.Size = new System.Drawing.Size(247, 22);
            this.открытьПапкуToolStripMenuItem.Text = "Открыть папку (O)";
            this.открытьПапкуToolStripMenuItem.Click += new System.EventHandler(this.открытьПапкуToolStripMenuItem_Click_1);
            // 
            // выравниваниеОсейToolStripMenuItem
            // 
            this.выравниваниеОсейToolStripMenuItem.Name = "выравниваниеОсейToolStripMenuItem";
            this.выравниваниеОсейToolStripMenuItem.Size = new System.Drawing.Size(247, 22);
            this.выравниваниеОсейToolStripMenuItem.Text = "Выравнивание осей";
            this.выравниваниеОсейToolStripMenuItem.Click += new System.EventHandler(this.выравниваниеОсейToolStripMenuItem_Click);
            // 
            // перенестиВЦентрToolStripMenuItem
            // 
            this.перенестиВЦентрToolStripMenuItem.Name = "перенестиВЦентрToolStripMenuItem";
            this.перенестиВЦентрToolStripMenuItem.Size = new System.Drawing.Size(247, 22);
            this.перенестиВЦентрToolStripMenuItem.Text = "Перенести в центр";
            this.перенестиВЦентрToolStripMenuItem.Click += new System.EventHandler(this.перенестиВЦентрToolStripMenuItem_Click);
            // 
            // ширинаНадписиToolStripMenuItem
            // 
            this.ширинаНадписиToolStripMenuItem.Name = "ширинаНадписиToolStripMenuItem";
            this.ширинаНадписиToolStripMenuItem.Size = new System.Drawing.Size(247, 22);
            this.ширинаНадписиToolStripMenuItem.Text = "Ширина надписи";
            this.ширинаНадписиToolStripMenuItem.Click += new System.EventHandler(this.ширинаНадписиToolStripMenuItem_Click);
            // 
            // видимостьToolStripMenuItem
            // 
            this.видимостьToolStripMenuItem.Name = "видимостьToolStripMenuItem";
            this.видимостьToolStripMenuItem.Size = new System.Drawing.Size(247, 22);
            this.видимостьToolStripMenuItem.Text = "Видимость (V)";
            this.видимостьToolStripMenuItem.Click += new System.EventHandler(this.видимостьToolStripMenuItem_Click);
            // 
            // связатьВPDFToolStripMenuItem
            // 
            this.связатьВPDFToolStripMenuItem.Name = "связатьВPDFToolStripMenuItem";
            this.связатьВPDFToolStripMenuItem.Size = new System.Drawing.Size(247, 22);
            this.связатьВPDFToolStripMenuItem.Text = "Связать в PDF";
            this.связатьВPDFToolStripMenuItem.Click += new System.EventHandler(this.связатьВPDFToolStripMenuItem_Click);
            // 
            // pDFЗеркальныеДеталиToolStripMenuItem
            // 
            this.pDFЗеркальныеДеталиToolStripMenuItem.Name = "pDFЗеркальныеДеталиToolStripMenuItem";
            this.pDFЗеркальныеДеталиToolStripMenuItem.Size = new System.Drawing.Size(247, 22);
            this.pDFЗеркальныеДеталиToolStripMenuItem.Text = "PDF зеркальные детали";
            this.pDFЗеркальныеДеталиToolStripMenuItem.Click += new System.EventHandler(this.pDFЗеркальныеДеталиToolStripMenuItem_Click);
            // 
            // удалитьПробелыToolStripMenuItem
            // 
            this.удалитьПробелыToolStripMenuItem.Name = "удалитьПробелыToolStripMenuItem";
            this.удалитьПробелыToolStripMenuItem.Size = new System.Drawing.Size(247, 22);
            this.удалитьПробелыToolStripMenuItem.Text = "Удалить пробелы";
            this.удалитьПробелыToolStripMenuItem.Click += new System.EventHandler(this.удалитьПробелыToolStripMenuItem_Click);
            // 
            // получитьНазванияToolStripMenuItem
            // 
            this.получитьНазванияToolStripMenuItem.Name = "получитьНазванияToolStripMenuItem";
            this.получитьНазванияToolStripMenuItem.Size = new System.Drawing.Size(247, 22);
            this.получитьНазванияToolStripMenuItem.Text = "Получить названия";
            this.получитьНазванияToolStripMenuItem.Click += new System.EventHandler(this.получитьНазванияToolStripMenuItem_Click);
            // 
            // переназватьToolStripMenuItem1
            // 
            this.переназватьToolStripMenuItem1.Name = "переназватьToolStripMenuItem1";
            this.переназватьToolStripMenuItem1.Size = new System.Drawing.Size(247, 22);
            this.переназватьToolStripMenuItem1.Text = "Переназвать";
            this.переназватьToolStripMenuItem1.Click += new System.EventHandler(this.переназватьToolStripMenuItem1_Click);
            // 
            // открытьПоШаблонуToolStripMenuItem
            // 
            this.открытьПоШаблонуToolStripMenuItem.Name = "открытьПоШаблонуToolStripMenuItem";
            this.открытьПоШаблонуToolStripMenuItem.Size = new System.Drawing.Size(247, 22);
            this.открытьПоШаблонуToolStripMenuItem.Text = "Открыть по шаблону";
            this.открытьПоШаблонуToolStripMenuItem.Click += new System.EventHandler(this.открытьПоШаблонуToolStripMenuItem_Click);
            // 
            // экспортВOBJToolStripMenuItem
            // 
            this.экспортВOBJToolStripMenuItem.Name = "экспортВOBJToolStripMenuItem";
            this.экспортВOBJToolStripMenuItem.Size = new System.Drawing.Size(247, 22);
            this.экспортВOBJToolStripMenuItem.Text = "Экспорт в OBJ";
            this.экспортВOBJToolStripMenuItem.Click += new System.EventHandler(this.ЭкспортВOBJToolStripMenuItem_Click);
            // 
            // экспортВGLBОдноТелоToolStripMenuItem
            // 
            this.экспортВGLBОдноТелоToolStripMenuItem.Name = "экспортВGLBОдноТелоToolStripMenuItem";
            this.экспортВGLBОдноТелоToolStripMenuItem.Size = new System.Drawing.Size(247, 22);
            this.экспортВGLBОдноТелоToolStripMenuItem.Text = "Экспорт в GLB (одно тело)";
            this.экспортВGLBОдноТелоToolStripMenuItem.Click += new System.EventHandler(this.экспортВGLBОдноТелоToolStripMenuItem_Click);
            // 
            // экспортВGLBСборкаToolStripMenuItem
            // 
            this.экспортВGLBСборкаToolStripMenuItem.Name = "экспортВGLBСборкаToolStripMenuItem";
            this.экспортВGLBСборкаToolStripMenuItem.Size = new System.Drawing.Size(247, 22);
            this.экспортВGLBСборкаToolStripMenuItem.Text = "Экспорт в GLB (сборка)";
            this.экспортВGLBСборкаToolStripMenuItem.Click += new System.EventHandler(this.экспортВGLBСборкаToolStripMenuItem_Click);
            // 
            // восстановитьToolStripMenuItem1
            // 
            this.восстановитьToolStripMenuItem1.Name = "восстановитьToolStripMenuItem1";
            this.восстановитьToolStripMenuItem1.Size = new System.Drawing.Size(247, 22);
            this.восстановитьToolStripMenuItem1.Text = "Восстановить";
            this.восстановитьToolStripMenuItem1.Click += new System.EventHandler(this.восстановитьToolStripMenuItem1_Click);
            // 
            // оЦToolStripMenuItem
            // 
            this.оЦToolStripMenuItem.Name = "оЦToolStripMenuItem";
            this.оЦToolStripMenuItem.Size = new System.Drawing.Size(247, 22);
            this.оЦToolStripMenuItem.Text = "ОЦ";
            this.оЦToolStripMenuItem.Click += new System.EventHandler(this.оЦToolStripMenuItem_Click);
            // 
            // шипыИПазыToolStripMenuItem
            // 
            this.шипыИПазыToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.шипToolStripMenuItem1,
            this.шипИнверсияToolStripMenuItem,
            this.выдавитьToolStripMenuItem1,
            this.добавитьВСборкуToolStripMenuItem1,
            this.пазToolStripMenuItem1,
            this.пазМультидетальToolStripMenuItem});
            this.шипыИПазыToolStripMenuItem.Name = "шипыИПазыToolStripMenuItem";
            this.шипыИПазыToolStripMenuItem.Size = new System.Drawing.Size(93, 20);
            this.шипыИПазыToolStripMenuItem.Text = "Шипы и пазы";
            // 
            // шипToolStripMenuItem1
            // 
            this.шипToolStripMenuItem1.Name = "шипToolStripMenuItem1";
            this.шипToolStripMenuItem1.Size = new System.Drawing.Size(183, 22);
            this.шипToolStripMenuItem1.Text = "Шип";
            this.шипToolStripMenuItem1.Click += new System.EventHandler(this.шипToolStripMenuItem1_Click);
            // 
            // шипИнверсияToolStripMenuItem
            // 
            this.шипИнверсияToolStripMenuItem.Name = "шипИнверсияToolStripMenuItem";
            this.шипИнверсияToolStripMenuItem.Size = new System.Drawing.Size(183, 22);
            this.шипИнверсияToolStripMenuItem.Text = "Шип инверсия";
            this.шипИнверсияToolStripMenuItem.Click += new System.EventHandler(this.шипИнверсияToolStripMenuItem_Click);
            // 
            // выдавитьToolStripMenuItem1
            // 
            this.выдавитьToolStripMenuItem1.Name = "выдавитьToolStripMenuItem1";
            this.выдавитьToolStripMenuItem1.Size = new System.Drawing.Size(183, 22);
            this.выдавитьToolStripMenuItem1.Text = "Выдавить";
            this.выдавитьToolStripMenuItem1.Click += new System.EventHandler(this.выдавитьToolStripMenuItem1_Click);
            // 
            // добавитьВСборкуToolStripMenuItem1
            // 
            this.добавитьВСборкуToolStripMenuItem1.Name = "добавитьВСборкуToolStripMenuItem1";
            this.добавитьВСборкуToolStripMenuItem1.Size = new System.Drawing.Size(183, 22);
            this.добавитьВСборкуToolStripMenuItem1.Text = "Добавить в сборку";
            this.добавитьВСборкуToolStripMenuItem1.Click += new System.EventHandler(this.добавитьВСборкуToolStripMenuItem1_Click);
            // 
            // пазToolStripMenuItem1
            // 
            this.пазToolStripMenuItem1.Name = "пазToolStripMenuItem1";
            this.пазToolStripMenuItem1.Size = new System.Drawing.Size(183, 22);
            this.пазToolStripMenuItem1.Text = "Паз";
            this.пазToolStripMenuItem1.Click += new System.EventHandler(this.пазToolStripMenuItem1_Click);
            // 
            // пазМультидетальToolStripMenuItem
            // 
            this.пазМультидетальToolStripMenuItem.Name = "пазМультидетальToolStripMenuItem";
            this.пазМультидетальToolStripMenuItem.Size = new System.Drawing.Size(183, 22);
            this.пазМультидетальToolStripMenuItem.Text = "Паз (Мультидеталь)";
            this.пазМультидетальToolStripMenuItem.Click += new System.EventHandler(this.пазМультидетальToolStripMenuItem_Click);
            // 
            // создатьФайлыToolStripMenuItem
            // 
            this.создатьФайлыToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.удалитьЭлементToolStripMenuItem,
            this.общиеРазмерыToolStripMenuItem,
            this.размерыДоГибовToolStripMenuItem,
            this.позицииToolStripMenuItem,
            this.крепежToolStripMenuItem2,
            this.подпозицииToolStripMenuItem,
            this.открепленныеРазмерыToolStripMenuItem,
            this.проекцияToolStripMenuItem,
            this.выноскиГибовToolStripMenuItem});
            this.создатьФайлыToolStripMenuItem.Name = "создатьФайлыToolStripMenuItem";
            this.создатьФайлыToolStripMenuItem.Size = new System.Drawing.Size(63, 20);
            this.создатьФайлыToolStripMenuItem.Text = "Удалить";
            this.создатьФайлыToolStripMenuItem.Click += new System.EventHandler(this.создатьФайлыToolStripMenuItem_Click);
            // 
            // удалитьЭлементToolStripMenuItem
            // 
            this.удалитьЭлементToolStripMenuItem.Name = "удалитьЭлементToolStripMenuItem";
            this.удалитьЭлементToolStripMenuItem.Size = new System.Drawing.Size(208, 22);
            this.удалитьЭлементToolStripMenuItem.Text = "Удалить элемент";
            this.удалитьЭлементToolStripMenuItem.Click += new System.EventHandler(this.удалитьЭлементToolStripMenuItem_Click_1);
            // 
            // общиеРазмерыToolStripMenuItem
            // 
            this.общиеРазмерыToolStripMenuItem.Name = "общиеРазмерыToolStripMenuItem";
            this.общиеРазмерыToolStripMenuItem.Size = new System.Drawing.Size(208, 22);
            this.общиеРазмерыToolStripMenuItem.Text = "Общие размеры";
            this.общиеРазмерыToolStripMenuItem.Click += new System.EventHandler(this.общиеРазмерыToolStripMenuItem_Click);
            // 
            // размерыДоГибовToolStripMenuItem
            // 
            this.размерыДоГибовToolStripMenuItem.Name = "размерыДоГибовToolStripMenuItem";
            this.размерыДоГибовToolStripMenuItem.Size = new System.Drawing.Size(208, 22);
            this.размерыДоГибовToolStripMenuItem.Text = "Размеры до гибов";
            this.размерыДоГибовToolStripMenuItem.Click += new System.EventHandler(this.размерыДоГибовToolStripMenuItem_Click);
            // 
            // позицииToolStripMenuItem
            // 
            this.позицииToolStripMenuItem.Name = "позицииToolStripMenuItem";
            this.позицииToolStripMenuItem.Size = new System.Drawing.Size(208, 22);
            this.позицииToolStripMenuItem.Text = "Позиции";
            this.позицииToolStripMenuItem.Click += new System.EventHandler(this.позицииToolStripMenuItem_Click);
            // 
            // крепежToolStripMenuItem2
            // 
            this.крепежToolStripMenuItem2.Name = "крепежToolStripMenuItem2";
            this.крепежToolStripMenuItem2.Size = new System.Drawing.Size(208, 22);
            this.крепежToolStripMenuItem2.Text = "Крепеж (X)";
            this.крепежToolStripMenuItem2.Click += new System.EventHandler(this.крепежToolStripMenuItem2_Click);
            // 
            // подпозицииToolStripMenuItem
            // 
            this.подпозицииToolStripMenuItem.Name = "подпозицииToolStripMenuItem";
            this.подпозицииToolStripMenuItem.Size = new System.Drawing.Size(208, 22);
            this.подпозицииToolStripMenuItem.Text = "Подпозиции";
            this.подпозицииToolStripMenuItem.Click += new System.EventHandler(this.подпозицииToolStripMenuItem_Click);
            // 
            // открепленныеРазмерыToolStripMenuItem
            // 
            this.открепленныеРазмерыToolStripMenuItem.Name = "открепленныеРазмерыToolStripMenuItem";
            this.открепленныеРазмерыToolStripMenuItem.Size = new System.Drawing.Size(208, 22);
            this.открепленныеРазмерыToolStripMenuItem.Text = "Открепленные размеры";
            this.открепленныеРазмерыToolStripMenuItem.Click += new System.EventHandler(this.открепленныеРазмерыToolStripMenuItem_Click);
            // 
            // проекцияToolStripMenuItem
            // 
            this.проекцияToolStripMenuItem.Name = "проекцияToolStripMenuItem";
            this.проекцияToolStripMenuItem.Size = new System.Drawing.Size(208, 22);
            this.проекцияToolStripMenuItem.Text = "Проекция";
            this.проекцияToolStripMenuItem.Click += new System.EventHandler(this.проекцияToolStripMenuItem_Click);
            // 
            // выноскиГибовToolStripMenuItem
            // 
            this.выноскиГибовToolStripMenuItem.Name = "выноскиГибовToolStripMenuItem";
            this.выноскиГибовToolStripMenuItem.Size = new System.Drawing.Size(208, 22);
            this.выноскиГибовToolStripMenuItem.Text = "Выноски гибов";
            this.выноскиГибовToolStripMenuItem.Click += new System.EventHandler(this.выноскиГибовToolStripMenuItem_Click);
            // 
            // извToolStripMenuItem
            // 
            this.извToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.заполнитьToolStripMenuItem,
            this.добавитьToolStripMenuItem,
            this.сформироватьКомплектToolStripMenuItem});
            this.извToolStripMenuItem.Name = "извToolStripMenuItem";
            this.извToolStripMenuItem.Size = new System.Drawing.Size(39, 20);
            this.извToolStripMenuItem.Text = "Изв";
            // 
            // заполнитьToolStripMenuItem
            // 
            this.заполнитьToolStripMenuItem.Name = "заполнитьToolStripMenuItem";
            this.заполнитьToolStripMenuItem.Size = new System.Drawing.Size(214, 22);
            this.заполнитьToolStripMenuItem.Text = "Заполнить";
            this.заполнитьToolStripMenuItem.Click += new System.EventHandler(this.заполнитьToolStripMenuItem_Click);
            // 
            // добавитьToolStripMenuItem
            // 
            this.добавитьToolStripMenuItem.Name = "добавитьToolStripMenuItem";
            this.добавитьToolStripMenuItem.Size = new System.Drawing.Size(214, 22);
            this.добавитьToolStripMenuItem.Text = "Добавить";
            this.добавитьToolStripMenuItem.Click += new System.EventHandler(this.добавитьToolStripMenuItem_Click);
            // 
            // сформироватьКомплектToolStripMenuItem
            // 
            this.сформироватьКомплектToolStripMenuItem.Name = "сформироватьКомплектToolStripMenuItem";
            this.сформироватьКомплектToolStripMenuItem.Size = new System.Drawing.Size(214, 22);
            this.сформироватьКомплектToolStripMenuItem.Text = "Сформировать комплект";
            this.сформироватьКомплектToolStripMenuItem.Click += new System.EventHandler(this.сформироватьКомплектToolStripMenuItem_Click);
            // 
            // перенаправитьToolStripMenuItem
            // 
            this.перенаправитьToolStripMenuItem.Name = "перенаправитьToolStripMenuItem";
            this.перенаправитьToolStripMenuItem.Size = new System.Drawing.Size(104, 20);
            this.перенаправитьToolStripMenuItem.Text = "Перенаправить";
            this.перенаправитьToolStripMenuItem.Click += new System.EventHandler(this.перенаправитьToolStripMenuItem_Click);
            // 
            // крепежToolStripMenuItem
            // 
            this.крепежToolStripMenuItem.Name = "крепежToolStripMenuItem";
            this.крепежToolStripMenuItem.Size = new System.Drawing.Size(78, 20);
            this.крепежToolStripMenuItem.Text = "Крепеж (F)";
            this.крепежToolStripMenuItem.Click += new System.EventHandler(this.крепежToolStripMenuItem_Click);
            // 
            // сайтToolStripMenuItem
            // 
            this.сайтToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.названиеМоделиToolStripMenuItem,
            this.спецификацияToolStripMenuItem,
            this.создатьСписокФайловToolStripMenuItem,
            this.моделиИзСпискаToolStripMenuItem,
            this.весьПроектToolStripMenuItem,
            this.gltfToolStripMenuItem,
            this.dИзСпискаФайловToolStripMenuItem,
            this.оптБрToolStripMenuItem});
            this.сайтToolStripMenuItem.Name = "сайтToolStripMenuItem";
            this.сайтToolStripMenuItem.Size = new System.Drawing.Size(45, 20);
            this.сайтToolStripMenuItem.Text = "Сайт";
            // 
            // названиеМоделиToolStripMenuItem
            // 
            this.названиеМоделиToolStripMenuItem.Name = "названиеМоделиToolStripMenuItem";
            this.названиеМоделиToolStripMenuItem.Size = new System.Drawing.Size(241, 22);
            this.названиеМоделиToolStripMenuItem.Text = "Название модели";
            this.названиеМоделиToolStripMenuItem.Click += new System.EventHandler(this.названиеМоделиToolStripMenuItem_Click);
            // 
            // спецификацияToolStripMenuItem
            // 
            this.спецификацияToolStripMenuItem.Name = "спецификацияToolStripMenuItem";
            this.спецификацияToolStripMenuItem.Size = new System.Drawing.Size(241, 22);
            this.спецификацияToolStripMenuItem.Text = "Спецификация";
            this.спецификацияToolStripMenuItem.Click += new System.EventHandler(this.спецификацияToolStripMenuItem_Click);
            // 
            // создатьСписокФайловToolStripMenuItem
            // 
            this.создатьСписокФайловToolStripMenuItem.Name = "создатьСписокФайловToolStripMenuItem";
            this.создатьСписокФайловToolStripMenuItem.Size = new System.Drawing.Size(241, 22);
            this.создатьСписокФайловToolStripMenuItem.Text = "Создать список файлов";
            this.создатьСписокФайловToolStripMenuItem.Click += new System.EventHandler(this.создатьСписокФайловToolStripMenuItem_Click);
            // 
            // моделиИзСпискаToolStripMenuItem
            // 
            this.моделиИзСпискаToolStripMenuItem.Name = "моделиИзСпискаToolStripMenuItem";
            this.моделиИзСпискаToolStripMenuItem.Size = new System.Drawing.Size(241, 22);
            this.моделиИзСпискаToolStripMenuItem.Text = "модели из списка";
            this.моделиИзСпискаToolStripMenuItem.Click += new System.EventHandler(this.моделиИзСпискаToolStripMenuItem_Click);
            // 
            // весьПроектToolStripMenuItem
            // 
            this.весьПроектToolStripMenuItem.Name = "весьПроектToolStripMenuItem";
            this.весьПроектToolStripMenuItem.Size = new System.Drawing.Size(241, 22);
            this.весьПроектToolStripMenuItem.Text = "Из списка файлов";
            this.весьПроектToolStripMenuItem.Click += new System.EventHandler(this.весьПроектToolStripMenuItem_Click);
            // 
            // gltfToolStripMenuItem
            // 
            this.gltfToolStripMenuItem.Name = "gltfToolStripMenuItem";
            this.gltfToolStripMenuItem.Size = new System.Drawing.Size(241, 22);
            this.gltfToolStripMenuItem.Text = "Маршрутки из списка файлов";
            this.gltfToolStripMenuItem.Click += new System.EventHandler(this.gltfToolStripMenuItem_Click);
            // 
            // dИзСпискаФайловToolStripMenuItem
            // 
            this.dИзСпискаФайловToolStripMenuItem.Name = "dИзСпискаФайловToolStripMenuItem";
            this.dИзСпискаФайловToolStripMenuItem.Size = new System.Drawing.Size(241, 22);
            this.dИзСпискаФайловToolStripMenuItem.Text = "3d из списка файлов";
            this.dИзСпискаФайловToolStripMenuItem.Click += new System.EventHandler(this.dИзСпискаФайловToolStripMenuItem_Click);
            // 
            // оптБрToolStripMenuItem
            // 
            this.оптБрToolStripMenuItem.Name = "оптБрToolStripMenuItem";
            this.оптБрToolStripMenuItem.Size = new System.Drawing.Size(241, 22);
            this.оптБрToolStripMenuItem.Text = "Опт Бр";
            this.оптБрToolStripMenuItem.Click += new System.EventHandler(this.оптБрToolStripMenuItem_Click);
            // 
            // изменитьОтверстияToolStripMenuItem
            // 
            this.изменитьОтверстияToolStripMenuItem.Name = "изменитьОтверстияToolStripMenuItem";
            this.изменитьОтверстияToolStripMenuItem.Size = new System.Drawing.Size(247, 22);
            this.изменитьОтверстияToolStripMenuItem.Text = "Изменить отверстия";
            this.изменитьОтверстияToolStripMenuItem.Click += new System.EventHandler(this.изменитьОтверстияToolStripMenuItem_Click);
            // 
            // CreateComponent
            // 
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Inherit;
            this.ClientSize = new System.Drawing.Size(827, 27);
            this.Controls.Add(this.menuStrip1);
            this.Location = new System.Drawing.Point(300, 130);
            this.MainMenuStrip = this.menuStrip1;
            this.Name = "CreateComponent";
            this.StartPosition = System.Windows.Forms.FormStartPosition.Manual;
            this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.CreateComponent_KeyDown);
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        public CreateComponent(Inventor.Document doc)
        {
            m_Parts = new Parts(Macros.StandardAddInServer.m_inventorApplication.ActiveDocument);
            desc = new XMLDoc(I.p() + @"\Description.xml", "Description");
            InitializeComponent();
            InterfaceDll.Data.EventHandler = new Data.MyEvent(func);
        }

        public void show()
        {
            this.ShowDialog();
        }

        protected virtual void OnMyEvent(Macros.MyEvArgs e)
        {
            Macros.MyHandler handler = myEvent;
            if (handler != null)
            {
                handler(this, e);
            }
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
            List<KeyValuePair<string, string>> replace = new List<KeyValuePair<string, string>>();
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
                        if (dgv[colDesc.Index, i].Value != null) elem.SetAttributeValue("D", dgv[colDesc.Index, i].Value.ToString());
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
                            if (dgv[colReplace.Index, i + k].Value == null) rep = "";
                            else rep = dgv[colReplace.Index, i + k].Value.ToString();

                            if (dgv[colLib.Index, i + k].Value == null) lib = "";
                            else lib = dgv[colLib.Index, i + k].Value.ToString();

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
                            || asmName == dgv[0, i + k].Value.ToString()));
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

        public string replaceUnion(List<KeyValuePair<string, string>> lst)
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
            Dictionary<string, string> dic = new Dictionary<string, string>() { { "name", "Название файла" }, { "description", "наименование" }, { "decNumber", "децимальный номер" }, { "base", "название файла для наследования" }, { "files", "Файлы для замены" }, { "replace", "Замена" }, { "count", "Кол-во" } };
            float[] weigth = { 0.2f, 0.2f, 0.1f, 0.2f, 0.1f, 0.2f, 0.1f };
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
                desc = desc ?? new XMLDoc(I.p() + @"\Description.xml", "Description");
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
            XMLDoc xmldoc = new XMLDoc(I.p() + @"\Description.xml", "Description");
            addItems(xmldoc.Doc, col, "Type", "name");
            tb.AutoCompleteCustomSource = col;
        }

        private void autoComplete(System.Windows.Forms.ComboBox cb)
        {
            XMLDoc xmldoc = new XMLDoc(I.p() + @"\Description.xml", "Description");
            HashSet<string> hs = new HashSet<string>();
            foreach (var item in xmldoc.El.Descendants("Value"))
            {
                if (item.Attribute("type") != null) hs.Add(item.Attribute("type").Value);
            }
            cb.Items.AddRange(hs.ToArray());
        }

        public void addItems(XDocument doc, AutoCompleteStringCollection col, string name, string attName)
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
                desc = desc ?? new XMLDoc(I.p() + @"\Description.xml", "Description");
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
                            addItems(desc.Doc, "name", "Value", DataCollection);
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
                            asmDoc = (AssemblyDocument)invApp.Documents.Open(doc.path() + "\\" + asmname, false);
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
                                        XMLDoc lib = new XMLDoc(I.p() + @"\ContentCenter.xml", "Content");
                                        addItems(lib.Doc, DataCollection);
                                    }
                                }
                            }
                            XMLDoc lib1 = new XMLDoc(I.p() + @"\ContentCenter.xml", "Content");
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

        public string fullName(AssemblyDocument doc, string name, ref bool isContent)
        {
            name = name.Substring(0, name.Length - 4);
            AssemblyComponentDefinition compDef = doc.ComponentDefinition;
            foreach (ComponentOccurrence occ in compDef.Occurrences)
            {
                if (occ.Name.ToLower().IndexOf(name.ToLower()) != -1)
                {
                    if (occ.DefinitionDocumentType == DocumentTypeEnum.kPartDocumentObject)
                        isContent = (((PartComponentDefinition)occ.Definition).IsContentMember) ? true : false;
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
                if (filter != "" && item.Attribute("type") != null && item.Attribute("type").Value != filter) continue;
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
                    name = item.Substring(item.LastIndexOf('\\') + 1, item.Length - 1 - item.LastIndexOf('\\'));
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

        private void удалитьЭлементToolStripMenuItem_Click_1(object sender, EventArgs e)
        {
            Document doc = I.aDoc();
            if (doc.DocumentType == DocumentTypeEnum.kAssemblyDocumentObject)
            {
                SelectSet ss = I.getSS(doc);
                foreach (ComponentOccurrence occ in ss)
                {
                    AssemblyComponentDefinition acd = ((AssemblyDocument)doc).ComponentDefinition;
                    Parts.removeOcc(acd, occ);
                }
            }
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
                if (pDoc.ComponentDefinition.ReferenceComponents.DerivedPartComponents.Count != 0)
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
        public static bool addLinkParam(Document to, Document from, string[] param)
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

        private void spikes(bool rev = true)
        {
            var ss = I.getSS();
            Vertex v = null;
            if (ss.Count == 1)
                v = ss[1] as Vertex;
            double R = 3.0 / 10, L = 5.0 / 10, H = 7.5 / 10;
            spike sp = new spike(Macros.StandardAddInServer.m_inventorApplication);
            sp.smcd = (SheetMetalComponentDefinition)((PartDocument)sp.invApp.ActiveDocument).ComponentDefinition;
            sp.R = R; sp.L = L; sp.H = H;
            sp.addSpikeDef(sp.smcd, R, R * 2, H);
            sp.addSpikeDef(sp.smcd, R, L, H, "Капелька");
            sp.addSketch(sp.smcd, H, R, L, rev, v);
        }

        private void шипToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        public void addCuts(AssemblyDocument aDoc, AssemblyComponentDefinition acd, string[] nameIn, string nameOut, String nameSketch, string nameCut)
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
            string path = invDoc.path + invDoc.getType + "(" + System.DateTime.Now.ToString("dd.MM.yyyy") + ").xml";
            XMLDoc xmlDoc = new XMLDoc(path, "Path");
            xmlDoc.Doc.Root.Add(new XElement("Name", invDoc.getType + "(" + System.DateTime.Now.ToString("dd.MM.yyyy") + ")"));
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

        private void drawing()
        {
            MyXML xml = new MyXML("AddDrawingInterface.xml");
            MyForm f = new MyForm(xml, "Чертеж"/*, null,*/ );
            DialogResult dr = f.f.ShowDialog();
            if (dr != DialogResult.OK) return;
            var values = getValues(f);
            bool open = f.chks[1].Checked, dims = f.chks[0].Checked, notSave = f.chks[2].Checked;
            if (f.cbs[0].Text == "Шаблон")
                foreach (Document item in I.visDocs())
                {
                    if (item.DocumentType == DocumentTypeEnum.kDrawingDocumentObject) continue;
                    drawing(item, values, open, notSave, dims);
                }
            else
                drawing(I.app.ActiveDocument, values, open, notSave, dims);
        }
        public void assembly()
        {
            var doc = I.aDoc();
            var ss = doc.SelectSet;
            string path = file.p(doc.FullDocumentName);
            string fs = u.OFD(path, multi: true);
            List<string> names = new List<string>();
            MyXML xml = new MyXML("AddAssemblyInterface.xml");
            MyForm f = new MyForm(xml, "Сборка");
            string type = u.getPropValue((Document)doc, "Type"),
                    dn = u.getPropValue((Document)doc, "DecNumber");
            if (ss.Count == 1)
            {
                var body = GetBodies(ss).ToList()[0];
                var spl1 = body.Name.Split('$');
                
                string num = u.getDN(spl1, dn, 0);
                f.cbs[0].Text = $"{type}.{dn}";
                var p = file.p(doc.FullDocumentName);
                var fi = $"{type}.{dn.Substring(0, dn.Length - 1)}";
                foreach (var item in file.getFiles(p, ".ipt"))
                {
                    var n = file.nameWithExt(item);
                    if (n.IndexOf(fi) != -1)
                        names.Add(n);
                }
            } else
            {
                //dn[dn.Length - 1] = '0';
                if (!addName(fs, f))
                    f.cbs[0].Text = $"{type}.{dn}";
            }
            DialogResult dr = f.f.ShowDialog();
            if (dr != DialogResult.OK) return;
            var vals = getValues(f, true);
            string fn = $"{path}{vals[0]} ({vals[1]}).iam";
            InvDocument<Document> id = new InvDocument<Document>(doc);
            var asm = id.addAsmDoc();
            asm.SaveAs(fn, false);
            var def = asm.ComponentDefinition;
            var mtx = I.getMatrix();
            foreach (var item in fs.Split('|'))
            {
                def.Occurrences.Add(item, mtx);
            }
            u.addProp((Document)asm, "Description", vals[1]);
            var spl = vals[0].Split('.');
            if (spl.Length > 1 )
            {
                var t = spl[0];
                var tmp = spl.Skip(1).ToArray();
                dn = String.Join(".", tmp);
                u.addProp((Document)asm, "Type", t);
                u.addProp((Document)asm, "DecNumber", dn);
                u.addProp((Document)asm, "Part Number", "=<Type>.<DecNumber>");

            } else
                u.addProp((Document)asm, "Part Number", vals[0]);
            asm.Save();
            I.open(asm.FullDocumentName, true);
        }
        public bool addName(string fs, MyForm f)
        {
            MyXML xml = new MyXML("AssemblyNames.xml");
            if (xml.elem == null) return false;
            foreach (var item in fs.Split('|'))
            {
                var name = file.name(item);
                foreach (XElement el in xml.elem.Elements())
                {
                    string find = MyXML.getAtt(el, "f");
                    var spl = find.Split(';');
                    if (inList(name, spl))
                    {
                        var desc = MyXML.getAtt(el, "name");
                        var dn = MyXML.getAtt(el, "dn");
                        var doc = I.open(item);
                        var t = u.getPropValue(doc, "Type");
                        f.cbs[0].Text = $"{t}.{dn}";
                        f.cbs[1].Text = desc;
                        return true;
                    }
                }
            }
            return false;
        }
        public bool inList(string n, string[] f)
        {
            foreach (var item in f)
            {
                if (n.IndexOf(item) != -1) return true;
            }
            return false;
        }
        private List<string> getValues(MyForm f, bool cbs = false)
        {
            List<string> values = new List<string>();
            for (int i = 0; i < f.cbs.Count(); i++)
            {
                values.Add(f.cbs[i].Text);
            }
            if (cbs) return values;
            values.Add(f.chks[3].Checked.ToString()); values.Add(f.chks[4].Checked.ToString());
            return values;
        }

        private void текущийДокументToolStripMenuItem_Click(object sender, EventArgs e)
        {
            drawing();
        }
        public void addEl(XElement par, string nameAtt, string attVal, string align)
        {
            par.Add(new XElement("view1", new XAttribute(nameAtt, attVal), new XAttribute("align", align)));
        }
        public DrawingDocument drawing(Document doc, List<string> values, bool open = false, bool notSave = false, bool dims = false)
        {
            InvDoc.InvDocument<Document> invDoc = new InvDoc.InvDocument<Document>(doc);
            XMLDoc xmlDoc = new XMLDoc(I.p() + @"\sheet.xml", "head");
            List<double> scales = new List<double>();
            string scale = null;
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
                if (notSave)
                {
                    if (file.check(name))
                    {
                        return open ? invDoc.openDrwDoc(name, true) : invDoc.openDrwDoc(name);
                    }
                }
                drw.SaveAs(name, false);
                if (doc.DocumentType == DocumentTypeEnum.kAssemblyDocumentObject || doc.DocumentType == DocumentTypeEnum.kPartDocumentObject)
                {
                    //    AssemblyDocument asmDoc = (AssemblyDocument)doc;
                    //}
                    //else if (doc.DocumentType == DocumentTypeEnum.kPartDocumentObject)
                    //{
                    //    PartDocument pDoc = (PartDocument)doc;
                    DesignViewRepresentation dvr = null;
                    Camera cam = null;
                    SheetMetalComponentDefinition smcd = invDoc.getSheetMetalCompDef;
                    string val = desc.ToLower();
                    invDoc.doc = (Document)drw;
                    //var ie = xmlDoc.Doc.Root.Descendants("sheet");
                    //Regex reg = new Regex(val);
                    XElement el = xmlDoc.El.Elements("view").FirstOrDefault(elem => elem.Attribute("name").Value.ToLower() == values[0].ToLower());
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
                        if (values[0] == "Шаблон")
                        {
                            dvr = InvDocument<PartDocument>.getDVR(doc, "drw");
                            if (dvr != null)
                            {
                                //if (!dvr.ModelAnnotationAutoScale)
                                //{
                                //    scale = dvr.ModelAnnotationScale.ToString();
                                 
                                //}
                                cam = dvr.Camera;
                                if (dvr.DesignViewType == DesignViewTypeEnum.kMasterDesignViewType)
                                {
                                    cam = null;
                                    values[0] = "Текущий";
                                }
                            }
                            else
                            {
                                values[0] = "Текущий";
                                el = xmlDoc.El.Elements("view").FirstOrDefault(elem => elem.Attribute("name").Value.ToLower() == values[0].ToLower());
                            }
                        }
                        if (scale != null) dv = addDV(scale, drw.Sheets[1], doc, el, "view1", scales);
                        else dv = addDV(values[1], drw.Sheets[1], doc, el, "view1", scales, cam: cam);

                        if (drw.PropertySets[6][8].Value.ToString() == "") drw.PropertySets[6][8].Value = dv.ScaleString;
                        XMLDoc xdoc = null; string tmp = null;
                        if (el.Attribute("TR") != null)
                        {
                            xdoc = new XMLDoc(I.p() + @"\" + el.Attribute("TR").Value, "TR");
                            addTR(drw.Sheets[1], xdoc);
                        }
                        else
                        {
                            if (doc.DocumentType == DocumentTypeEnum.kPartDocumentObject)
                            {
                                var prop = u.getPropValue(doc, "dxf");
                                if (!u.isNull(prop))
                                {
                                    tmp = xmlDoc.El.Element("TR").Elements().ElementAt(2).Attribute("value").Value;
                                }
                                else
                                    tmp = xmlDoc.El.Element("TR").Elements().ElementAt(0).Attribute("value").Value;
                            }
                            else if (doc.DocumentType == DocumentTypeEnum.kAssemblyDocumentObject)
                                tmp = xmlDoc.El.Element("TR").Elements().ElementAt(1).Attribute("value").Value;
                            xdoc = new XMLDoc(I.p() + @"\" + tmp, "TR");
                            addTR(drw.Sheets[1], xdoc);
                        }
                    }
                    var sh = invDoc.changeSheet(drw.Sheets[1], w, h);
                    Vector2d vec = I.tg.CreateVector2d(0, dv.Height / 2 + offset);
                    if (smcd != null)
                    {
                        if (!smcd.HasFlatPattern) smcd = null;
                        else if (smcd.Bends.Count == 0) smcd = null;
                        if (smcd == null)
                        {
                            //drw.SelectSet.Select(dv);
                            InvDocument<PartDocument>.addSketchedSymbol(sh, "АР", new string[] { "" }, dv.Position, vec);
                        }
                    }
                    if (smcd != null)
                    {
                        bool addSheet = true;
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
                                    addSheet = false;
                                    //w = u.convToDouble(values[2]);
                                    break;
                            }
                        }
                        if (addSheet)
                            sh = invDoc.addSheet(w, h);
                        if (el != null)
                        {
                            dv = addDV(values[1], sh, doc, el, "view2", scales);
                            dv.DisplayTangentEdges = true;
                            dv.DisplayForeshortenedTangentEdges = true;
                            if (el.Attribute("Dim") != null)
                            {
                                Dimensions dimens = new Dimensions(dv);
                                //Tables.addBends();
                            }
                            //drw.SelectSet.Select(dv);
                            DrawingViewLabel dvl = dv.Label;
                            var dv1 = sh.DrawingViews[1];
                            if (dv.Scale == dv1.Scale) 
                                dvl.FormattedText = @"<StyleOverride Font='AIGDT' Italic='false'>/</StyleOverride> Раскрой методом АР";
                            else
                                dvl.FormattedText = @"<StyleOverride Font='AIGDT' Italic='false'>/</StyleOverride> (<DrawingViewScale/>) Раскрой методом АР";
                            //dvl.Position = I.CP2d(dvl.Position.X, dvl.Position.Y - 2); 
                            dv.ShowLabel = true;
                            //InvDocument<PartDocument>.addSketchedSymbol(drw.Sheets[2], "Развертка 1:N АР", new string[] { (1/dv.Scale).ToString("#.#")},
                            //    dv.Position, vec);
                            if (dims) addDims(dv);
                        }
                        //drw.Sheets[1].Activate();
                    }
                }
                drw.Save2();
                if (notSave) return drw;
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
            return drw;
        }

        private void addDims(DrawingView dv)
        {
            gabs.add(dv); gabs.set();
            if (dv != null)
            {
                new CheckReflect(dv, I.CV2d(1)); new CheckReflect(dv, I.CV2d(0, 1));
                gabs.center();
                Dimensions dimens = new Dimensions(dv);
            }
        }

        private DrawingView addDV(string val, Sheet sh, Document doc, XElement el, string name, List<double> scales, Camera cam = null)
        {
            if (val == "0")
                return InvDocument<PartDocument>.addView(sh, doc, el, name, 0, scales, cam: cam);
            else
                return InvDocument<PartDocument>.addView(sh, doc, el, name, u.convToDouble(val), scales, cam: cam);
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
                    if (doc.DocumentType == DocumentTypeEnum.kPartDocumentObject &&
                    ((PartDocument)doc).ComponentDefinition.SurfaceBodies.Count != 1)
                        continue;

                    if (doc.DocumentType != DocumentTypeEnum.kDrawingDocumentObject)
                    {
                        pn = doc.PropertySets[3][2].Value.ToString();
                        if (pn == "") continue;
                        desc = doc.PropertySets[3][14].Value.ToString();
                        if (desc == "") continue;
                        ext = System.IO.Path.GetExtension(f);
                        if (doc.RequiresUpdate) doc.Update2(false);
                        //doc.Close();
                    }
                    else
                    {
                        if (doc.ReferencedDocuments.Count == 0)
                        {
                            //doc.Close(); 
                            continue;
                        }
                        pn = InvDoc.u.referendedDoc(doc).PropertySets[3][2].Value.ToString();
                        desc = InvDoc.u.referendedDoc(doc).PropertySets[3][14].Value.ToString();
                        ext = System.IO.Path.GetExtension(f);

                        //doc.Close();
                    }
                    string p = "";
                    int indx = pn.IndexOf("-", pn.Length - 4);
                    if (indx != -1 && ext == ".iam")
                    {
                        p = "^" + pn.Substring(indx + 1, 2);
                    }
                    name = pn + " (" + desc + ")" + p + ext;
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
                    MessageBox.Show("Файл " + name + " существует");
                }
                pb.UpdateProgress();
            }
            pb.Close();
            xmlDoc.save();

            app.SilentOperation = false;
        }

        public static XMLDoc xml(XElement el)
        {
            XMLDoc xdoc = new XMLDoc(I.p() + @"\rename.xml", "head");
            foreach (var item in el.Descendants())
            {
                if (!item.HasAttributes) continue;
                if (item.Attribute("old").Value == item.Attribute("new").Value) continue;
                XElement fel = xdoc.find(item.FirstAttribute.Name.ToString(), item.FirstAttribute.Value.ToString());
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
            XMLDoc xdoc = new XMLDoc(I.p() + @"\rename1.xml", "head");
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
                if (flag) { flag = false; n = System.IO.Path.GetFileNameWithoutExtension(el.Attribute("new").Value) + ".idw"; }
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
                if (rename(doc, el))
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
            if (pn == "") { el = null; return false; }
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

        private void copyAttrToolStripMenuItem1_Click_1(object sender, EventArgs e)
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

        private void addTRToolStripMenuItem1_Click_1(object sender, EventArgs e)
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
            MyLabel mylbl1 = new MyLabel(); int offsetY = 20;
            MyTextBox mytxt1 = new MyTextBox();
            MyComboBox mycb = new MyComboBox();
            f.Height = 100; f.Width = 400; f.WindowState = FormWindowState.Normal; f.Text = "Добавить свойство"; f.StartPosition = FormStartPosition.CenterScreen;
            System.Drawing.Point insPt = new System.Drawing.Point(5, 5);
            Label lbl1 = mylbl1.addLabel("Имя свойства", insPt, 100, 15);
            f.Controls.Add(lbl1);
            insPt.Y += offsetY;
            lbl1 = mylbl1.addLabel("Значение", insPt, 100, 15);
            f.Controls.Add(lbl1);
            insPt.X += 150; insPt.Y -= offsetY;
            System.Windows.Forms.ComboBox cb = mycb.addComboBox("", insPt, 200, 15, new string[] { "Изв", "ИзвД", "CountDXF", "thick", "NoFastener", "Break", "dxf", "model" });
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

        public static void pack(System.Collections.IEnumerable docs)
        {
            bool first = true;
            string locName = null;
            string name = null;
            foreach (Document doc in docs)
            {
                if (first)
                {
                    locName = packFolder(doc);
                    name = doc.FullDocumentName;
                    first = false;
                }
                if (locName == null) return;
                pack(doc, locName, name);
            }
            first = true;
            foreach (var item in System.IO.Directory.EnumerateFiles(locName, "*.ipj*"))
            {
                if (first)
                {
                    first = false;
                }
                else
                {
                    System.IO.File.Delete(item);
                }
            }
        }

        public static void pack(Document doc, string locName, string name)
        {
            PackAndGoLib.PackAndGoComponent packAndGoComp = new PackAndGoLib.PackAndGoComponent();
            //Document doc = Macros.StandardAddInServer.m_inventorApplication.ActiveDocument;
            PackAndGoLib.PackAndGo packAndGo = packAndGoComp.CreatePackAndGo(name, locName);

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

        public static string packFolder(Document doc)
        {
            string locName = doc.path() + "\\";
            locName += doc.name() + "\\";
            if (!System.IO.Directory.Exists(locName)) System.IO.Directory.CreateDirectory(locName);
            return locName;
        }

        private void комплектФайловToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            Document doc = Macros.StandardAddInServer.m_inventorApplication.ActiveDocument;
            string locName = packFolder(doc);
            pack(doc, locName, doc.FullDocumentName);
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

        public void pdfFromXML()
        {
            string path = I.curProjPath() + "\\";
            string dir = file.combine(path, "step");
            var xdoc = u.getXMLOFD();
            foreach (var item in xdoc.El.Elements())
            {
                string ffn = XMLDoc.getAttributeValue(item, "ffn");
                if (u.isNull(ffn)) continue;
                var doc = I.open(ffn, false, false);
                pdf(doc, dir);
            }
        }

        private void stepToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Exp3DPDF();
            //List<string> names = new List<string>();
            //string path = I.curProjPath() + "\\";
            //string dir = file.combine(path, "PDF3D");
            //file.createPath(dir);
            //ApplicationAddIn PDFAddin = null;
            //foreach (ApplicationAddIn item in I.app.ApplicationAddIns)
            //{
            //    if (item.ClassIdString == "{3EE52B28-D6E0-4EA4-8AA6-C2A266DEBB88}")
            //    {
            //        PDFAddin = item; break;
            //    }
            //}
            //dynamic pdfConv = PDFAddin.Automation;
            //NameValueMap oOptions = I.nvm;

            //foreach (Document item in I.app.Documents.VisibleDocuments)
            //{
            //    if (item.DocumentType == DocumentTypeEnum.kAssemblyDocumentObject)
            //    {
            //        item.Activate();
            //        oOptions.Value["FileOutputLocation"] = dir + file.name(item.FullDocumentName) + ".pdf";
            //        oOptions.Value["ExportAllProperties"] = true;
            //        oOptions.Value["GenerateAndAttachSTEPFile"] = true;
            //        oOptions.Value["ExportTemplate"] = @"c:\Users\Public\Documents\Autodesk\Inventor 2021\Templates\ru-RU\Blank.pdf";
            //        oOptions.Value["VisualizationQuality"] = AccuracyEnum.kHigh;

            //        String[] attFiles = new string[] { item.FullDocumentName };
            //        oOptions.Value["AttachedFiles "] = attFiles;

            //        pdfConv.Publish(item, oOptions);

            //        //names.Add(pdf(item, dir)); 
            //    }
            //}    
        }

        private string pdf(Document doc, string sdir)
        {
            TranslatorAddIn step = (TranslatorAddIn)Macros.StandardAddInServer.m_inventorApplication.ApplicationAddIns.ItemById["{90AF7F40-0C01-11D5-8E83-0010B541CD80}"];
            TranslationContext context = I.objs.CreateTranslationContext();
            NameValueMap nvm = I.objs.CreateNameValueMap();
            nvm.Add("ApplicationProtocolType", 3);
            nvm.Add("IncludeSketches", false);
            context.Type = IOMechanismEnum.kFileBrowseIOMechanism;
            DataMedium dm = I.objs.CreateDataMedium();
            if (doc.DocumentType == DocumentTypeEnum.kAssemblyDocumentObject)
            {
                AssemblyDocument asm = doc as AssemblyDocument;
                asm.ComponentDefinition.RepresentationsManager.DesignViewRepresentations[1].Activate();
                asm.ObjectVisibility.AllWorkFeatures = false;
                asm.ObjectVisibility.Sketches = false;
                asm.ObjectVisibility.Sketches3D = false;
            }
            string name = sdir + doc.name() + ".stp";
            dm.FileName = name;
            step.SaveCopyAs(doc, context, nvm, dm);
            RunProcessAsync(@"C:\Program Files (x86)\Adobe\Acrobat 10.0\Acrobat\Acrobat.exe", name);
            return name;
            //             System.Diagnostics.ProcessStartInfo start = new System.Diagnostics.ProcessStartInfo();
            //             start.FileName = @"C:\Program Files (x86)\Adobe\Acrobat 10.0\Acrobat\Acrobat.exe";
            //             start.Arguments = name;
            //             start.WindowStyle = System.Diagnostics.ProcessWindowStyle.Hidden;
            //             var process = System.Diagnostics.Process.Start(start);
            //             process.WaitForExit();
            //             System.IO.File.Delete(name);
            //this.Close();
        }

        static System.Threading.Tasks.Task<int> RunProcessAsync(string fileName, string name)
        {
            var tcs = new System.Threading.Tasks.TaskCompletionSource<int>();

            var process = new System.Diagnostics.Process
            {
                StartInfo = { FileName = fileName, Arguments = name },
                EnableRaisingEvents = true
            };

            process.Exited += (sender, args) =>
            {
                tcs.SetResult(process.ExitCode);
                process.Dispose();
            };

            process.Start();
            return tcs.Task;
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
            tbl = tbl ?? new TableInv(asm, I.p() + @"\Sequence.xml");
            tbl.addTable(bView);
            List<int> except = new List<int>();
            tbl.exc(ref except);
            tbl.rows.Sort((e1, e2) => e1.CompareTo(e2));
            tbl.reNumber(tbl.rows);
            if (TableInv.bvs != null)
                tbl.renumberBom(tbl.rows, TableInv.bvs);
            else tbl.renumberBom(tbl.rows, new List<BOMView>() { bView });
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
                        catch { }
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
        private void addParameter()
        {
            u.clear();
            Document doc = Macros.StandardAddInServer.m_inventorApplication.ActiveDocument;
            //             string name = ut.OFD(u.pathProj(), "XML(*.xml)|*.xml", false);
            //             if (name == null) return;
            //             XMLDoc xdoc = new XMLDoc(name, "head");
            string ffn = doc.FullFileName;
            XMLDoc xdoc = u.getXMLOFD();
            u.checkSubPath(doc, xdoc);
            xdoc.restoreName();
            xdoc.insert(f: new List<string>() { "Assembly", "Add", "Part", "Sketch", "Flange", "IMates",
                "Face", "Cut", "Plane", "CFlange", "Unfold", "Punch", "Array", "Mirror", "Hole"});
            u.paramFilter(doc, xdoc);
            //xdoc.ReplaceAlias(xdoc.El, "Alias");
            Document tdoc = u.changeProjProp(xdoc);
            if (!u.isNull(tdoc)) doc = tdoc;
            ut.addParameters(doc, xdoc, true);
            if (!u.isNull(u.sdoc) && u.sdoc.Dirty) { u.sdoc.Save(); }
            //             foreach (var el in xdoc.El.Elements("Parameter"))
            //             {
            //                 ut.addParameter(doc, el);
            //             }
            foreach (var el in xdoc.El.Elements("Plane"))
            {
                string Name = XMLDoc.getXAttributeValue(el, "Name"), BasePlane = XMLDoc.getXAttributeValue(el, "BasePlane"),
                    Offset = XMLDoc.getXAttributeValue(el, "Offset"), rev = XMLDoc.getXAttributeValue(el, "reverse");
                if (u.isNull(BasePlane)) continue;
                //if (InvDoc.Reflect.exist<WorkPlane>((doc as PartDocument).ComponentDefinition.WorkPlanes, Name)) continue;
                ut.addPlane(((PartDocument)doc).ComponentDefinition, Name, BasePlane, Offset, rev);
            }
            foreach (var el in xdoc.El.Elements("Sketch"))
            {
                string Name = XMLDoc.getXAttributeValue(el, "Name"), BasePlane = XMLDoc.getXAttributeValue(el, "BasePlane");
                if (InvDoc.Reflect.exist<PlanarSketch>((doc as PartDocument).ComponentDefinition.Sketches, Name)) continue;
                ut.addSketch(((PartDocument)doc).ComponentDefinition, Name, BasePlane);
            }
            ut.removeLink(doc, xdoc);
            ut.linkParameters(ffn, "", xdoc);
        }

        private void добавитьПараметрыToolStripMenuItem_Click(object sender, EventArgs e)
        {
            addParameter();
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
                    doc.SaveAs(name.Replace(".xml", ".ipt"), false);
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

        public static Profile addSpline(PartComponentDefinition compDef, WorkPlane wp, double start, double step, string vals, double scale)
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
            AssemblyDocument doc = Macros.StandardAddInServer.m_inventorApplication.Documents.Add(DocumentTypeEnum.kAssemblyDocumentObject, CreateVisible: true) as AssemblyDocument;
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
                iMate = (name.ToLower().EndsWith("отв")) ? name.Remove(name.Length - 3) : name;
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
            { p = compDef.Features.RectangularPatternFeatures.Add(col, dir, true, count, dist, ComputeType: PatternComputeTypeEnum.kAdjustToModelCompute); c = int.Parse(p.XCount.Value.ToString()) - 1; }
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
                c = (int.Parse(p.XCount.Value.ToString()) * int.Parse(p.YCount.Value.ToString())) - 1;
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
                InsertiMateDefinition insIMateDef1 = IMate.iMate_((hf as HoleFeature).Faces[1].Edges[numEdge], compDef, off);
                int numPat = (p.YCount == null) ? 2 : int.Parse(p.XCount.Value.ToString()) + 2;
                if (rev && p.PatternElements.Count > numPat + 1) numPat++;
                InsertiMateDefinition insIMateDef2 = IMate.iMate_(p.PatternElements[numPat].Faces[1].Edges[numEdge], compDef, off);
                swap(ref objs, rev, insIMateDef1, insIMateDef2);
                //if (comboBox3.Text != "")
                CompositeiMateDefinition compos;
                IMate.addName(compos = compDef.iMateDefinitions.AddCompositeiMateDefinition(objs), iMate);
                objs.Clear();
            }
            if (fastener != null && hf is HoleFeature)
            {
                int num = (numEdge == 2) ? 1 : 2;
                if (offsetFastener != null) off = ut.convToDouble(offsetFastener);
                InsertiMateDefinition insIMateDef = IMate.iMate_((hf as HoleFeature).Faces[1].Edges[num], compDef, off);
                IMate.addName(insIMateDef, fastener);
                //                 for (int i = 2; i <= p.PatternElements.Count; i++)
                //                 {
                //                     insIMateDef = im.iMate_(p.PatternElements[i].Faces[1].Edges[2], compDef, off);
                //                     im.addName(insIMateDef, fastener);
                //                 }
            }
            return p;
        }

        private void swap(ref ObjectCollection col, bool rev, InsertiMateDefinition im1, InsertiMateDefinition im2)
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
            if (col.Count == 0) return null;
            SketchHolePlacementDefinition def = compDef.Features.HoleFeatures.CreateSketchPlacementDefinition(col);
            HoleFeature hf;
            hf = compDef.Features.HoleFeatures.AddDrilledByDistanceExtent(def, diam, $"{compDef.Parameters["Толщина"].Name}*3", PartFeatureExtentDirectionEnum.kPositiveExtentDirection);

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
        //             XMLDoc xDoc = new XMLDoc(I.p() + @"\sheet.xml", "head");
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
            f.Height = 150 + 50; f.Width = 400; f.WindowState = FormWindowState.Normal; f.Text = "Массив"; f.StartPosition = FormStartPosition.CenterScreen;
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
            cb1.Items.AddRange(new string[] { "Отверстие", "Овал", "Другой элемент" });
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
            WorkPlane pl = ut.planeFromAxis(compDef, vec);
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

        private PlanarSketch addSketchFromDerSketch(PartComponentDefinition compDef, SketchPoint sp, string name = "", bool side = false)
        {
            UnitVector dir = (sp.Parent as PlanarSketch).PlanarEntityGeometry.Normal;
            Vector vec = dir.AsVector(); vec.ScaleBy(-1);
            Point pt = sp.Geometry3d;
            pt.TranslateBy(vec);
            PlanarSketch ps = null;
            int i = (side) ? 1 : 0;
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

        private void крепежToolStripMenuItem1_Click_1(object sender, EventArgs e)
        {
            u.transactStart(I.aDoc(), "АвтоКрепеж");
            //I.screenSilent(true);
            Fastener f = new Fastener(I.aDoc());
            f.add();
            //I.screenSilent(false);
            u.transactEnd();
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
            f.Height = 150 + 50; f.Width = 400; f.WindowState = FormWindowState.Normal; f.Text = "Массив данные"; f.StartPosition = FormStartPosition.CenterScreen;
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
            poz2 = InterfaceDll.MyLabel.position(txtbox, offsetX + 30, 0);
            txtbox = myTxt.addTextBox("", MyLabel.position(txtbox, -txtbox.Width, offsetY), w2, 15);
            f.Controls.Add(txtbox);
            lbl = mylbl1.addLabel("Отверстие", InterfaceDll.MyLabel.position(lbl, -lbl.Width, offsetY), 100, 15);
            f.Controls.Add(lbl);
            txtbox = myTxt.addTextBox("", InterfaceDll.MyLabel.position(txtbox, -txtbox.Width, offsetY), w3, 15);
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
            add.addParameters(PartsBtn.xDoc.El, new string[] { tbStep.Text, tbCount.Text });
            if (tbCountY.Text != "" && tbStepY.Text != "") add.addParameters(PartsBtn.xDoc.El, new string[] { tbStepY.Text, tbCountY.Text }, "Y");
            string suff = (tbDiamCounterName.Text.ToLower().EndsWith("отв")) ? "" : "Отв";
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
            public XElement addParameter(string n, string val, string suff = "", string typ = "mm")
            {
                return new XElement("Parameter", new XAttribute("Name", n + suff), new XAttribute("Value", val),
                    new XAttribute("Type", typ), new XAttribute("Group", group));
            }
            public void addParameters(XElement el, string[] vals, string y = "")
            {
                el.Add(addParameter(name, vals[0], suff1 + y));
                el.Add(addParameter(name, vals[1], suff2 + y, "ul"));
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
            XMLDoc im = new XMLDoc(I.p() + @"\Imate.xml", "head"),
                af = new XMLDoc(I.p() + @"\AutoFastener.xml", "Fasteners");
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
            invDoc.pathes = null;
            foreach (var item in invDoc.pathes)
            {
                if (item == doc.path() + "\\") continue;
                if (System.IO.File.Exists(item + "compare.xml"))
                {
                    XMLDoc tmp = new XMLDoc(item + "compare.xml", "head");
                    foreach (var elem in tmp.El.Elements())
                    {
                        xmlDoc.addXElement(elem);
                    }
                }
            }
            List<string> files = new List<string>();
            Macros.StandardAddInServer.m_inventorApplication.SilentOperation = true;
            files = invDoc.openFiles("*.ipt", true);
            files.AddRange(invDoc.openFiles("*.iam", true));
            files.AddRange(invDoc.openFiles("*.idw", true));
            files = files.Where(n => n.StartsWith(invDoc.path)).ToList();
            string curpath = invDoc.path;
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
                try
                {
                    Parts.replaceReference(invDoc.openDoc(item), xmlDoc, curpath);
                }
                catch (System.Exception ex)
                {
                    MessageBox.Show("Ошибка в файле " + item);
                }
            }
            Macros.StandardAddInServer.m_inventorApplication.SilentOperation = false;
        }

        private void скругленияToolStripMenuItem_Click_1(object sender, EventArgs e)
        {
            MyForm F = new MyForm("KPSInterface.xml", "Тест");
            F.f.ShowDialog();
            double rin = ut.convToDouble(F.cbs[1].Text), rout = ut.convToDouble(F.cbs[0].Text);
            foreach (Document doc in I.app.Documents.VisibleDocuments)
            {
                var def = I.getSMCD(doc);
                if (def == null) continue;
                for (int i = 0; i < def.SurfaceBodies.Count; i++)
                {
                    if (!def.SurfaceBodies[i + 1].Visible) continue;
                    fillet(def.SurfaceBodies[i + 1],
                        def.SurfaceBodies[i + 1].ConcaveEdges.Cast<Edge>(), def, rin, 3);
                    fillet(def.SurfaceBodies[i + 1],
                        def.SurfaceBodies[i + 1].ConvexEdges.Cast<Edge>(), def, rout, 3);
                }
            }
            this.Close();
        }
        private bool condition(Edge ed, double r, double t, bool kps = true)
        {
            if (ed.CurveType != CurveTypeEnum.kLineCurve) return false;
            if (condition(ed, t * 0.5)) return false;
            if (ed.StartVertex.Edges.Count != 3) return false;
            //if (ed.AssociativeID > 10) return false;
            List<Edge> edges = new List<Edge>();
            foreach (Edge e in ed.StartVertex.Edges)
            {
                if (e.Equals(ed)) continue;
                if (e.CurveType != CurveTypeEnum.kLineCurve) continue;
                if (!condition(e, r)) return false;
                edges.Add(e);
            }
            if (kps)
                if (!condition(edges)) return false;
            return true;
        }
        private bool condition(List<Edge> es)
        {
            if (es.Count != 2) return false;
            var v1 = es[0].StartVertex.Point.VectorTo(es[0].StopVertex.Point);
            var v2 = es[1].StartVertex.Point.VectorTo(es[1].StopVertex.Point);
            return v1.IsPerpendicularTo(v2, 0.01);
        }
        private bool condition(Edge e, double r)
        {
            Vector v = e.StartVertex.Point.VectorTo(e.StopVertex.Point);
            if (v.Length < r * 2) return false;
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

        private EdgeCollection fillet(SurfaceBody sb, IEnumerable<Edge> edges, SheetMetalComponentDefinition smcd,
            double r, double t)
        {
            r *= 0.1; t *= 0.1;
            EdgeCollection col = I.app.TransientObjects.CreateEdgeCollection();
            SheetMetalFeatures feat = smcd.Features as SheetMetalFeatures;
            foreach (Edge ed in edges)
            {
                if (condition(ed, r, t)) col.Add(ed);
            }
            if (col.Count != 0)
            {
                add(col, r, feat);
            }
            return col;
        }
        private EdgeCollection chamfer(SurfaceBody sb, IEnumerable<Edge> edges, SheetMetalComponentDefinition smcd,
            string v, double t)
        {
            EdgeCollection col = I.app.TransientObjects.CreateEdgeCollection();
            if (u.isNull(v)) return col;
            t *= 0.1;
            var spl = v.Split('x');
            var x = u.convToDouble(spl[0]) * 0.1;
            var y = u.convToDouble(spl[1]) * 0.1;
            var max = x > y ? x : y;

            SheetMetalFeatures feat = smcd.Features as SheetMetalFeatures;
            foreach (Edge ed in edges)
            {
                if (condition(ed, max, t, false)) col.Add(ed);
            }
            if (col.Count != 0)
            {
                addChamfer(col, x, y, feat);
            }
            return col;
        }

        public void addChamfer(EdgeCollection col, double x, double y, SheetMetalFeatures f)
        {
            try
            {
                var ch = f.ChamferFeatures.AddUsingDistance(col, x, false, false, false);
            }
            catch (Exception)
            {
            }
        }

        public void add(EdgeCollection col, double r, SheetMetalFeatures f)
        {
            try
            {
                var crd = f.CornerRoundFeatures.CreateCornerRoundDefinition(col, r);
                f.CornerRoundFeatures.Add(crd);
            }
            catch (Exception)
            {
            }
        }

        //         public void breakOp()
        //         {
        //             DrawingDocument drw = I.aDoc() as DrawingDocument;
        //             I.screenSilent(true);
        //             DrawingView dv = drw.ActiveSheet.DrawingViews[1];
        //             //BreakOperation br = dv.BreakOperations[1];
        //             double offset = 180 * 0.1, center = 200 * 0.1, gap = 0.2;
        //             offset *= dv.Scale; center *= dv.Scale;
        //             double w = dv.Width;
        // 
        //             BreakOperation br = dv.BreakOperations.Add(BreakOrientationEnum.kHorizontalBreakOrientation, I.CP2d(dv.Center.X - w / 2 + offset),
        //             I.CP2d(dv.Center.X - center / 2 + gap / 2), BreakStyleEnum.kRectangularBreakStyle, 10, gap, 1);
        //             dv.BreakOperations.Add(BreakOrientationEnum.kHorizontalBreakOrientation, I.CP2d(br.StartPoint.X + center),
        //             I.CP2d(dv.Center.X + dv.Width / 2 - offset), BreakStyleEnum.kRectangularBreakStyle, 10, gap, 1);
        // 
        //             List<Point2d> pts = new List<Point2d>();
        //             pts.Add(I.CP2d(dv.Height / 3 * dv.Scale, dv.Height / 2));
        //             pts.Add(I.CP2d(dv.Height * 2 * dv.Scale, 0));
        //             pts.Add(I.CP2d(-dv.Height * 2 * dv.Scale, 0));
        //             DrawingSketch ds = u.addSpline(dv, pts, true);
        //             u.addCut(ds);
        //             Vector2d v = I.CV2d(dv.Width / 2 / dv.Scale);
        //             u.translatePts(v, ref pts);
        //             ds = u.addSpline(dv, pts);
        //             u.addCut(ds);
        //             v = I.CV2d(-dv.Width / dv.Scale);
        //             u.translatePts(v, ref pts);
        //             ds = u.addSpline(dv, pts);
        //             u.addCut(ds);
        //             I.screenSilent(false);
        //         }

        public void breakOp()
        {
            DrawingDocument drw = I.aDoc() as DrawingDocument;
            //I.screenSilent(true);
            DrawingView dv = drw.ActiveSheet.DrawingViews[1];
            //BreakOperation br = dv.BreakOperations[1];
            double offset = 180 * 0.1, center = 200 * 0.1, gap = 0.2;
            offset *= dv.Scale; center *= dv.Scale;
            double w = dv.Width;

            BreakOperation br = dv.BreakOperations.Add(BreakOrientationEnum.kHorizontalBreakOrientation, I.CP2d(dv.Center.X - w / 2 + offset),
            I.CP2d(dv.Center.X - center / 2 + gap / 2), BreakStyleEnum.kRectangularBreakStyle, 10, gap, 1);
            dv.BreakOperations.Add(BreakOrientationEnum.kHorizontalBreakOrientation, I.CP2d(br.StartPoint.X + center),
            I.CP2d(dv.Center.X + dv.Width / 2 - offset), BreakStyleEnum.kRectangularBreakStyle, 10, gap, 1);

            List<Point2d> pts = new List<Point2d>();
            pts.Add(I.CP2d(dv.Height / 3 * dv.Scale, dv.Height / 2));
            pts.Add(I.CP2d(dv.Height * 2 * dv.Scale, 0));
            pts.Add(I.CP2d(-dv.Height * 2 * dv.Scale, 0));
            DrawingSketch ds = u.addSpline(dv, pts, true);
            u.addCut(ds);
            Vector2d v = I.CV2d(dv.Width / 2 / dv.Scale);
            u.translatePts(v, ref pts);
            ds = u.addSpline(dv, pts);
            u.addCut(ds);
            v = I.CV2d(-dv.Width / dv.Scale);
            u.translatePts(v, ref pts);
            ds = u.addSpline(dv, pts);
            u.addCut(ds);
            //I.screenSilent(false);
        }

        public void breakOp(List<DrawingDocument> drws, double gap = 0.2)
        {
            foreach (DrawingDocument drw in drws)
            {
                foreach (DrawingView dv in drw.ActiveSheet.DrawingViews.Cast<DrawingView>().Where(d => d.ViewType == DrawingViewTypeEnum.kStandardDrawingViewType))
                {
                    List<BODims> dims = getBO(dv, gap);
                    if (dims == null) continue;
                    addBO(dims, dv, gap);
                }
            }
        }

        //         public void addView(DrawingView dv, string name)
        //         {
        //             double w = dv.Width, h = dv.Height;
        //             Document doc = getModel(dv);
        //             string val = u.getPropValue(doc, name);
        //             if (val == "") return;
        //             var vals = val.Split(';');
        // 
        //         }

        public class BODims
        {
            public double sx = 0, ex = 0;
            public double sy = 0, ey = 0;
            public double scale = 0;
            public BreakOrientationEnum orient;
            public BODims(double v, bool h)
            {
                if (h)
                {
                    orient = BreakOrientationEnum.kHorizontalBreakOrientation;
                }
                else
                {
                    orient = BreakOrientationEnum.kVerticalBreakOrientation;
                }
                if (h)
                {
                    sx = v;
                }
                else
                {
                    sy = v;
                }
            }
            public void setEnd(double v, bool h, double gap)
            {
                if (h)
                {
                    ex = v + sx + gap;
                }
                else
                {
                    ey = v + sy + gap;
                }
            }
            //             public void setEnd(bool h, double v, double gap)
            //             {
            //                 if (h)
            //                 {
            //                     ex = v + sx + gap;
            //                 }
            //                 else
            //                 {
            //                     ey = v + sy + gap;
            //                 }
            //             }
            public void setScale(double v, double gap)
            {
                double tmp = (ex + ey) - (sx + sy) - gap;
                scale = tmp / v;
            }
            public void setScale(double v)
            {
                scale = v;
            }
        }

        public List<DrawingView> getDV(Sheet sh)
        {
            List<DrawingView> dvs = new List<DrawingView>();
            foreach (DrawingView item in sh.DrawingViews)
            {
                dvs.Add(item);
            }
            return dvs;
        }

        public Dictionary<DrawingView, DrawingView> addDV(Sheet sh, _Document model, List<DrawingView> dvs)
        {
            if (sh.DrawingViews.Count != 1) return null;
            Dictionary<DrawingView, DrawingView> dic = new Dictionary<DrawingView, DrawingView>();
            DrawingView dv = null, par = null;
            Point2d pos = null;
            bool first = true;
            string n = null, path = null;
            foreach (var item in dvs)
            {
                if (first)
                {
                    dv = sh.DrawingViews[1];
                    dv.Position = item.Position; dv.Scale = item.Scale; dv.Camera.ViewOrientationType = item.Camera.ViewOrientationType;
                    dv.ViewStyle = item.ViewStyle;
                    first = false;
                    dic.Add(item, dv);
                    continue;
                }
                switch (item.ViewType)
                {
                    case DrawingViewTypeEnum.kAssociativeDraftDrawingViewType:
                        break;
                    case DrawingViewTypeEnum.kAuxiliaryDrawingViewType:
                        break;
                    case DrawingViewTypeEnum.kCustomDrawingViewType:
                        break;
                    case DrawingViewTypeEnum.kDefaultDrawingViewType:
                        break;
                    case DrawingViewTypeEnum.kDetailDrawingViewType:
                        par = dic[item.ParentView];
                        DetailDrawingView ddv = item as DetailDrawingView;
                        pos = item.Position;
                        if (!item.Aligned) pos = position(par, item);
                        dv = sh.DrawingViews.AddDetailView(par, pos, ddv.ViewStyle, ddv.CircularFence, ddv.FenceCornerOne,
                            ddv.FenceCornerTwo, ddv.AttachPoint, ddv.Scale, ddv.ShowLabel) as DrawingView;
                        DetailDrawingView dd = dv as DetailDrawingView;
                        dd.DisplayConnectionLine = ddv.DisplayConnectionLine; dd.DisplayFullBoundary = ddv.DisplayFullBoundary;
                        dd.IsBreakLineSmooth = ddv.IsBreakLineSmooth;
                        //if (item.Aligned == false) dv.Aligned = false;
                        setPos(dv, item.Position, item.Scale);
                        dic.Add(item, dv);
                        break;
                    case DrawingViewTypeEnum.kDraftDrawingViewType:
                        break;
                    case DrawingViewTypeEnum.kOLEAttachmentDrawingViewType:
                        break;
                    case DrawingViewTypeEnum.kOverlayDrawingViewType:
                        break;
                    case DrawingViewTypeEnum.kProjectedDrawingViewType:
                        par = dic[item.ParentView];
                        pos = item.Position;
                        if (!item.Aligned) pos = position(par, item);
                        dv = sh.DrawingViews.AddProjectedView(par, pos, item.ViewStyle);
                        if (item.Aligned == false) dv.Aligned = false;
                        setPos(dv, item.Position, item.Scale);
                        dic.Add(item, dv);
                        break;
                    case DrawingViewTypeEnum.kSectionDrawingViewType:
                        par = dic[item.ParentView];
                        SectionDrawingView sdv = item as SectionDrawingView;
                        pos = position(item, par);
                        dv = sh.DrawingViews.AddSectionView(par, setDS(sdv, par), pos, sdv.ViewStyle, sdv.Scale, sdv.ShowLabel) as DrawingView;
                        if (item.Aligned == false) dv.Aligned = false;
                        setPos(dv, item.Position, item.Scale);
                        dic.Add(item, dv);
                        break;
                    case DrawingViewTypeEnum.kStandardDrawingViewType:
                        if (item.Name.StartsWith("0"))
                        {
                            if (n == null) n = u.getPropValue(model, "Part Number");
                            if (n == "") continue;
                            if (path == null) path = file.p(model.FullFileName);
                            string isp = u.findRegexFN(path, "\\^" + item.Name, false);
                            if (isp == "^" + item.Name) continue;
                            Document d = I.open(path + "\\" + isp, false, false);
                            dv = sh.DrawingViews.AddBaseView(d as _Document, item.Position, item.Scale, item.Camera.ViewOrientationType, item.ViewStyle);
                            dic.Add(item, dv);
                        }
                        break;
                    default:
                        break;
                }
                //dv.Name = getName(item.Name);
            }
            return dic;
        }

        public void names(Dictionary<DrawingView, DrawingView> dic)
        {
            foreach (var item in dic)
            {
                item.Value.Name = getName(item.Key.Name);
                item.Value.Label.Position = item.Key.Label.Position;
            }
        }

        public string getName(string name)
        {
            Regex r = new Regex("([-+][xy])");
            return r.Replace(name, "");
        }

        public DrawingSketch setDS(SectionDrawingView dv, DrawingView par)
        {
            SketchLine sl = dv.SectionLineSketch.SketchLines[1];
            DrawingSketch ds = par.Sketches.Add();
            ds.Edit();
            ds.SketchLines.AddByTwoPoints(sl.StartSketchPoint.Geometry, sl.EndSketchPoint.Geometry);
            ds.ExitEdit();
            return ds;
        }

        public DrawingSketch setSpline(DrawingView dv, DrawingSketch par)
        {
            SketchSpline spl = par.SketchSplines[1];
            var pts = u.getSplPoints(spl);
            DrawingSketch ds = u.addSpline(dv, pts);
            //ds.ExitEdit();
            return ds;
        }

        public DrawingSketch setSpline(DrawingView dv, Profile pr)
        {
            SketchSpline spl = pr.Parent.SketchSplines[1];
            var pts = u.getSplPoints(spl);
            //DrawingSketch ds = dv.Sketches.Add();
            //ds.Edit();
            DrawingSketch ds = u.addSpline(dv, pts);
            //ds.ExitEdit();
            return ds;
        }

        public Point2d setPos(SectionDrawingView dv, DrawingView par)
        {
            SketchLine sl = dv.SectionLineSketch.SketchLines[1];
            UnitVector2d uv = sl.Geometry.Direction;
            Point2d pt = par.Position;
            if (u.eq(uv.X, 0))
            {
                pt.X += 2; pt.Y = 0;
            }
            else
            {
                pt.Y -= 2; pt.X = 0;
            }
            return pt;
        }

        public void setPos(DrawingView dv, Point2d pos, double scale)
        {
            if (!u.eq(dv.Position, pos))
            {
                //if (dv.Aligned) dv.Aligned = false;
                //if (!dv.DisplayDefinitionInBase) dv.DisplayDefinitionInBase = true;
                dv.Position = pos;
            }
            if (dv.Scale != scale)
            {
                dv.ScaleFromBase = false; dv.Scale = scale;
            }

        }

        public Point2d position(DrawingView par, DrawingView dv)
        {
            Vector2d v = par.Position.VectorTo(dv.Position);
            if (u.eq(v.X, 0) || u.eq(v.Y, 0)) return dv.Position;

            return position(par);
        }

        public Point2d position(DrawingView par)
        {
            Point2d pt = par.Position;
            string name = par.Name;
            if (name.EndsWith("+x")) pt.X += par.Width;
            else if (name.EndsWith("-x")) pt.X -= par.Width;
            else if (name.EndsWith("+y")) pt.Y += par.Height;
            else if (name.EndsWith("-y")) pt.Y -= par.Height;
            else pt.X += par.Width;
            return pt;
        }

        public List<BODims> getBO(DrawingView dv, double gap)
        {
            bool h = true;
            Document doc = getModel(dv);
            string val = u.getPropValue(doc, "Break");
            if (val == "") return null;
            List<BODims> dims = new List<BODims>();
            var vals = val.Split(';');
            //ComponentDefinition def = I.getCD(doc);
            double scl = dv.Scale;
            double w = dv.Width;
            //u.scale(u.getMax(def.RangeBox), scl)*0.1;
            double l = dv.Left;
            double sum = 0;
            double g = 0;
            if (dv.BreakOperations.Count == (vals.Length - 1) / 2) return null;
            for (int i = 0; i < vals.Length - 1; i += 2)
            {
                double s = u.scale(u.convToDouble(vals[i], 0.1, 2), scl);
                BODims bod = null;
                sum += s;
                dims.Add(bod = new BODims(s - g, h));
                g = gap;
                bod.setScale(u.convToDouble(vals[i + 1]));
            }
            sum += u.scale(u.convToDouble(vals[vals.Count() - 1], 0.1, 2), scl);
            dims.Add(new BODims(sum, h));
            dims[dims.Count - 1].ex = w - sum;
            if (dims.Count == 0) return dims;
            return dims;
        }

        public List<BODims> getBO(IEnumerable<BreakOperation> bo, DrawingView dv, Document doc, double gap)
        {
            bool h = true;
            //int count = bo.Count();
            List<BODims> dims = new List<BODims>();
            double l = dv.Left, w = dv.Width;
            foreach (BreakOperation item in bo)
            {
                if (u.eq(item.EndPoint.X, 0)) h = false;
                double s = item.StartPoint.X + item.StartPoint.Y;
                dims.Add(new BODims(s - l, h));
                l = s;
            }
            double sum = 0;
            for (int i = 0; i < dims.Count; i++)
            {
                Transaction tr = I.beginTrans("t", doc);
                BreakOperation op = bo.ElementAt(i);
                op.Delete();
                sum += (dv.Width - w);
                dims[i].setEnd(dv.Width - w, h, gap);
                tr.End();
                I.app.TransactionManager.UndoTransaction();
            }
            if (dims.Count == 0) return dims;
            dims.Add(new BODims(w, h));
            dims[dims.Count - 1].ex = sum;
            for (int i = 0; i < dims.Count - 1; i++)
            {
                dims[i].setScale(sum, gap);
            }
            return dims;
        }

        public void addBO(List<BODims> dims, DrawingView dv, double gap)
        {
            if (dims.Count == 0) return;
            double newW = dv.Width - (dims[dims.Count - 1].sx);
            double oldW = dims[dims.Count - 1].ex;
            double scale = newW / oldW;
            //if (scale != 1) scale = (newW - dims[dims.Count - 1].sx) / (oldW - dims[dims.Count - 1].sx);
            BreakOperation br = null;
            double l = dv.Left;
            foreach (BODims item in dims)
            {
                if (item == dims[dims.Count - 1]) continue;
                double end = item.scale * newW + gap;
                br = dv.BreakOperations.Add(item.orient, I.CP2d(l + item.sx, l + item.sy), I.CP2d(l + item.sx + end, l + item.sx + end),
                    BreakStyleEnum.kRectangularBreakStyle, 10, gap, 1);
                l = br.StartPoint.X + br.StartPoint.Y;
                //oldW = newW;
                //newW = dv.Width;
            }
        }

        public object getDepth(object source, DrawingView dv)
        {
            AssemblyDocument asm = dv.ReferencedDocumentDescriptor.ReferencedDocument as AssemblyDocument;
            if (asm == null) return null;
            ObjectCollection col;
            List<string> lst;
            ComponentOccurrence occ = source as ComponentOccurrence;
            if (occ != null)
            {
                string v = u.getPropValue(occ.Definition.Document as Document, "Description");
                if (v == "") return null;
                lst = new List<string>();
                lst.Add(v);
                col = I.COC();
                u.findOcc(dv, lst, ref col);
                if (col.Count == 0) return null;
                return col;
            }
            col = source as ObjectCollection;
            if (col != null)
            {
                lst = new List<string>();
                foreach (ComponentOccurrence item in col)
                {
                    string v = u.getPropValue(occ.Definition.Document as Document, "Description");
                    if (v == "") return null;
                    lst.Add(v);
                }
                col = I.COC();
                u.findOcc(dv, lst, ref col);
                if (col.Count == 0) return null;
                return col;
            }
            return null;
        }

        public void addBOutO(DrawingView dv, BreakOutOperation par)
        {
            DrawingSketch ds = setSpline(dv, par.Profile);
            object source;
            double dept;
            par.GetDepth(out source, out dept);
            object col = getDepth(source, dv);
            if (col == null) return;
            dv.BreakOutOperations.Add(ds.Profiles.AddForSolid(), col, dept, par.SectionAllParts);
        }

        public void addBOutO(Dictionary<DrawingView, DrawingView> dic)
        {
            foreach (var item in dic)
            {
                foreach (BreakOutOperation boo in item.Key.BreakOutOperations)
                {
                    addBOutO(item.Value, boo);
                }
            }
        }

        public void addBO(Dictionary<DrawingView, List<BODims>> brs, Dictionary<DrawingView, DrawingView> dic, double gap)
        {
            foreach (var item in dic)
            {
                addBO(brs[item.Key], item.Value, gap);
            }
        }

        public List<DrawingDocument> addDrw(DocumentsEnumerator docs)
        {
            List<DrawingDocument> drws = new List<DrawingDocument>();
            DrawingDocument tmpl = docs.Cast<Document>().Where(d => d.DocumentType == DocumentTypeEnum.kDrawingDocumentObject).First() as DrawingDocument;
            List<string> vals = new List<string>() { "спереди", "0", (u.round(tmpl.ActiveSheet.Width * 10)).ToString(), "", "", "", "false", "false" };
            foreach (Document item in docs)
            {
                if (item is DrawingDocument) continue;
                drws.Add(drawing(item, vals, false, true));
            }
            return drws;
        }

        public _Document getModel(Sheet sh)
        {
            if (sh.DrawingViews.Count == 0) return null;
            return sh.DrawingViews[1].ReferencedDocumentDescriptor.ReferencedDocument as _Document;
        }

        public Document getModel(DrawingView dv)
        {
            return dv.ReferencedDocumentDescriptor.ReferencedDocument as Document;
        }

        //         public void crop(DrawingView dv, DrawingView par)
        //         {
        //             DrawingDocument drw = I.open((dv.Parent.Parent as DrawingDocument).FullDocumentName, true) as DrawingDocument;
        //             DrawingSketch ds = setSpline(dv, par);
        //             //drw = I.open(drw.FullFileName, true) as DrawingDocument;
        //             //drw.Sheets[1].Activate();
        //             //drw.SelectSet.Clear();
        //             drw.SelectSet.Select(dv.Sketches[1]);
        //             //drw.SelectSet.Select();
        //             if (drw.SelectSet.Count > 0)
        //             {
        //                 //I.app.CommandManager.StopActiveCommand();
        //                 //DrawingSketch dsv = drw.SelectSet[1] as DrawingSketch;
        //                 ControlDefinition cd = I.app.CommandManager.ControlDefinitions["DrawingCropViewCmd"];
        //                 cd.Execute2(false);
        //                 string na = I.app.CommandManager.ActiveCommand;
        //             }
        //         }

        public void addCrop(DrawingView dv, DrawingSketch parsketch)
        {
            DrawingSketch ds = setSpline(dv, parsketch);
        }

        public void addCrop(Dictionary<DrawingView, DrawingView> dic)
        {
            foreach (var item in dic)
            {
                foreach (DrawingSketch ds in item.Key.Sketches)
                {
                    if (ds.Name.EndsWith("о"))
                        addCrop(item.Value, ds);
                }
            }
        }

        public Dictionary<DrawingView, List<BODims>> getBrs(List<DrawingView> dvs, Document drw, double gap)
        {
            Dictionary<DrawingView, List<BODims>> brs = new Dictionary<DrawingView, List<BODims>>();
            foreach (var item in dvs)
            {
                brs.Add(item, getBO(item.BreakOperations.Cast<BreakOperation>().Where(p => p.IsSourceBreakOperation), item, drw, gap));
            }
            return brs;
        }

        public void templateDV(Documents docs)
        {
            List<DrawingDocument> drws = addDrw(docs.VisibleDocuments);
            DrawingDocument tmpl = docs.VisibleDocuments.Cast<Document>().Where(d => d.DocumentType == DocumentTypeEnum.kDrawingDocumentObject).First() as DrawingDocument;
            double gap = 0.2;
            DrawingView tmplDV = tmpl.ActiveSheet.DrawingViews[1];

            Sheet sh = tmpl.Sheets[1];


            List<DrawingView> dvs = getDV(sh);

            Dictionary<DrawingView, List<BODims>> brs = getBrs(dvs, tmpl as Document, gap);
            Dictionary<DrawingView, DrawingView> dic = null;

            foreach (DrawingDocument item in drws)
            {
                if (item.Equals(tmpl)) continue;

                //crop(item.ActiveSheet.DrawingViews[1], tmplDV);

                dic = addDV(item.Sheets[1], getModel(item.Sheets[1]), dvs);
                names(dic);
                if (dic != null)
                {
                    addBO(brs, dic, gap);
                    addBOutO(dic);
                    addCrop(dic);
                }

                //DrawingView dv = item.ActiveSheet.DrawingViews[1];
                //addBO(dimsH, dv, gap);
                //addBO(dimsV, dv, gap, BreakOrientationEnum.kVerticalBreakOrientation);

            }
        }

        public void imateName(AssemblyComponentDefinition acd)
        {
            if (acd == null) return;
            iMateDefinition imd1, imd2;
            string n1 = "", n2 = "", name = "iname";
            Document doc1 = null, doc2 = null;
            object o;
            var iMResults = u.gets<iMateResult>(acd.iMateResults, f => f.IsComposite);
            foreach (iMateResult imr in iMResults)
            {
                imr.GetInputs(out imd1, out o);
                imd2 = o as iMateDefinition;
                if (imd1 == null || imd2 == null) continue;
                doc1 = imr.Constraints[1].OccurrenceOne.ReferencedDocumentDescriptor.ReferencedDocument as Document;
                doc2 = imr.Constraints[1].OccurrenceTwo.ReferencedDocumentDescriptor.ReferencedDocument as Document;
                n1 = u.getPropValue(doc1, name); n2 = u.getPropValue(doc2, name);
                if (n1 == "" || n2 == "") continue;
                setIMateName(imd1, n1, n2); setIMateName(imd2, n2, n1);
            }
        }

        public void setIMateName(iMateDefinition iDef, string n1, string n2)
        {
            if (iDef.MatchList != null && (iDef.MatchList as Array).Length != 0) return;
            iDef.Name = n1 + "_" + n2;
            iDef.MatchList = new string[] { n2 + "_" + n1 };
        }

        public static void checkPDF(Document doc, string filter = "")
        {
            var ffn = doc.FullDocumentName;
            string pn = u.getPropValue(doc, "Part Number");
            if (pn == "") return;
            if (filter != "")
            {
                Regex r2 = new Regex(filter, RegexOptions.IgnoreCase);
                Match m2 = r2.Match(pn);
                if (m2.Value == null || m2.Value == "") return;
            }
            Regex r = new Regex(@"\d*\w*\.\d\d\.\d\d\d", RegexOptions.IgnoreCase);
            Match m = r.Match(pn);
            if (m.Value == null || m.Length == 0) return;
            string f = m.Value.Replace('.', '_');
            f = f.Replace('П', 'P');
            var p = file.p(ffn);
            var files = file.getFiles(file.combine(p, "PDF"), ".pdf");
            files = files.Where(e => e.IndexOf(f) != -1).OrderByDescending(e => System.IO.File.GetCreationTime(e));
            f = files.FirstOrDefault();
            if (f == null) return;
            f = file.nameWithExt(f);
            Property pr = u.getProp(doc, "Vendor");
            if (p != null) pr.Value = f;
            doc.Save();
        }

        private void browserNodesSort()
        {
            Document doc = I.aDoc();
            if (!(doc is AssemblyDocument)) return;
            BrowserPane pane = doc.BrowserPanes["Модель"];
            BrowserNode sn = null;
            sn = reorder(pane, sn, ".iam", "Сборки");
            sn = reorder(pane, sn, ".ipt", "Детали");
        }

        private ObjectCollection getNodes(BrowserNode n, string type)
        {
            ObjectCollection col = I.COC();
            var ie = u.gets<BrowserNode>(n.BrowserNodes, f => f.NativeObject is ComponentOccurrence);
            foreach (var item in ie)
            {
                ComponentOccurrence occ = item.NativeObject as ComponentOccurrence;
                if (occ.ReferencedDocumentDescriptor.FullDocumentName.EndsWith(type))
                    col.Add(item);
            }
            return col;
        }

        private BrowserNode reorder(BrowserPane p, BrowserNode sn, string type, string folder)
        {
            ObjectCollection col = getNodes(p.TopNode, type);
            BrowserNode n = null;
            if (col.Count != 0)
            {
                n = p.AddBrowserFolder(folder, col).BrowserNode;
                if (sn == null) sn = getNode(p.TopNode);
                p.Reorder(sn, false, n);
            }
            return n;
        }

        private BrowserNode getNode(BrowserNode n, int e = 1)
        {
            int i = 0;
            foreach (BrowserNode item in n.BrowserNodes)
            {
                if (i == e) return item;
                if (item.NativeObject is EndOfFeatures) i++;
            }
            return null;
        }

        public void projProp()
        {
            string prPath = I.app.DesignProjectManager.ActiveDesignProject.WorkspacePath;
            string nXML = "PrRepl.xml";
            MyXML bXML = new MyXML(nXML);
            string fn = file.fn(prPath, "\\" + nXML);
            if (file.check(fn))
            {
                MyXML newXML = new MyXML(fn);
                bXML.replace("Desc", 1, newXML.elem);
            }
            MyXML exc = new MyXML("PathFilter.xml");
            //exc.elem.Element("Filter").Add(new XElement("Value", ".ipt"));
            exc.remove("Value", "^");

            foreach (var item in bXML.elem.Elements())
            {
                string f = MyXML.getAtt(item, "desc");
                if (f == "") continue;
                var spl = file.getFiles(prPath, f, exc.elem.Element("Filter"));
                string pn = MyXML.getAtt(item, "nameProp");
                string val = MyXML.getAtt(item, "val");
                if (pn == "" || val == "") continue;
                foreach (var n in spl)
                {
                    Document doc = I.open(n);
                    u.addProp(doc, pn, val);
                }
            }
        }
        public static void Exp3DPDF(string path = "")
        {
            Inventor.ProgressBar pb = I.app.CreateProgressBar(false, I.app.Documents.VisibleDocuments.Count, "Создание 3D PDF");
            pb.Message = "Подождите...";
            //pb.UpdateProgress();
            //I.app.UserInterfaceManager.UserInteractionDisabled = true;
            ApplicationAddIn PDFAddin = null;

            foreach (ApplicationAddIn appAddin in I.app.ApplicationAddIns)
            {
                if (appAddin.ClassIdString == "{3EE52B28-D6E0-4EA4-8AA6-C2A266DEBB88}")
                {
                    PDFAddin = appAddin;
                    break;
                }
            }
            foreach (Document item in I.app.Documents.VisibleDocuments)
            {
                pb.Message = "Создается pdf для " + file.name(item.FullDocumentName);
                pb.UpdateProgress();
                if (item.DocumentType != DocumentTypeEnum.kAssemblyDocumentObject) continue;
                item.Activate();
                string fn = u.getPropValue(item, "model");
                if (fn == "") fn = file.name(item.FullDocumentName);
                Export3DPdf(PDFAddin, fn, path);
            }
            //I.app.UserInterfaceManager.UserInteractionDisabled = false;
            pb.Close();
        }
        public static void Export3DPdf(ApplicationAddIn PDFAddin, string fn, string save_path = "")
        {
            dynamic pdfConvertor3d = PDFAddin.Automation;
            //Reflect.getProp(pdfConvertor3d, "test");

            _Document oDoc = I.app.ActiveDocument;
            string path = I.dPr().TemplatesPath;
            // Create a NameValueMap object
            Inventor.NameValueMap oOptions = I.objs.CreateNameValueMap();
            var step = I.objs.CreateNameValueMap();
            step.Value["ApplicationProtocolType"] = 2;
            step.Value["IncludeSketches"] = false;
            step.Value["ExportFitTolerance"] = 0.001;
            string ffn = oDoc.FullDocumentName, p = file.p(ffn) + "PDF3D\\";
            if (save_path != "")
            {
                p = save_path + "PDF3D\\";
                file.createPath(p);
            }
            else
                file.createPath(p);
            //oOptions.Value["AttachedFiles"] = new string[] { oDoc.FullDocumentName };
            if (fn == null || fn == "") fn = file.name(ffn);
            oOptions.Value["FileOutputLocation"] = p + fn + "_3d.pdf";


            oOptions.Value["GenerateAndAttachSTEPFile"] = false;
            //oOptions.Value["STEPFileOptions "] = step;
            oOptions.Value["ExportTemplate"] = path + @"Blank.pdf";
            oOptions.Value["VisualizationQuality"] = AccuracyEnum.kLow;

            string[] sProps = new string[] {"{32853F0F-3444-11D1-9E93-0060B03C1CA6}:Part Number",
            "{32853F0F-3444-11D1-9E93-0060B03C1CA6}:Description",
            "{F29F85E0-4FF9-1068-AB91-08002B27B3D9}:Revision Number"};
            oOptions.Value["ExportAllProperties"] = false;
            oOptions.Value["ExportProperties"] = sProps;

            AssemblyComponentDefinition adef = ((AssemblyDocument)oDoc).ComponentDefinition;
            string name = adef.RepresentationsManager.ActiveDesignViewRepresentation.Name;
            oOptions.Value["ExportDesignViewRepresentations"] = new string[] { name };

            pdfConvertor3d.Publish(oDoc, oOptions);
        }

        private void dimBendTest()
        {
            DrawingDocument drw = I.aDoc() as DrawingDocument;
            var v = I.CV2d(1, 0);
            var pr = new Projections(drw.ActiveSheet.DrawingViews[1], v, I.CP2d(5, 20));
            pr.fill();
            pr.join();
            var b = I.Box(drw.ActiveSheet.DrawingViews[1]);
            var p = new Projection(b, v);
            pr.addHoles(p);
            pr.addCenter();
            pr.addBreaks();
        }
        private void projectTest()
        {
            DrawingDocument drw = I.aDoc() as DrawingDocument;
            var v = I.CV2d(1, 0);
            var dv = drw.ActiveSheet.DrawingViews[1];
            var dc = u.get<DrawingCurve>(dv.DrawingCurves,
                fi => fi.ProjectedCurveType == Curve2dTypeEnum.kCircleCurve2d);
            var p = new Projection(dc, v);
            var pb = new Projection(I.Box(dv), v);
            p.dir(pb);
        }
        private void boxTest()
        {
            DrawingDocument drw = I.aDoc() as DrawingDocument;
            var ss = drw.SelectSet;
            var boxes = new MyBoxes(drw.ActiveSheet.DrawingViews[1]);
            boxes.setBox();
            var x = boxes.getGrid(1, I.CV2d(1, 0));
            var y = boxes.getGrid(1, I.CV2d(0, 1));
            boxes.create(x, y);
            boxes.fill();
            boxes.arrange();
        }
        private void concaveTest()
        {
            PartDocument doc = I.aDoc() as PartDocument;
            var ss = doc.SelectSet;
            u.outherTest(ss[1] as Edge);
            u.outherTest(ss[2] as Edge);
        }
        public void cameraTest()
        {
            DrawingDocument drw = I.aDoc() as DrawingDocument;
            var ss = drw.SelectSet;
            Camera c = drw.ActiveSheet.DrawingViews[1].Camera;
            var mtx = c.ViewToWorld;
            double[] d = { };
            mtx.GetMatrixData(ref d);
            var pt = I.CP2d(1, 2);
            var pt3 = c.ViewToModelSpace(pt);
            var pt5 = c.ModelToViewSpace(pt3);
            var ptl1 = I.CP(1, 2, 0); var ptl2 = I.CP(1, 3, 0);
            Vector v = ptl1.VectorTo(ptl2);
            var ptc1 = c.ModelToViewSpace(ptl1); var ptc2 = c.ModelToViewSpace(ptl2);
            var v2 = ptc1.VectorTo(ptc2);
        }
        public void arrowTest()
        {
            DrawingDocument drw = I.aDoc() as DrawingDocument;
            var ss = drw.SelectSet;
            var el = ss[1];
            Reflect.getProp(el, "te");
        }
        public void arrangeTest()
        {
            DrawingDocument drw = I.aDoc() as DrawingDocument;
            var ss = drw.SelectSet;
            ArrangeDims ad = new ArrangeDims(ss.Cast<LinearGeneralDimension>(), drw.ActiveSheet.DrawingViews[1]);
            ad.sort();
            ad.setLevel();
            ad.setPos();
        }
        public void tbTest()
        {
            DrawingDocument drw = I.aDoc() as DrawingDocument;
            var tb = drw.ActiveSheet.TitleBlock;
            var ps = tb.Definition.Sketch;
            foreach (Inventor.TextBox text in ps.TextBoxes)
            {
                string rt = tb.GetResultText(text);
                string ft = text.FormattedText;
                if (ft.IndexOf("MaterialUpLine") != -1)
                {
                    drw.ActiveSheet.DrawingNotes.GeneralNotes.AddFitted(text.RangeBox.MinPoint, "s1");
                    drw.ActiveSheet.DrawingNotes.GeneralNotes.AddFitted(text.RangeBox.MaxPoint, "e1");
                }
            }
        }
        public void drawBoxTest()
        {
            DrawingDocument drw = I.aDoc() as DrawingDocument;
            var ss = drw.SelectSet;
            DrawingView dv = ss[3] as DrawingView;
            DrawBox db = new DrawBox(dv, I.CV2d(1, 0));
            db.set(ss[1]); db.set(ss[2]);
            db.setPr();
            var v = db.getDir();
            u.addDim(dv, db.getGI(0), db.getGI(1), v, db.mp);
        }
        public void alignDimTest()
        {
            DrawingDocument drw = I.aDoc() as DrawingDocument;
            var ss = drw.SelectSet;
            LinearGeneralDimension dim = ss[1] as LinearGeneralDimension;
            Point2d tp = dim.Text.Origin;
            LineSegment2d l = dim.DimensionLine as LineSegment2d;
            Point2d mp = u.midPt(l.StartPoint, l.EndPoint);
            Point2d pt = l.StartPoint;
            Vector2d n = u.normal(l, tp);
            mp.TranslateBy(n);
            if (n.Length > 0.4)
            {
                //dim.CenterText();
                dim.Text.Origin = mp;
                //u.addText(drw.ActiveSheet.DrawingViews[1], "mp", mp);
                var v = mp.VectorTo(tp);
                Point2d np = dim.Text.Origin;
                v.ScaleBy(1);
                np.TranslateBy(v);
                //u.addText(drw.ActiveSheet.DrawingViews[1], "mp1", mp);
                //dim.Text.Origin.X = mp.X; dim.Text.Origin.Y = mp.Y;
                dim.Text.Origin = np;
            }
        }
        public void nonOrdBoxTest()
        {
            DrawingDocument drw = I.aDoc() as DrawingDocument;
            var ss = drw.SelectSet;
            DrawingView dv = drw.ActiveSheet.DrawingViews[1];
            DrawingCurve dc = (ss[1] as DrawingCurveSegment).Parent;
            Vector2d v = dc.StartPoint.VectorTo(dc.EndPoint), n = I.CV2d(v);
            u.normal(n, null);
            Projection prv = new Projection(v, false), prn = new Projection(n, false);
            prv.set(dc.StartPoint, dc.EndPoint);
            u.setDist(n, 3);
            var ptn1 = I.CP2d(dc.StartPoint, n);
            n.ScaleBy(-1);
            var ptn2 = I.CP2d(dc.EndPoint, n);
            prn.set(ptn1, ptn2);
            u.addText(dv, "min", prn.pts[0].pt);
            u.addText(dv, "max", prn.pts[1].pt);
            foreach (DrawingCurve item in u.gets<DrawingCurve>(dv.DrawingCurves, fi =>
            fi.ProjectedCurveType == Curve2dTypeEnum.kCircleCurve2d))
            {
                Point2d cen = item.CenterPoint;
                bool c1 = prn.contains(cen), c2 = prv.contains(cen);
                if (c1 && c2)
                {
                    u.addText(dv, "pt", cen);
                }
            }
        }
        public void clTest()
        {
            DrawingDocument drw = I.aDoc() as DrawingDocument;
            var ss = drw.SelectSet;
            DrawingView dv = drw.ActiveSheet.DrawingViews[1];
            Centerline cl = ss[1] as Centerline;
            //double l = 0;
            //var v = cl.StartPoint.VectorTo(cl.EndPoint);
            //var n = I.CV2d(v); u.normal(n, null);
            //var dc = u.findMinDC(dv, cl.StartPoint, n,
            //    fi => fi.ProjectedCurveType == Curve2dTypeEnum.kLineSegmentCurve2d &&
            //    fi.EdgeType == DrawingEdgeTypeEnum.kUnknownEdge,
            //    (a, b) => a < b, out l);
            //if (dc != null) u.addText(dv, "p1", dc.MidPoint);
            //dc = u.findMinDC(dv, cl.EndPoint, n,
            //    fi => fi.ProjectedCurveType == Curve2dTypeEnum.kLineSegmentCurve2d &&
            //    fi.EdgeType == DrawingEdgeTypeEnum.kUnknownEdge,
            //    (a, b) => a < b, out l);
            //if (dc != null) u.addText(dv, "p2", dc.MidPoint);
            clDraw dcl = new clDraw(dv, cl);
            var vec = cl.StartPoint.VectorTo(cl.EndPoint);
            dcl.draw();
        }
        public void testRefer()
        {
            PartDocument doc = I.aDoc() as PartDocument;
            var ss = doc.SelectSet;
            Face f = ss[1] as Face;
            ReferenceFeature rf = f.CreatedByFeature as ReferenceFeature;
            var num = getNum(f, rf.Faces);
            SurfaceBody rsb = rf.ReferencedEntity as SurfaceBody;
            Face f2 = u.get<Face>(rsb.Faces, fi => u.eq(fi.Evaluator.RangeBox.MinPoint, num.MinPoint) &&
            u.eq(fi.Evaluator.RangeBox.MaxPoint, num.MaxPoint));
            string n = f2.CreatedByFeature.Name;
        }
        public Box getNum(Face f, Faces fs)
        {
            foreach (Face item in fs)
            {
                if (item.Equals(f)) return item.Evaluator.RangeBox;
            }
            return null;
        }
        public void unfoldTest()
        {
            Slice sl = new Slice(I.aDoc());
            sl.create();
        }
        public void culvTest()
        {
            PartDocument doc = I.aDoc() as PartDocument;
            var ss = doc.SelectSet;
            SketchEntity se = ss[1] as SketchEntity;
            PlanarSketch ps = se.Parent as PlanarSketch;
            SketchSpline spl = se as SketchSpline;
            var ev = spl.Geometry.Evaluator;
            double min, max;
            ev.GetParamExtents(out min, out max);
            int count = 10;
            double[] pars = new double[10];
            double[] dirs = { };
            double[] culv = { };
            pars[0] = min; pars[9] = max;
            double step = (max - min) / count;
            for (int i = 1; i < count - 1; i++)
            {
                pars[i] = min + step * i;
            }
            ev.GetCurvature(ref pars, ref dirs, ref culv);
            double tol = 10;
            double tmpR = 0;
            double par = 0;
            int s = 0;
            bool close = false;
            for (int i = 0; i < culv.Length; i++)
            {
                double r = 1 / culv[i];
                if (tmpR == 0) { tmpR = r; continue; }
                if (Math.Abs(tmpR - r) > tol)
                {
                    tmpR = 0;
                    if (i + 1 == pars.Length - 1)
                    {
                        addArc(ev, ps, pars, s, i, spl, true); close = true;
                    }
                    else
                        addArc(ev, ps, pars, s, i, spl);
                    s = i;
                }
            }
            if (!close)
            {
                addArc(ev, ps, pars, s, pars.Length, spl, true);
            }
        }
        public void addArc(Curve2dEvaluator ev, PlanarSketch ps, double[] pars, int s, int e, SketchSpline spl, bool end = false)
        {
            double sp = pars[s], ep = pars[e], mp = (pars[s] + pars[e]) * 0.5;
            object spt = u.getPointAtParam(ev, sp), ept = u.getPointAtParam(ev, ep);
            var mpt = u.getPointAtParam(ev, mp);
            if (end) ept = spl.EndSketchPoint;
            SketchArc sa = null;
            if (ps.SketchArcs.Count > 0) sa = ps.SketchArcs[ps.SketchArcs.Count];
            if (sa != null) spt = sa.StartSketchPoint;
            else spt = spl.StartSketchPoint;
            SketchArc a = ps.SketchArcs.AddByThreePoints(spt, mpt, ept);
            if (sa != null)
            {
                ps.GeometricConstraints.AddTangent((SketchEntity)sa, (SketchEntity)a);
            }
        }
        public void orient()
        {
            //MyXML xml = new MyXML("PlaneInterface.xml");
            //MyForm f = new MyForm(xml, "Плоскости"/*, null,*/ );
            //DialogResult dr = f.f.ShowDialog();
            //if (dr != DialogResult.OK) return;
            AssemblyDocument doc = I.aDoc() as AssemblyDocument;
            AssemblyComponentDefinition def = doc.ComponentDefinition;
            var ss = doc.SelectSet;
            if (ss.Count == 0) return;
            var view = I.app.ActiveView;
            var cam = view.Camera;
            var z = cam.Target.VectorTo(cam.Eye); z.Normalize();
            var y = cam.UpVector.AsVector();
            var x = z.CrossProduct(y);
            List<Vector> vecs = new List<Vector> { x, y, z };
            ComponentOccurrence occ = ss[1] as ComponentOccurrence;
            var wpsOcc = u.getWPs(occ);
            var wpsAsm = u.getWPs(def, vecs);
            var bv = I.CV(1, 1, 1);
            for (int i = 0; i < wpsOcc.Count; i++)
            {
                var wp1 = wpsAsm[i]; var wp2 = wpsOcc[i];
                var v1 = wp1.Plane.Normal; var v2 = wp2.Plane.Normal;
                double dp1 = u.dotProduct(v1.AsVector(), bv), dp2 = u.dotProduct(v2.AsVector(), bv);
                if (!u.eq(dp1, dp2)) def.Constraints.AddMateConstraint(wp1, wp2, 0);
                else def.Constraints.AddFlushConstraint(wp1, wp2, 0);
            }
        }
        public void adaptivePlane()
        {
            AssemblyDocument doc = I.aDoc() as AssemblyDocument;
            var def = doc.ComponentDefinition;
            var ss = doc.SelectSet;
            ComponentOccurrence occ = null;
            if (ss.Count == 0) return;
            foreach (var item in ss)
            {
                if (item is ComponentOccurrence) { occ = item as ComponentOccurrence; continue; }
                if (occ != null && item is WorkPlane) I.CWPP(def, occ, item as WorkPlane);
            }
        }
        public void mirrorTest()
        {
            Document doc = I.aDoc();
            MirrorIMate mim = new MirrorIMate(doc);
        }
        public void highTest()
        {
            PartDocument doc = I.aDoc() as PartDocument;
            var ss = doc.SelectSet;
            Highlight.set((Document)doc);

            //var ps = ss[1] as PlanarSketch;
            //var pt = u.get<SketchPoint>(ps.SketchPoints, fi => fi.HoleCenter);
            //var v = ps.PlanarEntityGeometry.Normal;
            //var r = u.reverse(v);
            //u.findUsingRay(doc.ComponentDefinition, pt.Geometry3d, v, (fi, l) => Highlight.add(fi));
        }
        public void convertSM()
        {
            PartDocument doc = I.aDoc() as PartDocument;
            doc.SubType = "{9C464203-9BAE-11D3-8BAD-0060B0CE6BB4}";
            doc.Save();
        }
        public void testPub()
        {
            PresentationDocument doc = I.aDoc() as PresentationDocument;
            var nvm = I.nvm;
            //var pub = doc.Publications[nvm];
            // var sb = pub.ActiveStoryboard;
        }
        public void testDataIO()
        {
            DrawingDocument drw = I.aDoc() as DrawingDocument;
            var p = file.p(drw.FullDocumentName);
            var n = file.name(drw.FullDocumentName) + ".xps";
            foreach (Sheet s in drw.Sheets)
            {
                var io = s.DataIO;
                string[] of = { };
                StorageTypeEnum[] st = { };
                io.GetOutputFormats(ref of, ref st);
                io.WriteDataToFile(of[0], System.IO.Path.Combine(p, n));
            }
        }
        public void testUnvisibleDims()
        {
            DrawingDocument drw = I.aDoc() as DrawingDocument;
            foreach (DrawingDimension d in drw.ActiveSheet.DrawingDimensions)
            {
                if (!d.Attached) continue;
            }
        }
        public void testPresentation()
        {
            Document doc = I.aDoc();
            PartComponentDefinition def = ((PartDocument)doc).ComponentDefinition;
            var pt = def.WorkPlanes[1].Plane.RootPoint;
            //pt.TranslateBy(def.WorkPlanes[1].Plane.Normal.AsVector());
            //Point[] pts = new Point[] { I.CP(1, 1, 0), I.CP(2, 2, 0), I.CP(0, 1, 0) };
            //u.clientLine(def, "test", "col", pts);
            //u.higlight(doc, def.WorkPlanes[1], null, null, null);
            var ss = doc.SelectSet;
            var edge = ss[1] as Edge;
            var euse = edge.EdgeUses[1];
            BSplineCurve2d bs = euse.Geometry as BSplineCurve2d;
            double[] poles = new double[] { }, knots = new double[] { }, weigth = new double[] { };
            int order, numPoles, numKnots;
            double plVec;
            bool ration, period, closed, planar;
            bs.GetBSplineInfo(out order, out numPoles, out numKnots, out ration, out period, out closed);
            bs.GetBSplineData(ref poles, ref knots, ref weigth);
            var ext = u.getParamExtents(bs.Evaluator);
            var p1 = u.getPointAtParam(bs.Evaluator, ext[0]);
            //u.booleanCut(doc, ss[1] as SurfaceBody, ss[2] as SurfaceBody, ss[3] as PlanarSketch);
        }
        public void testClientGraphics()
        {
            Document doc = I.aDoc();
            var ss = doc.SelectSet;
            if (ss.Count != 2) return;
            HoleModel hm = new HoleModel(doc);
            Edge e = ss[1] as Edge, e2 = ss[2] as Edge;
            Point pt1 = u.getCenter(e), pt2 = u.getCenter(e2);
            Vector v = pt1.VectorTo(pt2);
            v.Normalize();
            hm.setMtx(e, v);
            hm.find();
            var c = hm.edCount;
        }
        public void gridNote(double start, int c)
        {
            List<double> vals = new List<double>();
            var doc = I.aDoc();
            var def = I.getPCD(doc);
            var ps = def.Sketches.Add(def.WorkPlanes[1]);
            for (int i = 0; i < c; i++)
            {
                double v = Math.Pow((double)Math.Pow(2, i), (double)1 / 12);
                double n = v * start;
                double r = 1;
                if (i % 12 == 0) r *= 1.5;
                ps.SketchCircles.AddByCenterRadius(I.CP2d(0, n), r);
            }
            var prof = ps.Profiles.AddForSolid();
            var ex_def = def.Features.ExtrudeFeatures.CreateExtrudeDefinition(prof, PartFeatureOperationEnum.kNewBodyOperation);
            ex_def.SetDistanceExtent(1, PartFeatureExtentDirectionEnum.kPositiveExtentDirection);
            def.Features.ExtrudeFeatures.Add(ex_def);
        }
        public void oberton(double start, int c)
        {
            var tmp = I.aDoc();
            var p = file.p(tmp.FullDocumentName);
            List<string> names = new List<string>() { "a", "bb", "b", "c", "db", "d", "eb", "e", "f", "f#", "g", "g#" };
            for (int i = 0; i < 12; i++)
            {
                List<double> vals = new List<double>();
                var doc = I.newDoc($"{p}{names[i]}.ipt");
                var def = I.getPCD(doc);
                var ps = def.Sketches.Add(def.WorkPlanes[1]);
                double v = Math.Pow((double)Math.Pow(2, i), (double)1 / 12);
                double n = v * start;
                double r = 2;
                for (int j = 0; j < c; j++)
                {
                    ps.SketchCircles.AddByCenterRadius(I.CP2d(0, n + n*j), r);
                    r -= 0.1;
                }
                var prof = ps.Profiles.AddForSolid();
                var ex_def = def.Features.ExtrudeFeatures.CreateExtrudeDefinition(prof, PartFeatureOperationEnum.kNewBodyOperation);
                ex_def.SetDistanceExtent(1, PartFeatureExtentDirectionEnum.kPositiveExtentDirection);
                def.Features.ExtrudeFeatures.Add(ex_def);
               
                doc.Save();
            }
        }
        public void oberton_surface()
        {
            double s = 11, count = 12, n_count = 40;
            List<WorkPlane> pls = new List<WorkPlane>();
            List<PlanarSketch> pss = new List<PlanarSketch>();
            var doc = I.aDoc();
            var def = I.getPCD(doc);
            var col = I.COC();
            for (int j = 0; j < n_count; j++)
            {
                double p = j / 12.0;
                var n = s * Math.Pow(2, p);
                var pl = def.WorkPlanes.AddByPlaneAndOffset(def.WorkPlanes[1], n);
                var ps = def.Sketches.Add(pl);
                pss.Add(ps);
                var l = ps.SketchLines.AddByTwoPoints(I.CP2d(n, n), I.CP2d(n * count, n * count));
                var pr = ps.Profiles.AddForSurface(l);
                col.Add(pr);
            }
            var l_def = def.Features.LoftFeatures.CreateLoftDefinition(col, PartFeatureOperationEnum.kSurfaceOperation);
            def.Features.LoftFeatures.Add(l_def);
        }
        static void Rotate<t>(ref List<t> array, int c)
        {
            if (c < 0) throw new ArgumentException(nameof(c));
            if (c == 0) return;
            c = array.Count - c;

            int n = array.Count;
            c %= n;

            array = array.Skip(n - c).Concat(array.Take(n - c)).ToList();
        }
        public string get_chord(List<string> v, List<int> i, int s)
        {
            if (s != 0) Rotate(ref v, s);
            string r = "";
            foreach (var item in i)
            {
                r += $"sin({v[item]}*t) +";
            }
            r = r.TrimEnd('+');
            return r;
        }
        public List<Point2d> get_coord(List<Point2d> pts, List<string> names, List<int> i, int s, Dictionary<string, double> freq, string bn)
        {
            List<Point2d> ret = new List<Point2d>();
            if (s != 0)
            {
                Rotate(ref pts, s);
                Rotate(ref names, s);
            }
            foreach (var item in i)
            {
                string n = names[item];
                var pt = pts[item];
                pt = get_pt(pt, freq, bn, n);
                ret.Add(pt);
            }
            return ret;
        }
        public Asset createAsset(Document doc, TransientObjects to, byte r, byte g, byte b)
        {
            var pdoc = doc as PartDocument;
            var asset = pdoc.Assets.Add(AssetTypeEnum.kAssetTypeAppearance, "Generic");
            var col = to.CreateColor(r, g, b);

            ((ColorAssetValue)asset["generic_diffuse"]).Value = col;
            return asset;
        }
        public List<t> filter<t>(List<t> vals, List<int> pat)
        {
            List<t> r = new List<t>();
            for (int i = 0; i < pat.Count; i++)
            {
                if (pat[i] == 1) r.Add(vals[i]);
            }
            return r;
        }
        static int getNOD(int a, int b)
        {
            for (int i = (int)Math.Min(a, b); i > 1; i--)
                if (((a % i) == 0) & ((b % i) == 0))
                {
                    return i;
                }
            return 1;
        }
        static int getNOK(int a, int b)
        {
            return (int)a * b / getNOD(a, b);
        }
        static int getNOK(List<int> vals)
        {
            int nod = vals[0];
            int nok = nod;
            for (int i = 1; i < vals.Count; i++)
            {
                nod = getNOD(nod, vals[i]);
                nok = getNOK(nok, vals[i]);
            }
            return nok;
        }
        public Dictionary<string,double> get_freq(int count, double s, List<string> n)
        {
            double nts = 12.0;
            Dictionary<string, double> r = new Dictionary<string, double>();
            for (int i = 0; i < count; i++)
            {
                var v = s * Math.Pow(2, i / nts);
                r.Add(n[i], v);
            }
            return r;
        }
        public Point2d get_pt(Point2d pt, Dictionary<string,double> freq, string a, string b)
        {
            double fa = freq[a], fb = freq[b];
            double x = fb * pt.X, y = fa * pt.Y;
            x = log_scale(x); y = log_scale(y);
            return I.CP2d(x, y);
        }
        public double log_scale(double v)
        {
            return Math.Log10(v);
        }
        public List<int> modes(List<int> sc)
        {
            List<int> r = new List<int>();
            for (int i = 0; i < sc.Count; i++)
            {
                if (sc[i] == 1) r.Add(i);
            }
            return r;
        }
        public Dictionary<string, Asset> get_asset_dic(List<string> nts, List<Asset> assets)
        {
            return nts.Zip(assets, (k, v) => new { Key = k, Value = v }).ToDictionary(x => x.Key, x => x.Value);
        }
        public void draw_chords()
        {
            double bn = 110;
            var doc = I.aDoc();
            var path = file.p(doc.FullFileName);
            var def = I.getPCD(doc);
            var ss = doc.SelectSet;
            if (ss.Count == 0) return;
            var ps = ss[1] as PlanarSketch;
            var pl = ps.PlanarEntity;
            var to = I.app.TransientObjects;
            List<string> names_glob = new List<string>() { "a", "bb", "b", "c", "c#", "d", "eb", "e", "f", "f#", "g", "g#" };
            List<int> scale_glob = new List<int>() { 1, 0, 1, 0, 1, 1, 0, 1, 0, 1, 0, 1 };
            var freq_glob = get_freq(12, bn, names_glob);
            List<Point2d> coords_glob = new List<Point2d>(){I.CP2d(1, 1), I.CP2d(18,17), I.CP2d(9,8), I.CP2d(6,5), I.CP2d(5,4), I.CP2d(4,3),
                I.CP2d(7,5), I.CP2d(3,2), I.CP2d(8,5), I.CP2d(5,3), I.CP2d(16,9), I.CP2d(17,9) };
            List<int> interv = new List<int> { 0, 2, 4 };
            List<Asset> assets = new List<Asset>(){createAsset(doc, to, 255, 0, 0), createAsset(doc, to, 255, 69, 0), createAsset(doc, to, 255, 165, 0), createAsset(doc, to, 128, 128, 0),
                createAsset(doc, to, 0, 128, 0), createAsset(doc, to, 0, 255, 128), createAsset(doc, to, 0, 191, 255), createAsset(doc, to, 0,0,255), createAsset(doc, to, 0, 255, 0),
                createAsset(doc, to, 0,0,205), createAsset(doc, to, 75,0,130), createAsset(doc, to, 255,0,255)};
            var dic_asset = get_asset_dic(names_glob, assets);
            List<int> m = new List<int>() { 0, 9 };//modes(scale_glob);
            List<int> bns = new List<int>() { 0, 7};
            double offset = log_scale(bn);
            for (int i = 0; i < bns.Count; i++)
            {
                var names = new List<string>(names_glob);
                var num = bns[i];
                bn = freq_glob[names[num]];
                Rotate(ref names, bns[i]);
                var freq = get_freq(12, bn, names);
                pl = _draw_chords(m, scale_glob, coords_glob, names, offset, interv, dic_asset, freq, def, ps, pl);
            }

        }
        public object _draw_chords(List<int> m, List<int> scale_glob, List<Point2d> coords_glob, List<string> names_glob,
            double offset, List<int> interv, Dictionary<string,Asset> assets, Dictionary<string, double> freq,
            PartComponentDefinition def, PlanarSketch ps, object pl)
        {
            //double bn = 110;
            //var doc = I.aDoc();
            //var path = file.p(doc.FullFileName);
            //var def = I.getPCD(doc);
            //var ss = doc.SelectSet;
            //if (ss.Count == 0) return;
            //var ps = ss[1] as PlanarSketch;
            //var pl = ps.PlanarEntity;
            //var to = I.app.TransientObjects;
            //List<string> names_glob = new List<string>() { "a", "bb", "b", "c", "c#", "d", "eb", "e", "f", "f#", "g", "g#" };
            //var freq = get_freq(12, bn, names_glob);

            ////bn = freq[names_glob[3]];
            ////Rotate(ref names_glob, 3);
            ////freq = get_freq(12, bn, names_glob);
            //List<int> scale_glob = new List<int>() { 1, 0, 1, 0, 1, 1, 0, 1, 0, 1, 0, 1 };
            //List<Point2d> coords_glob = new List<Point2d>(){I.CP2d(1, 1), I.CP2d(18,17), I.CP2d(9,8), I.CP2d(6,5), I.CP2d(5,4), I.CP2d(4,3),
            //    I.CP2d(7,5), I.CP2d(3,2), I.CP2d(8,5), I.CP2d(5,3), I.CP2d(16,9), I.CP2d(17,9) };
            //List<int> interv = new List<int> { 0, 2, 4 };
            //List<Asset> assets = new List<Asset>(){createAsset(doc, to, 255, 0, 0), createAsset(doc, to, 255, 69, 0), createAsset(doc, to, 255, 165, 0), createAsset(doc, to, 128, 128, 0),
            //    createAsset(doc, to, 0, 128, 0), createAsset(doc, to, 0, 255, 128), createAsset(doc, to, 0, 191, 255), createAsset(doc, to, 0,0,255), createAsset(doc, to, 0, 255, 0), };

            //List<int> m = new List<int>() { 0 };//modes(scale_glob);

            for (int k = 0; k < m.Count; k++)
            {
                var scale = new List<int>(scale_glob);
                Rotate(ref scale, m[k]);
                var coords = filter(coords_glob, scale);
                var names = filter(names_glob, scale);

                var nok = getNOK(new List<int>() { 2, 4, 9 });

                for (int i = 0; i < coords.Count; i++)
                {
                    ps = def.Sketches.Add(pl);
                    var y = get_coord(coords, names, interv, i, freq, names[0]);
                    SketchLine l = null;
                    for (int j = 1; j < y.Count; j++)
                    {
                        l = ps.SketchLines.AddByTwoPoints(y[j - 1], y[j]);
                    }
                    l = ps.SketchLines.AddByTwoPoints(y[0], y[y.Count - 1]);
                    u.constr_line(ps);
                    var pr = ps.Profiles.AddForSolid();

                    var d = def.Features.ExtrudeFeatures.CreateExtrudeDefinition(pr, PartFeatureOperationEnum.kNewBodyOperation);
                    d.SetDistanceExtent(offset.ToString(), PartFeatureExtentDirectionEnum.kPositiveExtentDirection);
                    var extr = def.Features.ExtrudeFeatures.Add(d);
                    extr.SurfaceBody.Appearance = assets[names[i]];
                    //extr.SurfaceBody.Name = names[i];
                    //extr.Name = "extr " + names[i];
                    pl = def.WorkPlanes.AddByPlaneAndOffset(ps.PlanarEntity, offset.ToString());
                }
            }
            return pl;
        }
        public void equation_curves()
        {
            var doc = I.aDoc();
            var path = file.p(doc.FullFileName);
            var def = I.getPCD(doc);
            var ss = doc.SelectSet;
            if (ss.Count == 0) return;
            var ps = ss[1] as PlanarSketch;
            var pl = ps.PlanarEntity;
            double s = 110;
            List<string> vals = new List<string>() { $"{s}", $"9 / 8 * {s}", $"5 / 4 * {s}", $" 4 / 3 * {s }", $"3 / 2 * { s }", $"5 / 3 * {s}",  $"16 / 9 * {s}"};
            List<int> interv = new List<int> { 0, 2, 4 };
            var to = I.app.TransientObjects;
            List<Color> cols = new List<Color>() { to.CreateColor(255,0,0), to.CreateColor(255, 69, 0), to.CreateColor(255, 165, 0),
                to.CreateColor(128,128, 0), to.CreateColor(0, 128, 0), to.CreateColor(0, 255, 128), to.CreateColor(0, 191, 255), to.CreateColor(0,0,255), to.CreateColor(0,0,255),};
            List<Profile> prs = new List<Profile>();

            List<Asset> assets = new List<Asset>(){createAsset(doc, to, 255, 0, 0), createAsset(doc, to, 255, 69, 0), createAsset(doc, to, 255, 165, 0), createAsset(doc, to, 128, 128, 0),
                createAsset(doc, to, 0, 128, 0), createAsset(doc, to, 0, 255, 128), createAsset(doc, to, 0, 191, 255), createAsset(doc, to, 0,0,255), createAsset(doc, to, 0, 255, 0), };

            List<string> names = new List<string>() { "AM", "Bm", "C#m", "DM", "EM", "F#m", "G#Dim" };

            double offset = 2;

            for (int i = 0; i < vals.Count; i++)
            {               
                ps = def.Sketches.Add(pl);
                string y = get_chord(vals, interv, i);
                var ec = ps.SketchEquationCurves.Add(CurveEquationTypeEnum.kParametric, CoordinateSystemTypeEnum.kCartesian, "t", y, 0, 120);
                ec.OverrideColor = cols[i];
                var pr = ps.Profiles.AddForSurface(ec);

                var d = def.Features.ExtrudeFeatures.CreateExtrudeDefinition(pr, PartFeatureOperationEnum.kSurfaceOperation);
                d.SetDistanceExtent(offset.ToString(), PartFeatureExtentDirectionEnum.kPositiveExtentDirection);
                var extr = def.Features.ExtrudeFeatures.Add(d);
                var col = I.app.TransientObjects.CreateFaceCollection();
                col.Add(extr.SurfaceBody.Faces[1]);
                var th = def.Features.ThickenFeatures.Add(col, 0.001, PartFeatureExtentDirectionEnum.kSymmetricExtentDirection, PartFeatureOperationEnum.kNewBodyOperation);
                th.SurfaceBody.Appearance = assets[i];
                th.SurfaceBody.Name = names[i];
                pl = def.WorkPlanes.AddByPlaneAndOffset(ps.PlanarEntity, offset.ToString());
            }
        }
        public PlanarSketch getSketch(SheetMetalComponentDefinition def, Arc3d a, Edge e, ref List<SketchArc> arcs, ref List<SketchLine> lines)
        {
            PlanarSketch ps;
            Vector2d v;
            SketchPoint sp1, sp2;
            Face f = u.get<Face>(e.Faces, el => el.CreatedByFeature is PunchToolFeature);
            if (f != null) {
                PunchToolFeature punch = f.CreatedByFeature as PunchToolFeature;
                var pcp = punch.PunchCenterPoints;
                if (pcp.Count < 2) return null;
                sp1 = pcp[1] as SketchPoint; sp2 = pcp[2] as SketchPoint;
                ps = sp1.Parent as PlanarSketch;
                v = sp1.Geometry.VectorTo(sp2.Geometry); v.Normalize();
                v.ScaleBy(a.Radius);

                foreach (SketchPoint item in pcp.Cast<SketchPoint>().OrderBy(el => el.Geometry.X).ThenBy(el => el.Geometry.Y))
                {
                    arcs.Add(addArc(ps, item, v));
                }
            }
            else
            {
                f = u.get<Face>(e.Faces, fi => fi.SurfaceType == SurfaceTypeEnum.kPlaneSurface);
                ps = def.Sketches.Add(f);
                var el = u.gets<EdgeLoop>(f.EdgeLoops, fi => !fi.IsOuterEdgeLoop && fi.Edges.Count == 4);
                List<SketchPoint> pts = new List<SketchPoint>();
                foreach (EdgeLoop item in el)
                {
                    var arc = u.gets<Edge>(item.Edges, fi => fi.GeometryType == CurveTypeEnum.kCircularArcCurve);
                    List<SketchArc> lst = new List<SketchArc>();
                    foreach (var ar in arc)
                    {
                        var sa = ps.AddByProjectingEntity(ar) as SketchArc;
                        sa.Construction = true;
                        lst.Add(sa);
                    }
                    SketchLine sl = ps.SketchLines.AddByTwoPoints(lst[0].CenterSketchPoint, lst[1].CenterSketchPoint); sl.Construction = true;
                    SketchPoint sp = ps.SketchPoints.Add(I.CP2d());
                    ps.GeometricConstraints.AddMidpoint(sp, sl);
                    pts.Add(sp);
                }
                v = pts[0].Geometry.VectorTo(pts[1].Geometry); v.Normalize();
                v.ScaleBy(a.Radius);
                foreach (SketchPoint item in pts.OrderBy(fi => fi.Geometry.X).ThenBy(fi => fi.Geometry.Y))
                {
                    arcs.Add(addArc(ps, item, v));
                }
                var a1 = arcs[1];
                foreach (var item in arcs)
                {
                    if (item.Equals(a1)) continue;
                    ps.GeometricConstraints.AddEqualRadius(a1 as SketchEntity, item as SketchEntity);
                }
            }
            return ps;
        }
        public void slotSurface()
        {
            Document doc = I.aDoc();
            var def = I.getSMCD(doc);
            var ss = doc.SelectSet;
            if (ss.Count != 1) return;
            Edge e = ss[1] as Edge;
            if (e == null) return;
            Arc3d a = e.Geometry as Arc3d;
            if (a == null) return;
            List<SketchArc> arcs = new List<SketchArc>();
            List<SketchLine> lines = new List<SketchLine>();
            var ps = getSketch(def, a, e, ref arcs, ref lines);

            for (int i = 0; i < arcs.Count - 1; i++)
            {
                lines.Add(joinArcs(ps, arcs[i], arcs[i + 1]));
            }
            var prof = ps.Profiles.AddForSurface(arcs[0]);
            var sdef = def.Features.ExtrudeFeatures.CreateExtrudeDefinition(prof, PartFeatureOperationEnum.kSurfaceOperation);
            sdef.SetDistanceExtent(def.Thickness, PartFeatureExtentDirectionEnum.kNegativeExtentDirection);
            def.Features.ExtrudeFeatures.Add(sdef);
        }
        public SketchArc addArc(PlanarSketch ps, SketchPoint pt, Vector2d dir)
        {
            Point2d p1 = pt.Geometry.Copy(), p2 = pt.Geometry.Copy();
            p1.TranslateBy(dir);
            dir.ScaleBy(-1);
            p2.TranslateBy(dir);
            var a = ps.SketchArcs.AddByCenterStartEndPoint(pt, p1, p2);
            if (!a.CenterSketchPoint.Equals(pt))
            {
                a.CenterSketchPoint.Merge(pt);
            }
            return a;
        }

        public SketchPoint getPointArc(SketchArc a, Point2d pt)
        {
            List<SketchPoint> pts = new List<SketchPoint> { a.StartSketchPoint, a.EndSketchPoint };
            return pts[0].Geometry.DistanceTo(pt) <= pts[1].Geometry.DistanceTo(pt) ? pts[0] : pts[1];
        }

        public SketchLine joinArcs(PlanarSketch ps, SketchArc a1, SketchArc a2)
        {
            SketchPoint sp1 = getPointArc(a1, a2.CenterSketchPoint.Geometry), sp2 = getPointArc(a2, a1.CenterSketchPoint.Geometry);
            return ps.SketchLines.AddByTwoPoints(sp1, sp2);
        }

        private void тестToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //Point p1 = I.CP(1, 2, 1), p2 = I.CP(0, 0, 0), p3 = I.CP(3, 3, 3);
            //u.isPointOnEdge(p1, p2, p3);
            var doc = I.aDoc();
            var p = file.p(doc.FullDocumentName);
            MyXML xml = new MyXML($"{p}dims.xml", "head");
            DrawDims d = new DrawDims(xml.elem);
            //AddDeleteReplace r = new AddDeleteReplace(doc, this);
            //Repair r = new Repair(doc);
            
            //var def = I.getSMCD(doc);
            //foreach (PartFeature item in def.SurfaceBodies[1].AffectedByFeatures)
            //{
            //    if (item.Type == ObjectTypeEnum.kFaceFeatureObject)
            //    {
            //        FaceFeature f = item as FaceFeature;
            //        var bf = f.BendFeature;
            //    }
            //}
            //var asmPart = new Variable(doc as AssemblyDocument);
            //asmPart.addFromAsm();
            //SketchCopy copy = new SketchCopy(doc);
            //copy.add(ss[1] as PlanarSketch);
            //Variable var = new Variable(I.aDoc());
            //var.add_folder();
            //CutFlatPattern cut = new CutFlatPattern(I.aDoc());
            //var doc = I.aDoc();
            //var smcd = I.getSMCD(doc);
            //var f = smcd.SurfaceBodies[1].Faces[1];
            //var ss = I.getSS(doc);
            //PlanarSketch ps = ss[1] as PlanarSketch;
            //var cm = I.app.CommandManager;
            //var c = cm.ControlDefinitions["SketchProjectFlatPatternCmd"];
            //ps.Edit();
            ////ps.AddByProjectingEntity(f);
            //ss.Clear();
            //ss.Select(smcd.SurfaceBodies[1].Faces[2]);
            //cm.DoSelect(smcd.SurfaceBodies[1].Faces[2]);
            //c.Execute2(true);
            //ps.ExitEdit();
            //var mp = new MergePoints(ps);

            //MultiCut mc = new MultiCut(I.aDoc());

            //add_rects();
            //projectProperties.minMKarts();
            //slotSurface();
            //testClientGraphics();
            //testUnvisibleDims();
            //testPub();
            //convertSM();
            //Macros.MyEvArgs args = new Macros.MyEvArgs();
            //var b = sender as ToolStripMenuItem;
            //args.name = b.Name;
            //OnMyEvent(args);
            //this.Close();


            //this.Hide();
            //unmanaged.setForeground("Inventor");
            //I.app.CommandManager.Pick(SelectionFilterEnum.kAllCircularEntities, "тест");


            //mirrorTest();
            //adaptivePlane();
            //orient();
            //culvTest();
            //tbTest();
            //clTest();
            //unfoldTest();
            //testRefer();
            //alignDimTest();
            //nonOrdBoxTest();
            //drawBoxTest();
            //arrangeTest();
            //concaveTest();
            //cameraTest();
            //boxTest();
            //arrowTest();
            //arrangeTest();
            //highTest();
            //dimBendTest();
            //projectTest();
            //Exp3DPDF();
            //Draw dr = new Draw(I.aDoc() as DrawingDocument);


            //this.Close();

            //             foreach (DrawingCurve dc in dv.DrawingCurves)
            //             {
            //                 EdgeProxy ep = dc.ModelGeometry as EdgeProxy;
            //                 if (ep == null) continue;
            //                 Document doc = ep.ContainingOccurrence.Definition.Document as Document;
            //                 Property p = u.getProp(doc, "Description");
            //                 if (p == null) continue;
            //                 if (p.Value.ToString() == "Фланец правый")
            //                 {
            //                     dc.LineType = LineTypeEnum.kDottedLineType;
            //                 }
            //             }

            //             string p;
            //             I.aDoc().Close();
            //             DesignProject pr = I.app.DesignProjectManager.ActiveDesignProject;
            //             I.app.DesignProjectManager.DesignProjects[1].Activate();
            //             u.addPrPath(pr.FullFileName, "test", "path");
            //             pr.Activate(); 
            //              ProjectPaths ps = pr.LibraryPaths;
            //             ps.Add("БВ", ".\\Завеса");
            //             p = pr.FrequentlyUsedPaths[1].Path;
            //             p = file.trimEnd(p, "\\", 1);
            //             p = p + "Завеса";
            //             pr.FrequentlyUsedPaths[1].Path = p;
            //             string p = I.curProjPath();
            //             var sd = file.subDir(p, new string [] {"OldVersions"}, System.IO.SearchOption.AllDirectories);
            //             XMLDoc xml = u.XOFD( custom: sd);
            //             ContentOp.addColumn(xml);


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
            //             DrawingDocument from = I.app.Documents.Open(I.p() + @"\Cluster.idw", false) as DrawingDocument;
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
                    if (par.ParameterType == ParameterTypeEnum.kUserParameter) ut.parameterToXML(xmlDoc, par);
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
                double d = b1.Position.X - b2.Position.X;
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
                    if (b[i - 1].BalloonValueSets.Count == 1) offset -= 0.25 * r;
                    y = offset + r / 2;
                }
                pt = I.tg.CreatePoint2d(x, y);
                b[i].Position = pt;
            }
        }

        private void регионToolStripMenuItem_Click_1(object sender, EventArgs e)
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
        public class Bim
        {
            public List<string> docs;
            string suff = "_sat.ipt";
            string p;
            public Bim(string path, string suf)
            {
                docs = new List<string>();
                p = path;
                suff = suf;
            }
            public void get()
            {
                var ie = file.getFiles(p, suff);
                foreach (var item in ie)
                {
                    docs.Add(item);
                }
            }
            public void save()
            {
                I.silent(true);
                foreach (var item in docs)
                {
                    Document doc = I.open(item);
                    string t = u.getPropValue(doc, "Type");
                    PartComponentDefinition pcd = I.getPCD(doc);
                    BIMComponent b = pcd.BIMComponent;
                    b.ComponentDescription.OrientationType = BIMComponentOrientationTypeEnum.kViewCubeOrientationType;
                    b.ComponentDescription.ViewCubeOrientationOrigin = pcd.WorkPoints[1].Point;
                    if (file.check(file.p(item) + t + ".adsk") || file.check(file.p(item) + file.name(item) + ".adsk")) continue;
                    if (t != "")
                        b.ExportBuildingComponent(file.p(item) + t + ".adsk");
                    else b.ExportBuildingComponent(file.p(item) + file.name(item) + ".adsk");
                }
                I.silent(false);
            }
            public void saveSAT()
            {
                I.silent(true);
                foreach (var item in docs)
                {
                    Document doc = I.open(item);
                    string fn = doc.FullFileName;
                    var acd = I.getACD(doc);
                    acd.DataIO.WriteDataToFile("ACIS SAT", file.p(fn) + file.name(fn) + ".SAT");
                    doc.Close();
                }
                I.silent(false);
            }
        }
        public class Asms
        {
            public List<string> asms;
            public List<string> patches;
            public string path;
            public Parts p;
            static bool rev = false;
            static Vector v = null;
            public const string suff = "_adsk.ipt";
            public const string subPath = "adsk";
            public Asms(Document doc)
            {
                p = new Parts(doc);
                asms = new List<string>();
                patches = new List<string>();
                path = file.p(doc.FullFileName);
            }
            public void get()
            {
                var ie = file.getFiles(path, suff);
                foreach (var item in ie)
                {
                    if (findAsm(item))
                        patches.Add(item);
                }
            }
            public void create()
            {
                for (int i = 0; i < asms.Count; i++)
                {
                    create(asms[i], patches[i]);
                }
            }
            private void create(string a, string b)
            {
                I.silent(true);
                AssemblyDocument asm = p.addAsmDoc();
                AssemblyComponentDefinition acd = asm.ComponentDefinition;
                Matrix m = I.getMatrix();
                //UnitVector up1 = I.CUV(0, 0, 1), up2 = I.CUV(0, 0, 1), front1 = I.CUV(0, 0, 1), front2 = I.CUV(0, 0, 1);
                ComponentOccurrence ocp = acd.Occurrences.Add(b, m), oca = acd.Occurrences.Add(a, m);
                rotate(ocp, acd);
                rotate(oca, acd);
                ocp.Grounded = true;
                //oca.Grounded = false; ocp.Grounded = false;
                //getAxis(ocp, ref up1, ref front1);

                //getAxis(oca, ref up2, ref front2);
                //                 if (up1.IsParallelTo(up2, 0.2))
                //                 {
                //                     setAxis(up1, acd, ocp, true, ref m);
                //                     setAxis(front1, acd, ocp, false, ref m);
                //                     ocp.Transformation = m;
                //                     m = I.getMatrix();
                //                     setAxis(up2, acd, oca, true, ref m);
                //                     setAxis(front2, acd, oca, false, ref m);
                //                     oca.Transformation = m;
                //                 }
                //                 else
                //                 {
                //                     setAxis(front1, front2, acd, oca, false, ref m);
                //                     setAxis(up1, up2, acd, oca, true, ref m);
                //                 }

                string t = u.getPropValue(oca.Definition.Document as Document, "Type");

                string sp = file.combine(path, subPath);
                if (t != "")
                {
                    u.addProp(asm as Document, "Type", t);
                }
                if (!file.check(sp + file.name(a) + "_asm.iam"))
                    asm.SaveAs(sp + file.name(a) + "_asm.iam", false);
                asm.Close(true);
                I.silent(false);
            }
            public void rotate(ComponentOccurrence co, AssemblyComponentDefinition acd)
            {
                Matrix m = I.getMatrix();
                co.Grounded = false;
                UnitVector up = I.CUV(0, 0, 1), front = I.CUV(0, 1, 0);
                getAxis(co, ref up, ref front);
                setAxis(up, acd, co, true, ref m);
                setAxis(front, acd, co, false, ref m);
                co.Transformation = m;
            }
            public void setAxis(UnitVector v2, AssemblyComponentDefinition asm, ComponentOccurrence oc, bool top, ref Matrix m, bool tr = false)
            {
                Vector axis = I.CV();
                if (top) axis.Z = 1;
                else
                {
                    axis.Y = 1;
                    if (tr) axis.Y = -1;
                }
                //WorkAxis w1 = getWA(asm as ComponentDefinition, v1), w2 = getWA(oc.Definition, v2) ;
                //if (w2 == null && w1 == null) return;
                //object r; rev = true;
                //                 if (v2.DotProduct(w2.Line.Direction) < 0)
                //                 {
                //                     rev = true;
                //                 }
                //else rev = false;
                //oc.CreateGeometryProxy(w2, out r);
                //Matrix m = I.tg.CreateMatrix();
                //m.PreMultiplyBy(m);
                //                 if (axis.DotProduct(v2.AsVector()) < 0)
                //                 {
                //                     if (top) m.set_Cell(3, 3, -1);
                //                     else m.set_Cell(2, 2, -1);
                //                 }
                Matrix tmp = I.tg.CreateMatrix();
                tmp.SetToRotateTo(v2.AsVector(), axis);
                //oc.Transformation = m;
                m.PostMultiplyBy(tmp);
                //                 double[] d = new double[] { };
                //                 m.GetMatrixData(ref d);
                //                 d[1] = 1;
                //                 tmp.GetMatrixData(ref d);
                //                 d[1] = 1;
                //oc.Transformation = m;
                // double a = w1.Line.Direction.AngleTo(w2.Line.Direction);
                //                 if (rev)
                //                 {
                //                     Vector invV = v.Copy(); invV.ScaleBy(-1);
                //                     m.SetToRotateTo(v, invV);
                //                     oc.Transformation = m;
                //                 }
                //                 asm.Constraints.AddMateConstraint(w1, r, 0);

            }
            public void translateToCenter(string a)
            {
                I.silent(true);
                AssemblyDocument asm = p.addAsmDoc();
                AssemblyComponentDefinition acd = asm.ComponentDefinition;
                Matrix m = I.getMatrix();
                UnitVector up1 = I.CUV(0, 0, 1), up2 = I.CUV(0, 0, 1), front1 = I.CUV(0, 0, 1), front2 = I.CUV(0, 0, 1);
                ComponentOccurrence oca = acd.Occurrences.Add(a, m);
                oca.Grounded = false;

                getAxis(oca, ref up2, ref front2);

                setAxis(up2, acd, oca, true, ref m);
                setAxis(front2, acd, oca, false, ref m);

                acd = oca.Definition as AssemblyComponentDefinition;
                MassProperties mass = acd.MassProperties;
                Point pt = mass.CenterOfMass; pt.TransformBy(m);
                Vector tr = I.CV(-pt.X, -pt.Y, -pt.Z);

                m.SetTranslation(tr);
                oca.Transformation = m;


                string t = u.getPropValue(oca.Definition.Document as Document, "Type");

                string sp = path;
                if (t != "")
                {
                    u.addProp(asm as Document, "Type", t);
                }
                asm.SaveAs(sp + file.name(a) + "_.iam", false);
                asm.Close(true);
                I.silent(false);
            }
            public WorkAxis getWA(ComponentDefinition def, UnitVector v)
            {
                WorkAxes wa = null;
                PartComponentDefinition pc = def as PartComponentDefinition;
                if (pc != null) wa = pc.WorkAxes;
                AssemblyComponentDefinition ac = def as AssemblyComponentDefinition;
                if (ac != null) wa = ac.WorkAxes;
                Asms.v = null;
                foreach (WorkAxis item in wa)
                {
                    if (v.IsParallelTo(item.Line.Direction, 0.2))
                    {
                        Asms.v = v.AsVector();
                        rev = v.IsEqualTo(item.Line.Direction, 0.2) ? false : true;
                        return item;
                    }
                }
                return null;
            }
            public void getAxis(ComponentOccurrence oc, ref UnitVector up, ref UnitVector front)
            {
                Document doc = I.open(oc.ReferencedDocumentDescriptor.FullDocumentName, true, false);
                Inventor.View v = doc.Views[1];
                Camera c = v.Camera;
                c.ViewOrientationType = ViewOrientationTypeEnum.kTopViewOrientation;
                //up = c.UpVector;
                front = c.UpVector;
                c.ViewOrientationType = ViewOrientationTypeEnum.kFrontViewOrientation;
                //front = c.UpVector;
                up = c.UpVector;
                //doc.Close();
            }
            private bool findAsm(string ffn)
            {
                string f = file.sub(ffn, 0, -9) + ".iam";
                if (file.check(f))
                {
                    asms.Add(f);
                    return true;
                }
                return false;
            }
        }
        public class Sat
        {
            public string ffn;
            public Document doc;
            public AssemblyDocument asm;
            public DerivedAssemblyDefinition adef;
            public PartComponentDefinition def;
            public Parts p;
            string t;
            public bool op = false;
            public Sat(Document d)
            {
                p = new Parts(d);
                asm = d as AssemblyDocument;
                t = u.getPropValue(asm as Document, "Type");
                ffn = d.FullFileName;
            }
            public void open()
            {
                if (System.IO.File.Exists(ffn)) { doc = p.openDoc(ffn); op = true; }
                else doc = p.addPrtDoc() as Document;
            }
            public void run()
            {
                ffn = nameForSave();
                open();
                I.silent(true);
                addAsm();
                save();
                I.silent(false);
            }
            public void addAsm()
            {
                def = I.getPCD(doc);
                if (op) adef = def.ReferenceComponents.DerivedAssemblyComponents[1].Definition;
                else adef = def.ReferenceComponents.DerivedAssemblyComponents.CreateDefinition(asm.FullFileName);
                if (adef == null) return;
                adef.DeriveStyle = DerivedComponentStyleEnum.kDeriveAsSingleBodyNoSeams;
                adef.InclusionOption = DerivedComponentOptionEnum.kDerivedIncludeAll;
                adef.IncludeAllTopLevelParameters = DerivedComponentOptionEnum.kDerivedExcludeAll;
                adef.IncludeAllTopLevelSketches = DerivedComponentOptionEnum.kDerivedExcludeAll;
                adef.IncludeAllTopLevelWorkFeatures = DerivedComponentOptionEnum.kDerivedExcludeAll;
                adef.ReducedMemoryMode = true;
                if (op) def.ReferenceComponents.DerivedAssemblyComponents[1].Definition = adef;
                else def.ReferenceComponents.DerivedAssemblyComponents.Add(adef);
            }
            public void save()
            {
                if (t != null && t != "")
                {
                    u.addProp(doc, "Type", t);
                }
                if (op)
                    doc.Save2(true);
                else
                    doc.SaveAs(ffn, false);
            }
            public string nameForSave()
            {
                return file.p(ffn) + file.name(ffn) + "_sat.ipt";
            }
        }
        private void sketch()
        {
            Document doc = I.aDoc();
            MyXML xml = new MyXML("SketchInterface.xml");
            string path = u.pathProj();
            xml.cbel(path);
            MyForm F = new MyForm(xml, "Создать");
            if (F.f.ShowDialog() != DialogResult.OK) return;
            Elements.chks = F.getChks();
            if (F.cbs[0].Text.EndsWith(".xml")) Elements.path = path + "\\" + F.cbs[0].Text;
            Elements els = new Elements(doc);
        }
        private void create()
        {
            u.clear();
            Docs.clearStatic();
            Document doc = Macros.StandardAddInServer.m_inventorApplication.ActiveDocument;
            MyXML xml = new MyXML("CreateInterface.xml");
            string path = InvDoc.u.pathProj();
            xml.cbel(path);
            //             XElement cbel = xml.elem.Descendants("CB").FirstOrDefault();
            //             if (cbel != null)
            //             {
            //                 xml.set(cbel);
            //                 cbel.RemoveNodes();
            //                 foreach (var item in System.IO.Directory.GetFiles(path, "*.xml", System.IO.SearchOption.TopDirectoryOnly))
            //                 {
            //                     XElement tmp = new XElement("el", new XAttribute("val", System.IO.Path.GetFileName(item)));
            //                     //if (xml.find("val", item, 1) == null)
            //                     xml.addElem(tmp);
            //                 }
            //                 xml.setRoot();
            //             }
            MyForm F = new MyForm(xml, "Создать");
            var dr = F.f.ShowDialog();
            /*            return;*/
            if (dr != DialogResult.OK) return;
            Docs docs = new Docs();
            Docs.chk.Clear();
            if (F.chks != null)
            {
                for (int i = 0; i < F.chks.Count(); i++)
                {
                    Docs.chk.Add(F.chks[i].Checked);
                }
            }
            if (docs.empty()) return;
            string name = path + "\\" + F.cbs[0].Text;
            //ut.OFD(System.IO.Path.GetDirectoryName(doc.FullDocumentName), "XML files(*.xml)|*.xml");
            XMLDoc xdoc = new XMLDoc(name, "head");
            u.checkSubPath(doc, xdoc);
            xdoc.insert(f: new List<string>() {"Sketch", "Flange", "IMates", "Face", "Cut", "Plane", "CFlange", "Unfold", "Punch", "Array",
            "Mirror", "Hole", "AsmConstr", "FP", "Parameter", "Fillet"}, fpath: I.p() + @"\xml");
            XMLDoc.replaceId(xdoc);
            xdoc.setRoot();
            XMLDoc.removeEl(xdoc.El, "XML");
            u.paramFilter(doc, xdoc);

            I.silent(true);
            foreach (var el in xdoc.Doc.Descendants("Assembly"))
            {
                if (el.Parent.Name != "head") continue;
                if (el != null)
                {
                    string typ = el.Attribute("Type").Value;
                    create(el, typ, 1, xdoc: xdoc);
                    docs.clear();
                    docs = create(el, typ, 1, 1);
                    createAssembly(el, typ, docs);
                }
            }
            foreach (_Document item in Macros.StandardAddInServer.m_inventorApplication.Documents)
            {
                if (item.Open == false)
                {
                    item.ReleaseReference();
                    item.Close();
                }
            }
            I.silent(false);
        }

        private void создатьДеталиToolStripMenuItem_Click(object sender, EventArgs e)
        {
            create();
        }

        public class Docs
        {
            public HashSet<string> docs;
            public HashSet<string> adds;
            public HashSet<Perf> perfs;
            public static HashSet<string> names;
            static public List<bool> chk = new List<bool>();
            public Docs()
            {
                docs = new HashSet<string>();
                adds = new HashSet<string>();
                perfs = new HashSet<Perf>();
            }
            public bool empty()
            {
                foreach (var item in chk)
                {
                    if (item) return false;
                }
                return true;
            }
            public static void clearStatic()
            {
                if (!u.isNull(Docs.names)) names.Clear();
                else names = new HashSet<string>();
            }
            public void clear()
            {
                docs.Clear();
                adds.Clear();
                perfs.Clear();
            }
        }

        public class Perf
        {
            int num;
            string pn, d, typ, dec;
            string path;
            string comment;
            XElement el;
            static XElement b;
            public AssemblyDocument doc;
            public Perf(int n, string typ, string dec, string d, XElement el, string p)
            {
                num = n;
                this.dec = dec; this.typ = typ;
                this.pn = typ + "." + dec; this.d = d;
                this.el = el;
                path = MyXML.getAtt(el, "Path");
                if (path == "") path = p;
                comment = MyXML.getAtt(el, "Comment");
                if (!path.EndsWith("\\")) path += "\\";
            }
            public string nameForSave()
            {
                return pn + "-" + num.ToString("00") + " (" + d + ")^" + num.ToString("00") + ".iam";
            }
            public string nameForOpen()
            {
                return pn + " (" + d + ").iam";
            }
            public void adr()
            {
                b = new XElement("el");
                createXML(el, "el");
                Parts.removeReplace(doc.ComponentDefinition, b);
                b = null;
            }
            public void createXML(XElement el, string name)
            {
                string tmpPath = MyXML.getAtt(el, "Path");
                foreach (var item in el.Elements())
                {
                    XElement tmp = new XElement(name);
                    string a = "";
                    if (item.Name == "Part") a = ".ipt";
                    else if (item.Name == "Assembly") a = ".iam";
                    string attName = MyXML.getAtt(item, "Name");
                    string n = tmpPath + attName + a;
                    if (attName.StartsWith("v3#")) n = attName;
                    if (item.Name == "Add") createXML(item, "add");
                    else if (item.Name == "Delete") createXML(item, "delete");
                    else if (item.Name == "Replace") createXML(item, "replace");
                    if (name == "add") MyXML.addAtt(tmp, "name", n);
                    else if (name == "delete" || name == "replace") MyXML.addAtt(tmp, "name", MyXML.getAtt(item, "Name"), true);
                    MyXML.addAtt(tmp, "num", MyXML.getAtt(item, "Num"), true);
                    if (name == "replace")
                    {
                        if (tmpPath == "") tmpPath = path;
                        attName = MyXML.getAtt(item, "NameReplace");
                        n = tmpPath + attName + a;
                        if (attName.StartsWith("v3#")) n = attName;
                        MyXML.addAtt(tmp, "nameReplace", n, true);
                    }
                    MyXML.addAtt(tmp, "r", MyXML.getAtt(item, "r"), true);
                    if (name != "el") b.Add(tmp);
                }
            }
            public void addDoc()
            {
                doc = I.app.Documents.Add(DocumentTypeEnum.kAssemblyDocumentObject, CreateVisible: false) as AssemblyDocument;
            }
            public void openDoc()
            {
                doc = I.open(path + nameForOpen(), drw: false) as AssemblyDocument;
            }
            public void save()
            {
                bool open = false; string sn = path + nameForSave();

                if (System.IO.File.Exists(sn))
                {
                    doc = I.open(sn, drw: false) as AssemblyDocument;
                }
                else
                {
                    if (doc == null) openDoc();
                    doc.SaveAs(sn, false);
                    doc = I.open(sn) as AssemblyDocument;
                }
                adr();
                u.addProp(doc as Document, "DecNumber", dec + "-" + num.ToString("00"));
                u.addProp(doc as Document, "Description", d);
                if (comment != null) u.addProp(doc as Document, "Comments", comment);
                doc.Save2(true);
                doc.Close();
            }
        }

        public void CreateDrawing(XElement el, XElement par, string typ, string f)
        {
            string fileName = "";
            var fel = XMLDoc.find("Name", f, "Drawing", par);
            if (u.isNull(fel)) return;
            var vals = XMLDoc.getAttributeValue(fel, "params");
            if (u.isNull(vals)) return;
            string cnt = MyXML.getAtt(el, "Count");
            fileName = ut.nameForSave(el, typ, cnt);
            string sp = XMLDoc.getSubPath(el, "subPath");
            string no = XMLDoc.getAttributeValue(el, "noConstr");
            fileName = m_Parts.path + sp + fileName;
            string trfileName = trimEnd(fileName);
            string rname = fileName;
            if (!u.isNull(no)) rname += "?" + no;
            if (System.IO.File.Exists(trfileName) && Docs.chk[5])
            {
                doc = I.open(trfileName, false, false);
                var spl = u.getSpl(vals, ';').ToList();
                var drwName = file.p(trfileName) + file.name(trfileName) + ".idw";
                if (file.check(drwName)) return;
                if (!u.isNull(doc)) drawing(doc, spl);
            }
        }

        public Docs create(XElement el, string typ, int level, int maxLevel = 5, XMLDoc xdoc = null)
        {
            bool first = true;
            Docs docs = new Docs();
            int n = 0;
            string prev = null;
            foreach (var item in el.Elements())
            {
                if (item.HasElements && item.Name == "Assembly" && level < maxLevel)
                {
                    docs = create(item, typ, ++level, xdoc: xdoc);
                    // u.action<string>(docs2.docs, a => docs.docs.Add(a));
                }
                string drw = XMLDoc.getAttributeValue(item, "Draw");
                if (maxLevel > 1 && !u.isNull(drw))
                {
                    CreateDrawing(item, xdoc.Doc.Root, typ, drw);
                }
                if (item.Name == "Part")
                {
                    docs.docs.Add(createPart(item, typ, prev));
                    prev = docs.docs.LastOrDefault();
                }
                else if (item.Name == "Assembly")
                {
                    u.action<string>(docs.adds, a => docs.docs.Add(a));
                    string tmp = createAssembly(item, typ, docs);
                    if (first) { docs.docs.Clear(); u.action<string>(docs.adds, a => docs.docs.Add(a)); first = false; }
                    docs.docs.Add(tmp);
                }
                else if (item.Name == "Add") { add(item, ref docs); u.action<string>(docs.adds, a => docs.docs.Add(a)); }
                else if (item.Name == "Isp" && Docs.chk[0]) { n++; isp(item, n, typ, item.Parent, ref docs); }
            }
            if (Docs.chk[0])
            {
                createPerfs(docs);
            }
            return docs;
        }

        public void isp(XElement el, int n, string typ, XElement p, ref Docs docs)
        {
            Perf pe = new Perf(n, typ, p.Attribute("DecNumber").Value, p.Attribute("Name").Value, el, m_Parts.path);
            docs.perfs.Add(pe);
        }

        public string add(XElement el, ref Docs docs)
        {
            string path = null;
            if (el.Attribute("Path") != null)
            {
                path = el.Attribute("Path").Value;
                if (path.StartsWith(".\\")) path = path.Replace(".\\", I.curProjPath() + "\\");
            }
            else path = m_Parts.path;
            if (!path.EndsWith("\\")) path += "\\";
            foreach (var item in el.Elements())
            {
                string val = item.Attribute("Name").Value;
                string no = XMLDoc.getAttributeValue(item, "noConstr");
                if (u.isNull(no)) no = "";
                else no = "?" + no;
                if (item.Name == "Lib")
                {
                    path = "";
                    if (!val.StartsWith("v3#"))
                    {
                        XMLDoc lib = new XMLDoc(I.p() + @"\ContentCenter.xml", "Content");
                        XElement tmp = lib.getXElement(val, "Description", "TableRow");
                        val = tmp.Attribute("Id").Value;
                    }
                }
                if (val.StartsWith("-"))
                {
                    val = u.findRegexFN(path, val, true);
                }
                if (item.Attribute("Count") != null)
                {
                    docs.adds.Add(path + val + u.nameSuff(item) + '^' + item.Attribute("Count").Value + no);
                }
                else docs.adds.Add(path + val + u.nameSuff(item) + no);
            }
            return null;
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

        public string createPart(XElement el, string typ, string bn = null)
        {
            string fileName = "";
            //             if (el.Attribute("Path") != null)
            //             {
            //                 m_Parts.path = el.Attribute("Path").Value;
            //                 fileName = el.Attribute("Name").Value + ".ipt";
            //                 return m_Parts.path + fileName;
            //             }
            //             else 
            string cnt = MyXML.getAtt(el, "Count");
            fileName = ut.nameForSave(el, typ, cnt);
            string sp = XMLDoc.getSubPath(el, "subPath");
            string no = XMLDoc.getAttributeValue(el, "noConstr");
            fileName = m_Parts.path + sp + fileName;
            string trfileName = trimEnd(fileName);
            string rname = fileName;
            if (!u.isNull(no)) rname += "?" + no;
            PartDocument doc = null; bool open = false;
            if (System.IO.File.Exists(trfileName) && Docs.chk[4])
            {
                doc = m_Parts.openPrtDoc(trfileName);
                open = true;
            }
            else if (Docs.chk[2] && !System.IO.File.Exists(trfileName)) doc = m_Parts.addPrtDoc();
            else return rname;
            if (!Docs.chk[4] && !Docs.chk[2]) return rname;
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
                if (pln != DerivedPartMirrorPlaneEnum.kDerivedPartNoMirrorPlane && bn != null)
                {
                    bp = file.nameWithExt(bn);
                }
                string str = "";
                if (el.Attribute("P") != null)
                    str = el.Attribute("P").Value;
                doc = !open ? m_Parts.derivedDoc(m_Parts.path + sp + bp, skeches, null, namePar, pln) :
                    m_Parts.derivedDoc(m_Parts.path + sp + bp, skeches, doc, str);
                if (doc.ComponentDefinition is SheetMetalComponentDefinition && el.Attribute("SheetMetalStyle") != null)
                {
                    SheetMetalComponentDefinition smcd = doc.ComponentDefinition as SheetMetalComponentDefinition;
                    Point pt = I.tg.CreatePoint(0, 0, ut.convToDouble(smcd.Thickness.Value.ToString()));
                    SheetMetalStyle sms = smcd.SheetMetalStyles.OfType<SheetMetalStyle>().FirstOrDefault(e => e.Name == el.Attribute("SheetMetalStyle").Value);
                    if (sms != null) sms.Activate();
                }
                foreach (PlanarSketch ps in doc.ComponentDefinition.Sketches)
                {
                    ps.Visible = false;
                }
            }
            if (el.Attribute("Props") != null) u.addProps(doc as Document, el.Attribute("Props").Value.ToString());
            if (open) doc.Save2(true);
            else
            {
                doc.SaveAs(trfileName, false);
            }

            return rname;
            //doc.ReleaseReference();
            //doc.Close();
        }
        public string trimEnd(string fn, string sym = "^")
        {
            if (fn.IndexOf(sym) == -1) return fn;
            fn = fn.Remove(fn.IndexOf(sym));
            return fn;
        }

        public bool createPerfs(Docs docs)
        {
            if (docs.perfs != null && docs.perfs.Count > 0)
            {
                foreach (var item in docs.perfs)
                {
                    item.save();
                }
            }
            docs.perfs.Clear();
            return true;
        }

        public string createAssembly(XElement el, string typ, Docs docs)
        {
            bool open = false;
            AssemblyDocument doc = null;
            string fileName = "";
            fileName = ut.nameForSave(el, typ);
            if (el.Attribute("Path") != null)
                m_Parts.path = el.Attribute("Path").Value;
            string sp = XMLDoc.getSubPath(el, "subPath"), fn = m_Parts.path + sp + fileName;

            string no = XMLDoc.getAttributeValue(el, "noConstr");
            string rname = fn;
            if (!u.isNull(no)) rname += "?" + no;

            if (System.IO.File.Exists(fn) && Docs.chk[3])
            {
                doc = m_Parts.openAsmDoc(fn);
                open = true;
            }
            else if (!System.IO.File.Exists(fn) && Docs.chk[1]) doc = m_Parts.addAsmDoc();
            else return rname;
            string name = el.Attribute("Name").Value, decNumber = el.Attribute("DecNumber").Value;
            ut.partNumber(doc as Document, decNumber, typ);
            ut.addProp(doc as Document, "Description", name);

            //             foreach (var item in docs.docs)
            //             {
            //                 if(!ut.occsContaints(doc.ComponentDefinition.Occurrences, item, 1))
            //                 doc.ComponentDefinition.Occurrences.AddUsingiMates(item);       
            //             }
            if (!Docs.names.Contains(name))
            {
                addDocs(doc, docs.docs);
                //                 if (docs.adds != null && docs.adds.Count > 0)
                //                 {
                //                     addDocs(doc, docs.adds);
                //                 }
            }
            Docs.names.Add(name);
            //string fileName = ut.nameForSave(doc as Document);
            if (el.Attribute("Props") != null) u.addProps(doc as Document, el.Attribute("Props").Value.ToString());
            if (open) doc.Save2(true);
            else doc.SaveAs(fn, false);
            doc.ReleaseReference();
            doc.Close();

            return rname;
        }

        public void addDocs(AssemblyDocument doc, HashSet<string> docs)
        {
            foreach (var e in docs)
            {
                string item = u.getElem(e, 0, '?'), no = u.getElem(e, 1, '?');
                int cnt = 1; string n = item;
                if (item.IndexOf("^") != -1)
                {
                    var spl = item.Split('^');
                    cnt = int.Parse(spl[1]);
                    n = spl[0];
                }
                if (n.StartsWith("v3#"))
                {
                    n = ContentOp.memberForPlace(n);
                }
                var m = I.tg.CreateMatrix();
                for (int i = 0; i < cnt; i++)
                {
                    if (!ut.occsContaints(doc.ComponentDefinition.Occurrences, n, cnt))
                        if (u.isNull(no))
                            doc.ComponentDefinition.Occurrences.AddUsingiMates(n);
                        else doc.ComponentDefinition.Occurrences.Add(n, m);
                }
            }
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

        private void ручнаяГибкаToolStripMenuItem_Click_1(object sender, EventArgs e)
        {
            HandBend hb = new HandBend(I.aDoc());
            this.Close();
        }

        private void заменитьToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void спецToolStripMenuItem_Click_1(object sender, EventArgs e)
        {
            string fn = u.WOFD(I.p() + @"\", "XML (*.xml)|*.xml", new string[] { file.p(I.aDoc().FullFileName) });
            if (fn == "") { Macros.StandardAddInServer.xml = null; return; }
            Macros.StandardAddInServer.xml = new MyXML(fn);
        }

        private void добавитьИсполнениеToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Document doc = Macros.StandardAddInServer.m_inventorApplication.ActiveDocument;
            List<string> asms = TableInv.getAsms(doc.FullFileName); asms.Sort();
            int isp = 1;
            string name = asms[asms.Count - 1]; int l = name.Length;
            if (name[l - 7] == '^')
            {
                isp = int.Parse(name.Substring(l - 6, 2)) + 1;
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
            insPt.X += 150; insPt.Y -= offsetY * 2;
            System.Windows.Forms.ComboBox cb = mycb.addComboBox("", insPt, 200, 15, new string[] { "1 кВт", "3 кВт", "4 кВт", "6 кВт", "8 кВт", "10 кВт", "12 кВт", "18 кВт", "24 кВт",
                "27 кВт", "36 кВт", "54 кВт"});
            f.Controls.Add(cb);
            insPt.Y += offsetY;
            System.Windows.Forms.TextBox txt2 = mytxt1.addTextBox(isp.ToString("00"), insPt, 200, 15);
            insPt.Y += offsetY;
            cb = mycb.addComboBox("", insPt, 200, 15, new string[] { "A", "E", "W" });
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
            Regex regex = new Regex(@"\b([0-9\.\,]*)"), regex1 = new Regex(@"(КЭВ-)(\d*)(\w)(\d)(\d)(\d*)(.*)");
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
                    decu = match1.Groups[4].Value + match1.Groups[5].Value + match1.Groups[6].Value;
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
            desc = perf(desc, pn, ref type, ref decnumber, ref note, EWA, ref count);
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
            Document docAdd = Macros.StandardAddInServer.m_inventorApplication.Documents.OpenWithOptions(n, nvmOptions, false);
            if (docAdd.DocumentType == DocumentTypeEnum.kAssemblyDocumentObject)
            {
                AssemblyComponentDefinition compDef = (docAdd as AssemblyDocument).ComponentDefinition;
                compDef.RepresentationsManager.DesignViewRepresentations[2].Activate();
            }
            ut.addProp(docAdd, "Description", desc); ut.addProp(docAdd, "Type", type);
            if (count != "")
                ut.addProp(docAdd, "DecNumber", decnumber + "-" + count);
            else ut.addProp(docAdd, "DecNumber", decnumber);
            if (note != "")
                ut.addProp(docAdd, "Comments", note + " кВт"); 
            ut.addProp(docAdd, "Part Number", "=<Type>.<DecNumber>");
            if (EWA == "W" && System.IO.File.Exists(n = System.IO.Path.Combine(System.IO.Path.GetDirectoryName(n) + "\\W.xml")))
            {
                Parts.removeReplace(docAdd, n);
            }
            else if (EWA == "A" && System.IO.File.Exists(n = System.IO.Path.Combine(System.IO.Path.GetDirectoryName(n) + "\\A.xml")))
            {
                Parts.removeReplace(docAdd, n);
            }
            docAdd.Save2();
            docAdd.ReleaseReference();
            Macros.StandardAddInServer.m_inventorApplication.SilentOperation = false;
            f.Close();
        }

        private void обновитьToolStripMenuItem_Click_1(object sender, EventArgs e)
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
            SelectionFilterEnum[] enu = new SelectionFilterEnum[] { SelectionFilterEnum.kAllCircularEntities };
            foreach (ComponentOccurrence occ in asm.ComponentDefinition.Occurrences)
            {
                if (occ.iMateDefinitions.Count != 0)
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

                addwithDrw(rd.ReferencedDocument as Document, ref names);
            }

            addwithDrw(doc, ref names);
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

        private void обновитьПозицииToolStripMenuItem_Click_1(object sender, EventArgs e)
        {
            DrawingDocument m_Drw = Macros.StandardAddInServer.m_inventorApplication.ActiveDocument as DrawingDocument;
            AssemblyDocument m_AsmDoc = InvDoc.u.referendedDoc(m_Drw as Document) as AssemblyDocument;
            foreach (Balloon bal in m_Drw.ActiveSheet.Balloons)
            {
                if (bal.BalloonValueSets.Count == 1)
                {
                    Document rdoc = (Document)bal.ParentView.ReferencedDocumentDescriptor.ReferencedDocument;
                    if (rdoc.DocumentType == DocumentTypeEnum.kAssemblyDocumentObject)
                    {
                        Macros.StandardAddInServer.addBalloonValueSet(bal, (AssemblyDocument)rdoc);
                    }
                }

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
                dwg.SaveAsInventorDWG(dwg.FullDocumentName.Replace(".idw", ".dwg"), true);
            }
        }

        private void получитьТекстToolStripMenuItem_Click_1(object sender, EventArgs e)
        {
            InvDoc.InvDocument<AssemblyDocument> doc = new InvDoc.InvDocument<AssemblyDocument>(Macros.StandardAddInServer.m_inventorApplication.ActiveDocument as AssemblyDocument);
            BOMView view = doc.getBOMView(false, false, "Только детали", true, false);
            string n = u.WOFD(doc.path);
            XMLDoc xml = (n == "") ? new XMLDoc(doc.path + "English.xml", "head") : new XMLDoc(n, "head");
            InvDoc.GetDocs docs = new InvDoc.GetDocs();
            XElement root = xml.Doc.Root;
            foreach (Document item in docs.docs)
            {
                Property p = u.getProp(item, "Description");
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

        private void обновитьСтилиToolStripMenuItem_Click_1(object sender, EventArgs e)
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

        private void перевестиТекстToolStripMenuItem_Click_1(object sender, EventArgs e)
        {
            string n = u.OFD(Macros.StandardAddInServer.m_inventorApplication.ActiveDocument.path(), "xml|*.xml");
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
            MyXML xml = new MyXML(I.p() + @"\ContentCenterPaths.xml");
            List<string> lst = new List<string>();
            foreach (var item in xml.elem.Elements())
            {
                lst.Add(item.Value);
            }
            ut.action<FeatureBasedOccurrencePattern>(acd.OccurrencePatterns, p => deletePattern(p, lst));
            ut.action<RectangularOccurrencePattern>(acd.OccurrencePatterns, p => deletePattern(p, lst));
            ut.action<ComponentOccurrence>(acd.Occurrences, oc => oc.Delete(), f => check(f, lst));
        }

        static public bool check(ComponentOccurrence co, List<string> lst)
        {
            bool r = false;
            foreach (var item in lst)
            {
                if (co.ReferencedDocumentDescriptor.FullDocumentName.IndexOf(item) != -1) return !r;
            }
            return r;
            // return co.ReferencedDocumentDescriptor.FullDocumentName.IndexOf(@"\Content Center Files\") != -1 && ut.checkSquareLenght(co.RangeBox, 4); /*ut.getLenght(co.RangeBox) < 2.5;*/
        }

        static public void delete<T>(T f)
        {
            if (f == null) return;
            InvDoc.Reflect.runMethod<T, object>(f, "Delete", null);
        }

        static public void deletePattern(FeatureBasedOccurrencePattern pat, List<string> lst)
        {
            ObjectCollection colNew = I.objs.CreateObjectCollection();
            foreach (ComponentOccurrence co in pat.ParentComponents)
            {
                string name = ut.getName(co.Name, ':', 0);
                if (check(co, lst) && co.Name.IndexOf(name) == -1) colNew.Add(co);
            }
            if (colNew.Count != 0) pat.ParentComponents = colNew;
            else pat.Delete();
        }

        static public void deletePattern(RectangularOccurrencePattern pat, List<string> lst)
        {
            ObjectCollection colNew = I.objs.CreateObjectCollection();
            foreach (ComponentOccurrence co in pat.ParentComponents)
            {
                string name = ut.getName(co.Name, ':', 0);
                if (check(co, lst) && co.Name.IndexOf(name) == -1) colNew.Add(co);
            }
            if (colNew.Count == 1) pat.ParentComponents = colNew;
            else del.Add(pat);
        }

        private void переименоватьСборкуToolStripMenuItem_Click_1(object sender, EventArgs e)
        {
            renameAsm(I.app.ActiveDocument);
        }

        private void заменитьИмяКПToolStripMenuItem_Click(object sender, EventArgs e)
        {
            XMLDoc xdoc = new XMLDoc(I.p() + @"\rimate.xml", "head");
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
                ut.action<iMateDefinition>(compDef.iMateDefinitions, f => check(xdoc, f));
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
            xdoc.setRoot();
            XElement el = xdoc.find("old", i.Name);
            if (el == null)
                xdoc.addXElement("imate", new Dictionary<string, string>() { { "old", i.Name }, { "new", "" } });
            else if (el.Attribute("new") != null && el.Attribute("new").Value != "")
                i.Name = el.Attribute("new").Value;

        }

        public static void asmImate()
        {
            ut.action<Document>(I.app.Documents.VisibleDocuments, f => asmFindIMate(f));
            I.save(false);
            I.clearDocs();
        }

        public static void asmFindIMate(Document doc)
        {
            AssemblyComponentDefinition acompDef = I.getACD(doc);
            if (acompDef != null)
            {
                ut.action<iMateResult>(acompDef.iMateResults, f => getImate(f));
            }
        }
        public static void getImate(iMateResult r)
        {
            if (!r.IsComposite) return;
            Document d1 = getDoc(r.Constraints[1].OccurrenceOne),
                d2 = getDoc(r.Constraints[1].OccurrenceTwo);
            string v1 = ut.getPropValue(d1, "Description"),
                v2 = ut.getPropValue(d2, "Description");
            iMateDefinition one; object two;
            r.GetInputs(out one, out two);
            iMateDefinition twoImateDef = two as iMateDefinition;
            replaceImate(one, v2, false);
            replaceImate(twoImateDef, v2, true);
            if (!isContent(d1, "content")) I.addDoc(d1);
            if (!isContent(d2, "content")) I.addDoc(d2);
        }
        public static bool isContent(Document doc, string s)
        {
            return doc.FullDocumentName.ToLower().IndexOf(s) != -1 ? true : false;
        }
        public static Document getDoc(ComponentOccurrence occ)
        {
            return occ.ReferencedDocumentDescriptor.ReferencedDocument as Document;
        }
        public static void replaceImate(iMateDefinition i, string val, bool match)
        {
            if (match)
            {
                if (i.MatchList == null)
                    i.MatchList = new string[] { val };
                else
                {
                    string[] lst = i.MatchList as string[];
                    if (lst.Contains(val)) return;
                    lst = lst.Concat(new string[] { val }).ToArray();
                    i.MatchList = lst;
                }
            }
            else if (i.Name.StartsWith("iComp"))
            {
                i.Name = val;
            }
        }

        private void обновитьКрепежToolStripMenuItem_Click_1(object sender, EventArgs e)
        {
            ut.action<AssemblyDocument>(I.app.Documents.VisibleDocuments, d => updateFasteners(d), f => f.DocumentType == DocumentTypeEnum.kAssemblyDocumentObject);
        }

        static public void updateFasteners(AssemblyDocument asm)
        {
            //ContentOp c = new ContentOp();
            removeFasteners(asm as Document);
            addFasteners(asm, new ContentOp());
        }

        private void крепежToolStripMenuItem3_Click_1(object sender, EventArgs e)
        {
            ut.action<AssemblyDocument>(I.app.Documents.VisibleDocuments, d => addFasteners(d, new ContentOp()), f => f.DocumentType == DocumentTypeEnum.kAssemblyDocumentObject);
        }

        private void открытьПапкуToolStripMenuItem_Click_1(object sender, EventArgs e)
        {
            open();
        }
        private void open()
        {
            var doc = I.aDoc();
            var open = new OpenFile(doc);
            open.open();
            this.Close();
        }

        private void кластерToolStripMenuItem_Click_1(object sender, EventArgs e)
        {
            Cluster cl = new Cluster();
            cl.draw();
            if (cl.full)
            {

            }
            cl.txt();
        }

        private void габаритыToolStripMenuItem_Click_1(object sender, EventArgs e)
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
                sw.Write("Габариты: \t" + (v.X * 10).ToString("##.##") + " мм\n\t\t\t" + (v.Y * 10).ToString("##.##") + " мм\n\t\t\t" + (v.Z * 10).ToString("##.##") + " мм.\n");
            }
        }

        private void инициализацияToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Document doc = I.aDoc();
            string path = u.pathDoc(doc);
            init(path);
        }
        public void init(string path)
        {
            string[] files = System.IO.Directory.GetFiles(path, "*.i??", System.IO.SearchOption.TopDirectoryOnly);
            string[] ns = files.Select(f => System.IO.Path.GetFileName(f)).ToArray();
            string[] filter = new string[] { ".idw" };
            foreach (var item in files)
            {
                if (!filter.Contains(System.IO.Path.GetExtension(item))) continue;
                Document doc = I.open(item);
                foreach (DocumentDescriptor d in doc.ReferencedDocumentDescriptors)
                {
                    string n = d.FullDocumentName;
                    string p = System.IO.Path.GetDirectoryName(n);
                    n = System.IO.Path.GetFileName(n);
                    if (ns.Contains(n) && p != path)
                    {
                        d.ReferencedFileDescriptor.ReplaceReference(System.IO.Path.Combine(path, n));
                    }
                }
            }
        }

        private void шипИнверсияToolStripMenuItem_Click(object sender, EventArgs e)
        {
            spikes(false);
        }

        private void CreateComponent_KeyDown(object sender, KeyEventArgs e)
        {
            switch (e.KeyCode)
            {
                case Keys.D:
                    drawing();
                    break;
                case Keys.C:
                    create();
                    break;
                case Keys.O:
                    open();
                    break;
                case Keys.S:
                    setSupressBody();
                    break;
                case Keys.Q:
                    join();
                    break;
                case Keys.W:
                    this.Close();
                    break;
                case Keys.A:
                    addParameter();
                    break;
                case Keys.E:
                    surfaceVisible();
                    this.Close();
                    break;
                case Keys.R:
                    autoBlock();
                    this.Close();
                    break;
                case Keys.B:
                    var doc1 = I.aDoc();
                    if (doc1.DocumentType == DocumentTypeEnum.kDrawingDocumentObject)
                        breakView();
                    else if (doc1.DocumentType == DocumentTypeEnum.kPartDocumentObject)
                        setVisibleBody();
                    else if (doc1.DocumentType == DocumentTypeEnum.kAssemblyDocumentObject)
                    {
                        if (doc1.ActivatedObject is PartDocument)
                        {
                            setVisibleBody();
                        }
                    }
                    break;
                case Keys.N:
                    findBrowserNode();
                    break;
                case Keys.V:
                    visible();
                    break;
                case Keys.T:
                    insertBlock();
                    this.Close();
                    break;
                case Keys.F:
                    this.Hide();
                    ContentOp c = new ContentOp();
                    Transaction tr = Macros.StandardAddInServer.m_inventorApplication.TransactionManager.StartTransaction(Macros.StandardAddInServer.m_inventorApplication.ActiveDocument, "Крепеж");
                    addFasteners(I.app.ActiveDocument as AssemblyDocument, c);
                    tr.End();
                    break;
                case Keys.I:
                    this.Hide();
                    smartInsert();
                    this.Close();
                    break;
                case Keys.Z:
                    replaceDraw();
                    break;
                case Keys.H:
                    changeTypeBody();
                    break;
                case Keys.X:
                    var doc = I.aDoc();
                    if (doc.DocumentType == DocumentTypeEnum.kAssemblyDocumentObject)
                        removeFasteners(doc);
                    break;
                case Keys.M:
                    mark();
                    this.Close();
                    break;
                default:
                    break;
            }
        }

        private void экспортВSATToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Document doc = Macros.StandardAddInServer.m_inventorApplication.ActiveDocument;
            string path = file.combine(file.p(doc.FullFileName), "adsk");
            MyXML xml = new MyXML("SATInterface.xml");
            MyForm F = new MyForm(xml, "Создать");
            F.f.ShowDialog();
            /*            return;*/
            Docs docs = new Docs();
            if (F.chks != null)
            {
                Docs.chk.Clear();
                for (int i = 0; i < F.chks.Count(); i++)
                {
                    Docs.chk.Add(F.chks[i].Checked);
                }
            }
            if (Docs.chk[0])
            {
                Asms asms = new Asms(I.aDoc());
                asms.get();
                asms.create();
            }
            if (Docs.chk[1])
            {
                IEnumerable<Document> asms = I.getFiles<Document>(path, ".iam");
                foreach (var item in asms)
                {
                    Sat s = new Sat(item);
                    s.run();
                }
            }
            if (Docs.chk[2])
            {
                Bim b = new Bim(path, "_sat.ipt");
                b.get();
                b.save();
            }
            if (Docs.chk[3])
            {
                Bim b = new Bim(path, ".iam");
                b.get();
                b.saveSAT();
            }
        }

        private void кПToolStripMenuItem_Click(object sender, EventArgs e)
        {
            asmImate();
        }

        private void создатьЭскизыToolStripMenuItem_Click(object sender, EventArgs e)
        {
            sketch();
        }

        private void выравниваниеОсейToolStripMenuItem_Click(object sender, EventArgs e)
        {
            alignAxis();
        }

        public void alignAxis()
        {
            var acd = I.getACD(I.aDoc());
            ComponentOccurrence boc = u.get<ComponentOccurrence>(acd.Occurrences, f => f.Grounded);
            if (boc == null) return;
            boc.Grounded = false;
            alignAxis(boc, acd);
        }

        public void alignAxis(ComponentOccurrence oc, AssemblyComponentDefinition b)
        {
            int i = 0; WorkAxes wa = null;
            AssemblyComponentDefinition adef = oc.Definition as AssemblyComponentDefinition;
            if (adef != null) wa = adef.WorkAxes;
            PartComponentDefinition pdef = oc.Definition as PartComponentDefinition;
            if (pdef != null) wa = pdef.WorkAxes;
            if (wa == null) return;
            foreach (var item in wa)
            {
                i++;
                if (i > 2) break;
                object r;
                oc.CreateGeometryProxy(item, out r);
                if (r == null) continue;
                b.Constraints.AddMateConstraint(b.WorkAxes[i], r, 0);
            }
        }

        private void перенестиВЦентрToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Asms asm = new Asms(I.aDoc());
            asm.translateToCenter(I.aDoc().FullDocumentName);
        }

        private void совместитьКПToolStripMenuItem_Click(object sender, EventArgs e)
        {
            join();
        }

        public void join()
        {
            u.joinIMates(I.aDoc());
        }

        private void копироватьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            string p = I.curProjPath();
            var sd = file.subDir(p, new string[] { "OldVersions" }, System.IO.SearchOption.AllDirectories);
            XMLDoc xml = new XMLDoc(I.p() + @"\ChangeContentCenter.xml", "head");
            XMLDoc.removeEl(xml.El, "restore");
            XMLDoc.removeEl(xml.El, "change");
            ContentOp.addColumn(xml);
        }

        private void восстановитьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            string p = I.curProjPath();
            var sd = file.subDir(p, new string[] { "OldVersions" }, System.IO.SearchOption.AllDirectories);
            XMLDoc xml = new XMLDoc(I.p() + @"\ChangeContentCenter.xml", "head");
            XMLDoc.removeEl(xml.El, "copy");
            ContentOp.addColumn(xml);
        }

        private void копироватьВПапкуToolStripMenuItem_Click(object sender, EventArgs e)
        {
            string p = I.curProjPath();
            XMLDoc xml = new XMLDoc(p + "\\Поиск.xml", "head");
            xml.El = xml.El.Element("Find");
            string pat = xml.getAttributeValue("Name"), fold = "\\" + xml.getAttributeValue("folder") + "\\", fltr = xml.getAttributeValue("filter");
            file.createPath(p + fold);
            Regex regex = new Regex(pat);
            var files = file.getFiles(p, ".iam", System.IO.SearchOption.AllDirectories);
            if (!u.isNull(fltr))
            {
                var spl = u.getSpl(fltr, ';');
                u.action<string>(spl, a => files = files.Where(f => f.IndexOf(a) == -1));
            }
            //files = files.Where(f => f.IndexOf("OldVersions") == -1);
            files = files.Where(f => regex.IsMatch(file.name(f)));
            I.silent(true);
            foreach (var item in files)
            {
                string n = p + fold + file.nameWithExt(item);
                if (file.check(n)) continue;
                Document doc = I.open(item, false, false);
                if (u.isNull(doc)) continue;
                doc.SaveAs(n, true);
            }
            files = file.getFiles(p + fold, ".iam");
            xml.setRoot();
            foreach (var item in xml.El.Elements("Add"))
            {
                pat = XMLDoc.getAttributeValue(item, "Name");
                string suf = XMLDoc.getAttributeValue(item, "suf"), fn = XMLDoc.getAttributeValue(item, "file");
                Document doc = I.open(p + "\\" + fn);
                if (u.isNull(doc)) continue;
                addFile(files, pat, p + fold, suf, doc);
            }
            I.silent(false);
        }

        public void addFile(IEnumerable<string> files, string pat, string path, string suf, Document doc)
        {
            Regex regex = new Regex(pat);
            files = files.Where(f => regex.IsMatch(file.name(f)));
            foreach (var item in files)
            {
                string n = path + file.name(item) + suf + ".ipt";
                if (file.check(n)) continue;
                doc.SaveAs(n, true);
            }
        }

        private void dPDFToolStripMenuItem_Click(object sender, EventArgs e)
        {
            pdfFromXML();
        }

        private void подпозицииToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Document doc = I.aDoc();
            if (doc.DocumentType == DocumentTypeEnum.kDrawingDocumentObject)
            {
                SelectSet ss = I.getSS(doc);
                foreach (Balloon occ in ss)
                {
                    BalloonValueSets sets = occ.BalloonValueSets;
                    for (int i = sets.Count; i > 1; i--)
                    {
                        BalloonValueSet val = sets[i];
                        val.Delete();
                    }
                }
            }
        }

        private void ширинаНадписиToolStripMenuItem_Click(object sender, EventArgs e)
        {
            MyForm F = new MyForm("EditInterface.xml", "Ширина");
            //F.bnts[0].Click += fillet_Click;
            F.f.ShowDialog();
            I.screenSilent(true);
            foreach (Document doc in I.app.Documents.VisibleDocuments)
            {
                if (doc.DocumentType == DocumentTypeEnum.kDrawingDocumentObject)
                {
                    changeWidth(doc as DrawingDocument, F.cbs[0].Text, ut.convToDouble(F.cbs[1].Text));
                }
            }
            //F.f.Close();
            I.screenSilent(false);
            this.Close();
        }

        private void changeWidth(DrawingDocument doc, string name, double val)
        {
            foreach (TitleBlockDefinition td in doc.TitleBlockDefinitions)
            {
                DrawingSketch ds = null;
                td.Edit(out ds);
                //ds.Edit();
                foreach (Inventor.TextBox tb in ds.TextBoxes)
                {
                    if (tb.Text.IndexOf(name) != -1)
                    {
                        tb.WidthScale = val / 100;
                    }
                }
                //ds.ExitEdit();
                td.ExitEdit();
            }
        }

        private void позицииПоШаблонуToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Documents docs = Macros.StandardAddInServer.m_inventorApplication.Documents;
            DrawingDocument tmpl = I.aDoc() as DrawingDocument;
            Sheet sh = null;
            if (tmpl == null) return;
            string vName = "VIEW2";
            //SelectSet ss = I.getSS(tmpl as Document);
            //if (ss == null) return;
            foreach (var drw in tmpl.ActiveSheet.DrawingViews)
            {
                if (!(drw is DrawingView)) continue;
                vName = (drw as DrawingView).Name;
                //I.silent(true);
                foreach (DrawingDocument item in docs.VisibleDocuments)
                {
                    if (item.Equals(tmpl)) continue;
                    item.Activate();
                    foreach (Balloon b in tmpl.ActiveSheet.Balloons)
                    {
                        if (getName(b.ParentView.Name) != getName(vName)) continue;
                        List<Point2d> pts = u.getPts(b);
                        GeometryIntent i1 = b.Leader.AllLeafNodes[1].AttachedEntity;
                        if (i1 == null) continue;
                        DrawingCurve dc = i1.Geometry as DrawingCurve;
                        Balloon bal = u.addBallon(item.Sheets[1], pts, getName(vName), dc.CurveType);
                        //                         if (bal != null)
                        //                             Macros.StandardAddInServer.addBalloonValueSet(bal, item.ReferencedDocuments[1] as AssemblyDocument);
                    }
                }
            }
            //I.silent(false);
        }

        private void видыПоШаблонуToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Documents docs = Macros.StandardAddInServer.m_inventorApplication.Documents;
            templateDV(docs);
        }

        private void подготовитьШаблонToolStripMenuItem_Click(object sender, EventArgs e)
        {
            List<DrawingDocument> drws = new List<DrawingDocument>();
            foreach (DrawingDocument item in I.app.Documents.VisibleDocuments.Cast<DrawingDocument>())
            {
                drws.Add(item);
            }
            breakOp(drws);
            //breakOp();
        }

        private void комплектНесколькихФайловToolStripMenuItem_Click(object sender, EventArgs e)
        {
            pack(I.app.Documents.VisibleDocuments);
        }

        public void templDim(DrawingDocument drw, DrawingDocument tmplDrw)
        {
            //drw.Activate();
            Sheet sh = drw.Sheets[1];
            drw.StylesManager.ActiveStandardStyle.ActiveObjectDefaults.LinearDimensionStyle.LinearPrecision = LinearPrecisionEnum.kZeroDecimalPlaceLinearPrecision;
            foreach (GeneralDimension dim in tmplDrw.ActiveSheet.DrawingDimensions.GeneralDimensions)
            {
                if (!dim.Attached) continue;
                LinearGeneralDimension ldim = dim as LinearGeneralDimension;
                templDim(ldim, sh);
            }
        }

        public void templDim(LinearGeneralDimension ldim, Sheet sh)
        {
            Point2d sp = ldim.IntentOne.PointOnSheet, ep = ldim.IntentTwo.PointOnSheet, cen = ldim.Text.Origin;
            GeometryIntent i1 = ldim.IntentOne, i2 = ldim.IntentTwo;
            DrawingCurve st = i1.Geometry as DrawingCurve, et = i2.Geometry as DrawingCurve;
            //if (st.ProjectedCurveType != Curve2dTypeEnum.kLineSegmentCurve2d || et.ProjectedCurveType != Curve2dTypeEnum.kLineSegmentCurve2d) return;
            //             Vector2d sn = u.getNormal(st), en = u.getNormal(et);
            //             if (sn == null || en == null) return;
            DrawingCurve sc = u.findAtPoint(sh, sp, 0.2, (ldim.IntentOne.Geometry as DrawingCurve).CurveType);
            DrawingCurve ec = u.findAtPoint(sh, ep, 0.2, (ldim.IntentTwo.Geometry as DrawingCurve).CurveType);
            //             DrawingCurve sc = u.findAtPoint(sh, sp, dc => dc.CurveType == st.CurveType, 0.1);
            //             DrawingCurve ec = u.findAtPoint(sh, ep, dc => dc.CurveType == et.CurveType, 0.1);
            if (sc == null || ec == null) return;
            i1 = sh.CreateGeometryIntent(sc, u.getNearPoint(sc, sp));
            i2 = sh.CreateGeometryIntent(ec, u.getNearPoint(ec, ep));
            DimensionTypeEnum dtype = ldim.DimensionType;
            LineSegment2d dl = ldim.DimensionLine as LineSegment2d;
            if (u.eq(dl.Direction.X, 0)) dtype = DimensionTypeEnum.kVerticalDimensionType;
            if (u.eq(dl.Direction.Y, 0)) dtype = DimensionTypeEnum.kHorizontalDimensionType;
            sh.DrawingDimensions.GeneralDimensions.AddLinear(cen, i1, i2, dtype, ldim.ArrowheadsInside, ldim.Style, ldim.Layer);
        }

        private void размерыПоШаблонуToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Documents docs = Macros.StandardAddInServer.m_inventorApplication.Documents;
            DrawingDocument tmpl = I.aDoc() as DrawingDocument;
            foreach (DrawingDocument item in docs.VisibleDocuments)
            {
                if (item.Equals(tmpl)) continue;
                templDim(item, tmpl);
            }
        }

        public Dictionary<Centermark, Centermark> templCenters(DrawingDocument drw, DrawingDocument tmplDrw, ref Dictionary<Centermark, Centermark> dic)
        {
            //drw.Activate();
            Sheet sh = drw.Sheets[1];
            foreach (Centermark cen in tmplDrw.ActiveSheet.Centermarks)
            {
                Point2d pt = cen.Position;
                DrawingCurve dc = u.findAtPoint(sh, pt, 0.2, CurveTypeEnum.kCircleCurve);
                if (dc == null) continue;
                GeometryIntent i = sh.CreateGeometryIntent(dc, pt);
                if (!dic.ContainsKey(cen))
                    dic.Add(cen, sh.Centermarks.Add(i, cen.ExtensionLinesVisible));
            }
            return dic;
        }

        public void templCenterLines(DrawingDocument drw, DrawingDocument tmplDrw, Dictionary<Centermark, Centermark> dic)
        {
            //drw.Activate();
            Sheet sh = drw.Sheets[1];
            foreach (Centerline cen in tmplDrw.ActiveSheet.Centerlines)
            {
                ObjectCollection centres = I.COC();
                foreach (Centermark item in cen.FitPoints)
                {
                    Centermark c = dic[item];
                    if (c != null) centres.Add(c);
                }
                sh.Centerlines.Add(centres);
            }
        }

        private void осевыеПоШаблонуToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Documents docs = Macros.StandardAddInServer.m_inventorApplication.Documents;
            DrawingDocument tmpl = I.aDoc() as DrawingDocument;
            Dictionary<Centermark, Centermark> dic = new Dictionary<Centermark, Centermark>();
            foreach (DrawingDocument item in docs.VisibleDocuments)
            {
                if (item.Equals(tmpl)) continue;
                templCenters(item, tmpl, ref dic);
                templCenterLines(item, tmpl, dic);
            }
        }

        private void создатьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Hide();
            m_Parts.create();
            this.Show();
        }

        private void видимостьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            visible();
        }

        public void visible()
        {
#if INV14
#else
    I.app._LibraryDocumentModifiable = true;
#endif


            Document doc = I.app.ActiveEditObject as Document;
            PartComponentDefinition pcd = I.getPCD(doc);
            if (pcd != null)
            {
                setVisible(pcd);
            }
            AssemblyComponentDefinition acd = I.getACD(doc);
            if (acd != null)
            {
                setVisible(acd as ComponentDefinition);
                var p = doc.BrowserPanes[2];
                visibleNode(p.TopNode);
                //AssemblyDocument asm = acd.Parent;
                //asm.ObjectVisibility.AllWorkFeatures = false;
                //asm.ObjectVisibility.Sketches = false;
                //asm.ObjectVisibility.Sketches3D = false;
            }

#if INV14
#else
        I.app._LibraryDocumentModifiable = false;
#endif

        }

        public void visible(AssemblyComponentDefinition acd)
        {
            Document doc = acd.Document as Document;
            var p = doc.BrowserPanes[2];
            visibleNode(p.TopNode);
        }
        public void visibleNode(BrowserNode node)
        {
            foreach (BrowserNode n in node.BrowserNodes)
            {
                if (n.FullPath.IndexOf("пары") != -1) continue;
                if (n.NativeObject == null) continue;
                if (n.NativeObject is ComponentOccurrence || n.NativeObject is ComponentOccurrenceProxy ||
                    n.NativeObject is SheetMetalComponentDefinition || n.NativeObject is PartComponentDefinition ||
                    n.NativeObject is DerivedPartComponentProxy || n.NativeObject is DerivedPartComponent)
                    visibleNode(n);
                if (n.NativeObject is WorkPlaneProxy)
                {
                    WorkPlaneProxy e = n.NativeObject as WorkPlaneProxy;
                    e.Visible = false;
                }
                else if (n.NativeObject is WorkAxisProxy)
                {
                    var e = n.NativeObject as WorkAxisProxy;
                    e.Visible = false;
                }
                else if (n.NativeObject is WorkPointProxy)
                {
                    var e = n.NativeObject as WorkPointProxy;
                    e.Visible = false;
                }
                else if (n.NativeObject is PlanarSketchProxy)
                {
                    var e = n.NativeObject as PlanarSketchProxy;
                    e.Visible = false;
                }
            }
        }

        public void setVisible(ComponentDefinition def)
        {
            if (def.Type == ObjectTypeEnum.kPartComponentDefinitionObject ||
                def.Type == ObjectTypeEnum.kSheetMetalComponentDefinitionObject)
            { setVisible(def as PartComponentDefinition); }
            else if (def.Type == ObjectTypeEnum.kAssemblyComponentDefinitionObject)
            {
                var acd = def as AssemblyComponentDefinition;
                setVisible(acd);
                foreach (ComponentOccurrence occ in acd.Occurrences)
                {
                    setVisible(occ.Definition);
                }
            }
        }

        public void setVisible(PartComponentDefinition pcd)
        {
            u.action<PlanarSketch>(pcd.Sketches, a => a.Visible = false);
            u.action<WorkAxis>(pcd.WorkAxes, a => a.Visible = false);
            u.action<WorkPoint>(pcd.WorkPoints, a => a.Visible = false);
            u.action<WorkPlane>(pcd.WorkPlanes, a => a.Visible = false);
            u.action<WorkSurface>(pcd.WorkSurfaces, a => a.Visible = false);
        }
        public void setVisible(AssemblyComponentDefinition acd)
        {
            u.action<PlanarSketch>(acd.Sketches, a => a.Visible = false);
            u.action<WorkAxis>(acd.WorkAxes, a => a.Visible = false);
            u.action<WorkPoint>(acd.WorkPoints, a => a.Visible = false);
            u.action<WorkPlane>(acd.WorkPlanes, a => a.Visible = false);
        }

        private void названиеКПToolStripMenuItem_Click(object sender, EventArgs e)
        {
            imateName(I.getACD(I.aDoc()));
        }

        private void связатьВPDFToolStripMenuItem_Click(object sender, EventArgs e)
        {
            string p = file.p(I.aDoc().FullDocumentName);
            var files = file.getFiles(p, ".ipt");
            files = files.Concat(file.getFiles(p, ".iam"));
            foreach (var item in files)
            {
                Document doc = I.open(item, false, false);
                checkPDF(doc);
            }
        }

        private void удалитьПробелыToolStripMenuItem_Click(object sender, EventArgs e)
        {
            InvDocument<Document> invDoc = new InvDocument<Document>(Macros.StandardAddInServer.m_inventorApplication.ActiveDocument);
            List<string> files = invDoc.openFiles("*.ipt");
            foreach (var item in files)
            {
                deleteSpace(item);
            }
        }

        private void deleteSpace(string fn)
        {
            Document doc = I.open(fn, false, false);
            Property p = u.getProp(doc, "Part Number");
            if (p.Value.ToString() == " ")
            {
                p.Value = ""; doc.Save();
                doc.Close();
            }
        }

        private void стильРазмеровСборкиToolStripMenuItem_Click(object sender, EventArgs e)
        {
            asmStyle();
        }

        public void asmStyle()
        {
            DrawingDocument doc = I.aDoc() as DrawingDocument;
            if (doc == null) return;
            doc.StylesManager.DimensionStyles[1].LinearPrecision = LinearPrecisionEnum.kZeroDecimalPlaceLinearPrecision;
        }
        public void surfaceVisible()
        {
            PartDocument doc = I.aDoc() as PartDocument;
            if (doc == null) return;
            var ss = doc.SelectSet;
            if (ss.Count == 0)
            {
                visSurf(doc); return;
            }
            var f = ss[1] as Face;
            if (f == null) return;
            var sb = f.Parent;
            if (sb != null) visSurf(doc, sb);
        }
        public void visSurf(PartDocument doc, SurfaceBody sb = null)
        {
            if (sb != null) { sb.Visible = false; return; }
            var sbs = doc.ComponentDefinition.SurfaceBodies;
            foreach (SurfaceBody item in sbs)
            {
                item.Visible = true;
            }
        }

        private void переназватьToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            renameDesc(false);
        }

        public void renameDesc(bool add)
        {
            MyXML exc = new MyXML("PathFilter.xml");
            MyXML renXML = new MyXML("RenameDesc.xml");
            string p = I.app.DesignProjectManager.ActiveDesignProject.WorkspacePath;
            List<string> lst = file.getFiles(p, ".ipt", exc.elem.Element("Filter"));
            var ie = file.getFiles(p, ".iam", exc.elem.Element("Filter"));
            lst.AddRange(ie);

            foreach (string fn in lst)
            {
                renameDesc(renXML, fn, add);
            }
            if (add) renXML.save();
        }

        public void renameDesc(MyXML xml, string n, bool add)
        {
            Document doc = I.open(n, false, false);
            string pn = xml.elem.Name.ToString();
            Property p = u.getProp(doc, pn);
            xml.forElems(a => renameDesc(p, a, add, xml));
            if (p.Dirty)
                doc.Save2(false);
        }

        public void renameDesc(Property p, XElement el, bool add, MyXML xml)
        {
            string v = p.Value.ToString();
            if (v == "") return;
            if (add)
            {
                XElement tmp = xml.find("old", v, 1);
                if (tmp != null) return;
                xml.elem = MyXML.addXElement("Value", new Dictionary<string, string>() { { "old", v }, { "new", "" } }, xml.elem);
                return;
            }
            string old = MyXML.getAtt(el, "old");
            string n = MyXML.getAtt(el, "new");
            if (n == "") return;
            if (v != n && v.ToLower().IndexOf(old.ToLower()) != -1)
            {
                p.Value = n;
            }
        }

        public void remFromXML(List<string> lst, MyXML exc, string v)
        {
            lst.Remove(v);
            foreach (var item in lst)
            {
                exc.elem.Element("Filter").Add(new XElement("Value", item));
            }
            exc.remove("Value", v);
        }

        public static HashSet<SurfaceBody> GetBodies(SelectSet ss)
        {
            HashSet<SurfaceBody> hs = new HashSet<SurfaceBody>();
            foreach (var item in ss)
            {
                var ed = item as Edge; var f = item as Face;
                if (ed == null && f == null) continue;
                if (ed != null) f = ed.Faces[1];
                hs.Add(f.SurfaceBody);
            }
            return hs;
        }

        public static HashSet<SurfaceBodyProxy> GetProxyBodies(SelectSet ss)
        {
            HashSet<SurfaceBodyProxy> hs = new HashSet<SurfaceBodyProxy>();
            foreach (var item in ss)
            {
                var ed = item as EdgeProxy; var f = item as FaceProxy;
                if (ed == null && f == null) continue;
                if (ed != null) f = ed.Faces[1] as FaceProxy;
                hs.Add(f.SurfaceBody as SurfaceBodyProxy);
            }
            return hs;
        }

        public void setSupressBody()
        {
            var doc = I.aDoc();
            var ss = I.getSS();
            if (ss.Count == 0) 
            { 
                return; 
            }
            else
            {
                var def = I.getPCD(doc);
                if (def == null) return;
                if (ss[1] is Edge || ss[1] is Face)
                {
                    var bodies = GetBodies(ss);
                    foreach (var body in bodies)
                    {
                        foreach (PartFeature pf in body.AffectedByFeatures)
                        {
                            pf.Suppressed = true;
                        }
                    }
                }
                else if (ss[1] is SurfaceBody)
                {
                    foreach (var item in ss)
                    {
                        SurfaceBody sb = item as SurfaceBody;
                        if (sb == null) continue;
                        foreach (PartFeature pf in sb.AffectedByFeatures)
                        {
                            pf.Suppressed = true;
                        }
                    }
                }
            }
        }
        public SurfaceBody findBody(PartComponentDefinition def, string f1, string f2)
        {
            foreach (SurfaceBody item in def.SurfaceBodies)
            {
                if (item.Name.IndexOf(f1) != -1 && item.Name.ToLower().IndexOf(f2) != -1) return item;
            }
            return null;
        }
        public void changeTypeBody()
        {
            var doc = I.aDoc();
            var ss = I.getSS();
            if (ss.Count == 0)
            {
                return;
            }
            else
            {
                MyForm F = new MyForm("ChangeInterface.xml", "Сменить");
                F.f.ShowDialog();
                var f2 = F.cbs[0].Text.ToLower();
                if (f2 == "type_00") f2 = "";
                var def = I.getPCD(doc);
                if (def == null) return;
                if (ss[1] is Edge || ss[1] is Face)
                {
                    var bodies = GetBodies(ss);
                    foreach (var body in bodies)
                    {
                        body.Visible = false;
                        var str = body.Name;
                        var spl = str.Split('^');
                        if (spl.Length > 1) str = spl[1];
                        spl = str.Split('$');
                        var f1 = spl[0];
                        var fb = findBody(def, f1, f2);
                        if (fb != null) fb.Visible = true;
                    }
                }
            }
        }
        public void setVisibleBody(Document doc = null)
        {
            if (doc == null)
                doc = I.aDoc();
            var ss = I.getSS();
            if (ss.Count == 0)
            {
                var def = I.getPCD(doc);
                if (def != null)
                {
                    var type = Macros.StandardAddInServer.m_Variable.get();
                    foreach (SurfaceBody item in def.SurfaceBodies)
                    {
                        if (item.Name.StartsWith("!")) item.Visible = false;
                        if (type != null && type != "" && type != "Type" && !item.Name.StartsWith(type))
                        {
                            item.Visible = false; continue; 
                        }
                        if (!item.Visible)
                            item.Visible = true;
                    }
                }
                else
                {
                    var asm = I.getACD(doc);
                    var a = doc.ActivatedObject as PartDocument;
                    if (a == null) return;
                    if (asm != null)
                    {  
                        foreach (ComponentOccurrence occ in asm.Occurrences)
                        {
                            if (occ.ReferencedDocumentDescriptor.FullDocumentName != a.FullDocumentName) continue;
                            foreach (SurfaceBodyProxy item in occ.SurfaceBodies)
                            {
                                if (!item.Visible)
                                    item.Visible = true;
                            }
                        }
                    }
                }
                
                //I.sb.Clear();
                this.Close();
                doc.Update();
                return;
            }
            if (ss[1] is Edge || ss[1] is Face)
            {
                var bodies = GetBodies(ss);
                var type = Macros.StandardAddInServer.m_Variable.get();
                foreach (SurfaceBody item in bodies)
                {
                    if (item.Name.StartsWith("!")) item.Visible = false;
                    if (type != null && type != "" && type != "Type" && !item.Name.StartsWith(type)) continue;
                    item.Visible = false;
                    I.sb.Add(item);
                }
            }
            else if (ss[1] is EdgeProxy || ss[1] is FaceProxy)
            {
                var proxy = GetProxyBodies(ss);
                foreach (SurfaceBodyProxy item in proxy)
                {
                    item.Visible = false;
                }
            }
            doc.Update();
            this.Close();
        }

        public void openFromPattern()
        {
            MyForm F = new MyForm("OpenFilterInterface.xml", "Открыть");
            string prPath = I.app.DesignProjectManager.ActiveDesignProject.WorkspacePath;
            F.cbs[2].Text = prPath;
            F.cbs[2].DropDownHeight = 1080 * 3 / 2;
            MyXML exc = new MyXML("PathFilter.xml");
            var ie = file.getDirs(prPath, "", exc.elem.Element("Filter"), true);
            F.cbs[2].Items.AddRange(ie);
            //F.bnts[0].Click += fillet_Click;
            F.f.ShowDialog();
            if (F.close) return;
            string find = "", t = "";
            find = F.cbs[0].Text;
            t = F.cbs[1].Text;
            prPath = F.cbs[2].Text;
            exc.elem.Element("Filter").Add(new XElement("Value", ".xls"));
            exc.remove("Value", "^");
            List<string> vals = new List<string>() { ".iam", ".ipt", ".idw" };
            if (t != "")
            {
                remFromXML(vals, exc, t);
            }
            //MessageBox.Show("путь: " + prPath + "\n" + decNum);
            // return;
            var spl = file.getFiles(prPath, find, exc.elem.Element("Filter"));
            foreach (var item in spl)
            {
                I.open(item, true);
            }
        }

        private void получитьНазванияToolStripMenuItem_Click(object sender, EventArgs e)
        {
            renameDesc(true);
        }

        private void свойствоПоШаблонуToolStripMenuItem_Click(object sender, EventArgs e)
        {
            projProp();
        }

        private void ЭкспортВOBJToolStripMenuItem_Click(object sender, EventArgs e)
        {
            toOBJ o = new toOBJ();
        }

        private void экспортВGLBОдноТелоToolStripMenuItem_Click(object sender, EventArgs e)
        {
            new toGLB(GlbMode.Single);
        }

        private void экспортВGLBСборкаToolStripMenuItem_Click(object sender, EventArgs e)
        {
            new toGLB(GlbMode.Assembly);
        }

        private void восстановитьToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            var dp = I.app.DesignProjectManager.ActiveDesignProject;
            string ccp = dp.ContentCenterPath;
            var fs = file.getFiles(ccp, ".ipt", System.IO.SearchOption.AllDirectories);
            fs = fs.Where(f => f.IndexOf("OldVersions") == -1);
            foreach (Document item in I.app.Documents.VisibleDocuments)
            {
                recover(item, fs);
            }
        }

        private string find(IEnumerable<string> fs, string f)
        {
            f = file.name(f);
            f = f.Trim();
            if (f.StartsWith("Закл. _ гайка")) f = f.Substring(0, f.Length - 3);
            foreach (var item in fs)
            {
                if (item.IndexOf(f) != -1)
                {
                    return item;
                }
            }
            return null;
        }

        private void recover(Document doc, IEnumerable<string> fs)
        {
            foreach (DocumentDescriptor dd in doc.ReferencedDocumentDescriptors)
            {
                if (dd.FullDocumentName.EndsWith("iam"))
                {
                    Document rd = dd.ReferencedDocument as Document;
                    rd = I.open(rd.FullDocumentName, false, false);
                    if (rd == null) continue;
                    recover(rd, fs);
                }
                if (dd.ReferenceMissing)
                {
                    string f = find(fs, dd.FullDocumentName);
                    if (f != null)
                        try
                        {
                            dd.ReferencedFileDescriptor.ReplaceReference(f);
                        }
                        catch (Exception)
                        {
                        }
                }
            }
        }

        private Point2d connectPoint(SketchedSymbol sym)
        {
            foreach (SketchPoint item in sym.Definition.Sketch.SketchPoints)
            {
                if (!item.InsertionPoint && item.ConnectionPoint) return item.Geometry;
            }
            return I.CP2d(0, 0);
        }

        private void izvAdd()
        {
            DrawingDocument m_Drw = null;
            string fn = u.OFD(I.aPath(), "Drawing(*.idw)|*.idw;Part(*.ipt)|*.ipt");
            Document doc = I.open(fn, false, false);
            if (doc == null) return;
            if (doc.DocumentType == DocumentTypeEnum.kDrawingDocumentObject)
            {
                m_Drw = doc as DrawingDocument;
                doc = u.referendedDoc(doc);
            }
            string pn = u.getPropValue(doc, "Part Number"), desc = u.getPropValue(doc, "Description");
            MyForm F = new MyForm("IzvContentInterface.xml", "Извещение");
            //F.bnts[0].Click += fillet_Click;
            F.f.ShowDialog();
            string n = "ИзвС";
            List<string> vals = new List<string>();
            string act = F.cbs[0].Text, s1 = F.cbs[1].Text, s2 = F.cbs[2].Text, s3 = F.cbs[3].Text;
            vals.Add(pn + " (" + desc + ") " + act);
            if (s1 != null && s1 != "") { n = "ИзвС1"; vals.Add("1. " + s1); }
            if (s2 != null && s2 != "") { n = "ИзвС2"; vals.Add("2. " + s2); }
            if (s3 != null && s3 != "") { n = "ИзвС3"; vals.Add("3. " + s3); }
            if (m_Drw == null)
                vals.Add(doc.FullDocumentName);
            else vals.Add(m_Drw.FullDocumentName + @"|" + doc.FullDocumentName);
            DrawingDocument drw = (DrawingDocument)I.aDoc();
            Point2d pt;
            if (drw.ActiveSheet.TitleBlock.Name == "ГОСТ - Форма 1") pt = I.CP2d(2, 16.7);
            else pt = I.CP2d(2, 27.2);
            if (drw.ActiveSheet.SketchedSymbols.Count != 0)
            {
                SketchedSymbol sym = drw.ActiveSheet.SketchedSymbols[drw.ActiveSheet.SketchedSymbols.Count];
                pt = sym.Position;
                Point2d pt1 = connectPoint(sym);
                pt.X += pt1.X; pt.Y += pt1.Y;
            }
            drw.ActiveSheet.SketchedSymbols.Add(drw.SketchedSymbolDefinitions[n], pt, 0, 1, vals.ToArray());
            var props = izvProps((Document)drw, new List<string>() { "Изв", "Creation Time", "Применяемость", "Приложение" });
            var spl = props["Применяемость"].Value.ToString().Split(';').ToList();
            if (pn.Contains("."))
            {
                string type = pn.Split('.')[0];
                if (!spl.Contains(type)) spl.Add(type);
            }
            string v = String.Join<string>(";", spl);
            props["Применяемость"].Value = v.TrimStart(';');
            //u.addProp(doc, "Изв", props["Изв"].Value);
            //string d = props["Creation Time"].Expression.ToString();
            //d = d.Remove(d.Length - 4, 2);
            //u.addProp(doc, "ИзвД", d);
            //fillIzv(m_Drw);
        }

        private void fillIzv(DrawingDocument doc, string time, string izv, string izvName, string templ)
        {
            if (doc == null) return;

            DrawingDocument m_Drw = doc as DrawingDocument;
            if (m_Drw != null)
            {
                string author = null;

                foreach (Document rdoc in m_Drw.ReferencedDocuments)
                {
                    u.addProp(rdoc, "Изв", izv);
                    u.addProp(rdoc, "ИзвД", time);
                    author = u.getPropValue(rdoc, "Author");

                    var pdf = new PDFOp();
                    pdf.setNameIzv(rdoc, "1");

                    rdoc.Save2(true);
                }
                I.open(m_Drw.FullDocumentName, true, true);
                InvAddIn.Drawings.addIzv(m_Drw, izvName, time, author, templ);
            }
        }

        private Dictionary<string, Property> izvProps(Document doc, List<string> vals)
        {
            Dictionary<string, Property> props = new Dictionary<string, Property>();
            foreach (var item in vals)
            {
                props.Add(item, u.getProp(doc, item));
            }
            return props;
        }

        private void izvFill()
        {
            MyForm F = new MyForm("IzvFillInterface.xml", "Извещение");
            //F.bnts[0].Click += fillet_Click;
            F.f.ShowDialog();
            var tmp = I.aDoc();
            string num = F.cbs[0].Text, txt = F.cbs[1].Text;
            if (txt == null || txt == "") return;
            var spl = txt.Split(':');
            string code = spl[0], reason = spl[1];
            string path = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location) + "\\Изв.idw";
            Document doc = I.app.Documents.Add(DocumentTypeEnum.kDrawingDocumentObject, path);//I.aDoc();
            Property p1 = u.getProp(doc, "Причина"), p2 = u.getProp(doc, "Код");
            var props = izvProps(doc, new List<string>() { "Изв", "Creation Time" });
            string d = props["Creation Time"].Expression.ToString();
            string name = "ТПМШ." + num + "-" + d.Substring(d.Length - 2, 2);
            props["Изв"].Value = name;
            p1.Value = reason; p2.Value = code;
            var p = file.p(tmp.FullDocumentName);
            tmp.Close(true);
            doc.SaveAs(p + name + ".idw", false);
        }

        private void izvAppl()
        {

        }

        private string relativePath(string fn, string fn2)
        {
            return fn2;
        }

        private void complect()
        {
            var curPr = I.app.DesignProjectManager.ActiveDesignProject.Name;
            DrawingDocument drw = (DrawingDocument)I.aDoc();
            string num = u.getPropValue((Document)drw, "Изв");
            string path = I.aPath() + num + "\\";
            file.createPath(path);
            var props = izvProps((Document)drw, new List<string>() { "Изв", "Creation Time" });
            string d = props["Creation Time"].Expression.ToString();
            //d = d.Remove(d.Length - 4, 2);
            PDFOp pdf = new PDFOp();
            List<IzvData> lst = new List<IzvData>();

            //string txt = u.getPropValue((Document)drw, "Comments");
            //var spl = txt.Split('@');

            //foreach (string item in spl)
            //{
            //    Document doc = I.open(item, false, false);
            //    u.getPropValue(doc, "Revision Number");
            //}
            foreach (Sheet sh in drw.Sheets)
            {
                foreach (SketchedSymbol ss in sh.SketchedSymbols)
                {
                    if (!ss.Name.StartsWith("Изв")) continue;
                    var ssd = ss.Definition;
                    var tb = ssd.Sketch.TextBoxes[ssd.Sketch.TextBoxes.Count];
                    var str = ss.GetResultText(tb);
                    var izv = new IzvData(path, str, d, num);
                    izv.setTempl(drw.FullDocumentName);
                    lst.Add(izv);
                }
            }
            pdf.addPDF(drw, path);
            drw.Close();
            foreach (var item in lst)
            {
                PDFDXFIzv(item, pdf);
            }
            //List<Document> docs = new List<Document>();
            //foreach (Document item in I.app.Documents)
            //{
            //    //item.Close(false);
            //    docs.Add(item);
            //}
            //var pdfop = new PDFOp(docs);
            //foreach (Document item in I.app.Documents)
            //{
            //    item.Close(false);
            //}
            //I.app.DesignProjectManager.DesignProjects.ItemByName[curPr].Activate();
        }

        private void PDFDXFIzv(IzvData izv, PDFOp pdf)
        {
            var spl = izv.fn.Split('|');
            foreach (var item in spl)
            {
                PDFDXFIzv(pdf, izv, item);
            }
        }

        private void PDFDXFIzv(PDFOp pdf, IzvData izv, string fn)
        {
            var mgr = I.app.DesignProjectManager;
            var curPath = I.curProjPath();
            if (!fn.StartsWith(curPath))
            {
                var pr = changeProject(fn);
                if (pr != null)
                {
                    foreach (Document item in I.app.Documents)
                    {
                        item.Close(true);
                    }
                    mgr.DesignProjects.AddExisting(pr).Activate();
                }
            }
            var doc = I.open(fn);
            if (doc.DocumentType == DocumentTypeEnum.kDrawingDocumentObject)
                fillIzv(doc as DrawingDocument, izv.date, izv.num, izv.izv, izv.drwTempl);
            pdf.add(doc, izv.path);

            doc.Save();
        }

        private string changeProject(string fn)
        {
            bool flag = true;
            var path = file.p(fn);
            while (flag)
            {
                var fs = file.getFiles(path, ".ipj");
                if (fs.Count() != 0)
                {
                    return fs.ElementAt(0);
                }
                path = file.sub(path, "\\", 2);
            }
            return null;
        }

        private void добавитьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            izvAdd();
        }

        private void заполнитьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            izvFill();
        }

        private void сформироватьКомплектToolStripMenuItem_Click(object sender, EventArgs e)
        {
            complect();
        }

        private void создатьДеталиToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            createParts();
        }

        public void createParts()
        {
            PartDocument doc = I.aDoc() as PartDocument;
            string path = file.p(doc.FullDocumentName);
            if (doc == null) return;
            PartComponentDefinition def = doc.ComponentDefinition;
            string templPath = I.dPr().TemplatesPath;
            string type = u.getPropValue((Document)doc, "Type"), type1 = "", path1,
                    dn = u.getPropValue((Document)doc, "DecNumber");
            if (dn == "") dn = "00.001";
            string sms = I.getSMCD((Document)doc).ActiveSheetMetalStyle.Name;
            var smcd = I.getSMCD(doc as Document);
            var ss = doc.SelectSet;
            List<SurfaceBody> bodies = new List<SurfaceBody>();
            if (ss.Count > 0)
            {
                bodies = u.gets<SurfaceBody>(ss, f => f is SurfaceBody).ToList(); ;
            }
            else
            {
                bodies = u.gets<SurfaceBody>(def.SurfaceBodies, f => true).ToList();
            }
            for (int j = 1; j < bodies.Count + 1; j++)
            {
                if (!def.SurfaceBodies[j].Visible) continue;
                if (def.SurfaceBodies[j].Name.StartsWith("!")) continue;
                type1 = type;
                path1 = path;
                var spl = bodies[j-1].Name.Split('$');
                string num = u.getDN(spl, dn, j);
                if (spl[0].IndexOf("^") != -1)
                {
                    var spl1 = spl[0].Split('^');
                    spl[0] = spl1[1];
                    var t = u.getPropValue(doc as Document, spl1[0]);
                    var tbase = u.getPropValue(doc as Document, "Type");
                    Regex r = new Regex(@"\d+");
                    var m = r.Match(t);
                    if (m.Groups.Count > 0)
                    {
                        var p = file.findPath(path, m.Groups[0].Value);
                        if (p != "") path1 = $"{p}\\";
                        else
                        {
                            path1 = getPath(t, path, tbase);   
                        }
                    }
                    else
                    {
                        path1 = getPath(t, path, tbase);
                    }
                    file.createPath(path1);
                    type1 = t;
                }
                string name = type1 + "." + num + " (" + spl[0] + ")";
                Document d = I.newDoc(path1 + name + ".ipt", templPath + "Листовой.ipt", false);
                u.addProps(d, "Type=" + type1 + "$" + "DecNumber=" + num + "$" + "Description=" + spl[0]);
                u.addProp(d, "Part Number", "=<Type>.<DecNumber>");
                I.addSB(d, doc.FullDocumentName, bodies[j-1]);

                if (smcd != null)
                {

#if INV14
                    sms = smcd.ActiveSheetMetalStyle.Name;
#else
                    sms = smcd.GetBodySheetMetalStyle(bodies[j - 1]).Name;
#endif

                    I.setSMS(d, sms);
                }
                d.Save();
            }
        }

        public string getPath(string t, string path, string tbase)
        {
            Regex r = new Regex(@"\d+");
            var m = r.Match(t);
            var m1 = r.Match(tbase);
            return (m1.Value == m.Value) ? path :
            $"{path}{t}\\";
        }

        public void jointParts()
        {
            PartDocument doc = I.aDoc() as PartDocument;
            string path = file.p(doc.FullDocumentName);
            if (doc == null) return;
            PartComponentDefinition def = doc.ComponentDefinition;
            string templPath = I.dPr().TemplatesPath;
            string type = u.getPropValue((Document)doc, "Type"),
                    dn = u.getPropValue((Document)doc, "DecNumber");
            if (dn == "") dn = "00.001";
            string sms = I.getSMCD((Document)doc).ActiveSheetMetalStyle.Name;
            var ie = file.getFiles(path, ".ipt", System.IO.SearchOption.AllDirectories);
            for (int j = 1; j < def.SurfaceBodies.Count + 1; j++)
            {
                if (def.SurfaceBodies[j].Name.StartsWith("!")) continue;
                if (!def.SurfaceBodies[j].Visible) continue;
                var spl = def.SurfaceBodies[j].Name.Split('$');
                string num = u.getDN(spl, dn, j), name;
                
                if (spl[0].IndexOf('^') != -1)
                {
                    var tmp =  spl[0].Split('^');
                    name = u.getPropValue(doc as Document, tmp[0]) + '.' + num + " (" + tmp[1] + ")";
                }
                else
                    name = type + "." + num + " (" + spl[0] + ")";
                var ffn = findDoc(name, ie);
                if (ffn == null) continue;
                Document d = I.open(ffn);
                //Document d = I.newDoc(path + name + ".ipt", templPath + "Листовой.ipt", false);
                var iMs = findIMates(def, def.SurfaceBodies[j].Name);
                if (iMs.Count == 0) continue;
                I.addSB(d, d.FullDocumentName, null, iMs);
                d.Save2(false);
            }
        }
        public string findDoc(string name, IEnumerable<string> files)
        {
            foreach (var f in files)
            {
                var n = file.name(f);
                if (name == n) return f;
            }
            return null;
        }
        public List<string> findIMates(PartComponentDefinition def, string sbName)
        {
            List<string> imNames = new List<string>();
            var ie = u.gets<iMateDefinition>(def.iMateDefinitions, f => !f.Exported && !f.Suppressed);
            foreach (iMateDefinition item in ie)
            {
                if (checkIM(sbName, item)) imNames.Add(item.Identifier);
            }
            return imNames;
        }
        public bool checkIM(string sbName, iMateDefinition imDef)
        {
            InsertiMateDefinition iMdef = imDef as InsertiMateDefinition;
            CompositeiMateDefinition cMdef = imDef as CompositeiMateDefinition;
            if (cMdef != null) iMdef = cMdef[1] as InsertiMateDefinition;
            if (iMdef != null)
            {
                if (iMdef.Entity == null) return false;
                var e = iMdef.Entity as Edge;
                if (e.Parent.Name == sbName)
                    return true;
            } else
            {
                foreach (var item in cMdef)
                {
                    dynamic im = item;
                    var ent = im.Entity;
                    if (ent == null) return false;
                    if (ent is WorkPlane) continue;
                    if (ent.Parent.Name == sbName)
                        return true;
                }
            }
            return false;
        }

        private void changeMat()
        {
            var doc = I.aDoc();
            var p = file.p(doc.FullDocumentName);
            var fs = file.getFiles(p, ".ipt");
            var mats = I.app.StylesManager.Materials;
            Inventor.Material m = null;
            foreach (Inventor.Material item in mats)
            {
                var o = item.Name;
                if (o == "ОЦ") m = item;
            }
            foreach (var item in fs)
            {
                var part = I.open(item);
                var smcd = I.getSMCD(part);
                if (smcd == null) continue;
                var st = smcd.ActiveSheetMetalStyle;
                var n = st.Name;
                if (!n.ToLower().StartsWith("оцинк")) continue;
                if (m != null) st.Material = m;
                smcd.Material = m;
            }
        }

        private void оЦToolStripMenuItem_Click(object sender, EventArgs e)
        {
            changeMat();
        }

        public void breakView()
        {
            DrawingDocument drw = I.aDoc() as DrawingDocument;
            if (drw == null) return;
            var ss = drw.SelectSet;
            if (ss.Count != 1) return;
            DrawingView dv = ss[1] as DrawingView;
            if (dv == null) return;
            MyForm F = new MyForm("BreakInterface.xml", "Разрыв вида");
            F.f.ShowDialog();
            string dir = F.cbs[0].Text;
            double w = ut.convToDouble(F.cbs[1].Text);
            w *= 0.1;

            var v = dir.ToLower() == "горизонтальный" ? I.CV2d(0, 1) : I.CV2d(1, 0);
            var pr = new Projections(dv, v, I.CP2d(5, 20));
            pr.fill();
            pr.join();
            var b = I.Box(dv);
            var p = new Projection(b, v);
            double wCut = p.L - w;
            pr.addHoles(p);
            pr.addCenter();
            pr.setCut(wCut);
            pr.addBreaks();
        }

        private void разрывВидаToolStripMenuItem_Click(object sender, EventArgs e)
        {
            breakView();
        }
        public void findBrowserNode()
        {
            var pDoc = I.aDoc();
            var ss = pDoc.SelectSet;
            if (ss.Count != 1) return;
            var br = pDoc.BrowserPanes.ActivePane;
            var e = ss[1] as Edge; var f = ss[1] as Face;
            PartFeature pf = null;
            if (f != null) pf = f.CreatedByFeature;
            if (e != null)
            {
                var f1 = e.Faces[1]; var f2 = e.Faces[2];
                pf = f1.Evaluator.Area <= f2.Evaluator.Area ? f1.CreatedByFeature : f2.CreatedByFeature;
            }
            if (pf == null) return;
            var node = br.GetBrowserNodeFromObject((object)pf);
            node.DoSelect();
            this.Close();
        }
        private void найтиЭлементToolStripMenuItem_Click(object sender, EventArgs e)
        {
            findBrowserNode();
        }
        public void renameBrNodes()
        {
            var doc = I.aDoc();
            if (doc.DocumentType != DocumentTypeEnum.kPartDocumentObject) return;
            var br = doc.BrowserPanes.ActivePane;
            foreach (BrowserNode node in br.TopNode.BrowserNodes[1].BrowserNodes)
            {
                var pf = node.NativeObject as PartFeature;
                if (pf == null) continue;
                if (getPrefixSurf(doc, pf))
                {
                    string repl = "";
                    if (pf.Type == ObjectTypeEnum.kHoleFeatureObject)
                    {
                        repl += "D";
                        getFeatName(pf, ref repl);
                        pf.Name = pf.Name.Replace("Отверстие", repl);
                    }
                    if (pf.Type == ObjectTypeEnum.kFlangeFeatureObject)
                    {
                        repl += "Фл";
                        getFeatName(pf, ref repl);
                        pf.Name = pf.Name.Replace("Фланец", repl);
                    }
                    if (pf.Type == ObjectTypeEnum.kContourFlangeFeatureObject)
                    {
                        repl += "ФлКонт";
                        getFeatName(pf, ref repl);
                        pf.Name = pf.Name.Replace("Фланец по контуру", repl);
                    }
                    if (pf.Type == ObjectTypeEnum.kCornerRoundFeatureObject)
                    {
                        repl += "R";
                        getFeatName(pf, ref repl);
                        pf.Name = pf.Name.Replace("УглСкругление", repl);
                    }
                }
            }
        }
        public bool getPrefixSurf(Document doc, PartFeature pf)
        {
            PartComponentDefinition def = I.getPCD(doc);
            if (def != null && def.SurfaceBodies.Count > 1)
            {
                string str = "";
                var sb = pf.SurfaceBodies[1];
                var spl = sb.Name.Split('$');
                var spl1 = spl[0].Split(' ');
                List<char> fonetics = new List<char>() { 'а', 'е', 'и', 'о', 'у', 'ю', 'я', 'ы', 'э', 'ё' };
                int i = 0;
                foreach (var w in spl1)
                {
                    foreach (var item in w)
                    {
                        if (i == 0) str += item.ToString().ToUpper();
                        else if (i == 1) str += item;
                        else
                        {
                            if (!fonetics.Contains(item)) str += item;
                            else break;
                        }
                        i++;
                    }
                    i = 0;
                }
                if (pf.Name.StartsWith(str)) return false;
                pf.Name = str + "_" + pf.Name;
                return true;
            }
            return false;
        }
        public void getFeatName(PartFeature pf, ref string repl)
        {
            double d = (double)pf.FeatureDimensions[1].Parameter.Value;
            repl += (d * 10).ToString();
            if (pf.FeatureDimensions.Count > 1)
            {
                d = u.radToDeg((double)pf.FeatureDimensions[2].Parameter.Value);
                repl += "x" + d;
            }
            repl += "  :";
        }
        private void переименоватьУзлыToolStripMenuItem_Click(object sender, EventArgs e)
        {
            renameBrNodes();
        }

        public void flat()
        {
            MyForm F = new MyForm("FPInterface.xml", "Развертка");
            F.f.ShowDialog();
            foreach (var item in I.getDocs(fi => fi.DocumentType == DocumentTypeEnum.kPartDocumentObject))
            {
                item.Activate();
                if (F.chks[0].Checked)
                    flat(item, F.cbs[0].Text, F.chks[1].Checked);
                else
                    flat(item, "", F.chks[1].Checked);
            }
        }

        public void flat(Document doc, string mat, bool num)
        {
            var smcd = I.getSMCD(doc);
            if (smcd == null) return;
            if (smcd.HasFlatPattern)
            {
                rename_and_cut(smcd.FlatPattern, mat, num);
                return;
            }
            if (smcd.SurfaceBodies.Count != 1) return;
            var sb = smcd.SurfaceBodies[1];
            var fs = u.gets<Face>(sb.Faces, fi => fi.SurfaceType == SurfaceTypeEnum.kPlaneSurface)
                .OrderByDescending(el => el.Evaluator.Area);
            var f = getFace(fs);
            //var ed = u.get<Edge>(f.Edges, fi => fi.CurveType == CurveTypeEnum.kLineCurve && checkEdge(fi, f));
            //if (ed == null) return;
            setView();
            var st = FPOp.getStyle(f);
            if (st != null)
            {
                var st1 = FPOp.findStyle(smcd, st.Name);
                if (st1 != null)
                {
                    st1.Activate();
                }
            }
            smcd.Unfold2(f);
            FlatPattern fp = smcd.FlatPattern;
            Edge edf = null;
            //getFBR(fp, ed, ref edf);
            FlatBendResult fbr = null;
            rename_and_cut(fp, mat, num);

            //if (fp.FlatBendResults.Count == 0) return;
            //fbr = getFBR(fp);
            //Edge e; Face fa;
            //getFace(fbr, out e, out fa);
            //var fpo = fp.FlatPatternOrientations.ActiveFlatPatternOrientation;

            //setAlign(fpo, fa, e);

            //var v1 = I.BoxCenter(sb.RangeBox).VectorTo(cenFace(fa));
            //var v2 = I.BoxCenter(fp.SurfaceBodies[1].RangeBox).VectorTo(u.midPt(e));
            //v1.Normalize(); v2.Normalize();
            //var a = v1.DotProduct(v2);
            //if (a < 0) fpo.FlipAlignmentAxis = true;

            //fp.ExitEdit();
        }
        public void rename_and_cut(FlatPattern fp, string mat, bool num)
        {
            if (fp != null)
            {
                var smcd = fp.Parent;
                var doc = smcd.Document as Document;
                var name = smcd.ActiveSheetMetalStyle.Name;
                if (mat != "" && name.ToLower().StartsWith(mat))
                {
                    if (fp.FlatBendResults.Count != 0 && fp.Features.CutFeatures.Count == 0)
                    {
                        var fpOp = new FPOp(doc, I.app);
                        //fpOp.Cut(fp);
                    }
                }
                if (fp.FlatBendResults.Count != 0 && num)
                    FPOp.renameBendOrder(fp);
            }
        }
        public static void setView(Face f = null, Camera cam = null)
        {
            var v = I.app.ActiveView;
            var c = v.Camera;
            var pt = c.Eye;
            if (f == null)
            {
                c.Eye = I.CP(0, 0, pt.Z);
                c.Target = I.CP();
                c.UpVector = I.CUV(0, 1, 0);
                c.ApplyWithoutTransition();
            }
            else
            {
                Point cen = f.Evaluator.RangeBox.MaxPoint;
                //Point cen = u.midPt(f.Evaluator.RangeBox);
                c.Eye = I.CP(cen.X+2, cen.Y+2, cen.Z);
                c.Target = I.CP(cen.X, cen.Y, 0);
                c.UpVector = I.CUV(0, 1, 0);
                c.ApplyWithoutTransition();
            }
        }
        public bool checkEdge(Edge e, Face bf)
        {
            var f = u.get<Face>(e.Faces, fi => !fi.Equals(bf));
            return f.SurfaceType == SurfaceTypeEnum.kCylinderSurface;
        }
        public Point cenFace(Face f)
        {
            var e = u.get<Edge>(f.Edges, fi => fi.CurveType == CurveTypeEnum.kLineCurve);
            return u.midPt(e);
        }
        public FlatBendResult getFBR(FlatPattern fp)
        {
            var fbrs = u.gets<FlatBendResult>(fp.FlatBendResults, fi => true)
                .OrderByDescending(fi => u.getLenght(fi.Edge));
            foreach (var item in fbrs)
            {
                var es = u.gets<Edge>(fp.SurfaceBodies[1].Edges, fi => fi.CurveType == CurveTypeEnum.kLineCurve &&
                u.isParallel(fi, item.Edge));
                if (es == null) continue;
                foreach (var e in es)
                {
                    if (u.getLenght(e) * 0.7 <= u.getLenght(item.Edge)) return item;
                }
            }
            return null;
        }
        public void getFBR(FlatPattern fp, Edge e, ref Edge ed)
        {
            Face f = null; ed = null;
            foreach (FlatBendResult item in fp.FlatBendResults)
            {
                getFace(item, out ed, out f);
                var e1 = u.get<Edge>(f.Edges, fi => fi.Equals(e));
                if (e1 != null) { ed = item.Edge; break; }
            }
        }
        public Face getFace(IEnumerable<Face> fs)
        {
            Face f1 = fs.ElementAt(0), f2 = fs.ElementAt(1);
            Point pt = I.CP();
            Plane pl1 = f1.Geometry as Plane, pl2 = f2.Geometry as Plane;
            double d1 = u.distToPlane(pl1, pt), d2 = u.distToPlane(pl2, pt);
            return d1 >= d2 ? f1 : f2;
        }
        //public void setAlign(FlatPatternOrientation fpo, Face f, Edge e, bool rev = false)
        //{
        //    Cylinder c = f.Geometry as Cylinder;
        //    LineSegment l = e.Geometry as LineSegment;
        //    fpo.AlignmentAxis = e;
        //    if (c.AxisVector.IsPerpendicularTo(l.Direction)) fpo.AlignmentType = AlignmentTypeEnum.kVerticalAlignment;
        //    else if (c.AxisVector.IsParallelTo(l.Direction))
        //        fpo.AlignmentType = AlignmentTypeEnum.kHorizontalAlignment;
        //    fpo.FlipAlignmentAxis = rev;
        //}
        public void getFace(FlatBendResult fbr, out Edge e, out Face f)
        {
            var b = fbr.Bend;
            e = fbr.Edge;
            f = null;
            if (b.BackFaces.Count > 0) f = b.BackFaces[1];
            else if (b.FrontFaces.Count > 0) f = b.FrontFaces[1];
        }

        private void разверткаToolStripMenuItem_Click(object sender, EventArgs e)
        {
            flat();
        }

        private void убратьВинтToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Slice sl = new Slice(I.aDoc());
            sl.unfold();
            sl.addSketch();
            sl.addExtrude();
        }

        private void совместитьПлоскостиToolStripMenuItem_Click(object sender, EventArgs e)
        {
            orient();
        }

        private void адаптивныеПлоскостиToolStripMenuItem_Click(object sender, EventArgs e)
        {
            adaptivePlane();
        }

        private void связатьКПToolStripMenuItem_Click(object sender, EventArgs e)
        {
            jointParts();
        }

        private void зеркльныеКПToolStripMenuItem_Click(object sender, EventArgs e)
        {
            MirrorIMate mim = new MirrorIMate(I.aDoc());
        }

        private void копироватьПроектToolStripMenuItem_Click(object sender, EventArgs e)
        {
            var pdoc = I.aDoc() as PartDocument;
            if (pdoc == null) return;
            CopyProduct c = new CopyProduct(pdoc);
        }

        private void перенестиКПВСборкуToolStripMenuItem_Click(object sender, EventArgs e)
        {
            var doc = I.aDoc();
            MoveIM i = new MoveIM(doc);
            var ss = doc.SelectSet;
            i.move(ss);
        }

        private void копироватьКПToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Macros.MyEvArgs args = new Macros.MyEvArgs();
            var b = sender as ToolStripMenuItem;
            args.name = b.Name;
            OnMyEvent(args);
            this.Close();
            //this.Hide();
            //CopyIMate c = new CopyIMate(I.aDoc());
            //this.Close();
        }

        private void перенаправитьToolStripMenuItem1_Click_1(object sender, EventArgs e)
        {
            Document doc = Macros.StandardAddInServer.m_inventorApplication.ActiveDocument;
            //string newName = InvDoc.u.OFD(System.IO.Path.GetDirectoryName(doc.FullDocumentName));
            foreach (Document docum in Macros.StandardAddInServer.m_inventorApplication.Documents.VisibleDocuments)
            {
                if (docum.DocumentType == DocumentTypeEnum.kDrawingDocumentObject)
                {
                    var newName = Parts.findDocumentName(docum as DrawingDocument);
                    if (newName == null) continue;
                    var p = file.p(docum.FullDocumentName);
                    var fs = file.getFiles(p, newName);
                    foreach (var item in fs)
                    {
                        if (item == docum.FullDocumentName) continue;
                        Parts.replaceFullFileName((DrawingDocument)docum, item);
                    }
                }
            }
        }

        private void названиеМоделиToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Document doc = I.aDoc();
            Site site = new Site(doc);
            site.act("model");
        }

        private void спецификацияToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Document doc = I.aDoc();
            Site site = new Site(doc);
            site.act("xml");
        }

        private void тест2ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            var doc = I.aDoc();
            DrawDims d = new DrawDims(doc);
            //draw_chords();
            //equation_curves();
        }

        private void поверхностьПодОвалToolStripMenuItem_Click(object sender, EventArgs e)
        {
            slotSurface();
        }

        private void весьПроектToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Document doc = I.aDoc();
            Site site = new Site(doc);
            site.act("OpenFiles");
            site.act("listRead");
        }

        private void создатьСписокФайловToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Document doc = I.aDoc();
            Site site = new Site(doc);
            site.act("listWrite");
        }

        private void gltfToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Document doc = I.aDoc();
            string p = file.p(doc.FullDocumentName);
            Site site = new Site(doc);
            site.act("OpenFiles");
            Dictionary<string, string> props = new Dictionary<string, string>()
            {
                {"PartNumber", "Part Number" }, {"Description", "Description"}, {"RevisionNumber", "Revision Number"}
            };
            foreach (Document item in site.docs)
            {
                XElement el = new XElement("MKart");
                AssemblyDocument adoc = doc as AssemblyDocument;
                var acd = adoc.ComponentDefinition;
                var bom = acd.BOM;
                bom.PartsOnlyViewEnabled = true;
                //foreach (BOMView item in bom.BOMViews)
                //{
                //    var te = item;
                //}
                var view = bom.BOMViews[3];
                projectProperties.minMKart(view, props, el);
                Excel.InvExcel exc = new Excel.InvExcel($"{p}{file.name(item.FullDocumentName)}.xlsx");
                exc.add(el);
            }
            //Document doc = I.aDoc();
            //Gltf gltf = new Gltf(doc);
        }

        private void моделиИзСпискаToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Document doc = I.aDoc();
            Site site = new Site(doc);
            site.act("OpenFiles");
            site.act("model");
        }

        private void dИзСпискаФайловToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Document doc = I.aDoc();
            Site site = new Site(doc);
            site.vis = true;
            site.act("OpenFiles");
            site.act("3dPDF");
        }

        private void данныеИзTitleToolStripMenuItem_Click(object sender, EventArgs e)
        {
            var adoc = I.aDoc();
            if (adoc.DocumentType != DocumentTypeEnum.kDrawingDocumentObject) return;
            DrawingDocument doc = (DrawingDocument)adoc;
            MyForm F = new MyForm("BaseCBInterface.xml", "Свойство");
            foreach (TitleBlockDefinition item in doc.TitleBlockDefinitions)
            {
                F.cbs[0].Items.Add(item.Name);
            }

            F.f.ShowDialog();
            string name = F.cbs[0].Text;
            u.getTitleBlockToXML(doc, name);
            //double rin = ut.convToDouble(F.cbs[1].Text);
        }

        private void сохранитьDWGToolStripMenuItem_Click(object sender, EventArgs e)
        {
            string p = file.p(I.aDoc().FullDocumentName);
            foreach (Document item in I.visDocs())
            {
                if (item.DocumentType == DocumentTypeEnum.kDrawingDocumentObject)
                    u.saveToDWG((DrawingDocument)item, p);
                else if (item.DocumentType != DocumentTypeEnum.kDrawingDocumentObject)
                {
                    u.createDrawing(item, p);
                }
            }
        }

        private void вставкаПоСценариюToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Hide();
            smartInsert();
            this.Show();
        }

        public void smartInsert()
        {
            ContentOp cop = new ContentOp();
            cop.insertPart();
        }

        private void фаскиToolStripMenuItem_Click(object sender, EventArgs e)
        {
            MyForm F = new MyForm("ChamferInterface.xml", "head");
            F.f.ShowDialog();
            string rin = F.cbs[1].Text, rout = F.cbs[0].Text;
            foreach (Document doc in I.app.Documents.VisibleDocuments)
            {
                var def = I.getSMCD(doc);
                if (def == null) continue;
                for (int i = 0; i < def.SurfaceBodies.Count; i++)
                {
                    chamfer(def.SurfaceBodies[i + 1],
                        def.SurfaceBodies[i + 1].ConcaveEdges.Cast<Edge>(), def, rin, 5);
                    chamfer(def.SurfaceBodies[i + 1],
                        def.SurfaceBodies[i + 1].ConvexEdges.Cast<Edge>(), def, rout, 5);
                }
            }
            this.Close();
        }

        private void граниВDXFToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Macros.MyEvArgs args = new Macros.MyEvArgs();
            var b = sender as ToolStripMenuItem;
            args.name = b.Name;
            OnMyEvent(args);
            this.Close();
        }

        private void тестSplToolStripMenuItem_Click(object sender, EventArgs e)
        {
            var doc = I.aDoc();
            var ss = doc.SelectSet;
            Edge ed = ss[1] as Edge;
            if (ed == null) return;
            BSplineCurve spl = ed.Geometry as BSplineCurve;
            u.splineInfo(spl);
        }

        private void заменитьМодельZToolStripMenuItem_Click(object sender, EventArgs e)
        {
            replaceDraw();
        }

        public void replaceDraw()
        {
            var doc = I.aDoc();
            if (doc.DocumentType != DocumentTypeEnum.kDrawingDocumentObject) return;
            var drw = doc as DrawingDocument;
            string p = file.p(drw.FullDocumentName), fn = drw.FullDocumentName;
            string o = u.OFD(p, fn: fn);
            doc = I.open(o);
            string nn = doc.FullDocumentName, p1 = file.p(nn), name = file.name(nn);
            nn = $"{p1}{name}.idw";
            drw.SaveAs(nn, true);
            doc = I.open(nn, true);
            drw = doc as DrawingDocument;
            Parts.replaceFullFileName(drw, o);
        }

        private void создатьСборкуToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            assembly();
        }

        private void автоЭскизыToolStripMenuItem_Click(object sender, EventArgs e)
        {
            var doc = I.aDoc();
            var ah = new AutoHoles(doc);
        }

        private void автоКПToolStripMenuItem_Click(object sender, EventArgs e)
        {
            var doc = I.aDoc();
            var ah = new AutoiMates(doc);
        }

        private void фланецОтступToolStripMenuItem_Click(object sender, EventArgs e)
        {
            var doc = I.aDoc();
            u.setFlangeOffset(doc);
        }

        private void оптБрToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Document doc = I.aDoc();
            Site site = new Site(doc);
            site.act("model2");
        }

        private void пазМультидетальToolStripMenuItem_Click(object sender, EventArgs e)
        {
            var smcd = I.getSMCD();
            var ss = I.getSS();

            if (ss.Count != 2) return;
            var bodies = GetBodies(ss);
            PlanarSketch psp = ss[1] as PlanarSketch;
            if (psp == null) return;
            
            //var pc = psp.ProjectedCuts.Add();
            try
            {
                Offset.offsetAdaptive(smcd, psp.Name, 0.11, 0.8, 0.63, bodies);
                psp.Visible = false;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private void скрытьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            setVisibleBody();
        }

        private void автоБлокRToolStripMenuItem_Click(object sender, EventArgs e)
        {
            autoBlock();
        }

        private void autoBlock()
        {
            Block b;
            var doc = I.aDoc();
            var smcd = I.getSMCD(doc);
            if (smcd == null) return;
            var ss = I.getSS(doc);
            if (ss.Count == 0) b = new Block(null, smcd);
            else
            {
                PlanarSketch ps = ss[1] as PlanarSketch;
                b = new Block(ps, smcd);
            }
        }

        void copyBlock()
        {
            var doc = I.aDoc();
            string path = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
            var fn = $"{path}\\Blocks.ipt";
            var from = I.open(fn, false, false);
            Block.copy(from, doc);
        }

        void insertBlock()
        {
            var doc = I.aDoc();
            string path = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
            var fn = $"{path}\\Blocks.ipt";
            var from = I.open(fn, false, false);
            var sb = Block.copy(from, doc);
            Block.insertBlock(doc, sb);
        }

        private void копироватьБлокToolStripMenuItem_Click(object sender, EventArgs e)
        {
            copyBlock();
        }

        private void вставитьБлокToolStripMenuItem_Click(object sender, EventArgs e)
        {
            insertBlock();
        }

        private void экспортВSTPToolStripMenuItem_Click(object sender, EventArgs e)
        {
            var tmp = I.aDoc();
            var step = "STEP\\";
            var p = $"{file.p(tmp.FullDocumentName)}{step}";
            file.createPath(p);

            var docs = I.getDocs(f =>
                f.DocumentType == DocumentTypeEnum.kAssemblyDocumentObject &&
                u.getPropValue(f, "Part Number").Contains("00.000"));

            var cont = I.app.TransientObjects.CreateTranslationContext();
            var tr = I.app.ApplicationAddIns.ItemById["{90AF7F40-0C01-11D5-8E83-0010B541CD80}"] as TranslatorAddIn;
            if (tr == null) return;

            foreach (Document doc in docs)
            {
                var model = InvAddIn.Site.getModel(doc);
                if (string.IsNullOrWhiteSpace(model) || model == "КЭВ-")
                    model = file.name(doc.FullDocumentName);

                var nvm = I.objs.CreateNameValueMap();
                if (tr.HasSaveCopyAsOptions[doc, cont, nvm])
                {
                    nvm.Value["IncludeSketches"] = false;
                    nvm.Value["ApplicationProtocolType"] = 5;

                    cont.Type = IOMechanismEnum.kFileBrowseIOMechanism;
                    var data = I.app.TransientObjects.CreateDataMedium();
                    data.FileName = $"{p}{model}.stp";
                    tr.SaveCopyAs(doc, cont, nvm, data);
                }
            }
        }

        private void pDFЗеркальныеДеталиToolStripMenuItem_Click(object sender, EventArgs e)
        {
            string p = file.p(I.aDoc().FullDocumentName);
            var files = file.getFiles(p, ".ipt");
            files = files.Concat(file.getFiles(p, ".iam"));
            foreach (var item in files)
            {
                Document doc = I.open(item, false, false);
                if (doc != null)
                checkPDF(doc, @"-\d\d");
            }
        }

        private void объединитьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            MultiCut mc = new MultiCut(I.aDoc());
        }

        private void меткиToolStripMenuItem_Click(object sender, EventArgs e)
        {
            mark();   
        }

        private void mark()
        {
            I.screenSilent(true);
            try
            {
                SelectSet ss = I.getSS(I.aDoc());
                if (ss.Count > 1000 && ss[1] is SurfaceBody)
                {
                    var tg = I.beginTrans("cut", ((SurfaceBody)ss[1]).ComponentDefinition.Document as Document);
                    var ie = u.gets<SurfaceBody>(ss, f => f is SurfaceBody).Select(s => s.Name).ToList();
                    foreach (string sb in ie)
                    {
                        CutFlatPattern c = new CutFlatPattern(sb);                     
                    }
                    tg.End();
                }
                else
                {
                    foreach (Document item in I.app.Documents.VisibleDocuments)
                    {
                        if (!(item.DocumentType == DocumentTypeEnum.kPartDocumentObject)) continue;
                        var tg = I.beginTrans("cut", item);
                        CutFlatPattern c = new CutFlatPattern(item);
                        tg.End();
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
            finally
            {
                I.screenSilent(false);
            }
        }

        public void link_fastener()
        {
#if INV14
#else
            I.app._LibraryDocumentModifiable = true;
#endif

            var doc = I.aDoc();
            SelectSet ss = I.getSS(doc);
            if (ss.Count == 0) return;
            bool first = true;
            string prop = "";
            Document doc1 = null, doc2 = null;
            foreach (var item in ss)
            {
                ComponentOccurrence occ = item as ComponentOccurrence;
                if (occ == null) continue;
                doc2 = occ.ReferencedDocumentDescriptor.ReferencedDocument as Document;
                if (first) {
                    doc1 = doc2;
                }
                else
                {
                    string p = u.getPropValue(doc2, "Description");
                    prop += $"{p};";
                }
                first = false;
            }
            if (ss.Count > 1)
                prop = prop.TrimEnd(';');
            u.addProp(doc1, "NoFastener", prop, false);
            doc1.Save2();
#if INV14
#else
            I.app._LibraryDocumentModifiable = false;
#endif

        }

        private void добавитьИсполнениеToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            Variable var = new Variable(I.aDoc());
            var.add();
        }

        private void связатьКрепежToolStripMenuItem_Click(object sender, EventArgs e)
        {
            link_fastener();
        }

        private void исполненияToolStripMenuItem_Click(object sender, EventArgs e)
        {
            AsmRepl ar = new AsmRepl(I.aDoc());
        }

        private void открепленныеРазмерыToolStripMenuItem_Click(object sender, EventArgs e)
        {
            DeleteEntities del = new DeleteEntities(I.aDoc());
        }

        private void объединитьГибыToolStripMenuItem_Click(object sender, EventArgs e)
        {
            var item = I.aDoc();
            flat(item, "", true);
        }

        private void проекцияToolStripMenuItem_Click(object sender, EventArgs e)
        {
            DeleteEntities del = new DeleteEntities(I.aDoc());
        }

        private void расширитьИсполнениеToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Variable var = new Variable(I.aDoc(), this);
            var.extend();
        }

        private void масштабToolStripMenuItem_Click(object sender, EventArgs e)
        {
            var doc = I.aDoc();
            var ss = doc.SelectSet;
            List<PlanarSketch> lst = new List<PlanarSketch>();
            foreach (var item in ss)
            {
                if (item is PlanarSketch) lst.Add(item as PlanarSketch);
            }
            MyForm F = new MyForm("ScaleDimInterface.xml", "Scale");
            F.f.ShowDialog();
            string name = F.cbs[0].Text, val = F.cbs[1].Text;
            var cd = I.getPCD(doc);
            cd.Parameters.UserParameters.AddByExpression(name, val, "ul");

            foreach (var ps in lst)
            {
                foreach (DimensionConstraint dc in ps.DimensionConstraints)
                {
                    if (dc.Parameter.get_Units() == "град" || dc.Parameter.get_Units() == "grad") continue;
                    dc.Parameter.Expression += $"*{name}";
                }
            }
            this.Close();
        }

        private void подавитьТелоSToolStripMenuItem_Click(object sender, EventArgs e)
        {
            setSupressBody();
        }

        private void сменитьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            changeTypeBody();
        }

        private void добавитьКПToolStripMenuItem_Click(object sender, EventArgs e)
        {
            var doc = I.aDoc();
            BodyIMate bim = new BodyIMate(doc);
            //bim.add();
        }

        private void добавитьУдалитьЗаменитьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Macros.MyEvArgs args = new Macros.MyEvArgs();
            var b = sender as ToolStripMenuItem;
            args.name = b.Name;
            OnMyEvent(args);
            this.Close();
        }

        private void выноскиГибовToolStripMenuItem_Click(object sender, EventArgs e)
        {
            DeleteEntities ents = new DeleteEntities(I.aDoc());
            ents.deleteLeaders();
        }

        private void изменитьОтверстияToolStripMenuItem_Click(object sender, EventArgs e)
        {
            changeParam();
        }

        private void открытьПоШаблонуToolStripMenuItem_Click(object sender, EventArgs e)
        {
            openFromPattern();
        }

        private void changeParam()
        {
            var doc = I.aDoc();
            if (doc.DocumentType != DocumentTypeEnum.kPartDocumentObject) return;
            var pdoc = doc as PartDocument;
            MyForm my = new MyForm("ChangeParamInterface.xml", "Изменить");
            my.f.ShowDialog();
            var t = my.cbs[0].Text;
            var f = my.cbs[1].Text;
            var r = my.cbs[2].Text;

            if (t.ToLower() == "отв")
            {
                double d = 0;
                d = u.convToDouble(f) / 10;
                if (r.IndexOf("=") != -1)
                {
                    var spl = r.Split('=');
                    var val = u.addParameter(doc, spl[0].Trim(' '), spl[1], "мм");
                    val.Expression = spl[1];
                    r = val.Name;
                }
                foreach (HoleFeature item in pdoc.ComponentDefinition.Features.HoleFeatures)
                {
                    var p = item.HoleDiameter;
                    if (u.eq(p.ModelValue, d))
                    {
                        p.Expression = r;
                    }
                }
            }
            else if (t.ToLower() == "кп")
            {
                foreach (iMateDefinition item in pdoc.ComponentDefinition.iMateDefinitions)
                {
                    if (item.Name == f) item.Name = r;
                }
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
            string file = I.p() + @"\Parameters.xml";
            if (System.IO.File.Exists(file))
                param = new XMLDoc(file, "head");
        }

        protected override void ButtonDefinition_OnExecute(NameValueMap context)
        {

            try
            {
                Macros.StandardAddInServer.forms.Add(cc);
                if (Macros.StandardAddInServer.activeteForm()) cc = new CreateComponent(I.aDoc());
                Macros.MyEvents.add(cc);
                cc.show();

                /*System.Windows.Forms.Application.Run(cc = new CreateComponent(InventorApplication.ActiveDocument));*/
                if (Macros.MyEvents.name == "копироватьКПToolStripMenuItem")
                {
                    var d = I.aDoc();
                    var ss = d.SelectSet;
                    if (ss.Count == 0) return;
                    var comp = new composeIM(ss[1], false);
                    if (d.DocumentType == DocumentTypeEnum.kAssemblyDocumentObject)
                    {
                        AssemblyDocument asm = d as AssemblyDocument;
                        var m = new CopyIMate(d.ActivatedObject as Document, comp, false);
                    }
                    else
                    {
                        var m = new CopyIMate(d, comp, false);
                    }
                }
                else if (Macros.MyEvents.name == "тестToolStripMenuItem")
                {
                    var m = new MatePlane(I.aDoc());
                }
                else if (Macros.MyEvents.name == "граниВDXFToolStripMenuItem")
                {
                    u.faceToDXF();
                }
                else if (Macros.MyEvents.name == "добавитьУдалитьЗаменитьToolStripMenuItem")
                {
                    var doc = I.aDoc();
                    AddDeleteReplace r = new AddDeleteReplace(doc);
                    while (r.run()) ;
                    r.f.f.Close();
                    r.f.f = null;
                }
                Macros.MyEvents.name = "";
            }
            //catch (InvalidOperationException ex)
            //{
            //    cc.Show();
            //}
            catch (Exception ex)
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

        public Parts(Inventor.Document doc) : base(doc)
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
                filePath = (System.IO.File.Exists(path + "\\" + "Parts.xml")) ? path + "\\" + "Parts.xml" : I.p() + @"\Parts.xml";
            n = new InvDoc.XML(filePath);

            List<XMLData> data = new List<XMLData>();
            data = n.ReadXML("Parts");
            if (libraryPath.Count == 0)
                libraryPath = null;
            WithElem<XMLData>(data, addP);
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
                    XMLDoc lib = new XMLDoc(I.p() + @"\ContentCenter.xml", "Content");
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

        public static void replaceReference(Document doc, XMLDoc xDoc, string curPath)
        {
            string b = "BASE";
            foreach (DocumentDescriptor item in doc.ReferencedDocumentDescriptors)
            {
                if (item.FullDocumentName.IndexOf("Content Center") != -1) continue;
                //if (doc.DocumentType == DocumentTypeEnum.kDrawingDocumentObject)
                //{
                string newName = xDoc.find(xDoc.El, "old", "new", ut.nameUtil(item.FullDocumentName, true));
                if (newName == "")
                {
                    string fn = file.nameWithExt(item.FullDocumentName);
                    fn = file.fn(curPath, fn);
                    if (file.check(curPath, fn)) newName = fn;
                    else if (fn.ToUpper().IndexOf(b) != -1)
                    {
                        var files = file.getFiles(curPath, b + "*");
                        if (files != null) newName = files.First();
                    }

                    //}
                }
                if (newName != "" && System.IO.File.Exists(newName) && newName != item.FullDocumentName)
                {
                    item.ReferencedFileDescriptor.ReplaceReference(newName);
                }
                if (doc.Dirty)
                {
                    doc.Update2(false);
                    doc.Save2();
                }
            }
        }

        public static void replaceFullFileName(PartDocument doc, string newName)
        {
            PartComponentDefinition compDef;
            compDef = doc.ComponentDefinition;
            //FileDescriptor fd;
            if (compDef.ReferenceComponents.DerivedPartComponents.Count != 0)
            {
                var fd = compDef.ReferenceComponents.DerivedPartComponents[1].ReferencedDocumentDescriptor;

                //fd = compDef.Parameters.DerivedParameterTables[1].ReferencedDocumentDescriptor.ReferencedFileDescriptor;
                if (newName != "")
                    fd.ReferencedFileDescriptor.ReplaceReference(newName);
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
                    newName = InvDoc.u.OFD(System.IO.Path.GetDirectoryName(doc.FullDocumentName), fn: doc.FullDocumentName);
                    doc.ReferencedDocumentDescriptors[i + 1].ReferencedFileDescriptor.ReplaceReference(newName);
                }
            }
            doc.Update2(false);
        }

        public static string findDocumentName(DrawingDocument doc)
        {
            FileDescriptor fd;
            fd = doc.ReferencedDocumentDescriptors[1].ReferencedFileDescriptor;
            var bp = file.p(doc.FullFileName);
            var bn = file.name(doc.FullFileName);
            var fn = fd.FullFileName;
            var p = file.p(fn);
            var n = file.name(fn);
            if (bp == p) return null;
            string nn = "";
            return bn;
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
                oldName = newName.Substring(newName.LastIndexOf('\\') + 1, newName.Length - newName.LastIndexOf('\\') - 5);
                PartComponentDefinition compDef = doc.ComponentDefinition;
                SketchLine sl = null;
                List<WorkPoint> wps = selectFromSketch(doc, ref sl, ps1);
                XMLDoc xmldoc = null;
                string n = doc.path() + "\\" + oldName + ".xml";
                if (System.IO.File.Exists(n))
                    xmldoc = xmldoc ?? new XMLDoc(n, "Property");
                iFeature feature = null;
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
                                    ((iFeatureParameterInput)input).Value = double.Parse(at.Value) / 10;
                                }
                            }
                        }
                    }

                    Edge edge1 = null; Edge edge2 = null;
                    feature = compDef.Features.iFeatures.Add(ifd);
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
                        else getImateEdge(oldName, ref edge1, ref edge2, feature);
                        if (asm == null)
                            addImate(oldName, edge1, edge2, (Document)doc, feature);
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
                }
                if (feature != null) return feature;
            }
            return null;
        }

        public void addImate(string name, Edge edge1, Edge edge2, Document doc, iFeature feature)
        {
            IMate im = new IMate(doc);
            XMLDoc xmldoc = new XMLDoc(I.p() + @"\AutoFastener.xml", "Fasteners");
            var item = xmldoc.Doc.Descendants("IMate").FirstOrDefault(e => e.Value == name);
            if (item != null)
            {
                Double d = double.Parse(item.Attribute("Offset").Value.Replace('.', ','));
                im.imate(item.Attribute("name").Value, name, d, 0, null, doc, edge1, edge2, true);
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
            IMate im = new IMate(doc, null);
            XMLDoc xmldoc = new XMLDoc(I.p() + @"\AutoFastener.xml", "Fasteners");
            var item = xmldoc.Doc.Descendants("IMate").FirstOrDefault(e => e.Value == name);
            if (item != null)
            {
                Double d = double.Parse(item.Attribute("Offset").Value.Replace('.', ','));
                im.imate(item.Attribute("name").Value, name, d, 0, null, doc, edge1, edge2, true);
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

        public void getImateEdge(string name, ref Edge edge1, ref Edge edge2, iFeature feature)
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
                        if ((int)(cyl.Radius * 1000) == 300) edge1 = f.Edges[1];
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
                            dist = (int)(pl.DistanceTo(((Circle)edge1.Geometry).Center) * 1000);
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
            for (int i = 0; i < strs.Length; i += 2)
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
            for (int i = 0; i < names.Length; i += 2)
            {
                ComponentOccurrence occ = findOcc(names[i], acd.Occurrences);
                Parts.removeOcc(acd, occ, false);
                if (occ != null && fileExist(names[i + 1]))
                    occ.Replace(name, true);
                else if (occ != null && names[i + 1].StartsWith("v3#"))
                {
                    occ.Replace(findInContentCenter(names[i + 1]), true);
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

        public static List<ComponentOccurrence> findOccs(AssemblyComponentDefinition acd, string name, string num)
        {
            List<ComponentOccurrence> occs = new List<ComponentOccurrence>();
            string[] nums;
            if (num.IndexOf(";") != -1)
            {
                if (num.EndsWith(";")) num = num.Remove(num.Length - 2);
                nums = num.Split(';');
            }
            else
            {
                nums = new string[] { num };
            }
            foreach (ComponentOccurrence occ in acd.Occurrences)
            {
                foreach (var item in nums)
                {
                    if (occ.Name.ToLower() == name.ToLower() + ":" + item.ToLower()) occs.Add(occ);
                }
            }
            return occs;
        }

        public static void removeReplace(AssemblyComponentDefinition acd, XElement el, string path = null)
        {
            foreach (var row in el.Elements())
            {
                string name = row.Attribute("name").Value;
                string num = MyXML.getAtt(row, "num");
                if (path != null) name = path + name;
                List<ComponentOccurrence> occs;
                if (num != "") occs = findOccs(acd, name, num);
                else occs = findOccs(acd, name);
                if (row.Name.ToString().ToLower() == "add")
                {
                    int cnt = 1; string n = name;
                    string cs = MyXML.getAtt(row, "Count");
                    if (cs != "")
                    {
                        cnt = int.Parse(cs);
                    }
                    if (n.StartsWith("v3#"))
                    {
                        n = ContentOp.memberForPlace(n);
                    }
                    for (int i = 0; i < cnt; i++)
                    {
                        if (!ut.occsContaints(acd.Occurrences, n, cnt))
                            acd.Occurrences.AddUsingiMates(n);
                    }
                }
                if (row.Name.ToString().ToLower() == "delete")
                {
                    foreach (ComponentOccurrence occ in occs)
                    {
                        if (row.Attribute("r") != null)
                            Parts.removeOcc(acd, occ, true);
                        else
                            occ.Delete();
                    }
                }
                else if (row.Name.ToString().ToLower() == "replace")
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
            XMLDoc xmlDoc = new XMLDoc(I.p() + @"\ContentCenter.xml", "Content");
            createNodeContent(cc.TreeViewTopNode, xmlDoc);
            xmlDoc.save();
        }

        private void createNodeContent(ContentTreeViewNode basenode, XMLDoc xmlDoc)
        {
            Dictionary<string, string> vals = new Dictionary<string, string>(5);
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
            XMLDoc xmlDoc = new XMLDoc(I.p() + @"\Детали.xml", "Head");
            string memberfilename = "";
            string failuremessage;
            MemberManagerErrorsEnum err;
            ContentCenter cc = invApp.ContentCenter;
            string ccp = invApp.DesignProjectManager.ActiveDesignProject.ContentCenterPath;
            foreach (var item in xmlDoc.Doc.Root.Descendants())
            {
                ContentFamily fam = (ContentFamily)cc.GetContentObject(item.Value);
                string folder = file.concateDir(ccp, @"ru-RU\" + fam.MemberDirectory);
                file.removeFiles(folder);
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
                    createTableRowContent(fam, xmlDoc.addXElement(elem, "family", vals));
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
                    elem.Add(new XElement("TableRow", new XAttribute("Id", item.ContentIdentifier), new XAttribute("Description", item[c].Value), new XAttribute("Extra", a.DisplayHeading + "=" + item[a].Value)));
                else if (c != null)
                    elem.Add(new XElement("TableRow", new XAttribute("Id", item.ContentIdentifier), new XAttribute("Description", item[c].Value)));
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

        public static void selectOcc(ComponentOccurrence occ, SelectSet ss)
        {
            if (occ == null) return;
            var occs = I.COC();
            foreach (AssemblyConstraint constr in occ.Constraints)
            {
                if (constr.OccurrenceOne != null && constr.OccurrenceTwo != null && constr.OccurrenceTwo.Equals(occ) &&
                    constr.OccurrenceOne.ReferencedDocumentDescriptor.FullDocumentName.IndexOf(@"Content Center Files\") != -1)
                    occs.Add(constr.OccurrenceOne);
            }
            foreach (ComponentOccurrence item in occs)
            {
                selectOcc(item, ss);
                ss.SelectMultiple(occs);
            }
        }

        private static ComponentOccurrence findOcc(string name, ComponentOccurrences co)
        {
            foreach (ComponentOccurrence cop in co)
            {
                string copname = cop.Name.Substring(0, cop.Name.IndexOf(':'));
                if (copname == name.Substring(0, name.Length - 4)) return cop;
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
                    if (data.attr.ContainsKey("base")) { fileExist(data.attr["base"]); baseDoc = name; };
                    doc = (!fileExist(data.val)) ?
                    (data.attr.ContainsKey("base")) ? (Document)derivedDoc(baseDoc, data.attr["d"]) : (Document)addPrtDoc() :
                    (Document)openPrtDoc(name);
                    break;
                case "asm":
                    if (data.attr.ContainsKey("base")) { fileExist(data.attr["base"]); baseDoc = name; };
                    doc = (!fileExist(data.val)) ?
                        (data.attr.ContainsKey("base")) ? (Document)copyAsm(path + data.val, baseDoc) : (Document)addAsmDoc() :
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
                            addProp(doc, "Type", item.Value);
                            break;
                        case "decnumber":
                            addProp(doc, "DecNumber", item.Value);
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
                        if (data.Attribute("base") != null) { val = data.Attribute("base").Value; }
                        else if (data.Parent.Attribute("base") != null) { val = data.Parent.Attribute("base").Value; }
                        if (val != "") { fileExist(val); baseDoc = name; };
                        doc = (!fileExist(data.Value)) ?
                        (val != "") ? (Document)derivedDoc(baseDoc, data.Attribute("d").Value) : (Document)addPrtDoc() :
                        (Document)openPrtDoc(name);
                        break;
                    case "asm":
                        if (data.Attribute("base") != null) { val = data.Attribute("base").Value; }
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
                eds = (cond.StartsWith("<")) ?
                    f.Edges.OfType<Edge>().Where(e => ut.getLenght(e) < val) :
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
            return doc.FullFileName.Substring(0, doc.FullFileName.LastIndexOf('\\'));
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
            return doc.FullFileName.Substring(doc.FullFileName.LastIndexOf('\\') + 1, doc.FullFileName.Length - 1 - doc.FullFileName.LastIndexOf('\\') - 4);
        }
    }
}
