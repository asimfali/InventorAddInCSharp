using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Inventor;
using InvDoc;
using System.Xml.Linq;
using InterfaceDll;
using System.Text.RegularExpressions;

namespace InvAddIn
{
    class Site
    {
        List<SiteBOM> sites = new List<SiteBOM>();
        public string ty = null;
        MyXML xml;
        XElement el;
        Document doc;
        public bool vis = false;
        string fn = "site_files.txt";
        List<string> fs = new List<string>();
        public List<Document> docs = new List<Document>();
        public Site(Document doc)
        {
            this.fn = file.p(doc.FullDocumentName) + fn;
            this.doc = doc;
        }
        public void act(string t = "xml")
        {
            this.ty = t;
            if (ty == "listWrite")
            {
                createList();
                writeFile(fn);
                return;
            }
            else if (ty == "OpenFiles")
            {
                openFiles(fn);
            }
            if (ty == "xml" || ty == "listRead")
            {
                string fn = file.p(doc.FullDocumentName) + "kd.a.tepl";
                if (file.check(fn)) System.IO.File.Delete(fn);
                xmlData(fn);
            }
            else if (ty == "model" || ty == "model2" || ty == "3dPDF")
            {
                run();
            }
        }
        public void writeFile(string fn)
        {
            using (System.IO.StreamWriter f = new System.IO.StreamWriter(fn))
            {
                foreach (var item in fs)
                {
                    f.WriteLine(item);
                }
            }
        }
        public void openFiles(string fn)
        {
            fs = System.IO.File.ReadAllLines(fn).ToList();
            foreach (var item in fs)
            {
                docs.Add(I.open(item, vis, false));
            }
        }
        public void createList(string pattern = "00.000")
        {
            var pr = I.app.DesignProjectManager.ActiveDesignProject;
            var p = pr.WorkspacePath;
            foreach (var item in file.getFiles(p, ".iam", System.IO.SearchOption.AllDirectories))
            {
                if (item.IndexOf("OldVersions") != -1) continue;
                var n = file.name(item);
                if (n.IndexOf(pattern) != -1)
                    fs.Add(item);
            }
        }
        public void xmlData(string fn)
        {
            xml = new MyXML(fn, "root");
            el = xml.getRoot();
            var tmps = run();
            foreach (var item in tmps)
            {
                el.Add(item.Elements());
            }
            xml.save(fn);
        }
        public List<XElement> run(bool vis = true)
        {
            List<XElement> tmps = new List<XElement>();
            IEnumerable<Document> docs;
            if (this.docs.Count > 0) docs = this.docs;
            else docs = I.app.Documents.VisibleDocuments.OfType<Document>();
            foreach (Document doc in docs)
            {
                if (doc == null) continue;
                if (ty == "xml" || ty == "listRead")
                {
                    if (doc.DocumentType != DocumentTypeEnum.kAssemblyDocumentObject) continue;
                    XElement tmp = new XElement("root");
                    var bom = new SiteBOM(doc, tmp);
                    tmps.Add(tmp);
                }
                if (ty == "model")
                {
                    createModel(doc);
                }
                if (ty == "model2")
                {
                    createModel(doc, true);
                }
                if (ty == "3dPDF")
                {
                    var p = file.p(doc.FullDocumentName);
                    InvAddIn.CreateComponent.Exp3DPDF(p);
                    break;
                }
            }
            return tmps;
        }
        public static string getModel(Document doc)
        {
            string t = u.getPropValue(doc, "Part Number");
            t = t.Replace('Е', 'E');
            t = t.Replace('А', 'A');
            string n = u.getPropValue(doc, "Comments");
            string r1 = getRegex(t, @"-([^\.]*)\.+?", 1);
            if (r1 == "")
                r1 = getRegex(t, @"(\w*)", 1);
            string r2 = getRegex(n, @"(\d*)", 1);
            Dictionary<char, string> repl = new Dictionary<char, string>();
            Dictionary<char, string> repl1 = new Dictionary<char, string>();
            repl['M'] = "0"; repl['W'] = "1"; repl['Ф'] = "0";
            string model = getModelProp(doc);
            if (model != "") return model;
            if (r1.StartsWith("П"))
                return "КЭВ-" + r2 + r1;
            foreach (KeyValuePair<char, string> item in repl)
            {
                int ind = int.Parse(item.Value);
                if (r1[ind] == item.Key)
                {
                    if (item.Key == 'M' || item.Key == 'Ф')
                        repl1[item.Key] = item.Key + getVent(doc, "о_в", a => getRegex(a, @"(\d\d)", 1));
                    else if (item.Key == 'W')
                        repl1[item.Key] = item.Key + getVent(doc, "ТЕРМА", a => getRegex(a, @"(\d) ряд", 1));
                }
            }
            r1 = "";
            foreach (KeyValuePair<char, string> item in repl1)
            {
                r1 += item.Value;
            }
            if (r1.StartsWith("Ф")) r1 += "ПМ";
            return "КЭВ-" + r2 + r1;
        }
        public static string getVent(Document doc, string find, Func<string, string> f)
        {
            var tmp = "";
            foreach (Document item in doc.ReferencedDocuments)
            {
                var n = item.FullDocumentName;
                n = file.name(n);
                if (n.StartsWith(find))
                {
                    tmp = f(n);
                    if (tmp.Length > 1){
                        double v = double.Parse(tmp) / 10;
                        tmp = v.ToString();
                    }
                }
            }
            return tmp;
        }
        public static string getModelProp(Document doc)
        {
            string tmp = "";
            foreach (Document item in doc.ReferencedDocuments)
            {
                string n = u.getPropValue(item, "model");
                if (n != "") 
                    return n;
            }
            return tmp;
        }
        public void createModel(Document doc, bool copy = false)
        {
            var m = getModel(doc);
            if (copy)
            {
                StringBuilder m1 = new StringBuilder(m);
                MyForm F = new MyForm("ChangeNumberInterface.xml", "Изменить");
                F.f.ShowDialog();
                var num = int.Parse(F.cbs[0].Text);
                var val = char.Parse(F.cbs[1].Text);
                m1[m1.Length - num] = val;
                m = $"{m}, {m1}";
            }
            u.addProp(doc, "model", m);
            doc.Save2(true);
        }
        public static string getRegex(string str, string pat, int c)
        {
            Regex reg = new Regex(pat);
            var m = reg.Match(str);
            if (m != null && m.Groups.Count >= c)
            {
                return m.Groups[c].Value;
            }
            return null; 
        }
    }
    class SiteBOM
    {
        BOM bom;
        AssemblyDocument asm;
        BOMView bv;

        Dictionary<string, string> propsNames;
        public SiteBOM(Document doc, XElement el)
        {
            propsNames = new Dictionary<string, string> { { "PN", "Part Number" },
                { "Desc", "Description" },
                { "DXF", "Revision Number"},
                { "Model", "model"} };
            var dics = u.getProps(doc, propsNames);
            set(doc, null, ref dics);
            var tmp = MyXML.addXElement("row", dics);
            el.Add(tmp);
            //el = setProps(doc, el);
            asm = doc as AssemblyDocument;
            bom = asm.ComponentDefinition.BOM;
            bom.StructuredViewEnabled = true;
            bom.StructuredViewFirstLevelOnly = false;
            bv = u.get<BOMView>(bom.BOMViews, e => e.Name == "Структурированный");
            addRows(bv.BOMRows, tmp);
        }
        public void addRows(BOMRowsEnumerator rowsEnum, XElement el)
        {
            foreach (BOMRow row in rowsEnum)
            {
                var docum = row.ComponentDefinitions[1].Document as Document;
                var dics = u.getProps(docum, propsNames);
                trimDXF(dics);
                set(docum, row,ref dics);
                var tmp = MyXML.addXElement("row", dics);
                if (row.ChildRows != null)
                {
                    addRows(row.ChildRows, tmp);
                }
                el.Add(tmp);
            }
        }
        public void set(Document doc, BOMRow row, ref Dictionary<string, string> dics)
        {
            if (row != null)
                dics["Count"] = row.ItemQuantity.ToString();
            else dics["Count"] = "1";
            PDFOp pdf = new PDFOp();
            var tmpPDF = pdf.getFN(doc);
            if (tmpPDF != "") dics["PDF"] = tmpPDF;
        }
        public void trimDXF(Dictionary<string, string> dic)
        {
            if (dic.ContainsKey("DXF"))
            {
                var v = dic["DXF"];
                var i = v.IndexOf("  ");
                if (i != -1)
                {
                    dic["DXF"] = v.Substring(0, i);

                }
            }
        }
    }

    class Gltf
    {
        Document doc;
        TranslatorAddIn obj;
        TranslationContext context = I.objs.CreateTranslationContext();
        NameValueMap nvm = I.objs.CreateNameValueMap();
        DataMedium dm;
        string objFN = "", path, fn;
        //Arctron.Obj2Gltf.Converter conv;
        //Arctron.Obj2Gltf.GltfOptions opt;
        public Gltf(Document doc)
        {
            this.doc = doc;
            obj = (TranslatorAddIn)I.app.ApplicationAddIns.ItemById["{F539FB09-FC01-4260-A429-1818B14D6BAC}"];
            context.Type = IOMechanismEnum.kFileBrowseIOMechanism;
            dm = I.objs.CreateDataMedium();
            nvm.Add("ExportUnits", 1);
            nvm.Add("ExportFileStructure", 0);
            nvm.Add("Resolution", 1);
            nvm.Add("SurfaceDeviation", 60);
            nvm.Add("NormalDeviation", 14);
            nvm.Add("MaxEdgeLength", 100);
            nvm.Add("AspectRatio", 21.5);
            //opt = new Arctron.Obj2Gltf.GltfOptions();
            //opt.Binary = false;
            //opt.WithBatchTable = false;
            objFN = getObj();
            //opt.Name = this.fn + ".gltf";
            //getGltf();
        }
        public string getObj()
        {
            string fn = doc.FullFileName;
            path = file.p(fn) + "gltf";
            file.dir(path);
            this.fn = file.name(fn);
            dm.FileName = path + "\\" + this.fn + ".obj";
            if (!System.IO.File.Exists(dm.FileName))
                obj.SaveCopyAs(doc, context, nvm, dm);
            return dm.FileName;
        }
        //public void getGltf()
        //{
        //    conv = new Arctron.Obj2Gltf.Converter(objFN, opt);
        //    conv.WriteFile(path + "\\" + this.fn + ".gltf");
        //}
    }
}
