using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Inventor;

namespace InvAddIn
{
    public class VariableData
    {
        protected Document doc;
        protected AttributeSet attSet;
        protected AttributeSets attSets;
        protected Inventor.Attribute attr;
        public DataGridView _dgv { get; set; }
        //public Document Doc { get; set; }

        public VariableData(Document d) { doc = d; }

        public void AttribAdd<T>(string name, T val, ValueTypeEnum VT)
        {
            attSets = doc.AttributeSets;
            if (attSets.NameIsUsed[name]) AttribDelete(name);
            attSet = attSets.Add(name); attSet.Add(name, VT, val);
            attSet = null;
        }

        public void AttribAdd<T,TVal>(T ob ,string name, TVal val, ValueTypeEnum VT)
        {
            AttributeSets attSets = InvDoc.Reflect.getProp<T, AttributeSets>(ob, "AttributeSets");
                if (attSets.NameIsUsed[name]) AttribDelete(name);
                attSet = attSets.Add(name); attSet.Add(name, VT, val);
                attSet = null;
        }

        public AttributeSet AttribSetAdd(object ob, string nameAttSet)
        {
            PartFeature pf;   
            if (ob is PartFeature)
            {
                pf = ob as PartFeature;
                attSets = pf.AttributeSets;
                if (attSets.NameIsUsed[nameAttSet]) AttribDelete(nameAttSet);
                attSet = attSets.Add(nameAttSet);
                return attSet;
            }
            return null;
        }

        public AttributeSet AttribSetAdd(SketchPoint ob, string nameAttSet)
        {
            attSets = ob.AttributeSets;
            if (attSets.NameIsUsed[nameAttSet]) AttribDelete(nameAttSet);
            attSet = attSets.Add(nameAttSet);
            return attSet;
        }

        public void AttribAdd<T>(AttributeSet attSet, string name, T val, ValueTypeEnum VT)
        {
            attSet.Add(name, VT, val);
        }

        public void AttribAdd<T>(string name, List<T> vals, ValueTypeEnum VT)
        {
            attSet = doc.AttributeSets.Add(name);
            for (int i = 0; i < vals.Count; i++)
            {
                attSet.Add(name + i, VT, vals[i]); 
            }
            attSet = null;
        }

        public void AttribDelete(string name)
        {
            try
            {
                attSets = doc.AttributeSets;
                attSet = attSets[name];
                foreach (Inventor.Attribute att in attSet)
                {
                    att.Delete();
                }
                attSets[name].Delete();
            }
            catch
            {
            }
        }

        public void AttribsDelete(string name)
        {
            var IE = doc.AttributeSets.OfType<AttributeSet>().Where(a => a.Name.StartsWith(name));
            foreach (AttributeSet attSet in IE)
            {
                AttribDelete(attSet.Name);
            }
        }

        public Inventor.Attribute getAttrib(string name)
        {
            var ie = doc.AttributeSets.OfType<AttributeSet>().Where(a => a.Name.StartsWith(name));
            foreach (var item in ie)
            {
                if (item.Count != 0) return item[1];
            }
            return null;
        }

        public Inventor.Attribute getAttrib<T>(T ob, string name)
        {
            AttributeSets attSets = InvDoc.Reflect.getProp<T,AttributeSets>(ob, "AttributeSets");          
                var ie = attSets.OfType<AttributeSet>().Where(a => a.Name.StartsWith(name));
                foreach (var item in ie)
                {
                    if (item.Count != 0) return item[1];
                }
            return null;
        }

        public Inventor.AttributeSet getAttrSet<T>(T ob, string name)
        {
           AttributeSets attSets = InvDoc.Reflect.getProp<T,AttributeSets>(ob, "AttributeSets");
           var ie = attSets.OfType<AttributeSet>().Where(a => a.Name.StartsWith(name));
                return ie.FirstOrDefault(); 
        }

        public Inventor.AttributeSet getAttrSet(SketchPoint ob, string name)
        {
            var ie = ob.AttributeSets.OfType<AttributeSet>().Where(a => a.Name.StartsWith(name));
            return ie.FirstOrDefault();
        }

        public Inventor.Attribute getAttrib(AttributeSet attSet, string name)
        {
            return attSet[name];
        }

    }

    public class style
    {
        public Font f { get; set; }
        public int row { get; set; }
        public int col { get; set; }
        style(Font fon, int r, int c)
        {
            f = fon;
            row = r;
            col = c;
        }
    }

    public class VariableDataForSpec : VariableData
    {
        public string rowName, cellName, val = "";
        public Dictionary<string, string> dic { get; set; }
        public List<string> lst = new List<string>();
        public int i = 0, j = 0;

        public VariableDataForSpec(Document d) : base(d) { }

        public void AttribDelete(Dictionary<string, string> d)
        {
            foreach (var key in d)
            {
                AttribsDelete(key.Key);
            }
        }

        public void AttribAdd(string name,string [] vals)
        {
            try
            {
                for (int i = 0; i < vals.Length; i++)
                {
                    AttribAdd<string>(name + '_' + i, vals[i], ValueTypeEnum.kStringType);
                }
            }
            catch (Exception ex) { MessageBox.Show(ex.ToString());}
        }

        public List<string> AttribLoad(string name)
        {
            List<string> vals = new List<string>();
           var ie = doc.AttributeSets.OfType<AttributeSet>().Where(a => a.Name.StartsWith(name));
           foreach (var item in ie)
           {
              if (item.Count == 0) {AttribDelete(name); return vals;}
              vals.Add(item[1].Value.ToString()); 
           }
           return vals;
        }

        //public void AttribAdd(TableInv tbl, Dictionary<string, string> d)
        //{           
        //    dic = d;
        //    foreach (var key in dic)
        //    {
        //        AttribDelete(key.Key);
        //    }
        //    foreach (var row in tbl.rows)
        //    {
        //        string name = "Row"; int count = 0;
        //        rowName = name + count;
        //        AttributeSetsEnumerator ase = doc.AttributeManager.FindAttributeSets(rowName);
        //        foreach (AttributeSet item in ase)
        //        {
        //            item.AddAttributes
        //        }
        //        attSet = doc.AttributeSets.Add(rowName);
        //        foreach (var cel in row.cells)
        //        {
        //            string val1 = "";
        //                if (cel.underline) val1 += "{u}";
        //                if (cel.bold) val1 += "{b}";
        //                if (cel.italic) val1 += "{i}";
        //            val = val + val1 + cel.value + ';';
        //        }
        //        val += '|';
        //        val += row.group; val += row.varRow;
        //        val += '|';
        //        if (row.varRows.Count != 0)
        //        {
        //            foreach (var item in row.varRows)
        //            {
        //                val += item.rowInd; 
        //            }
        //        }
        //        attr = attSet.Add("String", ValueTypeEnum.kStringType, val);
        //        val = "";
        //    }

        //    foreach (var key in dic)
        //    {
        //        AttribAdd<string>(key.Key, key.Value, ValueTypeEnum.kStringType);
        //    }
        //}

        public void AttribAdd(DataGridView dgv, Dictionary<string,string> d)
        {
            _dgv = dgv;
            dic = d;

            foreach (var key in dic)
	            {
		            AttribsDelete(key.Key); 
	            }

            foreach (DataGridViewRow row in _dgv.Rows)
            {
                string name = "Row"; int count = 0;
                if (row.Cells[4].Value != null)
                {
                    string sVal = row.Cells[4].Value.ToString();
                    if (sVal == "Переменные данные для исполнений:")
                    {
                        name = "Isp" + count + "_";
                    }
                    if (name.StartsWith("Isp") && sVal[sVal.Length-3] == '-')
                    {
                        count++;
                        name = "Isp" + count + "_";
                    }
                    rowName = name + i;
                    i++;
                    attSet = doc.AttributeSets.Add(rowName);
                    if (i == _dgv.RowCount)
                        break;
                    foreach (DataGridViewCell cel in row.Cells)
                    {
                        string val1 = ""; 
                        if (cel.Style.Font != null)
                        {
                            if (cel.Style.Font.Underline) val1 += "{u}";
                            if (cel.Style.Font.Bold) val1 += "{b}";
                            if (cel.Style.Font.Italic) val1 += "{i}";
                        }
                        val = val + val1 + (string)cel.FormattedValue + ';';
                    }
                    attr = attSet.Add("String", ValueTypeEnum.kStringType, val);
                    val = "";
                }
                //j = 0;
            }
            foreach (var key in dic)
            {
                AttribAdd<string>(key.Key, key.Value, ValueTypeEnum.kStringType);
            }
        }

        //public TableInv load(string start)
        //{
        //    TableInv tbl = new TableInv();
        //    foreach (Inventor.AttributeSet attrset in doc.AttributeSets)
        //    {
        //        if (attrset.Name.StartsWith(start) && attrset.Count != 0)
        //        {
        //            i += 1;
        //            if (i == doc.AttributeSets.Count)
        //                break;
        //            string[] data = attrset[1].Value.ToString().Split('|');


        //            string[] str = attrset[1].Value.ToString().Split(';');
        //            //Array.Resize(ref str, str.Length - 1);
        //            string[] stl = new string[9];
        //            for (int j = 0; j < str.Length; j++)
        //            {
        //                if (str[j].StartsWith("{"))
        //                {
        //                    int ind = str[j].LastIndexOf('}');
        //                    stl[j] = str[j].Substring(0, ind + 1);
        //                    str[j] = str[j].Substring(ind + 1, str[j].Length - ind - 1);
        //                }
        //            }

        //            int num = _dgv.Rows.Add("", str[1], str[2], str[3], str[4], str[5], str[6], str[7]);

        //            for (int j = 0; j < stl.Length; j++)
        //            {
        //                if (stl[j] == "{u}")
        //                {
        //                    _dgv[j, num].Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
        //                    _dgv[j, num].Style.Font = new Font(dgv.Font, FontStyle.Underline);
        //                }
        //                if (stl[j] == "{b}")
        //                {
        //                    _dgv[j, num].Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
        //                    _dgv[j, num].Style.Font = new Font(dgv.Font, FontStyle.Bold);
        //                }
        //                if (stl[j] == "{i}")
        //                {
        //                    _dgv[j, num].Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
        //                    _dgv[j, num].Style.Font = new Font(dgv.Font, FontStyle.Italic);
        //                }
        //            }
        //        }
        //    }
        //}
    }
}
