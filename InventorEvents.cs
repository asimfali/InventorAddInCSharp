using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using InvDoc;
using Inventor;
using System.Xml.Linq;
using System.Reflection;
using addin = Macros.StandardAddInServer;

namespace InvAddIn
{
    public class InventorEvents
    {
        public CommandManager cmd;
        public InteractionEvents iEv;
        public UserInputEvents uInputEv;
        public UserInterfaceEvents uInterfaceEv;
        public ApplicationEvents appEv;
        public SelectEvents selEv;
        public SelectionFilterEnum filter;

        public InventorEvents()
        {
            cmd = addin.m_inventorApplication.CommandManager;

            uInterfaceEv = addin.m_inventorApplication.UserInterfaceManager.UserInterfaceEvents;
            appEv = addin.m_inventorApplication.ApplicationEvents;
        }
        public void createEvents()
        {
            iEv = cmd.CreateInteractionEvents();
            selEv = iEv.SelectEvents;
        }
    }

    public class InventorPRoperties
    {
        private Document doc;
        protected Property pr;
        public List<Property> props = new List<Property>();
        public string[] names;
        public Property this[int index]
        {
            get { return props[index]; }
            set { props.Insert(index, value); }
        }

        public Property this[string name]
        {
            get { return props.FirstOrDefault(p => p.Name == name); }
        }

        public Document Doc
        {
            get { return doc; }
            set { doc = value; }
        }

        public void get(string[] names)
        {
            foreach (var item in names)
            {
                get(item);
            }
        }

        public InventorPRoperties(Document doc,string [] names)
        {
            this.doc = doc;
            this.names = names;
            foreach (var item in names)
            {
                get(item);
            }
        }

        public InventorPRoperties()
        {
        }

        public T get<T>(Property pr)
        {
            if (pr == null) return default(T);
            return (T)pr.Value;
        }

        public Property get(string name)
        {

            pr = u.getProp(doc, name);
            if (pr != null) props.Add(pr);
            else u.addProp(doc, name, "");
            return pr;
//             foreach (PropertySet item in doc.PropertySets)
//             {
//                 pr = item.OfType<Property>().FirstOrDefault(p => p.Name.ToLower() == name.ToLower());
//                 if (pr != null) { props.Add(pr); return pr; }
//             }
//             props.Add(null);
//             return null;
        }

        public List<Property> gets(string[] names)
        {
            foreach (var item in names)
            {
                get(item);
            }
            return props;
        }

        public bool change<T>(Property p, T val)
        {
            T v = get<T>(p);
            if (v != null && v.Equals(val)) return false;
            p.Value = val;
            return true;
        }

        public bool changes(List<string> vals, int end, int start = 0)
        {
            bool flag = false;
            if (vals.Count != props.Count) return flag;
            for (int i = start; i < end; i++)
            {
                if (change<string>(props[i], vals[i])) flag = true;;
            }
            return flag;
        }

        public Property add<T>(string name, T val)
        {
            pr = get(name);
            if (pr == null||change<T>(pr, val))
            {
                pr = doc.PropertySets["Inventor User Defined Properties"].Add(val, name);
                doc.Update2();
                if (names.Contains(name))
                {
                    props[names.ToList().IndexOf(name)] = pr;
                }
                else props.Add(pr);
            }
            return pr;
        }
    }

    public class Materials
    {
        List<Material> materials = new List<Material>();
        public Materials()
        {
        }
        public void add(Material mat)
        {
            materials.Add(mat);
        }
        public void remove(Material mat)
        {
            materials.Remove(mat);
        }
        public Material this[object doc]
        {
            get { return materials.FirstOrDefault(m => m.Doc.Equals(doc)); }
        }
    }

    public class Mass
    {
        double m;
        public Mass(_Document doc)
        {
            if (doc.DocumentType == DocumentTypeEnum.kPartDocumentObject)
            {
                Property p = InvDoc.u.getProp(doc, "Mass");
                m = InvDoc.u.convToDouble(p.Value.ToString());
                InvDoc.u.addProp(doc, "Mass&", conv());
            }
        }
        private string conv()
        {
            m /= 1000;
            return m < 0.05 ? m.ToString("0.000"): m.ToString("0.00");
        }
    }

    public class Material : InventorPRoperties, IDisposable
    {
        bool changed = false;
        InventorEvents ev = new InventorEvents();
        static XDocument xDoc;
        Material mat;
        Materials materials = new Materials();
        SheetMetalComponentDefinition smcd;
        PartDocument pDoc;
        public Material(Document doc, string [] arr /*{ "MaterialUpLine", "MaterialDownLine", "MaterialLine", "MaterialCenter" }*/): base(doc, arr)
        {
            pDoc = doc as PartDocument;
            if (pDoc == null) return;
            string path = doc.pathUtil();
            smcd = pDoc.ComponentDefinition as SheetMetalComponentDefinition;
            if (smcd == null) return;
            xDoc = xDoc ?? XDocument.Load(@"C:\ProgramData\Autodesk\Inventor Addins\Material.xml");
            if (System.IO.File.Exists(path + "\\Material.xml"))
                xDoc = XDocument.Load(path + "\\Material.xml");
        }
        public Material()
        {
            ev.appEv.OnOpenDocument += appEv_OnOpenDocument;
            ev.appEv.OnSaveDocument += appEv_OnSaveDocument;
            ev.appEv.OnCloseDocument += appEv_OnCloseDocument;
        }

        void appEv_OnCloseDocument(_Document DocumentObject, string FullDocumentName, EventTimingEnum BeforeOrAfter, NameValueMap Context, out HandlingCodeEnum HandlingCode)
        {
            HandlingCode = HandlingCodeEnum.kEventHandled;
            if (BeforeOrAfter != EventTimingEnum.kBefore) return;
            mat = materials[DocumentObject];
            materials.remove(mat);
            if (DocumentObject.DocumentType == DocumentTypeEnum.kDrawingDocumentObject)
            {
                gabs.unset();
                ArrDim.unset();
            }
        }

        void appEv_OnSaveDocument(_Document DocumentObject, EventTimingEnum BeforeOrAfter, NameValueMap Context, out HandlingCodeEnum HandlingCode)
        {
            HandlingCode = HandlingCodeEnum.kEventHandled;
            if (BeforeOrAfter != EventTimingEnum.kBefore) return;
            mat = materials[DocumentObject];
            Mass m = new Mass(DocumentObject);
            mat.check(mat.smcd);
            //if (mat.changed) { DocumentObject.Save2(); mat.changed = false; }
            //materials.remove(mat);
        }

        void appEv_OnOpenDocument(_Document DocumentObject, string FullDocumentName, EventTimingEnum BeforeOrAfter, NameValueMap Context, out HandlingCodeEnum HandlingCode)
        {
            HandlingCode = HandlingCodeEnum.kEventHandled;
            if (BeforeOrAfter != EventTimingEnum.kAfter) return;
            mat = new Material(DocumentObject, new string[] { "MaterialUpLine", "MaterialDownLine", "MaterialLine", "MaterialCenter" });
            materials.add(mat);
        }

        bool check(SheetMetalComponentDefinition smcd1)
        {
            if (xDoc == null) return false;
            double t = u.convToDouble(smcd1.Thickness.Value.ToString(),10, 3);
            string matName = smcd1.Material.Name;
            XElement el = xDoc.Root.Descendants("Material").FirstOrDefault(e => e.FirstAttribute.Value == matName);
            if (el == null) return false;
            //el.Element("MaterialUpLine").Value = el.Element("MaterialUpLine").Value.Replace("t", t.ToString("0.0", System.Globalization.CultureInfo.CreateSpecificCulture("da-DK")));
            List<string> vals = new List<string>() { el.Element("MaterialUpLine").Value.Replace("t", t.ToString("0.0", System.Globalization.CultureInfo.CreateSpecificCulture("da-DK"))), el.Element("MaterialDownLine").Value, el.Element("MaterialLine").Value, el.Element("MaterialCenter").Value };
            if (base[0] != null)
            {
                changed = changes(vals, 3);
            }
            return changed;
        }

        public void Dispose()
        {
            ev.appEv.OnOpenDocument -= appEv_OnOpenDocument;
            ev.appEv.OnSaveDocument -= appEv_OnSaveDocument;
            ev.appEv.OnCloseDocument -= appEv_OnCloseDocument;
        }
    }
}
