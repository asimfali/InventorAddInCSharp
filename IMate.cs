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

namespace InvAddIn
{
    public partial class IMate : Form
    {
        Inventor.Application invApp;
        Document Doc;
        CommandManager CmdMgr;
        PartComponentDefinition PartCompDef;
        AssemblyComponentDefinition AsmCompDef;
        InsertiMateDefinition insIMateDef;
        AngleiMateDefinition angliMate;
        MateiMateDefinition iMateDef;
        Edge edge1, edge2;
        public bool flag;
        InteractionEvents intEvts;
        SelectEvents sel;
        KeyboardEvents key;
        List<Edge> edges = new List<Edge>();
        ObjectCollection objs;
        PlanarSketch ps;
        private List<XMLData> data;
        private Transaction tr;
        double offset;
        string str, nameImate;
        //delegate double conv(string str);
        char separator = System.Globalization.CultureInfo.CurrentCulture.NumberFormat.CurrencyDecimalSeparator[0];

        public IMate() { }
        public IMate(Inventor.Document oDoc, string name = "", string typ = "", string offset = "")
        {
            invApp = (Inventor.Application)oDoc.Parent;
            Doc = (Document)oDoc.ActivatedObject;
            if (Doc == null) Doc = oDoc;
            CmdMgr = invApp.CommandManager;
            objs = invApp.TransientObjects.CreateObjectCollection();
            InitializeComponent();
            XML n;
            string filePath = @"C:\ProgramData\Autodesk\Inventor Addins\Imate.xml";
            if (!System.IO.File.Exists(filePath)) MessageBox.Show("Отсутствует файл " + filePath);
            n = new InvDoc.XML(filePath);
            data = new List<XMLData>();
            data = n.ReadXML("Imates");
            foreach (XMLData s in data)
            {
                this.comboBox3.Items.Add(s.val);  
            }
            this.comboBox2.Text = offset;
            this.comboBox1.Text = typ;
            this.comboBox3.Text = name;
            //this.Deactivate += (sender, e) => MessageBox.Show("Deactivated!");
        }

        public void addName(CompositeiMateDefinition iMate, string name)
        {
            if (name != "")
            {
                iMate.Name = name;
                string[] names = new string[1]; names[0] = name;
                iMate.MatchList = names;
            }
        }

        public void addName(InsertiMateDefinition iMate, string name)
        {
            if (name != "")
            {
                iMate.Name = name;
                string[] names = new string[1]; names[0] = name;
                //iMate.MatchList = names;
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            this.Hide();
            switch (this.comboBox1.Text)
            {
                case "Вставка":
                    selEdge(ref edge1, ref edge2, "Выберите первое отверстие", "Выберите второе отверстие", SelectionFilterEnum.kAllCircularEntities);
                break;
                case "Совмещение по оси":
                    selEdge(ref edge1, ref edge2, "Выберите первое отверстие", "Выберите второе отверстие", SelectionFilterEnum.kAllCircularEntities);
                break;
                case "Овал":
                    selOval(ref edge1, "Выберите овал", SelectionFilterEnum.kAllCircularEntities);
                break;
                case "ТЭН-резистор":
                    selEdge(ref edge1, ref edge2, "Выберите первое отверстие", "Выберите отверстие для совмещения по оси", SelectionFilterEnum.kAllCircularEntities);
                break;
                case "ТЭН":
                    selEdge(ref edge1, ref edge2, "Выберите первое отверстие", "Выберите отверстие для совмещения по оси", SelectionFilterEnum.kAllCircularEntities);
                break;
                case "По элементу":
                    selEdge("Выберите элемент", SelectionFilterEnum.kPartFeatureFilter);
                break;
                case "Одиночная вставка":
                    //selEdge(ref edge1, ref edge2, "Выберите отверстие", "", SelectionFilterEnum.kAllCircularEntities);
                selOp(ref edges);
                break;
                case "DIN-рейка":
                    selEdge(ref edge1, ref edge2, "Выберите отверстие", "", SelectionFilterEnum.kAllCircularEntities);
                    selEdge(ref edge2, ref edge1, "Выберите второе отверстие", "", SelectionFilterEnum.kAllCircularEntities);
                break;
                default:
                break;
            }
            Func<string, double> conv = s =>
            {
                char rep = (separator == ',') ? '.' : ',';
                return (s.IndexOf(separator) != -1) ? Convert.ToDouble(s) : Convert.ToDouble(s.Replace(rep, separator));
            };
            if (edge1 != null)
            imate(this.comboBox1.Text, this.comboBox3.Text, conv (this.comboBox2.Text),Doc, edge1, edge2);
            else if (edges.Count != 0)
            {
                foreach (Edge ed in edges)
                {
                   imate(this.comboBox1.Text, this.comboBox3.Text, conv (this.comboBox2.Text),Doc, ed, edge2); 
                }
            }
            IMateBtn.name = comboBox3.Text; IMateBtn.typ = comboBox1.Text; IMateBtn.offset = comboBox2.Text;
        }

        public void selEdge(ref Edge edge1,ref Edge edge2, string promt1, string promt2, SelectionFilterEnum filter)
        {
            if (promt1 != "")
            edge1 = (Edge)CmdMgr.Pick(filter, promt1);
            if (promt2 != "")
            edge2 = (Edge)CmdMgr.Pick(filter, promt2);
        }
        public void selOval(ref Edge edge1, string promt1, SelectionFilterEnum filter)
        {
            if (promt1 != "")
                edge1 = (Edge)CmdMgr.Pick(filter, promt1);
        }
        public void selEdge(string promt, SelectionFilterEnum filter)
        {
            PartFeature pf = CmdMgr.Pick(filter, promt) as PartFeature;
            HoleFeature hf = pf as HoleFeature;
            if (hf == null) return;
            //ICollection<Edge> col = null; 
            foreach (Face item in hf.SideFaces)
            {
                Edge ed = checkBox1.Checked ? item.Edges[2] : item.Edges[1];
                edges.Add(ed);
            }
//             System.Collections.IEnumerator en = PartCompDef.Features.GetEnumerator();
//             en.Reset();
//             while(en.Current != null)
//             {
//                en.
//             }
            PartCompDef = hf.Parent as PartComponentDefinition;
//             for (int i = 1; i < PartCompDef.Features.Count; i++)
// 			{
//                 MirrorFeature mf = PartCompDef.Features[i] as MirrorFeature;
//                 if (mf == null) continue;
//                 HoleFeature holeF = PartCompDef.Features[i - 1] as HoleFeature;
//                 if (hf == null)
//                     holeF = PartCompDef.Features[i - 2] as HoleFeature;
//                 else
//                     continue;
//                 foreach (PartFeature item in mf.ParentFeatures)
//                 {
//                     if (item is HoleFeature && holeF.Equals(item))
//                     {
//                         if (mf.Faces[1].SurfaceType == SurfaceTypeEnum.kCylinderSurface)
//                             edges.Add(mf.Faces[1].Edges[1]);
//                     }
//                 }
// 			}
            foreach (MirrorFeature item in PartCompDef.Features.MirrorFeatures)
            {
                foreach (PartFeature m in item.ParentFeatures)
                {
                    if (m.Equals(hf))
                    {
                        if (item.Faces[1].SurfaceType != SurfaceTypeEnum.kCylinderSurface) continue;
                        Edge ed = checkBox1.Checked ? item.Faces[1].Edges[2] : item.Faces[1].Edges[1];
                        edges.Add(ed);
                    }
                } 
            }
        }

        public void imate(string str, string name, double offset, Document Doc, object edge1, object edge2, bool close = false)
        {
            try
            {
                switch (Doc.DocumentType)
                {
                    case DocumentTypeEnum.kAssemblyDocumentObject:
                        AsmCompDef = ((AssemblyDocument)Doc).ComponentDefinition;
                        break;
                    case DocumentTypeEnum.kPartDocumentObject:
                        PartCompDef = ((PartDocument)Doc).ComponentDefinition;
                        break;
                    default:
                        break;
                }
                //offset = conv(this.comboBox2.Text);
                if (str == "Вставка")
                {
                    switch (Doc.DocumentType)
                    {
                        case DocumentTypeEnum.kAssemblyDocumentObject:
                            insIMateDef = iMate_(edge1, AsmCompDef, offset);
                            objs.Add(insIMateDef);
                            insIMateDef = iMate_(edge2, AsmCompDef, offset);
                            objs.Add(insIMateDef);
                            //if (comboBox3.Text != "")
                            addName(AsmCompDef.iMateDefinitions.AddCompositeiMateDefinition(objs),name);
                            break;
                        case DocumentTypeEnum.kPartDocumentObject:
                            insIMateDef = iMate_(edge1, PartCompDef, offset);
                            objs.Add(insIMateDef);
                            insIMateDef = iMate_(edge2, PartCompDef, offset);
                            objs.Add(insIMateDef);
                            addName(PartCompDef.iMateDefinitions.AddCompositeiMateDefinition(objs),name);
                            break;
                    }
                }
                else if (str == "Совмещение по оси")
                {
                    
                    switch (Doc.DocumentType)
                    {
                        case DocumentTypeEnum.kPartDocumentObject:
                            iMateDef = iMate__(edge1, PartCompDef, offset);
                            objs.Add(iMateDef);
                            iMateDef = iMate__(edge2, PartCompDef, offset);
                            objs.Add(iMateDef);
                            iMateDef = (((Edge)edge1).Faces[1].SurfaceType == SurfaceTypeEnum.kPlaneSurface) ? PartCompDef.iMateDefinitions.AddMateiMateDefinition(((Edge)edge1).Faces[1], offset / 10, BiasPoint: ((Edge)edge1).PointOnEdge) :
                                PartCompDef.iMateDefinitions.AddMateiMateDefinition(((Edge)edge1).Faces[2], offset / 10, BiasPoint: ((Edge)edge1).PointOnEdge);
                            objs.Add(iMateDef);
                            addName(PartCompDef.iMateDefinitions.AddCompositeiMateDefinition(objs),name);
                            break;
                        case DocumentTypeEnum.kAssemblyDocumentObject:
                            iMateDef = iMate__(edge1, AsmCompDef, offset);
                            objs.Add(iMateDef);
                            iMateDef = iMate__(edge2, AsmCompDef, offset);
                            objs.Add(iMateDef);
                            iMateDef = (((Edge)edge1).Faces[1].SurfaceType == SurfaceTypeEnum.kPlaneSurface) ? AsmCompDef.iMateDefinitions.AddMateiMateDefinition(((Edge)edge1).Faces[1], offset / 10, BiasPoint: ((Edge)edge1).PointOnEdge) :
                                AsmCompDef.iMateDefinitions.AddMateiMateDefinition(((Edge)edge1).Faces[2], offset / 10, BiasPoint: ((Edge)edge1).PointOnEdge);
                            objs.Add(iMateDef);
                            addName(AsmCompDef.iMateDefinitions.AddCompositeiMateDefinition(objs), name);
                            break;
                    }
                }
                else if (str == "Овал")
                {

                    switch (Doc.DocumentType)
                    {
                        case DocumentTypeEnum.kPartDocumentObject:
                            Face f1;
                            f1 = (((Edge)edge1).Faces[1].SurfaceType == SurfaceTypeEnum.kPlaneSurface) ? ((Edge)edge1).Faces[1]: ((Edge)edge1).Faces[2];
                            iMateDef = findAxis(PartCompDef, edge1 as Edge);
                            objs.Add(iMateDef);
                            iMateDef = PartCompDef.iMateDefinitions.AddMateiMateDefinition(f1, offset / 10);
                            objs.Add(iMateDef);
                            addName(PartCompDef.iMateDefinitions.AddCompositeiMateDefinition(objs), name);
                            break;
                        case DocumentTypeEnum.kAssemblyDocumentObject:
                            Face fp1 = (((Edge)edge1).Faces[1].SurfaceType == SurfaceTypeEnum.kPlaneSurface) ? ((EdgeProxy)edge1).Faces[1] : ((Edge)edge1).Faces[2];
                            Face fp2 = (((Edge)edge1).Faces[1].SurfaceType == SurfaceTypeEnum.kCylinderSurface) ? ((Edge)edge1).Faces[1] : ((Edge)edge1).Faces[2];
                            iMateDef = AsmCompDef.iMateDefinitions.AddMateiMateDefinition(fp2, 0, InferredTypeEnum.kInferredLine);
                            objs.Add(iMateDef);
                            iMateDef = AsmCompDef.iMateDefinitions.AddMateiMateDefinition(fp1, 0);
                            objs.Add(iMateDef);
                            addName(AsmCompDef.iMateDefinitions.AddCompositeiMateDefinition(objs), name);
                            break;
                    }
                }
                else if (str == "ТЭН-резистор")
                {
                    
                    switch (Doc.DocumentType)
                    {
                        case DocumentTypeEnum.kAssemblyDocumentObject:
                            insIMateDef = iMate_(edge1, AsmCompDef, offset);
                            objs.Add(insIMateDef);
                            iMateDef = iMate__(edge2, AsmCompDef, offset);
                            objs.Add(iMateDef);
                            addName(AsmCompDef.iMateDefinitions.AddCompositeiMateDefinition(objs),name);
                            break;
                        case DocumentTypeEnum.kPartDocumentObject:
                            insIMateDef = iMate_(edge1, PartCompDef, offset);
                            objs.Add(insIMateDef);
                            iMateDef = iMate__(edge2, PartCompDef, offset);
                            objs.Add(iMateDef);
                            addName(PartCompDef.iMateDefinitions.AddCompositeiMateDefinition(objs),name);
                            break;
                    }
                }
                else if (str == "ТЭН")
                {
                    
                    switch (Doc.DocumentType)
                    {
                        case DocumentTypeEnum.kAssemblyDocumentObject:
                            iMateDef = iMate__(edge1, AsmCompDef, offset);
                            objs.Add(iMateDef);
                            iMateDef = (((Edge)edge1).Faces[1].SurfaceType == SurfaceTypeEnum.kPlaneSurface) ? PartCompDef.iMateDefinitions.AddMateiMateDefinition(((Edge)edge1).Faces[1], offset / 10, BiasPoint: ((Edge)edge1).PointOnEdge) :
                                PartCompDef.iMateDefinitions.AddMateiMateDefinition(((Edge)edge1).Faces[2], offset / 10, BiasPoint: ((Edge)edge1).PointOnEdge);
                            objs.Add(iMateDef);
                            iMateDef = iMate__(edge2, AsmCompDef, offset);
                            objs.Add(iMateDef);
                            addName(AsmCompDef.iMateDefinitions.AddCompositeiMateDefinition(objs),name);
                            break;
                        case DocumentTypeEnum.kPartDocumentObject:
                            iMateDef = iMate__(edge1, PartCompDef, offset);
                            objs.Add(iMateDef);
                            iMateDef = (((Edge)edge1).Faces[1].SurfaceType == SurfaceTypeEnum.kPlaneSurface) ? PartCompDef.iMateDefinitions.AddMateiMateDefinition(((Edge)edge1).Faces[1], offset / 10, BiasPoint: ((Edge)edge1).PointOnEdge) :
                                PartCompDef.iMateDefinitions.AddMateiMateDefinition(((Edge)edge1).Faces[2], offset / 10, BiasPoint: ((Edge)edge1).PointOnEdge);
                            objs.Add(iMateDef);
                            iMateDef = iMate__(edge2, PartCompDef, offset);
                            objs.Add(iMateDef);
                            addName(PartCompDef.iMateDefinitions.AddCompositeiMateDefinition(objs),name);
                            break;
                    }
                }
//                 else if (str == "Шип")
//                 {
//                     
//                     //edge2 = (Edge)CmdMgr.Pick(SelectionFilterEnum.kAllLinearEntities, "Выберите верхнюю часть выреза под замыкающий шип");
//                     switch (Doc.DocumentType)
//                     {
//                         case DocumentTypeEnum.kAssemblyDocumentObject:
//                             thorn<AssemblyComponentDefinition>(edge1, AsmCompDef, ref objs);
//                             //thorn<AssemblyComponentDefinition>(edge2, AsmCompDef, ref objs);
//                             addName(PartCompDef.iMateDefinitions.AddCompositeiMateDefinition(objs),name);
//                             break;
//                         case DocumentTypeEnum.kPartDocumentObject:
//                             thorn<PartComponentDefinition>(edge1, PartCompDef, ref objs);
//                             //thorn<PartComponentDefinition>(edge2, PartCompDef, ref objs);
//                             addName(PartCompDef.iMateDefinitions.AddCompositeiMateDefinition(objs),name);
//                             break;
//                     }
//                 }
                else if (str == "Одиночная вставка" || str == "По элементу")
                {

                    //edge2 = (Edge)CmdMgr.Pick(SelectionFilterEnum.kAllLinearEntities, "Выберите верхнюю часть выреза под замыкающий шип");
                    switch (Doc.DocumentType)
                    {
                        case DocumentTypeEnum.kAssemblyDocumentObject:
                            insIMateDef = iMate_(edge1, AsmCompDef, offset);
                            addName(insIMateDef, name);
                            break;
                        case DocumentTypeEnum.kPartDocumentObject:
                            insIMateDef = iMate_(edge1, PartCompDef, offset);
                            addName(insIMateDef, name);
                            break;
                    }
                }
                else if (str == "DIN-рейка")
                {
                    switch (Doc.DocumentType)
                    {
                        case DocumentTypeEnum.kAssemblyDocumentObject:
                            iMateDef = iMate__(edge1, AsmCompDef, offset, InferredTypeEnum.kInferredPoint);
                            objs.Add(iMateDef);
                            iMateDef = iMate__(edge2, AsmCompDef, offset, InferredTypeEnum.kInferredPoint);
                            objs.Add(iMateDef);
                            iMateDef = iMate__(edge2, AsmCompDef, offset, InferredTypeEnum.kNoInference);
                            objs.Add(iMateDef);
                            addName(AsmCompDef.iMateDefinitions.AddCompositeiMateDefinition(objs), name);
                            break;
                        case DocumentTypeEnum.kPartDocumentObject:
                            iMateDef = iMateMate__(edge1, PartCompDef, InferredTypeEnum.kInferredPoint);
                            objs.Add(iMateDef);
                            iMateDef = iMateMate__(edge2, PartCompDef, InferredTypeEnum.kInferredPoint);
                            objs.Add(iMateDef);
                            iMateDef = iMateMate__(edge2, PartCompDef, InferredTypeEnum.kNoInference);
                            objs.Add(iMateDef);
                            addName(PartCompDef.iMateDefinitions.AddCompositeiMateDefinition(objs), name);
                            break;
                    }
                }
                objs.Clear();
            }
            catch (Exception ex)
            {
                tr.End();
                MessageBox.Show(ex.ToString());
                this.Close();
            }
            finally
            {
                
                if (close)
                    this.Close();
                else
                {
                    this.Show();
                    this.button1.Focus();
                }
            }
        }

        private MateiMateDefinition findAxis(PartComponentDefinition def ,Edge ed)
        {
            if (ed.TangentiallyConnectedEdges.Count == 1)
            {
                Face f = (ed.Faces[1].SurfaceType == SurfaceTypeEnum.kCylinderSurface) ? ed.Faces[1] : ed.Faces[2];
                return def.iMateDefinitions.AddMateiMateDefinition(f, 0, InferredTypeEnum.kInferredLine);
            }
            SelectionFilterEnum[] sels = new SelectionFilterEnum[]{SelectionFilterEnum.kWorkAxisFilter};
            Arc3d c = ed.Geometry as Arc3d;
            ObjectsEnumerator col = def.FindUsingPoint(ed.PointOnEdge, ref sels, c.Radius * 5, false);
            WorkAxis ax = null;
            foreach (var item in col.OfType<WorkAxis>())
	        {
                ax = item;
                //Face f = (ed.Faces[1].SurfaceType == SurfaceTypeEnum.kCylinderSurface) ? ed.Faces[1]: ed.Faces[2];
                if (ax.DefinitionType == WorkAxisDefinitionEnum.kNormalToSurfaceWorkAxis)
                {
                    return def.iMateDefinitions.AddMateiMateDefinition(ax, 0, InferredTypeEnum.kInferredLine);
                }
                ax = null;
	        }
            if (ax == null)
            {
                return iMate__(ed, def);
            }
            return null;
        }

//         private void thorn<T>(object oEdge, T compDef, ref ObjectCollection objs, double offset = 0.15)
//         {
// 
//             if (((Edge)oEdge).Faces.Count == 2)
//             {
//                 Vertex v = ((Edge)oEdge).StartVertex;
//                 Face maxF, minF, extFace;
//                 int val = (int)(((Edge)oEdge).StartVertex.Point.DistanceTo(((Edge)oEdge).StopVertex.Point)*1000); 
//                 bool isThorn = (val > 78 && val < 101) ? true : false;
//                 minF = (((Edge)oEdge).Faces[1].Evaluator.Area < ((Edge)oEdge).Faces[2].Evaluator.Area) ? ((Edge)oEdge).Faces[1] : ((Edge)oEdge).Faces[2];
//                 //if (oEdge.Geometry)
//                 if (isThorn)
//                 {
//                     //extFace = v.Faces.OfType<Face>().FirstOrDefault(fac => fac.Edges.Count > 4);
//                     //maxF = (oEdge.Faces[1].Evaluator.Area > oEdge.Faces[2].Evaluator.Area) ? oEdge.Faces[1] : oEdge.Faces[2];
//                     Edge ed = (minF.Edges[1].Equals(oEdge)) ? 
//                         (minF.Edges[1].StartVertex.Point.IsEqualTo(minF.Edges[4].StopVertex.Point,0.1))? minF.Edges[2]: minF.Edges[4] :
//                         (minF.Edges[1].StartVertex.Point.IsEqualTo(minF.Edges[4].StopVertex.Point, 0.1)) ? minF.Edges[1] : minF.Edges[3]; 
//                     maxF = (ed.Faces[1].Equals(minF)) ? ed.Faces[2] : ed.Faces[1];
//                     extFace = (((Edge)oEdge).Faces[1].Evaluator.Area > ((Edge)oEdge).Faces[2].Evaluator.Area) ? ((Edge)oEdge).Faces[1] : ((Edge)oEdge).Faces[2];
//                 }
//                 else
//                 {
//                     //Edge ed = (minF.Edges[1].Equals(oEdge))?minF.Edges[2]:minF.Edges[1];
//                     Edge ed = (minF.Edges[1].Equals(oEdge)) ?
//                         (minF.Edges[1].StartVertex.Point.IsEqualTo(minF.Edges[4].StopVertex.Point, 0.1)) ? minF.Edges[4] : minF.Edges[2] :
//                         (minF.Edges[1].StartVertex.Point.IsEqualTo(minF.Edges[4].StopVertex.Point, 0.1)) ? minF.Edges[3] : minF.Edges[1];
//                     maxF = (ed.Faces[1].Equals(minF))?ed.Faces[2]:ed.Faces[1];
//                     extFace = (((Edge)oEdge).Faces[1].Evaluator.Area > ((Edge)oEdge).Faces[2].Evaluator.Area) ? ((Edge)oEdge).Faces[1] : ((Edge)oEdge).Faces[2];
//                 }
//                 //Type t = compDef.GetType();
//                 if (compDef is PartComponentDefinition)
//                 {
//                     iMateDef = ((PartComponentDefinition)compDef).iMateDefinitions.AddMateiMateDefinition(minF, offset / 10, BiasPoint: ((Edge)oEdge).PointOnEdge);
//                     objs.Add(iMateDef);
//                     iMateDef = ((PartComponentDefinition)compDef).iMateDefinitions.AddMateiMateDefinition(extFace, 0, BiasPoint: ((Edge)oEdge).PointOnEdge);
//                     objs.Add(iMateDef);
//                     iMateDef = ((PartComponentDefinition)compDef).iMateDefinitions.AddMateiMateDefinition(maxF, offset / 10, BiasPoint: ((Edge)oEdge).PointOnEdge);
//                     objs.Add(iMateDef);
//                 }
//                 else if (compDef is AssemblyComponentDefinition/*t == typeof(AssemblyComponentDefinition)*/)
//                 {
//                     iMateDef = ((AssemblyComponentDefinition)compDef).iMateDefinitions.AddMateiMateDefinition(minF, offset / 10, BiasPoint: ((Edge)oEdge).PointOnEdge);
//                     objs.Add(iMateDef);
//                     iMateDef = ((AssemblyComponentDefinition)compDef).iMateDefinitions.AddMateiMateDefinition(extFace, 0, BiasPoint: ((Edge)oEdge).PointOnEdge);
//                     objs.Add(iMateDef);
//                     iMateDef = ((AssemblyComponentDefinition)compDef).iMateDefinitions.AddMateiMateDefinition(maxF, offset / 10, BiasPoint: ((Edge)oEdge).PointOnEdge);
//                     objs.Add(iMateDef);
//                 }
//             }
//         }

        private MateiMateDefinition mateiMate<T>(T compDef, Face f, double offset = 0.15)
        {
            if (compDef.GetType() == typeof(PartComponentDefinition))
                return ((PartComponentDefinition)compDef).iMateDefinitions.AddMateiMateDefinition(f, offset / 10, BiasPoint: f.Edges[1].StartVertex);
            else if (compDef.GetType() == typeof(AssemblyComponentDefinition))
                return ((AssemblyComponentDefinition)compDef).iMateDefinitions.AddMateiMateDefinition(f, offset / 10, BiasPoint: f.Edges[1].StartVertex);
            return null;
        }

        public Inventor.InsertiMateDefinition iMate_(object oEdge, PartComponentDefinition compDef, double offset = 0)
        {
            return compDef.iMateDefinitions.AddInsertiMateDefinition(oEdge, true, offset / 10) ;
        }

        public Inventor.InsertiMateDefinition iMate_(object oEdge, AssemblyComponentDefinition compDef, double offset = 0)
        {
            return compDef.iMateDefinitions.AddInsertiMateDefinition(oEdge, true, offset / 10);
        }
        public Inventor.MateiMateDefinition iMate__(object oEdge, AssemblyComponentDefinition compDef, double offset = 0, InferredTypeEnum inf = InferredTypeEnum.kInferredLine)
        {
            MateiMateDefinition oImate = null; Inventor.Face oFace; Inventor.Point pt;
            switch (inf)
            {
                case InferredTypeEnum.kInferredLine:
                    oFace = (((Edge)oEdge).Faces[1].SurfaceType == SurfaceTypeEnum.kCylinderSurface) ? ((Edge)oEdge).Faces[1] : ((Edge)oEdge).Faces[2];
                    if (((Edge)oEdge).TangentiallyConnectedEdges.Count == 1)
                    {
                        oImate = compDef.iMateDefinitions.AddMateiMateDefinition(oFace, offset, inf, oFace.PointOnFace);
                    }
                    break;
                case InferredTypeEnum.kInferredPoint:
                    pt = (((Edge)oEdge).Geometry as Circle).Center;
                    oImate = compDef.iMateDefinitions.AddMateiMateDefinition(oEdge, offset, inf);
                    break;
                case InferredTypeEnum.kNoInference:
                    oFace = (((Edge)oEdge).Faces[1].SurfaceType == SurfaceTypeEnum.kPlaneSurface) ? ((Edge)oEdge).Faces[1] : ((Edge)oEdge).Faces[2];
                    oImate = compDef.iMateDefinitions.AddMateiMateDefinition(oFace, offset, inf);
                    break;
                default:
                    break;
            }      
            return oImate;
        }

        public Inventor.MateiMateDefinition iMateMate__(object oEdge, PartComponentDefinition compDef, InferredTypeEnum inf = InferredTypeEnum.kInferredPoint)
        {
            MateiMateDefinition def = null; Inventor.Face oFace;
            switch (inf)
            {
                case InferredTypeEnum.kInferredLine:
                    break;
                case InferredTypeEnum.kInferredPoint:
                    def = compDef.iMateDefinitions.AddMateiMateDefinition(oEdge, 0, inf);
                    break;
                case InferredTypeEnum.kNoInference:
                    oFace = (((Edge)oEdge).Faces[1].SurfaceType == SurfaceTypeEnum.kPlaneSurface) ? ((Edge)oEdge).Faces[1] : ((Edge)oEdge).Faces[2];
                    def = compDef.iMateDefinitions.AddMateiMateDefinition(oFace, 0, inf);
                    break;
                default:
                    break;
            }
            return def;
        }

        public Inventor.MateiMateDefinition iMate__(object oEdge, PartComponentDefinition compDef, double offset = 0)
        {
            MateiMateDefinition oImate = null ; Inventor.Face oFace; WorkPoint wp; WorkAxis wa;
            SketchArc sa1, sa2; SketchLine lin; MidpointConstraint mpc; SketchPoint sp; object a = 0.0;
            oFace = (((Edge)oEdge).Faces[1].SurfaceType == SurfaceTypeEnum.kCylinderSurface)? ((Edge)oEdge).Faces[1] : ((Edge)oEdge).Faces[2];
            if (((Edge)oEdge).TangentiallyConnectedEdges.Count == 1)
             {
                                /*oFace = oEdge.Faces[1].SurfaceType == SurfaceTypeEnum.kCylinderSurface ? oEdge.Faces[1] : oEdge.Faces[2];*/
//                                 wa = compDef.WorkAxes.AddByRevolvedFace(oFace);
//                                 wa.Visible = false;
                oImate = compDef.iMateDefinitions.AddMateiMateDefinition(oFace, 0, InferredTypeEnum.kInferredLine, oFace.PointOnFace);
             }

            else if (((Edge)oEdge).TangentiallyConnectedEdges.Count == 4 || ((Edge)oEdge).TangentiallyConnectedEdges.Count == 3)
            {
                ps = (((Edge)oEdge).Faces[1].SurfaceType == SurfaceTypeEnum.kPlaneSurface)? compDef.Sketches.Add(((Edge)oEdge).Faces[1]): compDef.Sketches.Add(((Edge)oEdge).Faces[2]);
                switch (((Edge)((Edge)oEdge).TangentiallyConnectedEdges[1]).GeometryType)
                {
                    case CurveTypeEnum.kCircularArcCurve:
                        sa1 = (SketchArc)ps.AddByProjectingEntity(((Edge)oEdge).TangentiallyConnectedEdges[1]);
                        sa2 = (SketchArc)ps.AddByProjectingEntity(((Edge)oEdge).TangentiallyConnectedEdges[3]);
                        break;
                    default :
                        sa1 = (SketchArc)ps.AddByProjectingEntity(((Edge)oEdge).TangentiallyConnectedEdges[2]);
                        sa2 = (SketchArc)ps.AddByProjectingEntity(((Edge)oEdge).TangentiallyConnectedEdges[4]);
                        break;
                }
                oFace = (((Edge)oEdge).Faces[1].SurfaceType == SurfaceTypeEnum.kCylinderSurface) ? ((Edge)oEdge).Faces[2] : ((Edge)oEdge).Faces[1];
                lin = ps.SketchLines.AddByTwoPoints(sa1.CenterSketchPoint, sa2.CenterSketchPoint);
                sp = ps.SketchPoints.Add(invApp.TransientGeometry.CreatePoint2d());
                mpc = ps.GeometricConstraints.AddMidpoint(sp,lin);
                ps.Visible = false;
                wp = compDef.WorkPoints.AddByPoint(sp);
                wa = compDef.WorkAxes.AddByNormalToSurface(oFace,wp);
                wp.Visible = false; wa.Visible = false;
                oImate = compDef.iMateDefinitions.AddMateiMateDefinition(wa, a, InferredTypeEnum.kInferredLine, oFace.PointOnFace);
                
            }
            return oImate;
        }

        private void IMate_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape) this.Dispose(true);
        }

        private void IMate_FormClosed(object sender, FormClosedEventArgs e)
        {
            tr.End();
        }

        private void IMate_Load(object sender, EventArgs e)
        {
            tr = invApp.TransactionManager.StartTransaction((_Document)Doc, "Конструктивные пары");
        }

        public void selOp(ref List<Edge> edgeCmp)
        {
            CommandManager cmdMgr = Macros.StandardAddInServer.m_inventorApplication.CommandManager;
            //input = cmdMgr.UserInputEvents;
            try
            {
                intEvts = cmdMgr.CreateInteractionEvents();
                intEvts.InteractionDisabled = false;
                sel = intEvts.SelectEvents;
                sel.WindowSelectEnabled = true;
                sel.AddSelectionFilter(SelectionFilterEnum.kPartEdgeCircularFilter);
                sel.OnSelect += new SelectEventsSink_OnSelectEventHandler(select);
                key = intEvts.KeyboardEvents;
                key.OnKeyPress += new KeyboardEventsSink_OnKeyPressEventHandler(keyOp);
                intEvts.Start();
                intEvts.StatusBarText = "Выберите отверстия:";
                //input.OnSelect += new UserInputEventsSink_OnSelectEventHandler(select);
                //face = (Face)cmdMgr.Pick(SelectionFilterEnum.kPartFaceFilter, "Выберите поверхность:");

                flag = true;
                while (flag)
                {
                    Macros.StandardAddInServer.m_inventorApplication.UserInterfaceManager.DoEvents();
                }

                for (int i = 0; i < sel.SelectedEntities.Count; i++)
                {
                    try
                    {
                        Edge ed = (Edge)sel.SelectedEntities[i + 1];
                        Inventor.Face f = (ed.Faces[1].SurfaceType == SurfaceTypeEnum.kCylinderSurface) ? ed.Faces[1] : ed.Faces[2];
                        edgeCmp.Add(ed);
                        //ed = (f.Edges[1].Equals(ed)) ? f.Edges[2] : f.Edges[1];
                        //edgeCmp1.Add(ed);
                    }
                    catch (Exception) { }
                }

                intEvts.Stop();

                sel.OnSelect -= new SelectEventsSink_OnSelectEventHandler(select);
                key.OnKeyPress -= new KeyboardEventsSink_OnKeyPressEventHandler(keyOp);
                sel = null;
                key = null;
                intEvts = null;
            }
            catch (Exception)
            {
                sel = null;
                key = null;
                intEvts = null;
            }
        }

        private void select(ObjectsEnumerator JustSelectedEntities, SelectionDeviceEnum SelectionDevice, Inventor.Point ModelPosition, Point2d ViewPosition, Inventor.View View)
        {
        }
        private void keyOp(int KeyASCII)
        {
            if (KeyASCII == 32)
            {
                Edge ed = (Edge)sel.SelectedEntities[sel.SelectedEntities.Count];
                double r = ((Circle)ed.Geometry).Radius;
                Inventor.Face f = (ed.Faces[1].SurfaceType == SurfaceTypeEnum.kPlaneSurface) ? ed.Faces[1] : ed.Faces[2];
                foreach (Edge e in f.Edges)
                {
                    if (e.GeometryType == CurveTypeEnum.kCircleCurve && ((Circle)e.Geometry).Radius == r)
                    {
                        sel.AddToSelectedEntities(e);
                    }
                }
            }
            else if (KeyASCII == 13)
            {
                flag = false;
            }
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            ComboBox cb = sender as ComboBox;
            if (cb.Text == "По элементу")
            {
                checkBox1.Visible = true;  
            }
        }
    }

    internal class IMateBtn : Button
    {
        public static IMate m_IMate;
        public static string name = "", typ = "", offset = "";
        public Inventor.Document pDoc { get; set; }
        public static IMate getIMate
        {
            get
            {
                return m_IMate;
            }
        }

        #region "Methods"
        public IMateBtn(string displayName, string internalName, CommandTypesEnum commandType, string clientId, string description, string tooltip, Icon standardIcon, Icon largeIcon, ButtonDisplayEnum buttonDisplayType)
            : base(displayName, internalName, commandType, clientId, description, tooltip, standardIcon, largeIcon, buttonDisplayType)
        {
        }
        public IMateBtn(string displayName, string internalName, CommandTypesEnum commandType, string clientId, string description, string tooltip, ButtonDisplayEnum buttonDisplayType)
            : base(displayName, internalName, commandType, clientId, description, tooltip, buttonDisplayType)
        {
        }

        protected override void ButtonDefinition_OnExecute(NameValueMap context)
        {
            //InvAddIn.Tables tbl = new Tables(InventorApplication.ActiveDocument, InventorApplication, "fsdf");
            //tbl.addLeader(true);
            m_IMate = new InvAddIn.IMate(InventorApplication.ActiveDocument, name, typ, offset);      
            Macros.StandardAddInServer.forms.Add(m_IMate);
            //InterfaceDll.MyEvents ev = new InterfaceDll.MyEvents(m_IMate);
            //ev.addKeyEvent();
            if (Macros.StandardAddInServer.activeteForm()) System.Windows.Forms.Application.Run(m_IMate);
        }

        #endregion
    }
}
