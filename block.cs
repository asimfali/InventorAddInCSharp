using System.Collections.Generic;
using System.Linq;
using Inventor;
using InvDoc;
using InterfaceDll;
using System.Xml.Linq;

namespace InvAddIn
{
    enum type_part { hole, cut  };
    class Block
    {
        static public MyForm f;
        static Dictionary<string, XElement> dims;
        public PlanarSketch ps;
        public string name, direct;
        static MyXML xml;
       
        public SheetMetalComponentDefinition def;
        //public IEnumerable<IGrouping<string, SketchPoint>> pts;
        public IEnumerable<IGrouping<string, SketchEntity>> ents;
        public List<HoleFeature> hfs = new List<HoleFeature>();
        List<iMateDefinition> ims = new List<iMateDefinition>();
        List<string> cuts = new List<string>();
        public Block(PlanarSketch ps, SheetMetalComponentDefinition def)
        {
            var tr = I.beginTrans("blocks");
            this.def = def;
            this.ps = ps;
            if (ps == null) this.ps = findSketch((Document)def.Document);
            xml = new MyXML("AutoBlocks.xml", "Blocks");
            run();
            tr.End();
        }
        public static void setXML(string name)
        {
            xml = new MyXML(name, "Blocks");
        }
        public void run()
        {
            //var ie = u.gets<SketchPoint>(ps.SketchPoints, f => f.HoleCenter && f.SketchBlockPath.Count != 0);
            //pts = ie.GroupBy(e => e.SketchBlockPath[1].Name);
            //foreach (var item in pts)
            //{
            //    name = item.Key.Split(':')[0];
            //    if (!item.Key.StartsWith(name)) continue;
            //    run(name, item);
            //}
            var ie2 = u.gets<SketchEntity>(ps.SketchEntities, f => f.SketchBlockPath.Count != 0);
            ents = ie2.GroupBy(e => e.SketchBlockPath[1].Name);
            foreach (var item in ents)
            {
                name = item.Key.Split(':')[0];
                if (!item.Key.StartsWith(name)) continue;
                run(name, item);
            }
        }
        public void run(string n, IGrouping<string, SketchEntity> gr)
        {
            var el = MyXML.find(xml.elem, "name", n);
            if (el == null) return;
            foreach (var item in el.Elements())
            {
                if (item.Name == "Hole") addHole(item, gr);
                if (item.Name == "Imate" && hfs.Count > 0) addImate(item);
                if (item.Name == "Cut") addCut(item, gr);
            }
        }
        public bool filterSE(SketchEntity se, type_part t)
        {
            PlanarSketch ps = se.Parent as PlanarSketch;
            try
            {
                foreach (var item in ps.Dependents)
                {
                    var extrude = item as ExtrudeFeature;
                    if (extrude != null && t == type_part.cut)
                    {
                        var se1 = extrude.Profile[1][1].SketchEntity;
                        if (se1.SketchBlockPath.Count == 0) continue;
                        if (se.SketchBlockPath[1].Name == se1.SketchBlockPath[1].Name)
                            return true;
                    }
                    var hole = item as HoleFeature;
                    if (hole != null && t == type_part.hole)
                    {
                        //if (hfs.Contains(hole)) continue;
                        if (!(se is SketchPoint)) continue;
                        var def = hole.PlacementDefinition as SketchHolePlacementDefinition;
                        if (def == null) return false;
                        int i = 1;
                        foreach (SketchPoint sp in def.HoleCenterPoints)
                        {
                            if (u.eq(sp.Geometry, ((SketchPoint)se).Geometry)) return true;
                            i++;
                        }
                    }
                }
            }
            catch (System.Exception)
            {
            }
            return false;
        }
        public void addCut(XElement el, IGrouping<string, SketchEntity> gr)
        {
            direct = MyXML.getAtt(el, "Direction");
            var pr = ps.Profiles.AddForSolid();
            var col = I.COC();
            var name = gr.Key.Split(':')[0];
            if (cuts.Contains(name)) return;
            for (int i = 0; i < pr.Count; i++)
            {
                ProfilePath pp = pr[i + 1];

                if (filterSE(pp[1].SketchEntity, type_part.cut)) continue;
                var name1 = pp[1].SketchEntity.SketchBlockPath[1].Name;
                name1 = name1.Split(':')[0];
                if (name == name1)
                {
                    foreach (ProfileEntity item in pp)
                    {
                        var se = item.SketchEntity;
                        if (se.Construction) continue;
                        if (se is SketchPoint) continue;
                        col.Add(se);
                    }
                }
            }
            if (col.Count == 0) return;
            var pr1 = ps.Profiles.AddForSolid(false, col);
            SheetMetalFeatures smf = def.Features as SheetMetalFeatures;
            var cut_def = smf.ExtrudeFeatures.CreateExtrudeDefinition(pr1, PartFeatureOperationEnum.kCutOperation);
            //var cut_def = smf.CutFeatures.CreateCutDefinition(pr1);
            var v = ps.PlanarEntityGeometry.Normal; v = u.round(v);
            PartFeatureExtentDirectionEnum dir = PartFeatureExtentDirectionEnum.kNegativeExtentDirection;
            v = u.reverse(v);
            if (direct != "Встречно")
            {
                dir = PartFeatureExtentDirectionEnum.kPositiveExtentDirection;
                v = u.reverse(v);
            }
            SketchEntity ent = col[1] as SketchEntity;
            SketchPoint sp = getPoint(ent);
            if (sp == null) return;
            var fd = getDist(sp, v);
            string sbName;
            if (fd == null)
            {
                Face f = ps.PlanarEntity as Face;
                if (f == null)
                {
                    PartComponentDefinition pDef = ps.Parent as PartComponentDefinition;
                    sbName = pDef.SurfaceBodies[1].Name;
                }
                sbName = f.SurfaceBody.Name;
            } else sbName = fd.sbName;

            cut_def.SetThroughAllExtent(dir);

            //var cut = smf.CutFeatures.Add(cut_def);
            var cut = smf.ExtrudeFeatures.Add(cut_def);
            var c = getBody(sp.Geometry3d);
            if (c.Count == 0) return;
            cut.SetAffectedBodies(c);
            if (cut.HealthStatus == HealthStatusEnum.kDriverLostHealth)
            {
                var doc = def.Document as Document;
                doc.Update();
            }
            cuts.Add(name);
        }
        public ObjectCollection getBody(Point pt)
        {
            var col = I.COC();
            PartComponentDefinition def = ps.Parent as PartComponentDefinition;
            var f = new[] { SelectionFilterEnum.kPartBodyFilter };
            var objs = def.FindUsingPoint(pt, ref f, 1);
            foreach (SurfaceBody item in objs)
            {
                if (item.Visible) col.Add(item);
            }
            return col;
        }
        public SketchPoint getPoint(SketchEntity ent)
        {
            SketchPoint sp = null;
            switch (ent.Type)
            {
                case ObjectTypeEnum.kSketchArcObject:
                    sp = ((SketchArc)ent).CenterSketchPoint;
                    break;
                case ObjectTypeEnum.kSketchLineObject:
                    sp = ((SketchLine)ent).StartSketchPoint;
                    break;
                case ObjectTypeEnum.kSketchSplineObject:
                    sp = ((SketchSpline)ent).StartSketchPoint;
                    break;
                case ObjectTypeEnum.kSketchCircleObject:
                    sp = ((SketchCircle)ent).CenterSketchPoint;
                    break;
                default:
                    break;
            }
            return sp;
        }
        public void addImate(XElement el)
        {
            string n = MyXML.getAtt(el, "name");
            if (el.HasElements)
            {
                foreach (var item in el.Elements())
                {
                    addImate(item);
                }
                var col = I.COC();
                foreach (var item in ims)
                {
                    col.Add(item);
                }
                var comp = def.iMateDefinitions.AddCompositeiMateDefinition(col);
                IMate.addName(comp, n);
                ims.Clear();
                hfs.Clear();
                return;
            }
            string hs = MyXML.getAtt(el, "hole"), pt = MyXML.getAtt(el, "point"),
                t = MyXML.getAtt(el, "t"), d = MyXML.getAtt(el, "D"), dir = MyXML.getAtt(el, "Direction");
            double D = MyXML.convToDouble(d, 2); D *= 0.1;
            var hole = hfs[int.Parse(hs) - 1];
            var f = hole.Faces[int.Parse(pt)];
            var e = dir == "Сонаправленно" ? f.Edges[1] : f.Edges[2];

            switch (t)
            {
                case "ins":
                    ims.Add(def.iMateDefinitions.AddInsertiMateDefinition(e, true, D) as iMateDefinition);
                    break;
                case "axis":
                    ims.Add(def.iMateDefinitions.AddMateiMateDefinition(f, 0, InferredTypeEnum.kInferredLine) as iMateDefinition);
                    break;
                case "mate":
                    f = u.get<Face>(e.Faces, fi => fi.SurfaceType == SurfaceTypeEnum.kPlaneSurface);
                    ims.Add(def.iMateDefinitions.AddMateiMateDefinition(f, D) as iMateDefinition);
                    break;
                case "flush":
                    f = u.get<Face>(e.Faces, fi => fi.SurfaceType == SurfaceTypeEnum.kPlaneSurface);
                    ims.Add(def.iMateDefinitions.AddFlushiMateDefinition(f, D) as iMateDefinition);
                    break;
                default:
                    break;
            }
        }
        public void addHole(XElement el, IGrouping<string, SketchEntity> gr)
        {
            direct = MyXML.getAtt(el, "Direction");
            double d = MyXML.convToDouble(MyXML.getAtt(el, "D"))*0.1;
            var pts = MyXML.getAtt(el, "Points").Split(',');
            int i = 1;
            var col = I.COC();
            var pt = u.gets<SketchPoint>(gr, f => f.HoleCenter);
            foreach (var p in pt)
            {

                if (pts.Contains(i.ToString()) || (pts.Length == 1 && pts[0] == ""))
                {
                    if (filterSE((SketchEntity)p, type_part.hole)) continue;
                    col.Add(p);
                }
                i++;
            }
            if (col.Count != 0) addHole(col, d);
        }
        public FaceData getDist(SketchPoint sp, UnitVector v)
        {
            double d = 0;
            var pt = sp.Geometry3d;
            List<FaceData> tmp = new List<FaceData>();
            u.findUsingRay((PartComponentDefinition)def, pt, v, (o, l) =>
            {
                FaceData fd = new FaceData(o, (Point)l);
                if (fd.f != null)
                    tmp.Add(fd);
            });
            if (tmp.Count >= 2) return tmp[1];
            if (tmp.Count == 1) return tmp[0];
            return null;
        }
        public void addHole(ObjectCollection col, double d)
        {
            var v = ps.PlanarEntityGeometry.Normal; v = u.round(v);
            PartFeatureExtentDirectionEnum dir = PartFeatureExtentDirectionEnum.kNegativeExtentDirection;
            if (direct == "Встречно") {
                dir = PartFeatureExtentDirectionEnum.kPositiveExtentDirection;
                v = u.reverse(v);
            }
            var fd = getDist(col[1] as SketchPoint, v);
            var hf_def = def.Features.HoleFeatures.CreateSketchPlacementDefinition(col);
            var hf = def.Features.HoleFeatures.AddDrilledByDistanceExtent(hf_def, (d*10).ToString(), ((fd.d + 0.5)*10).ToString(), dir);
            hf.SetAffectedBodies(getBody(((SketchPoint)col[1]).Geometry3d));
            hfs.Add(hf);
        }

        public static PlanarSketch findSketch(Document doc)
        {
            if (doc.DocumentType != DocumentTypeEnum.kPartDocumentObject) return null;
            var br = doc.BrowserPanes.ActivePane;
            BrowserNode prev = null;
            foreach (BrowserNode node in br.TopNode.BrowserNodes[1].BrowserNodes)
            {
                var pf = node.NativeObject as PartFeature;
                if (node.NativeObject is EndOfFeatures) break;
                prev = node;
            }
            return prev.NativeObject as PlanarSketch;
        }
        public static SketchBlockDefinition getBlockDef(Document doc)
        {
            Block.setXML("AutoBlocks.xml");
            MyXML xml = new MyXML("CopyBlockInterface.xml");
            f = new MyForm(xml, "Блок");
            var pcd = I.getPCD(doc);
            Dictionary<string, SketchBlockDefinition> dic = new Dictionary<string, SketchBlockDefinition>();
            u.action<SketchBlockDefinition>(pcd.SketchBlockDefinitions, a => dic.Add(a.Name, a));
            f.cbs[0].Items.Clear();
            f.cbs[0].Items.AddRange(dic.Keys.ToArray());
            f.cbs[0].SelectedIndexChanged += Block_SelectedIndexChanged;
            f.cbs[1].Items.Clear();
            f.cbs[1].Text = "";
            f.lbls[1].Visible = false;
            f.cbs[1].Visible = false;
            f.f.ShowDialog();
            if (f.f.DialogResult != System.Windows.Forms.DialogResult.OK) return null;
            SketchBlockDefinition block;
            dic.TryGetValue(f.cbs[0].Text, out block);
            return block;
        }

        private static void Block_SelectedIndexChanged(object sender, System.EventArgs e)
        {
            if (dims == null) dims = new Dictionary<string, XElement>();
            else dims.Clear();
            var val = f.cbs[0].Text;
            var el = MyXML.find(xml.elem, "name", val);
            if (el == null) return;
            el = MyXML.getEl(el, "Dim");
            if (el == null) return;
            foreach (var item in el.Elements())
            {
                var n = MyXML.getAtt(item, "name");
                if (n == "") continue;
                dims.Add(n, item);
            }
            if (dims.Count > 0)
            {
                f.cbs[1].Items.AddRange(dims.Keys.ToArray());
                f.lbls[1].Visible = true;
                f.cbs[1].Visible = true;
            }
        }

        public static SketchBlockDefinition getBlock(Document doc, PartComponentDefinition pcd, SketchBlockDefinition sb)
        {
            var d = u.get<SketchBlockDefinition>(pcd.SketchBlockDefinitions, fi => fi.Name == sb.Name);
            return d;
        }
        public static SketchBlockDefinition copy(Document from, Document to)
        {
            var pcd = I.getPCD(to);
            SketchBlockDefinition blockFrom = getBlockDef(from);
            if (blockFrom == null) return null;

            if (f.cbs[1].Text != "" && Block.dims.Count > 0)
            {
                if (blockFrom != null)
                    changeDims(blockFrom, dims[f.cbs[1].Text]);
            }

            var BlockTo = getBlock(to, pcd, blockFrom);
            if (BlockTo == null)
            {
                BlockTo = blockFrom.CopyTo((_Document)to);
            }
            else if (f.cbs[1].Text != "" && Block.dims.Count > 0)
            {
                if (BlockTo != null)
                    changeDims(BlockTo, dims[f.cbs[1].Text]);
            }
            //if (f.cbs[1].Text != "" && Block.dims.Count > 0)
            //{
            //    if (block1 != null)
            //        changeDims(block1, dims[f.cbs[1].Text]);
            //}
            return BlockTo;
        }
        public static void changeDims(SketchBlockDefinition def, XElement el)
        {
            var smcd = def.Parent as SheetMetalComponentDefinition;
            foreach (var item in el.Elements())
            {
                if (item.Name == "Dim")
                {
                    var t = MyXML.getAtt(item, "t");
                    var indx = int.Parse(MyXML.getAtt(item, "index"));
                    var val = MyXML.getAtt(item, "D");
                    if (indx > def.DimensionConstraints.Count) continue;
                    var dim = def.DimensionConstraints[indx];
                    dim.Parameter.Expression = val;
                } else if (item.Name == "Param")
                {
                    var n = MyXML.getAtt(item, "name");
                    var val = MyXML.getAtt(item, "D");
                    var p = u.get<Parameter>(smcd.Parameters, f => f.Name == n);
                    if (p == null) continue;
                    p.Expression = val;
                }
            }
            ((Document)smcd.Document).Update();
        }
        public static bool checkConstraint(SketchPoint sp)
        {
            foreach (var item in sp.Constraints)
            {
                var con = item as CoincidentConstraint;
                if (con == null) continue;
                if (con.EntityOne.SketchBlockPath.Count != 0 ||
                    con.EntityTwo.SketchBlockPath.Count != 0) return true; 
            }
            return false;
        }
        public static void insertBlock(Document doc, SketchBlockDefinition block)
        {
            if (block == null) return;
            SelectSet ss = doc.SelectSet;
            PlanarSketch ps = ss.Count == 0 ? findSketch(doc) : ss[1] as PlanarSketch;
            if (ps == null) return;
            var pts = u.gets<SketchPoint>(ps.SketchPoints, fi => fi.HoleCenter && fi.SketchBlockPath.Count == 0);
            foreach (var item in pts)
            {
                //if (item.SketchBlockPath.Count != 0) continue;
                //if (checkConstraint(item)) continue;
                var sb = ps.SketchBlocks.AddByDefinition(block, item.Geometry);
                
                var sp = u.get<SketchPoint>(ps.SketchPoints, fi => u.eq(item.Geometry, fi.Geometry) && fi.SketchBlockPath.Count != 0
                && fi.SketchBlockPath[1].Name == sb.Name);
                 
                if (sp == null) continue;
                ps.GeometricConstraints.AddCoincident((SketchEntity)sp, (SketchEntity)item);
                item.HoleCenter = false;
            }
        }
    }
}
