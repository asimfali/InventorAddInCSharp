//#define INV14

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Linq;
using InterfaceDll;
using InvDoc;
using Inventor;

namespace InvAddIn
{
    class Variable
    {
        public Document doc;
        public PartDocument docFrom;
        public AssemblyDocument asm;
        public SheetMetalComponentDefinition smcd, smcdFrom;
        public SheetMetalFeatures smf, smfFrom;
        public PlanarSketch ps, ps_new;
        public SelectSet ss;
        public SurfaceBody sb;
        public PartFeature start_pf, pf;
        public BrowserPane pane;
        public BrowserNode node;
        public ComponentOccurrence occFrom, occTo;
        public string t, v, val, sb_name, nname = "";
        SketchCopy copy;
        Matrix mtx = null;
        Form form;
        List<iMateDefinition> ims = new List<iMateDefinition>();
        int num;
        readonly Plane pl;
        readonly int count = 0;
        //double w, offset, thick;
        MyForm F;
        MyXML xml;

        public Variable(Document doc, Form form = null)
        {
            this.doc = doc;
            copy = new SketchCopy(doc);
            if (form != null) this.form = form;
            if (init()) return;
        }
        public Variable(AssemblyDocument asm)
        {
            this.asm = asm;
            ss = asm.SelectSet;
            if (asm == null) return;
            initAsm(asm);
            copy = new SketchCopy(doc);
            copy.setMtx(mtx, smcdFrom as PartComponentDefinition);
            addUCS();
        }
        public void addUCS()
        {
            var def = smcd.UserCoordinateSystems.CreateDefinition();
            def.Transformation = mtx;
            smcd.UserCoordinateSystems.Add(def);
        }
        public double get_thick(SurfaceBody sb)
        {
            var st = smcd.GetBodySheetMetalStyle(sb);
            return u.convToDouble(st.Thickness)/10;
        }
        public PartDocument getDoc(ComponentOccurrence occ, ref SheetMetalComponentDefinition def)
        {
            def = occ.DefinitionReference.ReferencedDefinition as SheetMetalComponentDefinition;
            if (def == null) return null;
            return def.Document as PartDocument;
        }
        public void initAsm(AssemblyDocument asm)
        {
            if (ss.Count != 2) return;
            occFrom = ss[1] as ComponentOccurrence;
            occTo = ss[2] as ComponentOccurrence;
            mtx = occFrom.Transformation;
            if (occFrom == null || occTo == null) return;
            docFrom = getDoc(occFrom, ref smcdFrom);
            smfFrom = I.getSMF(docFrom as Document);
            doc = getDoc(occTo, ref smcd) as Document;
            init();
        }
        public bool init()
        {
            smcd = I.getSMCD(doc);
            if (smcd == null) return false;
            smf = smcd.Features as SheetMetalFeatures;
            ims = u.gets<iMateDefinition>(smcd.iMateDefinitions, f => !f.Suppressed).ToList();
            u.action(ims, a => a.Suppressed = true);
            return true;
        }
        public bool check_error<T>(IEnumerable<T> elems)
        {
            if (elems.Count() == 0)
            {
                MessageBox.Show("Не выбраны тела");
                return true;
            }
            return false;
        }
        public void add()
        {
            ss = doc.SelectSet;
            var bodies = u.gets<SurfaceBody>(ss, f => f is SurfaceBody).Select(e => e.Name).ToList();
            if (check_error(bodies)) return;
            if (!set_t()) return;
            pane = doc.BrowserPanes.ActivePane;
            foreach (var item in bodies)
            {
                with_body(item);
            }
            u.action(ims, a => a.Suppressed = false);
        }
        public void addFromAsm()
        {
            var bodies = u.gets<SurfaceBody>(smcdFrom.SurfaceBodies, f => f is SurfaceBody).Select(e => e.Name).ToList();
            if(check_error(bodies)) return;
            pane = doc.BrowserPanes.ActivePane;
            nname = file.name(docFrom.FullDocumentName);
            foreach (var item in bodies)
            {
                with_body(item);
            }
            u.action(ims, a => a.Suppressed = false);
        }
        public void extend()
        {
            ss = doc.SelectSet;
            SurfaceBody body = ss[ss.Count] as SurfaceBody;
            if (body == null) return;
            sb = body; sb_name = sb.Name;
            copy.setBody(sb);
            if (ss == null) return;
            //if (!set_t()) return;
            pane = doc.BrowserPanes.ActivePane;
            var pfs = u.gets<PartFeature>(ss, f => f is PartFeature);
            if (check_error(pfs)) return;
            copy.setOldBody(pfs.ElementAt(0).SurfaceBody);
            extend_body(body, pfs);
            u.action(ims, a => a.Suppressed = false);
        }
        public Edge find_edge(SurfaceBody n, Edge e)
        {
            var old = e.Geometry as Circle;
            var f = copy.lst(e.Geometry);
            if (mtx != null)
                copy.tr(f);
            foreach (Edge ed in n.Edges)
            {
                if (ed.CurveType != CurveTypeEnum.kCircleCurve) continue;
                var g = ed.Geometry as Circle;
                if (u.eq(g.Center, f[0] as Point) && u.eq(g.Radius, old.Radius))
                    return ed;
            }
            return null;
            //return u.findAtPoint(smcd as PartComponentDefinition, e.PointOnEdge, n, new[] { SelectionFilterEnum.kAllCircularEntities }) as Edge;
            //ObjectsEnumerator en = smcd.FindUsingPoint(e.PointOnEdge, new[] { SelectionFilterEnum.kPartEdgeFilter }, 1, true);
            //if (en.Count == 0) return null;
            //foreach (var item in en)
            //{
            //    Edge ed = item as Edge;
            //    if (ed == null) continue;
            //    if (ed.Parent.Name == n.Name) return ed;
            //}
            //return null;
        }
        public void add_imates(SurfaceBody old)
        {
            foreach (dynamic im in smcd.iMateDefinitions)
            {
                if (im is InsertiMateDefinition)
                {
                    add_ins_imate(old, im as InsertiMateDefinition);
                }
                else if (im is CompositeiMateDefinition)
                {
                    add_comp(old, im as CompositeiMateDefinition);
                }
            }
        }
        public CompositeiMateDefinition add_comp(SurfaceBody old, CompositeiMateDefinition im)
        {
            var col = I.COC();
            foreach (var item in im)
            {
                if (item is InsertiMateDefinition)
                {
                    var ins = add_ins_imate(old, item as InsertiMateDefinition);
                    col.Add(ins);
                }
                else return null;
            }
            return smcd.iMateDefinitions.AddCompositeiMateDefinition(col, im.Name, im.MatchList);
        }
        public InsertiMateDefinition add_ins_imate(SurfaceBody old, InsertiMateDefinition im)
        {
            Edge e = im.Entity as Edge;
            if (e.Parent != old) return null;
            var ne = find_edge(sb, e);
            if (ne == null) return null;
            InsertiMateDefinition ins = smcd.iMateDefinitions.AddInsertiMateDefinition(ne, im.AxesOpposed, im.Distance.ModelValue);
            ins.Name = im.Name;
            return ins;
        }
        public void add_folder()
        {
            var new_pane = doc.BrowserPanes.Add("test", "test");
            pane = doc.BrowserPanes.ActivePane;
            var nodes = get_nodes(pane, "hole");
            add_child(pane, "Отверстия", nodes);
        }
        public void add_child(BrowserPane p, string name, ObjectCollection col)
        {
            p.AddBrowserFolder(name, col);
        }
        public ObjectCollection get_nodes(BrowserPane p, string t)
        {
            var col = I.COC();
            foreach (BrowserNode item in p.TopNode.BrowserNodes[1].BrowserNodes)
            {
                if (t == "hole" && item.NativeObject is HoleFeature)
                {
                    col.Add(item);
                }
            }
            return col;
        }
        public bool set_t()
        {
            F = new MyForm("VariableInterface.xml", "Исполнение");
            set_num(F.cbs[0].Text);
            F.cbs[0].SelectedValueChanged += Variable_SelectionChangeCommitted;
            var dr = F.f.ShowDialog();
            if (dr != DialogResult.OK) return false;
            v = F.cbs[0].Text; val = F.cbs[2].Text;
            F.cbs[0].SelectedValueChanged -= Variable_SelectionChangeCommitted;
            num = u.convToInt(F.cbs[1].Text);
            return set_t(v, num, val);
        }

        private void Variable_SelectionChangeCommitted(object sender, EventArgs e)
        {
            ComboBox cb = sender as ComboBox;
            set_num(cb.Text);
        }
        public void set_num(string val)
        {
            var props = u.gets<Property>(doc.PropertySets[4], f => f.Name.ToLower().StartsWith(val.ToLower()));
            F.cbs[1].Text = (props.Count()).ToString("00");
        }

        public bool set_t(string v, int num, string val)
        {
            t = u.getPropValue(doc, v);
            if (t == "") u.addProp(doc, v, "");
            else
            {
                t = u.getPropValue(doc, $"{v}_{num:00}");
                if (t == "")
                    u.addProp(doc, $"{v}_{num:00}", val);
            }
            return true;
        }
        public string get_t(string name, int num)
        {
            return u.getPropValue(doc, $"{name}_{num:00}");
        }

        public void set_node()
        {
            //smcd.SetEndOfPartToTopOrBottom(false);
            foreach (BrowserNode item in pane.TopNode.BrowserNodes[1].BrowserNodes)
            {
                if (item.NativeObject is EndOfFeatures) break;
                else if (item.NativeObject is PartFeature) 
                    this.start_pf = item.NativeObject as PartFeature;
            }
        }
        public int get_ind(PartFeature pf)
        {
            int i = 0;
            foreach (BrowserNode item in pane.TopNode.BrowserNodes[1].BrowserNodes)
            {
                if (item.NativeObject == pf)
                {
                    return i;
                }
                i++;
            }
            return i;
        }
        public void sort_node(ref IEnumerable<PartFeature> lst)
        {
            lst = lst.OrderBy(e => get_ind(e));
        }
        public PartFeature find_feature(string name)
        {
            if (asm == null)
                return u.get<PartFeature>(smf, f => f.Name == name);
            else
                return u.get<PartFeature>(smfFrom, f => f.Name == name);
        }
        public SurfaceBody GetBody(string name)
        {
            if (asm == null)
                return u.get<SurfaceBody>(smcd.SurfaceBodies, f => f.Name == name);
            else
                return u.get<SurfaceBody>(smcdFrom.SurfaceBodies, f => f.Name == name);
        }
        public void with_body(string name)
        {
            SurfaceBody body = GetBody(name);
            copy.setOldBody(body);
            IEnumerable<PartFeature> lst;
            //set_node();
            if (asm != null)
                lst = u.gets<PartFeature>(smcdFrom.Features, f => true);
            else
                lst = u.gets<PartFeature>(body.AffectedByFeatures, f => true);
            with_part_feature(body.CreatedByFeature, true);
            set_name(name);
            if (sb != null)
            {
                copy.setBody(sb);
                sb_name = sb.Name;
            }
            body = GetBody(name);
            if (lst.Count() == 1) return;
            if (asm == null)
                sort_node(ref lst);
            foreach (PartFeature item in lst)
            {
                body = GetBody(name);
                if (item.Equals(body.CreatedByFeature)) continue;
                var feat = find_feature(item.Name);
                with_part_feature(feat);
            }
            //start_pf.SetEndOfPart(false);
            smcd.SetEndOfPartToTopOrBottom(false);
            body = GetBody(name);
            //add_imates(body);
            if (asm != null)
                docFrom.Close(true);
        }
        public void extend_body(SurfaceBody body, IEnumerable<PartFeature> lst)
        {
            //set_node();
            if (lst.Count() == 0) return;
            sort_node(ref lst);
            foreach (PartFeature item in lst)
            {
                if (item.Equals(u.get<SurfaceBody>(smcd.SurfaceBodies, el => el.Name == sb_name).CreatedByFeature)) continue;
                with_part_feature(item);
            }
            //start_pf.SetEndOfPart(false);
            smcd.SetEndOfPartToTopOrBottom(false);
        }
        public void set_name(string name)
        {
            //sb = getBody(name);
            if (nname == "")
                name = $"{this.v}_{num:00}^{name}";
            else
                name = nname;
            sb.Name = name;
        }
        public void set_style(SurfaceBody oldBody, SurfaceBody newBody)
        {
            var def = oldBody.ComponentDefinition as SheetMetalComponentDefinition;
            if (def == null) return;
            var st = def.GetBodySheetMetalStyle(oldBody);
            var def1 = newBody.ComponentDefinition as SheetMetalComponentDefinition;
            if (def1 == null) return;
            if (asm != null)
            {
                st = u.get<SheetMetalStyle>(def1.SheetMetalStyles, f => f.Name == st.Name);
            }
            def1.SetBodySheetMetalStyle(newBody, st);
        }
        public void with_part_feature(PartFeature pf, bool first = false)
        {
            //if (pf is FaceFeature)
            //{
            //    ss.Select(pf);
            //    ss.Delete();
            //}
            //if (asm != null)
            //{
            //    if (pf is FaceFeature)
            //    {
            //        pf.SetEndOfPart(true);
            //        pf.SetEndOfPart(false);
            //    }
            //    else
            //    {
            //        smcdFrom.SetEndOfPartToTopOrBottom(false);
            //    }
            //}
            if (pf is FaceFeature)
            {
                copy_face_feature(pf as FaceFeature, first);
            }
            else if (pf is ContourFlangeFeature)
            {
                copy_contour(pf as ContourFlangeFeature, first);
            }
            else if (pf is FlangeFeature) copy_flange(pf as FlangeFeature);
            else if (pf is CutFeature) copy_cut_feature(pf as CutFeature);
            else if (pf is ExtrudeFeature) copy_extrude(pf as ExtrudeFeature);
            else if (pf is HoleFeature) copy_hole(pf as HoleFeature);
            else if (pf is MirrorFeature) copy_mirror(pf as MirrorFeature);
            else if (pf is CornerRoundFeature) copy_fillet(pf as CornerRoundFeature);
            else if (pf is CornerChamferFeature) copy_chamfer(pf as CornerChamferFeature);
            else if (pf is RectangularPatternFeature) copy_arr(pf as RectangularPatternFeature);
            else if (pf is CircularPatternFeature) copy_arr(pf as CircularPatternFeature);
            else if (pf is HemFeature) copy_Hem(pf as HemFeature);
            
        }
        public void copy_sketch(string type = "hole", object def = null)
        {
            HashSet<string> blocks = new HashSet<string>();
            if (type == "cut")
            {
                var f = get_plane((CutDefinition)def);
                ps_new = smcd.Sketches.Add(f);
            } 
            else ps_new = smcd.Sketches.Add(ps.PlanarEntity);
            foreach (SketchEntity item in ps.SketchEntities)
            {
                if (item.Construction) continue;
                if (type == "hole" && !(item is SketchPoint)) continue;
                if (type == "face" && filter(item, ((FaceFeatureDefinition)def).Profile)) continue;
                if (type == "path" && filter(item, ((ContourFlangeDefinition)def).Path)) continue;
                if (type == "extrude" && filter(item, ((ExtrudeDefinition)def).Profile)) continue;
                if (type == "cut" && filter(item, ((CutDefinition)def).Profile)) continue;

                if (item.SketchBlockPath.Count != 0)
                {
                    var b = item.SketchBlockPath[1];
                    if (blocks.Contains(b.Name)) continue;
                    var bl = ps_new.SketchBlocks.AddByDefinition(b.Definition, b.Position);
                    bl.Transformation = b.Transformation;
                    blocks.Add(b.Name);
                    continue;
                }
                var se = ps_new.AddByProjectingEntity(item);
                var sp = se as SketchPoint;
                if (sp != null)
                {
                    sp.HoleCenter = ((SketchPoint)item).HoleCenter;
                }
            }
        }
        public void add_dim(SketchBlock bl, PlanarSketch ps)
        {

        }
        public Object get_plane(CutDefinition def)
        {
            var se = def.Profile[1][1].SketchEntity;
            ps = se.Parent as PlanarSketch;
            var plane = ps.PlanarEntityGeometry;
            //var f = ps.PlanarEntity as Face;
            //if (f != null)
            //{

            //    return f;
            //}
            foreach (Face fa in sb.Faces)
            {
                if (fa.SurfaceType != SurfaceTypeEnum.kPlaneSurface) continue;
                var pt1 = fa.PointOnFace;
                var n1 = u.normal(fa, ref pt1);
                var pt2 = plane.RootPoint;
                var n2 = plane.Normal;
                if (u.eq(pt1, n1, pt2, n2.AsVector())) return fa;
            }
            return null;
        }
        public bool filter(SketchEntity se, Profile pr)
        {
            foreach (ProfilePath path in pr)
            {
                foreach (ProfileEntity ent in path)
                {
                    if (ent.SketchEntity.Equals(se)) return false;
                }
            }
            return true;
        }
        public bool filter(SketchEntity se, Path pr)
        {
            foreach (PathEntity ent in pr)
            {
                if (ent.SketchEntity.Equals(se)) return false;
            }
            return true;
        }
        public void copy_sketch(HoleFeature f)
        {
            var pts = f.HoleCenterPoints;
            
            var pt = pts[1] as SketchPoint;
            ps = pt.Parent as PlanarSketch;
            ps_new = smcd.Sketches.Add(ps.PlanarEntity);
            foreach (var item in pts)
            {
                pt = ps_new.AddByProjectingEntity(item) as SketchPoint;
                pt.HoleCenter = true;
            }
        }
        public Path get_path(Path p, bool first = true)
        {
            SketchEntity se;
            Path np = null;
            var col = I.COC();
            foreach (PathEntity item in p)
            {
                se = item.SketchEntity as SketchEntity;
                col.Add(copy.getSE(se, se.Type));
            }
            np = smf.CreateSpecifiedPath(col);
            //if (first)
            //    se = p[1].SketchEntity as SketchEntity;
            //else
            //    se = p[p.Count].SketchEntity as SketchEntity;
            //var sl = se as SketchLine;

            //if (sl != null)
            //{
            //    var sl_new = u.get<SketchLine>(ps_new.SketchLines, f => u.eq(sl, f));
            //    if (sl_new == null && ps_new.SketchLines.Count == 1) sl_new = ps_new.SketchLines[1];
            //    np = smf.CreatePath(sl_new);
            //}
            //var a = se as SketchArc;
            //if (a != null)
            //{
            //    var sl_new = u.get<SketchArc>(ps_new.SketchArcs, f => u.eq(a, f));
            //    np = smf.CreatePath(sl_new);
            //}
            return np;
        }
        public Profile get_profile(PartFeature pf = null)
        {
            dynamic p = pf;
            Profile pr = null;
            if (pf is CutFeature || pf is ExtrudeFeature || pf is FaceFeature)
            {
                //CutFeature cut = pf as CutFeature;
                var obs = get_profile_entities(p.Definition.Profile);
                pr = ps_new.Profiles.AddForSolid(false, obs);
            }
            if (pr == null)
                pr = ps_new.Profiles.AddForSolid(false);
            if (pr != null) with_profile(p.Definition.Profile, pr);
            return pr;
        }
        public void with_profile(Profile p1, Profile p2)
        {
            for (int i = 1; i < p2.Count+1; i++)
            {
                p2[i].AddsMaterial = p1[i].AddsMaterial;
            }
        }
        public ObjectCollection get_profile_entities(Profile p)
        {
            var col = I.COC();
            foreach (ProfilePath item in p)
            {
                foreach (ProfileEntity pe in item)
                {
                    SketchEntity se = pe.SketchEntity;
                    var ne = copy.getSE(se, se.Type);
                    col.Add(ne);
                }
            }
            return col;
        }
        public bool check_vector(Path p1, Path p2)
        {
            SketchEntity se1 = p1[1].SketchEntity as SketchEntity, se2 = p2[1].SketchEntity as SketchEntity;
            return se1.RangeBox.IsDisjoint(se2.RangeBox);
        }
        public ContourFlangeFeature copy_contour(ContourFlangeFeature c, bool n = true)
        {
            if (asm == null)
                c.SetEndOfPart(false);
            //var def = c.Definition.Copy();
            var p = c.Definition.Path;
            var se = p[1].SketchEntity as SketchEntity;
            ps = se.Parent as PlanarSketch;
            ps_new = copy.checkSketch(ps);
            //copy_sketch("path", c.Definition);
            var np = get_path(p);
            //if (check_vector(np, p))
            //    np = get_path(p, false);
            var def1 = smf.ContourFlangeFeatures.CreateContourFlangeDefinition(np);
            def1.ExtentDirection = c.Definition.ExtentDirection;
            def1.Operation = c.Definition.Operation; def1.ApplyAutoMitering = c.Definition.ApplyAutoMitering;
            if (c.Definition.DefaultWidthExtentType == PartFeatureExtentEnum.kDistanceExtent)
            {
                var ext = c.Definition.DefaultWidthExtent as DistanceExtent;
                def1.SetDistanceExtent(c.FeatureDimensions[1].Parameter.Expression, ext.Direction);
            }
#if INV14
#else
            if (n)
            {
                def1.Operation = PartFeatureOperationEnum.kNewBodyOperation;
            }
            else
            {
                def1.Operation = PartFeatureOperationEnum.kJoinOperation;
                def1.AffectedBodies = get_body();
            }
#endif

            var f = smf.ContourFlangeFeatures.Add(def1);
            if (n)
            {
                set_style(c.SurfaceBody, f.SurfaceBody);
                sb = f.SurfaceBody;
            }
            //if (!eqBox(f.RangeBox, c.RangeBox, 0.01))
            //{
            //    f.Definition.ExtentDirection = f.Definition.ExtentDirection == PartFeatureExtentDirectionEnum.kNegativeExtentDirection ?
            //        PartFeatureExtentDirectionEnum.kPositiveExtentDirection : PartFeatureExtentDirectionEnum.kNegativeExtentDirection;
            //}
            return f;
        }
        public bool eqBox(Box b1, Box b2, double d)
        {
            Vector v1 = b1.MinPoint.VectorTo(b1.MaxPoint),
                v2 = b2.MinPoint.VectorTo(b2.MaxPoint);
            double dist = Math.Abs(v1.Length - v2.Length);
            return dist <= d;
        }
        public List<Edge> getEdges(Faces fs, double r, ref List<Edge> ed)
        {
            List<Edge> eds = new List<Edge>();
            var fa = u.gets<Face>(fs, f => f.SurfaceType == SurfaceTypeEnum.kCylinderSurface &&
            u.eq(((Cylinder)f.Geometry).Radius, r));
            foreach (var item in fa)
            {
                foreach (Edge e in item.Edges)
                {
                    if (e.CurveType != CurveTypeEnum.kLineCurve) continue;
                    if (e.Faces[1].CreatedByFeature != e.Faces[2].CreatedByFeature)
                    {
                        var e1 = find_edge(sb.Edges, e, r + get_thick(sb));
                        ed.Add(e);
                        if (e1 != null)
                            eds.Add(e1);
                        break;
                    }
                }
            }
            return eds;
        }
        public EdgeCollection get_edges(FlangeFeature ff, out Dictionary<Edge, PartFeatureExtent> offsets)
        {
            var col = I.app.TransientObjects.CreateEdgeCollection();
            List<Edge> eds = new List<Edge>(); List<Edge> ed = new List<Edge>();
            offsets = new Dictionary<Edge, PartFeatureExtent>();
            ff.SetEndOfPart(true);
            if (ff.Definition.Edges == null)
            {
                dynamic br = ff.Definition.BendRadius;
                eds = getEdges(ff.Faces, br.Value, ref ed);
            } else
            {
                foreach (Edge item in ff.Definition.Edges)
                {
                    eds.Add(item);
                }
            }
            int i = 0;
            foreach (Edge item in eds)
            {
                //var p1 = u.midPt(item);
                //var f = copy.lst(p1);
                //if (mtx != null) copy.tr(f);
                //var e = u.get<Edge>(sb.Edges, el => u.eq(u.midPt(el), f[0] as Point));
                //if (e != null)
                //{
                    col.Add(item);
                    PartFeatureExtent offset;
                    if (ed.Count == eds.Count) { continue; }
                    else offset = ff.Definition.GetWidthExtent(item);
                    if (offset != null)
                    {
                    var edge = find_edge(sb.Edges, item, 0.01);
                    offsets.Add(edge, offset);
                    }
                       
                //}
                i++;
            }
            if (offsets.Count != 0)
            {
                col.Clear();
                foreach (var item in offsets)
                {
                    col.Add(item.Key);
                }
            }
            return col;
        }
        public EdgeCollection get_edges(IEnumerable<Edge> edges)
        {
            sb = getBody(sb_name);
            var col = I.app.TransientObjects.CreateEdgeCollection();
            foreach (var item in edges)
            {
                var mp = u.midPt(item);
                var f = copy.lst(mp); if (mtx != null) copy.tr(f);
                var e = u.get<Edge>(sb.Edges, el => u.eq(f[0] as Point, u.midPt(el)));
                if (e != null) col.Add(e);
            }
            return col;
        }
        public void add_offsets(Dictionary<Edge, PartFeatureExtent> offsets, FlangeDefinition ff, EdgeCollection eds)
        {
            foreach (Edge ed in eds)
            {
                foreach (var item in offsets)
                {
                    if (!ed.Equals(item.Key)) continue;
                    var width = item.Value as OffsetWidthExtent;
                    if (width != null)
                    {
                        ff.SetOffsetWidthExtent(ed, ed.StartVertex, width.OffsetDistanceOne.Value, ed.StopVertex, width.OffsetDistanceTwo.Value);
                        continue;
                    }
                    var ext = item.Value as EdgeWidthExtent;
                    if (ext != null)
                    {
                        ff.SetEdgeWidthExtent(ed);
                        continue;
                    }
                    var cen = item.Value as CenteredWidthExtent;
                    if (cen != null)
                    {
                        ff.SetCenteredWidthExtent(ed, cen.Width.Value);
                        continue;
                    }
                }
            }
        }
        public FlangeFeature copy_flange(FlangeFeature c)
        {
            Dictionary<Edge, PartFeatureExtent> offsets;
            var eds = get_edges(c, out offsets);
            var def = smf.FlangeFeatures.CreateFlangeDefinition(eds, c.FeatureDimensions[2].Parameter.Expression,
                c.FeatureDimensions[1].Parameter.Expression);
            if (asm == null)
                c.SetEndOfPart(false);
            //def.Edges = eds;
            if (def.HeightExtentType == PartFeatureExtentEnum.kFlangeToExtent)
            {
                var ext = c.Definition.HeightExtent as ToHeightExtent;
                def = smf.FlangeFeatures.CreateFlangeDefinition(eds, c.FeatureDimensions[2].Parameter.Name, c.FeatureDimensions[1].Parameter.Name);
                def.ApplyAutoMitering = c.Definition.ApplyAutoMitering;
                def.BendOptions.BendReliefDepth = c.Definition.BendOptions.BendReliefDepth;
                def.BendOptions.BendReliefShape = c.Definition.BendOptions.BendReliefShape;
                def.BendOptions.BendReliefWidth = c.Definition.BendOptions.BendReliefWidth;
                def.BendOptions.BendTransition = c.Definition.BendOptions.BendTransition;
                def.BendOptions.BendTransitionArcRadius = c.Definition.BendOptions.BendTransitionArcRadius;
                def.SetToHeightExtent(ext.ToEntity, ext.Offset);
            }
            add_offsets(offsets, def, eds);
            var f = smf.FlangeFeatures.Add(def);

            //fill_parameters(f, c);
            smcdFrom.SetEndOfPartToTopOrBottom(false);
            return f;
        }
        public FlangeFeature copy_face_flange(FaceFeature c)
        {
            BendFeature bf = c.BendFeature;
            Dictionary<Edge, PartFeatureExtent> offsets;
            var col = I.app.TransientObjects.CreateEdgeCollection();
            List<Edge> eds = new List<Edge>(); List<Edge> ed = new List<Edge>();
            offsets = new Dictionary<Edge, PartFeatureExtent>();
            dynamic br = bf.Definition.BendRadius;
            sb = getBody(sb_name);
            eds = getEdges(bf.Faces, br.Value, ref ed);
            foreach (var item in eds)
            {
                col.Add(item);
            }

            var def = smf.FlangeFeatures.CreateFlangeDefinition(col, c.FeatureDimensions[2].Parameter.Expression,
                c.FeatureDimensions[1].Parameter.Expression);
            if (asm == null)
                c.SetEndOfPart(false);
            //def.Edges = eds;
            //add_offsets(offsets, def);
            var f = smf.FlangeFeatures.Add(def);
            //fill_parameters(f, c);
            return f;
        }
        public HemFeature copy_Hem(HemFeature c)
        {
            EdgeCollection eds;
            if (c.Definition.Edges == null)
            {
                var ed = find_Edge(c.Faces, c as PartFeature);
                eds = I.app.TransientObjects.CreateEdgeCollection();
                eds.Add(ed);
                //form.Close();
                //var ed = I.app.CommandManager.Pick(SelectionFilterEnum.kPartEdgeFilter, "Выберите ребро:");
                //eds = get_edges(u.gets<Edge>(new[] { ed }, fi => true));
                //form.Show();
            }
            else
            {
                eds = get_edges(u.gets<Edge>(c.Definition.Edges, fi => true));
            }
            c.SetEndOfPart(false);
            var def = smf.HemFeatures.CreateHemDefinition(eds);
            if (c.Definition.HemType == HemTypeEnum.kSingleHemType)
            {
                var ht = c.Definition.HemTypeDefinition as SingleHemDefinition;
                def.SetSingleHemType(ht.Gap, ht.Length.Name);
            }
            var f = smf.HemFeatures.Add(def);
            f.SetAffectedBodies(get_body());
            return f;
        }
        public Edge find_edge(Faces fs, Edge e)
        {
            Edge ret = null;
            var mp = u.midPt(e);
            var f = copy.lst(mp);
            if (mtx != null)
                copy.tr(f);
            foreach (Face item in fs)
            {
                if (item.SurfaceType != SurfaceTypeEnum.kPlaneSurface) continue;
                var dist = u.distToPlane(item.Geometry as Plane, f[0] as Point);
                if (Math.Abs(dist) < 0.001)
                {
                    double d = 10000;
                    foreach (Edge ed in item.Edges)
                    {
                        if (ed.CurveType != CurveTypeEnum.kLineCurve) continue;
                        var vec = u.midPt(ed).VectorTo(f[0] as Point);
                        if (vec.Length < d) { d = vec.Length; ret = ed; }
                    }
                    return ret;
                }
            }
            return null;
        }
        public Edge find_edge(Edges eds, Edge e, double tol)
        {
            var mp = u.midPt(e);
            var f = copy.lst(mp);
            if (mtx != null)
                copy.tr(f);
            foreach (Edge item in eds)
            {
                if (item.CurveType != CurveTypeEnum.kLineCurve) continue;
                if (u.isPointOnEdge(f[0] as Point, item, tol + 0.01)) return item;
            }
            return null;
        }
        public Edge find_Edge(Faces faces, PartFeature c)
        {
            var uv = I.CUV(1, 0, 0);
            var f = u.gets<Face>(faces, fi => fi.SurfaceType == SurfaceTypeEnum.kCylinderSurface).OrderBy(el => ((Cylinder)el.Geometry).Radius).First();
            Edge ret = null;
            foreach (Edge e in f.Edges)
            {
                Face f1 = e.Faces[1], f2 = e.Faces[2];
                if (f1.CreatedByFeature != f2.CreatedByFeature)
                {
                    f1 = f1.CreatedByFeature == c ? f1 : f2;
                    return find_edge(sb.Faces, e);
                    //u.higlight(doc, f1, f2, null, null);
                }
            }
            return null;
        }
        public void fill_parameters(FlangeFeature f1, FlangeFeature f2)
        {
            for (int i = 1; i < f1.Parameters.Count; i++)
            {
                if (f1.Parameters[i].Expression != f2.Parameters[i].Expression)
                    f1.Parameters[i].Value = f2.Parameters[i].Value;
            }
        }
        public List<Edge> get_edges(Faces fs)
        {
            List<Edge> eds = new List<Edge>();
            foreach (Face item in fs)
            {
                eds.Add(get_edge(item));
            }
            return eds;
        }
        public Edge get_edge(Face f)
        {
            List<Face> fs = new List<Face>();
            var eds = u.gets<Edge>(f.Edges, fi => fi.CurveType == CurveTypeEnum.kLineCurve);
            foreach (var ed in eds)
            {
                foreach (Face face in sb.Faces)
                {
                    if (face.SurfaceType != SurfaceTypeEnum.kPlaneSurface) continue;
                
                    var mp = u.midPt(ed);
                    var fi = copy.lst(mp);
                    if (mtx != null)
                        copy.tr(fi);
                    var d = u.distToPlane(face.Geometry as Plane, fi[0] as Point);
                    if (d < 0.01)
                    {
                        if (face.Evaluator.RangeBox.Contains(fi[0] as Point))
                        {
                            fs.Add(face); break;
                        }
                    }
                }
            }
            if (fs.Count == 2)
            {
                //u.higlight(doc, fs[0], fs[1], null, null);
                //return null;
                foreach (Edge item in fs[0].Edges)
                {
                    if (item.Faces[1] == fs[1] || item.Faces[2] == fs[1])
                        return item;
                }
            }
            return null;
        }
        public CornerRoundFeature copy_fillet(CornerRoundFeature c)
        {
            c.SetEndOfPart(true);
            EdgeCollection eds = null;
            var esi = c.Definition.EdgeSetItem[1];
            if (esi.Edges == null)
            {
                var es = get_edges(c.Faces);
                //return null;
                eds = I.app.TransientObjects.CreateEdgeCollection();
                foreach (var item in es)
                {
                    if (item == null) continue;
                    eds.Add(item);
                }
            }
            else
            eds = get_edges(u.gets<Edge>(esi.Edges, fi => true));
            if (asm == null)
                c.SetEndOfPart(false);
            var def = smf.CornerRoundFeatures.CreateCornerRoundDefinition(eds, get_value(((PartFeature)c), c.Parameters.Count));
            var f = smf.CornerRoundFeatures.Add(def);
            f.SetAffectedBodies(get_body());
            if (smcdFrom != null)
                smcdFrom.SetEndOfPartToTopOrBottom(false);
            return f;
        }
        public CornerChamferFeature copy_chamfer(CornerChamferFeature c)
        {
            var eds = get_edges(u.gets<Edge>(c.Definition.CornerEdges, fi => true));
            if (eds.Count == 0) return null;
            if (asm == null)
                c.SetEndOfPart(false);
            var def = smf.CornerChamferFeatures.CreateCornerChamferDefinition(eds, get_value(((PartFeature)c), c.Parameters.Count));
            var f = smf.CornerChamferFeatures.Add(def);
            f.SetAffectedBodies(get_body());
            return f;
        }
        public ObjectCollection get_features(IEnumerable<PartFeature> lst)
        {
            var col = I.COC();
            sb = getBody(sb_name);
            foreach (PartFeature pf in lst)
            {
                var f = copy.lst(pf.RangeBox.MinPoint);
                if (mtx != null)
                    copy.tr(f);
                foreach (PartFeature item in sb.AffectedByFeatures)
                {
                    if (item.RangeBox == null) continue;
                    if (u.eq(item.RangeBox.MinPoint, f[0] as Point, 1.5))
                    {
                        col.Add(item);
                    }
                }
            }
            return col;
        }
        
        public MirrorFeature copy_mirror(MirrorFeature c)
        {
            try
            {
                //var def = c.Definition.Copy();
                //def.AffectedBodies = get_body();
               
                var feats = u.gets<PartFeature>(c.ParentFeatures, fi => fi is PartFeature);
                //def.ParentFeatures = get_features(feats);
                var col = get_features(feats);
                var pl = copy.findPlane(c.Definition.MirrorPlaneEntity);
                var def = smf.MirrorFeatures.CreateDefinition(col, pl, PatternComputeTypeEnum.kIdenticalCompute);
                if (asm == null)
                    c.SetEndOfPart(false);
                def.AffectedBodies = get_body();
                var f = smf.MirrorFeatures.AddByDefinition(def);
                return f;
            }
            catch (Exception)
            {
                return null;
            }
            return null;
        }
        public RectangularPatternFeature copy_arr(RectangularPatternFeature c)
        {
            var feats = u.gets<PartFeature>(c.ParentFeatures, fi => fi is PartFeature);
            //def.ParentFeatures = get_features(feats);
            object ax = copy.findAxis(c.Definition.XDirectionEntity);
            var def1 = smf.RectangularPatternFeatures.CreateDefinition(get_features(feats), ax, c.Definition.NaturalXDirection, c.XCount.Value, c.XSpacing.Value,
                c.XDirectionSpacingType);
            def1.XDirectionMidPlanePattern = c.Definition.XDirectionMidPlanePattern;
            if (asm == null)
                c.SetEndOfPart(false);
            if (c.Definition.YDirectionEntity != null)
            {
                object ay = copy.findAxis(c.Definition.YDirectionEntity);
                def1.YDirectionEntity = ay;
                def1.NaturalYDirection = c.Definition.NaturalYDirection;
                def1.YCount = c.YCount.Value;
                def1.YSpacing = c.YSpacing.Value;
                def1.YDirectionMidPlanePattern = c.Definition.YDirectionMidPlanePattern;
                def1.YDirectionSpacingType = c.Definition.YDirectionSpacingType;

            }
            def1.AffectedBodies = get_body();
            var f = smf.RectangularPatternFeatures.AddByDefinition(def1);
            //f.SetAffectedBodies(get_body());
            return f;
        }

        public CircularPatternFeature copy_arr(CircularPatternFeature c)
        {
            var feats = u.gets<PartFeature>(c.ParentFeatures, fi => fi is PartFeature);
            //def.ParentFeatures = get_features(feats);
            object ax = copy.findAxis(c.AxisEntity);
            var def1 = smf.CircularPatternFeatures.CreateDefinition(get_features(feats), ax, c.NaturalAxisDirection, c.Count.Value, c.Angle.Value, c.FitWithinAngle);
            if (asm == null)
                c.SetEndOfPart(false);
            def1.AffectedBodies = get_body();
            var f = smf.CircularPatternFeatures.AddByDefinition(def1);
            //f.SetAffectedBodies(get_body());
            return f;
        }

        public FaceFeature copy_face_feature(FaceFeature ff, bool n = true)
        {
            if (ff.BendFeature != null)
            {
                copy_face_flange(ff);
                return null;
            }
            if (asm == null)
                ff.SetEndOfPart(false);
            var def = ff.Definition;
            var se = def.Profile[1][1].SketchEntity;
            ps = se.Parent as PlanarSketch;
            ps_new = copy.checkSketch(ps);
            //copy_sketch("face", def);
            var pr = get_profile(ff as PartFeature);
            var def1 = smf.FaceFeatures.CreateFaceFeatureDefinition(pr);
            def1.Direction = def.Direction;
#if INV14
#else
            if (n)
                def1.Operation = PartFeatureOperationEnum.kNewBodyOperation;
            else
            {
                def1.Operation = PartFeatureOperationEnum.kJoinOperation;
                def1.AffectedBodies = get_body();
            }
#endif
            var f = smf.FaceFeatures.Add(def1);
            if (n)
            {
                set_style(ff.SurfaceBody, f.SurfaceBody);
                sb = f.SurfaceBody;
            }
            return f;
        }
        public SurfaceBody getBody(string name)
        {
            return u.get<SurfaceBody>(smcd.SurfaceBodies, f => f.Name == name);
        }
        public ObjectCollection get_body()
        {
            sb = getBody(sb_name);
            var col = I.COC();
            col.Add(sb);
            return col;
        }
        public ObjectCollection get_pts(HoleFeature hf)
        {
            var col = I.COC();
            var hpts = hf.HoleCenterPoints;
            foreach (SketchPoint pt in hpts)
            {
                col.Add(copy.getPoint(pt as SketchEntity, ObjectTypeEnum.kSketchPointObject));
            }

            //var pts = u.gets<SketchPoint>(ps_new.SketchPoints, f => f.HoleCenter);
            //foreach (var item in pts)
            //{
            //    col.Add(item);
            //}
            return col;
        }
        public CutFeature copy_cut_feature(CutFeature c)
        {
            if (asm == null)
                c.SetEndOfPart(false);
            var def = c.Definition;
            var se = def.Profile[1][1].SketchEntity;
            ps = se.Parent as PlanarSketch;
            ps_new = copy.checkSketch(ps);
            //copy_sketch("cut", def);
            var pr = get_profile(c as PartFeature);
            var def1 = smf.CutFeatures.CreateCutDefinition(pr);
            var ext = def.Extent;
            if (def.ExtentType == PartFeatureExtentEnum.kDistanceExtent)
            {
                var d = ext as DistanceExtent;
                def1.SetDistanceExtent(d.Distance.Value, d.Direction);
            }
            if (def.CutAcrossBends)
            {
                var b = c.SurfaceBody;
                var smbd = b.ComponentDefinition as SheetMetalComponentDefinition;
                if (smbd != null)
                {
                    def1.SetCutAcrossBendsExtent(smbd.Thickness.ModelValue);
                }
            }
            c.SurfaceBody.Visible = false;
            var f = smf.CutFeatures.Add(def1);
            f.SetAffectedBodies(get_body());
            c.SurfaceBody.Visible = true;
            return f;
        }
        public ExtrudeFeature copy_extrude(ExtrudeFeature c)
        {
            if (asm == null)
                c.SetEndOfPart(false);
            var def = c.Definition;
            var def1 = c.Definition.Copy();
            var se = def.Profile[1][1].SketchEntity;
            ps = se.Parent as PlanarSketch;
            ps_new = copy.checkSketch(ps);
            //copy_sketch("extrude", def);
            var pr = get_profile(c as PartFeature);
            def1.Profile = pr;
            var f = smf.ExtrudeFeatures.Add(def1);
            f.SetAffectedBodies(get_body());
            return f;
        }
        public object get_value(PartFeature pf, int num)
        {
            if (asm == null)
                return pf.Parameters[num].Name;
            else
                return pf.Parameters[num].Value;
        }
        public HoleFeature copy_hole(HoleFeature hf)
        {
            if (hf.HoleCenterPoints.Count == 0) return null;
            if (asm == null)
                hf.SetEndOfPart(false);
            ps_new = copy.checkSketch(hf.Sketch);
            Plane pl1 = ps_new.PlanarEntityGeometry, pl2 = hf.Sketch.PlanarEntityGeometry;
            var el = copy.lst(pl2.Normal);
            if (mtx != null) 
                copy.tr(el);
            dynamic ext = hf.Extent;
            PartFeatureExtentDirectionEnum dir;
            dir = (PartFeatureExtentDirectionEnum)ext.Direction;
            if (!u.isCollinear(pl1.Normal, el[0] as UnitVector))
            {
                dir = (PartFeatureExtentDirectionEnum)ext.Direction == PartFeatureExtentDirectionEnum.kNegativeExtentDirection ?
                            PartFeatureExtentDirectionEnum.kPositiveExtentDirection : PartFeatureExtentDirectionEnum.kNegativeExtentDirection;
            }
            //copy_sketch(hf);
            var def = smf.HoleFeatures.CreateSketchPlacementDefinition(get_pts(hf));
            HoleFeature f = null;
            if (hf.HoleType == HoleTypeEnum.kDrilledHole && hf.ExtentType == PartFeatureExtentEnum.kDistanceExtent)
            {
                try
                {
                    //DistanceExtent de = (DistanceExtent)hf.Extent;
                    //var dir = de.Direction;
                    f = smf.HoleFeatures.AddDrilledByDistanceExtent(def, get_value(hf as PartFeature, hf.Parameters.Count), get_value(hf as PartFeature, hf.Parameters.Count - 1),
                        dir);
                }
                catch (Exception)
                {

                    throw;
                }

            }
            else if (hf.HoleType == HoleTypeEnum.kDrilledHole && hf.ExtentType == PartFeatureExtentEnum.kThroughAllExtent)
            {
                object diam = "";
                if (asm == null)
                    diam = get_value(hf as PartFeature, hf.Parameters.Count);
                else
                    diam = hf.Parameters[hf.Parameters.Count].Expression;
                f = smf.HoleFeatures.AddDrilledByThroughAllExtent(def, diam, dir);

            }
            if (f.SurfaceBodies.Count == 0)
                f.SetAffectedBodies(get_body());
            return f;
        }
    }

    class AsmRepl
    {
        string path;
        AssemblyComponentDefinition def;
        Document doc;
        List<string> lst;

        public AsmRepl(Document doc)
        {
            this.doc = doc;
            path = file.p(doc.FullDocumentName);
            lst = findPathes(path);
            run();
        }

        static public List<string> findPathes(string path)
        {
            List<string> lst = new List<string>();
            var ie = System.IO.Directory.EnumerateDirectories(path);
            foreach (var item in ie)
            {
                var name = file.name(item);
                if (name.StartsWith("КЭВ")) lst.Add(item);
            }
            return lst;
        }
        public void replProp(IEnumerable<string> files, string dn)
        {
            Regex reg = new Regex(@"\w(\d+)\w", RegexOptions.IgnoreCase);
            Regex reg1 = new Regex(@"\d+", RegexOptions.IgnoreCase);

            foreach (var item in files)
            {
                var path = file.p(item);
                var name = file.nameWithExt(item);
                var m = reg.Match(name);
                if (!m.Success) continue;
                var dec = m.Groups[1].Value;

                var spl = name.Split('.');
                spl[0] = reg1.Replace(spl[0], dec);
                //spl[0] = "";
                var repl = path + String.Join(".", spl);
                if (file.check(repl)) continue;
                System.IO.File.Move(item, repl);
                var asm = I.open(repl);
                u.changeProp(asm, "Type", dn);
                asm.Update2();
                asm.Save2(false);
            }
        }
        public void run()
        {
            foreach (var p in lst)
            {
                var ie = file.getFiles(p, ".iam");
                var dn = file.nameWithExt(p);
                replProp(ie, dn);
                var files = System.IO.Directory.EnumerateFiles(p).Where(f => (f.EndsWith(".iam") || f.EndsWith(".ipt")));
                foreach (var item in ie)
                {
                    var asm = I.open(item);
                    repl(asm, files);
                }
            }
        }
        public void repl(Document doc, IEnumerable<string> files)
        {
            Regex reg = new Regex(@"\d\d\.\d\d\d-*\d*", RegexOptions.IgnoreCase);
            if (doc.RequiresUpdate) doc.Update2(false);
            foreach (DocumentDescriptor item in doc.ReferencedDocumentDescriptors)
            {
                var name = file.nameWithExt(item.FullDocumentName);
                var m1 = reg.Match(name);
                if (m1.Value == null || m1.Value == "") continue;
                foreach (var f in files)
                {
                    var name1 = file.name(f);
                    var m2 = reg.Match(name1);
                    if (m2.Value == null || m2.Value == "") continue;
                    if (m1.Value == m2.Value)
                    {
                        if (item.FullDocumentName != f)
                        {
                            item.ReferencedFileDescriptor.ReplaceReference(f);
                        }
                    }
                }
            }
            if (doc.Dirty) doc.Save2();
        }
        public void replace(Document doc, IEnumerable<string> files)
        {
            Regex reg = new Regex(@"\d\d\.\d\d\d-*\d*", RegexOptions.IgnoreCase);
            if (doc.RequiresUpdate) doc.Update2(false);
            def = ((AssemblyDocument)doc).ComponentDefinition;
            List<string> filter = new List<string>();
            Dictionary<ComponentOccurrence, string> repl = new Dictionary<ComponentOccurrence, string>();
            foreach (ComponentOccurrence occ in def.Occurrences)
            {
                var name = file.name(occ.ReferencedDocumentDescriptor.FullDocumentName);
                var m1 = reg.Match(name);
                if (m1.Value == null || m1.Value == "") continue;
                foreach (var f in files)
                {
                    var name1 = file.name(f);
                    var m2 = reg.Match(name1);
                    if (m2.Value == null || m2.Value == "") continue;
                    if (m1.Value == m2.Value)
                    {
                        if (filter.Contains(name)) continue;
                        if (f != occ.ReferencedDocumentDescriptor.FullDocumentName)
                        {
                            repl.Add(occ, f);
                            //doc.Update();
                            filter.Add(name);
                        }
                    }
                }
            }
            foreach (var item in repl)
            {
                if (item.Key.ReferencedDocumentDescriptor.FullDocumentName != item.Value)
                    item.Key.Replace2(item.Value, true);
            }
            if (doc.Dirty)
                doc.Save2();
        }
    }

    public class DeleteEntities
    {
        Document doc;
        DrawingDocument drw;
        SelectSet ss;
        PartDocument part;
        Dictionary<SketchEntity, SketchEntity> ent_constr = new Dictionary<SketchEntity, SketchEntity>();
        Sketch ps, ps2;
        public DeleteEntities(Document doc)
        {
            this.doc = doc;
            if (doc.DocumentType == DocumentTypeEnum.kPartDocumentObject)
            {
                part = doc as PartDocument;
                ss = part.SelectSet;
                //if (ss.Count != 1) return;
                deleteProjection();
            }
            else if (doc.DocumentType == DocumentTypeEnum.kDrawingDocumentObject)
            {
                drw = doc as DrawingDocument;
                deleteDims();
                deleteAxis();
                deleteMarks();
            }
        }
        public void deleteDims()
        {
            var dims = u.gets<DrawingDimension>(drw.ActiveSheet.DrawingDimensions, f => true);
            foreach (DrawingDimension item in dims)
            {
                if (!item.Attached) item.Delete();
            }
        }
        public void deleteAxis()
        {
            var cl = u.gets<Centerline>(drw.ActiveSheet.Centerlines, f => true);
            foreach (Centerline item in cl)
            {
                if (!item.Attached) item.Delete();
            }
        }
        public void deleteMarks()
        {
            var cl = u.gets<Centerline>(drw.ActiveSheet.Centermarks, f => true);
            foreach (Centermark item in cl)
            {
                if (!item.Attached) item.Delete();
            }
        }
        public void deleteLeaders()
        {
            var ls = u.gets<LeaderNote>(drw.ActiveSheet.DrawingNotes, f => f.Type == ObjectTypeEnum.kLeaderNoteObject);
            foreach (LeaderNote item in ls)
            {
                item.Delete();
            }
        }
        public void deleteProjection()
        {
            ps = null;
            IEnumerable<SketchEntity> ie;
            if (ss.Count == 1)
            {
                var se = ss[1] as SketchEntity;
                if (se == null) return;
                ps = se.Parent;
                ie = u.gets<SketchEntity>(ps.SketchEntities, f => true);
            }
            else ie = u.gets<SketchEntity>(ss, f => true);
            get_constr(ie);
            if (ps == null) ps = ie.ElementAt(0).Parent;

            foreach (var e in ie)
            {
                if (e.Reference)
                {
                    e.Reference = false;
                }
            }
            constrainEnt();
            ps2.Visible = true;
        }
        public void get_constr(IEnumerable<SketchEntity> ents)
        {
            foreach (SketchEntity se in ents)
            {
                var r_ent = se.ReferencedEntity as SketchEntity;
                if (r_ent == null) continue;
                ent_constr.Add(se, r_ent);
            }
        }
        public SketchEntity get_se(SketchEntity se)
        {
            foreach (SketchEntity item in ps.SketchEntities)
            {
                if (u.eq(item.RangeBox, se.RangeBox, 2))
                    return item;
            }
            return null;
        }
        public void with_se(KeyValuePair<SketchEntity,SketchEntity> p)
        {
            if (ps2 == null) ps2 = p.Value.Parent;
            SketchEntity se2 = p.Value;
            foreach (var c in se2.Constraints)
            {
                GeometricConstraint g = c as GeometricConstraint;
                if (g != null)
                    add_geom(g);
                DimensionConstraint d = c as DimensionConstraint;
                if (d != null)
                    add_dim(d);

            }
        }
        public bool check(SketchEntity se, SketchEntity se2)
        {
            foreach (GeometricConstraint item in ps.GeometricConstraints)
            {
                dynamic d = item;
                if (d is PerpendicularConstraint || d is ParallelConstraint || d is CollinearConstraint)
                {
                    if ((d.EntityOne == se || d.EntityTwo == se) &&
                        (d.EntityOne == se2 || d.EntityTwo == se2)) return true;
                }
                else if (d is HorizontalAlignConstraint || d is VerticalAlignConstraint)
                {
                    if ((d.PointOne == se && d.PointTwo == se2) ||
                        (d.PointOne == se2 && d.PointTwo == se)) return true;
                }
                else if (d is EqualLengthConstraint)
                {
                    if ((d.LineOne == se && d.LineTwo == se2) ||
                        (d.LineOne == se2 && d.LineTwo == se)) return true;
                }
            }
            return false;
        }
        public bool check<T>(SketchEntity se)
        {
            foreach (var item in se.Constraints)
            {
                if (item.GetType() == typeof(T)) return true;
            }
            return false;
        }
        public void add_dim(DimensionConstraint c)
        {
            SketchEntity sen1, sen2;
            dynamic con = c;
            try
            {
                if (c is TwoPointDistanceDimConstraint)
                {
                    sen1 = get_se(con.PointOne); sen2 = get_se(con.PointTwo);
                    var d = ps.DimensionConstraints.AddTwoPointDistance((SketchPoint)sen1, (SketchPoint)sen2, con.Orientation, con.TextPoint) as TwoPointDistanceDimConstraint;
                    d.Parameter.Expression = con.Parameter.Name;
                }
                else if (c is OffsetDimConstraint)
                {
                    sen1 = get_se(con.Line); sen2 = get_se(con.Entity);
                    var d = ps.DimensionConstraints.AddOffset((SketchLine)sen1, sen2, con.TextPoint, con.LinearDiameter);
                    d.Parameter.Expression = con.Parameter.Name;
                }
                else if (c is DiameterDimConstraint)
                {
                    sen1 = get_se(con.Entity);
                    var d = ps.DimensionConstraints.AddDiameter(sen1, con.TextPoint);
                    d.Parameter.Expression = con.Parameter.Name;
                }
                else if (c is RadiusDimConstraint)
                {
                    sen1 = get_se(con.Entity);
                    var d = ps.DimensionConstraints.AddRadius(sen1, con.TextPoint);
                    d.Parameter.Expression = con.Parameter.Name;
                }
            }
            catch (Exception)
            {
            }
        }
        public void add_geom(GeometricConstraint c)
        {
            SketchEntity sen1, sen2;
            dynamic con = c;
            try
            {
                if (c is ParallelConstraint)
                {
                    sen1 = get_se(con.EntityOne); sen2 = get_se(con.EntityTwo);
                    if (check(sen1, sen2)) return;
                    ps.GeometricConstraints.AddParallel(sen1, sen2);
                }
                else if (c is PerpendicularConstraint)
                {
                    sen1 = get_se(con.EntityOne); sen2 = get_se(con.EntityTwo);
                    if (check(sen1, sen2)) return;
                    ps.GeometricConstraints.AddPerpendicular(sen1, sen2);
                }
                else if (c is CollinearConstraint)
                {
                    sen1 = get_se(con.EntityOne); sen2 = get_se(con.EntityTwo);
                    if (check(sen1, sen2)) return;
                    ps.GeometricConstraints.AddCollinear(sen1, sen2);
                }
                else if (c is HorizontalConstraint)
                {
                    sen1 = get_se(con.Entity);
                    if (check<HorizontalConstraint>(sen1)) return;
                    ps.GeometricConstraints.AddHorizontal(sen1);
                }
                else if (c is VerticalConstraint)
                {
                    sen1 = get_se(con.Entity);
                    if (check<VerticalConstraint>(sen1)) return;
                    ps.GeometricConstraints.AddVertical(sen1);
                }
                else if (c is HorizontalAlignConstraint)
                {
                    sen1 = get_se(con.PointOne); sen2 = get_se(con.PointTwo);
                    if (check(sen1, sen2)) return;
                    ps.GeometricConstraints.AddHorizontalAlign((SketchPoint)sen1, (SketchPoint)sen2);
                }
                else if (c is VerticalAlignConstraint)
                {
                    sen1 = get_se(con.PointOne); sen2 = get_se(con.PointTwo);
                    if (check(sen1, sen2)) return;
                    ps.GeometricConstraints.AddVerticalAlign((SketchPoint)sen1, (SketchPoint)sen2);
                }
                else if (c is EqualLengthConstraint)
                {
                    sen1 = get_se(con.LineOne); sen2 = get_se(con.LineTwo);
                    if (check(sen1, sen2)) return;
                    ps.GeometricConstraints.AddEqualLength((SketchLine)sen1, (SketchLine)sen2);
                }
                else if (c is TangentSketchConstraint)
                {
                    sen1 = get_se(con.EntityOne); sen2 = get_se(con.EntityTwo);
                    ps.GeometricConstraints.AddTangent(sen1, sen2);
                }
            }
            catch (Exception)
            {
            }
            
        }
        public void constrainEnt()
        {
            foreach (KeyValuePair<SketchEntity, SketchEntity> p in ent_constr)
            {
                with_se(p);
            }
        }
    }
    public class OpenFile
    {
        public bool o = true;
        Document doc;
        string path;
        SelectSet ss;
        List<string> pathes = new List<string>();
        List<string> fil = new List<string>() { "OldVersions", "PDF", "DXF", "ContentCenter" };
        Dictionary<string, List<string>> files = new Dictionary<string, List<string>>();
        List<Document> docs = new List<Document>();
        public OpenFile(Document doc)
        {
            this.doc = doc;
            if (doc.DocumentType != DocumentTypeEnum.kPartDocumentObject) { of(); return; }
            init();
        }
        public void of()
        {
            string p = doc.FullFileName;
            file.open(p);
        }
        public void init()
        {
            ss = doc.SelectSet;
            path = file.p(doc.FullDocumentName);
        }
        public void findPath(string p)
        {
            pathes.Add(p);
            var path = System.IO.Directory.EnumerateDirectories(p, "*", System.IO.SearchOption.AllDirectories).Where(d => filter(d));
            pathes.AddRange(path);
            addFiles();
        }
        public bool filter(string s)
        {
            foreach (var item in fil)
            {
                if (s.ToLower().IndexOf(item.ToLower()) != -1) return false;
            }
            return true;
        }
        public void openFiles()
        {
            findPath(path);
            foreach (var item in files)
            {
                foreach (var v in item.Value)
                {
                    var d = I.open(v, false, false);
                    var pcd = I.getPCD(d);
                    if (pcd.SurfaceBodies.Count != 1) continue;
                    if (pcd.SurfaceBodies[1].CreatedByFeature.Type != ObjectTypeEnum.kReferenceFeatureObject) continue;
                    docs.Add(d);
                }
            }
        }
        public void addFiles()
        {
            foreach (var item in pathes)
            {
                if (files.ContainsKey(item)) continue;
                var ipt = file.getFiles(item, ".ipt").ToList();
                //var iam = file.getFiles(item, ".iam");
                //ipt.AddRange(iam);
                files[item] = ipt;
            }
        }
        public void open()
        {
            
            var bodies = CreateComponent.GetBodies(ss);
            if (bodies.Count == 0) { of(); return; }
            openFiles();
            foreach (SurfaceBody sb in bodies)
            {
                if (!sb.Exported) continue;
                var n = check(sb);
                if (n != "")
                {
                    I.open(n, true, false);
                }
            }
        }
        public string check(SurfaceBody sb)
        {
            foreach (Document d in docs)
            {
                var pcd = I.getPCD(d);
                var rf = pcd.SurfaceBodies[1].CreatedByFeature as ReferenceFeature;
                if (rf.ReferencedEntity.Equals(sb)) return d.FullDocumentName;
            }
            return "";
        }

    }
    public class Repair
    {
        string p, dn, t;
        Document doc;
        IEnumerable<string> files;
        List<string> fil = new List<string>() { "OldVersions" };
        List<Document> docs = new List<Document>();
        public Repair(Document doc)
        {
            this.doc = doc;
            init();
            add();
        }
        public void init()
        {
            p = file.p(doc.FullDocumentName);
            files = file.getFiles(p, ".ipt", System.IO.SearchOption.AllDirectories).Where(f => filter(f));
            dn = u.getPropValue(doc, "DecNumber");
            if (dn == "") dn = "00.001";
            t = u.getPropValue(doc, "Type");
        }
        public bool filter(string fn)
        {
            foreach (var item in fil)
            {
                if (fn.ToLower().IndexOf(item.ToLower()) != -1) return false;
            }
            return true;
        }
        public void add()
        {
            foreach (var item in files)
            {
                getDoc(item);
            }
            foreach (var item in docs)
            {
                var n = getDN(item);
                var spl = n.Split('$');
                var num = u.getDN(spl, dn, 1);
                u.addProp(item, "DecNumber", num);
                u.addProp(item, "Type", t);
                doc.Save();
            }
        }
        public string getDN(Document doc)
        {
            var def = I.getPCD(doc);
            if (def.SurfaceBodies.Count != 1) return "";
            var sb = def.SurfaceBodies[1];
            if (sb.CreatedByFeature.Type != ObjectTypeEnum.kReferenceFeatureObject) return "";
            var rf = sb.CreatedByFeature as ReferenceFeature;
            dynamic ent = rf.ReferencedEntity;
            return ent.Name;
        }
        public void getDoc(string fn)
        {
            var doc = I.open(fn);
            foreach (var item in doc.ReferencedDocuments)
            {
                if (item.Equals(this.doc)) docs.Add(doc);
            }
        }
    }
    public class AddDeleteReplace
    {
        AssemblyDocument doc;
        AssemblyComponentDefinition def;
        string path;
        string name, fn;
        public MyForm f;
        MyXML xml;
        public CommandManager mgr;
        ComponentOccurrence occ;
        List<string> fil = new List<string>() { "OldVersions", "PDF", "DXF", "ContentCenter" };
        List<string> pathes = new List<string>();
        Dictionary<string, List<string>> files = new Dictionary<string, List<string>>();
        List<string> skip = new List<string>();

        public AddDeleteReplace(Document doc)
        {
            if (doc.DocumentType != DocumentTypeEnum.kAssemblyDocumentObject) return;
            this.doc = doc as AssemblyDocument;
            def = this.doc.ComponentDefinition;
            init();
        }
        public void init()
        {
            f = new MyForm("AddDeleteReplaceInterface.xml", "Сборка");
            var btn = new Btn(f.offsetX, f.offsetY, f.w, f.h, f.pt, f.f, ClickRun, "Выполнить", "run");
            Control l = f.f.Controls[f.f.Controls.Count - 2];
            btn.center(l, f.offsetY + 5);
            path = file.p(doc.FullDocumentName);
            mgr = I.app.CommandManager;
            f.cbs[1].TextChanged += change;
            addXML();
        }

        private void change(object sender, EventArgs e)
        {
            ComboBox cb = sender as ComboBox;
            if (cb.Text.EndsWith(".xml"))
            {
                addXML();
            }
        }

        public void addXML()
        {
            fn = $"{path}{f.cbs[1].Text}";
            xml = new MyXML(fn, "head");
        }
        public bool run()
        {
            f.f.ShowDialog();
            if (f.f.DialogResult != DialogResult.OK) return false;
            if (f.cbs[0].Text == "Добавить")
            {
                add();
            } else if (f.cbs[0].Text == "Удалить")
            {
                delete();
            } else if (f.cbs[0].Text == "Заменить")
            {
                replace();
            }
            xml.save();
            return f.f.DialogResult == DialogResult.OK;
        }
        public ComponentOccurrence sel()
        {
            f.f.Hide();
            var occ = mgr.Pick(SelectionFilterEnum.kAssemblyOccurrenceFilter, "Выберите элемент") as ComponentOccurrence;
            return occ;
        }
        public void add()
        {
            var fs = u.OFD(path, multi: true);
            foreach (var item in fs.Split('|'))
            {
                var n = file.nameWithExt(item);
                XElement el = new XElement("Add");
                xml.addElem(el);
                MyXML.addAtt(el, "val", n);
            }
        }
        public void delete()
        {
            occ = sel();
            var n = occ.Name.Split(':')[0];
            XElement el = new XElement("Delete");
            xml.addElem(el);
            MyXML.addAtt(el, "val", n);
        }
        public void replace()
        {
            occ = sel();
            var n = occ.Name.Split(':')[0];
            var fs = u.OFD(path);
            XElement el = new XElement("Replace");
            xml.addElem(el);
            MyXML.addAtt(el, "find", n);
            MyXML.addAtt(el, "val", file.nameWithExt(fs));
        }
        public void add(XElement el)
        {
            name = find(el, path);
            if (name == "") return;
            var mtx = I.getMatrix();
            var oc = def.Occurrences.Add(name, mtx);
            u.joinIMate(oc);
        }
        public void delete(XElement el)
        {
            string v = MyXML.getAtt(el, "val");
            string all = MyXML.getAtt(el, "all");
            foreach (ComponentOccurrence occ in def.Occurrences)
            {
                var n = occ.Name;

                if (n.IndexOf(v) != -1)
                {
                    var acd = occ.Parent;
                    Parts.removeOcc(acd, occ);
                    if (all != "") return;
                }
            }
        }
        public void replace(XElement el)
        {
            string v = MyXML.getAtt(el, "find");
            var to = find(el, path);
            string all = MyXML.getAtt(el, "all");
            if (to == "") return;
            foreach (ComponentOccurrence occ in def.Occurrences)
            {
                var n = occ.Name;

                if (n.IndexOf(v) != -1)
                {
                    occ.Replace2(to, false);
                    if (all != "") return;
                }
            }
        }
        public void findPath(string p)
        {
            pathes.Add(p);
            var path = System.IO.Directory.EnumerateDirectories(p, "*", System.IO.SearchOption.AllDirectories).Where(d => filter(d));
            pathes.AddRange(path);
            addFiles();
        }
        public void addFiles()
        {
            foreach (var item in pathes)
            {
                if (files.ContainsKey(item)) continue;
                var ipt = file.getFiles(item, ".ipt").ToList();
                var iam = file.getFiles(item, ".iam");
                ipt.AddRange(iam);
                files[item] = ipt;
            }
        }
        public string find(string v, List<string> fs)
        {
            foreach (var item in fs)
            {
                var n = file.nameWithExt(item);
                if (n.IndexOf(v) != -1) return item;
            }
            return "";
        }
        public string find(string v, string path)
        {
            string fn = "";
            path = path.TrimEnd('\\');
            if (files.ContainsKey(path))
            {
                fn = find(v, files[path]);
                if (fn != "") return fn;
                skip.Add(path);
            }
            foreach (var item in files)
            {
                if (skip.Contains(item.Key)) continue;
                fn = find(v, item.Value);
                if (fn != "") return fn;
                skip.Add(item.Key);
            }
            return "";
        }
        public string find(XElement el, string path)
        {
            string v = MyXML.getAtt(el, "val");
            var f = find(v, path);
            if (f == "")
            {
                foreach (ProjectPath item in I.app.DesignProjectManager.ActiveDesignProject.LibraryPaths)
                {
                    findPath(item.Path);
                    f = find(v, item.Path);
                    if (f != "") return f;
                }
            }
            return f;
        }
        public bool filter(string s)
        {
            foreach (var item in fil)
            {
                if (s.ToLower().IndexOf(item.ToLower()) != -1) return false;
            }
            return true;
        }
        public void runXML()
        {
            foreach (var el in xml.elem.Elements())
            {
                if (el.Name == "Add") add(el);
                else if (el.Name == "Delete") delete(el);
                else if (el.Name == "Replace") replace(el);
            }
        }
        void ClickRun(object sender, EventArgs e)
        {
            string mainPath = I.app.DesignProjectManager.ActiveDesignProject.WorkspacePath;
            findPath(mainPath);
            
            foreach (Document item in I.app.Documents.VisibleDocuments)
            {
                if (item.DocumentType != DocumentTypeEnum.kAssemblyDocumentObject) continue;
                //if (item.Equals(doc)) continue;
                def = ((AssemblyDocument)item).ComponentDefinition;
                runXML();
            }
            f.f.DialogResult = DialogResult.Abort;
        }
        void ClickAdd(object sender, EventArgs e)
        {
            f.f.DialogResult = DialogResult.OK;
        }
    }
}
