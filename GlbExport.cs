using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Inventor;
using InvDoc;

namespace InvAddIn
{
    public enum GlbMode { Single, Assembly }

    public class GlbOptions
    {
        public GlbMode Mode = GlbMode.Assembly;
        public double Tolerance = 0.01;          // хордовое отклонение тесселяции, см
        public bool PlanarRetriangulate = true;  // плоские грани: только вершины контура
        public int PlanarMaxVertices = 200000;   // больше - оставляем тесселяцию Inventor
        public double Scale = 0.01;              // см -> м
        public bool Quantize = true;             // KHR_mesh_quantization: файл меньше, вид тот же (three.js/Babylon читают штатно)
        public int CircleSegments = 16;          // окружности (отверстия, трубы, контуры) не грубее этого числа сегментов
        public double SimplifyRatio = 0.0005;    // допуск упрощения сетки как доля диагонали модели (0 - без упрощения)
        public int SimplifyTimeLimitMs = 15000;  // предел времени упрощения одного тела
        public int DebugDumpTris = 150000;       // тела крупнее - исходная сетка в glb\_debug_*.txt (0 - не писать)
        public bool CullHidden = true;           // Single: убрать геометрию, не видимую снаружи
        public int CullDirections = 64;          // направлений обзора (равномерно по сфере)
        public double CullCell = 5;              // мм: ячейка карты видимых областей частично видимых граней
        public int CullResolution = 1024;        // разрешение z-буфера
        public double CullHoleSize = 20;         // мм: отверстия и щели меньше считаются закрытыми (перфорация)
        public double CullOpeningSize = 150;     // мм: при поиске внешних деталей закрыты и проёмы до этого размера
        public double ExternalMinShare = 0.03;   // деталь внешняя, если снаружи видна такая доля её поверхности
        public bool DropHoleWalls = true;        // Single: не выгружать стенки мелких отверстий (перфорация) -
        public double HoleWallMaxDepth = 3;      //   вогнутый цилиндр диаметром <= CullHoleSize и глубиной <= этого, мм
        public bool CullSeeThrough = false;      // 2-й проход: false - грани внешних деталей, видимые сквозь перфорацию, остаются
                                                 //   (сквозь отверстия видна ровная внутренняя сторона корпуса, а не фон)
        public bool TrimPartialFaces = true;    // обрезать частично видимые грани по видимым областям (меньше треугольников,
                                                 //   но сквозь перфорацию появляются "дыры" в внутренних поверхностях)
        public double SeeThroughTol = 0.5;       // мм: при поиске внешних деталей пиксель внутри заклеенного мелкого отверстия
                                                 //   (заклейка подняла глубину больше этого) не засчитывается
        public bool FinalCull = false;           // Single: финальная видимость по готовой сетке (подложки перфорации - преграда):
                                                 //   убирает то, что скрыто за перфорацией, а сквозь одиночные отверстия всё видно
        public bool HoleBacking = true;          // тёмная подложка в отверстиях перфорации (одиночные отверстия остаются сквозными)
        public double PerforationRadius = 3;     // отверстие - перфорация, если в пределах стольких своих диаметров
        public int PerforationNeighbors = 2;     //   есть не меньше стольких других отверстий
        public double SmallFaceArea = 200;       // мм²: скрытая грань меньше этого рядом с видимой всё равно выгружается (без дыр)
        public bool PerforatedDirectOnly = false; // перфорированный лист - только если виден напрямую (внутренняя сторона не нужна),
                                                  //   но тогда сквозь жалюзи/проёмы на её месте пятна
        public int PerforationMinHoles = 10;     //   и в группе соседствующих отверстий их не меньше этого (кронштейны, крепёж - не перфорация)
        public double[] BackingColor = { 0.02, 0.02, 0.02 };   // линейный цвет подложки
        public double CullKeepDepth = 15;        // мм: за "заклеенным" отверстием отбрасывается только то, что глубже -
                                                 //   мелкие наружные выемки и проточки (гермоввод) остаются
    }

    // Экспорт активного документа (деталь или сборка) в GLB
    // Single   - один узел, один меш, примитивы по материалам
    // Assembly - иерархия вхождений, каждое тело отдельным узлом, одинаковые тела - один меш
    public class toGLB
    {
        GlbOptions opt;
        GlbBuilder glb = new GlbBuilder();
        Dictionary<string, int> matIdx = new Dictionary<string, int>();
        Dictionary<string, GlbMesh> tessCache = new Dictionary<string, GlbMesh>();
        Dictionary<string, int> meshCache = new Dictionary<string, int>();
        GlbMesh merged;
        Dictionary<string, int> errors = new Dictionary<string, int>();
        Dictionary<string, Asset> matAssets = new Dictionary<string, Asset>();
        Dictionary<string, BodyStat> stats = new Dictionary<string, BodyStat>();

        class BodyStat
        {
            public string name;
            public int count, tris, verts, planarOk, planarFail, trisBefore;
            public long msFacets, msSimplify;
            public Dictionary<string, int> byType = new Dictionary<string, int>();
        }
        BodyStat cur;
        double simplifyTol;   // см
        string outDir;
        long cullMs = -1, coarseTris, coarseMs, trimmedTris, extMs, extTris;
        int hiddenBodies, hiddenFaces, partialFaces, internalParts, droppedWalls;
        bool inSingle;
        class ExtRow { public string name; public double share; public bool ext, byRule; public int tris; }
        List<ExtRow> extReport = new List<ExtRow>();

        // правила GlbParts.txt: "+подстрока" - всегда внешняя, "-подстрока" - всегда внутренняя, "#" - комментарий.
        // Подстрока ищется без учёта регистра в имени вхождения/тела и в имени файла детали; последнее совпавшее правило главнее
        // (сначала читается файл из папки аддина, потом из папки модели)
        List<KeyValuePair<bool, string>> partRules = new List<KeyValuePair<bool, string>>();

        void loadRules(string fn)
        {
            try
            {
                if (!System.IO.File.Exists(fn)) return;
                foreach (string raw in System.IO.File.ReadAllLines(fn, Encoding.UTF8))
                {
                    string l = raw.Trim();
                    if (l.Length < 2 || l[0] == '#') continue;
                    if (l[0] == '+' || l[0] == '-') partRules.Add(new KeyValuePair<bool, string>(l[0] == '+', l.Substring(1).Trim().ToLower()));
                }
            }
            catch (Exception e) { error("Правила " + fn, e); }
        }

        int rule(Inst t)
        {
            string a = (t.name ?? "").ToLower(), b = "";
            try { b = ((Document)t.def.Document).DisplayName.ToLower(); } catch { }
            int r = 0;
            foreach (var pr in partRules)
                if (pr.Value.Length > 0 && (a.Contains(pr.Value) || b.Contains(pr.Value))) r = pr.Key ? 1 : -1;
            return r;
        }
        Dictionary<string, int> planarWhy = new Dictionary<string, int>();

        static void inc(Dictionary<string, int> d, string k, int v = 1)
        {
            int c;
            d.TryGetValue(k, out c);
            d[k] = c + v;
        }

        void used(string key, PartComponentDefinition def, SurfaceBody b)
        {
            BodyStat st;
            if (!stats.TryGetValue(key, out st))
            {
                string dn = "";
                try { dn = ((Document)def.Document).DisplayName; } catch { }
                stats[key] = st = new BodyStat { name = dn + " / " + b.Name };
            }
            st.count++;
        }

        void error(string where, Exception e)
        {
            string k = where + ": " + e.GetType().Name + " - " + e.Message;
            int c;
            errors.TryGetValue(k, out c);
            errors[k] = c + 1;
        }

        public toGLB(GlbMode mode) : this(I.app.ActiveDocument, new GlbOptions { Mode = mode }) { }

        public toGLB(Document doc, GlbOptions opt)
        {
            this.opt = opt;
            glb.Quantize = opt.Quantize;
            if (doc == null) return;
            if (doc.DocumentType != DocumentTypeEnum.kPartDocumentObject &&
                doc.DocumentType != DocumentTypeEnum.kAssemblyDocumentObject) return;
            var sw = System.Diagnostics.Stopwatch.StartNew();
            string name = file.name(doc.FullFileName);
            try
            {
                Box box = doc is PartDocument ? ((PartDocument)doc).ComponentDefinition.RangeBox : ((AssemblyDocument)doc).ComponentDefinition.RangeBox;
                simplifyTol = box.MinPoint.DistanceTo(box.MaxPoint) * opt.SimplifyRatio;
            }
            catch (Exception e) { error("Габарит", e); }
            loadRules(I.p() + "\\GlbParts.txt");
            loadRules(file.p(doc.FullFileName) + "GlbParts.txt");
            string path = file.p(doc.FullFileName) + "glb";
            file.dir(path);
            outDir = path;
            int root = opt.Mode == GlbMode.Single ? single(doc, name) : assembly(doc, name);
            string fn = path + "\\" + name + (opt.Mode == GlbMode.Single ? "_single" : "") + ".glb";
            glb.save(fn, root);
            try { report(path + "\\" + name + (opt.Mode == GlbMode.Single ? "_single" : "") + "_report.txt"); } catch (Exception e) { error("Отчёт", e); }
            I.app.StatusBarText = "GLB: " + fn;
            string msg = string.Format("{0}\nВершин: {1}\nТреугольников: {2} (до упрощения {7})\nДопуск упрощения: {8:0.###} мм\nМешей: {3}, материалов: {4}\nРазмер: {5:0.0} КБ\nВремя: {6:0.0} с",
                fn, glb.vertices, glb.triangles, glb.meshCount, matIdx.Count, new FileInfo(fn).Length / 1024.0, sw.Elapsed.TotalSeconds,
                stats.Values.Sum(x => (long)x.trisBefore * (opt.Mode == GlbMode.Single ? x.count : 1)), simplifyTol * 10);
            if (cullMs >= 0)
                msg += string.Format("\nВнутренних деталей: {0} (поиск {1:0.0} с по {2} тр.)", internalParts, extMs / 1000.0, extTris) +
                    (finalRemoved >= 0 ? string.Format("\nФинальная видимость: убрано {0} тр. ({1:0.0} с)", finalRemoved, finalMs / 1000.0) : "") +
                    (droppedWalls > 0 ? "\nСтенок мелких отверстий не выгружено: " + droppedWalls + (backingDisks > 0 ? ", подложек: " + backingDisks : "") +
                        " (в грубой сетке " + coarseDisks + ")" : "") +
                    (rescuedFaces > 0 ? "\nМелких граней рядом с видимыми добавлено: " + rescuedFaces : "") +
                    string.Format("\nНевидимые: тел {0}, граней {1} - не создавались\nЧастично видимых граней {4}, обрезано треугольников {5}\nГрубая сетка ({2} тр.): {6:0.0} с, видимость: {3:0.0} с",
                    hiddenBodies, hiddenFaces, coarseTris, (cullMs - coarseMs) / 1000.0, partialFaces, trimmedTris, coarseMs / 1000.0) +
                    (coarseOrderFail > 0 ? "\n(экранная сетка не подошла для " + coarseOrderFail + " тел - грубая тесселяция)" : "");
            if (errors.Count > 0)
                msg += "\n\nОшибки:\n" + string.Join("\n", errors.Take(10).Select(kv => "[" + kv.Value + "] " + kv.Key));
            MessageBox.Show(msg, "Экспорт в GLB");
        }

        #region обход документа
        // экземпляр тела в мировых координатах (Single)
        class Inst
        {
            public PartComponentDefinition def;
            public SurfaceBody body;
            public int index;
            public Asset app;
            public double[] world;
            public string name;
            public string key { get { return docName(def) + "|" + index + "|" + (app == null ? "" : app.DisplayName); } }
        }

        // Single: сначала по грубой сетке выясняем, какие грани видны снаружи,
        // точно тесселируем и упрощаем только их; всё остальное не создаём вовсе
        int single(Document doc, string name)
        {
            merged = new GlbMesh();
            var inst = new List<Inst>();
            if (doc is PartDocument)
            {
                var def = ((PartDocument)doc).ComponentDefinition;
                int i = 0;
                foreach (SurfaceBody b in def.SurfaceBodies)
                {
                    if (b.Visible) inst.Add(new Inst { def = def, body = b, index = i, world = identity(), name = b.Name });
                    i++;
                }
            }
            else walk(((AssemblyDocument)doc).ComponentDefinition.Occurrences, null, inst);

            inSingle = true;
            Dictionary<string, VisInfo> visible = null;
            if (opt.CullHidden)
                try
                {
                    // 1) внешние детали: видимость по грубой сетке тел с закрытыми отверстиями и проёмами
                    var cw = System.Diagnostics.Stopwatch.StartNew();
                    bool[] ext = externalInstances(inst);
                    extMs = cw.ElapsedMilliseconds;
                    var keep = new List<Inst>();
                    for (int i = 0; i < inst.Count; i++) if (ext[i]) keep.Add(inst[i]); else internalParts++;
                    inst = keep;
                    // 2) видимые грани внешних деталей (закрыта только перфорация)
                    cw.Restart();
                    visible = visibleFaces(inst);
                    cullMs = cw.ElapsedMilliseconds;
                }
                catch (Exception e) { error("Видимость", e); visible = null; }

            foreach (var t in inst)
            {
                VisInfo vi = null;
                // vi == null при наличии ключа - тело без грубой сетки, выгружаем целиком
                if (visible != null && (!visible.TryGetValue(t.key, out vi) || (vi != null && vi.faces.Count == 0))) { hiddenBodies++; continue; }
                append(tess(t.def, t.body, t.index, t.app, vi), t.world, merged);
            }
            if (opt.CullHidden && opt.FinalCull)
                try
                {
                    try { I.app.StatusBarText = "GLB: финальная видимость"; } catch { }
                    var fw = System.Diagnostics.Stopwatch.StartNew();
                    finalRemoved = Visibility.cull(merged, opt.CullDirections, opt.CullResolution, 0,
                        backingMat >= 0 ? new HashSet<int> { backingMat } : null);
                    finalMs = fw.ElapsedMilliseconds;
                }
                catch (Exception e) { error("Финальная видимость", e); }
            int mesh = glb.addMesh(name, merged);
            return glb.addNode(name, null, mesh, null);
        }

        // общая сетка экземпляров в мировых координатах (м)
        void worldMesh(List<Inst> inst, Dictionary<string, CoarseMesh> coarse,
            out float[] X, out float[] Y, out float[] Z, out int[] tri, out int[] owner, out int[] ownerTri)
        {
            int nv = 0, nt = 0;
            foreach (var t in inst) { var c = coarse[t.key]; if (c != null) { nv += c.P.Length / 3; nt += c.tri.Length / 3; } }
            X = new float[nv]; Y = new float[nv]; Z = new float[nv];
            tri = new int[nt * 3]; owner = new int[nt]; ownerTri = new int[nt];
            int vo = 0, to = 0;
            for (int ii = 0; ii < inst.Count; ii++)
            {
                var c = coarse[inst[ii].key];
                if (c == null) continue;
                double[] m = inst[ii].world;
                int n = c.P.Length / 3;
                for (int i = 0; i < n; i++)
                {
                    double x = c.P[3 * i], y = c.P[3 * i + 1], z = c.P[3 * i + 2];
                    X[vo + i] = (float)((m[0] * x + m[1] * y + m[2] * z + m[3]) * opt.Scale);
                    Y[vo + i] = (float)((m[4] * x + m[5] * y + m[6] * z + m[7]) * opt.Scale);
                    Z[vo + i] = (float)((m[8] * x + m[9] * y + m[10] * z + m[11]) * opt.Scale);
                }
                bool flip = det3(m) < 0;
                for (int t = 0; t < c.tri.Length / 3; t++, to++)
                {
                    tri[3 * to] = vo + c.tri[3 * t];
                    tri[3 * to + 1] = vo + c.tri[3 * t + (flip ? 2 : 1)];
                    tri[3 * to + 2] = vo + c.tri[3 * t + (flip ? 1 : 2)];
                    owner[to] = ii;
                    ownerTri[to] = t;
                }
                vo += n;
            }
        }

        // деталь (экземпляр тела) внешняя, если снаружи - с закрытыми отверстиями и проёмами - видна заметная доля её поверхности
        bool[] externalInstances(List<Inst> inst)
        {
            var coarse = new Dictionary<string, CoarseMesh>();
            foreach (var t in inst)
            {
                if (coarse.ContainsKey(t.key)) continue;
                try { I.app.StatusBarText = "GLB: внешние детали, сетка " + coarse.Count + "/" + inst.Count; } catch { }
                try { coarse[t.key] = bodyMesh(t.body); } catch (Exception e) { error("Грубая сетка тела", e); coarse[t.key] = null; }
            }
            float[] X, Y, Z; int[] tri, owner, ownerTri;
            worldMesh(inst, coarse, out X, out Y, out Z, out tri, out owner, out ownerTri);
            extTris = tri.Length / 3;
            try { I.app.StatusBarText = "GLB: внешние детали, видимость (" + extTris + " тр.)"; } catch { }
            bool[] vis = Visibility.visible(X, Y, Z, tri, opt.CullDirections, opt.CullResolution, opt.CullOpeningSize / 1000, 0,
                opt.CullHoleSize / 1000, opt.SeeThroughTol / 1000);
            var total = new double[inst.Count];
            var seen = new double[inst.Count];
            for (int t = 0; t < owner.Length; t++)
            {
                int a = tri[3 * t], b = tri[3 * t + 1], c = tri[3 * t + 2];
                double ux = X[b] - X[a], uy = Y[b] - Y[a], uz = Z[b] - Z[a], wx = X[c] - X[a], wy = Y[c] - Y[a], wz = Z[c] - Z[a];
                double cx = uy * wz - uz * wy, cy = uz * wx - ux * wz, cz = ux * wy - uy * wx;
                double ar = Math.Sqrt(cx * cx + cy * cy + cz * cz) / 2;
                total[owner[t]] += ar;
                if (vis[t]) seen[owner[t]] += ar;
            }
            var ext = new bool[inst.Count];
            for (int i = 0; i < inst.Count; i++)
            {
                double share = total[i] > 0 ? seen[i] / total[i] : 0;
                var ci0 = coarse[inst[i].key];
                bool noMesh = ci0 == null || ci0.tri.Length == 0 || total[i] <= 0;
                ext[i] = noMesh || (seen[i] > 0 && share >= opt.ExternalMinShare);   // нечем проверить - оставляем
                int r = rule(inst[i]);
                if (r != 0) ext[i] = r > 0;
                extReport.Add(new ExtRow { name = inst[i].name, share = share, ext = ext[i], byRule = r != 0, tris = ci0 == null ? -1 : ci0.tri.Length / 3 });
            }
            return ext;
        }

        // грубая сетка тела целиком (без разбивки по граням), треугольники наружу
        CoarseMesh bodyMesh(SurfaceBody body)
        {
            int vc = 0, fc = 0; double[] v = new double[] { }, n = new double[] { }; int[] ix = new int[] { };
            try { body.CalculateFacets(coarseTol, out vc, out fc, out v, out n, out ix); } catch { fc = 0; }
            // тело без сетки на уровне тела (бывает у тел из массива) - пробуем по граням
            if (fc == 0 || vc == 0) return coarseMesh(body);
            int off = Array.IndexOf(ix, 0) >= 0 ? 0 : 1;
            var c = new CoarseMesh { P = v, tri = new int[fc * 3], face = new int[fc] };
            for (int k = 0; k < fc; k++)
            {
                int a = ix[3 * k] - off, b = ix[3 * k + 1] - off, cc = ix[3 * k + 2] - off;
                orient(v, n, ref a, ref b, ref cc);
                c.tri[3 * k] = a; c.tri[3 * k + 1] = b; c.tri[3 * k + 2] = cc;
            }
            return c;
        }

        // видимые грани тела; для частично видимых - ячейки (см, в координатах тела), где грань видна
        public class VisInfo
        {
            public HashSet<int> faces = new HashSet<int>();
            public Dictionary<int, HashSet<Key3>> cells = new Dictionary<int, HashSet<Key3>>();
            public double cell;
        }

        double coarseTol { get { return Math.Max(simplifyTol * 3, 0.05); } }

        // грубая сетка каждого тела с номерами граней -> видимость -> видимые грани (и области) для каждого тела
        Dictionary<string, VisInfo> visibleFaces(List<Inst> inst)
        {
            var cwc = System.Diagnostics.Stopwatch.StartNew();
            var coarse = new Dictionary<string, CoarseMesh>();
            foreach (var t in inst)
            {
                string k = t.key;
                if (coarse.ContainsKey(k)) continue;
                try { I.app.StatusBarText = "GLB: грубая сетка " + coarse.Count + "/" + inst.Count; } catch { }
                try { coarse[k] = coarseMesh(t.body); } catch (Exception e) { error("Грубая сетка", e); coarse[k] = null; }
            }
            float[] X, Y, Z; int[] tri, owner, ownerTri;
            worldMesh(inst, coarse, out X, out Y, out Z, out tri, out owner, out ownerTri);
            int nt = owner.Length;
            coarseTris = nt;
            coarseMs = cwc.ElapsedMilliseconds;
            try { I.app.StatusBarText = "GLB: видимость (" + nt + " треугольников грубой сетки)"; } catch { }
            bool[] two = null;
            if (opt.HoleBacking && opt.DropHoleWalls)
            {
                // подложки перфорации - двусторонняя преграда: то, что только за перфорацией, не выгружаем,
                // а сквозь одиночные отверстия, щели и жалюзи видно то, что там есть
                var xl = new List<float>(X); var yl = new List<float>(Y); var zl = new List<float>(Z); var tl = new List<int>(tri);
                foreach (var t in inst)
                {
                    var c = coarse[t.key];
                    if (c == null) continue;
                    if (c.perf == null) c.perf = coarseHoles(c);
                    double[] m = t.world;
                    foreach (var h in c.perf)
                    {
                        int b0 = xl.Count;
                        foreach (Vec q in hexagon(h))
                        {
                            xl.Add((float)((m[0] * q.x + m[1] * q.y + m[2] * q.z + m[3]) * opt.Scale));
                            yl.Add((float)((m[4] * q.x + m[5] * q.y + m[6] * q.z + m[7]) * opt.Scale));
                            zl.Add((float)((m[8] * q.x + m[9] * q.y + m[10] * q.z + m[11]) * opt.Scale));
                        }
                        for (int k = 1; k < BackingSides - 1; k++) { tl.Add(b0); tl.Add(b0 + k); tl.Add(b0 + k + 1); }
                        coarseDisks++;
                    }
                }
                X = xl.ToArray(); Y = yl.ToArray(); Z = zl.ToArray(); tri = tl.ToArray();
                two = new bool[tri.Length / 3];
                for (int t = nt; t < two.Length; t++) two[t] = true;
            }
            // direct - видно снаружи напрямую, не через проёмы (нужно для перфорированных листов)
            bool[] direct = new bool[tri.Length / 3];
            bool[] vis = Visibility.visible(X, Y, Z, tri, opt.CullDirections, opt.CullResolution,
                opt.CullSeeThrough ? opt.CullHoleSize / 1000 : 0, opt.CullKeepDepth / 1000,
                opt.CullOpeningSize / 1000, opt.SeeThroughTol / 1000, two, direct);

            // видимые треугольники грубой сетки по телам (объединение по всем экземплярам)
            var visTri = new Dictionary<string, HashSet<int>>();
            foreach (var t in inst) if (!visTri.ContainsKey(t.key)) visTri[t.key] = new HashSet<int>();
            for (int t = 0; t < nt; t++)
            {
                if (!vis[t]) continue;
                var c = coarse[inst[owner[t]].key];
                // перфорированный лист (за ним подложки) - только если виден напрямую: внутренняя сторона не нужна
                if (opt.PerforatedDirectOnly && !direct[t] && c != null && c.perfFaces != null && c.perfFaces.Contains(c.face[ownerTri[t]])) continue;
                visTri[inst[owner[t]].key].Add(ownerTri[t]);
            }

            var res = new Dictionary<string, VisInfo>();
            double cs = opt.CullCell / 10, margin = Math.Max(3 * coarseTol, 2 * cs);
            foreach (var kv in visTri)
            {
                var c = coarse[kv.Key];
                if (c == null || c.tri.Length == 0) { res[kv.Key] = null; continue; }   // нечем проверить - тесселируем целиком
                var vi = new VisInfo { cell = cs };
                var total = new Dictionary<int, int>();
                var seen = new Dictionary<int, int>();
                for (int t = 0; t < c.face.Length; t++) { int f = c.face[t]; int n; total.TryGetValue(f, out n); total[f] = n + 1; }
                foreach (int t in kv.Value) { int f = c.face[t]; int n; seen.TryGetValue(f, out n); seen[f] = n + 1; vi.faces.Add(f); }
                // частично видимые грани: запоминаем ячейки вокруг видимых треугольников
                foreach (int t in kv.Value)
                {
                    int f = c.face[t];
                    if (!opt.TrimPartialFaces || seen[f] == total[f]) continue;
                    HashSet<Key3> cells;
                    if (!vi.cells.TryGetValue(f, out cells)) vi.cells[f] = cells = new HashSet<Key3>();
                    double x0 = double.MaxValue, y0 = double.MaxValue, z0 = double.MaxValue, x1 = double.MinValue, y1 = double.MinValue, z1 = double.MinValue;
                    for (int k = 0; k < 3; k++)
                    {
                        int v = c.tri[3 * t + k];
                        x0 = Math.Min(x0, c.P[3 * v]); x1 = Math.Max(x1, c.P[3 * v]);
                        y0 = Math.Min(y0, c.P[3 * v + 1]); y1 = Math.Max(y1, c.P[3 * v + 1]);
                        z0 = Math.Min(z0, c.P[3 * v + 2]); z1 = Math.Max(z1, c.P[3 * v + 2]);
                    }
                    for (long ix = (long)Math.Floor((x0 - margin) / cs); ix <= (long)Math.Floor((x1 + margin) / cs); ix++)
                        for (long iy = (long)Math.Floor((y0 - margin) / cs); iy <= (long)Math.Floor((y1 + margin) / cs); iy++)
                            for (long iz = (long)Math.Floor((z0 - margin) / cs); iz <= (long)Math.Floor((z1 + margin) / cs); iz++)
                                cells.Add(new Key3(ix, iy, iz));
                }
                // маленькие скрытые грани рядом с видимыми тоже выгружаем: их треугольники в грубой сетке
                // бывают мельче пикселя, и тогда в поверхности остаётся дыра
                var farea = new Dictionary<int, double>();
                var fkeys = new Dictionary<int, List<Key3>>();
                var visKeys = new HashSet<Key3>();
                for (int t = 0; t < c.face.Length; t++)
                {
                    int f = c.face[t];
                    var q = new Vec[3];
                    for (int k = 0; k < 3; k++) { int v = c.tri[3 * t + k]; q[k] = new Vec(c.P[3 * v], c.P[3 * v + 1], c.P[3 * v + 2]); }
                    double ar;
                    farea.TryGetValue(f, out ar);
                    farea[f] = ar + Vec.cross(q[1] - q[0], q[2] - q[0]).len() / 2;
                    List<Key3> ks;
                    if (!fkeys.TryGetValue(f, out ks)) fkeys[f] = ks = new List<Key3>();
                    foreach (Vec p in q) { var key = new Key3(p); ks.Add(key); if (vi.faces.Contains(f)) visKeys.Add(key); }
                }
                for (int pass = 0; pass < 2; pass++)
                    foreach (var fk in fkeys)
                        if (!vi.faces.Contains(fk.Key) && farea[fk.Key] <= opt.SmallFaceArea / 100 && fk.Value.Any(visKeys.Contains))
                        {
                            vi.faces.Add(fk.Key);
                            rescuedFaces++;
                            foreach (var key in fk.Value) visKeys.Add(key);
                        }
                // список невыгруженных граней - в отчёт
                var owner0 = inst.First(t => t.key == kv.Key);
                foreach (var fk in fkeys)
                {
                    if (vi.faces.Contains(fk.Key)) continue;
                    var cen = new Vec();
                    foreach (var key in fk.Value) cen = cen + new Vec(key.x * 1e-5, key.y * 1e-5, key.z * 1e-5);
                    cen = cen * (1.0 / fk.Value.Count);
                    double[] m = owner0.world;
                    hiddenList.Add(string.Format(CultureInfo.InvariantCulture, "  {0,9:0.0} мм²  центр ({1:0}, {2:0}, {3:0}) мм  {4} / грань {5}",
                        farea[fk.Key] * 100,
                        (m[0] * cen.x + m[1] * cen.y + m[2] * cen.z + m[3]) * 10,
                        (m[4] * cen.x + m[5] * cen.y + m[6] * cen.z + m[7]) * 10,
                        (m[8] * cen.x + m[9] * cen.y + m[10] * cen.z + m[11]) * 10, owner0.name, fk.Key));
                }
                partialFaces += vi.cells.Count;
                hiddenFaces += c.faceCount - vi.faces.Count;
                res[kv.Key] = vi;
            }
            return res;
        }

        class CoarseMesh { public double[] P, N; public int[] tri, face; public int faceCount; public List<Hole> perf; public HashSet<int> perfFaces; }

        // готовая экранная сетка Inventor с разбивкой по граням; если её нет - грубая тесселяция каждой грани
        CoarseMesh coarseMesh(SurfaceBody body)
        {
            var c = new CoarseMesh();
            var faces = new List<Face>();
            foreach (Face f in body.Faces) faces.Add(f);
            c.faceCount = faces.Count;
            try
            {
                // просим Inventor посчитать грубую сетку тела, затем забираем её с разбивкой по граням
                double tol = coarseTol;
                try
                {
                    int vc0, fc0; double[] v0 = new double[] { }, n0 = new double[] { }; int[] i0 = new int[] { };
                    body.CalculateFacets(tol, out vc0, out fc0, out v0, out n0, out i0);
                }
                catch { }
                int tc; double[] tols = new double[] { };
                body.GetExistingFacetTolerances(out tc, out tols);
                if (tc > 0)
                {
                    double best = tols.OrderBy(x => Math.Abs(Math.Log(Math.Max(x, 1e-9) / tol))).First();
                    int vc, fc; double[] v = new double[] { }, n = new double[] { }, tex = new double[] { }; int[] ix = new int[] { }, per = new int[] { };
                    body.GetExistingFacetsAndTextureMap2(best, out vc, out fc, out v, out n, out ix, out tex, out per);
                    int sum = per.Sum();
                    if (sum == fc) for (int i = 0; i < per.Length; i++) per[i] *= 3;   // если вдруг число треугольников, а не индексов
                    if (fc > 0 && per.Length == faces.Count && per.Sum() == ix.Length)
                    {
                        int off = Array.IndexOf(ix, 0) >= 0 ? 0 : 1;
                        c.P = v; c.N = n;
                        c.tri = new int[ix.Length];
                        c.face = new int[fc];
                        int t = 0, fi = 0, left = per.Length > 0 ? per[0] : 0;
                        for (int k = 0; k < fc; k++)
                        {
                            while (left <= 0 && fi < per.Length - 1) left = per[++fi];
                            int a = ix[3 * k] - off, b = ix[3 * k + 1] - off, cc = ix[3 * k + 2] - off;
                            orient(v, n, ref a, ref b, ref cc);
                            c.tri[t++] = a; c.tri[t++] = b; c.tri[t++] = cc;
                            c.face[k] = fi;
                            left -= 3;
                        }
                        if (orderOk(c, faces)) return c;
                        coarseOrderFail++;
                    }
                }
            }
            catch { }
            // запасной путь: грубая тесселяция по граням
            var P = new List<double>(); var T = new List<int>(); var F = new List<int>(); var NN = new List<double>();
            double ftol = coarseTol;
            for (int fi = 0; fi < faces.Count; fi++)
            {
                int vc = 0, fc = 0; double[] v = new double[] { }, n = new double[] { }; int[] ix = new int[] { };
                try { faces[fi].CalculateFacets(ftol, out vc, out fc, out v, out n, out ix); } catch { continue; }
                if (vc == 0 || fc == 0) continue;
                int off = Array.IndexOf(ix, 0) >= 0 ? 0 : 1, b0 = P.Count / 3;
                P.AddRange(v); NN.AddRange(n);
                for (int k = 0; k < fc; k++)
                {
                    int a = ix[3 * k] - off, b = ix[3 * k + 1] - off, cc = ix[3 * k + 2] - off;
                    orient(v, n, ref a, ref b, ref cc);
                    T.Add(b0 + a); T.Add(b0 + b); T.Add(b0 + cc); F.Add(fi);
                }
            }
            c.P = P.ToArray(); c.N = NN.ToArray(); c.tri = T.ToArray(); c.face = F.ToArray();
            return c;
        }

        int coarseOrderFail;

        // проверка, что разбивка по граням идёт в порядке body.Faces: вершины треугольников грани лежат в её габарите
        static bool orderOk(CoarseMesh c, List<Face> faces)
        {
            var check = new HashSet<int> { 0, faces.Count / 3, faces.Count / 2, 2 * faces.Count / 3, faces.Count - 1 };
            foreach (int fi in check)
            {
                int t = Array.IndexOf(c.face, fi);
                if (t < 0) continue;
                Box b;
                try { b = faces[fi].Evaluator.RangeBox; } catch { continue; }
                double e = 1e-3 + 0.01 * b.MinPoint.DistanceTo(b.MaxPoint);
                for (int k = 0; k < 3; k++)
                {
                    int v = c.tri[3 * t + k];
                    double x = c.P[3 * v], y = c.P[3 * v + 1], z = c.P[3 * v + 2];
                    if (x < b.MinPoint.X - e || x > b.MaxPoint.X + e || y < b.MinPoint.Y - e || y > b.MaxPoint.Y + e ||
                        z < b.MinPoint.Z - e || z > b.MaxPoint.Z + e) return false;
                }
            }
            return true;
        }

        // треугольник по нормалям поверхности (наружу)
        static void orient(double[] v, double[] n, ref int a, ref int b, ref int c)
        {
            double ux = v[3 * b] - v[3 * a], uy = v[3 * b + 1] - v[3 * a + 1], uz = v[3 * b + 2] - v[3 * a + 2];
            double wx = v[3 * c] - v[3 * a], wy = v[3 * c + 1] - v[3 * a + 1], wz = v[3 * c + 2] - v[3 * a + 2];
            double cx = uy * wz - uz * wy, cy = uz * wx - ux * wz, cz = ux * wy - uy * wx;
            double nx = n[3 * a] + n[3 * b] + n[3 * c], ny = n[3 * a + 1] + n[3 * b + 1] + n[3 * c + 1], nz = n[3 * a + 2] + n[3 * b + 2] + n[3 * c + 2];
            if (cx * nx + cy * ny + cz * nz < 0) { int x = b; b = c; c = x; }
        }

        int assembly(Document doc, string name)
        {
            var children = new List<int>();
            if (doc is PartDocument)
                children = bodies(((PartDocument)doc).ComponentDefinition, null, name);
            else
                foreach (ComponentOccurrence o in ((AssemblyDocument)doc).ComponentDefinition.Occurrences)
                {
                    int n = occNode(o, identity(), null);
                    if (n >= 0) children.Add(n);
                }
            return glb.addNode(name, null, -1, children);
        }

        // Single: список всех видимых тел с мировыми матрицами
        void walk(IEnumerable occs, Asset parentApp, List<Inst> inst)
        {
            foreach (ComponentOccurrence o in occs)
            {
                if (skip(o)) continue;
                Asset app = parentApp ?? overrideApp(o);
                if (o.DefinitionDocumentType == DocumentTypeEnum.kAssemblyDocumentObject)
                    walk(o.SubOccurrences, app, inst);
                else if (o.Definition is PartComponentDefinition)
                {
                    var def = (PartComponentDefinition)o.Definition;
                    double[] world = mat(o.Transformation);
                    int i = 0;
                    foreach (SurfaceBody b in def.SurfaceBodies)
                    {
                        if (b.Visible) inst.Add(new Inst { def = def, body = b, index = i, app = app, world = world, name = o.Name + " / " + b.Name });
                        i++;
                    }
                }
            }
        }

        // Assembly: узел на вхождение, локальная матрица относительно родителя
        int occNode(ComponentOccurrence o, double[] parentWorld, Asset parentApp)
        {
            if (skip(o)) return -1;
            Asset app = parentApp ?? overrideApp(o);
            double[] world = mat(o.Transformation);
            double[] local = mul(invRigid(parentWorld), world);
            var children = new List<int>();
            int mesh = -1;
            if (o.DefinitionDocumentType == DocumentTypeEnum.kAssemblyDocumentObject)
            {
                foreach (ComponentOccurrence s in o.SubOccurrences)
                {
                    int n = occNode(s, world, app);
                    if (n >= 0) children.Add(n);
                }
                if (children.Count == 0) return -1;
            }
            else if (o.Definition is PartComponentDefinition)
            {
                var def = (PartComponentDefinition)o.Definition;
                children = bodies(def, app, o.Name);
                // одно тело - меш прямо на узле вхождения
                if (children.Count == 1 && glb.nodeHasOnlyMesh(children[0]))
                {
                    mesh = glb.nodeMesh(children[0]);
                    glb.removeLastNode();
                    children.Clear();
                }
                if (children.Count == 0 && mesh < 0) return -1;
            }
            else return -1;
            return glb.addNode(o.Name, local, mesh, children);
        }

        List<int> bodies(PartComponentDefinition def, Asset app, string name)
        {
            var res = new List<int>();
            int i = 0;
            foreach (SurfaceBody b in def.SurfaceBodies)
            {
                if (b.Visible)
                {
                    string key = docName(def) + "|" + i + "|" + (app == null ? "" : app.DisplayName);
                    int mesh;
                    if (meshCache.ContainsKey(key)) used(key, def, b);
                    if (!meshCache.TryGetValue(key, out mesh))
                    {
                        mesh = glb.addMesh(b.Name, tess(def, b, i, app));
                        meshCache[key] = mesh;
                    }
                    if (mesh >= 0) res.Add(glb.addNode(b.Name, null, mesh, null));
                }
                i++;
            }
            return res;
        }

        static bool skip(ComponentOccurrence o)
        {
            try { return o.Suppressed || !o.Visible; } catch { return true; }
        }

        static Asset overrideApp(ComponentOccurrence o)
        {
            try { if (o.AppearanceSourceType == AppearanceSourceTypeEnum.kOverrideAppearance) return o.Appearance; } catch { }
            return null;
        }

        static string docName(PartComponentDefinition def)
        {
            try { return ((Document)def.Document).FullDocumentName; } catch { return def.GetHashCode().ToString(); }
        }
        #endregion

        #region тесселяция
        // vis - видимые грани (номера в порядке body.Faces) и области частично видимых; null - все грани целиком
        GlbMesh tess(PartComponentDefinition def, SurfaceBody body, int bodyIndex, Asset app, VisInfo vis = null)
        {
            string key = docName(def) + "|" + bodyIndex + "|" + (app == null ? "" : app.DisplayName);
            used(key, def, body);
            GlbMesh m;
            if (tessCache.TryGetValue(key, out m)) return m;
            m = new GlbMesh();
            cur = stats[key];
            try { I.app.StatusBarText = "GLB (" + tessCache.Count + "): " + cur.name; } catch { }
            var sw = System.Diagnostics.Stopwatch.StartNew();
            var bm = new BodyMesher();
            Asset bodyApp = null;
            try { bodyApp = body.Appearance; } catch { }
            int fi = -1;
            foreach (Face f in body.Faces)
            {
                fi++;
                if (vis != null && !vis.faces.Contains(fi)) continue;
                HashSet<Key3> cells = null;
                if (vis != null) vis.cells.TryGetValue(fi, out cells);
                Asset a = app;
                if (a == null) try { a = f.Appearance; } catch { }
                if (a == null) a = bodyApp;
                try { addFace(f, bm, material(a), cells, vis == null ? 0 : vis.cell); } catch (Exception e) { error("Грань", e); }
            }
            addBackings(bm);
            cur.msFacets = sw.ElapsedMilliseconds;
            sw.Restart();
            cur.trisBefore = bm.triCount;
            if (opt.DebugDumpTris > 0 && bm.triCount > opt.DebugDumpTris)
                try
                {
                    string dn = new string(cur.name.Select(ch => System.IO.Path.GetInvalidFileNameChars().Contains(ch) ? '_' : ch).ToArray());
                    bm.dump(outDir + "\\_debug_" + dn + ".txt");
                }
                catch (Exception e) { error("Отладочный файл", e); }
            bm.TimeLimitMs = opt.SimplifyTimeLimitMs;
            bm.MaxTurnDeg = 360.0 / Math.Max(6, opt.CircleSegments);
            try { bm.simplify(simplifyTol); } catch (Exception e) { error("Упрощение", e); }
            if (bm.timedOut) inc(planarWhy, "упрощение остановлено по времени: " + cur.name);
            foreach (var kv in bm.failures) inc(planarWhy, "после упрощения: " + kv.Key, kv.Value);
            bm.emit(m, opt.Scale);
            cur.msSimplify = sw.ElapsedMilliseconds;
            tessCache[key] = m;
            foreach (var p in m.prims.Values) { stats[key].tris += p.idx.Count / 3; stats[key].verts += p.vcount; }
            return m;
        }

        // cells - ячейки (размер cs, см), где грань видна; треугольники вне них не выгружаются
        void addFace(Face f, BodyMesher bm, int mat, HashSet<Key3> cells = null, double cs = 0)
        {
            int vc = 0, fc = 0;
            double[] v = new double[] { }, n = new double[] { };
            int[] ix = new int[] { };
            f.CalculateFacets(opt.Tolerance, out vc, out fc, out v, out n, out ix);
            if (vc == 0 || fc == 0) { error("Грань", new Exception("CalculateFacets вернул пустую сетку")); return; }
            int off = Array.IndexOf(ix, 0) >= 0 ? 0 : 1;   // индексы Inventor 1-based

            // сварка совпадающих вершин грани
            var map = new int[vc];
            var P = new List<Vec>();
            var N = new List<Vec>();
            var keys = new Dictionary<Key3, int>();
            for (int i = 0; i < vc; i++)
            {
                var pt = new Vec(v[3 * i], v[3 * i + 1], v[3 * i + 2]);
                var nr = new Vec(n[3 * i], n[3 * i + 1], n[3 * i + 2]);
                var k = new Key3(pt);
                int id;
                if (keys.TryGetValue(k, out id)) N[id] = N[id] + nr;
                else { id = P.Count; keys[k] = id; P.Add(pt); N.Add(nr); }
                map[i] = id;
            }

            // треугольники с ориентацией по нормалям поверхности
            var tris = new List<int>();
            Vec fn = new Vec();
            for (int t = 0; t < fc; t++)
            {
                int a = map[ix[3 * t] - off], b = map[ix[3 * t + 1] - off], c = map[ix[3 * t + 2] - off];
                if (a == b || b == c || a == c) continue;
                Vec cr = Vec.cross(P[b] - P[a], P[c] - P[a]);
                if (Vec.dot(cr, N[a] + N[b] + N[c]) < 0) { int x = b; b = c; c = x; cr = cr * -1; }
                fn = fn + cr;
                tris.Add(a); tris.Add(b); tris.Add(c);
            }
            if (tris.Count == 0) return;
            // тип поверхности не проверяем: у гнутой перфорации стенки отверстий - сплайны; форму определяет isHoleWall по сетке
            if (inSingle && opt.DropHoleWalls && f.SurfaceType != SurfaceTypeEnum.kPlaneSurface && isHoleWall(f, P, N))
            {
                droppedWalls++;
                if (opt.HoleBacking) holes.Add(holeOf(P, N));
                return;
            }

            bool flat = false;
            if (opt.PlanarRetriangulate && f.SurfaceType == SurfaceTypeEnum.kPlaneSurface && fn.len() > 0)
            {
                string why;
                var rt = PlanarTriangulator.run(P, tris, fn.norm(), opt.PlanarMaxVertices, out why);
                if (rt != null) { tris = rt; flat = true; cur.planarOk++; }
                else { cur.planarFail++; inc(planarWhy, why); }
            }
            inc(cur.byType, f.SurfaceType.ToString().Replace("Surface", "").Substring(1), tris.Count / 3);

            Vec fnu = fn.norm();
            bool trimmed = false;
            if (cells != null && cs > 0)
            {
                var kept = new List<int>(tris.Count);
                for (int t = 0; t < tris.Count; t += 3)
                {
                    bool keep = false;
                    for (int k = 0; k < 3 && !keep; k++)
                    {
                        Vec q = P[tris[t + k]];
                        keep = cells.Contains(new Key3((long)Math.Floor(q.x / cs), (long)Math.Floor(q.y / cs), (long)Math.Floor(q.z / cs)));
                    }
                    if (keep) { kept.Add(tris[t]); kept.Add(tris[t + 1]); kept.Add(tris[t + 2]); }
                }
                trimmed = kept.Count < tris.Count;
                trimmedTris += (tris.Count - kept.Count) / 3;
                tris = kept;
                if (tris.Count == 0) return;
            }
            // у обрезанной плоской грани контур уже не замкнутый - упрощаем её как обычную сетку
            int fid = bm.addFace(mat, flat && !trimmed ? fnu : new Vec());
            var g = new int[P.Count];
            for (int i = 0; i < P.Count; i++) g[i] = -1;
            foreach (int id in tris)
                if (g[id] < 0)
                {
                    g[id] = bm.vertex(P[id]);
                    bm.normal(g[id], fid, flat ? fnu : N[id].norm());
                }
            for (int t = 0; t < tris.Count; t += 3) bm.tri(fid, g[tris[t]], g[tris[t + 1]], g[tris[t + 2]]);
        }

        class Hole { public Vec c, ax; public double r; }
        List<Hole> holes = new List<Hole>();   // отверстия текущего тела (стенки не выгружены)

        // центр на середине толщины, ось и радиус отверстия по сетке его стенки
        Hole holeOf(List<Vec> P, List<Vec> N)
        {
            var ns = N.Select(n => n.norm()).ToList();
            Vec n0 = ns[0], n1 = ns.OrderBy(n => Math.Abs(Vec.dot(n, n0))).First();
            Vec ax = Vec.cross(n0, n1).norm();
            var c = new Vec();
            foreach (Vec p in P) c = c + p;
            c = c * (1.0 / P.Count);
            double lo = double.MaxValue, hi = double.MinValue, r = 0;
            foreach (Vec p in P) { double t = Vec.dot(p - c, ax); lo = Math.Min(lo, t); hi = Math.Max(hi, t); }
            c = c + ax * ((lo + hi) / 2);
            foreach (Vec p in P) { Vec d = p - c; r = Math.Max(r, (d - ax * Vec.dot(d, ax)).len()); }
            return new Hole { c = c, ax = ax, r = r };
        }

        // отверстия перфорации: у отверстия >= PerforationNeighbors соседей ближе PerforationRadius диаметров,
        // и в связной группе соседствующих отверстий их >= PerforationMinHoles
        List<Hole> perforation(List<Hole> hs)
        {
            int n = hs.Count;
            var par = new int[n];
            for (int i = 0; i < n; i++) par[i] = i;
            Func<int, int> find = null;
            find = x => { while (par[x] != x) { par[x] = par[par[x]]; x = par[x]; } return x; };
            var deg = new int[n];
            for (int i = 0; i < n; i++)
                for (int j = i + 1; j < n; j++)
                    // соседи - только отверстия того же размера (перфорация - много одинаковых отверстий вплотную)
                    if (Math.Abs(hs[i].r - hs[j].r) <= 0.15 * Math.Max(hs[i].r, hs[j].r) &&
                        (hs[j].c - hs[i].c).len() <= opt.PerforationRadius * 2 * Math.Min(hs[i].r, hs[j].r))
                    {
                        deg[i]++; deg[j]++;
                        par[find(i)] = find(j);
                    }
            var size = new Dictionary<int, int>();
            for (int i = 0; i < n; i++) { int r = find(i), c; size.TryGetValue(r, out c); size[r] = c + 1; }
            var res = new List<Hole>();
            for (int i = 0; i < n; i++)
                if (deg[i] >= opt.PerforationNeighbors && size[find(i)] >= opt.PerforationMinHoles) res.Add(hs[i]);
            return res;
        }

        void addBackings(BodyMesher bm)
        {
            foreach (var h in perforation(holes)) addBacking(bm, h);
            holes.Clear();
        }

        // восьмиугольник подложки, описанный вокруг отверстия, на середине толщины (углы едва выходят за край отверстия)
        const int BackingSides = 8;
        static Vec[] hexagon(Hole h)
        {
            double R = h.r * 1.02 / Math.Cos(Math.PI / BackingSides);
            Vec u = Math.Abs(h.ax.x) < 0.9 ? Vec.cross(h.ax, new Vec(1, 0, 0)).norm() : Vec.cross(h.ax, new Vec(0, 1, 0)).norm();
            Vec w = Vec.cross(h.ax, u);
            var r = new Vec[BackingSides];
            for (int k = 0; k < BackingSides; k++) { double a = k * 2 * Math.PI / BackingSides; r[k] = h.c + u * (R * Math.Cos(a)) + w * (R * Math.Sin(a)); }
            return r;
        }

        // отверстия перфорации по грубой сетке тела (стенки - по граням)
        List<Hole> coarseHoles(CoarseMesh c)
        {
            if (c.N == null) return new List<Hole>();
            var byFace = new Dictionary<int, HashSet<int>>();
            for (int t = 0; t < c.face.Length; t++)
            {
                HashSet<int> vs;
                if (!byFace.TryGetValue(c.face[t], out vs)) byFace[c.face[t]] = vs = new HashSet<int>();
                for (int k = 0; k < 3; k++) vs.Add(c.tri[3 * t + k]);
            }
            var hs = new List<Hole>();
            var wallOf = new Dictionary<Hole, int>();
            foreach (var kv in byFace)
            {
                var vs = kv.Value;
                var P = vs.Select(i => new Vec(c.P[3 * i], c.P[3 * i + 1], c.P[3 * i + 2])).ToList();
                var N = vs.Select(i => new Vec(c.N[3 * i], c.N[3 * i + 1], c.N[3 * i + 2])).ToList();
                if (P.Count >= 6 && isHoleWall(P, N)) { var h = holeOf(P, N); hs.Add(h); wallOf[h] = kv.Key; }
            }
            var perf = perforation(hs);
            // перфорированные грани листа: касаются стенок не меньше PerforationMinHoles отверстий перфорации
            var wallKeys = new HashSet<Key3>();
            foreach (var h in perf)
                foreach (int v in byFace[wallOf[h]]) wallKeys.Add(new Key3(new Vec(c.P[3 * v], c.P[3 * v + 1], c.P[3 * v + 2])));
            var perfWalls = new HashSet<int>(perf.Select(h => wallOf[h]));
            c.perfFaces = new HashSet<int>();
            foreach (var kv in byFace)
            {
                if (perfWalls.Contains(kv.Key)) continue;
                int touch = kv.Value.Count(v => wallKeys.Contains(new Key3(new Vec(c.P[3 * v], c.P[3 * v + 1], c.P[3 * v + 2]))));
                if (touch >= opt.PerforationMinHoles * 3) c.perfFaces.Add(kv.Key);
            }
            return perf;
        }

        // подложка: шестиугольник на середине толщины стенки, описанный вокруг отверстия - его края уходят в металл листа,
        // сквозь отверстие виден ровный тёмный фон; материал двусторонний
        void addBacking(BodyMesher bm, Hole h)
        {
            Vec[] hx = hexagon(h);
            int fid = bm.addFace(backingMaterial(), new Vec());
            var ids = new int[hx.Length];
            for (int k = 0; k < hx.Length; k++)
            {
                ids[k] = bm.vertex(hx[k]);
                bm.normal(ids[k], fid, h.ax);
            }
            for (int k = 1; k < hx.Length - 1; k++) bm.tri(fid, ids[0], ids[k], ids[k + 1]);
            backingDisks++;
        }

        int backingMat = -1, backingDisks;
        long finalRemoved = -1, finalMs;
        int coarseDisks, rescuedFaces;
        List<string> hiddenList = new List<string>();
        int backingMaterial()
        {
            if (backingMat < 0)
                backingMat = glb.addMaterial(new GlbMat
                {
                    name = "Подложка перфорации", r = opt.BackingColor[0], g = opt.BackingColor[1], b = opt.BackingColor[2],
                    metallic = 0, roughness = 0.9, doubleSided = true
                });
            return backingMat;
        }

        // стенка мелкого отверстия: вогнутый цилиндр малого диаметра и глубины.
        // Ось и радиус - по сетке (нормали цилиндра перпендикулярны оси): f.Geometry у граней гнутого листа не всегда Cylinder
        bool isHoleWall(Face f, List<Vec> P, List<Vec> N) { return isHoleWall(P, N); }

        bool isHoleWall(List<Vec> P, List<Vec> N)
        {
            if (P.Count < 6) return false;
            var ns = N.Select(n => n.norm()).ToList();
            Vec n0 = ns[0], n1 = ns.OrderBy(n => Math.Abs(Vec.dot(n, n0))).First();
            Vec ax = Vec.cross(n0, n1);
            if (ax.len() < 0.5) return false;   // нормали почти параллельны - не цилиндр
            ax = ax.norm();
            if (ns.Any(n => Math.Abs(Vec.dot(n, ax)) > 0.05)) return false;
            var c = new Vec();
            foreach (Vec p in P) c = c + p;
            c = c * (1.0 / P.Count);
            double lo = double.MaxValue, hi = double.MinValue, concave = 0, r = 0;
            for (int i = 0; i < P.Count; i++)
            {
                Vec d = P[i] - c;
                double t = Vec.dot(d, ax);
                lo = Math.Min(lo, t); hi = Math.Max(hi, t);
                Vec rad = d - ax * t;
                concave += Vec.dot(ns[i], rad);
                r += rad.len();
            }
            r /= P.Count;
            return concave < 0 && 2 * r <= opt.CullHoleSize / 10 && hi - lo <= opt.HoleWallMaxDepth / 10;
        }

        // добавление тела в общий меш с преобразованием (матрица в см)
        void append(GlbMesh src, double[] m, GlbMesh dst)
        {
            bool flip = det3(m) < 0;
            foreach (var kv in src.prims)
            {
                GlbPrim s = kv.Value, d = dst.get(kv.Key);
                int b = d.vcount;
                for (int i = 0; i < s.vcount; i++)
                {
                    double x = s.pos[3 * i], y = s.pos[3 * i + 1], z = s.pos[3 * i + 2];
                    d.pos.Add((float)(m[0] * x + m[1] * y + m[2] * z + m[3] * opt.Scale));
                    d.pos.Add((float)(m[4] * x + m[5] * y + m[6] * z + m[7] * opt.Scale));
                    d.pos.Add((float)(m[8] * x + m[9] * y + m[10] * z + m[11] * opt.Scale));
                    x = s.nrm[3 * i]; y = s.nrm[3 * i + 1]; z = s.nrm[3 * i + 2];
                    d.nrm.Add((float)(m[0] * x + m[1] * y + m[2] * z));
                    d.nrm.Add((float)(m[4] * x + m[5] * y + m[6] * z));
                    d.nrm.Add((float)(m[8] * x + m[9] * y + m[10] * z));
                }
                for (int i = 0; i < s.idx.Count; i += 3)
                {
                    d.idx.Add(b + s.idx[i]);
                    d.idx.Add(b + s.idx[i + (flip ? 2 : 1)]);
                    d.idx.Add(b + s.idx[i + (flip ? 1 : 2)]);
                }
            }
        }
        #endregion

        #region материалы
        int material(Asset a)
        {
            string key = a == null ? "" : a.DisplayName;
            int id;
            if (matIdx.TryGetValue(key, out id)) return id;
            id = glb.addMaterial(GlbMat.from(a));
            matIdx[key] = id;
            matAssets[key] = a;
            return id;
        }
        #endregion

        void report(string fn)
        {
            var sb = new StringBuilder();
            sb.AppendLine(string.Format("Время: тесселяция Inventor {0:0.0} с, упрощение {1:0.0} с (всего тел {2})",
                stats.Values.Sum(x => x.msFacets) / 1000.0, stats.Values.Sum(x => x.msSimplify) / 1000.0, stats.Count));
            sb.AppendLine("Самые долгие: " + string.Join("; ", stats.Values.OrderByDescending(x => x.msFacets + x.msSimplify).Take(5)
                .Select(x => string.Format("{0} ({1:0.0} + {2:0.0} с)", x.name, x.msFacets / 1000.0, x.msSimplify / 1000.0))));
            sb.AppendLine();
            if (extReport.Count > 0)
            {
                sb.AppendLine(string.Format("== Детали: доля поверхности, видимая снаружи (внешняя, если >= {0:0.#}%) ==", opt.ExternalMinShare * 100));
                if (partRules.Count > 0) sb.AppendLine("  правил из GlbParts.txt: " + partRules.Count);
                foreach (var row in extReport.OrderByDescending(x => x.share))
                    sb.AppendLine(string.Format("  {0,6:0.0}%  {1}{2}  {3,7} тр.  {4}", row.share * 100, row.ext ? "внешняя  " : "внутренняя",
                        row.byRule ? " (правило)" : row.tris <= 0 ? " (нет сетки)" : "           ", row.tris, row.name));
                sb.AppendLine();
            }
            if (hiddenList.Count > 0)
            {
                sb.AppendLine("== Невыгруженные (невидимые) грани внешних деталей ==");
                foreach (var l in hiddenList) sb.AppendLine(l);
                sb.AppendLine();
            }
            sb.AppendLine("== Тела по треугольникам (всего = на экземпляр x экземпляров) ==");
            foreach (var st in stats.Values.OrderByDescending(x => (long)x.tris * x.count).Take(60))
            {
                sb.AppendLine(string.Format("{0,9} = {1,7} x {2,-4} верш. {3,7}  (до упрощения {5})  {4}", (long)st.tris * st.count, st.tris, st.count, st.verts, st.name, st.trisBefore));
                sb.AppendLine("            " + string.Join(", ", st.byType.OrderByDescending(x => x.Value).Select(x => x.Key + " " + x.Value)) +
                    string.Format("  | плоских перетриангулировано {0} из {1}", st.planarOk, st.planarOk + st.planarFail));
            }
            if (planarWhy.Count > 0)
            {
                sb.AppendLine();
                sb.AppendLine("== Замечания (плоские грани, оставленные как у Inventor, и прочее) ==");
                foreach (var kv in planarWhy) sb.AppendLine("  " + kv.Value + "  " + kv.Key);
            }
            sb.AppendLine();
            sb.AppendLine("== Материалы ==");
            foreach (var kv in matAssets)
            {
                Asset a = kv.Value;
                sb.AppendLine("[" + kv.Key + "]");
                if (a == null) continue;
                try { sb.AppendLine("  Name = " + a.Name); } catch { }
                try
                {
                    foreach (AssetValue v in a) sb.AppendLine("  " + v.Name + " = " + GlbMat.valueText(v));
                }
                catch (Exception e) { sb.AppendLine("  ! " + e.Message); }
            }
            System.IO.File.WriteAllText(fn, sb.ToString(), Encoding.UTF8);
        }

        #region матрицы (row-major 4x4, перенос в см)
        static double[] identity() { return new double[] { 1, 0, 0, 0, 0, 1, 0, 0, 0, 0, 1, 0, 0, 0, 0, 1 }; }

        static double[] mat(Matrix m)
        {
            var r = new double[16];
            for (int i = 0; i < 4; i++)
                for (int j = 0; j < 4; j++)
                    r[i * 4 + j] = m.get_Cell(i + 1, j + 1);
            return r;
        }

        static double[] mul(double[] a, double[] b)
        {
            var r = new double[16];
            for (int i = 0; i < 4; i++)
                for (int j = 0; j < 4; j++)
                {
                    double s = 0;
                    for (int k = 0; k < 4; k++) s += a[i * 4 + k] * b[k * 4 + j];
                    r[i * 4 + j] = s;
                }
            return r;
        }

        static double[] invRigid(double[] m)
        {
            var r = identity();
            for (int i = 0; i < 3; i++)
                for (int j = 0; j < 3; j++) r[i * 4 + j] = m[j * 4 + i];
            for (int i = 0; i < 3; i++)
                r[i * 4 + 3] = -(r[i * 4] * m[3] + r[i * 4 + 1] * m[7] + r[i * 4 + 2] * m[11]);
            return r;
        }

        static double det3(double[] m)
        {
            return m[0] * (m[5] * m[10] - m[6] * m[9]) - m[1] * (m[4] * m[10] - m[6] * m[8]) + m[2] * (m[4] * m[9] - m[5] * m[8]);
        }
        #endregion
    }

    public struct Vec
    {
        public double x, y, z;
        public Vec(double x, double y, double z) { this.x = x; this.y = y; this.z = z; }
        public static Vec operator +(Vec a, Vec b) { return new Vec(a.x + b.x, a.y + b.y, a.z + b.z); }
        public static Vec operator -(Vec a, Vec b) { return new Vec(a.x - b.x, a.y - b.y, a.z - b.z); }
        public static Vec operator *(Vec a, double s) { return new Vec(a.x * s, a.y * s, a.z * s); }
        public static double dot(Vec a, Vec b) { return a.x * b.x + a.y * b.y + a.z * b.z; }
        public static Vec cross(Vec a, Vec b) { return new Vec(a.y * b.z - a.z * b.y, a.z * b.x - a.x * b.z, a.x * b.y - a.y * b.x); }
        public double len() { return Math.Sqrt(x * x + y * y + z * z); }
        public Vec norm() { double l = len(); return l > 0 ? this * (1 / l) : new Vec(0, 0, 1); }
    }

    // хэш для long-ключей (пара индексов): стандартный long.GetHashCode = lo ^ hi,
    // для пар близких индексов это почти всегда маленькие числа и словари вырождаются в цепочки коллизий
    public sealed class LongCmp : IEqualityComparer<long>
    {
        public static readonly LongCmp I = new LongCmp();
        public bool Equals(long a, long b) { return a == b; }
        public int GetHashCode(long k)
        {
            unchecked
            {
                ulong x = (ulong)k * 0x9E3779B97F4A7C15UL;
                x ^= x >> 29;
                return (int)x ^ (int)(x >> 32);
            }
        }
    }

    // числовой ключ словаря из трёх целых (сварка вершин по координатам, округлённым до 0.1 мкм)
    public struct Key3 : IEquatable<Key3>
    {
        public long x, y, z;
        public Key3(long x, long y, long z) { this.x = x; this.y = y; this.z = z; }
        public Key3(Vec p) : this((long)Math.Round(p.x * 1e5), (long)Math.Round(p.y * 1e5), (long)Math.Round(p.z * 1e5)) { }
        public bool Equals(Key3 o) { return x == o.x && y == o.y && z == o.z; }
        public override bool Equals(object o) { return o is Key3 && Equals((Key3)o); }
        public override int GetHashCode() { unchecked { return (int)(x * 73856093L ^ y * 19349663L ^ z * 83492791L); } }
    }

    // Перетриангуляция плоской грани: берём граничный контур сетки Inventor (он совпадает с соседними
    // гранями, поэтому щелей не будет), выкидываем внутренние точки и триангулируем earcut'ом
    static class PlanarTriangulator
    {
        public static List<int> run(List<Vec> P, List<int> tris, Vec n, int maxVerts, out string why)
        {
            why = null;
            // граничные рёбра: направленные рёбра без обратной пары
            var edges = new HashSet<long>(LongCmp.I);
            for (int t = 0; t < tris.Count; t += 3)
                for (int k = 0; k < 3; k++)
                    if (!edges.Add(key(tris[t + k], tris[t + (k + 1) % 3]))) { why = "дублирующиеся рёбра"; return null; }
            var next = new Dictionary<int, int>();
            foreach (long e in edges)
            {
                int a = (int)(e >> 32), b = (int)(e & 0xffffffff);
                if (edges.Contains(key(b, a))) continue;
                if (next.ContainsKey(a)) { why = "контуры касаются в вершине"; return null; }
                next[a] = b;
            }
            if (next.Count < 3) { why = "нет контура"; return null; }
            if (next.Count > maxVerts) { why = "слишком много вершин"; return null; }

            var loops = new List<List<int>>();
            var used = new HashSet<int>();
            foreach (int s in next.Keys)
            {
                if (used.Contains(s)) continue;
                var loop = new List<int>();
                int c = s;
                do
                {
                    if (!used.Add(c)) { why = "контур не замкнут"; return null; }
                    loop.Add(c);
                    if (!next.TryGetValue(c, out c)) { why = "контур не замкнут"; return null; }
                } while (c != s);
                if (loop.Count < 3) { why = "вырожденный контур"; return null; }
                loops.Add(loop);
            }
            // вершины внутри грани, которых нет на контуре - то, что мы выкидываем
            if (used.Count == tris.SelectMany(i => new[] { i }).Distinct().Count() && tris.Count / 3 == used.Count - 2 + 2 * (loops.Count - 1))
                return tris;   // сетка Inventor уже минимальна

            // проекция на плоскость (u x w = n, т.е. CCW в 2D = по нормали)
            Vec u = Math.Abs(n.x) < 0.9 ? Vec.cross(n, new Vec(1, 0, 0)).norm() : Vec.cross(n, new Vec(0, 1, 0)).norm();
            Vec w = Vec.cross(n, u);
            Func<int, double> X = i => Vec.dot(P[i], u), Y = i => Vec.dot(P[i], w);
            Func<int, int, int, double> area = (a, b, c) => 0.5 * ((X(b) - X(a)) * (Y(c) - Y(a)) - (X(c) - X(a)) * (Y(b) - Y(a)));
            Func<List<int>, double> loopArea = l =>
            {
                double s = 0;
                for (int i = 0; i < l.Count; i++) { int a = l[i], b = l[(i + 1) % l.Count]; s += X(a) * Y(b) - X(b) * Y(a); }
                return s * 0.5;
            };

            // внешний контур - максимальный по площади
            int oi = 0;
            for (int i = 1; i < loops.Count; i++)
                if (Math.Abs(loopArea(loops[i])) > Math.Abs(loopArea(loops[oi]))) oi = i;
            var order = new List<List<int>> { loops[oi] };
            order.AddRange(loops.Where((l, i) => i != oi));

            var ids = new List<int>();
            var coords = new List<double>();
            var holes = new List<int>();
            foreach (var l in order)
            {
                if (ids.Count > 0) holes.Add(ids.Count);
                foreach (int i in l) { ids.Add(i); coords.Add(X(i)); coords.Add(Y(i)); }
            }
            List<int> et = Earcut.run(coords.ToArray(), holes.ToArray());

            var res = new List<int>(et.Count);
            double srcArea = 0, dstArea = 0;
            for (int t = 0; t < tris.Count; t += 3) srcArea += area(tris[t], tris[t + 1], tris[t + 2]);
            for (int t = 0; t < et.Count; t += 3)
            {
                int a = ids[et[t]], b = ids[et[t + 1]], c = ids[et[t + 2]];
                double ar = area(a, b, c);
                if (ar < 0) { int x = b; b = c; c = x; ar = -ar; }
                dstArea += ar;
                res.Add(a); res.Add(b); res.Add(c);
            }
            if (res.Count == 0 || Math.Abs(dstArea - srcArea) > Math.Abs(srcArea) * 1e-4) { why = "площадь не сошлась"; return null; }
            return res;
        }

        static long key(int a, int b) { return ((long)a << 32) | (uint)b; }
    }

    // порт mapbox/earcut (ISC): триангуляция многоугольника с отверстиями
    static class Earcut
    {
        class Node
        {
            public int i; public double x, y; public int z; public bool steiner;
            public Node prev, next, prevZ, nextZ;
            public Node(int i, double x, double y) { this.i = i; this.x = x; this.y = y; }
        }

        public static List<int> run(double[] data, int[] holeIndices)
        {
            var tris = new List<int>();
            bool hasHoles = holeIndices != null && holeIndices.Length > 0;
            int outerLen = hasHoles ? holeIndices[0] * 2 : data.Length;
            Node outer = linkedList(data, 0, outerLen, true);
            if (outer == null || outer.next == outer.prev) return tris;
            if (hasHoles) outer = eliminateHoles(data, holeIndices, outer);
            double minX = 0, minY = 0, invSize = 0;
            if (data.Length > 80 * 2)
            {
                minX = double.MaxValue; minY = double.MaxValue;
                double maxX = double.MinValue, maxY = double.MinValue;
                for (int i = 0; i < outerLen; i += 2)
                {
                    minX = Math.Min(minX, data[i]); minY = Math.Min(minY, data[i + 1]);
                    maxX = Math.Max(maxX, data[i]); maxY = Math.Max(maxY, data[i + 1]);
                }
                invSize = Math.Max(maxX - minX, maxY - minY);
                invSize = invSize != 0 ? 32767 / invSize : 0;
            }
            earcutLinked(outer, tris, minX, minY, invSize, 0);
            return tris;
        }

        static Node linkedList(double[] data, int start, int end, bool clockwise)
        {
            Node last = null;
            if (clockwise == (signedArea(data, start, end) > 0))
                for (int i = start; i < end; i += 2) last = insertNode(i / 2, data[i], data[i + 1], last);
            else
                for (int i = end - 2; i >= start; i -= 2) last = insertNode(i / 2, data[i], data[i + 1], last);
            if (last != null && equals(last, last.next)) { removeNode(last); last = last.next; }
            return last;
        }

        static Node filterPoints(Node start, Node end = null)
        {
            if (start == null) return start;
            if (end == null) end = start;
            Node p = start;
            bool again;
            do
            {
                again = false;
                if (!p.steiner && (equals(p, p.next) || area(p.prev, p, p.next) == 0))
                {
                    removeNode(p);
                    p = end = p.prev;
                    if (p == p.next) break;
                    again = true;
                }
                else p = p.next;
            } while (again || p != end);
            return end;
        }

        static void earcutLinked(Node ear, List<int> tris, double minX, double minY, double invSize, int pass)
        {
            if (ear == null) return;
            if (pass == 0 && invSize != 0) indexCurve(ear, minX, minY, invSize);
            Node stop = ear;
            while (ear.prev != ear.next)
            {
                Node prev = ear.prev, next = ear.next;
                if (invSize != 0 ? isEarHashed(ear, minX, minY, invSize) : isEar(ear))
                {
                    tris.Add(prev.i); tris.Add(ear.i); tris.Add(next.i);
                    removeNode(ear);
                    ear = next.next;
                    stop = next.next;
                    continue;
                }
                ear = next;
                if (ear == stop)
                {
                    if (pass == 0) earcutLinked(filterPoints(ear), tris, minX, minY, invSize, 1);
                    else if (pass == 1)
                    {
                        ear = cureLocalIntersections(filterPoints(ear), tris);
                        earcutLinked(ear, tris, minX, minY, invSize, 2);
                    }
                    else if (pass == 2) splitEarcut(ear, tris, minX, minY, invSize);
                    break;
                }
            }
        }

        static bool isEar(Node ear)
        {
            Node a = ear.prev, b = ear, c = ear.next;
            if (area(a, b, c) >= 0) return false;
            double x0 = Math.Min(a.x, Math.Min(b.x, c.x)), y0 = Math.Min(a.y, Math.Min(b.y, c.y));
            double x1 = Math.Max(a.x, Math.Max(b.x, c.x)), y1 = Math.Max(a.y, Math.Max(b.y, c.y));
            Node p = c.next;
            while (p != a)
            {
                if (p.x >= x0 && p.x <= x1 && p.y >= y0 && p.y <= y1 &&
                    pointInTriangle(a.x, a.y, b.x, b.y, c.x, c.y, p.x, p.y) && area(p.prev, p, p.next) >= 0) return false;
                p = p.next;
            }
            return true;
        }

        static bool isEarHashed(Node ear, double minX, double minY, double invSize)
        {
            Node a = ear.prev, b = ear, c = ear.next;
            if (area(a, b, c) >= 0) return false;
            double x0 = Math.Min(a.x, Math.Min(b.x, c.x)), y0 = Math.Min(a.y, Math.Min(b.y, c.y));
            double x1 = Math.Max(a.x, Math.Max(b.x, c.x)), y1 = Math.Max(a.y, Math.Max(b.y, c.y));
            int minZ = zOrder(x0, y0, minX, minY, invSize), maxZ = zOrder(x1, y1, minX, minY, invSize);
            Node p = ear.prevZ, n = ear.nextZ;
            while (p != null && p.z >= minZ && n != null && n.z <= maxZ)
            {
                if (blocks(p, a, b, c, x0, y0, x1, y1)) return false;
                p = p.prevZ;
                if (blocks(n, a, b, c, x0, y0, x1, y1)) return false;
                n = n.nextZ;
            }
            while (p != null && p.z >= minZ)
            {
                if (blocks(p, a, b, c, x0, y0, x1, y1)) return false;
                p = p.prevZ;
            }
            while (n != null && n.z <= maxZ)
            {
                if (blocks(n, a, b, c, x0, y0, x1, y1)) return false;
                n = n.nextZ;
            }
            return true;
        }

        static bool blocks(Node p, Node a, Node b, Node c, double x0, double y0, double x1, double y1)
        {
            return p.x >= x0 && p.x <= x1 && p.y >= y0 && p.y <= y1 && p != a && p != c &&
                pointInTriangle(a.x, a.y, b.x, b.y, c.x, c.y, p.x, p.y) && area(p.prev, p, p.next) >= 0;
        }

        static Node cureLocalIntersections(Node start, List<int> tris)
        {
            Node p = start;
            do
            {
                Node a = p.prev, b = p.next.next;
                if (!equals(a, b) && intersects(a, p, p.next, b) && locallyInside(a, b) && locallyInside(b, a))
                {
                    tris.Add(a.i); tris.Add(p.i); tris.Add(b.i);
                    removeNode(p);
                    removeNode(p.next);
                    p = start = b;
                }
                p = p.next;
            } while (p != start);
            return filterPoints(p);
        }

        static void splitEarcut(Node start, List<int> tris, double minX, double minY, double invSize)
        {
            Node a = start;
            do
            {
                Node b = a.next.next;
                while (b != a.prev)
                {
                    if (a.i != b.i && isValidDiagonal(a, b))
                    {
                        Node c = splitPolygon(a, b);
                        a = filterPoints(a, a.next);
                        c = filterPoints(c, c.next);
                        earcutLinked(a, tris, minX, minY, invSize, 0);
                        earcutLinked(c, tris, minX, minY, invSize, 0);
                        return;
                    }
                    b = b.next;
                }
                a = a.next;
            } while (a != start);
        }

        static Node eliminateHoles(double[] data, int[] holeIndices, Node outer)
        {
            var queue = new List<Node>();
            for (int i = 0; i < holeIndices.Length; i++)
            {
                int start = holeIndices[i] * 2;
                int end = i < holeIndices.Length - 1 ? holeIndices[i + 1] * 2 : data.Length;
                Node list = linkedList(data, start, end, false);
                if (list == null) continue;
                if (list == list.next) list.steiner = true;
                queue.Add(getLeftmost(list));
            }
            queue.Sort((a, b) => a.x != b.x ? a.x.CompareTo(b.x) : a.y.CompareTo(b.y));
            foreach (Node h in queue) outer = eliminateHole(h, outer);
            return outer;
        }

        static Node eliminateHole(Node hole, Node outer)
        {
            Node bridge = findHoleBridge(hole, outer);
            if (bridge == null) return outer;
            Node bridgeReverse = splitPolygon(bridge, hole);
            filterPoints(bridgeReverse, bridgeReverse.next);
            return filterPoints(bridge, bridge.next);
        }

        static Node findHoleBridge(Node hole, Node outer)
        {
            Node p = outer, m = null;
            double hx = hole.x, hy = hole.y, qx = double.NegativeInfinity;
            do
            {
                if (hy <= p.y && hy >= p.next.y && p.next.y != p.y)
                {
                    double x = p.x + (hy - p.y) * (p.next.x - p.x) / (p.next.y - p.y);
                    if (x <= hx && x > qx)
                    {
                        qx = x;
                        m = p.x < p.next.x ? p : p.next;
                        if (x == hx) return m;
                    }
                }
                p = p.next;
            } while (p != outer);
            if (m == null) return null;
            Node stop = m;
            double mx = m.x, my = m.y, tanMin = double.PositiveInfinity;
            p = m;
            do
            {
                if (hx >= p.x && p.x >= mx && hx != p.x &&
                    pointInTriangle(hy < my ? hx : qx, hy, mx, my, hy < my ? qx : hx, hy, p.x, p.y))
                {
                    double tan = Math.Abs(hy - p.y) / (hx - p.x);
                    if (locallyInside(p, hole) &&
                        (tan < tanMin || (tan == tanMin && (p.x > m.x || (p.x == m.x && sectorContainsSector(m, p))))))
                    {
                        m = p;
                        tanMin = tan;
                    }
                }
                p = p.next;
            } while (p != stop);
            return m;
        }

        static bool sectorContainsSector(Node m, Node p) { return area(m.prev, m, p.prev) < 0 && area(p.next, m, m.next) < 0; }

        static void indexCurve(Node start, double minX, double minY, double invSize)
        {
            Node p = start;
            do
            {
                if (p.z == 0) p.z = zOrder(p.x, p.y, minX, minY, invSize);
                p.prevZ = p.prev;
                p.nextZ = p.next;
                p = p.next;
            } while (p != start);
            p.prevZ.nextZ = null;
            p.prevZ = null;
            sortLinked(p);
        }

        static Node sortLinked(Node list)
        {
            int inSize = 1, numMerges;
            do
            {
                Node p = list, tail = null;
                list = null;
                numMerges = 0;
                while (p != null)
                {
                    numMerges++;
                    Node q = p;
                    int pSize = 0;
                    for (int i = 0; i < inSize; i++)
                    {
                        pSize++;
                        q = q.nextZ;
                        if (q == null) break;
                    }
                    int qSize = inSize;
                    while (pSize > 0 || (qSize > 0 && q != null))
                    {
                        Node e;
                        if (pSize != 0 && (qSize == 0 || q == null || p.z <= q.z)) { e = p; p = p.nextZ; pSize--; }
                        else { e = q; q = q.nextZ; qSize--; }
                        if (tail != null) tail.nextZ = e; else list = e;
                        e.prevZ = tail;
                        tail = e;
                    }
                    p = q;
                }
                tail.nextZ = null;
                inSize *= 2;
            } while (numMerges > 1);
            return list;
        }

        static int zOrder(double x0, double y0, double minX, double minY, double invSize)
        {
            int x = (int)((x0 - minX) * invSize), y = (int)((y0 - minY) * invSize);
            x = (x | (x << 8)) & 0x00FF00FF; x = (x | (x << 4)) & 0x0F0F0F0F;
            x = (x | (x << 2)) & 0x33333333; x = (x | (x << 1)) & 0x55555555;
            y = (y | (y << 8)) & 0x00FF00FF; y = (y | (y << 4)) & 0x0F0F0F0F;
            y = (y | (y << 2)) & 0x33333333; y = (y | (y << 1)) & 0x55555555;
            return x | (y << 1);
        }

        static Node getLeftmost(Node start)
        {
            Node p = start, leftmost = start;
            do
            {
                if (p.x < leftmost.x || (p.x == leftmost.x && p.y < leftmost.y)) leftmost = p;
                p = p.next;
            } while (p != start);
            return leftmost;
        }

        static bool pointInTriangle(double ax, double ay, double bx, double by, double cx, double cy, double px, double py)
        {
            return (cx - px) * (ay - py) >= (ax - px) * (cy - py) &&
                   (ax - px) * (by - py) >= (bx - px) * (ay - py) &&
                   (bx - px) * (cy - py) >= (cx - px) * (by - py);
        }

        static bool isValidDiagonal(Node a, Node b)
        {
            return a.next.i != b.i && a.prev.i != b.i && !intersectsPolygon(a, b) &&
                ((locallyInside(a, b) && locallyInside(b, a) && middleInside(a, b) &&
                  (area(a.prev, a, b.prev) != 0 || area(a, b.prev, b) != 0)) ||
                 (equals(a, b) && area(a.prev, a, a.next) > 0 && area(b.prev, b, b.next) > 0));
        }

        static double area(Node p, Node q, Node r) { return (q.y - p.y) * (r.x - q.x) - (q.x - p.x) * (r.y - q.y); }
        static bool equals(Node a, Node b) { return a.x == b.x && a.y == b.y; }

        static bool intersects(Node p1, Node q1, Node p2, Node q2)
        {
            int o1 = Math.Sign(area(p1, q1, p2)), o2 = Math.Sign(area(p1, q1, q2));
            int o3 = Math.Sign(area(p2, q2, p1)), o4 = Math.Sign(area(p2, q2, q1));
            if (o1 != o2 && o3 != o4) return true;
            if (o1 == 0 && onSegment(p1, p2, q1)) return true;
            if (o2 == 0 && onSegment(p1, q2, q1)) return true;
            if (o3 == 0 && onSegment(p2, p1, q2)) return true;
            if (o4 == 0 && onSegment(p2, q1, q2)) return true;
            return false;
        }

        static bool onSegment(Node p, Node q, Node r)
        {
            return q.x <= Math.Max(p.x, r.x) && q.x >= Math.Min(p.x, r.x) && q.y <= Math.Max(p.y, r.y) && q.y >= Math.Min(p.y, r.y);
        }

        static bool intersectsPolygon(Node a, Node b)
        {
            Node p = a;
            do
            {
                if (p.i != a.i && p.next.i != a.i && p.i != b.i && p.next.i != b.i && intersects(p, p.next, a, b)) return true;
                p = p.next;
            } while (p != a);
            return false;
        }

        static bool locallyInside(Node a, Node b)
        {
            return area(a.prev, a, a.next) < 0 ?
                area(a, b, a.next) >= 0 && area(a, a.prev, b) >= 0 :
                area(a, b, a.prev) < 0 || area(a, a.next, b) < 0;
        }

        static bool middleInside(Node a, Node b)
        {
            Node p = a;
            bool inside = false;
            double px = (a.x + b.x) / 2, py = (a.y + b.y) / 2;
            do
            {
                if (((p.y > py) != (p.next.y > py)) && p.next.y != p.y && (px < (p.next.x - p.x) * (py - p.y) / (p.next.y - p.y) + p.x))
                    inside = !inside;
                p = p.next;
            } while (p != a);
            return inside;
        }

        static Node splitPolygon(Node a, Node b)
        {
            Node a2 = new Node(a.i, a.x, a.y), b2 = new Node(b.i, b.x, b.y), an = a.next, bp = b.prev;
            a.next = b; b.prev = a;
            a2.next = an; an.prev = a2;
            b2.next = a2; a2.prev = b2;
            bp.next = b2; b2.prev = bp;
            return b2;
        }

        static Node insertNode(int i, double x, double y, Node last)
        {
            Node p = new Node(i, x, y);
            if (last == null) { p.prev = p; p.next = p; }
            else { p.next = last.next; p.prev = last; last.next.prev = p; last.next = p; }
            return p;
        }

        static void removeNode(Node p)
        {
            p.next.prev = p.prev;
            p.prev.next = p.next;
            if (p.prevZ != null) p.prevZ.nextZ = p.nextZ;
            if (p.nextZ != null) p.nextZ.prevZ = p.prevZ;
        }

        static double signedArea(double[] data, int start, int end)
        {
            double sum = 0;
            for (int i = start, j = end - 2; i < end; i += 2) { sum += (data[j] - data[i]) * (data[i + 1] + data[j + 1]); j = i; }
            return sum;
        }
    }

    // Сетка тела, сшитая по позициям, с точными нормалями поверхности для каждой пары (вершина, грань).
    // simplify() - схлопывание рёбер по квадрикам (QEM) с сохранением границ граней:
    // вершина v может слиться в u, только если все грани v содержат и u (рёбра, углы и границы материалов не плывут).
    // Плоские грани как сетку не упрощаем: упрощается их контур, после чего грань заново триангулируется earcut'ом.
    public class BodyMesher
    {
        public List<Vec> P = new List<Vec>();
        public List<int> T = new List<int>();        // тройки вершин
        public List<int> TF = new List<int>();       // грань треугольника
        public List<int> faceMat = new List<int>();  // материал грани
        List<Vec> faceN = new List<Vec>();           // нормаль плоской грани, (0,0,0) - криволинейная
        Dictionary<Key3, int> keys = new Dictionary<Key3, int>();
        Dictionary<long, Vec> nrm = new Dictionary<long, Vec>(LongCmp.I);
        public double ErrorFactor = 1.0;             // множитель допуска
        public double MaxTurnDeg = 22.5;             // не больше этого угла на сегмент окружности (16 сегментов)
        public double MaxEdgeCos { get { return Math.Cos(MaxTurnDeg * Math.PI / 180) - 1e-9; } }   // то же для нормалей на ребре
        public int MaxValence = 64;                  // вершины с большим числом треугольников не трогаем
        public int TimeLimitMs = 0;                  // 0 - без ограничения
        public bool timedOut;

        // исходная сетка в текст для отладки: v x y z / t a b c f / f planar nx ny nz / n v f nx ny nz
        public void dump(string fn)
        {
            var ci = CultureInfo.InvariantCulture;
            using (var w = new StreamWriter(fn, false, new UTF8Encoding(false)))
            {
                w.WriteLine(P.Count + " " + TF.Count + " " + faceMat.Count + " " + nrm.Count);
                foreach (Vec p in P) w.WriteLine(p.x.ToString("R", ci) + " " + p.y.ToString("R", ci) + " " + p.z.ToString("R", ci));
                for (int t = 0; t < TF.Count; t++) w.WriteLine(T[3 * t] + " " + T[3 * t + 1] + " " + T[3 * t + 2] + " " + TF[t]);
                for (int f = 0; f < faceMat.Count; f++)
                    w.WriteLine((planar(f) ? 1 : 0) + " " + faceN[f].x.ToString("R", ci) + " " + faceN[f].y.ToString("R", ci) + " " + faceN[f].z.ToString("R", ci));
                foreach (var kv in nrm)
                    w.WriteLine((kv.Key >> 32) + " " + (int)(kv.Key & 0xffffffff) + " " + kv.Value.x.ToString("R", ci) + " " + kv.Value.y.ToString("R", ci) + " " + kv.Value.z.ToString("R", ci));
            }
        }

        static long nk(int v, int f) { return ((long)v << 32) | (uint)f; }

        // planeN - нормаль, если грань плоская и триангулирована по контуру (иначе new Vec())
        public int addFace(int mat, Vec planeN) { faceMat.Add(mat); faceN.Add(planeN); return faceMat.Count - 1; }

        bool planar(int f) { return faceN[f].x != 0 || faceN[f].y != 0 || faceN[f].z != 0; }

        public int vertex(Vec p)
        {
            var k = new Key3(p);
            int id;
            if (!keys.TryGetValue(k, out id)) { id = P.Count; keys[k] = id; P.Add(p); }
            return id;
        }

        public void normal(int v, int f, Vec n)
        {
            Vec o;
            nrm[nk(v, f)] = nrm.TryGetValue(nk(v, f), out o) ? o + n : n;
        }

        public void tri(int f, int a, int b, int c)
        {
            if (a == b || b == c || a == c) return;
            T.Add(a); T.Add(b); T.Add(c); TF.Add(f);
        }

        public int triCount { get { return TF.Count; } }

        // состояние упрощения
        int nv, nt;
        bool[] alive, locked;
        List<int>[] vt, vf;
        Dictionary<long, int> lnext, lprev, loopOf;
        Dictionary<int, int> bnext, bprev, bLoopOf;       // открытая граница
        Dictionary<int, List<Vec>> babs;                  // точки, убранные с граничного ребра (начало ребра -> точки)
        List<int> bLoopLen;
        List<int> loopLen, loopMin;
        public int MinLoop = 6;

        public int simplify(double tol)
        {
            nv = P.Count; nt = TF.Count;
            if (nt == 0 || tol <= 0) return nt;
            alive = new bool[nt];
            locked = new bool[nv];
            vt = new List<int>[nv];
            vf = new List<int>[nv];
            lnext = new Dictionary<long, int>(LongCmp.I);
            lprev = new Dictionary<long, int>(LongCmp.I);
            loopOf = new Dictionary<long, int>(LongCmp.I);
            loopLen = new List<int>();
            loopMin = new List<int>();
            for (int i = 0; i < nv; i++) { vt[i] = new List<int>(6); vf[i] = new List<int>(2); }

            // замкнутость: вершины на открытой границе (поверхностные тела) не трогаем
            var edgeUse = new Dictionary<long, int>(LongCmp.I);
            for (int t = 0; t < nt; t++)
            {
                alive[t] = true;
                for (int k = 0; k < 3; k++)
                {
                    int a = T[3 * t + k], b = T[3 * t + (k + 1) % 3];
                    long e = a < b ? nk(a, b) : nk(b, a);
                    int c;
                    edgeUse.TryGetValue(e, out c);
                    edgeUse[e] = c + 1;
                    int pos = vf[a].BinarySearch(TF[t]);
                    if (pos < 0) vf[a].Insert(~pos, TF[t]);
                }
            }
            // открытая граница (обрезанные/выброшенные грани, поверхностные тела): вершины двигаются только вдоль неё
            bnext = new Dictionary<int, int>(); bprev = new Dictionary<int, int>();
            babs = new Dictionary<int, List<Vec>>();
            bLoopOf = new Dictionary<int, int>(); bLoopLen = new List<int>();
            foreach (var kv in edgeUse)
                if (kv.Value > 2) { locked[(int)(kv.Key >> 32)] = true; locked[(int)(kv.Key & 0xffffffff)] = true; }
            for (int t = 0; t < nt; t++)
                for (int k = 0; k < 3; k++)
                {
                    int a = T[3 * t + k], b = T[3 * t + (k + 1) % 3];
                    if (edgeUse[a < b ? nk(a, b) : nk(b, a)] != 1) continue;
                    if (bnext.ContainsKey(a) || bprev.ContainsKey(b)) { locked[a] = locked[b] = true; continue; }
                    bnext[a] = b; bprev[b] = a;
                }
            foreach (int s0 in bnext.Keys.ToList())
            {
                if (bLoopOf.ContainsKey(s0)) continue;
                int c = s0, len = 0;
                bool closed = true;
                do
                {
                    bLoopOf[c] = bLoopLen.Count;
                    len++;
                    if (!bnext.TryGetValue(c, out c) || len > nv) { closed = false; break; }
                } while (c != s0);
                bLoopLen.Add(len);
                if (!closed) foreach (var kv in bLoopOf) if (kv.Value == bLoopLen.Count - 1) locked[kv.Key] = true;
            }
            foreach (int v in bnext.Keys) if (!bprev.ContainsKey(v)) locked[v] = true;
            foreach (int v in bprev.Keys) if (!bnext.ContainsKey(v)) locked[v] = true;

            // контуры плоских граней
            var byFace = new Dictionary<int, List<int>>();
            for (int t = 0; t < nt; t++)
                if (planar(TF[t]))
                {
                    List<int> l;
                    if (!byFace.TryGetValue(TF[t], out l)) byFace[TF[t]] = l = new List<int>();
                    l.Add(t);
                }
            foreach (var kv in byFace)
            {
                int f = kv.Key;
                var dir = new HashSet<long>(LongCmp.I);
                foreach (int t in kv.Value)
                    for (int k = 0; k < 3; k++) dir.Add(nk(T[3 * t + k], T[3 * t + (k + 1) % 3]));
                bool rigid = false;
                foreach (long e in dir)
                {
                    int a = (int)(e >> 32), b = (int)(e & 0xffffffff);
                    if (dir.Contains(nk(b, a))) continue;
                    if (lnext.ContainsKey(nk(a, f)) || lprev.ContainsKey(nk(b, f))) { rigid = true; continue; }
                    lnext[nk(a, f)] = b;
                    lprev[nk(b, f)] = a;
                }
                if (rigid)
                {
                    fail("плоская грань заблокирована (контуры касаются)");
                    foreach (int t in kv.Value) for (int k = 0; k < 3; k++) locked[T[3 * t + k]] = true;
                }
                // номера контуров и их длины: контур не упрощаем меньше MinLoop вершин
                foreach (int t in kv.Value)
                    for (int k = 0; k < 3; k++)
                    {
                        int s = T[3 * t + k], c = s, len = 0;
                        if (loopOf.ContainsKey(nk(s, f)) || !lnext.ContainsKey(nk(s, f))) continue;
                        do
                        {
                            loopOf[nk(c, f)] = loopLen.Count;
                            len++;
                        } while (lnext.TryGetValue(nk(c, f), out c) && c != s && len <= nv);
                        loopLen.Add(len);
                        loopMin.Add(Math.Min(len, MinLoop));
                    }
            }

            // смежность криволинейных граней
            for (int t = 0; t < nt; t++)
            {
                if (planar(TF[t])) continue;
                vt[T[3 * t]].Add(t); vt[T[3 * t + 1]].Add(t); vt[T[3 * t + 2]].Add(t);
            }

            double lim = tol * ErrorFactor;
            var touched = new bool[nv];
            var dirty = new bool[nv];   // после первого прохода пересматриваем только рёбра у изменившихся вершин
            for (int i = 0; i < nv; i++) dirty[i] = true;
            var clock = System.Diagnostics.Stopwatch.StartNew();
            timedOut = false;
            for (int pass = 0; pass < 100; pass++)
            {
                if (TimeLimitMs > 0 && clock.ElapsedMilliseconds > TimeLimitMs) { timedOut = true; break; }
                var cand = new List<KeyValuePair<double, long>>();
                var seen = new HashSet<long>(LongCmp.I);
                Action<int, int> consider = (a, b) =>
                {
                    if (!dirty[a] && !dirty[b]) return;
                    if (vt[a].Count > MaxValence || vt[b].Count > MaxValence) return;   // веера - дорого и незачем
                    if (!seen.Add(a < b ? nk(a, b) : nk(b, a))) return;
                    double best = double.MaxValue;
                    long move = -1;
                    if (allowed(a, b)) { double c = cost(a, b); if (c < best) { best = c; move = nk(a, b); } }
                    if (allowed(b, a)) { double c = cost(b, a); if (c < best) { best = c; move = nk(b, a); } }
                    if (move >= 0 && best <= lim) cand.Add(new KeyValuePair<double, long>(best, move));
                };
                for (int t = 0; t < nt; t++)
                    if (alive[t] && !planar(TF[t]))
                        for (int k = 0; k < 3; k++) consider(T[3 * t + k], T[3 * t + (k + 1) % 3]);
                foreach (var kv in lnext) consider((int)(kv.Key >> 32), kv.Value);
                if (cand.Count == 0) break;
                cand.Sort((x, y) => x.Key.CompareTo(y.Key));
                Array.Clear(touched, 0, nv);
                int done = 0;
                foreach (var c in cand)
                {
                    int v = (int)(c.Value >> 32), u = (int)(c.Value & 0xffffffff);
                    if (touched[v] || touched[u]) continue;
                    if (vt[u].Count > MaxValence) continue;
                    if (!allowed(v, u) || !canCollapse(v, u)) continue;
                    collapse(v, u, touched);
                    done++;
                }
                if (done == 0) break;
                Array.Copy(touched, dirty, nv);
            }

            // сборка: криволинейные треугольники + заново триангулированные плоские грани
            var nT = new List<int>();
            var nTF = new List<int>();
            for (int t = 0; t < nt; t++)
                if (alive[t] && !planar(TF[t])) { nT.Add(T[3 * t]); nT.Add(T[3 * t + 1]); nT.Add(T[3 * t + 2]); nTF.Add(TF[t]); }
            foreach (var kv in byFace)
            {
                int f = kv.Key;
                if (kv.Value.Any(t => locked[T[3 * t]] && !lnext.ContainsKey(nk(T[3 * t], f))) || !retriangulate(f, kv.Value, nT, nTF))
                    foreach (int t in kv.Value) { nT.Add(T[3 * t]); nT.Add(T[3 * t + 1]); nT.Add(T[3 * t + 2]); nTF.Add(f); }
            }
            T = nT; TF = nTF;
            return TF.Count;
        }

        public Dictionary<string, int> failures = new Dictionary<string, int>();

        bool fail(string why)
        {
            int c;
            failures.TryGetValue(why, out c);
            failures[why] = c + 1;
            return false;
        }

        bool retriangulate(int f, List<int> origTris, List<int> outT, List<int> outTF)
        {
            // контуры после схлопываний
            var loops = new List<List<int>>();
            var used = new HashSet<int>();
            foreach (int t in origTris)
                for (int k = 0; k < 3; k++)
                {
                    int s = T[3 * t + k];
                    if (used.Contains(s) || !lnext.ContainsKey(nk(s, f))) continue;
                    var loop = new List<int>();
                    int c = s;
                    do
                    {
                        if (!used.Add(c) || loop.Count > nv) return fail("контур пересекает сам себя");
                        loop.Add(c);
                        if (!lnext.TryGetValue(nk(c, f), out c)) return fail("контур разорван");
                    } while (c != s);
                    if (loop.Count < 3) return fail("вырожденный контур");
                    loops.Add(loop);
                }
            if (loops.Count == 0) return fail("нет контура");
            // если ничего не схлопнулось - оставляем как было
            bool changed = false;
            foreach (int t in origTris) for (int k = 0; k < 3; k++) if (!used.Contains(T[3 * t + k])) changed = true;
            if (!changed) return false;

            Vec n = faceN[f];
            Vec u = Math.Abs(n.x) < 0.9 ? Vec.cross(n, new Vec(1, 0, 0)).norm() : Vec.cross(n, new Vec(0, 1, 0)).norm();
            Vec w = Vec.cross(n, u);
            Func<List<int>, double> area = l =>
            {
                double s = 0;
                for (int i = 0; i < l.Count; i++)
                {
                    Vec a = P[l[i]], b = P[l[(i + 1) % l.Count]];
                    s += Vec.dot(a, u) * Vec.dot(b, w) - Vec.dot(b, u) * Vec.dot(a, w);
                }
                return s * 0.5;
            };
            int oi = 0;
            for (int i = 1; i < loops.Count; i++) if (Math.Abs(area(loops[i])) > Math.Abs(area(loops[oi]))) oi = i;
            var order = new List<List<int>> { loops[oi] };
            order.AddRange(loops.Where((l, i) => i != oi));
            var ids = new List<int>();
            var coords = new List<double>();
            var holes = new List<int>();
            foreach (var l in order)
            {
                if (ids.Count > 0) holes.Add(ids.Count);
                foreach (int i in l) { ids.Add(i); coords.Add(Vec.dot(P[i], u)); coords.Add(Vec.dot(P[i], w)); }
            }
            double expect = Math.Abs(area(order[0])) - order.Skip(1).Sum(l => Math.Abs(area(l)));
            var et = Earcut.run(coords.ToArray(), holes.ToArray());
            double got = 0;
            var res = new List<int>();
            for (int i = 0; i < et.Count; i += 3)
            {
                int a = ids[et[i]], b = ids[et[i + 1]], c = ids[et[i + 2]];
                Vec cr = Vec.cross(P[b] - P[a], P[c] - P[a]);
                if (Vec.dot(cr, n) < 0) { int x = b; b = c; c = x; }
                got += Math.Abs(Vec.dot(cr, n)) * 0.5;
                res.Add(a); res.Add(b); res.Add(c);
            }
            if (res.Count == 0 || Math.Abs(got - expect) > Math.Abs(expect) * 1e-3) return fail("earcut: площадь не сошлась");
            outT.AddRange(res);
            for (int i = 0; i < res.Count / 3; i++) outTF.Add(f);
            return true;
        }

        bool allowed(int v, int u)
        {
            if (locked[v]) return false;
            int bn;
            if (bnext.TryGetValue(v, out bn))
            {
                // на открытой границе - только к соседу по границе, контур не короче MinLoop
                if (bn != u && bprev[v] != u) return false;
                if (bLoopLen[bLoopOf[v]] <= MinLoop) return false;
            }
            if (locked[v] || !subset(vf[v], vf[u])) return false;
            foreach (int f in vf[v])
            {
                if (!planar(f)) continue;
                int n, p;
                // по плоской грани двигаемся только вдоль её контура, контур не короче 3
                if (!lnext.TryGetValue(nk(v, f), out n) || !lprev.TryGetValue(nk(v, f), out p)) return false;
                if (n != u && p != u) return false;
                int li;
                if (!loopOf.TryGetValue(nk(v, f), out li) || loopLen[li] <= loopMin[li]) return false;
            }
            return true;
        }

        bool canCollapse(int v, int u)
        {
            // топология криволинейной части: общих соседей столько же, сколько треугольников на ребре
            var nu = new HashSet<int>();
            foreach (int t in vt[u]) if (alive[t]) for (int k = 0; k < 3; k++) nu.Add(T[3 * t + k]);
            var nvs = new HashSet<int>();
            int shared = 0;
            foreach (int t in vt[v])
            {
                if (!alive[t]) continue;
                bool hasU = false;
                for (int k = 0; k < 3; k++) { nvs.Add(T[3 * t + k]); if (T[3 * t + k] == u) hasU = true; }
                if (hasU) shared++;
            }
            int common = 0;
            foreach (int x in nvs) if (x != u && x != v && nu.Contains(x)) common++;
            if (common != shared) return false;
            // геометрия: треугольники не переворачиваются и не вырождаются
            foreach (int t in vt[v])
            {
                if (!alive[t]) continue;
                int a = T[3 * t], b = T[3 * t + 1], c = T[3 * t + 2];
                if (a == u || b == u || c == u) continue;
                Vec n0 = Vec.cross(P[b] - P[a], P[c] - P[a]);
                Vec pa = a == v ? P[u] : P[a], pb = b == v ? P[u] : P[b], pc = c == v ? P[u] : P[c];
                Vec n1 = Vec.cross(pb - pa, pc - pa);
                double l0 = n0.len(), l1 = n1.len();
                if (l1 == 0 || l1 <= 1e-12 * l0) return false;
                if (Vec.dot(n0, n1) < 0.3 * l0 * l1) return false;
                Vec sn;
                if (nrm.TryGetValue(nk(u, TF[t]), out sn) && Vec.dot(n1, sn) <= 0) return false;
            }
            return true;
        }

        void collapse(int v, int u, bool[] touched)
        {
            int bn;
            if (bnext.TryGetValue(v, out bn))
            {
                int bp = bprev[v];
                List<Vec> l, lv;
                if (!babs.TryGetValue(bp, out l)) babs[bp] = l = new List<Vec>();
                l.Add(P[v]);
                if (babs.TryGetValue(v, out lv)) { l.AddRange(lv); babs.Remove(v); }
                bnext[bp] = bn; bprev[bn] = bp;
                bnext.Remove(v); bprev.Remove(v);
                bLoopLen[bLoopOf[v]]--;
                touched[bp] = touched[bn] = true;
            }
            foreach (int t in vt[v])
            {
                if (!alive[t]) continue;
                if (T[3 * t] == u || T[3 * t + 1] == u || T[3 * t + 2] == u) { alive[t] = false; continue; }
                for (int k = 0; k < 3; k++) if (T[3 * t + k] == v) T[3 * t + k] = u;
                vt[u].Add(t);
            }
            vt[v].Clear();
            vt[u].RemoveAll(t => !alive[t]);
            foreach (int f in vf[v])
            {
                if (!planar(f)) continue;
                int n = lnext[nk(v, f)], p = lprev[nk(v, f)];
                lnext[nk(p, f)] = n;
                lprev[nk(n, f)] = p;
                lnext.Remove(nk(v, f));
                lprev.Remove(nk(v, f));
                loopLen[loopOf[nk(v, f)]]--;
                touched[n] = touched[p] = true;
            }
            touched[v] = touched[u] = true;
            foreach (int t in vt[u]) for (int k = 0; k < 3; k++) touched[T[3 * t + k]] = true;
        }

        static bool subset(List<int> a, List<int> b)
        {
            foreach (int x in a) if (b.BinarySearch(x) < 0) return false;
            return true;
        }

        // отход от поверхности после переноса v в u: максимум по новым рёбрам.
        // Вершины всегда лежат на исходной поверхности, поэтому ошибка не накапливается.
        double cost(int v, int u)
        {
            double worst = 0;
            int bn;
            if (bnext.TryGetValue(v, out bn))
            {
                // новое граничное ребро p-n: отклонение всех убранных с него точек (точно, без накопления)
                int bp = bprev[v];
                if (!turnOk(bprev[bp], bp, v, bn, bnext[bn])) return double.MaxValue;
                Vec a = P[bp], b = P[bn];
                worst = segDist(P[v], a, b);
                List<Vec> l;
                if (babs.TryGetValue(bp, out l)) foreach (Vec q in l) worst = Math.Max(worst, segDist(q, a, b));
                if (babs.TryGetValue(v, out l)) foreach (Vec q in l) worst = Math.Max(worst, segDist(q, a, b));
            }
            foreach (int t in vt[v])
            {
                if (!alive[t]) continue;
                int a = T[3 * t], b = T[3 * t + 1], c = T[3 * t + 2];
                if (a == u || b == u || c == u) continue;
                if (a != v) worst = Math.Max(worst, sag(u, a, TF[t]));
                if (b != v) worst = Math.Max(worst, sag(u, b, TF[t]));
                if (c != v) worst = Math.Max(worst, sag(u, c, TF[t]));
            }
            foreach (int f in vf[v])
            {
                if (!planar(f)) continue;
                int n = lnext[nk(v, f)], p = lprev[nk(v, f)];
                int o = n == u ? p : n;   // новое ребро контура u-o
                if (!turnOk(lprev[nk(p, f)], p, v, n, lnext[nk(n, f)])) return double.MaxValue;
                double s = segDist(P[v], P[u], P[o]);
                foreach (int g in vf[u])
                    if (!planar(g) && vf[o].BinarySearch(g) >= 0) s = Math.Max(s, sag(u, o, g));
                worst = Math.Max(worst, s);
            }
            return worst;
        }

        // контур pp-p-v-n-nn после удаления v: излом в p и n не больше MaxTurn (окружности не грубее CircleSegments),
        // но уже существующие изломы (настоящие углы контура) не мешают
        bool turnOk(int pp, int p, int v, int n, int nn)
        {
            double lim = MaxTurnDeg * Math.PI / 180;
            double newP = turn(P[pp], P[p], P[n]), newN = turn(P[p], P[n], P[nn]);
            if (newP > lim && newP > turn(P[pp], P[p], P[v]) + 1e-6) return false;
            if (newN > lim && newN > turn(P[v], P[n], P[nn]) + 1e-6) return false;
            return true;
        }

        static double turn(Vec a, Vec b, Vec c)
        {
            Vec u = b - a, w = c - b;
            double l = u.len() * w.len();
            return l > 0 ? Math.Acos(Math.Max(-1, Math.Min(1, Vec.dot(u, w) / l))) : 0;
        }

        // отход хорды a-b от поверхности грани f по нормалям на концах: c/2 * tan(θ/4) (точно для окружности)
        double sag(int a, int b, int f)
        {
            Vec na, nb;
            if (!nrm.TryGetValue(nk(a, f), out na) || !nrm.TryGetValue(nk(b, f), out nb)) return 0;
            double cos = Vec.dot(na.norm(), nb.norm());
            if (cos < MaxEdgeCos) return double.MaxValue;
            return (P[b] - P[a]).len() / 2 * Math.Tan(Math.Acos(Math.Min(1, cos)) / 4);
        }

        static double segDist(Vec p, Vec a, Vec b)
        {
            Vec ab = b - a;
            double l2 = Vec.dot(ab, ab);
            double t = l2 > 0 ? Math.Max(0, Math.Min(1, Vec.dot(p - a, ab) / l2)) : 0;
            return (p - (a + ab * t)).len();
        }

        // выгрузка в меш: вершина дублируется только там, где различаются нормали или материал
        public void emit(GlbMesh m, double scale)
        {
            var map = new Dictionary<Key3, int>[faceMat.Count == 0 ? 0 : faceMat.Max() + 1];
            for (int t = 0; t < TF.Count; t++)
            {
                int f = TF[t], mat = faceMat[f];
                GlbPrim p = m.get(mat);
                if (map[mat] == null) map[mat] = new Dictionary<Key3, int>();
                for (int k = 0; k < 3; k++)
                {
                    int v = T[3 * t + k];
                    Vec n;
                    if (!nrm.TryGetValue(nk(v, f), out n) || n.len() == 0)
                    {
                        int a = T[3 * t], b = T[3 * t + 1], c = T[3 * t + 2];
                        n = Vec.cross(P[b] - P[a], P[c] - P[a]);
                    }
                    n = n.norm();
                    var key = new Key3(v, (long)Math.Round(n.x * 1000) * 4096 + (long)Math.Round(n.y * 1000), (long)Math.Round(n.z * 1000));
                    int li;
                    if (!map[mat].TryGetValue(key, out li))
                    {
                        li = p.vcount;
                        map[mat][key] = li;
                        Vec q = P[v] * scale;
                        p.pos.Add((float)q.x); p.pos.Add((float)q.y); p.pos.Add((float)q.z);
                        p.nrm.Add((float)n.x); p.nrm.Add((float)n.y); p.nrm.Add((float)n.z);
                    }
                    p.idx.Add(li);
                }
            }
        }
    }

    // Удаление геометрии, не видимой снаружи: ортографический z-буфер из многих направлений (равномерно по сфере).
    // Треугольник остаётся, если хотя бы из одного направления он виден лицевой стороной,
    // плюс его соседи по вершинам - чтобы на границе видимой области не было щелей.
    public static class Visibility
    {
        // серая морфология (max или min) квадратным окном 2r+1, по строкам и столбцам, алгоритм van Herk / Gil-Werman
        static void morph(float[] img, int res, int r, bool isMax, float[][] buf)
        {
            int w = 2 * r + 1, n = res + 2 * r;
            if (buf[0] == null || buf[0].Length < n) for (int i = 0; i < 4; i++) buf[i] = new float[n];
            float[] p = buf[0], g = buf[1], h = buf[2], o = buf[3];
            for (int axis = 0; axis < 2; axis++)
                for (int ln = 0; ln < res; ln++)
                {
                    for (int i = 0; i < n; i++)
                    {
                        int j = i - r;
                        p[i] = j < 0 || j >= res ? float.NegativeInfinity : img[axis == 0 ? ln * res + j : j * res + ln];
                    }
                    for (int i = 0; i < n; i++)
                        g[i] = i % w == 0 ? p[i] : (isMax ? Math.Max(g[i - 1], p[i]) : Math.Min(g[i - 1], p[i]));
                    for (int i = n - 1; i >= 0; i--)
                        h[i] = i == n - 1 || (i + 1) % w == 0 ? p[i] : (isMax ? Math.Max(h[i + 1], p[i]) : Math.Min(h[i + 1], p[i]));
                    for (int i = 0; i < res; i++)
                        o[i] = isMax ? Math.Max(h[i], g[i + w - 1]) : Math.Min(h[i], g[i + w - 1]);
                    for (int i = 0; i < res; i++) img[axis == 0 ? ln * res + i : i * res + ln] = o[i];
                }
        }

        // видимость треугольников снаружи (X,Y,Z в метрах, треугольники ориентированы наружу)
        class Buf
        {
            public float[] zb, cl, cl2, sx, sy, sz;
            public int[] ib;
            public List<int> small = new List<int>();
            public float[][] line = new float[4][];
            public Buf(int nv, int res)
            {
                zb = new float[res * res]; cl = new float[res * res]; cl2 = new float[res * res]; ib = new int[res * res];
                sx = new float[nv]; sy = new float[nv]; sz = new float[nv];
            }
        }

        // keepGap (м): поверхность за "заклеенным" отверстием отбрасывается, только если она глубже keepGap -
        // сквозное отверстие (за ним далеко) отличается от наружной выемки (дно близко)
        // holeClose/holeTol (м): пиксели внутри заклеенных мелких отверстий (заклейка подняла глубину больше holeTol) не засчитываются -
        // так деталь, стоящая сразу за перфорацией, не считается видимой
        public static bool[] visible(float[] X, float[] Y, float[] Z, int[] tri, int dirs, int res, double closeSize, double keepGap = 0,
            double holeClose = 0, double holeTol = 0, bool[] twoSided = null, bool[] directOut = null)
        {
            int nv = X.Length, nt = tri.Length / 3;
            if (nt == 0 || dirs <= 0 || res <= 0) { var all = new bool[nt]; for (int i = 0; i < nt; i++) all[i] = true; return all; }
            double minX = X.Min(), maxX = X.Max(), minY = Y.Min(), maxY = Y.Max(), minZ = Z.Min(), maxZ = Z.Max();
            double cx = (minX + maxX) / 2, cy = (minY + maxY) / 2, cz = (minZ + maxZ) / 2, R = 0;
            for (int i = 0; i < nv; i++)
                R = Math.Max(R, (X[i] - cx) * (X[i] - cx) + (Y[i] - cy) * (Y[i] - cy) + (Z[i] - cz) * (Z[i] - cz));
            R = Math.Sqrt(R) * 1.001;
            if (R <= 0) return new bool[nt];

            var vis = new bool[nt];
            double golden = Math.PI * (3 - Math.Sqrt(5));
            double s = res / (2 * R);
            float eps = (float)(3 / s);                    // три пикселя по глубине
            float epsCl = (float)Math.Max(3 / s, keepGap); // допуск относительно "заклеенной" карты
            int r2 = holeClose > 0 ? (int)Math.Ceiling(holeClose / 2 * s) : 0;
            float tol2 = (float)holeTol;
            int r = (int)Math.Ceiling(closeSize / 2 * s);
            // направления независимы - считаем параллельно, у каждого потока свои буферы
            // (в vis только пишется true, гонка безопасна)
            System.Threading.Tasks.Parallel.For(0, dirs, () => new Buf(nv, res), (di, state, B) =>
            {
                float[] zb = B.zb, cl = B.cl, cl2 = B.cl2, sx = B.sx, sy = B.sy, sz = B.sz;
                int[] ib = B.ib;
                List<int> small = B.small;
                float[][] line = B.line;
                double yy = 1 - 2 * (di + 0.5) / dirs, rr = Math.Sqrt(1 - yy * yy), phi = golden * di;
                var d = new Vec(rr * Math.Cos(phi), yy, rr * Math.Sin(phi));   // зритель в +d, смотрит в -d
                Vec u = Math.Abs(d.x) < 0.9 ? Vec.cross(d, new Vec(1, 0, 0)).norm() : Vec.cross(d, new Vec(0, 1, 0)).norm();
                Vec w = Vec.cross(d, u);   // u x w = d: лицевые треугольники в проекции CCW
                for (int i = 0; i < nv; i++)
                {
                    double dx = X[i] - cx, dy = Y[i] - cy, dz = Z[i] - cz;
                    sx[i] = (float)((dx * u.x + dy * u.y + dz * u.z) * s + res / 2.0);
                    sy[i] = (float)((dx * w.x + dy * w.y + dz * w.z) * s + res / 2.0);
                    sz[i] = (float)(dx * d.x + dy * d.y + dz * d.z);
                }
                for (int k = 0; k < zb.Length; k++) { zb[k] = float.NegativeInfinity; ib[k] = -1; }
                small.Clear();
                for (int t = 0; t < nt; t++)
                {
                    int a = tri[3 * t], b = tri[3 * t + 1], c = tri[3 * t + 2];
                    float ax = sx[a], ay = sy[a], bx = sx[b], by = sy[b], qx = sx[c], qy = sy[c];
                    float area = (bx - ax) * (qy - ay) - (qx - ax) * (by - ay);
                    if (area <= 0)
                    {
                        if (twoSided == null || !twoSided[t] || area == 0) continue;   // обратная сторона
                        int tmp = b; b = c; c = tmp;                                  // двусторонний: разворачиваем
                        bx = sx[b]; by = sy[b]; qx = sx[c]; qy = sy[c];
                        area = -area;
                    }
                    int x0 = Math.Max(0, (int)Math.Ceiling(Math.Min(ax, Math.Min(bx, qx)) - 0.5f));
                    int x1 = Math.Min(res - 1, (int)Math.Floor(Math.Max(ax, Math.Max(bx, qx)) - 0.5f));
                    int y0 = Math.Max(0, (int)Math.Ceiling(Math.Min(ay, Math.Min(by, qy)) - 0.5f));
                    int y1 = Math.Min(res - 1, (int)Math.Floor(Math.Max(ay, Math.Max(by, qy)) - 0.5f));
                    bool covered = false;
                    float az = sz[a], bz = sz[b], qz = sz[c];
                    for (int py = y0; py <= y1; py++)
                    {
                        float fy = py + 0.5f;
                        for (int px = x0; px <= x1; px++)
                        {
                            float fx = px + 0.5f;
                            float e0 = (qx - bx) * (fy - by) - (qy - by) * (fx - bx);
                            float e1 = (ax - qx) * (fy - qy) - (ay - qy) * (fx - qx);
                            float e2 = (bx - ax) * (fy - ay) - (by - ay) * (fx - ax);
                            if (e0 < 0 || e1 < 0 || e2 < 0) continue;
                            covered = true;
                            float z = (e0 * az + e1 * bz + e2 * qz) / area;
                            int k = py * res + px;
                            if (z > zb[k]) { zb[k] = z; ib[k] = t; }
                        }
                    }
                    if (!covered) small.Add(t);
                }
                // "закрытие" карты глубины: отверстия и щели меньше closeSize заполняются глубиной окружающей
                // поверхности, поэтому всё, что видно только сквозь перфорацию, считается закрытым
                Array.Copy(zb, cl, zb.Length);
                if (r > 0) { morph(cl, res, r, true, line); morph(cl, res, r, false, line); }
                if (r2 > 0) { Array.Copy(zb, cl2, zb.Length); morph(cl2, res, r2, true, line); morph(cl2, res, r2, false, line); }
                for (int k = 0; k < ib.Length; k++)
                {
                    if (ib[k] < 0 || zb[k] < cl[k] - epsCl) continue;
                    bool direct = r2 == 0 || cl2[k] - zb[k] <= tol2;   // не сквозь заклеенный проём
                    if (directOut != null) { vis[ib[k]] = true; if (direct) directOut[ib[k]] = true; }
                    else if (direct) vis[ib[k]] = true;
                }
                // треугольники мельче пикселя: сравниваем глубину центра с буфером
                foreach (int t in small)
                {
                    if (vis[t]) continue;
                    int a = tri[3 * t], b = tri[3 * t + 1], c = tri[3 * t + 2];
                    int px = (int)((sx[a] + sx[b] + sx[c]) / 3), py = (int)((sy[a] + sy[b] + sy[c]) / 3);
                    if (px < 0 || py < 0 || px >= res || py >= res) continue;
                    float z = (sz[a] + sz[b] + sz[c]) / 3;
                    int pk = py * res + px;
                    if (z < zb[pk] - eps || z < cl[pk] - epsCl) continue;
                    bool direct = r2 == 0 || cl2[pk] - z <= tol2;
                    if (directOut != null) { vis[t] = true; if (direct) directOut[t] = true; }
                    else if (direct) vis[t] = true;
                }
                return B;
            }, B => { });

            return vis;
        }

        // closeSize - размер (м) отверстий и щелей, которые считаются закрытыми
        public static long cull(GlbMesh m, int dirs, int res, double closeSize, HashSet<int> twoSidedMats = null)
        {
            var keys = m.prims.Where(kv => kv.Value.idx.Count > 0).Select(kv => kv.Key).ToList();
            var prims = keys.Select(k => m.prims[k]).ToList();
            int nv = 0, nt = 0;
            foreach (var p in prims) { nv += p.vcount; nt += p.idx.Count / 3; }
            if (nt == 0 || dirs <= 0 || res <= 0) return 0;
            var X = new float[nv]; var Y = new float[nv]; var Z = new float[nv];
            var tri = new int[nt * 3];
            var two = new bool[nt];
            int vo = 0, to = 0;
            for (int pi = 0; pi < prims.Count; pi++)
            {
                var p = prims[pi];
                bool ts = twoSidedMats != null && twoSidedMats.Contains(keys[pi]);
                for (int i = 0; i < p.vcount; i++) { X[vo + i] = p.pos[3 * i]; Y[vo + i] = p.pos[3 * i + 1]; Z[vo + i] = p.pos[3 * i + 2]; }
                for (int i = 0; i < p.idx.Count; i++) { tri[to + i] = vo + p.idx[i]; if (i % 3 == 0) two[(to + i) / 3] = ts; }
                to += p.idx.Count;
                vo += p.vcount;
            }
            var vis = visible(X, Y, Z, tri, dirs, res, closeSize, 0, 0, 0, two);

            // соседи видимых треугольников (по совпадающим позициям) тоже остаются
            var mark = new HashSet<Key3>();
            Func<int, Key3> key = i => new Key3((long)Math.Round(X[i] * 1e6), (long)Math.Round(Y[i] * 1e6), (long)Math.Round(Z[i] * 1e6));
            for (int t = 0; t < nt; t++)
                if (vis[t]) for (int k = 0; k < 3; k++) mark.Add(key(tri[3 * t + k]));
            long removed = 0;
            to = 0;
            foreach (var p in prims)
            {
                var map = new Dictionary<int, int>();
                var np = new GlbPrim();
                for (int i = 0; i < p.idx.Count; i += 3, to++)
                {
                    bool keep = vis[to];
                    for (int k = 0; k < 3 && !keep; k++) keep = mark.Contains(key(tri[3 * to + k]));
                    if (!keep) { removed++; continue; }
                    for (int k = 0; k < 3; k++)
                    {
                        int vi = p.idx[i + k], li;
                        if (!map.TryGetValue(vi, out li))
                        {
                            li = np.vcount;
                            map[vi] = li;
                            for (int q = 0; q < 3; q++) { np.pos.Add(p.pos[3 * vi + q]); np.nrm.Add(p.nrm[3 * vi + q]); }
                        }
                        np.idx.Add(li);
                    }
                }
                p.pos = np.pos; p.nrm = np.nrm; p.idx = np.idx;
            }
            return removed;
        }
    }

    public class GlbPrim
    {
        public List<float> pos = new List<float>(), nrm = new List<float>();
        public List<int> idx = new List<int>();
        public int vcount { get { return pos.Count / 3; } }
    }

    // геометрия тела, разбитая по материалам (ключ - индекс материала glTF)
    public class GlbMesh
    {
        public Dictionary<int, GlbPrim> prims = new Dictionary<int, GlbPrim>();
        public GlbPrim get(int mat)
        {
            GlbPrim p;
            if (!prims.TryGetValue(mat, out p)) prims[mat] = p = new GlbPrim();
            return p;
        }
    }

    public class GlbMat
    {
        public string name = "default";
        public double r = 0.8, g = 0.8, b = 0.8, a = 1, metallic = 0, roughness = 0.5;
        public bool doubleSided;

        static readonly string[] metalWords = { "металл", "сталь", "метал", "оцинк", "алюмин", "хром", "латун", "медь", "нерж",
            "metal", "steel", "alumin", "chrome", "brass", "copper", "zinc", "galvan", "iron" };

        public static GlbMat from(Asset a)
        {
            var m = new GlbMat();
            if (a == null) return m;
            m.name = a.DisplayName;
            Color c = null;
            foreach (string cn in colorNames)
            {
                try { c = ((ColorAssetValue)a[cn]).Value; } catch { }
                if (c != null) break;
            }
            if (c != null) { m.r = lin(c.Red / 255.0); m.g = lin(c.Green / 255.0); m.b = lin(c.Blue / 255.0); }

            string metalType = choice(a, "metal_type");
            if (metalType != null)
            {
                // Metal: цвет определяется типом металла, metal_color - только для анодирования
                string fin = choice(a, "metal_finish") ?? "";
                m.roughness = fin.Contains("semi") ? 0.3 : fin.Contains("polish") ? 0.15 : fin.Contains("satin") ? 0.4 :
                              fin.Contains("brush") ? 0.45 : fin.Contains("hammer") ? 0.5 : 0.35;
                if (metalType.Contains("galvanized_alu"))
                    m.metallic = 0.3;       // анодированный алюминий: окрашенное покрытие
                else
                {
                    m.metallic = 1;
                    double[] mc = metalColor(metalType);
                    if (mc != null) { m.r = mc[0]; m.g = mc[1]; m.b = mc[2]; }
                }
                return m;
            }

            string pv = choice(a, "plasticvinyl_application") ?? choice(a, "wallpaint_finish") ?? choice(a, "metallicpaint_finish");
            if (pv != null)
            {
                // пластик / краски: шероховатость по отделке
                m.metallic = 0;
                m.roughness = pv.Contains("matte") || pv.Contains("flat") ? 0.8 : pv.Contains("eggshell") ? 0.65 :
                              pv.Contains("platinum") ? 0.55 : pv.Contains("pearl") ? 0.45 : pv.Contains("semigloss") ? 0.35 :
                              pv.Contains("gloss") || pv.Contains("polish") || pv.Contains("smooth") ? 0.2 : 0.5;
                if ((choice(a, "plasticvinyl_type") ?? "").Contains("transparent")) m.a = 0.5;
                return m;
            }

            if (choice(a, "glazing_transmittance_color") != null || has(a, "solidglass_transmittance_custom_color"))
            {
                m.metallic = 0; m.roughness = 0.05; m.a = 0.35;
                return m;
            }

            // Generic и прочее
            try { m.a = 1 - ((FloatAssetValue)a["generic_transparency"]).Value; } catch { }
            bool metal = false;
            try { metal = ((BooleanAssetValue)a["generic_is_metal"]).Value; } catch { }
            string ln = m.name.ToLower();
            if (!metal && !has(a, "generic_is_metal")) metal = metalWords.Any(w => ln.Contains(w));
            double gloss = -1;
            try { gloss = ((FloatAssetValue)a["generic_glossiness"]).Value; } catch { }
            m.metallic = metal ? 1 : 0;
            m.roughness = gloss >= 0 ? Math.Max(0.05, Math.Min(1, 1 - gloss)) : (metal ? 0.35 : 0.5);
            return m;
        }

        static string choice(Asset a, string name)
        {
            try { return ((ChoiceAssetValue)a[name]).Value; } catch { return null; }
        }

        static bool has(Asset a, string name)
        {
            try { return a[name] != null; } catch { return false; }
        }

        // линейные базовые цвета металлов (PBR, F0)
        static double[] metalColor(string t)
        {
            if (t.Contains("copper")) return new[] { 0.955, 0.638, 0.538 };
            if (t.Contains("bronze")) return new[] { 0.80, 0.55, 0.35 };
            if (t.Contains("brass")) return new[] { 0.910, 0.778, 0.423 };
            if (t.Contains("chrome")) return new[] { 0.549, 0.556, 0.554 };
            if (t.Contains("stainless")) return new[] { 0.67, 0.64, 0.60 };
            if (t.Contains("galvanized")) return new[] { 0.62, 0.62, 0.62 };
            if (t.Contains("zinc")) return new[] { 0.664, 0.824, 0.850 };
            if (t.Contains("alu")) return new[] { 0.913, 0.922, 0.924 };
            if (t.Contains("nickel")) return new[] { 0.660, 0.609, 0.526 };
            if (t.Contains("gold")) return new[] { 1.0, 0.766, 0.336 };
            if (t.Contains("silver")) return new[] { 0.972, 0.960, 0.915 };
            if (t.Contains("titan")) return new[] { 0.542, 0.497, 0.449 };
            if (t.Contains("lead")) return new[] { 0.63, 0.63, 0.64 };
            if (t.Contains("steel") || t.Contains("iron")) return new[] { 0.56, 0.57, 0.58 };
            return null;
        }

        // основной цвет в разных схемах Autodesk (Generic, Metal, Metallic Paint, Plastic, Wall Paint...)
        static readonly string[] colorNames = { "generic_diffuse", "metal_color", "metallicpaint_base_color", "plasticvinyl_color",
            "wallpaint_color", "ceramic_color", "masonrycmu_color", "concrete_color", "stone_color", "hardwood_tint_color",
            "glazing_transmittance_map", "solidglass_transmittance_custom_color", "mirror_tintcolor", "water_tint_color" };

        public static string valueText(AssetValue v)
        {
            try
            {
                if (v is ColorAssetValue)
                {
                    Color c = ((ColorAssetValue)v).Value;
                    return "Color(" + c.Red + "," + c.Green + "," + c.Blue + ")" + (((ColorAssetValue)v).HasConnectedTexture ? " +texture" : "");
                }
                if (v is FloatAssetValue) return ((FloatAssetValue)v).Value.ToString(CultureInfo.InvariantCulture);
                if (v is IntegerAssetValue) return "int " + ((IntegerAssetValue)v).Value;
                if (v is BooleanAssetValue) return "bool " + ((BooleanAssetValue)v).Value;
                if (v is ChoiceAssetValue) return "choice " + ((ChoiceAssetValue)v).Value;
                if (v is StringAssetValue) return "\"" + ((StringAssetValue)v).Value + "\"";
                return v.ValueType.ToString();
            }
            catch (Exception e) { return "! " + e.Message; }
        }

        static double lin(double c) { return c <= 0.04045 ? c / 12.92 : Math.Pow((c + 0.055) / 1.055, 2.4); }
    }

    // минимальный writer GLB 2.0: один буфер, POSITION/NORMAL/indices, PBR-материалы
    public class GlbBuilder
    {
        static readonly CultureInfo ci = CultureInfo.InvariantCulture;
        MemoryStream bin = new MemoryStream();
        List<string> views = new List<string>(), accessors = new List<string>(), meshes = new List<string>(), materials = new List<string>();
        List<Node> nodes = new List<Node>();
        public long vertices, triangles;
        public int meshCount { get { return meshes.Count; } }

        class Node { public string name; public double[] m; public int mesh = -1; public List<int> children; }

        public int addMaterial(GlbMat m)
        {
            var sb = new StringBuilder();
            sb.Append("{\"name\":").Append(str(m.name));
            sb.Append(",\"pbrMetallicRoughness\":{\"baseColorFactor\":[")
              .Append(num(m.r)).Append(',').Append(num(m.g)).Append(',').Append(num(m.b)).Append(',').Append(num(m.a))
              .Append("],\"metallicFactor\":").Append(num(m.metallic))
              .Append(",\"roughnessFactor\":").Append(num(m.roughness)).Append('}');
            if (m.a < 0.999) sb.Append(",\"alphaMode\":\"BLEND\"");
            if (m.doubleSided) sb.Append(",\"doubleSided\":true");
            sb.Append('}');
            materials.Add(sb.ToString());
            return materials.Count - 1;
        }

        public bool Quantize = true;             // KHR_mesh_quantization: позиции uint16, нормали int8, индексы uint16
        List<double[]> dequant = new List<double[]>();   // на меш: {ox, oy, oz, s} - перенос и масштаб узла; null - без квантизации

        public int addMesh(string name, GlbMesh mesh)
        {
            var prims = new List<string>();
            // общий габарит меша - одно преобразование на меш
            double[] lo = { double.MaxValue, double.MaxValue, double.MaxValue }, hi = { double.MinValue, double.MinValue, double.MinValue };
            foreach (var p in mesh.prims.Values)
                for (int i = 0; i < p.pos.Count; i++) { lo[i % 3] = Math.Min(lo[i % 3], p.pos[i]); hi[i % 3] = Math.Max(hi[i % 3], p.pos[i]); }
            double[] dq = null;
            if (Quantize && lo[0] <= hi[0])
            {
                double ext = Math.Max(hi[0] - lo[0], Math.Max(hi[1] - lo[1], hi[2] - lo[2]));
                dq = new[] { lo[0], lo[1], lo[2], ext > 0 ? ext / 65535 : 1 };
            }
            foreach (var kv in mesh.prims)
            {
                GlbPrim p = kv.Value;
                if (p.idx.Count == 0) continue;
                if (dq == null) { prims.Add(primFloat(p, kv.Key)); continue; }
                // куски не больше 65535 вершин - индексы uint16
                var map = new Dictionary<int, int>();
                var chunk = new List<int>();
                for (int t = 0; t < p.idx.Count; t += 3)
                {
                    if (map.Count > 65535 - 3) { prims.Add(primQuant(p, map, chunk, dq, kv.Key)); map.Clear(); chunk.Clear(); }
                    for (int k = 0; k < 3; k++)
                    {
                        int v = p.idx[t + k], li;
                        if (!map.TryGetValue(v, out li)) { li = map.Count; map[v] = li; }
                        chunk.Add(li);
                    }
                }
                if (chunk.Count > 0) prims.Add(primQuant(p, map, chunk, dq, kv.Key));
                triangles += p.idx.Count / 3;
            }
            if (prims.Count == 0) return -1;
            meshes.Add("{\"name\":" + str(name) + ",\"primitives\":[" + string.Join(",", prims) + "]}");
            dequant.Add(dq);
            return meshes.Count - 1;
        }

        string primQuant(GlbPrim p, Dictionary<int, int> map, List<int> idx, double[] dq, int mat)
        {
            int vc = map.Count;
            var pos = new byte[vc * 8];   // uint16 x3 + выравнивание до 8
            var nrm = new byte[vc * 4];   // int8 x3 + выравнивание до 4
            float[] qmin = { float.MaxValue, float.MaxValue, float.MaxValue }, qmax = { float.MinValue, float.MinValue, float.MinValue };
            foreach (var kv in map)
            {
                int v = kv.Key, o = kv.Value;
                for (int k = 0; k < 3; k++)
                {
                    int q = (int)Math.Round((p.pos[3 * v + k] - dq[k]) / dq[3]);
                    q = Math.Max(0, Math.Min(65535, q));
                    pos[o * 8 + 2 * k] = (byte)(q & 0xff); pos[o * 8 + 2 * k + 1] = (byte)(q >> 8);
                    qmin[k] = Math.Min(qmin[k], q); qmax[k] = Math.Max(qmax[k], q);
                    int nq = (int)Math.Round(Math.Max(-1, Math.Min(1, p.nrm[3 * v + k])) * 127);
                    nrm[o * 4 + k] = (byte)(sbyte)nq;
                }
            }
            var ib = new byte[idx.Count * 2];
            for (int i = 0; i < idx.Count; i++) { ib[2 * i] = (byte)(idx[i] & 0xff); ib[2 * i + 1] = (byte)(idx[i] >> 8); }
            int ap = accessor(view(pos, 34962, 8), 5123, vc, "VEC3", qmin, qmax);
            int an = accessor(view(nrm, 34962, 4), 5120, vc, "VEC3", null, null, true);
            int ai = accessor(view(ib, 34963), 5123, idx.Count, "SCALAR", null, null);
            vertices += vc;
            return string.Format(ci, "{{\"attributes\":{{\"POSITION\":{0},\"NORMAL\":{1}}},\"indices\":{2},\"material\":{3}}}", ap, an, ai, mat);
        }

        string primFloat(GlbPrim p, int mat)
        {
            int vc = p.vcount;
            float[] min = { float.MaxValue, float.MaxValue, float.MaxValue }, max = { float.MinValue, float.MinValue, float.MinValue };
            for (int i = 0; i < vc; i++)
                for (int k = 0; k < 3; k++)
                {
                    min[k] = Math.Min(min[k], p.pos[3 * i + k]);
                    max[k] = Math.Max(max[k], p.pos[3 * i + k]);
                }
            int ap = accessor(view(floats(p.pos), 34962), 5126, vc, "VEC3", min, max);
            int an = accessor(view(floats(p.nrm), 34962), 5126, vc, "VEC3", null, null);
            byte[] ib;
            int ct;
            if (vc < 65536)
            {
                var s = p.idx.Select(i => (ushort)i).ToArray();
                ib = new byte[s.Length * 2]; Buffer.BlockCopy(s, 0, ib, 0, ib.Length); ct = 5123;
            }
            else
            {
                var s = p.idx.Select(i => (uint)i).ToArray();
                ib = new byte[s.Length * 4]; Buffer.BlockCopy(s, 0, ib, 0, ib.Length); ct = 5125;
            }
            int ai = accessor(view(ib, 34963), ct, p.idx.Count, "SCALAR", null, null);
            vertices += vc;
            triangles += p.idx.Count / 3;
            return string.Format(ci, "{{\"attributes\":{{\"POSITION\":{0},\"NORMAL\":{1}}},\"indices\":{2},\"material\":{3}}}", ap, an, ai, mat);
        }

        public int addNode(string name, double[] m, int mesh, List<int> children)
        {
            nodes.Add(new Node { name = name, m = m, mesh = mesh, children = children });
            return nodes.Count - 1;
        }

        public bool nodeHasOnlyMesh(int n) { return n == nodes.Count - 1 && nodes[n].m == null && (nodes[n].children == null || nodes[n].children.Count == 0); }
        public int nodeMesh(int n) { return nodes[n].mesh; }
        public void removeLastNode() { nodes.RemoveAt(nodes.Count - 1); }

        public void save(string fn, int root)
        {
            var sb = new StringBuilder();
            sb.Append("{\"asset\":{\"version\":\"2.0\",\"generator\":\"Macros Inventor GLB\"},\"scene\":0,\"scenes\":[{\"nodes\":[").Append(root).Append("]}]");
            // узлы с квантизованным мешем: меш уходит в дочерний узел с переносом и масштабом (деквантизация)
            var extra = new List<string>();
            var nj = new List<string>();
            bool quant = false;
            foreach (var n in nodes)
            {
                double[] dq = n.mesh >= 0 ? dequant[n.mesh] : null;
                if (dq == null) { nj.Add(nodeJson(n, -1)); continue; }
                quant = true;
                int child = nodes.Count + extra.Count;
                extra.Add(string.Format(ci, "{{\"mesh\":{0},\"translation\":[{1},{2},{3}],\"scale\":[{4},{4},{4}]}}",
                    n.mesh, num(dq[0]), num(dq[1]), num(dq[2]), num(dq[3])));
                nj.Add(nodeJson(n, child));
            }
            nj.AddRange(extra);
            if (quant) sb.Append(",\"extensionsUsed\":[\"KHR_mesh_quantization\"],\"extensionsRequired\":[\"KHR_mesh_quantization\"]");
            sb.Append(",\"nodes\":[").Append(string.Join(",", nj)).Append(']');
            if (meshes.Count > 0) sb.Append(",\"meshes\":[").Append(string.Join(",", meshes)).Append(']');
            if (materials.Count > 0) sb.Append(",\"materials\":[").Append(string.Join(",", materials)).Append(']');
            if (accessors.Count > 0)
            {
                sb.Append(",\"accessors\":[").Append(string.Join(",", accessors)).Append(']');
                sb.Append(",\"bufferViews\":[").Append(string.Join(",", views)).Append(']');
                sb.Append(",\"buffers\":[{\"byteLength\":").Append(bin.Length).Append("}]");
            }
            sb.Append('}');

            byte[] json = Encoding.UTF8.GetBytes(sb.ToString());
            int jsonLen = (json.Length + 3) & ~3;
            int binLen = (int)((bin.Length + 3) & ~3);
            using (var w = new BinaryWriter(System.IO.File.Create(fn)))
            {
                w.Write(0x46546C67u);
                w.Write(2u);
                w.Write((uint)(12 + 8 + jsonLen + (bin.Length > 0 ? 8 + binLen : 0)));
                w.Write((uint)jsonLen); w.Write(0x4E4F534Au);
                w.Write(json);
                for (int i = json.Length; i < jsonLen; i++) w.Write((byte)0x20);
                if (bin.Length > 0)
                {
                    w.Write((uint)binLen); w.Write(0x004E4942u);
                    w.Write(bin.GetBuffer(), 0, (int)bin.Length);
                    for (long i = bin.Length; i < binLen; i++) w.Write((byte)0);
                }
            }
        }

        // meshChild >= 0: меш не на этом узле, а в дочернем узле деквантизации
        string nodeJson(Node n, int meshChild)
        {
            var sb = new StringBuilder("{\"name\":" + str(n.name));
            if (n.mesh >= 0 && meshChild < 0) sb.Append(",\"mesh\":").Append(n.mesh);
            var ch = new List<int>();
            if (n.children != null) ch.AddRange(n.children);
            if (meshChild >= 0) ch.Add(meshChild);
            if (ch.Count > 0) sb.Append(",\"children\":[").Append(string.Join(",", ch)).Append(']');
            if (n.m != null && !isIdentity(n.m))
            {
                // row-major (перенос в см) -> column-major (м)
                var c = new List<string>();
                for (int col = 0; col < 4; col++)
                    for (int row = 0; row < 4; row++)
                        c.Add(num(n.m[row * 4 + col] * (col == 3 && row < 3 ? 0.01 : 1)));
                sb.Append(",\"matrix\":[").Append(string.Join(",", c)).Append(']');
            }
            return sb.Append('}').ToString();
        }

        static bool isIdentity(double[] m)
        {
            for (int i = 0; i < 16; i++) if (Math.Abs(m[i] - (i % 5 == 0 ? 1 : 0)) > 1e-12) return false;
            return true;
        }

        int view(byte[] data, int target, int stride = 0)
        {
            while (bin.Length % 4 != 0) bin.WriteByte(0);
            long off = bin.Length;
            bin.Write(data, 0, data.Length);
            views.Add(string.Format(ci, "{{\"buffer\":0,\"byteOffset\":{0},\"byteLength\":{1},\"target\":{2}{3}}}", off, data.Length, target,
                stride > 0 ? ",\"byteStride\":" + stride : ""));
            return views.Count - 1;
        }

        int accessor(int view, int type, int count, string kind, float[] min, float[] max, bool normalized = false)
        {
            var s = string.Format(ci, "{{\"bufferView\":{0},\"componentType\":{1},\"count\":{2},\"type\":\"{3}\"", view, type, count, kind);
            if (normalized) s += ",\"normalized\":true";
            if (min != null)
                s += ",\"min\":[" + string.Join(",", min.Select(v => num(v))) + "],\"max\":[" + string.Join(",", max.Select(v => num(v))) + "]";
            accessors.Add(s + "}");
            return accessors.Count - 1;
        }

        static byte[] floats(List<float> l)
        {
            var a = l.ToArray();
            var b = new byte[a.Length * 4];
            Buffer.BlockCopy(a, 0, b, 0, b.Length);
            return b;
        }

        static string num(double v) { return v.ToString("R", ci); }
        static string num(float v) { return v.ToString("R", ci); }

        static string str(string s)
        {
            var sb = new StringBuilder("\"");
            foreach (char c in s ?? "")
            {
                if (c == '"') sb.Append("\\\"");
                else if (c == '\\') sb.Append("\\\\");
                else if (c < 0x20) sb.Append("\\u").Append(((int)c).ToString("x4"));
                else sb.Append(c);
            }
            return sb.Append('"').ToString();
        }
    }
}
