// Проверка загрузки Macros.dll: запустить в папке аддина (рядом с Macros.dll). Отчёт - CheckLoad.txt рядом с exe.
using System; using System.IO; using System.Linq; using System.Reflection; using System.Runtime.InteropServices;
class CheckLoad
{
    [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    static extern IntPtr CreateFile(string name, uint access, uint share, IntPtr sec, uint disp, uint flags, IntPtr tmpl);
    [DllImport("kernel32.dll")] static extern bool CloseHandle(IntPtr h);

    // отметка Windows "получен из интернета" (поток Zone.Identifier) - из-за неё .NET молча не грузит сборку
    static bool Blocked(string f)
    {
        IntPtr h = CreateFile(f + ":Zone.Identifier", 0x80000000, 1, IntPtr.Zero, 3, 0, IntPtr.Zero);
        if (h == new IntPtr(-1)) return false;
        CloseHandle(h);
        return true;
    }

    // interop Inventor - из папки Inventor (внутри Inventor он уже загружен, а отдельной программе его надо найти)
    static string InventorInterop()
    {
        foreach (string pf in new[] { Environment.GetEnvironmentVariable("ProgramW6432"), Environment.GetEnvironmentVariable("ProgramFiles") })
        {
            if (string.IsNullOrEmpty(pf)) continue;
            string root = Path.Combine(pf, "Autodesk");
            if (!Directory.Exists(root)) continue;
            foreach (string inv in Directory.GetDirectories(root, "Inventor 20*").OrderByDescending(x => x))
                foreach (string sub in new[] { @"Bin\Public Assemblies", "Bin" })
                {
                    string f = Path.Combine(Path.Combine(inv, sub), "Autodesk.Inventor.Interop.dll");
                    if (File.Exists(f)) return f;
                }
        }
        return null;
    }

    static string dir;
    static StreamWriter log;

    static void W(string s)
    {
        try { Console.WriteLine(s); } catch { }
        try { log.WriteLine(s); log.Flush(); } catch { }
    }

    static int Main(string[] args)
    {
        dir = args.Length > 0 ? args[0] : AppDomain.CurrentDomain.BaseDirectory;
        string rep = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "CheckLoad.txt");
        try { log = new StreamWriter(rep, false, new System.Text.UTF8Encoding(true)); }
        catch { rep = Path.Combine(Path.GetTempPath(), "CheckLoad.txt"); log = new StreamWriter(rep, false, new System.Text.UTF8Encoding(true)); }
        AppDomain.CurrentDomain.UnhandledException += (s, e) => { W("НЕОБРАБОТАННАЯ ОШИБКА: " + e.ExceptionObject); };
        int bad = 0;
        try { bad = Run(); }
        catch (Exception ex) { bad++; W("ОШИБКА проверки: " + ex); }
        W(bad == 0 ? "\nЗагрузка в порядке. Если Inventor всё равно не грузит - разблокируйте файлы (Свойства -> Разблокировать)."
                   : "\nНайдено ошибок: " + bad);
        W("Отчёт: " + rep);
        log.Close();
        try { Console.WriteLine("\nНажмите Enter..."); Console.ReadLine(); } catch { }
        return bad;
    }

    static int Run()
    {
        string dll = Path.Combine(dir, "Macros.dll");
        W("Дата: " + DateTime.Now + "   Папка: " + dir + "   64-бит процесс: " + Environment.Is64BitProcess + "   .NET " + Environment.Version);
        if (!File.Exists(dll)) { W("Нет Macros.dll"); return 1; }
        int blocked = 0;
        foreach (var f in Directory.GetFiles(dir, "*.dll"))
        {
            bool b = Blocked(f);
            if (b) blocked++;
            W("   файл " + Path.GetFileName(f) + "  " + File.GetLastWriteTime(f) + "  " + new FileInfo(f).Length + " байт" + (b ? "   ЗАБЛОКИРОВАН Windows" : ""));
        }
        if (blocked > 0) W("ОШИБКА: заблокировано файлов: " + blocked + " - Inventor их не загрузит. Разблокировать: Свойства -> Разблокировать, или PowerShell: Get-ChildItem \"" + dir + "\" | Unblock-File");
        string interop = InventorInterop();
        W("Interop Inventor: " + (interop ?? "не найден в Program Files\\Autodesk"));
        AppDomain.CurrentDomain.AssemblyResolve += (s, e) =>
        {
            try
            {
                string n = new AssemblyName(e.Name).Name;
                if (n == "Autodesk.Inventor.Interop" && interop != null) return Assembly.LoadFrom(interop);
                string f = Path.Combine(dir, n + ".dll");
                return File.Exists(f) ? Assembly.LoadFrom(f) : null;
            }
            catch (Exception ex) { W("ОШИБКА загрузки " + e.Name + ": " + ex.Message); return null; }
        };
        int bad = 0;
        Assembly asm;
        try { asm = Assembly.LoadFrom(dll); W("OK  Macros.dll " + asm.GetName().Version); }
        catch (Exception ex) { W("ОШИБКА загрузки Macros.dll: " + ex); return 1; }
        foreach (var r in asm.GetReferencedAssemblies().OrderBy(x => x.Name))
        {
            try { var a = Assembly.Load(r); W("OK  " + r.Name + " " + r.Version + "  <- " + a.Location); }
            catch (Exception ex) { bad++; W("ОШИБКА  " + r.Name + " " + r.Version + ": " + ex.Message); }
        }
        try { asm.GetTypes(); W("OK  все типы Macros.dll загружаются"); }
        catch (ReflectionTypeLoadException ex)
        {
            bad++;
            foreach (var le in ex.LoaderExceptions.Where(x => x != null).Select(x => x.Message).Distinct().Take(15))
                W("ОШИБКА типа: " + le);
        }
        catch (Exception ex) { bad++; W("ОШИБКА типов: " + ex); }
        try
        {
            var t = asm.GetType("Macros.StandardAddInServer");
            W(t == null ? "ОШИБКА: нет класса Macros.StandardAddInServer" : "OK  класс аддина найден, GUID " + t.GUID);
            if (t != null) { Activator.CreateInstance(t); W("OK  объект аддина создаётся"); }
        }
        catch (Exception ex) { bad++; W("ОШИБКА создания аддина: " + (ex.InnerException ?? ex)); }
        return bad + blocked;
    }
}
