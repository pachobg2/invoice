using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text.Json;
using System.Text.Json.Nodes;
using Microsoft.Web.WebView2.Core;
using Microsoft.Web.WebView2.WinForms;

namespace Faktura;

static class Program
{
    [DllImport("user32.dll")] static extern bool SetForegroundWindow(IntPtr h);
    [DllImport("user32.dll")] static extern IntPtr FindWindow(string cls, string title);
    [DllImport("user32.dll")] static extern bool ShowWindow(IntPtr h, int cmd);

    [STAThread]
    static void Main(string[] args)
    {
        // started by the updater: let the old instance finish closing first
        int wi = Array.IndexOf(args, "--wait-pid");
        if (wi >= 0 && wi + 1 < args.Length && int.TryParse(args[wi + 1], out var wp)) { try { Process.GetProcessById(wp).WaitForExit(20000); } catch { } }
        var where = (Path.GetDirectoryName(Environment.ProcessPath) ?? "").ToLowerInvariant();
        var id = Convert.ToHexString(System.Security.Cryptography.SHA1.HashData(System.Text.Encoding.UTF8.GetBytes(where)))[..12];
        using var mutex = new Mutex(true, @"Local\Faktura.SingleInstance." + id, out bool first);
        if (!first)
        {
            var h = FindWindow(null, "Фактури");
            if (h != IntPtr.Zero) { ShowWindow(h, 9); SetForegroundWindow(h); }
            return;
        }
        Updater.CleanUp();
        ApplicationConfiguration.Initialize();
        Application.Run(new MainForm());
    }
}

class Config
{
    public string BackupDir { get; set; }
    public string PdfDir { get; set; }
    public int X { get; set; } = -1;
    public int Y { get; set; } = -1;
    public int W { get; set; }
    public int H { get; set; }
    public bool Max { get; set; }
}

class MainForm : Form
{
    const int KeepBackups = 30;
    static readonly string ExeDir = Path.GetDirectoryName(Environment.ProcessPath)!;
    static readonly string TestRoot = Environment.GetEnvironmentVariable("FAKTURA_TEST_ROOT"); // tests only: keeps real data untouched
    static readonly string PortableDir = Path.Combine(ExeDir, "invoice-data");
    static readonly string LegacyDataDir = Path.Combine(ExeDir, "Faktura-data");   // name used before 2.5.2 - renamed automatically
    static readonly string AppDataDir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "Faktura");
    string DataDir, DataFile, ConfigFile; // ONE folder holds everything; see ResolveDataDir()
    static readonly string Docs = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);

    readonly WebView2 web = new WebView2 { Dock = DockStyle.Fill, DefaultBackgroundColor = Color.FromArgb(243, 244, 246) };
    Config cfg = new();
    DateTime lastBak = DateTime.MinValue;
    bool closing;
    bool portableMode;   // invoice-data\portable.flag exists: run from a flash drive, keep the browser profile in %TEMP% and delete it on exit
    string wvDir;

    string BackupDir => Path.Combine(DataDir, "backups");

    public MainForm()
    {
        Text = "Фактури";
        ResolveDataDir();
        MigrateOldData();
        portableMode = File.Exists(Path.Combine(DataDir, "portable.flag"));
        try { cfg = JsonSerializer.Deserialize<Config>(File.ReadAllText(ConfigFile)) ?? new Config(); } catch { cfg = new Config(); }
        try { Icon = Icon.ExtractAssociatedIcon(Environment.ProcessPath!); } catch { }
        var sc = DeviceDpi / 96f;
        MinimumSize = new Size((int)(900 * sc), (int)(600 * sc));
        var area = Screen.PrimaryScreen!.WorkingArea;
        var w = Math.Min(cfg.W > 0 ? Math.Max(cfg.W, MinimumSize.Width) : (int)(1320 * sc), area.Width); var h = Math.Min(cfg.H > 0 ? Math.Max(cfg.H, MinimumSize.Height) : (int)(900 * sc), area.Height);
        StartPosition = FormStartPosition.Manual;
        var onScreen = cfg.X >= 0 && Screen.AllScreens.Any(s => s.WorkingArea.Contains(new Point(cfg.X + 60, cfg.Y + 20)));
        Bounds = onScreen ? new Rectangle(cfg.X, cfg.Y, w, h) : new Rectangle(area.Left + (area.Width - w) / 2, area.Top + (area.Height - h) / 2, w, h);
        if (cfg.Max) WindowState = FormWindowState.Maximized;
        Controls.Add(web);
        AutoBackupDaily();
    }

    // EVERYTHING lives in one folder next to the exe, "invoice-data":
    //   data.json (all documents, clients, settings), data.json.bak, backups\ (automatic archives), config.json (window position).
    // Copy the exe together with that folder to move the whole app. If the exe's folder is read-only the same layout is used in %AppData%\Faktura.
    void ResolveDataDir()
    {
        if (TestRoot != null && Environment.GetEnvironmentVariable("FAKTURA_TEST_PORTABLE") != "1") DataDir = Path.Combine(TestRoot, "data");
        else
        {
            try
            {
                var pd = PortableDir;
                if (!Directory.Exists(PortableDir) && Directory.Exists(LegacyDataDir))   // data folder of the older builds: rename it, keep working in it if that fails
                {
                    try { Directory.Move(LegacyDataDir, PortableDir); } catch { pd = LegacyDataDir; }
                }
                Directory.CreateDirectory(pd);
                var probe = Path.Combine(pd, ".write-test"); File.WriteAllText(probe, "x"); File.Delete(probe);
                DataDir = pd;
            }
            catch { DataDir = AppDataDir; }
        }
        Directory.CreateDirectory(DataDir); Directory.CreateDirectory(Path.Combine(DataDir, "backups"));
        DataFile = Path.Combine(DataDir, "data.json"); ConfigFile = Path.Combine(DataDir, "config.json");
    }

    // Gathers data left behind by earlier builds into the one folder (nothing is deleted; originals are moved into backups\ when they are duplicates).
    void MigrateOldData()
    {
        var backups = Path.Combine(DataDir, "backups");
        // under test the "old" places are replaced by test folders so real files are never touched
        var oldApp = TestRoot != null ? Path.Combine(TestRoot, "old_appdata") : AppDataDir;
        var oldDocs = TestRoot != null ? Path.Combine(TestRoot, "old_docs") : Path.Combine(Docs, "Фактури-backup");
        static string Full(string p) => Path.GetFullPath(p).TrimEnd('\\');
        void Move(string from, string to, bool keepBoth = true)
        {
            try
            {
                if (!File.Exists(from)) return;
                Directory.CreateDirectory(Path.GetDirectoryName(to)!);
                if (keepBoth && File.Exists(to)) // never overwrite an older archive: add a time stamp instead
                    to = Path.Combine(Path.GetDirectoryName(to)!, Path.GetFileNameWithoutExtension(to) + "_" + DateTime.Now.ToString("yyyyMMdd_HHmmss") + Path.GetExtension(to));
                File.Move(from, to, !keepBoth);
            }
            catch { }
        }
        string oldBackupDir = null;
        try { oldBackupDir = JsonSerializer.Deserialize<Config>(File.ReadAllText(Path.Combine(oldApp, "config.json")))?.BackupDir; } catch { }

        // 1) data file of the previous build (it sat directly next to the exe)
        var beside = Path.Combine(ExeDir, "Faktura-data.json");
        if (!File.Exists(DataFile)) { Move(beside, DataFile); Move(beside + ".bak", DataFile + ".bak"); }
        else { Move(beside, Path.Combine(backups, "old_beside_exe.json")); Move(beside + ".bak", Path.Combine(backups, "old_beside_exe.bak.json")); }

        // 2) data + settings of the very first build (%AppData%\Faktura)
        if (Full(oldApp) != Full(DataDir))
        {
            var oldData = Path.Combine(oldApp, "data.json");
            if (File.Exists(oldData)) { try { if (!File.Exists(DataFile)) File.Copy(oldData, DataFile); } catch { } Move(oldData, Path.Combine(backups, "old_appdata.json")); }
            Move(Path.Combine(oldApp, "data.json.bak"), Path.Combine(backups, "old_appdata.bak.json"));
            if (!File.Exists(ConfigFile)) Move(Path.Combine(oldApp, "config.json"), ConfigFile);
        }

        // 3) automatic archives that were kept elsewhere (Documents\Фактури-backup or a folder chosen earlier)
        foreach (var dir in new[] { oldDocs, oldBackupDir }.Where(d => !string.IsNullOrWhiteSpace(d)).Select(d => d!).Distinct())
        {
            try
            {
                if (!Directory.Exists(dir) || Full(dir) == Full(backups)) continue;
                foreach (var f in new DirectoryInfo(dir).GetFiles("faktura_*.json"))
                { var to = Path.Combine(backups, f.Name); if (!File.Exists(to)) f.MoveTo(to); }
                try { Directory.Delete(dir); } catch { } // only succeeds when it is empty now
            }
            catch { }
        }
    }

    protected override async void OnLoad(EventArgs e)
    {
        base.OnLoad(e);
        try
        {
            wvDir = portableMode ? Path.Combine(Path.GetTempPath(), "Faktura-wv2-" + Guid.NewGuid().ToString("N"))
                  : TestRoot != null ? Path.Combine(TestRoot, "wv2") : Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Faktura", "WebView2");
            if (portableMode) CleanStaleProfiles();
            var env = await CoreWebView2Environment.CreateAsync(null, wvDir);
            await web.EnsureCoreWebView2Async(env);
        }
        catch (Exception ex)
        {
            MessageBox.Show("Не може да се стартира Microsoft Edge WebView2 Runtime.\r\n\r\nИнсталирайте го от https://go.microsoft.com/fwlink/p/?LinkId=2124703\r\n\r\n" + ex.Message, "Фактури", MessageBoxButtons.OK, MessageBoxIcon.Error);
            closing = true; Close(); return;
        }
        var core = web.CoreWebView2;
        // the whole web app is embedded in this exe and served from memory at https://faktura.app/
        core.AddWebResourceRequestedFilter("https://faktura.app/*", CoreWebView2WebResourceContext.All);
        core.WebResourceRequested += OnWebResource;
        core.Settings.IsStatusBarEnabled = false;
        core.Settings.IsPasswordAutosaveEnabled = false;
        core.Settings.IsGeneralAutofillEnabled = false;
        core.Settings.AreDevToolsEnabled = true;
        core.WebMessageReceived += OnMessage;
        core.NewWindowRequested += (s, a) => { a.Handled = true; OpenExternal(a.Uri); };
        core.NavigationStarting += (s, a) => { if (!a.Uri.StartsWith("https://faktura.app/")) { a.Cancel = true; OpenExternal(a.Uri); } };
        core.Navigate("https://faktura.app/index.html");
    }

    const string ReadmeText =
"INVOICE - PORTABLE VERSION / ФАКТУРИ - ПРЕНОСИМА ВЕРСИЯ\r\n" +
"=======================================================\r\n\r\n" +
"Стартирайте програмата с пряката връзка „invoice“ в началото на флашката\r\n" +
"(или директно invoice-lite.exe от тази папка).\r\n\r\n" +
"Работи на Windows XP, Vista, 7, 8, 10 и 11 и издава и печата фактури.\r\n" +
"Изисква .NET Framework 2.0 или 3.5 (има го в Windows 7; на Windows XP се инсталира от сайта на Microsoft).\r\n" +
"Файлът invoice-lite.exe.config трябва да стои до програмата.\r\n\r\n" +
"ВАЖНО\r\n" +
"  * Не изтривайте и не преименувайте папка invoice-data - в нея са всички данни.\r\n" +
"  * Не издавайте фактури с номера от основното приложение, докато флашката е на път.\r\n\r\n" +
"КАТО СЕ ВЪРНЕТЕ В ОФИСА\r\n" +
"  Включете флашката, отворете основното приложение и натиснете бутона за синхронизация в горното меню.\r\n" +
"  Приложението само намира данните във флашката (или ви моли да изберете файла invoice-data\\data.json).\r\n\r\n" +
"Ако преместите папката, пряката връзка продължава да работи, стига да стоят заедно.\r\n";

    // a USB stick was plugged in: tell the page (after a moment, the volume needs time to mount)
    System.Windows.Forms.Timer driveTimer;
    protected override void WndProc(ref Message m)
    {
        base.WndProc(ref m);
        if (m.Msg == 0x0219 && (long)m.WParam == 0x8000) ArmDriveTimer();   // WM_DEVICECHANGE / DBT_DEVICEARRIVAL
    }

    void ArmDriveTimer()
    {
        if (driveTimer == null)
        {
            driveTimer = new System.Windows.Forms.Timer { Interval = 2500 };
            driveTimer.Tick += (s, e) => { driveTimer.Stop(); try { web?.CoreWebView2?.PostWebMessageAsJson("{\"ev\":\"drive\"}"); } catch { } };
        }
        driveTimer.Stop(); driveTimer.Start();
    }

    /* Drives that may be prepared: removable drives (USB sticks); if there are none, external (non-system) disks. */
    List<(string root, string label, double freeGb, double totalGb, bool removable)> ListDrives()
    {
        var all = new List<(string root, string label, double freeGb, double totalGb, bool removable)>();
        var fake = Environment.GetEnvironmentVariable("FAKTURA_TEST_DRIVES");   // tests only: simulate drives
        if (TestRoot != null && fake != null)
        {
            foreach (var r in fake.Split(';', StringSplitOptions.RemoveEmptyEntries)) all.Add((r.EndsWith('\\') ? r : r + "\\", "тест", 1, 1, true));
            return all;
        }
        var sys = Path.GetPathRoot(Environment.SystemDirectory) ?? "C:\\";
        foreach (var d in DriveInfo.GetDrives())
            try
            {
                if (!d.IsReady) continue;
                bool rem = d.DriveType == DriveType.Removable;
                if (!rem && (d.DriveType != DriveType.Fixed || string.Equals(d.RootDirectory.FullName, sys, StringComparison.OrdinalIgnoreCase))) continue;
                all.Add((d.RootDirectory.FullName, d.VolumeLabel, Math.Round(d.AvailableFreeSpace / 1073741824.0, 1), Math.Round(d.TotalSize / 1073741824.0, 1), rem));
            }
            catch { }
        var usb = all.Where(d => d.removable).ToList();
        return (usb.Count > 0 ? usb : all);
    }

    // folder names: current first, then the names used by older builds (so older flash drives are still found)
    static readonly string[] DataFolderNames = { "invoice-data", "Faktura-data" };
    static readonly string[] PortableFolderNames = { "invoice-portable", "Фактури-преносима" };

    /* Looks for flash drives that carry a prepared portable edition: remembered folders first, then every removable drive
       (<drive>:\invoice-portable, <drive>:\ and one level of sub-folders). Returns the candidates, best first. */
    List<object> FindPortable(List<string> knownFolders, HashSet<string> ids)
    {
        var paths = new List<string>();
        void Add(string dataJson) { try { if (File.Exists(dataJson) && !paths.Contains(dataJson, StringComparer.OrdinalIgnoreCase)) paths.Add(dataJson); } catch { } }
        foreach (var k in knownFolders) foreach (var dn in DataFolderNames) Add(Path.Combine(k, dn, "data.json"));
        var fake = Environment.GetEnvironmentVariable("FAKTURA_TEST_DRIVES"); // tests only: simulate drives
        var roots = new List<(string root, bool removable)>();
        if (fake != null) foreach (var r in fake.Split(';', StringSplitOptions.RemoveEmptyEntries)) roots.Add((r, true));
        else
            foreach (var d in DriveInfo.GetDrives())
                try { if (d.IsReady && (d.DriveType == DriveType.Removable || d.DriveType == DriveType.Fixed)) roots.Add((d.RootDirectory.FullName, d.DriveType == DriveType.Removable)); } catch { }
        foreach (var (root, removable) in roots)
        {
            foreach (var pf in PortableFolderNames) foreach (var dn in DataFolderNames) Add(Path.Combine(root, pf, dn, "data.json"));
            foreach (var dn in DataFolderNames) Add(Path.Combine(root, dn, "data.json"));
            if (!removable) continue;
            try
            {
                foreach (var sub in new DirectoryInfo(root).GetDirectories().Take(300))
                {
                    if ((sub.Attributes & FileAttributes.System) != 0 || (sub.Attributes & FileAttributes.Hidden) != 0) continue;
                    foreach (var dn in DataFolderNames) Add(Path.Combine(sub.FullName, dn, "data.json"));
                }
            }
            catch { }
        }
        var own = Path.GetFullPath(DataFile);
        var found = new List<(object item, bool known, DateTime when)>();
        foreach (var path in paths)
        {
            try
            {
                if (string.Equals(Path.GetFullPath(path), own, StringComparison.OrdinalIgnoreCase)) continue; // that's this app's own data
                var fi = new FileInfo(path); string id = null, created = null; int docs = 0;
                try
                {
                    var o = JsonNode.Parse(File.ReadAllText(path));
                    id = o?["portable"]?["id"]?.ToString(); created = o?["portable"]?["created"]?.ToString();
                    docs = (o?["invoices"]?.AsArray().Count ?? 0) + (o?["receipts"]?.AsArray().Count ?? 0);
                }
                catch { continue; }                           // not one of ours
                var folder = Path.GetDirectoryName(Path.GetDirectoryName(path))!;
                bool match = (id != null && ids.Contains(id)) || knownFolders.Any(k => string.Equals(Path.GetFullPath(k), Path.GetFullPath(folder), StringComparison.OrdinalIgnoreCase));
                found.Add((new { path, folder, id, created, docs, modified = fi.LastWriteTime.ToString("dd.MM.yyyy HH:mm"), drive = Path.GetPathRoot(path), match }, match, fi.LastWriteTime));
            }
            catch { }
        }
        return found.OrderByDescending(f => f.known).ThenByDescending(f => f.when).Select(f => f.item).ToList();
    }

    static void CleanStaleProfiles()
    {
        try
        {
            foreach (var d in new DirectoryInfo(Path.GetTempPath()).GetDirectories("Faktura-wv2-*"))
                if ((DateTime.Now - d.LastWriteTime).TotalHours > 1) try { d.Delete(true); } catch { }
        }
        catch { }
    }

    protected override void OnFormClosed(FormClosedEventArgs e)
    {
        base.OnFormClosed(e);
        if (!portableMode || wvDir == null) return;
        try { web.Dispose(); } catch { }
        try // the browser processes need a moment to let go of the folder
        {
            Process.Start(new ProcessStartInfo("cmd.exe", $"/c ping 127.0.0.1 -n 6 >nul & rmdir /s /q \"{wvDir}\"") { CreateNoWindow = true, UseShellExecute = false, WindowStyle = ProcessWindowStyle.Hidden });
        }
        catch { }
    }

    static readonly System.Reflection.Assembly Asm = typeof(MainForm).Assembly;

    void OnWebResource(object s, CoreWebView2WebResourceRequestedEventArgs e)
    {
        var env = web.CoreWebView2.Environment;
        var name = Uri.UnescapeDataString(new Uri(e.Request.Uri).AbsolutePath.TrimStart('/'));
        if (name.Length == 0) name = "index.html";
        using var src = name.Contains("..") ? null : Asm.GetManifestResourceStream("web/" + name);
        if (src == null) { e.Response = env.CreateWebResourceResponse(null, 404, "Not Found", "Content-Type: text/plain"); return; }
        var ms = new MemoryStream(); src.CopyTo(ms); ms.Position = 0;
        var mime = Path.GetExtension(name).ToLowerInvariant() switch
        {
            ".html" => "text/html; charset=utf-8",
            ".js" => "application/javascript; charset=utf-8",
            ".png" => "image/png",
            ".ico" => "image/x-icon",
            _ => "application/octet-stream"
        };
        e.Response = env.CreateWebResourceResponse(ms, 200, "OK", "Content-Type: " + mime + "\r\nCache-Control: no-store");
    }

    static void OpenExternal(string uri)
    {
        if (uri.StartsWith("http://") || uri.StartsWith("https://"))
            try { Process.Start(new ProcessStartInfo(uri) { UseShellExecute = true }); } catch { }
    }

    /* ---------------- data ---------------- */
    void SaveData(string json)
    {
        Directory.CreateDirectory(DataDir);
        var tmp = DataFile + ".tmp";
        File.WriteAllText(tmp, json, new System.Text.UTF8Encoding(false));
        if (File.Exists(DataFile) && (DateTime.Now - lastBak).TotalMinutes > 2) { File.Copy(DataFile, DataFile + ".bak", true); lastBak = DateTime.Now; }
        File.Move(tmp, DataFile, true);
    }

    string BackupNow(string json)
    {
        Directory.CreateDirectory(BackupDir);
        var path = Path.Combine(BackupDir, $"faktura_{DateTime.Now:yyyyMMdd_HHmmss}.json");
        File.WriteAllText(path, json, new System.Text.UTF8Encoding(false));
        Prune();
        return path;
    }

    void Prune()
    {
        var files = new DirectoryInfo(BackupDir).GetFiles("faktura_*.json").OrderByDescending(f => f.Name).ToList();
        foreach (var f in files.Skip(KeepBackups)) try { f.Delete(); } catch { }
    }

    void AutoBackupDaily()
    {
        try
        {
            if (!File.Exists(DataFile)) return;
            Directory.CreateDirectory(BackupDir);
            var today = $"faktura_{DateTime.Now:yyyyMMdd}_";
            if (new DirectoryInfo(BackupDir).GetFiles(today + "*.json").Length == 0)
            { File.Copy(DataFile, Path.Combine(BackupDir, $"faktura_{DateTime.Now:yyyyMMdd_HHmmss}.json")); Prune(); }
        }
        catch { }
    }

    void SaveConfig() { try { File.WriteAllText(ConfigFile, JsonSerializer.Serialize(cfg)); } catch { } }

    /* ---------------- bridge ---------------- */
    async void OnMessage(object s, CoreWebView2WebMessageReceivedEventArgs e)
    {
        JsonNode msg;
        try { msg = JsonNode.Parse(e.WebMessageAsJson); } catch { return; }
        int? id = msg?["id"]?.GetValue<int>();
        string cmd = msg?["cmd"]?.GetValue<string>();
        var arg = msg?["arg"];
        object result = null; string error = null; bool restartAfter = false;
        try
        {
            switch (cmd)
            {
                case "load": result = File.Exists(DataFile) ? File.ReadAllText(DataFile) : null; break;
                case "save": SaveData(arg.GetValue<string>()); break;
                case "backup": result = BackupNow(arg.GetValue<string>()); break;
                case "info":
                    {
                        var dir = new DirectoryInfo(BackupDir);
                        var files = dir.Exists ? dir.GetFiles("faktura_*.json").OrderByDescending(f => f.Name).ToList() : new List<FileInfo>();
                        result = new { dataDir = DataDir, backupDir = BackupDir, count = files.Count, keep = KeepBackups, last = files.Count > 0 ? files[0].LastWriteTime.ToString("dd.MM.yyyy HH:mm") : null };
                        break;
                    }
                case "about":
                    result = new
                    {
                        version = typeof(MainForm).Assembly.GetName().Version is Version av ? $"{av.Major}.{av.Minor}.{av.Build}" : "?",
                        exe = Environment.ProcessPath,
                        dataDir = DataDir,
                        webview2 = web.CoreWebView2.Environment.BrowserVersionString,
                        dotnet = System.Runtime.InteropServices.RuntimeInformation.FrameworkDescription
                    };
                    break;
                case "listDrives":
                    result = ListDrives().Select(d => new { d.root, d.label, d.freeGb, d.totalGb, d.removable }).ToList();
                    break;
                case "makePortable":
                    {
                        var json = arg["state"]!.GetValue<string>();
                        var chosen = arg["drive"]!.GetValue<string>();
                        // only a whole drive from the list is accepted - never a folder
                        var drive = ListDrives().FirstOrDefault(d => string.Equals(d.root, chosen, StringComparison.OrdinalIgnoreCase));
                        if (drive.root == null) throw new Exception("Избраното устройство не е в списъка на флашките.");
                        var driveRoot = drive.root; var dest = Path.Combine(driveRoot, "invoice-portable");
                        var dd = Path.Combine(dest, "invoice-data"); var df = Path.Combine(dd, "data.json");
                        // a stick prepared by an older build keeps its data in the old folder names: warn about unsynced documents there too
                        var existing = new List<string> { df, Path.Combine(driveRoot, "Фактури-преносима", "Faktura-data", "data.json") };
                        bool cancelled = false;
                        foreach (var ex in existing)
                        {
                            if (!File.Exists(ex)) continue;
                            int docs = 0;
                            try { var o = JsonNode.Parse(File.ReadAllText(ex)); docs = (o?["invoices"]?.AsArray().Count ?? 0) + (o?["receipts"]?.AsArray().Count ?? 0); } catch { }
                            if (docs > 0 && TestRoot == null && MessageBox.Show(this, $"В „{Path.GetDirectoryName(Path.GetDirectoryName(ex))}“ има {docs} документа, които може още да не са синхронизирани с основното приложение.\n\nДа продължа ли с подготовката на флашката?", "Преносима версия", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2) != DialogResult.Yes) { cancelled = true; break; }
                        }
                        if (cancelled) break;
                        if (File.Exists(df)) File.Copy(df, df + ".old", true); // keep the previous file as a safety copy
                        Directory.CreateDirectory(dd);
                        // the portable edition is the Lite program only (runs on Windows XP .. 11)
                        foreach (var res in new[] { "invoice-lite.exe", "invoice-lite.exe.config" })
                        {
                            using var rs = Asm.GetManifestResourceStream("lite/" + res) ?? throw new Exception("В програмата липсва вградената Lite версия.");
                            using var fs = File.Create(Path.Combine(dest, res)); rs.CopyTo(fs);
                        }
                        File.WriteAllText(Path.Combine(dest, "README.txt"), ReadmeText, new System.Text.UTF8Encoding(true));
                        File.WriteAllText(df, json, new System.Text.UTF8Encoding(false));
                        string shortcut = null;
                        try
                        {
                            shortcut = Path.Combine(driveRoot, "invoice.lnk");
                            Lnk.Create(shortcut, Path.Combine(dest, "invoice-lite.exe"), dest, "Фактури — преносима версия");
                        }
                        catch { shortcut = null; }               // e.g. a drive that cannot take a file in its root
                        result = new { path = dest, shortcut };
                        break;
                    }
                case "findPortable":
                    {
                        var known = new List<string>();
                        foreach (var k in arg?["paths"]?.AsArray() ?? new JsonArray()) if (k?.GetValue<string>() is string kp && kp.Length > 0) known.Add(kp);
                        var ids = new HashSet<string>();
                        foreach (var k in arg?["ids"]?.AsArray() ?? new JsonArray()) if (k?.GetValue<string>() is string ki && ki.Length > 0) ids.Add(ki);
                        result = FindPortable(known, ids);
                        break;
                    }
                case "readFile":
                    {
                        var path = arg.GetValue<string>();
                        if (!path.EndsWith("data.json", StringComparison.OrdinalIgnoreCase) || !DataFolderNames.Contains(Path.GetFileName(Path.GetDirectoryName(path)) ?? "", StringComparer.OrdinalIgnoreCase)) throw new Exception("Позволено е само четене на invoice-data\\data.json.");
                        result = File.ReadAllText(path);
                        break;
                    }
                case "testDriveArrival": if (TestRoot != null) ArmDriveTimer(); break;   // tests only
                case "sendData":
                    {
                        var info = arg["info"]?.GetValue<string>() ?? "";
                        var zip = SendData.MakeZip(arg["state"]!.GetValue<string>(), info);
                        var url = SendData.Endpoint; var test = TestRoot != null;
                        if (test && Environment.GetEnvironmentVariable("FAKTURA_SEND_URL") is string tu && tu.Length > 0) url = tu;   // tests only
                        if (url.Length == 0) { result = new { sent = false, reason = "noendpoint", message = "", file = zip, kb = 0L }; break; }
                        if (!SendData.Allowed(url, test)) throw new Exception("Невалиден адрес за изпращане.");
                        try { await SendData.Upload(zip, info, Updater.Current.ToString(), url); result = new { sent = true, reason = "", message = "", file = zip, kb = new FileInfo(zip).Length / 1024 }; }
                        catch (Exception ex) { result = new { sent = false, reason = "error", message = ex.Message, file = zip, kb = 0L }; }
                        break;
                    }
                case "revealFile":
                    {
                        var path = Path.GetFullPath(arg.GetValue<string>());
                        if (!path.StartsWith(SendData.FolderPath + Path.DirectorySeparatorChar, StringComparison.OrdinalIgnoreCase) || !File.Exists(path)) throw new Exception("Файлът не е намерен.");
                        if (TestRoot == null) Process.Start(new ProcessStartInfo("explorer.exe", $"/select,\"{path}\"") { UseShellExecute = true });
                        break;
                    }
                case "updateCheck": result = await Updater.Check(); break;
                case "updateInstall":
                    await Updater.Install(pct => { try { web.CoreWebView2.PostWebMessageAsJson(JsonSerializer.Serialize(new { ev = "upd", pct })); } catch { } });
                    restartAfter = true;
                    break;
                case "liteRefresh":
                    result = Updater.LiteRefresh(Asm, arg["folder"]!.GetValue<string>(), arg["apply"]?.GetValue<bool>() ?? false);
                    break;
                case "openDataDir":
                    Process.Start(new ProcessStartInfo("explorer.exe", $"\"{DataDir}\"") { UseShellExecute = true });
                    break;
                case "pdf":
                    {
                        var name = string.Concat((arg?["name"]?.GetValue<string>() ?? "faktura.pdf").Select(c => Path.GetInvalidFileNameChars().Contains(c) ? '_' : c));
                        var initial = !string.IsNullOrEmpty(cfg.PdfDir) && Directory.Exists(cfg.PdfDir) ? cfg.PdfDir : Path.Combine(Docs, "Фактури");
                        Directory.CreateDirectory(initial);
                        string target;
                        if (TestRoot != null) target = Path.Combine(TestRoot, name);
                        else
                        {
                            using var dlg = new SaveFileDialog { Filter = "PDF (*.pdf)|*.pdf", FileName = name, InitialDirectory = initial, OverwritePrompt = true, Title = "Запази фактурата като PDF" };
                            if (dlg.ShowDialog(this) != DialogResult.OK) break;
                            target = dlg.FileName; cfg.PdfDir = Path.GetDirectoryName(target); SaveConfig();
                        }
                        var ps = web.CoreWebView2.Environment.CreatePrintSettings();
                        ps.Orientation = CoreWebView2PrintOrientation.Portrait;
                        ps.PageWidth = 8.27; ps.PageHeight = 11.69;
                        ps.MarginTop = ps.MarginBottom = ps.MarginLeft = ps.MarginRight = 0;
                        ps.ShouldPrintBackgrounds = true; ps.ShouldPrintHeaderAndFooter = false;
                        if (!await web.CoreWebView2.PrintToPdfAsync(target, ps)) throw new Exception("PDF не беше създаден.");
                        result = target;
                        break;
                    }
                default: error = "unknown command " + cmd; break;
            }
        }
        catch (Exception ex) { error = ex.Message; }
        if (id != null)
        {
            try { web.CoreWebView2.PostWebMessageAsJson(JsonSerializer.Serialize(new { id, result, error })); } catch { }
        }
        if (restartAfter) { await Task.Delay(400); Close(); }   // the new version is already starting; closing saves the state first
    }

    /* ---------------- shutdown: flush the latest state to disk ---------------- */
    protected override async void OnFormClosing(FormClosingEventArgs e)
    {
        if (closing) { base.OnFormClosing(e); return; }
        e.Cancel = true; closing = true;
        try
        {
            cfg.Max = WindowState == FormWindowState.Maximized;
            var b = cfg.Max ? RestoreBounds : Bounds;
            cfg.X = b.X; cfg.Y = b.Y; cfg.W = b.Width; cfg.H = b.Height; SaveConfig();
            if (web.CoreWebView2 != null)
            {
                var t = web.CoreWebView2.ExecuteScriptAsync("window.__state ? window.__state() : null");
                if (await Task.WhenAny(t, Task.Delay(2500)) == t && t.Result != "null")
                {
                    var j = JsonSerializer.Deserialize<string>(t.Result);
                    if (!string.IsNullOrEmpty(j)) SaveData(j);
                }
            }
        }
        catch { }
        Close();
    }
}
