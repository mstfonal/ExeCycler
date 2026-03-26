using System;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.IO;
using System.Net.Http;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text.Json;
using System.Threading;
using System.Windows.Forms;
using Microsoft.Win32;

namespace AutoStart
{
    static class Program
    {
        [STAThread]
        static void Main(string[] args)
        {
            if (args.Length > 0 && args[0] == "--rollback")
            {
                string backupPath = args.Length > 1 ? args[1] : "";
                PerformRollback(backupPath);
                return;
            }
            if (args.Length > 0 && args[0] == "--updated")
                Thread.Sleep(1500);
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new CyclerContext());
        }
        private static void PerformRollback(string backupPath)
        {
            try
            {
                string currentExe = Application.ExecutablePath;
                if (!string.IsNullOrEmpty(backupPath) && File.Exists(backupPath))
                {
                    Thread.Sleep(1500);
                    File.Copy(backupPath, currentExe, overwrite: true);
                    Process.Start(new ProcessStartInfo(currentExe) { UseShellExecute = true });
                }
            }
            catch { }
        }
    }

    static class NativeMethods
    {
        [DllImport("user32.dll")]
        public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
        [DllImport("user32.dll")]
        public static extern bool SetForegroundWindow(IntPtr hWnd);
        [DllImport("user32.dll")]
        public static extern IntPtr FindWindow(string? lpClassName, string lpWindowName);
        public const int SW_MINIMIZE = 6;
    }

    class CyclerConfig
    {
        public string ExePath { get; set; } = "";
        public int RunSeconds { get; set; } = 600;
        public int OffSeconds { get; set; } = 10;
        public int HeartbeatSeconds { get; set; } = 30;
        public int StopWaitSeconds { get; set; } = 5;
        public bool AutoUpdate { get; set; } = true;
        public bool MinimizeTarget { get; set; } = true;
    }

    class UpdateInfo
    {
        public string TagName { get; set; } = "";
        public string DownloadUrl { get; set; } = "";
    }

    class CyclerContext : ApplicationContext
    {
        private const string GITHUB_REPO = "mstfonal/AutoStart";
        private static readonly string CURRENT_VERSION =
            Assembly.GetExecutingAssembly().GetName().Version?.ToString(3) ?? "1.0.0";

        private NotifyIcon? _tray;
        private Thread? _workerThread;
        private Thread? _updateThread;
        private volatile bool _running = true;
        private CyclerConfig _config = new CyclerConfig();
        private string _configPath = "";
        private string _logPath = "";
        private string _status = "Starting...";
        private int _cycleCount = 0;
        private readonly object _statusLock = new object();
        private readonly object _logLock = new object();
        private SynchronizationContext? _syncCtx;

        public CyclerContext()
        {
            _syncCtx = SynchronizationContext.Current;
            string appData = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData), "AutoStart");
            Directory.CreateDirectory(appData);
            _logPath = Path.Combine(appData, "cyclerlog.txt");
            _configPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "config.json");
            InitTray();
            RegisterAutostart();
            _config = LoadOrCreateConfig();
            _workerThread = new Thread(WorkerLoop) { IsBackground = true, Name = "CyclerWorker" };
            _workerThread.Start();
            if (_config.AutoUpdate)
            {
                _updateThread = new Thread(UpdateLoop) { IsBackground = true, Name = "UpdateChecker" };
                _updateThread.Start();
            }
        }

        private static Icon CreateAppIcon()
        {
            using var bmp = new Bitmap(32, 32);
            using var g = Graphics.FromImage(bmp);
            g.SmoothingMode = SmoothingMode.AntiAlias;
            g.Clear(Color.Transparent);
            using var bgBrush = new SolidBrush(Color.FromArgb(30, 140, 60));
            g.FillEllipse(bgBrush, 1, 1, 30, 30);
            using var borderPen = new Pen(Color.FromArgb(20, 100, 40), 1.5f);
            g.DrawEllipse(borderPen, 1, 1, 30, 30);
            using var font = new Font("Arial", 14, FontStyle.Bold, GraphicsUnit.Pixel);
            using var brush = new SolidBrush(Color.White);
            var sf = new StringFormat { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center };
            g.DrawString("C", font, brush, new RectangleF(0, 0, 32, 32), sf);
            return Icon.FromHandle(bmp.GetHicon());
        }

        private void InitTray()
        {
            var menu = new ContextMenuStrip();
            menu.Items.Add("Status", null, OnShowStatus);
            menu.Items.Add("Settings", null, OnShowSettings);
            menu.Items.Add("Check for Updates", null, OnCheckUpdate);
            menu.Items.Add("Open Log", null, OnOpenLog);
            menu.Items.Add(new ToolStripSeparator());
            menu.Items.Add("Exit", null, OnExit);
            _tray = new NotifyIcon
            {
                Icon = CreateAppIcon(),
                Text = "AutoStart v" + CURRENT_VERSION,
                ContextMenuStrip = menu,
                Visible = true
            };
            _tray.DoubleClick += OnShowStatus;
        }

        private void SetStatus(string s)
        {
            lock (_statusLock) { _status = s; }
            string tip = "AutoStart v" + CURRENT_VERSION + " - " + s;
            if (tip.Length > 63) tip = tip.Substring(0, 63);
            _syncCtx?.Post(_ => { if (_tray != null) _tray.Text = tip; }, null);
        }

        private void ShowBalloon(string title, string message, ToolTipIcon icon = ToolTipIcon.Info)
        {
            _syncCtx?.Post(_ => { _tray?.ShowBalloonTip(5000, title, message, icon); }, null);
        }

        private void OnShowStatus(object? sender, EventArgs e)
        {
            string s;
            lock (_statusLock) { s = _status; }
            MessageBox.Show(
                "Version: " + CURRENT_VERSION +
                "\nStatus: " + s +
                "\nCycles: " + _cycleCount +
                "\n\nEXE: " + (_config?.ExePath ?? "-") +
                "\nRun: " + _config?.RunSeconds + "s  |  Off: " + _config?.OffSeconds + "s" +
                "\nAuto-update: " + (_config?.AutoUpdate == true ? "On" : "Off") +
                "\nMinimize target: " + (_config?.MinimizeTarget == true ? "On" : "Off"),
                "AutoStart", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void OnShowSettings(object? sender, EventArgs e)
        {
            if (_syncCtx != null)
                _syncCtx.Post(_ => ShowSettingsForm(), null);
            else
                ShowSettingsForm();
        }

        private void ShowSettingsForm()
        {
            var form = new Form
            {
                Text = "AutoStart - Settings",
                Width = 480, Height = 330,
                FormBorderStyle = FormBorderStyle.FixedDialog,
                MaximizeBox = false, MinimizeBox = false,
                StartPosition = FormStartPosition.CenterScreen
            };
            int y = 20;
            var lblExe = new Label { Text = "Selected EXE:", Left = 20, Top = y, Width = 160, AutoSize = false, TextAlign = ContentAlignment.MiddleLeft };
            string exeShort = string.IsNullOrEmpty(_config.ExePath) ? "(None)" : Path.GetFileName(_config.ExePath);
            var lblExePath = new Label { Text = exeShort, Left = 190, Top = y, Width = 260, AutoSize = false, Height = 20, TextAlign = ContentAlignment.MiddleLeft, ForeColor = Color.DarkBlue };
            y += 30;
            var btnReset = new Button { Text = "Reset EXE (Pick New)", Left = 190, Top = y, Width = 260, Height = 28 };
            btnReset.Click += (s, ev) =>
            {
                string picked = PickExe();
                if (!string.IsNullOrEmpty(picked))
                {
                    _config.ExePath = picked;
                    SaveConfig(_config);
                    lblExePath.Text = Path.GetFileName(picked);
                    Log("EXE changed: " + picked);
                    MessageBox.Show("EXE updated:\n" + picked + "\n\nTakes effect on next cycle.", "AutoStart", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            };
            y += 38;
            var lblRun = new Label { Text = "Run duration (sec):", Left = 20, Top = y + 3, Width = 160 };
            var txtRun = new TextBox { Text = _config.RunSeconds.ToString(), Left = 190, Top = y, Width = 80 };
            y += 35;
            var lblOff = new Label { Text = "Off duration (sec):", Left = 20, Top = y + 3, Width = 160 };
            var txtOff = new TextBox { Text = _config.OffSeconds.ToString(), Left = 190, Top = y, Width = 80 };
            y += 35;
            var chkUpdate = new CheckBox { Text = "Auto-update", Left = 190, Top = y, Width = 260, Checked = _config.AutoUpdate };
            y += 28;
            var chkMin = new CheckBox { Text = "Minimize target window", Left = 190, Top = y, Width = 260, Checked = _config.MinimizeTarget };
            y += 40;
            var btnSave = new Button { Text = "Save", Left = 190, Top = y, Width = 110, Height = 28 };
            btnSave.Click += (s, ev) =>
            {
                if (int.TryParse(txtRun.Text, out int r) && r > 0) _config.RunSeconds = r;
                if (int.TryParse(txtOff.Text, out int o) && o >= 0) _config.OffSeconds = o;
                _config.AutoUpdate = chkUpdate.Checked;
                _config.MinimizeTarget = chkMin.Checked;
                SaveConfig(_config);
                Log("Settings saved.");
                MessageBox.Show("Settings saved.", "AutoStart", MessageBoxButtons.OK, MessageBoxIcon.Information);
                form.Close();
            };
            var btnCancel = new Button { Text = "Cancel", Left = 310, Top = y, Width = 100, Height = 28 };
            btnCancel.Click += (s, ev) => form.Close();
            form.Controls.AddRange(new Control[] { lblExe, lblExePath, btnReset, lblRun, txtRun, lblOff, txtOff, chkUpdate, chkMin, btnSave, btnCancel });
            form.ShowDialog();
        }

        private void OnCheckUpdate(object? sender, EventArgs e)
        {
            new Thread(() =>
            {
                var info = CheckForUpdate();
                if (info != null) ApplyUpdate(info);
                else ShowBalloon("AutoStart", "No update available. Version: " + CURRENT_VERSION);
            }) { IsBackground = true }.Start();
        }

        private void OnOpenLog(object? sender, EventArgs e)
        {
            if (File.Exists(_logPath)) Process.Start(new ProcessStartInfo(_logPath) { UseShellExecute = true });
            else MessageBox.Show("Log file not created yet.", "AutoStart");
        }

        private void OnExit(object? sender, EventArgs e)
        {
            _running = false;
            if (_tray != null) _tray.Visible = false;
            Log("=== Exit requested by user ===");
            Application.Exit();
        }

        private void UpdateLoop()
        {
            Thread.Sleep(60000);
            while (_running)
            {
                try
                {
                    var info = CheckForUpdate();
                    if (info != null)
                    {
                        Log("Update found: " + info.TagName);
                        ShowBalloon("AutoStart - Update", "New version: " + info.TagName + "  Current: " + CURRENT_VERSION + "  Applying in 60s...");
                        for (int i = 0; i < 120 && _running; i++) Thread.Sleep(500);
                        if (_running) ApplyUpdate(info);
                    }
                }
                catch (Exception ex) { Log("Update check error: " + ex.Message); }
                for (int i = 0; i < 43200 && _running; i++) Thread.Sleep(500);
            }
        }

        private UpdateInfo? CheckForUpdate()
        {
            try
            {
                using var client = new HttpClient();
                client.Timeout = TimeSpan.FromSeconds(10);
                client.DefaultRequestHeaders.Add("User-Agent", "ExeCycler/" + CURRENT_VERSION);
                string json = client.GetStringAsync("https://api.github.com/repos/" + GITHUB_REPO + "/releases/latest").GetAwaiter().GetResult();
                using var doc = JsonDocument.Parse(json);
                var root = doc.RootElement;
                string tagName = root.GetProperty("tag_name").GetString() ?? "";
                if (!IsNewerVersion(tagName.TrimStart('v'), CURRENT_VERSION)) return null;
                string dlUrl = "";
                if (root.TryGetProperty("assets", out var assets))
                    foreach (var asset in assets.EnumerateArray())
                    {
                        string name = asset.GetProperty("name").GetString() ?? "";
                        if (name.EndsWith(".exe", StringComparison.OrdinalIgnoreCase))
                        { dlUrl = asset.GetProperty("browser_download_url").GetString() ?? ""; break; }
                    }
                if (string.IsNullOrEmpty(dlUrl)) return null;
                return new UpdateInfo { TagName = tagName, DownloadUrl = dlUrl };
            }
            catch (Exception ex) { Log("Update check failed: " + ex.Message); return null; }
        }

        private void ApplyUpdate(UpdateInfo info)
        {
            string currentExe = Application.ExecutablePath;
            string exeDir = Path.GetDirectoryName(currentExe) ?? AppDomain.CurrentDomain.BaseDirectory;
            string backupPath = Path.Combine(exeDir, "AutoStart_backup.exe");
            string tempPath = Path.Combine(exeDir, "AutoStart_new.exe");
            Log("Applying update: " + info.TagName);
            ShowBalloon("AutoStart", "Downloading update " + info.TagName + "...");
            try
            {
                using var client = new HttpClient();
                client.Timeout = TimeSpan.FromSeconds(120);
                client.DefaultRequestHeaders.Add("User-Agent", "ExeCycler/" + CURRENT_VERSION);
                byte[] bytes = client.GetByteArrayAsync(info.DownloadUrl).GetAwaiter().GetResult();
                if (bytes.Length < 1024) throw new Exception("Downloaded file too small.");
                File.WriteAllBytes(tempPath, bytes);
                Log("Downloaded: " + bytes.Length + " bytes");
                if (File.Exists(backupPath)) File.Delete(backupPath);
                File.Copy(currentExe, backupPath);
                Log("Backup created: " + backupPath);
                // PowerShell hidden - no UAC, no CMD window
                string nl = "\r\n";
                string safe_temp = tempPath.Replace("'", "''");
                string safe_cur = currentExe.Replace("'", "''");
                string ps1Path = Path.Combine(exeDir, "update_helper.ps1");
                string ps1 = "Start-Sleep -Seconds 2" + nl
                    + "Copy-Item -Path '" + safe_temp + "' -Destination '" + safe_cur + "' -Force" + nl
                    + "Remove-Item -Path '" + safe_temp + "' -ErrorAction SilentlyContinue" + nl
                    + "Start-Process -FilePath '" + safe_cur + "' -ArgumentList '--updated'" + nl
                    + "Remove-Item -Path $MyInvocation.MyCommand.Path -ErrorAction SilentlyContinue" + nl;
                File.WriteAllText(ps1Path, ps1);
                _running = false;
                Process.Start(new ProcessStartInfo
                {
                    FileName = "powershell.exe",
                    Arguments = "-NonInteractive -WindowStyle Hidden -ExecutionPolicy Bypass -File \"" + ps1Path + "\"",
                    UseShellExecute = false,
                    CreateNoWindow = true,
                    WindowStyle = ProcessWindowStyle.Hidden
                });
                Thread.Sleep(500);
                _syncCtx?.Post(_ => { if (_tray != null) _tray.Visible = false; Application.Exit(); }, null);
            }
            catch (Exception ex)
            {
                Log("UPDATE ERROR: " + ex.Message);
                try { if (File.Exists(tempPath)) File.Delete(tempPath); } catch { }
                ShowBalloon("AutoStart - Update Failed", ex.Message, ToolTipIcon.Error);
            }
        }

        private static bool IsNewerVersion(string remote, string local)
        {
            try { return new Version(remote) > new Version(local); }
            catch { return false; }
        }

        private CyclerConfig LoadOrCreateConfig()
        {
            if (File.Exists(_configPath))
            {
                try
                {
                    var cfg = JsonSerializer.Deserialize<CyclerConfig>(File.ReadAllText(_configPath));
                    if (cfg != null)
                    {
                        Log("Config loaded: " + _configPath);
                        if (string.IsNullOrEmpty(cfg.ExePath) || !File.Exists(cfg.ExePath))
                        {
                            Log("Saved EXE not found, re-picking: " + cfg.ExePath);
                            cfg.ExePath = PickExe();
                            SaveConfig(cfg);
                        }
                        return cfg;
                    }
                }
                catch (Exception ex) { Log("Config read error: " + ex.Message); }
            }
            var newCfg = new CyclerConfig { ExePath = PickExe() };
            SaveConfig(newCfg);
            return newCfg;
        }

        private void SaveConfig(CyclerConfig cfg)
        {
            try { File.WriteAllText(_configPath, JsonSerializer.Serialize(cfg, new JsonSerializerOptions { WriteIndented = true })); }
            catch (Exception ex) { Log("Config save error: " + ex.Message); }
        }

        private string PickExe()
        {
            string result = "";
            var t = new Thread(() =>
            {
                using var dlg = new OpenFileDialog { Filter = "Applications (*.exe)|*.exe", Title = "Select EXE to cycle" };
                if (dlg.ShowDialog() == DialogResult.OK) result = dlg.FileName;
            });
            t.SetApartmentState(ApartmentState.STA);
            t.Start(); t.Join();
            if (string.IsNullOrEmpty(result))
            {
                MessageBox.Show("No EXE selected. Exiting.", "AutoStart", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                Application.Exit(); Environment.Exit(0);
            }
            Log("EXE selected: " + result);
            return result;
        }

        private void RegisterAutostart()
        {
            try
            {
                using var key = Registry.CurrentUser.OpenSubKey(@"SOFTWARE\Microsoft\Windows\CurrentVersion\Run", true);
                key?.SetValue("AutoStart", "\"" + Application.ExecutablePath + "\"");
                Log("Autostart registered.");
            }
            catch (Exception ex) { Log("Autostart error: " + ex.Message); }
        }

        private void WorkerLoop()
        {
            Log("=== AutoStart v" + CURRENT_VERSION + " started ===");
            Log("Config: Run=" + _config.RunSeconds + "s, Off=" + _config.OffSeconds + "s, EXE=" + _config.ExePath);
            while (_running)
            {
                _cycleCount++;
                Log("=== CYCLE " + _cycleCount + " START ===");
                SetStatus("Cycle " + _cycleCount + " - launching EXE");
                if (GetProcessCount() == 0)
                {
                    if (!StartExe("Cycle start"))
                    {
                        Log("Launch failed. Waiting " + _config.OffSeconds + "s.");
                        SetStatus("Launch failed - waiting");
                        SleepCancellable(_config.OffSeconds * 1000);
                        continue;
                    }
                }
                else Log("EXE already running.");
                SetStatus("Cycle " + _cycleCount + " - running (" + _config.RunSeconds + "s)");
                var runEnd = DateTime.UtcNow.AddSeconds(_config.RunSeconds);
                while (DateTime.UtcNow < runEnd && _running)
                {
                    SleepCancellable(_config.HeartbeatSeconds * 1000);
                    if (!_running) break;
                    if (GetProcessCount() == 0) { Log("Heartbeat: EXE gone, restarting"); StartExe("Heartbeat restart"); }
                }
                if (!_running) break;
                Log("Run time elapsed. Stopping EXE.");
                SetStatus("Cycle " + _cycleCount + " - stopping");
                StopExe();
                Log("Off phase: " + _config.OffSeconds + "s");
                SetStatus("Cycle " + _cycleCount + " - off (" + _config.OffSeconds + "s)");
                SleepCancellable(_config.OffSeconds * 1000);
            }
            Log("=== Worker stopped ===");
        }

        private int GetProcessCount()
        {
            string exeName = Path.GetFileNameWithoutExtension(_config.ExePath);
            string targetLow = _config.ExePath.ToLowerInvariant();
            int count = 0;
            foreach (var p in Process.GetProcessesByName(exeName))
            {
                try { if (p.MainModule?.FileName?.ToLowerInvariant() == targetLow) count++; }
                catch { }
                finally { p.Dispose(); }
            }
            return count;
        }

        private bool StartExe(string reason)
        {
            Log("START: " + reason);
            try
            {
                var psi = new ProcessStartInfo
                {
                    FileName = _config.ExePath,
                    WorkingDirectory = Path.GetDirectoryName(_config.ExePath),
                    WindowStyle = ProcessWindowStyle.Minimized,
                    UseShellExecute = true
                };
                var p = Process.Start(psi);
                Log("START OK PID=" + p?.Id);
                if (_config.MinimizeTarget && p != null)
                {
                    new Thread(() =>
                    {
                        try
                        {
                            for (int i = 0; i < 20; i++)
                            {
                                Thread.Sleep(250); p.Refresh();
                                if (p.MainWindowHandle != IntPtr.Zero)
                                { NativeMethods.ShowWindow(p.MainWindowHandle, NativeMethods.SW_MINIMIZE); break; }
                            }
                        }
                        catch { }
                    }) { IsBackground = true }.Start();
                }
                return true;
            }
            catch (Exception ex) { Log("START ERROR: " + ex.Message); return false; }
        }

        private void StopExe()
        {
            string exeName = Path.GetFileNameWithoutExtension(_config.ExePath);
            string targetLow = _config.ExePath.ToLowerInvariant();
            var procs = new System.Collections.Generic.List<Process>();
            foreach (var p in Process.GetProcessesByName(exeName))
            {
                try { if (p.MainModule?.FileName?.ToLowerInvariant() == targetLow) procs.Add(p); else p.Dispose(); }
                catch { p.Dispose(); }
            }
            if (procs.Count == 0) { Log("STOP: not running."); return; }
            foreach (var p in procs) { try { if (p.MainWindowHandle != IntPtr.Zero) p.CloseMainWindow(); } catch { } }
            var deadline = DateTime.UtcNow.AddSeconds(_config.StopWaitSeconds);
            while (DateTime.UtcNow < deadline) { Thread.Sleep(250); if (GetProcessCount() == 0) { Log("STOP: graceful OK"); foreach (var p in procs) p.Dispose(); return; } }
            foreach (var p in procs) { try { p.Kill(); Log("STOP: killed PID=" + p.Id); } catch (Exception ex) { Log("STOP kill error: " + ex.Message); } finally { p.Dispose(); } }
            Thread.Sleep(500);
            Log(GetProcessCount() == 0 ? "STOP: force kill OK." : "STOP: WARNING - process still running!");
        }

        private void SleepCancellable(int ms)
        {
            int elapsed = 0;
            while (elapsed < ms && _running) { Thread.Sleep(Math.Min(500, ms - elapsed)); elapsed += 500; }
        }

        private void Log(string message)
        {
            string line = "[" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "] " + message;
            try { lock (_logLock) File.AppendAllText(_logPath, line + Environment.NewLine, System.Text.Encoding.UTF8); }
            catch { }
        }
    }
}
