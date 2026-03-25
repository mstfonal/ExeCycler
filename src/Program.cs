using System;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Text.Json;
using System.Threading;
using System.Windows.Forms;
using Microsoft.Win32;

namespace ExeCycler
{
    static class Program
    {
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new CyclerContext());
        }
    }

    class CyclerConfig
    {
        public string ExePath { get; set; } = "";
        public int RunSeconds { get; set; } = 600;
        public int OffSeconds { get; set; } = 10;
        public int HeartbeatSeconds { get; set; } = 30;
        public int StopWaitSeconds { get; set; } = 5;
    }

    class CyclerContext : ApplicationContext
    {
        private NotifyIcon _tray;
        private Thread _workerThread;
        private volatile bool _running = true;
        private CyclerConfig _config;
        private string _configPath;
        private string _logPath;
        private string _status = "Başlatılıyor...";
        private int _cycleCount = 0;
        private readonly object _statusLock = new object();
        private readonly object _logLock = new object();

        public CyclerContext()
        {
            string appData = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData),
                "ExeCycler");
            Directory.CreateDirectory(appData);
            _logPath = Path.Combine(appData, "cyclerlog.txt");
            _configPath = Path.Combine(
                AppDomain.CurrentDomain.BaseDirectory, "config.json");

            InitTray();
            RegisterAutostart();
            _config = LoadOrCreateConfig();

            _workerThread = new Thread(WorkerLoop)
            {
                IsBackground = true,
                Name = "CyclerWorker"
            };
            _workerThread.Start();
        }

        private void InitTray()
        {
            var menu = new ContextMenuStrip();
            menu.Items.Add("Durum", null, OnShowStatus);
            menu.Items.Add("Log dosyasını aç", null, OnOpenLog);
            menu.Items.Add(new ToolStripSeparator());
            menu.Items.Add("Çıkış", null, OnExit);

            _tray = new NotifyIcon
            {
                Icon = SystemIcons.Application,
                Text = "EXE Cycler",
                ContextMenuStrip = menu,
                Visible = true
            };
            _tray.DoubleClick += OnShowStatus;
        }

        private void SetStatus(string s)
        {
            lock (_statusLock) { _status = s; }
            string tip = "EXE Cycler - " + s;
            if (tip.Length > 63) tip = tip.Substring(0, 63);
            if (_tray != null)
                _tray.Invoke((Action)(() => _tray.Text = tip));
        }

        private void OnShowStatus(object sender, EventArgs e)
        {
            string s;
            lock (_statusLock) { s = _status; }
            MessageBox.Show(
                $"Durum: {s}\nDöngü sayısı: {_cycleCount}\n\nEXE: {_config?.ExePath ?? "-"}\nRun: {_config?.RunSeconds}s | Off: {_config?.OffSeconds}s",
                "EXE Cycler", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void OnOpenLog(object sender, EventArgs e)
        {
            if (File.Exists(_logPath))
                Process.Start(new ProcessStartInfo(_logPath) { UseShellExecute = true });
            else
                MessageBox.Show("Log dosyası henüz oluşturulmadı.", "EXE Cycler");
        }

        private void OnExit(object sender, EventArgs e)
        {
            _running = false;
            _tray.Visible = false;
            Log("=== Kullanıcı isteğiyle çıkılıyor ===");
            Application.Exit();
        }

        private CyclerConfig LoadOrCreateConfig()
        {
            if (File.Exists(_configPath))
            {
                try
                {
                    string json = File.ReadAllText(_configPath);
                    var cfg = JsonSerializer.Deserialize<CyclerConfig>(json);
                    if (cfg != null)
                    {
                        Log($"Config yüklendi: {_configPath}");
                        if (string.IsNullOrEmpty(cfg.ExePath) || !File.Exists(cfg.ExePath))
                        {
                            cfg.ExePath = PickExe();
                            SaveConfig(cfg);
                        }
                        return cfg;
                    }
                }
                catch (Exception ex)
                {
                    Log($"Config okuma hatası: {ex.Message}");
                }
            }

            var newCfg = new CyclerConfig();
            newCfg.ExePath = PickExe();
            SaveConfig(newCfg);
            return newCfg;
        }

        private void SaveConfig(CyclerConfig cfg)
        {
            try
            {
                var opts = new JsonSerializerOptions { WriteIndented = true };
                File.WriteAllText(_configPath, JsonSerializer.Serialize(cfg, opts));
            }
            catch (Exception ex)
            {
                Log($"Config kaydetme hatası: {ex.Message}");
            }
        }

        private string PickExe()
        {
            string result = "";
            var t = new Thread(() =>
            {
                using var dlg = new OpenFileDialog
                {
                    Filter = "Uygulamalar (*.exe)|*.exe",
                    Title = "Döngüde çalıştırılacak EXE'yi seçin"
                };
                if (dlg.ShowDialog() == DialogResult.OK)
                    result = dlg.FileName;
            });
            t.SetApartmentState(ApartmentState.STA);
            t.Start();
            t.Join();

            if (string.IsNullOrEmpty(result))
            {
                MessageBox.Show("EXE seçilmedi. Program kapanıyor.", "EXE Cycler",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                Application.Exit();
                Environment.Exit(0);
            }

            Log($"EXE seçildi: {result}");
            return result;
        }

        private void RegisterAutostart()
        {
            try
            {
                string exePath = Application.ExecutablePath;
                using var key = Registry.CurrentUser.OpenSubKey(
                    @"SOFTWARE\Microsoft\Windows\CurrentVersion\Run", true);
                key?.SetValue("ExeCycler", $"\"{exePath}\"");
                Log("Autostart kaydedildi.");
            }
            catch (Exception ex)
            {
                Log($"Autostart kayıt hatası: {ex.Message}");
            }
        }

        private void WorkerLoop()
        {
            Log("=== EXE Cycler başlatıldı ===");
            Log($"Config: Run={_config.RunSeconds}s, Off={_config.OffSeconds}s, " +
                $"Heartbeat={_config.HeartbeatSeconds}s, EXE={_config.ExePath}");

            while (_running)
            {
                _cycleCount++;
                Log($"=== DÖNGÜ {_cycleCount} BAŞLADI ===");
                SetStatus($"Döngü {_cycleCount} - EXE başlatılıyor");

                if (GetProcessCount() == 0)
                {
                    bool started = StartExe("Döngü başlangıcı");
                    if (!started)
                    {
                        Log($"Başlatma başarısız. {_config.OffSeconds}s bekleniyor.");
                        SetStatus("Başlatma hatası - bekleniyor");
                        SleepCancellable(_config.OffSeconds * 1000);
                        continue;
                    }
                }
                else
                {
                    Log("EXE zaten çalışıyor.");
                }

                SetStatus($"Döngü {_cycleCount} - RUN ({_config.RunSeconds}s)");
                var runEnd = DateTime.UtcNow.AddSeconds(_config.RunSeconds);

                while (DateTime.UtcNow < runEnd && _running)
                {
                    SleepCancellable(_config.HeartbeatSeconds * 1000);
                    if (!_running) break;

                    int cnt = GetProcessCount();
                    if (cnt == 0)
                    {
                        Log("HEARTBEAT: EXE çalışmıyor -> yeniden başlatılıyor");
                        StartExe("RUN fazında EXE kapandı");
                    }
                }

                if (!_running) break;

                Log("RUN süresi doldu. EXE durduruluyor.");
                SetStatus($"Döngü {_cycleCount} - durduruluyor");
                StopExe();

                Log($"OFF fazı: {_config.OffSeconds}s bekleniyor");
                SetStatus($"Döngü {_cycleCount} - OFF ({_config.OffSeconds}s)");
                SleepCancellable(_config.OffSeconds * 1000);
            }

            Log("=== Worker durdu ===");
        }

        private int GetProcessCount()
        {
            string exeName = Path.GetFileNameWithoutExtension(_config.ExePath);
            string targetLow = _config.ExePath.ToLowerInvariant();
            int count = 0;
            foreach (var p in Process.GetProcessesByName(exeName))
            {
                try
                {
                    if (p.MainModule?.FileName?.ToLowerInvariant() == targetLow)
                        count++;
                }
                catch { }
                finally { p.Dispose(); }
            }
            return count;
        }

        private bool StartExe(string reason)
        {
            Log($"START: {reason}");
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
                Log($"START: başarılı PID={p?.Id}");
                return true;
            }
            catch (Exception ex)
            {
                Log($"START HATASI: {ex.Message}");
                return false;
            }
        }

        private void StopExe()
        {
            string exeName = Path.GetFileNameWithoutExtension(_config.ExePath);
            string targetLow = _config.ExePath.ToLowerInvariant();
            var processes = new System.Collections.Generic.List<Process>();

            foreach (var p in Process.GetProcessesByName(exeName))
            {
                try
                {
                    if (p.MainModule?.FileName?.ToLowerInvariant() == targetLow)
                        processes.Add(p);
                    else
                        p.Dispose();
                }
                catch { p.Dispose(); }
            }

            if (processes.Count == 0)
            {
                Log("STOP: zaten çalışmıyor.");
                return;
            }

            foreach (var p in processes)
            {
                try
                {
                    if (p.MainWindowHandle != IntPtr.Zero)
                        p.CloseMainWindow();
                }
                catch { }
            }

            var deadline = DateTime.UtcNow.AddSeconds(_config.StopWaitSeconds);
            while (DateTime.UtcNow < deadline)
            {
                Thread.Sleep(250);
                if (GetProcessCount() == 0)
                {
                    Log("STOP: graceful başarılı.");
                    foreach (var p in processes) p.Dispose();
                    return;
                }
            }

            foreach (var p in processes)
            {
                try { p.Kill(); Log($"STOP: kill PID={p.Id}"); }
                catch (Exception ex) { Log($"STOP: kill hatası PID={p.Id} - {ex.Message}"); }
                finally { p.Dispose(); }
            }

            Thread.Sleep(500);
            int final = GetProcessCount();
            if (final == 0)
                Log("STOP: force kill başarılı.");
            else
                Log($"STOP: UYARI - {final} process hala çalışıyor!");
        }

        private void SleepCancellable(int ms)
        {
            const int chunk = 500;
            int elapsed = 0;
            while (elapsed < ms && _running)
            {
                Thread.Sleep(Math.Min(chunk, ms - elapsed));
                elapsed += chunk;
            }
        }

        private void Log(string message)
        {
            string line = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] {message}";
            try
            {
                lock (_logLock)
                    File.AppendAllText(_logPath, line + Environment.NewLine,
                        System.Text.Encoding.UTF8);
            }
            catch { }
        }
    }
}
