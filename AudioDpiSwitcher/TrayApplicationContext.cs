using System.Diagnostics;
using AudioDpiSwitcher.Services;
using NAudio.CoreAudioApi;

namespace AudioDpiSwitcher;

internal sealed class TrayApplicationContext : ApplicationContext
{
    private readonly NotifyIcon _trayIcon;
    private readonly ContextMenuStrip _menu;
    private readonly AppSettings _settings;
    private readonly AudioService _audio = new();
    private readonly System.Windows.Forms.Timer _restoreTimer;
    private readonly System.Windows.Forms.Timer _updateTimer;
    private int _updateCheckRunning;

    private SettingsForm? _settingsForm;

    public TrayApplicationContext(bool openSettingsOnStart = false)
    {
        _settings = AppSettings.Load();
        I18n.SetLanguage(_settings.Lang);

        _menu = new ContextMenuStrip
        {
            Font = new Font("Segoe UI", 9.5f),
            ShowImageMargin = true,
            Padding = new Padding(2, 4, 2, 4),
        };
        BuildMenu();

        Icon? icon = null;
        try
        {
            var exePath = Environment.ProcessPath ?? Application.ExecutablePath;
            if (!string.IsNullOrWhiteSpace(exePath))
                icon = Icon.ExtractAssociatedIcon(exePath);
        }
        catch
        {
            // fallback below
        }

        _trayIcon = new NotifyIcon
        {
            Visible = true,
            Text = TruncateTip(I18n.T("app_name")),
            Icon = icon ?? SystemIcons.Application,
            ContextMenuStrip = _menu,
        };
        _trayIcon.MouseClick += (_, e) =>
        {
            if (e.Button == MouseButtons.Left)
                OpenSettings();
        };

        _restoreTimer = new System.Windows.Forms.Timer { Interval = 2500 };
        _restoreTimer.Tick += (_, _) => TryAutoRestore();
        _restoreTimer.Start();

        _updateTimer = new System.Windows.Forms.Timer { Interval = 8000 };
        _updateTimer.Tick += async (_, _) =>
        {
            _updateTimer.Interval = 6 * 60 * 60 * 1000;
            await CheckForUpdatesAsync(manual: false, install: true);
        };
        _updateTimer.Start();

        // Warm audio cache shortly after startup (STA-safe)
        var warm = new System.Windows.Forms.Timer { Interval = 100 };
        warm.Tick += (_, _) =>
        {
            warm.Stop();
            warm.Dispose();
            try { _audio.CaptureSnapshot(); }
            catch { /* ignore */ }
        };
        warm.Start();

        if (openSettingsOnStart)
        {
            var open = new System.Windows.Forms.Timer { Interval = 200 };
            open.Tick += (_, _) =>
            {
                open.Stop();
                open.Dispose();
                OpenSettings();
            };
            open.Start();
        }
    }

    private void BuildMenu()
    {
        _menu.Items.Clear();
        _menu.Items.Add(Item("settings", "\uE713", (_, _) => OpenSettings()));
        _menu.Items.Add(Item("windows_settings", "\uE767", (_, _) => OpenWindowsSound()));
        _menu.Items.Add(new ToolStripSeparator());
        _menu.Items.Add(Item("set_defaults_tray", "\uE74E", (_, _) => CaptureCurrentAsDefaults()));
        _menu.Items.Add(new ToolStripSeparator());
        _menu.Items.Add(Item("exit", "\uE7E8", (_, _) => ExitApp()));
    }

    private static ToolStripMenuItem Item(string key, string glyph, EventHandler onClick)
    {
        return new ToolStripMenuItem(I18n.T(key), Glyph(glyph), onClick)
        {
            Padding = new Padding(4, 3, 8, 3),
        };
    }

    private static Image Glyph(string glyph)
    {
        var bmp = new Bitmap(16, 16);
        using var g = Graphics.FromImage(bmp);
        g.Clear(Color.Transparent);
        g.TextRenderingHint = System.Drawing.Text.TextRenderingHint.AntiAliasGridFit;
        using var font = new Font("Segoe MDL2 Assets", 11f, FontStyle.Regular, GraphicsUnit.Pixel);
        using var brush = new SolidBrush(Color.FromArgb(32, 32, 32));
        g.DrawString(glyph, font, brush, new PointF(-1, 0));
        return bmp;
    }

    private void CaptureCurrentAsDefaults()
    {
        try
        {
            var snapshot = _audio.CaptureSnapshot(forceRefresh: true);
            var defs = snapshot.Defaults;

            _settings.PreferredOutput = defs.GetValueOrDefault("output");
            _settings.PreferredOutputComm = defs.GetValueOrDefault("output_comm");
            _settings.PreferredInput = defs.GetValueOrDefault("input");
            _settings.PreferredInputComm = defs.GetValueOrDefault("input_comm");
            _settings.Save();

            // Keep open settings form in sync if visible
            if (_settingsForm is { IsDisposed: false, Visible: true })
                _settingsForm.ReloadFromPreferred();
        }
        catch
        {
            // ignore transient COM errors
        }
    }

    public void ApplySettings(AppSettings settings)
    {
        _settings.Lang = settings.Lang;
        _settings.AutoRestoreAudio = settings.AutoRestoreAudio;
        _settings.AutoUpdate = settings.AutoUpdate;
        _settings.PreferredOutput = settings.PreferredOutput;
        _settings.PreferredOutputComm = settings.PreferredOutputComm;
        _settings.PreferredInput = settings.PreferredInput;
        _settings.PreferredInputComm = settings.PreferredInputComm;
        _settings.Save();

        I18n.SetLanguage(_settings.Lang);
        _trayIcon.Text = TruncateTip(I18n.T("app_name"));
        BuildMenu();
    }

    private void OpenSettings()
    {
        if (_settingsForm is { IsDisposed: false })
        {
            _settingsForm.Show();
            _settingsForm.Activate();
            return;
        }

        _settingsForm = new SettingsForm(_settings, _audio, this);
        _settingsForm.FormClosed += (_, _) => _settingsForm = null;
        _settingsForm.Show();
    }

    public async Task CheckForUpdatesAsync(bool manual, bool install = true)
    {
        if (Interlocked.Exchange(ref _updateCheckRunning, 1) == 1)
            return;

        try
        {
            if (!manual && !_settings.AutoUpdate)
                return;

            UpdateOffer? offer;
            try
            {
                offer = await UpdateService.CheckAsync().ConfigureAwait(true);
            }
            catch
            {
                SetUpdateStatus(I18n.T("update_fail"));
                return;
            }

            if (offer is null)
            {
                SetUpdateStatus(I18n.T("update_fail"));
                return;
            }

            var current = UpdateService.CurrentVersion.ToString(3);
            var latest = offer.Version.ToString(3);
            if (!UpdateService.IsNewer(offer))
            {
                SetUpdateStatus(I18n.T("update_ok", ("version", current)));
                return;
            }

            SetUpdateStatus(I18n.T("update_status", ("current", current), ("latest", latest)));
            if (!install || (!manual && !_settings.AutoUpdate))
                return;

            if (!UpdateService.CanReplaceRunningExe)
            {
                SetUpdateStatus(I18n.T("update_status", ("current", current), ("latest", latest))
                    + " " + I18n.T("update_dev"));
                return;
            }

            SetUpdateStatus(I18n.T("update_downloading", ("version", latest)));
            var downloaded = await UpdateService.DownloadAsync(offer).ConfigureAwait(true);
            _trayIcon.ShowBalloonTip(4000, I18n.T("app_name"), I18n.T("update_restarting", ("version", latest)), ToolTipIcon.Info);
            UpdateService.LaunchReplacer(downloaded);
            var exit = new System.Windows.Forms.Timer { Interval = 600 };
            exit.Tick += (_, _) =>
            {
                exit.Stop();
                exit.Dispose();
                ExitApp();
            };
            exit.Start();
        }
        finally
        {
            Interlocked.Exchange(ref _updateCheckRunning, 0);
        }
    }

    private void SetUpdateStatus(string text)
    {
        if (_settingsForm is { IsDisposed: false })
            _settingsForm.SetUpdateStatus(text);
    }

    private static void OpenWindowsSound()
    {
        try
        {
            Process.Start(new ProcessStartInfo
            {
                FileName = "mmsys.cpl",
                UseShellExecute = true,
            });
        }
        catch
        {
            try
            {
                Process.Start(new ProcessStartInfo
                {
                    FileName = "control.exe",
                    Arguments = "mmsys.cpl",
                    UseShellExecute = true,
                });
            }
            catch
            {
                // ignore
            }
        }
    }

    private void TryAutoRestore()
    {
        if (!_settings.AutoRestoreAudio)
            return;

        try
        {
            var defs = _audio.GetDefaults();
            RestoreOne(defs.GetValueOrDefault("output"), _settings.PreferredOutput, AudioRoleKind.Console, DataFlow.Render);
            RestoreOne(defs.GetValueOrDefault("output_comm"), _settings.PreferredOutputComm, AudioRoleKind.Communications, DataFlow.Render);
            RestoreOne(defs.GetValueOrDefault("input"), _settings.PreferredInput, AudioRoleKind.Console, DataFlow.Capture);
            RestoreOne(defs.GetValueOrDefault("input_comm"), _settings.PreferredInputComm, AudioRoleKind.Communications, DataFlow.Capture);
        }
        catch
        {
            // ignore transient COM errors
        }
    }

    private bool RestoreOne(string? currentId, string? preferredId, AudioRoleKind role, DataFlow flow)
    {
        if (string.IsNullOrWhiteSpace(preferredId))
            return false;

        if (string.Equals(currentId, preferredId, StringComparison.OrdinalIgnoreCase))
            return false;

        var pool = flow == DataFlow.Render ? _audio.ListPlayback() : _audio.ListRecording();
        if (!pool.Any(d => string.Equals(d.Id, preferredId, StringComparison.OrdinalIgnoreCase)))
            return false;

        _audio.SetDefault(preferredId, role);
        return true;
    }

    private static string TruncateTip(string text) =>
        text.Length > 63 ? text[..63] : text;

    private void ExitApp()
    {
        _restoreTimer.Stop();
        _updateTimer.Stop();
        _trayIcon.Visible = false;
        _trayIcon.Dispose();
        _audio.Dispose();
        ExitThread();
    }

    protected override void Dispose(bool disposing)
    {
        if (disposing)
        {
            _restoreTimer.Dispose();
            _updateTimer.Dispose();
            _trayIcon.Dispose();
            _audio.Dispose();
        }

        base.Dispose(disposing);
    }
}
