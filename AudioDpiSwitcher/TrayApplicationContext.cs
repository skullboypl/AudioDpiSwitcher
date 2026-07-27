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

    private SettingsForm? _settingsForm;

    public TrayApplicationContext()
    {
        _settings = AppSettings.Load();
        I18n.SetLanguage(_settings.Lang);

        _menu = new ContextMenuStrip();
        BuildMenu();

        Icon? icon = null;
        var iconPath = Path.Combine(AppContext.BaseDirectory, "Assets", "icon.ico");
        if (File.Exists(iconPath))
            icon = new Icon(iconPath);

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
    }

    private void BuildMenu()
    {
        _menu.Items.Clear();
        _menu.Items.Add(new ToolStripMenuItem(I18n.T("set_defaults_tray"), null, (_, _) => CaptureCurrentAsDefaults()));
        _menu.Items.Add(new ToolStripSeparator());
        _menu.Items.Add(new ToolStripMenuItem(I18n.T("settings"), null, (_, _) => OpenSettings()));
        _menu.Items.Add(new ToolStripMenuItem(I18n.T("windows_settings"), null, (_, _) => OpenWindowsSound()));
        _menu.Items.Add(new ToolStripSeparator());
        _menu.Items.Add(new ToolStripMenuItem(I18n.T("exit"), null, (_, _) => ExitApp()));
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
            _trayIcon.Dispose();
            _audio.Dispose();
        }

        base.Dispose(disposing);
    }
}
