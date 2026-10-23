using System.Diagnostics;
using AudioDpiSwitcher.Services;

namespace AudioDpiSwitcher;

internal sealed class SettingsForm : Form
{
    private readonly AppSettings _settings;
    private readonly AudioService _audio;
    private readonly TrayApplicationContext _tray;

    private readonly ComboBox _output = MakeDeviceCombo();
    private readonly ComboBox _outputComm = MakeDeviceCombo();
    private readonly ComboBox _input = MakeDeviceCombo();
    private readonly ComboBox _inputComm = MakeDeviceCombo();
    private readonly CheckBox _autoRestore = new()
    {
        AutoSize = true,
        MaximumSize = new Size(400, 0),
    };
    private Button _useCurrentBtn = null!;
    private Button _refreshBtn = null!;
    private Button _setDpiBtn = null!;
    private Button _applyBtn = null!;
    private Button _closeBtn = null!;
    private Label _dpiLabel = null!;
    private GroupBox _audioGroup = null!;
    private GroupBox _dpiGroup = null!;
    private GroupBox _langGroup = null!;
    private LinkLabel _footer = null!;

    private readonly ComboBox _lang = new() { DropDownStyle = ComboBoxStyle.DropDownList, Width = 220 };
    private readonly ComboBox _monitor = new() { DropDownStyle = ComboBoxStyle.DropDownList, Width = 400 };
    private readonly ComboBox _dpiPreset = new() { DropDownStyle = ComboBoxStyle.DropDownList, Width = 120 };

    private List<DpiDisplay> _displays = [];
    private bool _loading;
    private bool _dataLoaded;

    public SettingsForm(AppSettings settings, AudioService audio, TrayApplicationContext tray)
    {
        _settings = settings;
        _audio = audio;
        _tray = tray;

        FormBorderStyle = FormBorderStyle.FixedDialog;
        MaximizeBox = false;
        MinimizeBox = false;
        StartPosition = FormStartPosition.CenterScreen;
        ClientSize = new Size(500, 668);
        Font = new Font("Segoe UI", 9f);
        BackColor = Color.FromArgb(243, 243, 243);
        ShowInTaskbar = true;

        BuildUi();
        Localize();
        LoadFromSettings();

        // Show empty shell first; fill devices/monitors after paint
        Shown += async (_, _) => await EnsureDataLoadedAsync();
    }

    protected override void OnFormClosing(FormClosingEventArgs e)
    {
        // Keep instance warm — next open is instant
        if (e.CloseReason == CloseReason.UserClosing)
        {
            e.Cancel = true;
            Hide();
            return;
        }

        base.OnFormClosing(e);
    }

    private static ComboBox MakeDeviceCombo() => new()
    {
        DropDownStyle = ComboBoxStyle.DropDownList,
        Width = 444,
        IntegralHeight = false,
        MaxDropDownItems = 12,
    };

    private void BuildUi()
    {
        const int left = 16;
        const int width = 468;

        var audioGroup = new GroupBox
        {
            Location = new Point(left, 12),
            Size = new Size(width, 372),
            Padding = new Padding(12),
        };
        audioGroup.Tag = "group_audio";

        var hint = new Label
        {
            AutoSize = false,
            Location = new Point(14, 24),
            Size = new Size(440, 32),
            ForeColor = Color.FromArgb(80, 80, 80),
            Tag = "audio_hint",
        };
        audioGroup.Controls.Add(hint);

        var y = 58;
        AddLabeledCombo(audioGroup, "label_output", _output, ref y);
        AddLabeledCombo(audioGroup, "label_output_comm", _outputComm, ref y);
        AddLabeledCombo(audioGroup, "label_input", _input, ref y);
        AddLabeledCombo(audioGroup, "label_input_comm", _inputComm, ref y);

        _autoRestore.Location = new Point(14, y);
        _autoRestore.MaximumSize = new Size(440, 0);
        _autoRestore.Tag = "auto_restore";
        audioGroup.Controls.Add(_autoRestore);
        _audioGroup = audioGroup;

        _useCurrentBtn = new Button
        {
            Size = new Size(250, 30),
            Tag = "use_current_defaults",
        };
        _useCurrentBtn.Click += (_, _) => UseCurrentAsDefaults();
        audioGroup.Controls.Add(_useCurrentBtn);

        _refreshBtn = new Button
        {
            Size = new Size(120, 30),
            Tag = "refresh_devices",
        };
        _refreshBtn.Click += (_, _) => _ = ReloadDevicesAsync(forceRefresh: true);
        audioGroup.Controls.Add(_refreshBtn);
        LayoutAudioActionRow();

        var dpiTop = audioGroup.Bottom + 10;
        var dpiGroup = new GroupBox
        {
            Location = new Point(left, dpiTop),
            Size = new Size(width, 128),
            Tag = "group_dpi",
        };

        var monLabel = new Label { AutoSize = true, Location = new Point(14, 24), Tag = "monitor" };
        _monitor.Location = new Point(14, 44);
        _monitor.Width = 440;
        _monitor.SelectedIndexChanged += OnMonitorSelectedHandler;

        _dpiLabel = new Label { AutoSize = true, Location = new Point(14, 84), Tag = "dpi_preset" };
        _dpiPreset.Width = 110;
        foreach (var p in DpiService.Presets)
            _dpiPreset.Items.Add($"{p}%");

        _setDpiBtn = new Button
        {
            Size = new Size(180, 28),
            Tag = "set_dpi",
        };
        _setDpiBtn.Click += (_, _) => ApplyDpi();

        dpiGroup.Controls.Add(monLabel);
        dpiGroup.Controls.Add(_monitor);
        dpiGroup.Controls.Add(_dpiLabel);
        dpiGroup.Controls.Add(_dpiPreset);
        dpiGroup.Controls.Add(_setDpiBtn);
        _dpiGroup = dpiGroup;

        var langTop = dpiGroup.Bottom + 10;
        var langGroup = new GroupBox
        {
            Location = new Point(left, langTop),
            Size = new Size(width, 64),
            Tag = "group_language",
        };
        _lang.Location = new Point(14, 26);
        _lang.Width = 220;
        langGroup.Controls.Add(_lang);
        _langGroup = langGroup;

        _applyBtn = new Button
        {
            Size = new Size(88, 30),
            Tag = "apply",
        };
        _applyBtn.Click += (_, _) => ApplyAll();

        _closeBtn = new Button
        {
            Size = new Size(88, 30),
            Tag = "close",
        };
        _closeBtn.Click += (_, _) => Close();

        _footer = new LinkLabel
        {
            AutoSize = true,
            LinkColor = Color.FromArgb(0, 102, 204),
            ActiveLinkColor = Color.FromArgb(0, 70, 160),
            Text = $"© {DateTime.Now.Year} skullboypl",
        };
        var nameStart = _footer.Text.IndexOf("skullboypl", StringComparison.Ordinal);
        _footer.LinkArea = new LinkArea(nameStart, "skullboypl".Length);
        _footer.LinkClicked += (_, _) =>
        {
            try
            {
                Process.Start(new ProcessStartInfo
                {
                    FileName = "https://github.com/skullboypl",
                    UseShellExecute = true,
                });
            }
            catch
            {
                // ignore
            }
        };

        Controls.Add(audioGroup);
        Controls.Add(dpiGroup);
        Controls.Add(langGroup);
        Controls.Add(_applyBtn);
        Controls.Add(_closeBtn);
        Controls.Add(_footer);

        Reflow();
        AcceptButton = _applyBtn;
        CancelButton = _closeBtn;
    }

    private void LayoutAudioActionRow()
    {
        var buttonsY = _autoRestore.Bottom + 10;
        _useCurrentBtn.Location = new Point(14, buttonsY);
        _refreshBtn.Location = new Point(294, buttonsY);

        _refreshBtn.Left = _audioGroup.Width - 14 - _refreshBtn.Width;
        _useCurrentBtn.Width = Math.Max(160, _refreshBtn.Left - 24 - 14);
        _audioGroup.Height = buttonsY + _useCurrentBtn.Height + 16;
    }

    private void Reflow()
    {
        const int left = 16;
        LayoutAudioActionRow();

        _dpiGroup.Top = _audioGroup.Bottom + 10;
        _dpiLabel.Location = new Point(14, 84);
        _dpiPreset.Location = new Point(_dpiLabel.Right + 12, 78);
        _setDpiBtn.Location = new Point(_dpiPreset.Right + 12, 78);
        var dpiBottom = Math.Max(_monitor.Bottom, _setDpiBtn.Bottom) + 16;
        if (dpiBottom > 48)
            _dpiGroup.Height = dpiBottom;

        _langGroup.Top = _dpiGroup.Bottom + 10;
        var buttonTop = _langGroup.Bottom + 14;
        _closeBtn.Location = new Point(_dpiGroup.Right - _closeBtn.Width, buttonTop);
        _applyBtn.Location = new Point(_closeBtn.Left - 8 - _applyBtn.Width, buttonTop);
        _footer.Location = new Point(left, _closeBtn.Bottom + 14);
        ClientSize = new Size(_audioGroup.Right + left, _footer.Bottom + 16);
    }

    private static void AddLabeledCombo(Control parent, string tag, ComboBox combo, ref int y)
    {
        var label = new Label { AutoSize = true, Location = new Point(14, y), Tag = tag };
        y += 16;
        combo.Location = new Point(14, y);
        y += 32;
        parent.Controls.Add(label);
        parent.Controls.Add(combo);
    }

    private void Localize()
    {
        Text = I18n.T("settings_title");
        foreach (Control c in Controls)
            LocalizeControl(c);

        // Refresh language combo labels while preserving selection
        var selected = _lang.SelectedItem is LangItem cur ? cur.Code : (_settings.Lang ?? "auto");
        _lang.Items.Clear();
        _lang.Items.Add(new LangItem("auto", I18n.T("lang_auto")));
        _lang.Items.Add(new LangItem("en", I18n.T("lang_en")));
        _lang.Items.Add(new LangItem("pl", I18n.T("lang_pl")));
        _lang.Items.Add(new LangItem("de", I18n.T("lang_de")));
        _lang.DisplayMember = nameof(LangItem.Label);
        for (var i = 0; i < _lang.Items.Count; i++)
        {
            if (_lang.Items[i] is LangItem li &&
                string.Equals(li.Code, selected, StringComparison.OrdinalIgnoreCase))
            {
                _lang.SelectedIndex = i;
                break;
            }
        }

        Reflow();
    }

    private void LocalizeControl(Control control)
    {
        if (control.Tag is string key)
            control.Text = I18n.T(key);

        foreach (Control child in control.Controls)
            LocalizeControl(child);
    }

    public void ReloadFromPreferred()
    {
        if (IsDisposed)
            return;

        _ = ReloadDevicesAsync(forceRefresh: false);
        LoadFromSettings();
    }

    private async Task EnsureDataLoadedAsync()
    {
        // Re-load if first attempt produced empty DPI list (e.g. after API fix)
        if (_dataLoaded && _displays.Count > 0)
            return;

        await Task.Yield();

        UseWaitCursor = true;
        Enabled = false;
        try
        {
            await ReloadDevicesAsync(forceRefresh: false);
            await LoadDisplaysAsync();
            _dataLoaded = _displays.Count > 0;
        }
        finally
        {
            Enabled = true;
            UseWaitCursor = false;
        }
    }

    private async Task ReloadDevicesAsync(bool forceRefresh)
    {
        _loading = true;
        try
        {
            // Yield so UI stays responsive; NAudio COM must stay on STA/UI thread
            await Task.Yield();
            var snapshot = _audio.CaptureSnapshot(forceRefresh);
            FillDeviceCombo(_output, snapshot.Playback, _settings.PreferredOutput);
            FillDeviceCombo(_outputComm, snapshot.Playback, _settings.PreferredOutputComm);
            FillDeviceCombo(_input, snapshot.Recording, _settings.PreferredInput);
            FillDeviceCombo(_inputComm, snapshot.Recording, _settings.PreferredInputComm);
        }
        finally
        {
            _loading = false;
        }
    }

    private async Task LoadDisplaysAsync(string? selectKey = null)
    {
        var displays = await Task.Run(() => DpiService.ListDisplays().ToList()).ConfigureAwait(true);
        _displays = displays;

        _monitor.BeginUpdate();
        try
        {
            _monitor.SelectedIndexChanged -= OnMonitorSelectedHandler;
            _monitor.Items.Clear();
            foreach (var d in _displays)
            {
                var res = d.Width > 0 && d.Height > 0 ? $"{d.Width}×{d.Height}" : "?×?";
                _monitor.Items.Add($"{d.Index}. {d.Name} - {res} ({d.CurrentDpi}%)");
            }
        }
        finally
        {
            _monitor.EndUpdate();
            _monitor.SelectedIndexChanged += OnMonitorSelectedHandler;
        }

        var select = 0;
        if (!string.IsNullOrEmpty(selectKey))
        {
            for (var i = 0; i < _displays.Count; i++)
            {
                if (_displays[i].StableKey == selectKey)
                {
                    select = i;
                    break;
                }
            }
        }

        if (_monitor.Items.Count > 0)
        {
            _monitor.SelectedIndex = select;
            OnMonitorSelected();
        }
    }

    private void OnMonitorSelectedHandler(object? sender, EventArgs e) => OnMonitorSelected();

    private void OnMonitorSelected()
    {
        if (_monitor.SelectedIndex < 0 || _monitor.SelectedIndex >= _displays.Count)
            return;

        SelectDpiPreset((int)_displays[_monitor.SelectedIndex].CurrentDpi);
    }

    private void SelectDpiPreset(int pct)
    {
        var label = $"{pct}%";
        var i = _dpiPreset.Items.IndexOf(label);
        if (i >= 0)
            _dpiPreset.SelectedIndex = i;
        else if (_dpiPreset.Items.Count > 0 && _dpiPreset.SelectedIndex < 0)
            _dpiPreset.SelectedIndex = 0;
    }

    private void UseCurrentAsDefaults()
    {
        _ = UseCurrentAsDefaultsAsync();
    }

    private async Task UseCurrentAsDefaultsAsync()
    {
        _loading = true;
        UseWaitCursor = true;
        try
        {
            await Task.Yield();
            var snapshot = _audio.CaptureSnapshot(forceRefresh: true);
            var defs = snapshot.Defaults;

            FillDeviceCombo(_output, snapshot.Playback, defs.GetValueOrDefault("output"));
            FillDeviceCombo(_outputComm, snapshot.Playback, defs.GetValueOrDefault("output_comm"));
            FillDeviceCombo(_input, snapshot.Recording, defs.GetValueOrDefault("input"));
            FillDeviceCombo(_inputComm, snapshot.Recording, defs.GetValueOrDefault("input_comm"));

            _settings.PreferredOutput = defs.GetValueOrDefault("output");
            _settings.PreferredOutputComm = defs.GetValueOrDefault("output_comm");
            _settings.PreferredInput = defs.GetValueOrDefault("input");
            _settings.PreferredInputComm = defs.GetValueOrDefault("input_comm");
            _settings.Save();
            _tray.ApplySettings(_settings);
        }
        finally
        {
            _loading = false;
            UseWaitCursor = false;
        }
    }

    private void FillDeviceCombo(ComboBox combo, IReadOnlyList<AudioDeviceInfo> devices, string? selectedId)
    {
        combo.BeginUpdate();
        try
        {
            combo.Items.Clear();
            combo.DisplayMember = nameof(DeviceItem.Name);
            combo.ValueMember = nameof(DeviceItem.Id);
            combo.Items.Add(new DeviceItem("", I18n.T("none")));
            var select = 0;
            for (var i = 0; i < devices.Count; i++)
            {
                var d = devices[i];
                combo.Items.Add(new DeviceItem(d.Id, d.Name));
                if (!string.IsNullOrWhiteSpace(selectedId) &&
                    string.Equals(d.Id, selectedId, StringComparison.OrdinalIgnoreCase))
                    select = i + 1;
            }

            combo.SelectedIndex = select;
        }
        finally
        {
            combo.EndUpdate();
        }
    }

    private void LoadFromSettings()
    {
        _autoRestore.Checked = _settings.AutoRestoreAudio;
        var lang = (_settings.Lang ?? "auto").ToLowerInvariant();
        for (var i = 0; i < _lang.Items.Count; i++)
        {
            if (_lang.Items[i] is LangItem li && li.Code == lang)
            {
                _lang.SelectedIndex = i;
                return;
            }
        }

        if (_lang.Items.Count > 0)
            _lang.SelectedIndex = 0;
    }

    private void ApplyDpi()
    {
        if (_monitor.SelectedIndex < 0 || _monitor.SelectedIndex >= _displays.Count ||
            _dpiPreset.SelectedItem is null)
            return;

        var pctText = _dpiPreset.SelectedItem.ToString()!.TrimEnd('%');
        if (!int.TryParse(pctText, out var pct))
            return;

        var display = _displays[_monitor.SelectedIndex];
        var key = display.StableKey;
        if (!DpiService.SetDpi(display, pct))
            return;

        _ = LoadDisplaysAsync(selectKey: key);
    }

    private void ApplyAll()
    {
        if (_loading)
            return;

        _settings.AutoRestoreAudio = _autoRestore.Checked;
        _settings.Lang = _lang.SelectedItem is LangItem langItem ? langItem.Code : "auto";
        _settings.PreferredOutput = SelectedId(_output);
        _settings.PreferredOutputComm = SelectedId(_outputComm);
        _settings.PreferredInput = SelectedId(_input);
        _settings.PreferredInputComm = SelectedId(_inputComm);
        _settings.Save();

        if (!string.IsNullOrWhiteSpace(_settings.PreferredOutput))
            _audio.SetDefault(_settings.PreferredOutput, AudioRoleKind.Console);
        if (!string.IsNullOrWhiteSpace(_settings.PreferredOutputComm))
            _audio.SetDefault(_settings.PreferredOutputComm, AudioRoleKind.Communications);
        if (!string.IsNullOrWhiteSpace(_settings.PreferredInput))
            _audio.SetDefault(_settings.PreferredInput, AudioRoleKind.Console);
        if (!string.IsNullOrWhiteSpace(_settings.PreferredInputComm))
            _audio.SetDefault(_settings.PreferredInputComm, AudioRoleKind.Communications);

        _tray.ApplySettings(_settings);
        Localize();
        _ = ReloadDevicesAsync(forceRefresh: true);
    }

    private static string? SelectedId(ComboBox combo)
    {
        if (combo.SelectedItem is DeviceItem item)
            return string.IsNullOrWhiteSpace(item.Id) ? null : item.Id;
        return null;
    }

    private sealed record DeviceItem(string Id, string Name)
    {
        public override string ToString() => Name;
    }

    private sealed record LangItem(string Code, string Label)
    {
        public override string ToString() => Label;
    }
}
