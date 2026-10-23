using System.Diagnostics;
using AudioDpiSwitcher.Services;

namespace AudioDpiSwitcher;

internal sealed class SettingsForm : Form
{
    private readonly AppSettings _settings;
    private readonly AudioService _audio;
    private readonly TrayApplicationContext _tray;

    private readonly ComboBox _output = MakeCombo();
    private readonly ComboBox _outputComm = MakeCombo();
    private readonly ComboBox _input = MakeCombo();
    private readonly ComboBox _inputComm = MakeCombo();
    private readonly CheckBox _autoRestore = new() { AutoSize = true };
    private readonly CheckBox _autoUpdate = new() { AutoSize = true, MaximumSize = new Size(640, 0) };

    private readonly Label _curOutput = MakeValue();
    private readonly Label _curOutputComm = MakeValue();
    private readonly Label _curInput = MakeValue();
    private readonly Label _curInputComm = MakeValue();

    private readonly ComboBox _lang = MakeCombo();
    private readonly ComboBox _monitor = MakeCombo();
    private readonly ComboBox _dpiPreset = MakeCombo();
    private readonly Label _updateStatus = new() { AutoSize = true, MaximumSize = new Size(640, 0) };

    private readonly SettingsStack _pagePreferred = new();
    private readonly SettingsStack _pageCurrent = new();
    private readonly SettingsStack _pageDisplay = new();
    private readonly SettingsStack _pageApp = new();
    private readonly TabChip[] _tabButtons = new TabChip[4];
    private Panel _footerBar = null!;
    private int _pageIndex;
    private readonly Button _applyBtn = new() { AutoSize = true, MinimumSize = new Size(96, 32) };
    private readonly Button _closeBtn = new() { AutoSize = true, MinimumSize = new Size(96, 32) };
    private readonly Button _useCurrentBtn = new() { AutoSize = true, MinimumSize = new Size(180, 32) };
    private readonly Button _refreshCurrentBtn = new() { AutoSize = true, MinimumSize = new Size(110, 32) };
    private readonly Button _refreshPreferredBtn = new() { AutoSize = true, MinimumSize = new Size(110, 32) };
    private readonly Button _setDpiBtn = new() { AutoSize = true, MinimumSize = new Size(140, 32) };
    private readonly Button _checkUpdateBtn = new() { AutoSize = true, MinimumSize = new Size(160, 32) };
    private readonly LinkLabel _footer = new() { AutoSize = true };

    private List<DpiDisplay> _displays = [];
    private bool _loading;
    private bool _dataLoaded;

    public SettingsForm(AppSettings settings, AudioService audio, TrayApplicationContext tray)
    {
        _settings = settings;
        _audio = audio;
        _tray = tray;

        Font = new Font("Segoe UI", 9.5f);
        DoubleBuffered = true;
        BackColor = Color.White;
        StartPosition = FormStartPosition.CenterScreen;
        MinimumSize = new Size(520, 460);
        Size = new Size(640, 520);
        ShowInTaskbar = true;

        BuildUi();
        Localize();
        LoadFromSettings();
        Shown += async (_, _) => await EnsureDataLoadedAsync();
    }

    public void SetUpdateStatus(string text)
    {
        if (IsDisposed)
            return;
        _updateStatus.Text = text;
    }

    protected override void OnFormClosing(FormClosingEventArgs e)
    {
        if (e.CloseReason == CloseReason.UserClosing)
        {
            e.Cancel = true;
            Hide();
            return;
        }

        base.OnFormClosing(e);
    }

    protected override CreateParams CreateParams
    {
        get
        {
            const int composited = 0x02000000;
            var cp = base.CreateParams;
            cp.ExStyle |= composited;
            return cp;
        }
    }

    private void BuildUi()
    {
        FillPreferred(_pagePreferred);
        FillCurrent(_pageCurrent);
        FillDisplay(_pageDisplay);
        FillApp(_pageApp);

        var pages = new SettingsStack[] { _pagePreferred, _pageCurrent, _pageDisplay, _pageApp };
        var host = new Panel { Dock = DockStyle.Fill, BackColor = Color.White };
        for (var i = 0; i < pages.Length; i++)
        {
            pages[i].Dock = DockStyle.Fill;
            pages[i].Visible = i == 0;
            host.Controls.Add(pages[i]);
        }

        var keys = new[] { "tab_preferred", "tab_current", "tab_display", "tab_app" };
        var strip = new TabStrip();
        for (var i = 0; i < keys.Length; i++)
        {
            var index = i;
            var chip = new TabChip { Tag = keys[i], Selected = i == 0 };
            chip.Click += (_, _) => ShowPage(index);
            _tabButtons[i] = chip;
            strip.Controls.Add(chip);
        }

        _footerBar = new Panel
        {
            Dock = DockStyle.Bottom,
            Height = 56,
            BackColor = Color.White,
        };
        _footerBar.Paint += (_, e) =>
        {
            using var pen = new Pen(Color.FromArgb(229, 229, 229));
            e.Graphics.DrawLine(pen, 0, 0, _footerBar.Width, 0);
        };
        _footer.LinkColor = Color.FromArgb(0, 102, 204);
        _footer.ActiveLinkColor = Color.FromArgb(0, 70, 160);
        _footer.Text = $"© {DateTime.Now.Year} skullboypl";
        var nameStart = _footer.Text.IndexOf("skullboypl", StringComparison.Ordinal);
        _footer.LinkArea = new LinkArea(nameStart, "skullboypl".Length);
        _footer.Location = new Point(20, 18);
        _footer.LinkClicked += (_, _) => OpenGithub();
        StyleButton(_applyBtn);
        StyleButton(_closeBtn);
        _applyBtn.Tag = "apply";
        _closeBtn.Tag = "close";
        _applyBtn.Click += (_, _) => ApplyAll();
        _closeBtn.Click += (_, _) => Close();
        _footerBar.Controls.Add(_footer);
        _footerBar.Controls.Add(_applyBtn);
        _footerBar.Controls.Add(_closeBtn);
        _footerBar.Resize += (_, _) => PlaceBarButtons(_footerBar);

        Controls.Add(host);
        Controls.Add(strip);
        Controls.Add(_footerBar);
        AcceptButton = _applyBtn;
        CancelButton = _closeBtn;
        PlaceBarButtons(_footerBar);
    }

    private void ShowPage(int index)
    {
        if (index == _pageIndex)
            return;

        var pages = new SettingsStack[] { _pagePreferred, _pageCurrent, _pageDisplay, _pageApp };
        _pageIndex = index;
        SuspendLayout();
        for (var i = 0; i < pages.Length; i++)
        {
            var on = i == index;
            if (pages[i].Visible != on)
                pages[i].Visible = on;
            _tabButtons[i].Selected = on;
        }

        ResumeLayout(true);
    }

    private void FillPreferred(SettingsStack page)
    {
        page.Add(Hint("preferred_hint"), StackRow.Hint);
        AddField(page, "label_output", _output);
        AddField(page, "label_output_comm", _outputComm);
        AddField(page, "label_input", _input);
        AddField(page, "label_input_comm", _inputComm);
        _autoRestore.Tag = "auto_restore";
        page.Add(_autoRestore, StackRow.Check);
        StyleButton(_refreshPreferredBtn);
        _refreshPreferredBtn.Tag = "refresh_devices";
        _refreshPreferredBtn.Click += (_, _) => _ = ReloadDevicesAsync(forceRefresh: true);
        page.Add(ButtonRow(_refreshPreferredBtn), StackRow.Buttons);
    }

    private void FillCurrent(SettingsStack page)
    {
        page.Add(Hint("current_hint"), StackRow.Hint);
        AddReadout(page, "label_output", _curOutput);
        AddReadout(page, "label_output_comm", _curOutputComm);
        AddReadout(page, "label_input", _curInput);
        AddReadout(page, "label_input_comm", _curInputComm);
        StyleButton(_useCurrentBtn);
        StyleButton(_refreshCurrentBtn);
        _useCurrentBtn.Tag = "use_current_defaults";
        _refreshCurrentBtn.Tag = "refresh_devices";
        _useCurrentBtn.Click += (_, _) => UseCurrentAsDefaults();
        _refreshCurrentBtn.Click += (_, _) => _ = RefreshCurrentAsync();
        page.Add(ButtonRow(_useCurrentBtn, _refreshCurrentBtn), StackRow.Buttons);
    }

    private void FillDisplay(SettingsStack page)
    {
        page.Add(Hint("display_hint"), StackRow.Hint);
        AddField(page, "monitor", _monitor);
        _monitor.SelectedIndexChanged += OnMonitorSelectedHandler;
        foreach (var preset in DpiService.Presets)
            _dpiPreset.Items.Add($"{preset}%");
        page.Add(Caption("dpi_preset"), StackRow.Caption);
        StyleButton(_setDpiBtn);
        _setDpiBtn.Tag = "set_dpi";
        _setDpiBtn.Click += (_, _) => ApplyDpi();
        page.Add(ScaleRow(_dpiPreset, _setDpiBtn), StackRow.Field);
    }

    private void FillApp(SettingsStack page)
    {
        page.Add(Hint("app_hint"), StackRow.Hint);
        AddField(page, "group_language", _lang);
        _autoUpdate.Tag = "auto_update";
        page.Add(_autoUpdate, StackRow.Check);
        _updateStatus.ForeColor = Color.FromArgb(80, 80, 80);
        page.Add(_updateStatus, StackRow.Text);
        StyleButton(_checkUpdateBtn);
        _checkUpdateBtn.Tag = "check_update";
        _checkUpdateBtn.Click += (_, _) => _ = _tray.CheckForUpdatesAsync(manual: true);
        page.Add(ButtonRow(_checkUpdateBtn), StackRow.Buttons);
    }

    private static void AddField(SettingsStack page, string tag, ComboBox combo)
    {
        page.Add(Caption(tag), StackRow.Caption);
        page.Add(combo, StackRow.Field);
    }

    private static void AddReadout(SettingsStack page, string tag, Label value)
    {
        page.Add(Caption(tag), StackRow.Caption);
        page.Add(value, StackRow.Text);
    }

    private static Label Hint(string tag) => new()
    {
        AutoSize = true,
        Tag = tag,
        ForeColor = Color.FromArgb(90, 90, 90),
    };

    private static Label Caption(string tag) => new()
    {
        AutoSize = true,
        Tag = tag,
        ForeColor = Color.FromArgb(70, 70, 70),
    };

    private static Label MakeValue() => new()
    {
        AutoSize = true,
        Font = new Font("Segoe UI", 10f),
        Text = "...",
    };

    private static ComboBox MakeCombo() => new()
    {
        DropDownStyle = ComboBoxStyle.DropDownList,
        IntegralHeight = false,
        MaxDropDownItems = 14,
    };

    private static void StyleButton(Button button)
    {
        button.AutoSize = true;
        button.MinimumSize = new Size(96, 32);
        button.Padding = new Padding(10, 4, 10, 4);
        button.FlatStyle = FlatStyle.System;
    }

    private static Panel ButtonRow(params Button[] buttons)
    {
        var row = new Panel { Height = 36 };
        var x = 0;
        foreach (var button in buttons)
        {
            button.Location = new Point(x, 0);
            row.Controls.Add(button);
            x += button.PreferredSize.Width + 8;
        }

        return row;
    }

    private static Panel ScaleRow(ComboBox preset, Button apply)
    {
        var row = new Panel { Height = 32 };
        preset.SetBounds(0, 1, 120, 30);
        apply.Location = new Point(128, 0);
        row.Controls.Add(preset);
        row.Controls.Add(apply);
        return row;
    }

    private static void PlaceBarButtons(Panel bar)
    {
        var close = bar.Controls.OfType<Button>().First(b => Equals(b.Tag, "close"));
        var apply = bar.Controls.OfType<Button>().First(b => Equals(b.Tag, "apply"));
        close.Location = new Point(bar.ClientSize.Width - close.Width - 16, 12);
        apply.Location = new Point(close.Left - apply.Width - 8, 12);
    }

    private void Localize()
    {
        Text = I18n.T("settings_title");
        foreach (Control c in Controls)
            LocalizeControl(c);

        var selected = _lang.SelectedItem is LangItem cur ? cur.Code : (_settings.Lang ?? "auto");
        _lang.Items.Clear();
        _lang.Items.Add(new LangItem("auto", I18n.T("lang_auto")));
        _lang.Items.Add(new LangItem("en", I18n.T("lang_en")));
        _lang.Items.Add(new LangItem("pl", I18n.T("lang_pl")));
        _lang.Items.Add(new LangItem("de", I18n.T("lang_de")));
        for (var i = 0; i < _lang.Items.Count; i++)
        {
            if (_lang.Items[i] is LangItem li &&
                string.Equals(li.Code, selected, StringComparison.OrdinalIgnoreCase))
            {
                _lang.SelectedIndex = i;
                break;
            }
        }

        if (string.IsNullOrWhiteSpace(_updateStatus.Text))
        {
            _updateStatus.Text = I18n.T(
                "update_status",
                ("current", UpdateService.CurrentVersion.ToString(3)),
                ("latest", "..."));
        }

        PlaceBarButtons(_footerBar);
        PerformLayout();
    }

    private void LocalizeControl(Control control)
    {
        if (control.Tag is string key && control is not ComboBox)
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
        if (_dataLoaded && _displays.Count > 0)
            return;

        await Task.Yield();
        UseWaitCursor = true;
        try
        {
            await ReloadDevicesAsync(forceRefresh: false);
            await LoadDisplaysAsync();
            _dataLoaded = _displays.Count > 0;
        }
        finally
        {
            UseWaitCursor = false;
        }
    }

    private async Task ReloadDevicesAsync(bool forceRefresh)
    {
        _loading = true;
        try
        {
            await Task.Yield();
            var snapshot = _audio.CaptureSnapshot(forceRefresh);
            FillDeviceCombo(_output, snapshot.Playback, _settings.PreferredOutput);
            FillDeviceCombo(_outputComm, snapshot.Playback, _settings.PreferredOutputComm);
            FillDeviceCombo(_input, snapshot.Recording, _settings.PreferredInput);
            FillDeviceCombo(_inputComm, snapshot.Recording, _settings.PreferredInputComm);
            ShowCurrent(snapshot);
        }
        finally
        {
            _loading = false;
        }
    }

    private async Task RefreshCurrentAsync()
    {
        await Task.Yield();
        var snapshot = _audio.CaptureSnapshot(forceRefresh: true);
        ShowCurrent(snapshot);
    }

    private void ShowCurrent(AudioSnapshot snapshot)
    {
        _curOutput.Text = _audio.FindName(snapshot.Playback, snapshot.Defaults.GetValueOrDefault("output"));
        _curOutputComm.Text = _audio.FindName(snapshot.Playback, snapshot.Defaults.GetValueOrDefault("output_comm"));
        _curInput.Text = _audio.FindName(snapshot.Recording, snapshot.Defaults.GetValueOrDefault("input"));
        _curInputComm.Text = _audio.FindName(snapshot.Recording, snapshot.Defaults.GetValueOrDefault("input_comm"));
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
            foreach (var display in _displays)
            {
                var res = display.Width > 0 && display.Height > 0
                    ? $"{display.Width}x{display.Height}"
                    : "?";
                _monitor.Items.Add($"{display.Index}. {display.Name} - {res} ({display.CurrentDpi}%)");
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
            ShowCurrent(snapshot);

            _settings.PreferredOutput = defs.GetValueOrDefault("output");
            _settings.PreferredOutputComm = defs.GetValueOrDefault("output_comm");
            _settings.PreferredInput = defs.GetValueOrDefault("input");
            _settings.PreferredInputComm = defs.GetValueOrDefault("input_comm");
            _settings.Save();
            _tray.ApplySettings(_settings);
            ShowPage(0);
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
            combo.Items.Add(new DeviceItem("", I18n.T("none")));
            var select = 0;
            for (var i = 0; i < devices.Count; i++)
            {
                var device = devices[i];
                combo.Items.Add(new DeviceItem(device.Id, device.Name));
                if (!string.IsNullOrWhiteSpace(selectedId) &&
                    string.Equals(device.Id, selectedId, StringComparison.OrdinalIgnoreCase))
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
        _autoUpdate.Checked = _settings.AutoUpdate;
        var lang = (_settings.Lang ?? "auto").ToLowerInvariant();
        for (var i = 0; i < _lang.Items.Count; i++)
        {
            if (_lang.Items[i] is LangItem item && item.Code == lang)
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
        if (!DpiService.SetDpi(display, pct))
            return;

        _ = LoadDisplaysAsync(display.StableKey);
    }

    private void ApplyAll()
    {
        if (_loading)
            return;

        _settings.AutoRestoreAudio = _autoRestore.Checked;
        _settings.AutoUpdate = _autoUpdate.Checked;
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

    private static void OpenGithub()
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
    }

    private sealed record DeviceItem(string Id, string Name)
    {
        public override string ToString() => Name;
    }

    private sealed record LangItem(string Code, string Label)
    {
        public override string ToString() => Label;
    }

    private enum StackRow
    {
        Hint,
        Caption,
        Text,
        Field,
        Check,
        Buttons,
    }

    private sealed class SettingsStack : Panel
    {
        private readonly List<(Control Control, StackRow Kind)> _rows = new();
        private bool _inLayout;

        public SettingsStack()
        {
            DoubleBuffered = true;
            BackColor = Color.White;
            Padding = new Padding(20, 16, 20, 16);
        }

        public void Add(Control control, StackRow kind)
        {
            _rows.Add((control, kind));
            Controls.Add(control);
        }

        protected override void OnLayout(LayoutEventArgs levent)
        {
            if (_inLayout)
                return;

            _inLayout = true;
            try
            {
                var x = Padding.Left;
                var y = Padding.Top;
                var width = Math.Max(120, ClientSize.Width - Padding.Horizontal);
                foreach (var (control, kind) in _rows)
                {
                    var gap = kind is StackRow.Caption ? 4 : 12;
                    if (kind == StackRow.Field)
                    {
                        var fieldHeight = control is Panel ? 34 : 30;
                        control.SetBounds(x, y, width, fieldHeight);
                        y += fieldHeight + gap;
                        continue;
                    }

                    if (kind == StackRow.Buttons)
                    {
                        PlaceButtons(control, x, y, width);
                        y += control.Height + gap;
                        continue;
                    }

                    control.MaximumSize = new Size(width, 0);
                    var preferred = control.GetPreferredSize(new Size(width, 0));
                    control.SetBounds(x, y, width, Math.Max(18, preferred.Height));
                    y += control.Height + gap;
                }
            }
            finally
            {
                _inLayout = false;
            }
        }

        private static void PlaceButtons(Control row, int x, int y, int width)
        {
            var cursor = 0;
            var height = 32;
            foreach (Control child in row.Controls)
            {
                var preferred = child.PreferredSize;
                var w = Math.Max(child.MinimumSize.Width, preferred.Width);
                var h = Math.Max(32, preferred.Height);
                child.SetBounds(cursor, 0, w, h);
                cursor += w + 8;
                height = Math.Max(height, h);
            }

            row.SetBounds(x, y, width, height);
        }
    }

    private sealed class TabStrip : Panel
    {
        public TabStrip()
        {
            Dock = DockStyle.Top;
            Height = 44;
            BackColor = Color.FromArgb(243, 243, 243);
            Padding = new Padding(12, 8, 12, 0);
            DoubleBuffered = true;
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            base.OnPaint(e);
            using var pen = new Pen(Color.FromArgb(229, 229, 229));
            e.Graphics.DrawLine(pen, 0, Height - 1, Width, Height - 1);
        }

        protected override void OnLayout(LayoutEventArgs levent)
        {
            var x = Padding.Left;
            foreach (Control child in Controls)
            {
                var width = Math.Max(72, child.PreferredSize.Width);
                child.SetBounds(x, 6, width, Height - 7);
                x += width + 4;
            }
        }
    }

    private sealed class TabChip : Control
    {
        private bool _selected;
        private bool _hover;

        public TabChip()
        {
            SetStyle(ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer | ControlStyles.ResizeRedraw | ControlStyles.UserPaint, true);
            Cursor = Cursors.Hand;
            Font = new Font("Segoe UI", 9.5f);
            Height = 36;
        }

        public bool Selected
        {
            get => _selected;
            set
            {
                if (_selected == value)
                    return;
                _selected = value;
                Invalidate();
            }
        }

        protected override void OnTextChanged(EventArgs e)
        {
            base.OnTextChanged(e);
            Parent?.PerformLayout();
            Invalidate();
        }

        protected override void OnMouseEnter(EventArgs e)
        {
            _hover = true;
            Invalidate();
            base.OnMouseEnter(e);
        }

        protected override void OnMouseLeave(EventArgs e)
        {
            _hover = false;
            Invalidate();
            base.OnMouseLeave(e);
        }

        public override Size GetPreferredSize(Size proposedSize)
        {
            var text = TextRenderer.MeasureText(Text, Font);
            return new Size(text.Width + 28, 36);
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            var g = e.Graphics;
            var back = _selected
                ? Color.White
                : _hover ? Color.FromArgb(232, 232, 232) : Color.FromArgb(243, 243, 243);
            using (var brush = new SolidBrush(back))
                g.FillRectangle(brush, ClientRectangle);

            if (_selected)
            {
                using var accent = new Pen(Color.FromArgb(0, 120, 212), 2);
                g.DrawLine(accent, 10, Height - 2, Width - 10, Height - 2);
            }

            var color = _selected ? Color.FromArgb(20, 20, 20) : Color.FromArgb(80, 80, 80);
            TextRenderer.DrawText(
                g,
                Text,
                Font,
                ClientRectangle,
                color,
                TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter | TextFormatFlags.EndEllipsis);
        }
    }
}
