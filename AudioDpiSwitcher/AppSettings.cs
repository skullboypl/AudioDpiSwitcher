using System.Text.Json;

namespace AudioDpiSwitcher;

internal sealed class AppSettings
{
    private static readonly string SettingsDir = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
        "AudioDpiSwitcher");

    private static readonly string SettingsPath = Path.Combine(SettingsDir, "settings.json");

    /// <summary>auto | pl | en | de</summary>
    public string Lang { get; set; } = "auto";

    public bool AutoRestoreAudio { get; set; } = true;

    public bool AutoUpdate { get; set; } = true;

    public string? PreferredOutput { get; set; }
    public string? PreferredOutputComm { get; set; }
    public string? PreferredInput { get; set; }
    public string? PreferredInputComm { get; set; }

    public static AppSettings Load()
    {
        try
        {
            if (!File.Exists(SettingsPath))
                return new AppSettings();

            var json = File.ReadAllText(SettingsPath);
            return JsonSerializer.Deserialize<AppSettings>(json) ?? new AppSettings();
        }
        catch
        {
            return new AppSettings();
        }
    }

    public void Save()
    {
        Directory.CreateDirectory(SettingsDir);
        var json = JsonSerializer.Serialize(this, new JsonSerializerOptions { WriteIndented = true });
        File.WriteAllText(SettingsPath, json);
    }
}
