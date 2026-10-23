namespace AudioDpiSwitcher.Services;

internal static class StartupService
{
    private const string ShortcutName = "AudioDpiSwitcher.lnk";

    public static bool IsEnabled() => OurShortcuts().Any();

    public static void SetEnabled(bool enabled)
    {
        var canonical = Path.Combine(Folder(), ShortcutName);
        if (enabled)
        {
            var exe = Environment.ProcessPath;
            if (string.IsNullOrWhiteSpace(exe))
                return;

            WriteShortcut(canonical, exe);
        }

        foreach (var link in OurShortcuts())
        {
            if (enabled && string.Equals(link, canonical, StringComparison.OrdinalIgnoreCase))
                continue;

            try
            {
                File.Delete(link);
            }
            catch
            {
                // leave a shortcut we cannot remove
            }
        }
    }

    private static string Folder() =>
        Environment.GetFolderPath(Environment.SpecialFolder.Startup);

    private static IEnumerable<string> OurShortcuts()
    {
        var folder = Folder();
        if (!Directory.Exists(folder))
            yield break;

        foreach (var link in Directory.EnumerateFiles(folder, "*.lnk"))
        {
            var name = Path.GetFileName(link);
            if (name.StartsWith("AudioDpiSwitcher", StringComparison.OrdinalIgnoreCase))
            {
                yield return link;
                continue;
            }

            var target = ReadTarget(link);
            if (target.EndsWith("AudioDpiSwitcher.exe", StringComparison.OrdinalIgnoreCase))
                yield return link;
        }
    }

    private static void WriteShortcut(string path, string exe)
    {
        var shellType = Type.GetTypeFromProgID("WScript.Shell")
            ?? throw new InvalidOperationException("WScript.Shell is unavailable.");
        dynamic shell = Activator.CreateInstance(shellType)!;
        dynamic shortcut = shell.CreateShortcut(path);
        shortcut.TargetPath = exe;
        shortcut.WorkingDirectory = Path.GetDirectoryName(exe) ?? "";
        shortcut.Description = "AudioDpiSwitcher";
        shortcut.Save();
    }

    private static string ReadTarget(string path)
    {
        try
        {
            var shellType = Type.GetTypeFromProgID("WScript.Shell");
            if (shellType is null)
                return "";

            dynamic shell = Activator.CreateInstance(shellType)!;
            dynamic shortcut = shell.CreateShortcut(path);
            return (string?)shortcut.TargetPath ?? "";
        }
        catch
        {
            return "";
        }
    }
}
