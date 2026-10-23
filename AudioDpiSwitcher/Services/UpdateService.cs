using System.Diagnostics;
using System.Net.Http;
using System.Reflection;
using System.Text.Json;

namespace AudioDpiSwitcher.Services;

internal sealed record UpdateOffer(Version Version, string Tag, string DownloadUrl);

internal static class UpdateService
{
    private static readonly HttpClient Http = CreateClient();

    public static Version CurrentVersion =>
        Assembly.GetExecutingAssembly().GetName().Version ?? new Version(0, 0);

    /// <summary>Single-file publish has an empty assembly location. Debug builds must not replace themselves.</summary>
    public static bool CanReplaceRunningExe =>
        string.IsNullOrEmpty(Assembly.GetExecutingAssembly().Location)
        && !string.IsNullOrEmpty(Environment.ProcessPath);

    public static async Task<UpdateOffer?> CheckAsync(CancellationToken cancellationToken = default)
    {
        using var response = await Http.GetAsync(
            "https://api.github.com/repos/skullboypl/AudioDpiSwitcher/releases/latest",
            cancellationToken).ConfigureAwait(false);
        response.EnsureSuccessStatusCode();
        await using var stream = await response.Content.ReadAsStreamAsync(cancellationToken).ConfigureAwait(false);
        using var doc = await JsonDocument.ParseAsync(stream, cancellationToken: cancellationToken).ConfigureAwait(false);
        var root = doc.RootElement;
        var tag = root.GetProperty("tag_name").GetString() ?? "";
        var versionText = tag.Trim().TrimStart('v', 'V');
        if (!Version.TryParse(versionText, out var version))
            return null;

        string? url = null;
        if (root.TryGetProperty("assets", out var assets))
        {
            foreach (var asset in assets.EnumerateArray())
            {
                var name = asset.GetProperty("name").GetString();
                if (!string.Equals(name, "AudioDpiSwitcher.exe", StringComparison.OrdinalIgnoreCase))
                    continue;
                url = asset.GetProperty("browser_download_url").GetString();
                break;
            }
        }

        if (string.IsNullOrWhiteSpace(url))
            return null;

        return new UpdateOffer(version, tag, url);
    }

    public static bool IsNewer(UpdateOffer offer)
    {
        var current = new Version(CurrentVersion.Major, CurrentVersion.Minor, Math.Max(CurrentVersion.Build, 0));
        var remote = new Version(offer.Version.Major, offer.Version.Minor, Math.Max(offer.Version.Build, 0));
        return remote > current;
    }

    public static async Task<string> DownloadAsync(UpdateOffer offer, CancellationToken cancellationToken = default)
    {
        var path = Path.Combine(Path.GetTempPath(), $"AudioDpiSwitcher-{offer.Version}.exe");
        await using var stream = await Http.GetStreamAsync(offer.DownloadUrl, cancellationToken).ConfigureAwait(false);
        await using var file = File.Create(path);
        await stream.CopyToAsync(file, cancellationToken).ConfigureAwait(false);
        return path;
    }

    public static void LaunchReplacer(string downloadedExe)
    {
        var destination = Environment.ProcessPath
            ?? throw new InvalidOperationException("Running executable path is unknown.");
        var script = Path.Combine(Path.GetTempPath(), "AudioDpiSwitcher-update.ps1");
        File.WriteAllBytes(script, ReadScript());

        Process.Start(new ProcessStartInfo
        {
            FileName = "powershell",
            Arguments =
                $"-NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -File \"{script}\" -ProcessId {Environment.ProcessId} -Source \"{downloadedExe}\" -Destination \"{destination}\"",
            UseShellExecute = false,
            CreateNoWindow = true,
        });
    }

    private static byte[] ReadScript()
    {
        var assembly = Assembly.GetExecutingAssembly();
        using var stream = assembly.GetManifestResourceStream("AudioDpiSwitcher.update.ps1")
            ?? throw new InvalidOperationException("update.ps1 is not embedded.");
        using var memory = new MemoryStream();
        stream.CopyTo(memory);
        return memory.ToArray();
    }

    private static HttpClient CreateClient()
    {
        var client = new HttpClient { Timeout = TimeSpan.FromMinutes(2) };
        client.DefaultRequestHeaders.UserAgent.ParseAdd("AudioDpiSwitcher");
        client.DefaultRequestHeaders.Accept.ParseAdd("application/vnd.github+json");
        return client;
    }
}
