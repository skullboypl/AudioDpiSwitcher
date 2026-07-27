using System.Runtime.InteropServices;
using NAudio.CoreAudioApi;

namespace AudioDpiSwitcher.Services;

internal enum AudioRoleKind
{
    Console,
    Communications,
}

internal sealed record AudioDeviceInfo(string Id, string Name, bool IsPlayback);

internal sealed class AudioService : IDisposable
{
    private readonly MMDeviceEnumerator _enumerator = new();
    private readonly object _cacheLock = new();
    private IReadOnlyList<AudioDeviceInfo>? _playbackCache;
    private IReadOnlyList<AudioDeviceInfo>? _recordingCache;
    private DateTime _cacheUtc = DateTime.MinValue;
    private static readonly TimeSpan CacheTtl = TimeSpan.FromSeconds(8);

    public IReadOnlyList<AudioDeviceInfo> ListPlayback(bool forceRefresh = false) =>
        GetCached(DataFlow.Render, forceRefresh);

    public IReadOnlyList<AudioDeviceInfo> ListRecording(bool forceRefresh = false) =>
        GetCached(DataFlow.Capture, forceRefresh);

    public void InvalidateCache()
    {
        lock (_cacheLock)
        {
            _playbackCache = null;
            _recordingCache = null;
            _cacheUtc = DateTime.MinValue;
        }
    }

    /// <summary>Snapshot of lists + defaults in one pass (for settings UI).</summary>
    public AudioSnapshot CaptureSnapshot(bool forceRefresh = false)
    {
        var outs = ListPlayback(forceRefresh);
        var ins = ListRecording(forceRefresh);
        return new AudioSnapshot(outs, ins, GetDefaults());
    }

    private IReadOnlyList<AudioDeviceInfo> GetCached(DataFlow flow, bool forceRefresh)
    {
        lock (_cacheLock)
        {
            var fresh = !forceRefresh && DateTime.UtcNow - _cacheUtc < CacheTtl;
            if (flow == DataFlow.Render)
            {
                if (fresh && _playbackCache is not null)
                    return _playbackCache;
            }
            else if (fresh && _recordingCache is not null)
            {
                return _recordingCache;
            }
        }

        var list = Enumerate(flow);
        lock (_cacheLock)
        {
            if (flow == DataFlow.Render)
                _playbackCache = list;
            else
                _recordingCache = list;
            _cacheUtc = DateTime.UtcNow;
        }

        return list;
    }

    private List<AudioDeviceInfo> Enumerate(DataFlow flow)
    {
        var result = new List<AudioDeviceInfo>();
        var collection = _enumerator.EnumerateAudioEndPoints(flow, DeviceState.Active);
        for (var i = 0; i < collection.Count; i++)
        {
            var device = collection[i];
            result.Add(new AudioDeviceInfo(
                device.ID,
                device.FriendlyName,
                flow == DataFlow.Render));
        }

        return result.OrderBy(d => d.Name, StringComparer.CurrentCultureIgnoreCase).ToList();
    }

    public string? GetDefaultId(DataFlow flow, Role role)
    {
        try
        {
            using var device = _enumerator.GetDefaultAudioEndpoint(flow, role);
            return device.ID;
        }
        catch
        {
            return null;
        }
    }

    public Dictionary<string, string?> GetDefaults()
    {
        return new Dictionary<string, string?>
        {
            ["output"] = GetDefaultId(DataFlow.Render, Role.Multimedia),
            ["output_comm"] = GetDefaultId(DataFlow.Render, Role.Communications),
            ["input"] = GetDefaultId(DataFlow.Capture, Role.Multimedia),
            ["input_comm"] = GetDefaultId(DataFlow.Capture, Role.Communications),
        };
    }

    public string FindName(IEnumerable<AudioDeviceInfo> pool, string? id)
    {
        if (string.IsNullOrWhiteSpace(id))
            return I18n.T("unknown");

        foreach (var d in pool)
        {
            if (string.Equals(d.Id, id, StringComparison.OrdinalIgnoreCase))
                return d.Name;
        }

        return I18n.T("unknown");
    }

    public void SetDefault(string deviceId, AudioRoleKind role)
    {
        if (string.IsNullOrWhiteSpace(deviceId))
            return;

        PolicyConfigClient.SetDefaultEndpoint(deviceId, role);
        InvalidateCache();
    }

    public void Dispose() => _enumerator.Dispose();
}

internal sealed record AudioSnapshot(
    IReadOnlyList<AudioDeviceInfo> Playback,
    IReadOnlyList<AudioDeviceInfo> Recording,
    Dictionary<string, string?> Defaults);

/// <summary>
/// Sets default audio endpoints via PolicyConfig COM (AudioDeviceCmdlets / SoundSwitch pattern).
/// </summary>
internal static class PolicyConfigClient
{
    public static void SetDefaultEndpoint(string deviceId, AudioRoleKind roleKind)
    {
        var client = new CPolicyConfigClient();
        var policy = (IPolicyConfig)client;

        try
        {
            if (roleKind == AudioRoleKind.Console)
            {
                policy.SetDefaultEndpoint(deviceId, 0); // eConsole
                policy.SetDefaultEndpoint(deviceId, 1); // eMultimedia
            }
            else
            {
                policy.SetDefaultEndpoint(deviceId, 2); // eCommunications
            }
        }
        finally
        {
            Marshal.FinalReleaseComObject(policy);
        }
    }

    [ComImport]
    [Guid("870af99c-171d-4f9e-af0d-e63df40c2bc9")]
    private class CPolicyConfigClient
    {
    }

    [ComImport]
    [Guid("F8679F50-850A-41CF-9C72-430F290290C8")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    private interface IPolicyConfig
    {
        void GetMixFormat([MarshalAs(UnmanagedType.LPWStr)] string pszDeviceName, out IntPtr ppFormat);
        void GetDeviceFormat([MarshalAs(UnmanagedType.LPWStr)] string pszDeviceName, int bDefault, out IntPtr ppFormat);
        void ResetDeviceFormat([MarshalAs(UnmanagedType.LPWStr)] string pszDeviceName);
        void SetDeviceFormat([MarshalAs(UnmanagedType.LPWStr)] string pszDeviceName, IntPtr pEndpointFormatCapture, IntPtr pEndpointFormatRender);
        void GetProcessingPeriod([MarshalAs(UnmanagedType.LPWStr)] string pszDeviceName, int bDefault, out long pmftDefaultPeriod, out long pmftMinimumPeriod);
        void SetProcessingPeriod([MarshalAs(UnmanagedType.LPWStr)] string pszDeviceName, ref long pmftPeriod);
        void GetShareMode([MarshalAs(UnmanagedType.LPWStr)] string pszDeviceName, IntPtr pMode);
        void SetShareMode([MarshalAs(UnmanagedType.LPWStr)] string pszDeviceName, IntPtr mode);
        void GetPropertyValue([MarshalAs(UnmanagedType.LPWStr)] string pszDeviceName, ref PropertyKey key, out PropVariant pv);
        void SetPropertyValue([MarshalAs(UnmanagedType.LPWStr)] string pszDeviceName, ref PropertyKey key, ref PropVariant pv);
        void SetDefaultEndpoint([MarshalAs(UnmanagedType.LPWStr)] string pszDeviceName, int role);
        void SetEndpointVisibility([MarshalAs(UnmanagedType.LPWStr)] string pszDeviceName, int bVisible);
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct PropertyKey
    {
        public Guid fmtid;
        public int pid;
    }

    [StructLayout(LayoutKind.Explicit)]
    private struct PropVariant
    {
        [FieldOffset(0)] public short vt;
        [FieldOffset(8)] public IntPtr pointerValue;
    }
}
