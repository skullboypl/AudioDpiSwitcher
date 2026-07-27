using System.Runtime.InteropServices;
using Microsoft.Win32;

namespace AudioDpiSwitcher.Services;

/// <summary>
/// Native per-monitor DPI get/set via undocumented DisplayConfig DPI IOCTLs
/// (same approach as https://github.com/imniko/SetDPI / lihas windows-DPI-scaling-sample).
/// Extended scale table for Windows 11.
/// </summary>
internal static class DpiService
{
    // Observed / extrapolated from Windows Settings (Win10/11)
    private static readonly uint[] DpiVals =
        [100, 125, 150, 175, 200, 225, 250, 300, 350, 400, 450, 500];

    public static readonly int[] Presets =
        [100, 125, 150, 175, 200, 225, 250, 300, 350, 400];

    public static IReadOnlyList<DpiDisplay> ListDisplays()
    {
        if (!TryGetPaths(out var paths, out var modes, out var virtualAware))
            return [];

        var result = new List<DpiDisplay>(paths.Length);
        for (var i = 0; i < paths.Length; i++)
        {
            var path = paths[i];
            var adapter = path.targetInfo.adapterId;
            var targetId = path.targetInfo.id;
            var sourceId = path.sourceInfo.id;
            var name = GetTargetName(adapter, targetId) ?? $"Monitor #{i + 1}";
            var info = GetDpiScalingInfo(adapter, sourceId);
            TryGetSourceResolution(path, modes, virtualAware, out var width, out var height);
            result.Add(new DpiDisplay(
                Index: i + 1,
                Name: name,
                AdapterId: adapter,
                SourceId: sourceId,
                TargetId: targetId,
                CurrentDpi: info.Current,
                RecommendedDpi: info.Recommended,
                MinDpi: info.Minimum,
                MaxDpi: info.Maximum,
                Width: width,
                Height: height));
        }

        return result;
    }

    public static string? GetDpiValue(int monitorIndex1Based)
    {
        var displays = ListDisplays();
        if (monitorIndex1Based < 1 || monitorIndex1Based > displays.Count)
            return null;
        return displays[monitorIndex1Based - 1].CurrentDpi.ToString();
    }

    public static bool SetDpi(int percent, int monitorIndex1Based = 1)
    {
        var displays = ListDisplays();
        if (monitorIndex1Based < 1 || monitorIndex1Based > displays.Count)
            return false;

        var d = displays[monitorIndex1Based - 1];
        var ok = SetDpiScaling(d.AdapterId, d.SourceId, (uint)percent);
        if (ok && monitorIndex1Based == 1)
            TryUpdatePrimaryAppliedDpiRegistry(percent);
        return ok;
    }

    public static bool SetDpi(DpiDisplay display, int percent)
    {
        var ok = SetDpiScaling(display.AdapterId, display.SourceId, (uint)percent);
        if (ok && display.Index == 1)
            TryUpdatePrimaryAppliedDpiRegistry(percent);
        return ok;
    }

    private static void TryUpdatePrimaryAppliedDpiRegistry(int percent)
    {
        try
        {
            // Same legacy touch-up as SetDPI for primary monitor
            using var key = Registry.CurrentUser.OpenSubKey(
                @"Control Panel\Desktop\WindowMetrics", writable: true);
            if (key is null)
                return;
            key.SetValue("AppliedDPI", (int)(percent * 0.96), RegistryValueKind.DWord);
        }
        catch
        {
            // non-fatal
        }
    }

    private static bool TryGetPaths(
        out DISPLAYCONFIG_PATH_INFO[] paths,
        out DISPLAYCONFIG_MODE_INFO[] modes,
        out bool virtualAware)
    {
        paths = [];
        modes = [];
        virtualAware = false;

        // QDC_ALL_PATHS = 0x1, QDC_ONLY_ACTIVE_PATHS = 0x2, QDC_VIRTUAL_MODE_AWARE = 0x10
        const uint qdcOnlyActive = 0x00000002;
        const uint qdcVirtualAware = 0x00000010;

        foreach (var flags in new[] { qdcOnlyActive | qdcVirtualAware, qdcOnlyActive })
        {
            var err = GetDisplayConfigBufferSizes(flags, out var numPaths, out var numModes);
            if (err != 0 || numPaths == 0)
                continue;

            var pathBuf = new DISPLAYCONFIG_PATH_INFO[numPaths];
            var modeBuf = new DISPLAYCONFIG_MODE_INFO[numModes];
            err = QueryDisplayConfig(flags, ref numPaths, pathBuf, ref numModes, modeBuf, IntPtr.Zero);
            if (err != 0)
                continue;

            if (numPaths != pathBuf.Length)
                Array.Resize(ref pathBuf, (int)numPaths);
            if (numModes != modeBuf.Length)
                Array.Resize(ref modeBuf, (int)numModes);

            paths = pathBuf;
            modes = modeBuf;
            virtualAware = (flags & qdcVirtualAware) != 0;
            return paths.Length > 0;
        }

        return false;
    }

    private static void TryGetSourceResolution(
        DISPLAYCONFIG_PATH_INFO path,
        DISPLAYCONFIG_MODE_INFO[] modes,
        bool virtualAware,
        out uint width,
        out uint height)
    {
        width = 0;
        height = 0;

        // Prefer matching by adapter + source id (more reliable than modeInfoIdx across Win10/11)
        for (var i = 0; i < modes.Length; i++)
        {
            var mode = modes[i];
            // DISPLAYCONFIG_MODE_INFO_TYPE_SOURCE = 1
            if (mode.infoType != 1)
                continue;
            if (mode.id != path.sourceInfo.id)
                continue;
            if (mode.adapterId.LowPart != path.sourceInfo.adapterId.LowPart ||
                mode.adapterId.HighPart != path.sourceInfo.adapterId.HighPart)
                continue;

            width = mode.modeInfo.sourceMode.width;
            height = mode.modeInfo.sourceMode.height;
            if (width > 0 && height > 0)
                return;
        }

        // Fallback: modeInfoIdx (low 16 bits when virtual-mode aware)
        var modeIdx = path.sourceInfo.modeInfoIdx;
        if (virtualAware)
            modeIdx &= 0xFFFF;

        if (modeIdx < modes.Length && modes[modeIdx].infoType == 1)
        {
            width = modes[modeIdx].modeInfo.sourceMode.width;
            height = modes[modeIdx].modeInfo.sourceMode.height;
        }
    }

    private static string? GetTargetName(LUID adapterId, uint targetId)
    {
        var deviceName = new DISPLAYCONFIG_TARGET_DEVICE_NAME();
        deviceName.header.type = DISPLAYCONFIG_DEVICE_INFO_TYPE.DISPLAYCONFIG_DEVICE_INFO_GET_TARGET_NAME;
        deviceName.header.size = (uint)Marshal.SizeOf<DISPLAYCONFIG_TARGET_DEVICE_NAME>();
        deviceName.header.adapterId = adapterId;
        deviceName.header.id = targetId;

        if (DisplayConfigGetDeviceInfo(ref deviceName) != 0)
            return null;

        var name = deviceName.monitorFriendlyDeviceName;
        if (string.IsNullOrWhiteSpace(name))
            name = deviceName.monitorDevicePath;
        return string.IsNullOrWhiteSpace(name) ? null : name.Trim();
    }

    private static DpiScalingInfo GetDpiScalingInfo(LUID adapterId, uint sourceId)
    {
        var packet = new DISPLAYCONFIG_SOURCE_DPI_SCALE_GET();
        packet.header.type = (DISPLAYCONFIG_DEVICE_INFO_TYPE)(-3); // GET_DPI_SCALE
        packet.header.size = (uint)Marshal.SizeOf<DISPLAYCONFIG_SOURCE_DPI_SCALE_GET>();
        packet.header.adapterId = adapterId;
        packet.header.id = sourceId;

        if (DisplayConfigGetDeviceInfo(ref packet) != 0)
            return default;

        var cur = packet.curScaleRel;
        if (cur < packet.minScaleRel) cur = packet.minScaleRel;
        if (cur > packet.maxScaleRel) cur = packet.maxScaleRel;

        var minAbs = Math.Abs(packet.minScaleRel);
        if (minAbs + packet.maxScaleRel + 1 > DpiVals.Length)
            return default;

        return new DpiScalingInfo(
            Minimum: DpiVals[minAbs + packet.minScaleRel],
            Recommended: DpiVals[minAbs],
            Current: DpiVals[minAbs + cur],
            Maximum: DpiVals[minAbs + packet.maxScaleRel],
            Ok: true);
    }

    private static bool SetDpiScaling(LUID adapterId, uint sourceId, uint dpiPercentToSet)
    {
        var info = GetDpiScalingInfo(adapterId, sourceId);
        if (!info.Ok)
            return false;

        if (dpiPercentToSet == info.Current)
            return true;

        dpiPercentToSet = Math.Clamp(dpiPercentToSet, info.Minimum, info.Maximum);

        var idxTarget = IndexOfDpi(dpiPercentToSet);
        var idxRecommended = IndexOfDpi(info.Recommended);
        if (idxTarget < 0 || idxRecommended < 0)
            return false;

        var setPacket = new DISPLAYCONFIG_SOURCE_DPI_SCALE_SET();
        setPacket.header.type = (DISPLAYCONFIG_DEVICE_INFO_TYPE)(-4); // SET_DPI_SCALE
        setPacket.header.size = (uint)Marshal.SizeOf<DISPLAYCONFIG_SOURCE_DPI_SCALE_SET>();
        setPacket.header.adapterId = adapterId;
        setPacket.header.id = sourceId;
        setPacket.scaleRel = idxTarget - idxRecommended;

        return DisplayConfigSetDeviceInfo(ref setPacket) == 0;
    }

    private static int IndexOfDpi(uint value)
    {
        for (var i = 0; i < DpiVals.Length; i++)
        {
            if (DpiVals[i] == value)
                return i;
        }

        return -1;
    }

    // —— P/Invoke ——

    [DllImport("user32.dll")]
    private static extern int GetDisplayConfigBufferSizes(
        uint flags, out uint numPathArrayElements, out uint numModeInfoArrayElements);

    [DllImport("user32.dll")]
    private static extern int QueryDisplayConfig(
        uint flags,
        ref uint numPathArrayElements,
        [Out] DISPLAYCONFIG_PATH_INFO[] pathArray,
        ref uint numModeInfoArrayElements,
        [Out] DISPLAYCONFIG_MODE_INFO[] modeInfoArray,
        IntPtr currentTopologyId);

    [DllImport("user32.dll")]
    private static extern int DisplayConfigGetDeviceInfo(ref DISPLAYCONFIG_SOURCE_DPI_SCALE_GET requestPacket);

    [DllImport("user32.dll")]
    private static extern int DisplayConfigGetDeviceInfo(ref DISPLAYCONFIG_TARGET_DEVICE_NAME requestPacket);

    [DllImport("user32.dll")]
    private static extern int DisplayConfigSetDeviceInfo(ref DISPLAYCONFIG_SOURCE_DPI_SCALE_SET setPacket);

    private enum DISPLAYCONFIG_DEVICE_INFO_TYPE : int
    {
        DISPLAYCONFIG_DEVICE_INFO_GET_TARGET_NAME = 2,
    }

    [StructLayout(LayoutKind.Sequential)]
    internal struct LUID
    {
        public uint LowPart;
        public int HighPart;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct DISPLAYCONFIG_DEVICE_INFO_HEADER
    {
        public DISPLAYCONFIG_DEVICE_INFO_TYPE type;
        public uint size;
        public LUID adapterId;
        public uint id;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct DISPLAYCONFIG_SOURCE_DPI_SCALE_GET
    {
        public DISPLAYCONFIG_DEVICE_INFO_HEADER header;
        public int minScaleRel;
        public int curScaleRel;
        public int maxScaleRel;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct DISPLAYCONFIG_SOURCE_DPI_SCALE_SET
    {
        public DISPLAYCONFIG_DEVICE_INFO_HEADER header;
        public int scaleRel;
    }

    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
    private struct DISPLAYCONFIG_TARGET_DEVICE_NAME
    {
        public DISPLAYCONFIG_DEVICE_INFO_HEADER header;
        public uint flags;
        public int outputTechnology;
        public ushort edidManufactureId;
        public ushort edidProductCodeId;
        public uint connectorInstance;
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 64)]
        public string monitorFriendlyDeviceName;
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 128)]
        public string monitorDevicePath;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct DISPLAYCONFIG_PATH_SOURCE_INFO
    {
        public LUID adapterId;
        public uint id;
        public uint modeInfoIdx;
        public uint statusFlags;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct DISPLAYCONFIG_PATH_TARGET_INFO
    {
        public LUID adapterId;
        public uint id;
        public uint modeInfoIdx;
        public uint outputTechnology;
        public uint rotation;
        public uint scaling;
        public DISPLAYCONFIG_RATIONAL refreshRate;
        public uint scanLineOrdering;
        public int targetAvailable;
        public uint statusFlags;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct DISPLAYCONFIG_RATIONAL
    {
        public uint Numerator;
        public uint Denominator;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct DISPLAYCONFIG_PATH_INFO
    {
        public DISPLAYCONFIG_PATH_SOURCE_INFO sourceInfo;
        public DISPLAYCONFIG_PATH_TARGET_INFO targetInfo;
        public uint flags;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct DISPLAYCONFIG_MODE_INFO
    {
        public uint infoType;
        public uint id;
        public LUID adapterId;
        public DISPLAYCONFIG_MODE_INFO_union modeInfo;
    }

    [StructLayout(LayoutKind.Explicit, Size = 48)]
    private struct DISPLAYCONFIG_MODE_INFO_union
    {
        [FieldOffset(0)] public DISPLAYCONFIG_TARGET_MODE targetMode;
        [FieldOffset(0)] public DISPLAYCONFIG_SOURCE_MODE sourceMode;
        [FieldOffset(0)] public DISPLAYCONFIG_DESKTOP_IMAGE_INFO desktopImageInfo;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct DISPLAYCONFIG_TARGET_MODE
    {
        public DISPLAYCONFIG_VIDEO_SIGNAL_INFO targetVideoSignalInfo;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct DISPLAYCONFIG_VIDEO_SIGNAL_INFO
    {
        public ulong pixelRate;
        public DISPLAYCONFIG_RATIONAL hSyncFreq;
        public DISPLAYCONFIG_RATIONAL vSyncFreq;
        public DISPLAYCONFIG_2DREGION activeSize;
        public DISPLAYCONFIG_2DREGION totalSize;
        public uint videoStandard;
        public uint scanLineOrdering;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct DISPLAYCONFIG_2DREGION
    {
        public uint cx;
        public uint cy;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct DISPLAYCONFIG_SOURCE_MODE
    {
        public uint width;
        public uint height;
        public uint pixelFormat;
        public POINTL position;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct POINTL
    {
        public int x;
        public int y;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct DISPLAYCONFIG_DESKTOP_IMAGE_INFO
    {
        public POINTL PathSourceSize;
        public RECTL DesktopImageRegion;
        public RECTL DesktopImageClip;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct RECTL
    {
        public int left;
        public int top;
        public int right;
        public int bottom;
    }

    private readonly record struct DpiScalingInfo(
        uint Minimum,
        uint Maximum,
        uint Current,
        uint Recommended,
        bool Ok);
}

internal sealed record DpiDisplay(
    int Index,
    string Name,
    DpiService.LUID AdapterId,
    uint SourceId,
    uint TargetId,
    uint CurrentDpi,
    uint RecommendedDpi,
    uint MinDpi,
    uint MaxDpi,
    uint Width,
    uint Height)
{
    public string StableKey => $"{AdapterId.LowPart}:{AdapterId.HighPart}:{TargetId}:{SourceId}";
}
