using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Media;

namespace PPEditer.Services;

/// <summary>Per-monitor geometry in WPF device-independent units.</summary>
public sealed class MonitorInfo
{
    public int    Index     { get; init; }
    public double Left      { get; init; }
    public double Top       { get; init; }
    public double Width     { get; init; }
    public double Height    { get; init; }
    public bool   IsPrimary { get; init; }

    public string DisplayName => IsPrimary
        ? $"모니터 {Index + 1}  ({(int)Width}×{(int)Height})  ★ 주 모니터"
        : $"모니터 {Index + 1}  ({(int)Width}×{(int)Height})";
}

/// <summary>Enumerates connected monitors via Win32 EnumDisplayMonitors.</summary>
public static class ScreenHelper
{
    // ── Win32 interop ─────────────────────────────────────────────────

    private delegate bool EnumMonitorProc(
        IntPtr hMonitor, IntPtr hdcMonitor, ref NativeRect lprcMonitor, IntPtr dwData);

    [DllImport("user32.dll")]
    private static extern bool EnumDisplayMonitors(
        IntPtr hdc, IntPtr lprcClip, EnumMonitorProc lpfnEnum, IntPtr dwData);

    [StructLayout(LayoutKind.Sequential)]
    private struct NativeRect { public int Left, Top, Right, Bottom; }

    // ── Enum ──────────────────────────────────────────────────────────

    /// <summary>Returns all connected monitors in WPF device-independent units.</summary>
    public static IReadOnlyList<MonitorInfo> GetMonitors()
    {
        var rects = new List<NativeRect>();

        // Named method required — C# < 14 doesn't allow `ref` in discard-parameter lambdas.
        bool Callback(IntPtr hMon, IntPtr hdc, ref NativeRect r, IntPtr data)
        {
            rects.Add(r);
            return true;
        }

        EnumDisplayMonitors(IntPtr.Zero, IntPtr.Zero, Callback, IntPtr.Zero);

        // Convert physical pixels → WPF device-independent units.
        GetDpiScale(out double sx, out double sy);

        var result = new List<MonitorInfo>(rects.Count);
        for (int i = 0; i < rects.Count; i++)
        {
            var r = rects[i];
            result.Add(new MonitorInfo
            {
                Index     = i,
                Left      = r.Left   * sx,
                Top       = r.Top    * sy,
                Width     = (r.Right  - r.Left) * sx,
                Height    = (r.Bottom - r.Top)  * sy,
                IsPrimary = r.Left == 0 && r.Top == 0,
            });
        }
        return result;
    }

    // ── Helpers ───────────────────────────────────────────────────────

    private static void GetDpiScale(out double sx, out double sy)
    {
        sx = sy = 1.0;
        var win = Application.Current?.MainWindow;
        if (win is null) return;
        var src = PresentationSource.FromVisual(win);
        var m   = src?.CompositionTarget?.TransformFromDevice;
        if (m.HasValue) { sx = m.Value.M11; sy = m.Value.M22; }
    }

    /// <summary>Positions and maximizes <paramref name="wnd"/> on the monitor at <paramref name="monitorIndex"/>.</summary>
    public static void MaximizeOnMonitor(Window wnd, IReadOnlyList<MonitorInfo> monitors, int monitorIndex)
    {
        if (monitorIndex < 0 || monitorIndex >= monitors.Count) return;
        var m = monitors[monitorIndex];
        // Place window inside target monitor, then maximize — WPF maximizes on the monitor the window is on.
        wnd.WindowState = WindowState.Normal;
        wnd.Left        = m.Left + 1;
        wnd.Top         = m.Top  + 1;
        wnd.WindowState = WindowState.Maximized;
    }
}
