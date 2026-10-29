using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Automation;
using System.Windows.Forms;

// Local foreground-only adapter. Never enumerates browser history or attaches a debugger.
internal sealed class OnlineBrowserBridge : IDisposable
{
    [DllImport("user32.dll")] private static extern IntPtr GetForegroundWindow();
    [DllImport("user32.dll")] private static extern uint GetWindowThreadProcessId(IntPtr window, out uint process);
    [DllImport("user32.dll")] private static extern bool IsWindow(IntPtr window);
    [DllImport("user32.dll")] private static extern bool IsIconic(IntPtr window);
    [DllImport("user32.dll")] private static extern bool SetForegroundWindow(IntPtr window);
    [DllImport("user32.dll")] private static extern bool ShowWindowAsync(IntPtr window, int command);
    private readonly object gate = new object();
    private System.Threading.Timer watcher;
    private int polling;
    private Target target;
    private bool disposed;
    public volatile bool Enabled;
    private sealed class Target
    {
        internal IntPtr Window;
        internal string Url, Title;
        internal AutomationElement Focus;
    }
    public void Start()
    {
        watcher = new System.Threading.Timer(delegate { Observe(); }, null, 0, 350);
    }
    internal static string DocumentUrl(string value)
    {
        if (String.IsNullOrWhiteSpace(value)) return null;
        if (value.StartsWith("www.kdocs.cn/", StringComparison.OrdinalIgnoreCase) || value.StartsWith("kdocs.cn/", StringComparison.OrdinalIgnoreCase)) value = "https://" + value;
        Uri uri;
        if (!Uri.TryCreate(value, UriKind.Absolute, out uri) || uri.Scheme != "https" || !String.IsNullOrEmpty(uri.UserInfo) || !uri.IsDefaultPort) return null;
        if (uri.Host != "www.kdocs.cn" && uri.Host != "kdocs.cn") return null;
        string path = uri.AbsolutePath;
        if (!(path.StartsWith("/l/", StringComparison.Ordinal) || path.StartsWith("/view/l/", StringComparison.Ordinal))) return null;
        string id = path.Substring(path.LastIndexOf('/') + 1);
        if (id.Length == 0) return null;
        foreach (char c in id) if (!Char.IsLetterOrDigit(c) && c != '-' && c != '_') return null;
        // Query changes do not change the document. Host variants refer to the same service.
        return "https://www.kdocs.cn/l/" + id;
    }
    private static bool BrowserWindow(IntPtr hwnd)
    {
        uint pid;
        GetWindowThreadProcessId(hwnd, out pid);
        try { using (Process p = Process.GetProcessById((int)pid)) return p.ProcessName.Equals("msedge", StringComparison.OrdinalIgnoreCase); }
        catch { return false; }
    }
    private static string ReadUrl(AutomationElement window)
    {
        var edits = window.FindAll(TreeScope.Descendants, new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Edit));
        foreach (AutomationElement edit in edits)
        {
            string name = edit.Current.Name ?? "", id = edit.Current.AutomationId ?? "";
            if (!(id.Equals("addressEditBox", StringComparison.OrdinalIgnoreCase) || name.IndexOf("地址", StringComparison.OrdinalIgnoreCase) >= 0 || name.IndexOf("address", StringComparison.OrdinalIgnoreCase) >= 0)) continue;
            object pattern;
            if (edit.TryGetCurrentPattern(ValuePattern.Pattern, out pattern)) return DocumentUrl(((ValuePattern)pattern).Current.Value);
        }
        return null;
    }
    private static bool InPage(AutomationElement element, IntPtr hwnd)
    {
        bool document = false;
        for (int i = 0; element != null && i < 90; i++, element = TreeWalker.ControlViewWalker.GetParent(element))
        {
            if (element.Current.ControlType == ControlType.Document) document = true;
            if (element.Current.NativeWindowHandle == hwnd.ToInt32()) return document;
        }
        return false;
    }
    private void Observe()
    {
        if (disposed || !Enabled || Interlocked.Exchange(ref polling, 1) != 0) return;
        try
        {
            IntPtr hwnd = GetForegroundWindow();
            if (!BrowserWindow(hwnd)) return;
            AutomationElement focused = AutomationElement.FocusedElement;
            if (focused == null || !InPage(focused, hwnd)) return;
            var window = AutomationElement.FromHandle(hwnd);
            string url = ReadUrl(window);
            if (url == null || hwnd != GetForegroundWindow()) return;
            var captured = new Target { Window = hwnd, Url = url, Title = window.Current.Name, Focus = focused };
            lock (gate) target = captured;
        }
        catch { /* A navigation or busy accessibility provider is not a target. */ }
        finally { Interlocked.Exchange(ref polling, 0); }
    }
    public string Description
    {
        get { lock (gate) return target == null ? "请先在 Edge 金山表格中点选目标格" : "网页目标：" + target.Title; }
    }
    public string Send()
    {
        Target captured;
        lock (gate) captured = target;
        if (captured == null) throw new InvalidOperationException("尚未识别网页目标。请在 Edge 金山表格里单击目标格，等一秒，再回轻图点击发送。");
        if (!IsWindow(captured.Window) || !BrowserWindow(captured.Window)) throw new InvalidOperationException("网页窗口已关闭，请重新选择目标格。");
        var window = AutomationElement.FromHandle(captured.Window);
        if (ReadUrl(window) != captured.Url) throw new InvalidOperationException("网页已切换，已取消发送。请在正确的金山表格里重新选格。");
        RestoreIfMinimized(captured.Window);
        if (!ActivateWindow(captured.Window)) throw new InvalidOperationException("Windows 未允许切回 Edge，已停止发送。请先在网页点选目标格，再点击新版拼图按钮。");
        // Restore the captured page control, never the address bar or a toolbar input.
        captured.Focus.SetFocus();
        AutomationElement focus = AutomationElement.FocusedElement;
        if (GetForegroundWindow() != captured.Window || focus == null || !InPage(focus, captured.Window) || ReadUrl(window) != captured.Url)
            throw new InvalidOperationException("网页焦点已改变，已取消粘贴。请重新选中目标格。");
        SendKeys.SendWait("^v");
        return captured.Title;
    }
    internal static void RestoreIfMinimized(IntPtr window)
    {
        // SW_RESTORE also removes maximization/fullscreen. Do not touch a visible window.
        if (IsIconic(window)) ShowWindowAsync(window, 9);
    }
    internal static bool ActivateWindow(IntPtr window)
    {
        if (GetForegroundWindow() == window) return true;
        SetForegroundWindow(window);
        // Activation may finish after SetForegroundWindow returns. The observed foreground
        // window is authoritative, rather than the API's return value or an immediate check.
        Stopwatch watch = Stopwatch.StartNew();
        while (watch.ElapsedMilliseconds < 800)
        {
            if (GetForegroundWindow() == window) return true;
            Thread.Sleep(15);
        }
        return GetForegroundWindow() == window;
    }
    public void Dispose()
    {
        disposed = true;
        if (watcher != null) watcher.Dispose();
    }
}
