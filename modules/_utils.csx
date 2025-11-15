// ============================
// UNIVERSAL UTILITIES MODULE
// ============================
// Loaded by all modules automatically using: #load "_utils.csx"
// Provides: ClickAt, ClickUi, TypeText, PressKey, Wait, FindUi, MoveMouse, SafeInvoke, ClickRelative
// ============================

using System;
using System.Threading;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using System.Windows.Automation;
using System.Drawing;

// -------------------------------
// Win32 API for mouse events and window rect
// -------------------------------
[DllImport("user32.dll")]
static extern void mouse_event(int flags, int dx, int dy, int data, int extraInfo);

const int MOUSEEVENTF_LEFTDOWN = 0x02;
const int MOUSEEVENTF_LEFTUP = 0x04;
const int MOUSEEVENTF_RIGHTDOWN = 0x08;
const int MOUSEEVENTF_RIGHTUP = 0x10;

[System.Runtime.InteropServices.StructLayout(
    System.Runtime.InteropServices.LayoutKind.Sequential)]
public struct RECT
{
    public int Left;
    public int Top;
    public int Right;
    public int Bottom;
}

[DllImport("user32.dll")]
public static extern bool GetWindowRect(IntPtr hWnd, ref RECT rect);

[DllImport("user32.dll")]
static extern IntPtr GetForegroundWindow();

// -------------------------------
// Reliable Click Anywhere
// -------------------------------
void ClickAt(int x, int y)
{
    Cursor.Position = new System.Drawing.Point(x, y);
    Thread.Sleep(50);

    mouse_event(MOUSEEVENTF_LEFTDOWN, x, y, 0, 0);
    mouse_event(MOUSEEVENTF_LEFTUP, x, y, 0, 0);
}

// -------------------------------
// Click relative to the current application window
// cx, cy are floats between 0..1 (fractional position inside the window rect)
// -------------------------------
void ClickRelative(double cx, double cy)
{
    try
    {
        IntPtr hwnd = (IntPtr)AppContext.Window.Current.NativeWindowHandle;
        if (hwnd == IntPtr.Zero)
        {
            // fallback to foreground window if AppContext not available
            hwnd = GetForegroundWindow();
            if (hwnd == IntPtr.Zero) return;
        }

        RECT r = new RECT();
        if (!GetWindowRect(hwnd, ref r)) return;

        int width = Math.Max(1, r.Right - r.Left);
        int height = Math.Max(1, r.Bottom - r.Top);

        int x = r.Left + (int)Math.Round(Math.Max(0.0, Math.Min(1.0, cx)) * width);
        int y = r.Top + (int)Math.Round(Math.Max(0.0, Math.Min(1.0, cy)) * height);

        ClickAt(x, y);
    }
    catch { }
}

// -------------------------------
// Move Mouse Without Clicking
// -------------------------------
void MoveMouse(int x, int y)
{
    Cursor.Position = new System.Drawing.Point(x, y);
}

// -------------------------------
// Reliable Keyboard Typing
// -------------------------------
void TypeText(string text)
{
    foreach (char c in text)
    {
        SendKeys.SendWait(c.ToString());
        Thread.Sleep(5); // stability
    }
}

// -------------------------------
// Press a specific key
// -------------------------------
void PressKey(string key)
{
    SendKeys.SendWait(key);
    Thread.Sleep(30);
}

// -------------------------------
// Find UI element by name
// -------------------------------
AutomationElement FindUi(string name)
{
    return AppContext.Window.FindFirst(
        TreeScope.Descendants,
        new PropertyCondition(AutomationElement.NameProperty, name)
    );
}

// -------------------------------
// Click a UI element by Name
// -------------------------------
void ClickUi(string name)
{
    var el = FindUi(name);
    if (el == null)
        return;

    try
    {
        var pattern = el.GetCurrentPattern(InvokePattern.Pattern) as InvokePattern;
        pattern?.Invoke();
        Thread.Sleep(150);
    }
    catch {}
}

// -------------------------------
// SafeInvoke - Click by AutomationElement
// -------------------------------
void SafeInvoke(AutomationElement el)
{
    try
    {
        var pattern = el.GetCurrentPattern(InvokePattern.Pattern) as InvokePattern;
        pattern?.Invoke();
        Thread.Sleep(120);
    }
    catch {}
}

// -------------------------------
// Wait helper
// -------------------------------
void Wait(int ms)
{
    Thread.Sleep(ms);
}
