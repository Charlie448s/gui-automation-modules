// ============================
// UNIVERSAL UTILITIES MODULE
// ============================
// Loaded by all modules automatically using: #load "_utils.csx"
// Provides: ClickAt, ClickUi, TypeText, PressKey, Wait, FindUi, MoveMouse, SafeInvoke
// ============================

using System;
using System.Threading;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using System.Windows.Automation;

// -------------------------------
// Win32 API for mouse events
// -------------------------------
[DllImport("user32.dll")]
static extern void mouse_event(int flags, int dx, int dy, int data, int extraInfo);

const int MOUSEEVENTF_LEFTDOWN = 0x02;
const int MOUSEEVENTF_LEFTUP = 0x04;
const int MOUSEEVENTF_RIGHTDOWN = 0x08;
const int MOUSEEVENTF_RIGHTUP = 0x10;

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
