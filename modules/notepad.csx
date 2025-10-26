// notepad.csx
using System;
using System.Threading;
using System.Windows.Forms;
using System.Windows.Automation;

if (AppContext == null) throw new Exception("AppContext is null inside module.");
if (string.IsNullOrWhiteSpace(Action)) throw new Exception("No Action provided.");

void Focus()
{
    try
    {
        var hwnd = (IntPtr)AppContext.Window.Current.NativeWindowHandle;
        if (hwnd != IntPtr.Zero) Win32.SetForegroundWindow(hwnd);
        Thread.Sleep(200);
    }
    catch { }
}

void Send(string keys, int delay = 150)
{
    SendKeys.SendWait(keys);
    Thread.Sleep(delay);
}

static class Win32
{
    [System.Runtime.InteropServices.DllImport("user32.dll")]
    public static extern bool SetForegroundWindow(IntPtr hWnd);
}

string actionName = Action;
string actionParam = "";
int colon = Action.IndexOf(':');
if (colon >= 0)
{
    actionName = Action.Substring(0, colon).Trim();
    actionParam = Action.Substring(colon + 1).Trim();
}

Console.WriteLine($"[notepad.csx] Action='{actionName}', Param='{actionParam}'");
Focus();

try
{
    switch (actionName.ToLower())
    {
        case "new_file":
            // Ctrl+N
            Send("^n", 200);
            break;

        case "type":
            if (!string.IsNullOrEmpty(actionParam))
            {
                Send(actionParam, 50);
            }
            else
            {
                Console.WriteLine("[notepad.csx] type action requires a parameter.");
            }
            break;

        case "save_as":
            // Ctrl+Shift+S triggers Save As; then paste path and Enter
            Send("^(+s)", 300); // fallback; sometimes not reliable -> use Alt+F A
            Thread.Sleep(300);
            // fallback to Alt+F, A
            Send("%(f)", 200);
            Thread.Sleep(120);
            Send("a", 400);
            Thread.Sleep(600);

            if (!string.IsNullOrEmpty(actionParam))
            {
                // send the path
                Send(actionParam, 300);
                Send("{ENTER}", 300);
            }
            else
            {
                Console.WriteLine("[notepad.csx] save_as requires full path parameter.");
            }
            break;

        case "close":
            // Alt+F4
            Send("%{F4}", 200);
            break;

        default:
            Console.WriteLine($"[notepad.csx] Unknown action: {actionName}");
            break;
    }
}
catch (Exception ex)
{
    Console.WriteLine($"[notepad.csx] Error executing action: {ex.Message}");
}
