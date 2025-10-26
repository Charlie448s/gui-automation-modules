// code.csx
using System;
using System.Threading;
using System.Windows.Forms;
using System.Windows.Automation;

// Globals available:
//   AppContext  -> ApplicationContext
//   Action      -> string

if (AppContext == null) throw new Exception("AppContext is null inside module.");
if (string.IsNullOrWhiteSpace(Action)) throw new Exception("No Action provided.");

void FocusWindow()
{
    try
    {
        var hwnd = (IntPtr)AppContext.Window.Current.NativeWindowHandle;
        if (hwnd != IntPtr.Zero)
        {
            // Bring to front (best-effort)
            Win32.SetForegroundWindow(hwnd);
            Thread.Sleep(200);
        }
    }
    catch { /* best-effort */ }
}

void Send(string keys, int after = 150)
{
    SendKeys.SendWait(keys);
    Thread.Sleep(after);
}

// Simple Win32 wrapper used only here
static class Win32
{
    [System.Runtime.InteropServices.DllImport("user32.dll")]
    public static extern bool SetForegroundWindow(IntPtr hWnd);
}

// Parse action that could include ':' parameter (e.g., command_palette:Python: Create Environment)
string actionName = Action;
string actionParam = "";
int colon = Action.IndexOf(':');
if (colon >= 0)
{
    actionName = Action.Substring(0, colon).Trim();
    actionParam = Action.Substring(colon + 1).Trim();
}

Console.WriteLine($"[code.csx] Action='{actionName}', Param='{actionParam}'");
FocusWindow();

try
{
    switch (actionName.ToLower())
    {
        case "open_terminal":
            // Ctrl+` opens integrated terminal
            Send("^`", 500);
            break;

        case "toggle_sidebar":
            // Ctrl+B toggles sidebar
            Send("^b", 250);
            break;

        case "toggle_panel":
            // Ctrl+J toggles panel
            Send("^j", 250);
            break;

        case "command_palette":
            // open command palette then type param + enter
            Send("^+p", 350);
            if (!string.IsNullOrEmpty(actionParam))
            {
                Send(actionParam, 300);
                Send("{ENTER}", 500);
            }
            break;

        case "create_python_env":
            // Try open command palette and run Python: Create Environment
            Send("^+p", 400);
            Send("Python: Create Environment", 500);
            Send("{ENTER}", 500);
            break;

        case "open_file":
            // actionParam expected to be relative/full path or filename
            // Use File: Open... via command palette for reliability
            Send("^+p", 350);
            Send($"File: Open File...", 300);
            Send("{ENTER}", 400);
            Thread.Sleep(700);
            if (!string.IsNullOrEmpty(actionParam))
            {
                Send(actionParam, 300);
                Send("{ENTER}", 400);
            }
            break;

        default:
            Console.WriteLine($"[code.csx] Unknown action: {actionName}");
            break;
    }
}
catch (Exception ex)
{
    Console.WriteLine($"[code.csx] Error executing action: {ex.Message}");
}
