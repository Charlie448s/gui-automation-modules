// code.csx
using System;
using System.Threading;
using System.Windows.Forms;
using System.Windows.Automation;

if (AppContext == null) throw new Exception("AppContext is null inside module.");
if (string.IsNullOrWhiteSpace(Action)) throw new Exception("No Action provided.");

// Helper: focus window
void FocusWindow()
{
    try
    {
        var hwnd = (IntPtr)AppContext.Window.Current.NativeWindowHandle;
        if (hwnd != IntPtr.Zero) Win32.SetForegroundWindow(hwnd);
        Thread.Sleep(200);
    }
    catch { }
}

// Helper: send keystrokes
void Send(string keys, int delay = 150)
{
    SendKeys.SendWait(keys);
    Thread.Sleep(delay);
}

// Win32 import
static class Win32
{
    [System.Runtime.InteropServices.DllImport("user32.dll")]
    public static extern bool SetForegroundWindow(IntPtr hWnd);
}

// Split action into name + param
string actionName = Action;
string actionParam = "";
int colon = Action.IndexOf(':');
if (colon >= 0)
{
    actionName = Action.Substring(0, colon).Trim();
    actionParam = Action.Substring(colon + 1).Trim();
}

Console.WriteLine($"? VS Code Automation Module Loaded!");
Console.WriteLine($" - App Name: code");
Console.WriteLine($" - Process ID: {AppContext.Window.Current.ProcessId}");
Console.WriteLine($" - Automating action: '{actionName}' (param: '{actionParam}')");

FocusWindow();

try
{
    switch (actionName.ToLower())
    {
        case "toggle-sidebar":
        case "toggle_sidebar":
            Send("^b", 300);
            break;

        case "open-terminal":
        case "open_terminal":
            Send("^`", 300);
            break;

        case "command_palette":
            Console.WriteLine(" - Automating: Opening Command Palette...");
            Send("^+p", 300);
            if (!string.IsNullOrEmpty(actionParam))
            {
                Console.WriteLine($" - Automating: Typing '{actionParam}'...");
                Send(actionParam, 300);
                Send("{ENTER}", 300);
            }
            break;

        case "create_python_env":
            Console.WriteLine(" - Automating: Creating Python environment...");
            Send("^+p", 300);
            Send("Python: Create Environment", 400);
            Send("{ENTER}", 400);
            break;

        default:
            Console.WriteLine($"[code.csx] Unknown action: {actionName}");
            break;
    }

    Console.WriteLine(" - Automation task complete.");
}
catch (Exception ex)
{
    Console.WriteLine($"[code.csx] Error executing action: {ex.Message}");
}
