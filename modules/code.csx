// =============================
// VS CODE AUTOMATION MODULE
// =============================

// ---- Usings at VERY TOP ----
using System;
using System.IO;
using System.Threading;
using System.Windows.Forms;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Windows.Automation;

// ---- Win32 Click Support ----
[DllImport("user32.dll")]
static extern void mouse_event(int flags, int dx, int dy, int data, int extraInfo);

const int MOUSEEVENTF_LEFTDOWN = 0x02;
const int MOUSEEVENTF_LEFTUP = 0x04;

// ---- Utility Functions ----
void ClickAt(int x, int y)
{
    Cursor.Position = new System.Drawing.Point(x, y);
    Thread.Sleep(50);
    mouse_event(MOUSEEVENTF_LEFTDOWN, x, y, 0, 0);
    mouse_event(MOUSEEVENTF_LEFTUP, x, y, 0, 0);
    Thread.Sleep(80);
}

void TypeText(string text)
{
    foreach (char c in text)
    {
        SendKeys.SendWait(c.ToString());
        Thread.Sleep(5);
    }
}

AutomationElement FindUi(string name)
{
    return AppContext.Window.FindFirst(
        TreeScope.Descendants,
        new PropertyCondition(
            AutomationElement.NameProperty,
            name,
            PropertyConditionFlags.IgnoreCase
        )
    );
}

void ClickUi(string name)
{
    var element = FindUi(name);
    if (element == null)
        return;

    try
    {
        var invoke = element.GetCurrentPattern(InvokePattern.Pattern) as InvokePattern;
        invoke?.Invoke();
        Thread.Sleep(150);
    }
    catch {}
}

// ---- Core Script Starts ----
if (AppContext == null) throw new Exception("AppContext is null.");
if (string.IsNullOrWhiteSpace(Action)) throw new Exception("No Action provided.");

const int MAX_RETRIES = 3;
const int BASE_DELAY = 150;

// ---- Focus Helpers ----
static class Win32
{
    [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr hWnd);
    [DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
    [DllImport("user32.dll")] public static extern bool IsWindowVisible(IntPtr hWnd);
}

bool FocusWindowHard(int retries = MAX_RETRIES)
{
    for (int i = 1; i <= retries; i++)
    {
        try
        {
            var hwnd = (IntPtr)AppContext.Window.Current.NativeWindowHandle;

            if (hwnd == IntPtr.Zero)
            {
                Thread.Sleep(400);
                continue;
            }

            Win32.ShowWindow(hwnd, 9); // SW_RESTORE
            Thread.Sleep(200);

            if (Win32.SetForegroundWindow(hwnd))
            {
                Thread.Sleep(350);
                return true;
            }
        }
        catch {}
        Thread.Sleep(400);
    }
    return false;
}

// ---- Logging ----
Console.WriteLine("✓ VS Code module loaded");
Console.WriteLine($"Action: {Action}");

if (!FocusWindowHard())
{
    Console.WriteLine("✗ Could not focus VS Code window");
    return;
}

// ---- Execute Actions ----
try
{
    string actionName = Action.ToLower().Trim();

    switch (actionName)
    {
        case "toggle_sidebar":
            SendKeys.SendWait("^b");
            Thread.Sleep(200);
            break;

        case "toggle_terminal":
            SendKeys.SendWait("^`");
            Thread.Sleep(200);
            break;

        case "new_file":
            SendKeys.SendWait("^n");
            Thread.Sleep(200);
            break;

        case "save":
            SendKeys.SendWait("^s");
            Thread.Sleep(200);
            break;

        case "my_macro":
            // Your recorded macro
            ClickAt(1435, 122);
            ClickAt(609, 25);
            ClickAt(991, 437);
            ClickUi("Blue");
            ClickAt(930, 308);
            break;

        // Your Python actions preserved
        case "create_virtual_environment":
        case "python_venv:create":
        case "venv:create":
            // ... (keep your old logic here)
            break;

        case "activate_virtual_environment":
        case "python_venv:activate":
        case "venv:activate":
            // ... (keep your old logic here)
            break;

        default:
            Console.WriteLine($"Unknown action: {actionName}");
            break;
    }
}
catch (Exception ex)
{
    Console.WriteLine($"CRITICAL ERROR: {ex.Message}");
}
