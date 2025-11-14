// ======================================================================
// VS CODE AUTOMATION MODULE
// This script is executed by the GUI automation engine when the user
// issues an AI-generated command such as "new file", "toggle terminal",
// or a macro like "my_macro".
// ======================================================================

// ----------------------------------------------------------------------
// 1. REQUIRED USING STATEMENTS (MUST BE AT TOP ONLY)
// ----------------------------------------------------------------------
using System;
using System.IO;
using System.Threading;
using System.Windows.Forms;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Windows.Automation;


// ----------------------------------------------------------------------
// 2. WIN32 API IMPORTS FOR MOUSE INPUT (REAL MOUSE CLICKS)
// ----------------------------------------------------------------------
[DllImport("user32.dll")]
static extern void mouse_event(int flags, int dx, int dy, int data, int extraInfo);

const int MOUSEEVENTF_LEFTDOWN = 0x02;
const int MOUSEEVENTF_LEFTUP   = 0x04;


// ======================================================================
// 3. UNIVERSAL UTILITY FUNCTIONS
// These functions support macros, automation, and UI interaction.
// ======================================================================

// ------------------------
// Move + Click at Position
// ------------------------
void ClickAt(int x, int y)
{
    // Move the mouse to the target screen coordinates
    Cursor.Position = new System.Drawing.Point(x, y);

    Thread.Sleep(50); // Allow few ms for stable movement

    // Press and release left mouse button (actual click)
    mouse_event(MOUSEEVENTF_LEFTDOWN, x, y, 0, 0);
    mouse_event(MOUSEEVENTF_LEFTUP,   x, y, 0, 0);

    Thread.Sleep(80); // Give UI time to react
}


// ------------------------
// Type Text Reliably
// ------------------------
void TypeText(string text)
{
    foreach (char c in text)
    {
        SendKeys.SendWait(c.ToString());
        Thread.Sleep(5); // Prevent input overflow
    }
}


// ------------------------
// Find UI Element by Visible Label
// ------------------------
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


// ------------------------
// Click UI Element via Automation
// ------------------------
void ClickUi(string name)
{
    var element = FindUi(name);

    if (element == null)
    {
        Console.WriteLine($"UI element '{name}' not found.");
        return;
    }

    try
    {
        var pattern = element.GetCurrentPattern(InvokePattern.Pattern) as InvokePattern;
        pattern?.Invoke();
        Thread.Sleep(150); // allow UI response
    }
    catch (Exception ex)
    {
        Console.WriteLine($"Error clicking UI element '{name}': {ex.Message}");
    }
}


// ======================================================================
// 4. APP CONTEXT VALIDATION
// Ensures the module is running inside a valid app automation context.
// ======================================================================
if (AppContext == null)
    throw new Exception("AppContext is null — module cannot run.");

if (string.IsNullOrWhiteSpace(Action))
    throw new Exception("No action provided to module.");


// ======================================================================
// 5. WINDOW FOCUS UTILITIES
// Ensures VS Code is brought to the front before automating it.
// ======================================================================
static class Win32
{
    [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr hWnd);
    [DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
    [DllImport("user32.dll")] public static extern bool IsWindowVisible(IntPtr hWnd);
}

bool FocusWindowHard(int retries = 3)
{
    for (int i = 0; i < retries; i++)
    {
        try
        {
            var hwnd = (IntPtr)AppContext.Window.Current.NativeWindowHandle;

            if (hwnd == IntPtr.Zero)
            {
                Thread.Sleep(300);
                continue;
            }

            Win32.ShowWindow(hwnd, 9); // Restore window
            Thread.Sleep(200);

            if (Win32.SetForegroundWindow(hwnd))
            {
                Thread.Sleep(300);
                return true;
            }
        }
        catch
        {
            // ignore errors and retry
        }

        Thread.Sleep(300);
    }

    return false;
}


// ======================================================================
// 6. BOOT LOG — Helps you see module activation
// ======================================================================
Console.WriteLine("✓ VS Code Automation Module Loaded");
Console.WriteLine($"Action requested: {Action}");

if (!FocusWindowHard())
{
    Console.WriteLine("✗ ERROR: Could not focus VS Code window.");
    return;
}


// ======================================================================
// 7. ACTION EXECUTION LOGIC
// Central switch-case where the automation commands are executed.
// ======================================================================
try
{
    string actionName = Action.ToLower().Trim();

    switch (actionName)
    {
        // ----------------------------------------------------------
        // TOGGLE SIDEBAR: Ctrl + B
        // ----------------------------------------------------------
        case "toggle_sidebar":
            Console.WriteLine("→ Toggling sidebar...");
            SendKeys.SendWait("^b");
            break;


        // ----------------------------------------------------------
        // OPEN TERMINAL: Ctrl + `
        // ----------------------------------------------------------
        case "toggle_terminal":
            Console.WriteLine("→ Toggling terminal...");
            SendKeys.SendWait("^`");
            break;


        // ----------------------------------------------------------
        // NEW FILE: Ctrl + N
        // ----------------------------------------------------------
        case "new_file":
            Console.WriteLine("→ Creating new file...");
            SendKeys.SendWait("^n");
            ClickAt(1435, 122);
            ClickAt(609, 25);
            ClickAt(991, 437);
            ClickUi("Blue");
            ClickAt(930, 308);
            break;


        // ----------------------------------------------------------
        // SAVE FILE: Ctrl + S
        // ----------------------------------------------------------
        case "save":
            Console.WriteLine("→ Saving file...");
            SendKeys.SendWait("^s");
            break;


        // ----------------------------------------------------------
        // MACRO INSERTED HERE
        // Corresponds to your RecordedMacro.csx output
        // ----------------------------------------------------------
        case "my_macro":
            Console.WriteLine("→ Running recorded macro...");
            ClickAt(1435, 122);
            ClickAt(609, 25);
            ClickAt(991, 437);
            ClickUi("Blue");
            ClickAt(930, 308);
            break;


        // ----------------------------------------------------------
        // PYTHON ENVIRONMENT CREATION
        // ----------------------------------------------------------
        case "create_virtual_environment":
        case "python_venv:create":
        case "venv:create":
            Console.WriteLine("→ Creating Python virtual environment...");

            SendKeys.SendWait("^`");      // open terminal
            Thread.Sleep(800);

            SendKeys.SendWait("cls{ENTER}");
            Thread.Sleep(300);

            SendKeys.SendWait("python -m venv .venv{ENTER}");
            Thread.Sleep(3500);

            Console.WriteLine("✔ Attempted .venv creation. Verify folder manually.");
            break;


        // ----------------------------------------------------------
        // PYTHON ENV ACTIVATION
        // ----------------------------------------------------------
        case "activate_virtual_environment":
        case "python_venv:activate":
        case "venv:activate":
            Console.WriteLine("→ Activating Python environment...");

            SendKeys.SendWait("^`");
            Thread.Sleep(600);

            string activateCmd =
                Environment.OSVersion.Platform == PlatformID.Win32NT
                ? ".venv\\Scripts\\activate{ENTER}"
                : "source .venv/bin/activate{ENTER}";

            SendKeys.SendWait(activateCmd);
            Thread.Sleep(800);

            Console.WriteLine("✔ Activation command sent.");
            break;


        // ----------------------------------------------------------
        // DEFAULT UNKNOWN COMMAND
        // ----------------------------------------------------------
        default:
            Console.WriteLine($"✗ UNKNOWN ACTION: '{actionName}'");
            break;
    }

    Console.WriteLine("✓ Automation complete.");
}
catch (Exception ex)
{
    Console.WriteLine($"CRITICAL ERROR: {ex.Message}");
}
