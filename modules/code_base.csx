// ======================================================================
// VS CODE AUTOMATION MODULE
// Workflow Fix: Seamless Create -> Activate
// ======================================================================

// ----------------------------------------------------------------------
// 1. REQUIRED USING STATEMENTS
// ----------------------------------------------------------------------
#load "_utils.csx"
using System;
using System.IO;
using System.Threading;
using System.Windows.Forms;
using System.Diagnostics;
using System.Runtime.InteropServices;

// ----------------------------------------------------------------------
// 2. GLOBAL STATE TRACKING
// ----------------------------------------------------------------------
// Single source of truth. true = script believes terminal is visible.
bool terminalOpen = false; 

// ======================================================================
// 3. APP CONTEXT VALIDATION
// ======================================================================
if (AppContext == null)
    throw new Exception("AppContext is null — module cannot run.");

if (string.IsNullOrWhiteSpace(Action))
    throw new Exception("No action provided to module.");

// ======================================================================
// 4. WINDOW FOCUS UTILITIES
// ======================================================================
static class Win32
{
    [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr hWnd);
}

bool FocusWindowHard(int retries = 3)
{
    for (int i = 0; i < retries; i++)
    {
        try
        {
            var hwnd = (IntPtr)AppContext.Window.Current.NativeWindowHandle;
            if (hwnd != IntPtr.Zero)
            {
                if (Win32.SetForegroundWindow(hwnd))
                {
                    Thread.Sleep(200);
                    return true;
                }
            }
        }
        catch { }
        Thread.Sleep(200);
    }
    return false;
}

// ======================================================================
// 5. TERMINAL STATE LOGIC
// ======================================================================

/// <summary>
/// INTELLIGENT TERMINAL CHECK:
/// 1. If 'terminalOpen' is true -> Do nothing (Terminal is ready).
/// 2. If 'terminalOpen' is false -> Toggle it open and wait.
/// </summary>
void EnsureTerminal()
{
    if (!terminalOpen)
    {
        Console.WriteLine("   (Opening terminal window...)");
        SendKeys.SendWait("^`"); // Ctrl + Backtick
        Thread.Sleep(600);       // Wait for animation
        terminalOpen = true;     // Mark as open
    }
    else
    {
        // It's already open. We do nothing so we don't disturb the flow.
        // Just a tiny safety yield.
        Thread.Sleep(50);
    }
}

// ======================================================================
// 6. ACTION EXECUTION
// ======================================================================
Console.WriteLine("✓ VS Code Automation Module Loaded");
Console.WriteLine($"Action requested: {Action}");

if (!FocusWindowHard())
{
    Console.WriteLine("✗ ERROR: Could not focus VS Code window.");
    return;
}

try
{
    string actionName = Action.ToLower().Trim();

    switch (actionName)
    {
        // ----------------------------------------------------------
        // TOGGLE SIDEBAR
        // ----------------------------------------------------------
        case "toggle_sidebar":
            Console.WriteLine("→ Toggling sidebar...");
            SendKeys.SendWait("^b");
            break;

        // ----------------------------------------------------------
        // TOGGLE TERMINAL
        // ----------------------------------------------------------
        case "toggle_terminal":
            Console.WriteLine("→ Toggling terminal...");
            SendKeys.SendWait("^`");
            // We must flip the state here so the tracker stays accurate
            terminalOpen = !terminalOpen; 
            break;

        // ----------------------------------------------------------
        // NEW FILE
        // ----------------------------------------------------------
        case "new_file":
            Console.WriteLine("→ Creating new file...");
            SendKeys.SendWait("^n");
            break;

        // ----------------------------------------------------------
        // SAVE FILE
        // ----------------------------------------------------------
        case "save":
            Console.WriteLine("→ Saving file...");
            SendKeys.SendWait("^s");
            break;

        // ----------------------------------------------------------
        // MACRO
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
        // PYTHON: CREATE VIRTUAL ENVIRONMENT
        // ----------------------------------------------------------
        case "create_virtual_environment":
        case "python_venv:create":
        case "venv:create":
            Console.WriteLine("→ Creating Python virtual environment...");
            
            EnsureTerminal(); // Opens if needed

            SendKeys.SendWait("cls{ENTER}");
            Thread.Sleep(300);

            SendKeys.SendWait("python -m venv .venv{ENTER}");
            
            // Wait for creation (it takes time)
            Thread.Sleep(3500);

            // FIX: Removed 'terminalOpen = false' here.
            // The terminal is STILL open after this finishes, so we leave the state as TRUE.
            
            Console.WriteLine("✔ Virtual environment created.");
            break;

        // ----------------------------------------------------------
        // DUPLICATE FILE
        // ----------------------------------------------------------
        case "duplicate_file":
        case "file:duplicate":
        case "duplicate":
            Console.WriteLine("→ Duplicating current file...");
            SendKeys.SendWait("^+p");
            Thread.Sleep(500);

            TypeText("File: Save As");
            Thread.Sleep(400);

            SendKeys.SendWait("{ENTER}");
            Thread.Sleep(700);

            string newName = $"copy_{DateTime.Now:HHmmss}";
            TypeText(newName);
            Thread.Sleep(300);

            SendKeys.SendWait("{ENTER}");
            Thread.Sleep(400);

            Console.WriteLine($"✔ File duplicated as {newName}");
            break;

        // ----------------------------------------------------------
        // PYTHON: ACTIVATE VIRTUAL ENVIRONMENT
        // ----------------------------------------------------------
        case "activate_virtual_environment":
        case "python_venv:activate":
        case "venv:activate":
            Console.WriteLine("→ Activating Python environment...");

            // Because we fixed the 'Create' step above, this function sees 
            // terminalOpen == true and skips the toggle logic entirely.
            // EnsureTerminal();

            Thread.Sleep(200); // Minimal wait

            string activateCmd = Environment.OSVersion.Platform == PlatformID.Win32NT
                ? ".venv\\Scripts\\activate"
                : "source .venv/bin/activate";

            SendKeys.SendWait(activateCmd);
            Thread.Sleep(100);
            SendKeys.SendWait("{ENTER}");

            Console.WriteLine("✔ Activation command sent.");
            break;

        // ----------------------------------------------------------
        // DEFAULT
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
