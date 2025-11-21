

// ======================================================================
// VS CODE AUTOMATION MODULE
// This script is executed by the GUI automation engine when the user
// issues an AI-generated command such as "new file", "toggle terminal",
// or a macro like "my_macro".
// ======================================================================
//hola amigo
// ----------------------------------------------------------------------
// 1. REQUIRED USING STATEMENTS (MUST BE AT TOP ONLY)
// ----------------------------------------------------------------------
#load "_utils.csx"
using System;
using System.IO;
using System.Threading;
using System.Windows.Forms;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Windows.Automation;




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
         case "duplicate_file":
 case "file:duplicate":
 case "duplicate":
    Console.WriteLine("→ Duplicating current file...");

    // Open Command Palette (Ctrl+Shift+P)
    SendKeys.SendWait("^+p");
    Thread.Sleep(500);

    // Type Save As
    TypeText("File: Save As");
    Thread.Sleep(400);

    SendKeys.SendWait("{ENTER}");
    Thread.Sleep(700);

    // Generate duplicate name with timestamp
    string newName = $"copy_{DateTime.Now:HHmmss}";
    TypeText(newName);
    Thread.Sleep(300);

    SendKeys.SendWait("{ENTER}");
    Thread.Sleep(400);

    Console.WriteLine($"✔ File duplicated as {newName}");
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
"
