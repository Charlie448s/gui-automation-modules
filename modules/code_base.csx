

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
 // Required for Clipboard access





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

string EscapeSendKeys(string s)
{
    return s
        .Replace("{", "{{}")
        .Replace("}", "{}}")
        .Replace("(", "{(}")
        .Replace(")", "{)}");
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
case "duplicate_file":
case "file:duplicate":
case "duplicate":
    Console.WriteLine("→ Duplicating current file...");

    // 1. GET CURRENT FILE PATH (via VSCode Command -> clipboard)
    SendKeys.SendWait("^+p");
    Thread.Sleep(200);
    TypeText("File: Copy Path of Active File");
    Thread.Sleep(200);
    SendKeys.SendWait("{ENTER}");
    Thread.Sleep(250);

    string fullPath = GetClipboardTextWithRetries(5, 200);

    if (string.IsNullOrEmpty(fullPath) || !File.Exists(fullPath))
    {
        Console.WriteLine("❌ Could not get valid file path from clipboard.");
        break;
    }

    string directory = Path.GetDirectoryName(fullPath);
    string fileNameNoExt = Path.GetFileNameWithoutExtension(fullPath);
    string extension = Path.GetExtension(fullPath);

    // 2. GENERATE NEW FILE NAME (hello_copy_1.py)
    int count = 1;
    string newFullPath;
    string newFileNameOnly;

    do
    {
        newFileNameOnly = $"{fileNameNoExt}_copy_{count}{extension}";
        newFullPath = Path.Combine(directory, newFileNameOnly);
        count++;
    } while (File.Exists(newFullPath));

    // 3. FIRST & BEST OPTION: copy file directly (avoid clipboard & sendkeys)
    try
    {
        File.Copy(fullPath, newFullPath);
        Console.WriteLine($"✔ File duplicated as {newFileNameOnly} (copied directly).");

        // Optional: open the new file in VS Code by asking VS Code to open it.
        // This uses the command palette to open the file by path.
        // If you don't want to auto-open, you can comment out the block below.
        Thread.Sleep(200);
        SendKeys.SendWait("^+p");
        Thread.Sleep(200);
        TypeText("File: Open File...");
        Thread.Sleep(200);
        SendKeys.SendWait("{ENTER}");
        Thread.Sleep(400);
        // Use a clipboard helper to paste the path into the Open dialog
        if (SetClipboardTextWithRetries(newFullPath, 5, 200))
        {
            SendKeys.SendWait("^v");
            Thread.Sleep(200);
            SendKeys.SendWait("{ENTER}");
        }

        break; // done
    }
    catch (Exception ex)
    {
        // Could be permission / locked file / IO issue — fall back to Save As approach
        Console.WriteLine($"⚠ Direct file copy failed ({ex.GetType().Name}): {ex.Message}");
        Console.WriteLine("→ Falling back to Save As dialog (will attempt robust clipboard).");
    }

    // 4. FALLBACK: use Save As dialog + robust clipboard helper (STA + retries)
    bool clipboardSet = SetClipboardTextWithRetries(newFullPath, 8, 200);

    if (!clipboardSet)
    {
        Console.WriteLine("❌ Failed to set clipboard after retries. Cannot complete Save As fallback.");
        break;
    }

    // Open Save As in VS Code
    SendKeys.SendWait("^+p");
    Thread.Sleep(300);

    TypeText("File: Save As");
    Thread.Sleep(300);
    SendKeys.SendWait("{ENTER}");
    Thread.Sleep(900); // give dialog time

    // Paste the path
    SendKeys.SendWait("^v");
    Thread.Sleep(300);
    SendKeys.SendWait("{ENTER}");
    Thread.Sleep(500);

    Console.WriteLine($"✔ File duplicated as {newFileNameOnly} (via Save As fallback).");
    break;

// ----------------- Helper functions -----------------

// Attempts to get clipboard text with retries
string GetClipboardTextWithRetries(int maxAttempts, int delayMs)
{
    for (int i = 0; i < maxAttempts; i++)
    {
        try
        {
            string text = GetClipboardTextSTA();
            if (!string.IsNullOrEmpty(text)) return text;
        }
        catch { /* ignore and retry */ }

        Thread.Sleep(delayMs);
    }
    return null;
}

// Attempts to set clipboard text with retries
bool SetClipboardTextWithRetries(string text, int maxAttempts, int delayMs)
{
    for (int i = 0; i < maxAttempts; i++)
    {
        try
        {
            SetClipboardTextSTA(text);
            // verify
            string verify = GetClipboardTextSTA();
            if (verify == text) return true;
        }
        catch { /* ignore and retry */ }

        Thread.Sleep(delayMs);
    }
    return false;
}

// STA-thread wrapper to get clipboard text (avoids "Clipboard operation did not succeed" in many cases)
string GetClipboardTextSTA()
{
    string result = null;
    Thread thread = new Thread(() =>
    {
        try
        {
            if (Clipboard.ContainsText())
                result = Clipboard.GetText();
        }
        catch { /* access denied or in use */ }
    });
    thread.SetApartmentState(ApartmentState.STA);
    thread.Start();
    thread.Join(1000); // wait up to 1s
    return result;
}

// STA-thread wrapper to set clipboard text
void SetClipboardTextSTA(string text)
{
    Thread thread = new Thread(() =>
    {
        try
        {
            Clipboard.SetText(text);
        }
        catch
        {
            // Let outer retry logic handle failures
            throw;
        }
    });
    thread.SetApartmentState(ApartmentState.STA);
    thread.Start();
    thread.Join(1000);
}



case "create_virtual_environment":
        case "python_venv:create":
        case "venv:create":
            Console.WriteLine("→ Creating Python virtual environment...");

            // FIX: Use Command Palette to ensure a NEW terminal is opened
            // (Prevents toggling closed if already open)
            SendKeys.SendWait("^+p");       // Ctrl + Shift + P
            Thread.Sleep(400);
            
            // We use SendKeys to type the command palette search
            SendKeys.SendWait("Terminal: Create New Terminal"); 
            Thread.Sleep(300);
            SendKeys.SendWait("{ENTER}");
            Thread.Sleep(1500);             // Wait for terminal to initialize

            // Now send the creation command
            SendKeys.SendWait("python -m venv .venv{ENTER}");
            Thread.Sleep(3500);             // Wait for venv creation to finish

            Console.WriteLine("✔ Attempted .venv creation.");
            break;

        // ----------------------------------------------------------
        // PYTHON ENV ACTIVATION
        // ----------------------------------------------------------
        case "activate_virtual_environment":
        case "python_venv:activate":
        case "venv:activate":
            Console.WriteLine("→ Activating Python environment...");

            // FIX: Use Command Palette to FOCUS the existing terminal
            // (Prevents closing it if it's already open)
            SendKeys.SendWait("^+p");       // Ctrl + Shift + P
            Thread.Sleep(400);
            
            SendKeys.SendWait("Terminal: Focus Terminal");
            Thread.Sleep(300);
            SendKeys.SendWait("{ENTER}");
            Thread.Sleep(800);              // Wait for focus switch

            string activateCmd =
                Environment.OSVersion.Platform == PlatformID.Win32NT
                ? ".venv\\Scripts\\activate{ENTER}"
                : "source .venv/bin/activate{ENTER}";

            SendKeys.SendWait(activateCmd);
            Thread.Sleep(800);

            Console.WriteLine("✔ Activation command sent.");
            break;
        // ----------------------------------------------------------
        // PYTHON ENV ACTIVATION
        // ----------------------------------------------------------
case "open_gitbash":
case "terminal:gitbash":
case "gitbash":
{
    Console.WriteLine("→ Opening Git Bash terminal...");

    // 1. Open Command Palette
    SendKeys.SendWait("^+p");
    Thread.Sleep(200);

    // 2. Use the UNIQUE command that ALWAYS opens the profile picker
    TypeText("Terminal: Select Default Profile");
    Thread.Sleep(300);

    SendKeys.SendWait("{ENTER}");
    Thread.Sleep(700); // time for the profile list to open

    // 3. Select Git Bash
    TypeText("Git Bash");
    Thread.Sleep(300);

    SendKeys.SendWait("{ENTER}");

    // 4. NOW create a new terminal using that default profile
    SendKeys.SendWait("^+p");
    Thread.Sleep(200);

    TypeText("Terminal: Create New Terminal");
    Thread.Sleep(200);

    SendKeys.SendWait("{ENTER}");

    Console.WriteLine("✔ Git Bash opened.");
    break;
}




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


 
