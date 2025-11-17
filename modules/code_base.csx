// ======================================================================
// VS CODE AUTOMATION MODULE — UPDATED
// Improvements:
//  - Added defensive checks & logging
//  - Centralized safe helper methods (SafeSendKeys, ClickAt, ClickUi)
//  - Added handling for duplicate-save (attempts to save with unique filename when a replace/confirm dialog appears)
//  - More robust focus logic and retries
//  - Preserves original actions and names (backwards compatible)
// ======================================================================

#load "_utils.csx"

// REQUIRED USING STATEMENTS
using System;
using System.IO;
using System.Threading;
using System.Windows.Forms;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Windows.Automation;


// ---------------------------
// Basic guards for AppContext and Action
// ---------------------------
if (AppContext == null)
    throw new Exception("AppContext is null — module cannot run.");

if (string.IsNullOrWhiteSpace(Action))
    throw new Exception("No action provided to module.");


// ---------------------------
// Win32 helpers
// ---------------------------
static class Win32
{
    [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr hWnd);
    [DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
    [DllImport("user32.dll")] public static extern bool IsWindowVisible(IntPtr hWnd);
}


// ---------------------------
// Utilities (Removed: ClickAt, ClickUi, SafeSendKeys, Mouse helpers)
// Now using utils.csx for all click/keyboard helpers
// ---------------------------

// Duplicate-save handling helpers
bool TryHandleReplaceDialogAndSaveUnique()
{
    // Common titles to detect: "Confirm Save As", "Confirm Save", "Replace File", "Save As"
    string[] candidates = new[] { "Confirm Save As", "Confirm Save", "Replace File", "Save As" };

    foreach (var candidate in candidates)
    {
        var dlg = FindTopLevelWindowByName(candidate, 800);
        if (dlg != null)
        {
            Console.WriteLine($"Detected dialog: {candidate} — attempting to save with unique filename.");

            try
            {
                // Try to find an Edit control for filename inside the dialog
                var edit = dlg.FindFirst(TreeScope.Descendants, new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Edit));
                string uniqueSuffix = DateTime.Now.ToString("yyyyMMdd_HHmmss");

                if (edit != null)
                {
                    var valuePattern = edit.GetCurrentPattern(ValuePattern.Pattern) as ValuePattern;
                    if (valuePattern != null)
                    {
                        string current = valuePattern.Current.Value ?? string.Empty;
                        string newName = current;

                        // If current contains a dot extension, insert suffix before extension
                        try
                        {
                            int lastDot = current.LastIndexOf('.');
                            if (lastDot > 0)
                                newName = current.Substring(0, lastDot) + "_" + uniqueSuffix + current.Substring(lastDot);
                            else
                                newName = current + "_" + uniqueSuffix;

                            valuePattern.SetValue(newName);
                            SafeSleep(120);

                            // Press Enter to confirm
                            SendKeys.SendWait("{ENTER}");
                            SafeSleep(300);

                            Console.WriteLine($"Saved using unique filename: {newName}");
                            return true;
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"Failed to set unique filename in dialog: {ex.Message}");
                        }
                    }
                }

                // If we couldn't find the edit, attempt a simple keyboard approach:
                // Send Ctrl+S again, then type timestamp, Enter.
                SafeSendKeys("^s", 300);
                SafeSendKeys(DateTime.Now.ToString("_yyyyMMdd_HHmmss") + "{ENTER}", 300);

                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error handling replace dialog: {ex.Message}");
            }
        }
    }

    return false;
}


// ---------------------------
// Boot log
// ---------------------------
Console.WriteLine("✓ VS Code Automation Module Loaded (updated)");
Console.WriteLine($"Action requested: {Action}");

if (!FocusWindowHard())
{
    Console.WriteLine("✗ ERROR: Could not focus VS Code window. Continuing but actions may fail.");
}


// ---------------------------
// Action execution
// ---------------------------
try
{
    string actionName = Action.ToLower().Trim();

    switch (actionName)
    {
        case "toggle_sidebar":
            Console.WriteLine("→ Toggling sidebar...");
            SafeSendKeys("^b");
            break;

        case "toggle_terminal":
            Console.WriteLine("→ Toggling terminal...");
            SafeSendKeys("^`");
            break;

        case "new_file":
            Console.WriteLine("→ Creating new file...");
            SafeSendKeys("^n");
            SafeSleep(200);

            // Keep original recorded clicks but protect them with try/catch
            ClickAt(1435, 122);
            ClickAt(609, 25);
            ClickAt(991, 437);
            ClickUi("Blue");
            ClickAt(930, 308);
            break;

        case "save":
            Console.WriteLine("→ Saving file...");
            SafeSendKeys("^s");
            SafeSleep(300);

            // If a replace/confirm dialog appears, try to save with a unique name
            if (TryHandleReplaceDialogAndSaveUnique())
            {
                Console.WriteLine("✔ Resolved duplicate-save by using unique filename.");
            }
            break;

        case "save_with_unique_name":
            // New action — force saving with a unique timestamped filename to avoid replace dialogs
            Console.WriteLine("→ Saving file with unique name...");
            SafeSendKeys("^s");
            SafeSleep(250);
            if (!TryHandleReplaceDialogAndSaveUnique())
            {
                // As a fallback: attempt typing a timestamp and press Enter
                SafeSendKeys(DateTime.Now.ToString("_yyyyMMdd_HHmmss") + "{ENTER}", 300);
                Console.WriteLine("✔ Attempted fallback unique save via keystrokes.");
            }
            break;

        case "my_macro":
            Console.WriteLine("→ Running recorded macro...");
            ClickAt(1435, 122);
            ClickAt(609, 25);
            ClickAt(991, 437);
            ClickUi("Blue");
            ClickAt(930, 308);
            break;

        case "create_virtual_environment":
        case "python_venv:create":
        case "venv:create":
            Console.WriteLine("→ Creating Python virtual environment...");
            SafeSendKeys("^`");      // open terminal
            SafeSleep(800);

            SafeSendKeys("cls{ENTER}");
            SafeSleep(300);

            SafeSendKeys("python -m venv .venv{ENTER}", 3500);

            Console.WriteLine("✔ Attempted .venv creation. Verify folder manually.");
            break;

        case "activate_virtual_environment":
        case "python_venv:activate":
        case "venv:activate":
            Console.WriteLine("→ Activating Python environment...");

            SafeSendKeys("^`");
            SafeSleep(600);

            string activateCmd =
                Environment.OSVersion.Platform == PlatformID.Win32NT
                ? ".venv\\Scripts\\activate{ENTER}"
                : "source .venv/bin/activate{ENTER}";

            SafeSendKeys(activateCmd);
            SafeSleep(800);

            Console.WriteLine("✔ Activation command sent.");
            break;

        default:
            Console.WriteLine($"✗ UNKNOWN ACTION: '{actionName}'");
            break;
    }

    Console.WriteLine("✓ Automation complete.");
}
catch (Exception ex)
{
    Console.WriteLine($"CRITICAL ERROR: {ex.Message}\n{ex.StackTrace}");
}

// ======================================================================
// End of script
// ======================================================================
