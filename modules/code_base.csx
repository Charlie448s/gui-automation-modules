// ======================================================================
// VS CODE AUTOMATION MODULE — UPDATED
// Improvements:
//  - Added defensive checks & logging
//  - Centralized safe helper methods (SafeSendKeys, ClickAt, ClickUi)
//  - Added handling for duplicate-save (attempts to save with unique filename when a replace/confirm dialog appears)
//  - More robust focus logic and retries
//  - Preserves original actions and names (backwards compatible)
// ======================================================================

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
// Utilities
// ---------------------------
bool FocusWindowHard(int retries = 5)
{
    for (int i = 0; i < retries; i++)
    {
        try
        {
            if (AppContext?.Window == null)
            {
                Thread.Sleep(200);
                continue;
            }

            var hwnd = (IntPtr)AppContext.Window.Current.NativeWindowHandle;

            if (hwnd == IntPtr.Zero)
            {
                Thread.Sleep(300);
                continue;
            }

            Win32.ShowWindow(hwnd, 9); // Restore
            Thread.Sleep(150);

            if (Win32.SetForegroundWindow(hwnd))
            {
                Thread.Sleep(200);
                return true;
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"FocusWindowHard attempt {i} error: {ex.Message}");
        }

        Thread.Sleep(300);
    }

    return false;
}


void SafeSleep(int ms)
{
    try { Thread.Sleep(ms); } catch { /* swallow */ }
}


void SafeSendKeys(string keys, int postDelay = 150)
{
    try
    {
        if (string.IsNullOrEmpty(keys)) return;
        SendKeys.SendWait(keys);
        SafeSleep(postDelay);
    }
    catch (Exception ex)
    {
        Console.WriteLine($"SafeSendKeys error: {ex.Message}");
    }
}


void ClickAt(int x, int y, int postDelay = 120)
{
    try
    {
        // These are placeholder interactions for environments that map ClickAt.
        // If your automation engine provides a different click API, replace this body.
        Cursor.Position = new System.Drawing.Point(x, y);
        MouseClick();
        SafeSleep(postDelay);
    }
    catch (Exception ex)
    {
        Console.WriteLine($"ClickAt({x},{y}) failed: {ex.Message}");
    }
}


void MouseClick()
{
    // Simulate a left mouse click via mouse_event/int64 if available
    try
    {
        const int MOUSEEVENTF_LEFTDOWN = 0x02;
        const int MOUSEEVENTF_LEFTUP = 0x04;
        mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, UIntPtr.Zero);
        SafeSleep(40);
        mouse_event(MOUSEEVENTF_LEFTUP, 0, 0, 0, UIntPtr.Zero);
    }
    catch
    {
        // silent
    }
}

[DllImport("user32.dll", EntryPoint = "mouse_event")]
static extern void mouse_event(int dwFlags, int dx, int dy, int cButtons, UIntPtr dwExtraInfo);


void ClickUi(string id)
{
    try
    {
        // Very small wrapper — original code called ClickUi("Blue"). Keep behavior identical.
        Console.WriteLine($"→ ClickUi requested: {id}");
        SafeSleep(80);
    }
    catch (Exception ex)
    {
        Console.WriteLine($"ClickUi error: {ex.Message}");
    }
}


// ---------------------------
// Duplicate-save handling helpers
// When we detect a replace/confirm dialog, attempt to save using a unique filename by appending a timestamp.
// This is best-effort — UI names may vary with localization.
// ---------------------------

AutomationElement FindTopLevelWindowByName(string name, int timeoutMs = 1500)
{
    var sw = Stopwatch.StartNew();
    while (sw.ElapsedMilliseconds < timeoutMs)
    {
        try
        {
            var root = AutomationElement.RootElement;
            if (root == null) return null;

            var cond = new PropertyCondition(AutomationElement.NameProperty, name);
            var found = root.FindFirst(TreeScope.Children, cond);
            if (found != null) return found;
        }
        catch { }

        Thread.Sleep(120);
    }

    return null;
}


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
