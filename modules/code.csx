#load "_utils.csx"
using System;
using System.IO;
using System.Threading;
using System.Windows.Forms;
using System.Windows.Automation;
using System.Diagnostics;
//hello 
// -------------------------------
// Globals provided by ModuleManager
// -------------------------------
//   AppContext  -> ApplicationContext
//   Action      -> string
// -------------------------------

if (AppContext == null) throw new Exception("AppContext is null.");
if (string.IsNullOrWhiteSpace(Action)) throw new Exception("No Action provided.");

const int MAX_RETRIES = 3;
const int BASE_DELAY = 150;

static class Win32
{
    [System.Runtime.InteropServices.DllImport("user32.dll")]
    public static extern bool SetForegroundWindow(IntPtr hWnd);
    [System.Runtime.InteropServices.DllImport("user32.dll")]
    public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
    [System.Runtime.InteropServices.DllImport("user32.dll")]
    public static extern bool IsWindowVisible(IntPtr hWnd);
}

bool IsWindowResponsive()
{
    try
    {
        var hwnd = (IntPtr)AppContext.Window.Current.NativeWindowHandle;
        return hwnd != IntPtr.Zero && Win32.IsWindowVisible(hwnd);
    }
    catch { return false; }
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
                Console.WriteLine("[vscode.csx] Window focused successfully");
                return true;
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"[vscode.csx] Focus error: {ex.Message}");
        }
        Thread.Sleep(400);
    }
    return false;
}

bool SendKeysWithRetry(string keys, int delay = BASE_DELAY)
{
    try
    {
        SendKeys.SendWait(keys);
        Thread.Sleep(delay);
        return true;
    }
    catch (Exception ex)
    {
        Console.WriteLine($"[vscode.csx] SendKeys error: {ex.Message}");
        return false;
    }
}

Console.WriteLine("✓ VS Code Automation Module Loaded!");
Console.WriteLine($" - App Name : vscode");
Console.WriteLine($" - Process  : {AppContext.Window.Current.ProcessId}");
Console.WriteLine($" - Action   : {Action}");

if (!FocusWindowHard())
{
    Console.WriteLine("[vscode.csx] ERROR: Failed to focus VS Code window. Aborting.");
    return;
}

try
{
    string actionName = Action.ToLower().Trim();

    switch (actionName)
    {
        case "toggle_sidebar":
            Console.WriteLine(" - Toggling sidebar...");
            SendKeysWithRetry("^b", 200);
            Console.WriteLine(" - Sidebar toggled");
            break;

        case "toggle_terminal":
            Console.WriteLine(" - Opening integrated terminal...");
            SendKeysWithRetry("^`", 200);
            Console.WriteLine(" - Terminal opened");
            break;

        case "new_file":
            Console.WriteLine(" - Creating new file...");
            SendKeysWithRetry("^n", 200);
            Console.WriteLine(" - New file created");
            break;

        case "save":
            Console.WriteLine(" - Saving file...");
            SendKeysWithRetry("^s", 200);
            Console.WriteLine(" - File saved");
            break;
         case "my_macro":
        ClickAt(1435, 122);
        ClickAt(609, 25);
        ClickAt(991, 437);
        ClickUi("Blue");
        ClickAt(930, 308);
        break;

        case "create_virtual_environment": // ✅ NEW direct action
        case "python_venv:create":          // ✅ NEW alias
        case "venv:create":
            try
            {
                Console.WriteLine(" - Initiating Python virtual environment creation...");
                FocusWindowHard();
                SendKeysWithRetry("^`", 300);
                Thread.Sleep(800);

                // Clear old terminal text
                SendKeysWithRetry("cls{ENTER}", 300);

                // Execute python venv command
                string createCmd = "python -m venv .venv{ENTER}";
                SendKeysWithRetry(createCmd, 800);

                Console.WriteLine(" - Waiting for environment creation to finish...");
                Thread.Sleep(4000); // allow process to run

                string projectDir = Directory.GetCurrentDirectory();
                string venvPath = Path.Combine(projectDir, ".venv");

                if (Directory.Exists(venvPath) && Directory.Exists(Path.Combine(venvPath, "Scripts")))
                {
                    Console.WriteLine($" - ✓ Virtual environment successfully created at: {venvPath}");
                }
                else
                {
                    Console.WriteLine(" - ✗ Failed to verify .venv creation. Check if Python is installed and accessible.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[vscode.csx] Virtual environment creation failed: {ex.Message}");
            }
            break;

        case "activate_virtual_environment": // ✅ NEW action for activation
        case "python_venv:activate":
        case "venv:activate":
            try
            {
                Console.WriteLine(" - Activating Python virtual environment...");
                FocusWindowHard();
                SendKeysWithRetry("^`", 300);
                Thread.Sleep(600);

                string activateCmd = Environment.OSVersion.Platform == PlatformID.Win32NT
                    ? ".venv\\Scripts\\activate{ENTER}"
                    : "source .venv/bin/activate{ENTER}";
                SendKeysWithRetry(activateCmd, 800);

                Thread.Sleep(1000);
                Console.WriteLine(" - ✓ Virtual environment activation command sent.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[vscode.csx] Activation error: {ex.Message}");
            }
            break;

        default:
            Console.WriteLine($"[vscode.csx] ERROR: Unknown action '{actionName}'");
            Console.WriteLine(" - Available actions: toggle_sidebar, toggle_terminal, new_file, save, create_virtual_environment, activate_virtual_environment");
            break;
    }

    Console.WriteLine(" - Automation task complete.");
}
catch (Exception ex)
{
    Console.WriteLine($"[vscode.csx] CRITICAL ERROR executing action: {ex.Message}");
    Console.WriteLine($"   Stack trace: {ex.StackTrace}");
}
