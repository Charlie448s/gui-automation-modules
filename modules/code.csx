// Robust Visual Studio Code automation module with comprehensive error handling
// Enhanced version with retry mechanisms, validation, and cross-platform support
// Compatible with ModuleManager.cs
//Version 2

using System;
using System.IO;
using System.Threading;
using System.Windows.Forms;
using System.Windows.Automation;
using System.Diagnostics;

// -------------------------------
// Globals provided by ModuleManager
// -------------------------------
//   AppContext  -> ApplicationContext
//   Action      -> string
// -------------------------------

if (AppContext == null) throw new Exception("AppContext is null.");
if (string.IsNullOrWhiteSpace(Action)) throw new Exception("No Action provided.");

// ----------- Configuration -----------
const int MAX_RETRIES = 3;
const int BASE_DELAY = 150;
const int OPERATION_TIMEOUT = 5000;

// ----------- Enhanced Utility Helpers -----------

bool IsWindowResponsive()
{
    try
    {
        var hwnd = (IntPtr)AppContext.Window.Current.NativeWindowHandle;
        return hwnd != IntPtr.Zero && Win32.IsWindowVisible(hwnd);
    }
    catch (Exception ex)
    {
        Console.WriteLine($"[vscode.csx] Window responsiveness check failed: {ex.Message}");
        return false;
    }
}

bool FocusWindowHard(int retries = MAX_RETRIES)
{
    for (int attempt = 1; attempt <= retries; attempt++)
    {
        try
        {
            if (!IsWindowResponsive())
            {
                Console.WriteLine($"[vscode.csx] Window not responsive on attempt {attempt}");
                Thread.Sleep(500 * attempt);
                continue;
            }

            var hwnd = (IntPtr)AppContext.Window.Current.NativeWindowHandle;
            if (hwnd == IntPtr.Zero)
            {
                Console.WriteLine($"[vscode.csx] Invalid window handle on attempt {attempt}");
                Thread.Sleep(300 * attempt);
                continue;
            }

            Win32.ShowWindow(hwnd, 9);   // SW_RESTORE
            Thread.Sleep(200);
            
            if (Win32.SetForegroundWindow(hwnd))
            {
                Thread.Sleep(350);
                Console.WriteLine($"[vscode.csx] Window focused successfully");
                return true;
            }
            
            Console.WriteLine($"[vscode.csx] SetForegroundWindow failed on attempt {attempt}");
            Thread.Sleep(300 * attempt);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"[vscode.csx] FocusWindowHard error (attempt {attempt}): {ex.Message}");
            if (attempt == retries) return false;
            Thread.Sleep(500 * attempt);
        }
    }
    return false;
}

bool SendKeysWithRetry(string keys, int delay = BASE_DELAY, int retries = 2)
{
    for (int attempt = 1; attempt <= retries; attempt++)
    {
        try
        {
            if (!IsWindowResponsive())
            {
                Console.WriteLine($"[vscode.csx] Window not responsive before SendKeys (attempt {attempt})");
                Thread.Sleep(300);
                continue;
            }

            SendKeys.SendWait(keys);
            Thread.Sleep(delay);
            return true;
        }
        catch (Exception ex)
        {
            Console.WriteLine($"[vscode.csx] SendKeys error (attempt {attempt}): {ex.Message}");
            if (attempt == retries) return false;
            Thread.Sleep(200 * attempt);
        }
    }
    return false;
}

bool WaitForClipboard(int maxWaitMs = 2000)
{
    var stopwatch = Stopwatch.StartNew();
    while (stopwatch.ElapsedMilliseconds < maxWaitMs)
    {
        try
        {
            string text = Clipboard.GetText();
            if (!string.IsNullOrEmpty(text))
                return true;
        }
        catch { }
        Thread.Sleep(100);
    }
    return false;
}

string SafeGetClipboardText(int retries = 3)
{
    for (int attempt = 1; attempt <= retries; attempt++)
    {
        try
        {
            return Clipboard.GetText() ?? string.Empty;
        }
        catch (Exception ex)
        {
            Console.WriteLine($"[vscode.csx] Clipboard read error (attempt {attempt}): {ex.Message}");
            if (attempt < retries) Thread.Sleep(200);
        }
    }
    return string.Empty;
}

void ClearClipboard()
{
    try
    {
        Clipboard.Clear();
    }
    catch (Exception ex)
    {
        Console.WriteLine($"[vscode.csx] Clipboard clear warning: {ex.Message}");
    }
}

static class Win32
{
    [System.Runtime.InteropServices.DllImport("user32.dll")]
    public static extern bool SetForegroundWindow(IntPtr hWnd);

    [System.Runtime.InteropServices.DllImport("user32.dll")]
    public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

    [System.Runtime.InteropServices.DllImport("user32.dll")]
    public static extern bool IsWindowVisible(IntPtr hWnd);
}

// ----------- Action parsing -----------

string actionName = Action ?? string.Empty;
string actionParam = string.Empty;
int colon = Action.IndexOf(':');

if (colon >= 0)
{
    actionName = Action.Substring(0, colon).Trim();
    actionParam = Action.Substring(colon + 1).Trim();
}

Console.WriteLine($"✓ VS Code Automation Module Loaded!");
Console.WriteLine($" - App Name : vscode");
Console.WriteLine($" - Process  : {AppContext.Window.Current.ProcessId}");
Console.WriteLine($" - Action   : {actionName}");
if (!string.IsNullOrEmpty(actionParam))
    Console.WriteLine($" - Parameter: {actionParam}");

if (!FocusWindowHard())
{
    Console.WriteLine("[vscode.csx] ERROR: Failed to focus VS Code window. Aborting.");
    return;
}

// ----------- Main automation logic with robust error handling -----------

try
{
    switch (actionName.ToLower())
    {
        case "command_palette":
        case "palette":
        case "open_palette":
            try
            {
                Console.WriteLine(" - Opening Command Palette...");
                if (!SendKeysWithRetry("^+p", 400))
                {
                    Console.WriteLine(" - ERROR: Failed to open Command Palette");
                    break;
                }

                if (!string.IsNullOrEmpty(actionParam))
                {
                    Console.WriteLine($" - Typing command '{actionParam}'...");
                    if (!SendKeysWithRetry(actionParam, 300))
                    {
                        Console.WriteLine(" - ERROR: Failed to type command");
                        break;
                    }
                    if (!SendKeysWithRetry("{ENTER}", 200))
                    {
                        Console.WriteLine(" - ERROR: Failed to execute command");
                        break;
                    }
                }
                Console.WriteLine(" - Command Palette operation completed");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[vscode.csx] Command Palette error: {ex.Message}");
            }
            break;

        case "menu":
        case "run_menu":
            try
            {
                if (string.IsNullOrEmpty(actionParam))
                {
                    Console.WriteLine(" - ERROR: Menu command requires a parameter");
                    break;
                }

                Console.WriteLine($" - Searching menu option '{actionParam}'...");
                if (!SendKeysWithRetry("^+p", 400))
                {
                    Console.WriteLine(" - ERROR: Failed to open Command Palette");
                    break;
                }
                if (!SendKeysWithRetry(actionParam, 300))
                {
                    Console.WriteLine(" - ERROR: Failed to type menu command");
                    break;
                }
                if (!SendKeysWithRetry("{ENTER}", 200))
                {
                    Console.WriteLine(" - ERROR: Failed to execute menu command");
                    break;
                }
                Console.WriteLine(" - Menu command executed");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[vscode.csx] Menu command error: {ex.Message}");
            }
            break;

        case "toggle_sidebar":
            try
            {
                Console.WriteLine(" - Toggling sidebar...");
                if (SendKeysWithRetry("^b", 200))
                    Console.WriteLine(" - Sidebar toggled");
                else
                    Console.WriteLine(" - ERROR: Failed to toggle sidebar");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[vscode.csx] Toggle sidebar error: {ex.Message}");
            }
            break;

        case "toggle_panel":
            try
            {
                Console.WriteLine(" - Toggling panel...");
                if (SendKeysWithRetry("^j", 200))
                    Console.WriteLine(" - Panel toggled");
                else
                    Console.WriteLine(" - ERROR: Failed to toggle panel");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[vscode.csx] Toggle panel error: {ex.Message}");
            }
            break;

        case "toggle_terminal":
        case "open_terminal":
            try
            {
                Console.WriteLine(" - Opening integrated terminal...");
                if (SendKeysWithRetry("^`", 200))
                    Console.WriteLine(" - Terminal opened");
                else
                    Console.WriteLine(" - ERROR: Failed to open terminal");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[vscode.csx] Terminal error: {ex.Message}");
            }
            break;

        case "new_file":
            try
            {
                Console.WriteLine(" - Creating new file...");
                if (SendKeysWithRetry("^n", 200))
                    Console.WriteLine(" - New file created");
                else
                    Console.WriteLine(" - ERROR: Failed to create new file");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[vscode.csx] New file error: {ex.Message}");
            }
            break;

        case "save":
            try
            {
                Console.WriteLine(" - Saving current file...");
                if (SendKeysWithRetry("^s", 200))
                    Console.WriteLine(" - File saved");
                else
                    Console.WriteLine(" - ERROR: Failed to save file");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[vscode.csx] Save error: {ex.Message}");
            }
            break;

        case "find":
            try
            {
                Console.WriteLine(" - Opening Find dialog...");
                if (SendKeysWithRetry("^f", 200))
                    Console.WriteLine(" - Find dialog opened");
                else
                    Console.WriteLine(" - ERROR: Failed to open Find dialog");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[vscode.csx] Find error: {ex.Message}");
            }
            break;

        case "replace":
            try
            {
                Console.WriteLine(" - Opening Replace dialog...");
                if (SendKeysWithRetry("^h", 200))
                    Console.WriteLine(" - Replace dialog opened");
                else
                    Console.WriteLine(" - ERROR: Failed to open Replace dialog");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[vscode.csx] Replace error: {ex.Message}");
            }
            break;

        case "go_to_file":
            try
            {
                Console.WriteLine(" - Opening Quick Open (Go to File)...");
                if (SendKeysWithRetry("^p", 200))
                    Console.WriteLine(" - Quick Open dialog opened");
                else
                    Console.WriteLine(" - ERROR: Failed to open Quick Open");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[vscode.csx] Go to file error: {ex.Message}");
            }
            break;

        case "run_command":
            try
            {
                if (string.IsNullOrEmpty(actionParam))
                {
                    Console.WriteLine(" - ERROR: run_command requires a parameter");
                    break;
                }

                Console.WriteLine($" - Running '{actionParam}' via Command Palette...");
                if (!SendKeysWithRetry("^+p", 400))
                {
                    Console.WriteLine(" - ERROR: Failed to open Command Palette");
                    break;
                }
                if (!SendKeysWithRetry(actionParam, 300))
                {
                    Console.WriteLine(" - ERROR: Failed to type command");
                    break;
                }
                if (!SendKeysWithRetry("{ENTER}", 200))
                {
                    Console.WriteLine(" - ERROR: Failed to execute command");
                    break;
                }
                Console.WriteLine(" - Command executed successfully");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[vscode.csx] Run command error: {ex.Message}");
            }
            break;

        case "list_commands":
            try
            {
                Console.WriteLine(" - Developer commands snapshot:");
                Console.WriteLine("   New File, Open Folder..., Save, Save All, Toggle Sidebar, Toggle Panel,");
                Console.WriteLine("   Run Task..., Debug: Start, Run Without Debugging, New Terminal,");
                Console.WriteLine("   Extensions: Install, Git: Commit, Git: Push, Git: Pull, AI Chat: Open,");
                Console.WriteLine("   Run All Tests, Go to Definition, Go to Symbol in Workspace..., etc.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[vscode.csx] List commands error: {ex.Message}");
            }
            break;

        case "python_venv":
        case "venv":
            try
            {
                if (string.IsNullOrEmpty(actionParam))
                {
                    Console.WriteLine(" - ERROR: Virtual environment command requires a parameter");
                    Console.WriteLine("   Usage: python_venv:create | python_venv:activate | python_venv:install <package>");
                    break;
                }

                if (!FocusWindowHard())
                {
                    Console.WriteLine(" - ERROR: Failed to focus window for terminal operations");
                    break;
                }

                Console.WriteLine(" - Opening terminal...");
                if (!SendKeysWithRetry("^`", 400))
                {
                    Console.WriteLine(" - ERROR: Failed to open terminal");
                    break;
                }

                Thread.Sleep(700);
                SendKeysWithRetry("cls{ENTER}", 300);

                string[] paramParts = actionParam.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
                if (paramParts.Length == 0)
                {
                    Console.WriteLine(" - ERROR: Invalid venv parameter");
                    break;
                }

                string venvCommand = paramParts[0].ToLower();

                switch (venvCommand)
                {
                    case "create":
                        Console.WriteLine(" - Creating virtual environment (.venv)...");
                        if (SendKeysWithRetry("python -m venv .venv{ENTER}", 1000))
                            Console.WriteLine(" - Virtual environment creation initiated");
                        else
                            Console.WriteLine(" - ERROR: Failed to create virtual environment");
                        break;

                    case "activate":
                        Console.WriteLine(" - Activating virtual environment...");
                        string activateCmd = Environment.OSVersion.Platform == PlatformID.Win32NT
                            ? ".venv\\Scripts\\activate{ENTER}"
                            : "source .venv/bin/activate{ENTER}";
                        
                        if (SendKeysWithRetry(activateCmd, 800))
                            Console.WriteLine(" - Virtual environment activation initiated");
                        else
                            Console.WriteLine(" - ERROR: Failed to activate virtual environment");
                        break;

                    case "install":
                        if (paramParts.Length < 2)
                        {
                            Console.WriteLine(" - ERROR: Please specify a package to install");
                            Console.WriteLine("   Usage: python_venv:install <package_name>");
                            break;
                        }

                        string pkg = string.Join(" ", paramParts, 1, paramParts.Length - 1);
                        Console.WriteLine($" - Installing package: {pkg}");
                        if (SendKeysWithRetry($"pip install {pkg}{{ENTER}}", 1000))
                            Console.WriteLine(" - Package installation initiated");
                        else
                            Console.WriteLine(" - ERROR: Failed to install package");
                        break;

                    default:
                        Console.WriteLine($" - ERROR: Unknown virtual environment command: {venvCommand}");
                        Console.WriteLine("   Available: create, activate, install <package>");
                        break;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[vscode.csx] Virtual environment error: {ex.Message}");
            }
            break;

        case "duplicate_file":
            try
            {
                if (!FocusWindowHard())
                {
                    Console.WriteLine(" - ERROR: Failed to focus window");
                    break;
                }

                Console.WriteLine(" - Attempting to duplicate current file...");
                ClearClipboard();

                if (!SendKeysWithRetry("^+p", 600))
                {
                    Console.WriteLine(" - ERROR: Failed to open Command Palette");
                    break;
                }

                Thread.Sleep(600);
                
                if (!SendKeysWithRetry("File: Copy Path of Active File{ENTER}", 800))
                {
                    Console.WriteLine(" - ERROR: Failed to execute copy path command");
                    break;
                }

                if (!WaitForClipboard(2000))
                {
                    Console.WriteLine(" - ERROR: Clipboard did not receive file path");
                    break;
                }

                Thread.Sleep(500);
                string filePath = SafeGetClipboardText();

                if (string.IsNullOrWhiteSpace(filePath))
                {
                    Console.WriteLine(" - ERROR: Could not retrieve file path from clipboard");
                    break;
                }

                filePath = filePath.Trim();
                Console.WriteLine($" - Current file path: {filePath}");

                if (!File.Exists(filePath))
                {
                    Console.WriteLine($" - ERROR: File does not exist: {filePath}");
                    break;
                }

                string dir = Path.GetDirectoryName(filePath);
                if (string.IsNullOrEmpty(dir))
                {
                    Console.WriteLine(" - ERROR: Could not determine file directory");
                    break;
                }

                string name = Path.GetFileNameWithoutExtension(filePath);
                string ext = Path.GetExtension(filePath);
                string dupFilePath = Path.Combine(dir, $"{name}(dup){ext}");

                int dupCounter = 1;
                while (File.Exists(dupFilePath))
                {
                    dupFilePath = Path.Combine(dir, $"{name}(dup{dupCounter}){ext}");
                    dupCounter++;
                    if (dupCounter > 100)
                    {
                        Console.WriteLine(" - ERROR: Too many duplicate files exist");
                        break;
                    }
                }

                File.Copy(filePath, dupFilePath, overwrite: false);
                Console.WriteLine($" - ✓ File duplicated successfully: {Path.GetFileName(dupFilePath)}");
                
                ClearClipboard();
            }
            catch (UnauthorizedAccessException ex)
            {
                Console.WriteLine($"[vscode.csx] Permission denied: {ex.Message}");
            }
            catch (IOException ex)
            {
                Console.WriteLine($"[vscode.csx] File I/O error: {ex.Message}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[vscode.csx] Duplicate file error: {ex.Message}");
            }
            finally
            {
                ClearClipboard();
            }
            break;

        default:
            Console.WriteLine($"[vscode.csx] ERROR: Unknown action '{actionName}'");
            Console.WriteLine(" - Available actions: command_palette, toggle_sidebar, toggle_panel,");
            Console.WriteLine("   toggle_terminal, new_file, save, find, replace, go_to_file,");
            Console.WriteLine("   python_venv, duplicate_file, list_commands");
            break;
    }

    Console.WriteLine(" - Automation task complete.");
}
catch (Exception ex)
{
    Console.WriteLine($"[vscode.csx] CRITICAL ERROR executing action: {ex.Message}");
    Console.WriteLine($"   Stack trace: {ex.StackTrace}");
}
