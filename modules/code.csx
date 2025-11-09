// vscode.csx
// Robust Visual Studio Code automation module with comprehensive error handling
// Enhanced with retry mechanisms, validation, and cross-platform support
// Compatible with ModuleManager.cs
// Version 5 7

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

// ----------- Config -----------
const int MAX_RETRIES = 3;
const int BASE_DELAY = 150;

// ----------- Win32 helpers -----------
static class Win32
{
    [System.Runtime.InteropServices.DllImport("user32.dll")]
    public static extern bool SetForegroundWindow(IntPtr hWnd);

    [System.Runtime.InteropServices.DllImport("user32.dll")]
    public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

    [System.Runtime.InteropServices.DllImport("user32.dll")]
    public static extern bool IsWindowVisible(IntPtr hWnd);
}

// ----------- Utility functions -----------
bool IsWindowResponsive()
{
    try
    {
        var hwnd = (IntPtr)AppContext.Window.Current.NativeWindowHandle;
        return hwnd != IntPtr.Zero && Win32.IsWindowVisible(hwnd);
    }
    catch (Exception ex)
    {
        Console.WriteLine($"[vscode.csx] Window check failed: {ex.Message}");
        return false;
    }
}

bool FocusWindowHard(int retries = MAX_RETRIES)
{
    for (int attempt = 1; attempt <= retries; attempt++)
    {
        try
        {
            var hwnd = (IntPtr)AppContext.Window.Current.NativeWindowHandle;
            if (hwnd == IntPtr.Zero)
            {
                Console.WriteLine("Invalid VS Code handle.");
                Thread.Sleep(300 * attempt);
                continue;
            }

            Win32.ShowWindow(hwnd, 9);
            Thread.Sleep(200);

            if (Win32.SetForegroundWindow(hwnd))
            {
                Thread.Sleep(300);
                return true;
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"[vscode.csx] Focus attempt {attempt} failed: {ex.Message}");
        }
        Thread.Sleep(400);
    }
    return false;
}

bool SendKeysWithRetry(string keys, int delay = BASE_DELAY, int retries = 2)
{
    for (int i = 0; i < retries; i++)
    {
        try
        {
            if (!IsWindowResponsive()) Thread.Sleep(200);
            SendKeys.SendWait(keys);
            Thread.Sleep(delay);
            return true;
        }
        catch (Exception ex)
        {
            Console.WriteLine($"[vscode.csx] SendKeys error: {ex.Message}");
            Thread.Sleep(200);
        }
    }
    return false;
}

void ClearClipboardSafe()
{
    try { Clipboard.Clear(); } catch { }
}

// ----------- Parse Action -----------
string actionName = Action;
string actionParam = "";
int idx = Action.IndexOf(':');
if (idx >= 0)
{
    actionName = Action[..idx];
    actionParam = Action[(idx + 1)..];
}

Console.WriteLine($"✓ VS Code Module Loaded");
Console.WriteLine($" - Process ID : {AppContext.Window.Current.ProcessId}");
Console.WriteLine($" - Action    : {actionName}");
if (!string.IsNullOrEmpty(actionParam)) Console.WriteLine($" - Param    : {actionParam}");

if (!FocusWindowHard())
{
    Console.WriteLine("✗ Unable to focus VS Code.");
    return;
}

// ----------- Main Logic -----------
try
{
    switch (actionName.ToLower())
    {
        // ---------- Basic ----------
        case "save": SendKeysWithRetry("^s"); break;
        case "new_file": SendKeysWithRetry("^n"); break;
        case "toggle_sidebar": SendKeysWithRetry("^b"); break;
        case "toggle_panel": SendKeysWithRetry("^j"); break;
        case "toggle_terminal":
        case "open_terminal": SendKeysWithRetry("^`"); break;

        // ---------- Duplicate File ----------
        case "duplicate_file":
            ClearClipboardSafe();
            SendKeysWithRetry("^+p", 300);
            SendKeysWithRetry("File: Copy Path of Active File{ENTER}", 600);
            Thread.Sleep(500);
            string path = Clipboard.GetText();
            if (string.IsNullOrWhiteSpace(path)) { Console.WriteLine("✗ Clipboard empty."); break; }

            string dir = Path.GetDirectoryName(path);
            string name = Path.GetFileNameWithoutExtension(path);
            string ext = Path.GetExtension(path);
            string copyPath = Path.Combine(dir, $"{name}_copy{ext}");
            int counter = 1;
            while (File.Exists(copyPath))
                copyPath = Path.Combine(dir, $"{name}_copy{counter++}{ext}");
            File.Copy(path, copyPath);
            Console.WriteLine($"✓ File duplicated → {copyPath}");
            break;

        // ---------- Git Bash ----------
        case "open_git_bash":
            Console.WriteLine("🧰 Opening Git Bash ...");
            SendKeysWithRetry("^`", 400);
            SendKeysWithRetry("bash{ENTER}", 400);
            break;

        // ---------- Python HTTP Server ----------
        case "python_http_server":
            Console.WriteLine("🌐 Running Python HTTP Server...");
            SendKeysWithRetry("^`", 400);
            SendKeysWithRetry("cls{ENTER}", 200);
            SendKeysWithRetry("python -m http.server{ENTER}", 800);
            break;

        // ---------- Virtual Env ----------
        case "python_venv:create":
            Console.WriteLine("🐍 Creating virtual environment...");
            SendKeysWithRetry("^`", 400);
            SendKeysWithRetry("python -m venv .venv{ENTER}", 800);
            break;

        case "python_venv:activate":
            Console.WriteLine("⚡ Activating virtual environment...");
            SendKeysWithRetry("^`", 400);
            string act = Environment.OSVersion.Platform == PlatformID.Win32NT
                ? ".venv\\Scripts\\activate{ENTER}"
                : "source .venv/bin/activate{ENTER}";
            SendKeysWithRetry(act, 700);
            break;

        case "python_venv:install requirements.txt":
        case "install_requirements":
            Console.WriteLine("📦 Installing requirements.txt...");
            SendKeysWithRetry("^`", 400);
            SendKeysWithRetry("pip install -r requirements.txt{ENTER}", 800);
            break;

        // ---------- Create File ----------
        default:
            if (actionName.StartsWith("create_file"))
            {
                string extn = ".txt";
                int c = actionName.IndexOf(':');
                if (c >= 0) extn = actionName[(c + 1)..];
                Console.Write($"🆕 Enter filename (without extension {extn}): ");
                string nameIn = Console.ReadLine();
                if (string.IsNullOrWhiteSpace(nameIn)) nameIn = "newfile";
                string filename = $"{nameIn}{extn}";
                SendKeysWithRetry("^n", 300);
                Thread.Sleep(400);
                SendKeysWithRetry("^+s", 400);
                Thread.Sleep(400);
                SendKeysWithRetry(filename + "{ENTER}", 300);
                Console.WriteLine($"✓ Created {filename}");
            }
            else
            {
                Console.WriteLine($"⚠ Unknown action ‘{actionName}’");
            }
            break;
    }

    Console.WriteLine("✅ VS Code automation completed.");
}
catch (Exception ex)
{
    Console.WriteLine($"[vscode.csx] ERROR: {ex.Message}");
}
