// vscode.csx
// Comprehensive Visual Studio Code automation module
// Compatible with ModuleManager.cs
// Version 5 — Extended Developer Actions

using System;
using System.IO;
using System.Threading;
using System.Windows.Forms;
using System.Windows.Automation;

// -------------------------------
// Globals provided by ModuleManager
// -------------------------------
//   AppContext  -> ApplicationContext
//   Action      -> string
// -------------------------------

if (AppContext == null) throw new Exception("AppContext is null.");
if (string.IsNullOrWhiteSpace(Action)) throw new Exception("No Action provided.");

// ----------- Utility Helpers -----------

bool FocusVSCode()
{
    try
    {
        var hwnd = (IntPtr)AppContext.Window.Current.NativeWindowHandle;
        if (hwnd == IntPtr.Zero) return false;

        Win32.ShowWindow(hwnd, 9);   // SW_RESTORE
        Thread.Sleep(150);
        return Win32.SetForegroundWindow(hwnd);
    }
    catch { return false; }
}

bool SendKeysWithDelay(string keys, int delay = 200)
{
    try
    {
        SendKeys.SendWait(keys);
        Thread.Sleep(delay);
        return true;
    }
    catch
    {
        return false;
    }
}

static class Win32
{
    [System.Runtime.InteropServices.DllImport("user32.dll")]
    public static extern bool SetForegroundWindow(IntPtr hWnd);

    [System.Runtime.InteropServices.DllImport("user32.dll")]
    public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
}

// ----------- Start Execution -----------

string actionName = Action.Trim().ToLower();
Console.WriteLine($"[VSCode.csx] → Action received: {actionName}");

if (!FocusVSCode())
{
    Console.WriteLine("[VSCode.csx] ✗ Failed to focus VS Code window.");
    return;
}

try
{
    switch (actionName)
    {
        // -------------------- BASIC ACTIONS --------------------
        case "save":
            Console.WriteLine("💾 Saving file...");
            SendKeysWithDelay("^s");
            break;

        case "new_file":
            Console.WriteLine("📝 Creating new file...");
            SendKeysWithDelay("^n");
            break;

        case "toggle_terminal":
            Console.WriteLine("🖥️ Toggling terminal...");
            SendKeysWithDelay("^`");
            break;

        case "toggle_sidebar":
            Console.WriteLine("📁 Toggling sidebar...");
            SendKeysWithDelay("^b");
            break;

        case "toggle_panel":
            Console.WriteLine("🧩 Toggling bottom panel...");
            SendKeysWithDelay("^j");
            break;

        // -------------------- DUPLICATE FILE --------------------
        case "duplicate_file":
            Console.WriteLine("📄 Duplicating current file...");
            SendKeysWithDelay("^+p", 300);
            SendKeysWithDelay("File: Copy Path of Active File{ENTER}", 800);
            Thread.Sleep(600);

            string path = Clipboard.GetText();
            if (string.IsNullOrEmpty(path))
            {
                Console.WriteLine("✗ No file path in clipboard.");
                break;
            }

            string dir = Path.GetDirectoryName(path);
            string name = Path.GetFileNameWithoutExtension(path);
            string ext = Path.GetExtension(path);

            string newPath = Path.Combine(dir, $"{name}_copy{ext}");
            int counter = 1;
            while (File.Exists(newPath))
            {
                newPath = Path.Combine(dir, $"{name}_copy{counter++}{ext}");
            }

            File.Copy(path, newPath);
            Console.WriteLine($"✓ File duplicated → {newPath}");
            break;

        // -------------------- PYTHON VIRTUAL ENVIRONMENT --------------------
        case "python_venv:create":
            Console.WriteLine("🐍 Creating virtual environment (.venv)...");
            SendKeysWithDelay("^`", 300);
            Thread.Sleep(400);
            SendKeysWithDelay("cls{ENTER}", 200);
            SendKeysWithDelay("python -m venv .venv{ENTER}", 800);
            break;

        case "python_venv:activate":
            Console.WriteLine("⚡ Activating virtual environment...");
            SendKeysWithDelay("^`", 300);
            Thread.Sleep(400);
            SendKeysWithDelay("cls{ENTER}", 200);
            string activateCmd = Environment.OSVersion.Platform == PlatformID.Win32NT
                ? ".venv\\Scripts\\activate{ENTER}"
                : "source .venv/bin/activate{ENTER}";
            SendKeysWithDelay(activateCmd, 700);
            break;

        case "python_venv:install requirements.txt":
        case "install_requirements":
            Console.WriteLine("📦 Installing packages from requirements.txt...");
            SendKeysWithDelay("^`", 300);
            Thread.Sleep(400);
            SendKeysWithDelay("pip install -r requirements.txt{ENTER}", 800);
            break;

        // -------------------- OPEN GIT BASH --------------------
        case "open_git_bash":
            Console.WriteLine("🧰 Opening Git Bash terminal...");
            SendKeysWithDelay("^`", 400);
            Thread.Sleep(400);
            SendKeysWithDelay("bash{ENTER}", 300);
            break;

        // -------------------- PYTHON HTTP SERVER --------------------
        case "python_http_server":
            Console.WriteLine("🌐 Running Python HTTP Server (port 8000)...");
            SendKeysWithDelay("^`", 400);
            Thread.Sleep(400);
            SendKeysWithDelay("cls{ENTER}", 200);
            SendKeysWithDelay("python -m http.server{ENTER}", 800);
            break;

        // -------------------- CREATE FILE BY EXTENSION --------------------
        default:
            if (actionName.StartsWith("create_file"))
            {
                string extension = ".txt";
                int colonIndex = actionName.IndexOf(':');
                if (colonIndex > 0)
                    extension = actionName.Substring(colonIndex + 1).Trim();

                Console.WriteLine($"🆕 Creating new {extension} file...");
                SendKeysWithDelay("^n", 300);
                Thread.Sleep(400);

                // Ask for filename
                Console.Write("Enter file name (without extension): ");
                string? nameInput = Console.ReadLine();
                if (string.IsNullOrWhiteSpace(nameInput))
                    nameInput = "newfile";

                string filename = $"{nameInput}{extension}";
                SendKeysWithDelay("^+s", 400);
                Thread.Sleep(600);
                SendKeysWithDelay(filename + "{ENTER}", 300);
                Console.WriteLine($"✓ Created file: {filename}");
            }
            else
            {
                Console.WriteLine($"⚠️ Unknown action: {actionName}");
            }
            break;
    }

    Console.WriteLine("✅ VS Code automation completed.");
}
catch (Exception ex)
{
    Console.WriteLine($"[VSCode.csx] ERROR: {ex.Message}");
}
