// vscode.csx
// Comprehensive Visual Studio Code automation module
// Author: ChatGPT (GPT-5)
// Compatible with ModuleManager.cs

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

// ----------- Utility helpers -----------

void FocusWindowHard()
{
    try
    {
        var hwnd = (IntPtr)AppContext.Window.Current.NativeWindowHandle;
        if (hwnd != IntPtr.Zero)
        {
            Win32.ShowWindow(hwnd, 9);   // SW_RESTORE
            Win32.SetForegroundWindow(hwnd);
            Thread.Sleep(350);
        }
    }
    catch (Exception ex)
    {
        Console.WriteLine($"[vscode.csx] FocusWindowHard error: {ex.Message}");
    }
}

void SendKeysWithDelay(string keys, int delay = 150)
{
    SendKeys.SendWait(keys);
    Thread.Sleep(delay);
}

static class Win32
{
    [System.Runtime.InteropServices.DllImport("user32.dll")]
    public static extern bool SetForegroundWindow(IntPtr hWnd);

    [System.Runtime.InteropServices.DllImport("user32.dll")]
    public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
}

// ----------- Action parsing -----------

string actionName = Action;
string actionParam = "";
int colon = Action.IndexOf(':');
if (colon >= 0)
{
    actionName = Action.Substring(0, colon).Trim();
    actionParam = Action.Substring(colon + 1).Trim();
}

Console.WriteLine($"? VS Code Automation Module Loaded!");
Console.WriteLine($" - App Name : vscode");
Console.WriteLine($" - Process  : {AppContext.Window.Current.ProcessId}");
Console.WriteLine($" - Action   : {actionName}");
if (!string.IsNullOrEmpty(actionParam))
    Console.WriteLine($" - Parameter: {actionParam}");

FocusWindowHard();

// ----------- Main automation logic -----------

try
{
    switch (actionName.ToLower())
    {
        case "command_palette":
        case "palette":
        case "open_palette":
            Console.WriteLine(" - Opening Command Palette...");
            SendKeysWithDelay("^+p", 400);

            if (!string.IsNullOrEmpty(actionParam))
            {
                Console.WriteLine($" - Typing command '{actionParam}'...");
                SendKeysWithDelay(actionParam, 300);
                SendKeysWithDelay("{ENTER}", 200);
            }
            break;

        case "menu":
        case "run_menu":
            // Same as command palette but forces the user’s menu item explicitly
            Console.WriteLine($" - Searching menu option '{actionParam}'...");
            SendKeysWithDelay("^+p", 400);
            SendKeysWithDelay(actionParam, 300);
            SendKeysWithDelay("{ENTER}", 200);
            break;

        case "toggle_sidebar":
            Console.WriteLine(" - Toggling sidebar...");
            SendKeysWithDelay("^b", 200);
            break;

        case "toggle_panel":
            Console.WriteLine(" - Toggling panel...");
            SendKeysWithDelay("^j", 200);
            break;

        case "toggle_terminal":
        case "open_terminal":
            Console.WriteLine(" - Opening integrated terminal...");
            SendKeysWithDelay("^`", 200);
            break;

        case "new_file":
            Console.WriteLine(" - Creating new file...");
            SendKeysWithDelay("^n", 200);
            break;

        case "save":
            Console.WriteLine(" - Saving current file...");
            SendKeysWithDelay("^s", 200);
            break;

        case "find":
            Console.WriteLine(" - Opening Find dialog...");
            SendKeysWithDelay("^f", 200);
            break;

        case "replace":
            Console.WriteLine(" - Opening Replace dialog...");
            SendKeysWithDelay("^h", 200);
            break;

        case "go_to_file":
            Console.WriteLine(" - Opening Quick Open (Go to File)...");
            SendKeysWithDelay("^p", 200);
            break;

        case "run_command":
            // Generic "run anything" in palette form
            Console.WriteLine($" - Running '{actionParam}' via Command Palette...");
            SendKeysWithDelay("^+p", 400);
            if (!string.IsNullOrEmpty(actionParam))
            {
                SendKeysWithDelay(actionParam, 300);
                SendKeysWithDelay("{ENTER}", 200);
            }
            break;

        case "list_commands":
            // Display the complete consolidated developer command list (trimmed)
            Console.WriteLine(" - Developer commands snapshot:");
            Console.WriteLine("   New File, Open Folder..., Save, Save All, Toggle Sidebar, Toggle Panel,");
            Console.WriteLine("   Run Task..., Debug: Start, Run Without Debugging, New Terminal,");
            Console.WriteLine("   Extensions: Install, Git: Commit, Git: Push, Git: Pull, AI Chat: Open,");
            Console.WriteLine("   Run All Tests, Go to Definition, Go to Symbol in Workspace..., etc.");
            break;
        case "python_venv":
        case "venv":
            {
                if (string.IsNullOrEmpty(actionParam))
                {
                    Console.WriteLine("Usage: python_venv:create | python_venv:activate | python_venv:install <package>");
                    break;
                }

                FocusWindowHard();
                SendKeysWithDelay("^`", 400); // Open terminal
                Thread.Sleep(500);

                if (actionParam.StartsWith("create", StringComparison.OrdinalIgnoreCase))
                {
                    Console.WriteLine(" - Creating virtual environment (.venv)...");
                    SendKeysWithDelay("python -m venv .venv{ENTER}", 1000);
                }
                else if (actionParam.StartsWith("activate", StringComparison.OrdinalIgnoreCase))
                {
                    Console.WriteLine(" - Activating virtual environment...");
                    SendKeysWithDelay(".venv\\Scripts\\activate{ENTER}", 800);
                }
                else if (actionParam.StartsWith("install", StringComparison.OrdinalIgnoreCase))
                {
                    string[] parts = actionParam.Split(' ');
                    if (parts.Length < 2)
                    {
                        Console.WriteLine("Please specify a package to install.");
                        break;
                    }

                    string pkg = parts[1];
                    Console.WriteLine($" - Installing package: {pkg}");
                    SendKeysWithDelay($"pip install {pkg}{{ENTER}}", 1000);
                }
                else
                {
                    Console.WriteLine($"Unknown virtual environment command: {actionParam}");
                }
            }
            break;

        // ----------- Duplicate Current File -----------
        case "duplicate_file":
            try
            {
                FocusWindowHard();
                Console.WriteLine(" - Attempting to duplicate current file...");

                // Open Command Palette and get full path of current file
                SendKeysWithDelay("^+p", 500);
                SendKeysWithDelay("Copy Path of Active File{ENTER}", 700);
                Thread.Sleep(600);

                // Clipboard read attempt
                string? filePath = Clipboard.GetText();
                if (string.IsNullOrEmpty(filePath) || !File.Exists(filePath))
                {
                    Console.WriteLine(" - Could not retrieve current file path from clipboard.");
                    break;
                }

                string dir = Path.GetDirectoryName(filePath)!;
                string name = Path.GetFileNameWithoutExtension(filePath);
                string ext = Path.GetExtension(filePath);

                string dupFilePath = Path.Combine(dir, $"{name}(dup){ext}");
                File.Copy(filePath, dupFilePath, overwrite: false);

                Console.WriteLine($" - File duplicated: {dupFilePath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[vscode.csx] Error duplicating file: {ex.Message}");
            }
            break;

        default:
            Console.WriteLine($"[vscode.csx] Unknown action: {actionName}");
            break;
    }

    Console.WriteLine(" - Automation task complete.");
}
catch (Exception ex)
{
    Console.WriteLine($"[vscode.csx] Error executing action: {ex.Message}");
}
