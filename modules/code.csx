// Advanced VS Code Python Development Automation Module
// Comprehensive routines for Python environment, package management, and development workflows
// Enhanced with intelligent file creation and robust error handling
// Compatible with ModuleManager.cs
// Version 1.0

using System;
using System.IO;
using System.Linq;
using System.Threading;
using System.Windows.Forms;
using System.Windows.Automation;
using System.Diagnostics;
using System.Text.RegularExpressions;

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
const int TERMINAL_COMMAND_DELAY = 1200; // Longer for terminal commands
const int CLIPBOARD_TIMEOUT = 3000;

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
        Console.WriteLine($"[code.csx] Window responsiveness check failed: {ex.Message}");
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
                Console.WriteLine($"[code.csx] Window not responsive on attempt {attempt}");
                Thread.Sleep(500 * attempt);
                continue;
            }

            var hwnd = (IntPtr)AppContext.Window.Current.NativeWindowHandle;
            if (hwnd == IntPtr.Zero)
            {
                Console.WriteLine($"[code.csx] Invalid window handle on attempt {attempt}");
                Thread.Sleep(300 * attempt);
                continue;
            }

            Win32.ShowWindow(hwnd, 9);   // SW_RESTORE
            Thread.Sleep(200);
            
            if (Win32.SetForegroundWindow(hwnd))
            {
                Thread.Sleep(350);
                Console.WriteLine($"[code.csx] Window focused successfully");
                return true;
            }
            
            Console.WriteLine($"[code.csx] SetForegroundWindow failed on attempt {attempt}");
            Thread.Sleep(300 * attempt);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"[code.csx] FocusWindowHard error (attempt {attempt}): {ex.Message}");
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
                Console.WriteLine($"[code.csx] Window not responsive before SendKeys (attempt {attempt})");
                Thread.Sleep(300);
                continue;
            }

            SendKeys.SendWait(keys);
            Thread.Sleep(delay);
            return true;
        }
        catch (Exception ex)
        {
            Console.WriteLine($"[code.csx] SendKeys error (attempt {attempt}): {ex.Message}");
            if (attempt == retries) return false;
            Thread.Sleep(200 * attempt);
        }
    }
    return false;
}

bool WaitForClipboard(int maxWaitMs = CLIPBOARD_TIMEOUT)
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
            Console.WriteLine($"[code.csx] Clipboard read error (attempt {attempt}): {ex.Message}");
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
        Console.WriteLine($"[code.csx] Clipboard clear warning: {ex.Message}");
    }
}

bool OpenTerminal()
{
    Console.WriteLine(" - Opening integrated terminal...");
    if (!SendKeysWithRetry("^`", 500))
    {
        Console.WriteLine(" - ERROR: Failed to open terminal");
        return false;
    }
    Thread.Sleep(800); // Wait for terminal to be ready
    return true;
}

bool ExecuteTerminalCommand(string command, int waitMs = TERMINAL_COMMAND_DELAY)
{
    Console.WriteLine($" - Executing: {command}");
    if (!SendKeysWithRetry($"{command}{{ENTER}}", waitMs))
    {
        Console.WriteLine($" - ERROR: Failed to execute command: {command}");
        return false;
    }
    return true;
}

bool OpenCommandPalette()
{
    if (!SendKeysWithRetry("^+p", 600))
    {
        Console.WriteLine(" - ERROR: Failed to open Command Palette");
        return false;
    }
    Thread.Sleep(400);
    return true;
}

bool ExecuteCommandPaletteCommand(string command, int waitMs = 800)
{
    if (!OpenCommandPalette()) return false;
    
    if (!SendKeysWithRetry(command, 300))
    {
        Console.WriteLine($" - ERROR: Failed to type command: {command}");
        return false;
    }
    
    if (!SendKeysWithRetry("{ENTER}", waitMs))
    {
        Console.WriteLine($" - ERROR: Failed to execute command");
        return false;
    }
    
    return true;
}

string GetCurrentWorkspacePath()
{
    try
    {
        ClearClipboard();
        
        if (!ExecuteCommandPaletteCommand("File: Copy Path of Active File", 1000))
        {
            Console.WriteLine(" - WARNING: Could not get active file path");
            return null;
        }

        if (!WaitForClipboard(2000))
        {
            Console.WriteLine(" - WARNING: Clipboard did not receive path");
            return null;
        }

        string filePath = SafeGetClipboardText()?.Trim();
        ClearClipboard();

        if (!string.IsNullOrEmpty(filePath) && File.Exists(filePath))
        {
            return Path.GetDirectoryName(filePath);
        }
    }
    catch (Exception ex)
    {
        Console.WriteLine($"[code.csx] Error getting workspace path: {ex.Message}");
    }
    
    return null;
}

string PromptForInput(string message)
{
    try
    {
        // Use InputBox for user input
        string input = Microsoft.VisualBasic.Interaction.InputBox(message, "VS Code Automation", "", -1, -1);
        return input?.Trim();
    }
    catch (Exception ex)
    {
        Console.WriteLine($"[code.csx] Input prompt error: {ex.Message}");
        return null;
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

Console.WriteLine($"✓ VS Code Python Development Module Loaded!");
Console.WriteLine($" - App Name : code / vscode");
Console.WriteLine($" - Process  : {AppContext.Window.Current.ProcessId}");
Console.WriteLine($" - Action   : {actionName}");
if (!string.IsNullOrEmpty(actionParam))
    Console.WriteLine($" - Parameter: {actionParam}");

if (!FocusWindowHard())
{
    Console.WriteLine("[code.csx] ERROR: Failed to focus VS Code window. Aborting.");
    return;
}

// ----------- Main automation logic -----------

try
{
    switch (actionName.ToLower())
    {
        // ==================== ENVIRONMENT MANAGEMENT ====================
        
        case "create_venv":
        case "venv_create":
            try
            {
                Console.WriteLine(" - Creating Python virtual environment (.venv)...");
                
                if (!OpenTerminal()) break;
                
                ExecuteTerminalCommand("cls", 300);
                
                if (ExecuteTerminalCommand("python -m venv .venv", 3000))
                {
                    Console.WriteLine(" - ✓ Virtual environment creation initiated");
                    Console.WriteLine(" - NOTE: This may take 10-30 seconds to complete");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[code.csx] Create venv error: {ex.Message}");
            }
            break;

        case "activate_venv":
        case "venv_activate":
            try
            {
                Console.WriteLine(" - Activating virtual environment...");
                
                if (!OpenTerminal()) break;
                
                ExecuteTerminalCommand("cls", 300);
                
                string activateCmd = Environment.OSVersion.Platform == PlatformID.Win32NT
                    ? ".venv\\Scripts\\activate"
                    : "source .venv/bin/activate";
                
                if (ExecuteTerminalCommand(activateCmd, 1000))
                {
                    Console.WriteLine(" - ✓ Virtual environment activated");
                    Console.WriteLine(" - TIP: Look for (.venv) in terminal prompt");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[code.csx] Activate venv error: {ex.Message}");
            }
            break;

        case "deactivate_venv":
        case "venv_deactivate":
            try
            {
                Console.WriteLine(" - Deactivating virtual environment...");
                
                if (!OpenTerminal()) break;
                
                if (ExecuteTerminalCommand("deactivate", 500))
                {
                    Console.WriteLine(" - ✓ Virtual environment deactivated");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[code.csx] Deactivate venv error: {ex.Message}");
            }
            break;

        case "select_interpreter":
        case "switch_interpreter":
            try
            {
                Console.WriteLine(" - Opening Python interpreter selector...");
                
                if (ExecuteCommandPaletteCommand("Python: Select Interpreter", 800))
                {
                    Console.WriteLine(" - ✓ Interpreter selector opened");
                    Console.WriteLine(" - TIP: Choose your .venv or desired Python version");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[code.csx] Select interpreter error: {ex.Message}");
            }
            break;

        // ==================== PACKAGE MANAGEMENT ====================

        case "install_package":
        case "pip_install":
            try
            {
                string package = actionParam;
                if (string.IsNullOrEmpty(package))
                {
                    Console.WriteLine(" - ERROR: Package name required");
                    Console.WriteLine("   Usage: install_package:<package_name>");
                    break;
                }

                Console.WriteLine($" - Installing package: {package}");
                
                if (!OpenTerminal()) break;
                
                ExecuteTerminalCommand("cls", 300);
                
                if (ExecuteTerminalCommand($"pip install {package}", 2000))
                {
                    Console.WriteLine($" - ✓ Package installation initiated: {package}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[code.csx] Install package error: {ex.Message}");
            }
            break;

        case "install_requirements":
        case "pip_install_requirements":
            try
            {
                Console.WriteLine(" - Installing packages from requirements.txt...");
                
                if (!OpenTerminal()) break;
                
                ExecuteTerminalCommand("cls", 300);
                
                string installCmd = actionParam?.ToLower() == "uv" 
                    ? "uv pip install -r requirements.txt"
                    : "pip install -r requirements.txt";
                
                if (ExecuteTerminalCommand(installCmd, 2000))
                {
                    Console.WriteLine(" - ✓ Requirements installation initiated");
                    Console.WriteLine(" - NOTE: This may take a while depending on package size");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[code.csx] Install requirements error: {ex.Message}");
            }
            break;

        case "upgrade_package":
        case "pip_upgrade":
            try
            {
                string package = actionParam;
                if (string.IsNullOrEmpty(package))
                {
                    Console.WriteLine(" - ERROR: Package name required");
                    Console.WriteLine("   Usage: upgrade_package:<package_name>");
                    break;
                }

                Console.WriteLine($" - Upgrading package: {package}");
                
                if (!OpenTerminal()) break;
                
                if (ExecuteTerminalCommand($"pip install --upgrade {package}", 2000))
                {
                    Console.WriteLine($" - ✓ Package upgrade initiated: {package}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[code.csx] Upgrade package error: {ex.Message}");
            }
            break;

        case "uninstall_package":
        case "pip_uninstall":
            try
            {
                string package = actionParam;
                if (string.IsNullOrEmpty(package))
                {
                    Console.WriteLine(" - ERROR: Package name required");
                    Console.WriteLine("   Usage: uninstall_package:<package_name>");
                    break;
                }

                Console.WriteLine($" - Uninstalling package: {package}");
                
                if (!OpenTerminal()) break;
                
                if (ExecuteTerminalCommand($"pip uninstall -y {package}", 1500))
                {
                    Console.WriteLine($" - ✓ Package uninstalled: {package}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[code.csx] Uninstall package error: {ex.Message}");
            }
            break;

        case "list_packages":
        case "pip_list":
            try
            {
                Console.WriteLine(" - Listing installed packages...");
                
                if (!OpenTerminal()) break;
                
                ExecuteTerminalCommand("cls", 300);
                
                if (ExecuteTerminalCommand("pip list", 1000))
                {
                    Console.WriteLine(" - ✓ Package list displayed in terminal");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[code.csx] List packages error: {ex.Message}");
            }
            break;

        case "outdated_packages":
        case "pip_outdated":
            try
            {
                Console.WriteLine(" - Checking for outdated packages...");
                
                if (!OpenTerminal()) break;
                
                ExecuteTerminalCommand("cls", 300);
                
                if (ExecuteTerminalCommand("pip list --outdated", 2000))
                {
                    Console.WriteLine(" - ✓ Outdated packages check initiated");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[code.csx] Outdated packages error: {ex.Message}");
            }
            break;

        case "freeze_requirements":
        case "pip_freeze":
            try
            {
                Console.WriteLine(" - Freezing dependencies to requirements.txt...");
                
                if (!OpenTerminal()) break;
                
                ExecuteTerminalCommand("cls", 300);
                
                if (ExecuteTerminalCommand("pip freeze > requirements.txt", 1000))
                {
                    Console.WriteLine(" - ✓ Requirements frozen to requirements.txt");
                    Console.WriteLine(" - TIP: Review the file to remove unwanted system packages");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[code.csx] Freeze requirements error: {ex.Message}");
            }
            break;

        // ==================== RUNNING CODE & SCRIPTS ====================

        case "run_file":
        case "python_run":
            try
            {
                string fileName = actionParam;
                Console.WriteLine($" - Running Python file{(string.IsNullOrEmpty(fileName) ? " (active file)" : $": {fileName}")}...");
                
                if (!OpenTerminal()) break;
                
                ExecuteTerminalCommand("cls", 300);
                
                string runCmd = string.IsNullOrEmpty(fileName) 
                    ? "python ${file}" // VS Code variable for active file
                    : $"python {fileName}";
                
                if (ExecuteTerminalCommand(runCmd, 1500))
                {
                    Console.WriteLine(" - ✓ Script execution initiated");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[code.csx] Run file error: {ex.Message}");
            }
            break;

        case "run_flask":
            try
            {
                Console.WriteLine(" - Starting Flask development server...");
                
                if (!OpenTerminal()) break;
                
                ExecuteTerminalCommand("cls", 300);
                
                // Set FLASK_APP if provided
                if (!string.IsNullOrEmpty(actionParam))
                {
                    string setCmd = Environment.OSVersion.Platform == PlatformID.Win32NT
                        ? $"set FLASK_APP={actionParam}"
                        : $"export FLASK_APP={actionParam}";
                    ExecuteTerminalCommand(setCmd, 500);
                }
                
                if (ExecuteTerminalCommand("flask run", 1500))
                {
                    Console.WriteLine(" - ✓ Flask server started");
                    Console.WriteLine(" - TIP: Usually runs on http://127.0.0.1:5000");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[code.csx] Run Flask error: {ex.Message}");
            }
            break;

        case "run_django":
            try
            {
                Console.WriteLine(" - Starting Django development server...");
                
                if (!OpenTerminal()) break;
                
                ExecuteTerminalCommand("cls", 300);
                
                if (ExecuteTerminalCommand("python manage.py runserver", 1500))
                {
                    Console.WriteLine(" - ✓ Django server started");
                    Console.WriteLine(" - TIP: Usually runs on http://127.0.0.1:8000");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[code.csx] Run Django error: {ex.Message}");
            }
            break;

        case "run_fastapi":
        case "run_uvicorn":
            try
            {
                string module = string.IsNullOrEmpty(actionParam) ? "main:app" : actionParam;
                Console.WriteLine($" - Starting FastAPI server: {module}");
                
                if (!OpenTerminal()) break;
                
                ExecuteTerminalCommand("cls", 300);
                
                if (ExecuteTerminalCommand($"uvicorn {module} --reload", 1500))
                {
                    Console.WriteLine(" - ✓ FastAPI server started with auto-reload");
                    Console.WriteLine(" - TIP: Usually runs on http://127.0.0.1:8000");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[code.csx] Run FastAPI error: {ex.Message}");
            }
            break;

        case "run_streamlit":
            try
            {
                string appFile = string.IsNullOrEmpty(actionParam) ? "app.py" : actionParam;
                Console.WriteLine($" - Starting Streamlit app: {appFile}");
                
                if (!OpenTerminal()) break;
                
                ExecuteTerminalCommand("cls", 300);
                
                if (ExecuteTerminalCommand($"streamlit run {appFile}", 1500))
                {
                    Console.WriteLine(" - ✓ Streamlit app started");
                    Console.WriteLine(" - NOTE: Will open in browser automatically");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[code.csx] Run Streamlit error: {ex.Message}");
            }
            break;

        case "http_server":
        case "serve_folder":
            try
            {
                string port = string.IsNullOrEmpty(actionParam) ? "8000" : actionParam;
                Console.WriteLine($" - Starting HTTP server on port {port}...");
                
                if (!OpenTerminal()) break;
                
                ExecuteTerminalCommand("cls", 300);
                
                if (ExecuteTerminalCommand($"python -m http.server {port}", 1000))
                {
                    Console.WriteLine($" - ✓ HTTP server started on port {port}");
                    Console.WriteLine($" - TIP: Access at http://localhost:{port}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[code.csx] HTTP server error: {ex.Message}");
            }
            break;

        case "run_jupyter":
        case "jupyter_notebook":
            try
            {
                Console.WriteLine(" - Starting Jupyter Notebook...");
                
                if (!OpenTerminal()) break;
                
                ExecuteTerminalCommand("cls", 300);
                
                if (ExecuteTerminalCommand("jupyter notebook", 1500))
                {
                    Console.WriteLine(" - ✓ Jupyter Notebook started");
                    Console.WriteLine(" - NOTE: Will open in browser automatically");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[code.csx] Jupyter error: {ex.Message}");
            }
            break;

        // ==================== DEBUGGING & TESTING ====================

        case "run_debug":
        case "start_debug":
            try
            {
                Console.WriteLine(" - Starting debugger...");
                
                if (SendKeysWithRetry("{F5}", 500))
                {
                    Console.WriteLine(" - ✓ Debugger started");
                    Console.WriteLine(" - TIP: Set breakpoints by clicking left of line numbers");
                }
                else
                {
                    Console.WriteLine(" - ERROR: Failed to start debugger");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[code.csx] Debug error: {ex.Message}");
            }
            break;

        case "run_pytest":
        case "pytest":
            try
            {
                string testPath = actionParam ?? "";
                Console.WriteLine($" - Running pytest{(string.IsNullOrEmpty(testPath) ? "" : $": {testPath}")}...");
                
                if (!OpenTerminal()) break;
                
                ExecuteTerminalCommand("cls", 300);
                
                string pytestCmd = string.IsNullOrEmpty(testPath) ? "pytest" : $"pytest {testPath}";
                
                if (ExecuteTerminalCommand(pytestCmd, 2000))
                {
                    Console.WriteLine(" - ✓ Pytest execution initiated");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[code.csx] Pytest error: {ex.Message}");
            }
            break;

        case "run_unittest":
        case "unittest":
            try
            {
                Console.WriteLine(" - Running unittest...");
                
                if (!OpenTerminal()) break;
                
                ExecuteTerminalCommand("cls", 300);
                
                if (ExecuteTerminalCommand("python -m unittest discover", 2000))
                {
                    Console.WriteLine(" - ✓ Unittest execution initiated");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[code.csx] Unittest error: {ex.Message}");
            }
            break;

        // ==================== GIT & VERSION CONTROL ====================

        case "git_init":
            try
            {
                Console.WriteLine(" - Initializing Git repository...");
                
                if (!OpenTerminal()) break;
                
                ExecuteTerminalCommand("cls", 300);
                
                if (ExecuteTerminalCommand("git init", 1000))
                {
                    Console.WriteLine(" - ✓ Git repository initialized");
                    Console.WriteLine(" - TIP: Don't forget to create .gitignore file");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[code.csx] Git init error: {ex.Message}");
            }
            break;

        case "git_status":
            try
            {
                Console.WriteLine(" - Checking Git status...");
                
                if (!OpenTerminal()) break;
                
                ExecuteTerminalCommand("cls", 300);
                
                if (ExecuteTerminalCommand("git status", 800))
                {
                    Console.WriteLine(" - ✓ Git status displayed");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[code.csx] Git status error: {ex.Message}");
            }
            break;

        case "git_add_all":
        case "git_add":
            try
            {
                Console.WriteLine(" - Staging all changes for commit...");
                
                if (!OpenTerminal()) break;
                
                if (ExecuteTerminalCommand("git add .", 800))
                {
                    Console.WriteLine(" - ✓ All changes staged");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[code.csx] Git add error: {ex.Message}");
            }
            break;

        case "git_commit":
            try
            {
                string message = actionParam;
                if (string.IsNullOrEmpty(message))
                {
                    Console.WriteLine(" - ERROR: Commit message required");
                    Console.WriteLine("   Usage: git_commit:<message>");
                    break;
                }

                Console.WriteLine($" - Committing changes: {message}");
                
                if (!OpenTerminal()) break;
                
                // Escape quotes in commit message
                message = message.Replace("\"", "\\\"");
                
                if (ExecuteTerminalCommand($"git commit -m \"{message}\"", 1000))
                {
                    Console.WriteLine(" - ✓ Changes committed");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[code.csx] Git commit error: {ex.Message}");
            }
            break;

        case "git_push":
            try
            {
                string branch = string.IsNullOrEmpty(actionParam) ? "main" : actionParam;
                Console.WriteLine($" - Pushing to remote: origin/{branch}");
                
                if (!OpenTerminal()) break;
                
                if (ExecuteTerminalCommand($"git push origin {branch}", 2000))
                {
                    Console.WriteLine($" - ✓ Push to origin/{branch} initiated");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[code.csx] Git push error: {ex.Message}");
            }
            break;

        case "git_pull":
            try
            {
                string branch = string.IsNullOrEmpty(actionParam) ? "main" : actionParam;
                Console.WriteLine($" - Pulling from remote: origin/{branch}");
                
                if (!OpenTerminal()) break;
                
                if (ExecuteTerminalCommand($"git pull origin {branch}", 2000))
                {
                    Console.WriteLine($" - ✓ Pull from origin/{branch} initiated");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[code.csx] Git pull error: {ex.Message}");
            }
            break;

        case "git_log":
            try
            {
                Console.WriteLine(" - Viewing Git log...");
                
                if (!OpenTerminal()) break;
                
                ExecuteTerminalCommand("cls", 300);
                
                if (ExecuteTerminalCommand("git log --oneline --graph -10", 800))
                {
                    Console.WriteLine(" - ✓ Git log displayed (last 10 commits)");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[code.csx] Git log error: {ex.Message}");
            }
            break;

        // ==================== FILE CREATION (SPECIAL FEATURE) ====================

        case "create_file":
        case "new_file":
            try
            {
                Console.WriteLine(" - Creating new file with intelligent naming...");
                
                string fileName = null;
                string fileExtension = null;

                // Parse parameter for file type hints
                string param = actionParam?.ToLower() ?? "";
                bool isPython = param.Contains("python") || param.Contains("py");
                bool isJavaScript = param.Contains("javascript") || param.Contains("js");
                bool isHtml = param.Contains("html");
                bool isCss = param.Contains("css");
                bool isJson = param.Contains("json");
                bool isMarkdown = param.Contains("markdown") || param.Contains("md");

                // Determine if extension is pre-specified
                if (isPython)
                {
                    fileExtension = ".py";
                    fileName = PromptForInput("Enter Python file name (without extension):");
                }
                else if (isJavaScript)
                {
                    fileExtension = ".js";
                    fileName = PromptForInput("Enter JavaScript file name (without extension):");
                }
                else if (isHtml)
                {
                    fileExtension = ".html";
                    fileName = PromptForInput("Enter HTML file name (without extension):");
                }
                else if (isCss)
                {
                    fileExtension = ".css";
                    fileName = PromptForInput("Enter CSS file name (without extension):");
                }
                else if (isJson)
                {
                    fileExtension = ".json";
                    fileName = PromptForInput("Enter JSON file name (without extension):");
                }
                else if (isMarkdown)
                {
                    fileExtension = ".md";
                    fileName = PromptForInput("Enter Markdown file name (without extension):");
                }
                else
                {
                    // No file type specified - ask for both name and extension
                    fileName = PromptForInput("Enter file name (e.g., script.py or document.txt):");
                    
                    if (!string.IsNullOrEmpty(fileName))
                    {
                        // Check if user already included extension
                        if (Path.HasExtension(fileName))
                        {
                            fileExtension = Path.GetExtension(fileName);
                            fileName = Path.GetFileNameWithoutExtension(fileName);
                        }
                        else
                        {
                            // Ask for extension separately
                            fileExtension = PromptForInput("Enter file extension (e.g., .py, .txt, .js):");
                            
                            // Ensure extension starts with dot
                            if (!string.IsNullOrEmpty(fileExtension) && !fileExtension.StartsWith("."))
                            {
                                fileExtension = "." + fileExtension;
                            }
                        }
                    }
                }

                // Validate inputs
                if (string.IsNullOrEmpty(fileName))
                {
                    Console.WriteLine(" - ERROR: File creation cancelled (no name provided)");
                    break;
                }

                if (string.IsNullOrEmpty(fileExtension))
                {
                    Console.WriteLine(" - WARNING: No extension provided, defaulting to .txt");
                    fileExtension = ".txt";
                }

                // Sanitize filename
                fileName = fileName.Trim();
                char[] invalidChars = Path.GetInvalidFileNameChars();
                foreach (char c in invalidChars)
                {
                    fileName = fileName.Replace(c.ToString(), "");
                }

                string fullFileName = fileName + fileExtension;
                Console.WriteLine($" - Creating file: {fullFileName}");

                // Get workspace path
                string workspacePath = GetCurrentWorkspacePath();
                if (string.IsNullOrEmpty(workspacePath))
                {
                    Console.WriteLine(" - WARNING: Could not determine workspace path");
                    Console.WriteLine(" - File will be created in current terminal directory");
                }

                // Create the file using VS Code command
                if (!OpenTerminal()) break;
                
                ExecuteTerminalCommand("cls", 300);

                // Create empty file (cross-platform)
                string createCmd = Environment.OSVersion.Platform == PlatformID.Win32NT
                    ? $"type nul > {fullFileName}"
                    : $"touch {fullFileName}";
                
                if (ExecuteTerminalCommand(createCmd, 800))
                {
                    Console.WriteLine($" - ✓ File created: {fullFileName}");
                    
                    // Open the file in editor
                    Thread.Sleep(500);
                    if (ExecuteCommandPaletteCommand($"File: Open File...", 600))
                    {
                        Thread.Sleep(400);
                        SendKeysWithRetry(fullFileName, 300);
                        SendKeysWithRetry("{ENTER}", 500);
                        Console.WriteLine($" - ✓ File opened in editor");
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[code.csx] Create file error: {ex.Message}");
            }
            break;

        // ==================== PROJECT SETUP & CONFIGURATION ====================

        case "create_gitignore":
        case "gitignore":
            try
            {
                Console.WriteLine(" - Creating .gitignore for Python project...");
                
                string workspacePath = GetCurrentWorkspacePath();
                if (string.IsNullOrEmpty(workspacePath))
                {
                    Console.WriteLine(" - ERROR: Could not determine workspace path");
                    break;
                }

                string gitignorePath = Path.Combine(workspacePath, ".gitignore");
                
                string gitignoreContent = @"# Python
__pycache__/
*.py[cod]
*$py.class
*.so
.Python
build/
develop-eggs/
dist/
downloads/
eggs/
.eggs/
lib/
lib64/
parts/
sdist/
var/
wheels/
*.egg-info/
.installed.cfg
*.egg

# Virtual Environment
.venv/
venv/
ENV/
env/

# IDE
.vscode/
.idea/
*.swp
*.swo
*~

# Testing
.pytest_cache/
.coverage
htmlcov/

# Environment variables
.env
.env.local

# OS
.DS_Store
Thumbs.db";

                File.WriteAllText(gitignorePath, gitignoreContent);
                Console.WriteLine($" - ✓ .gitignore created at: {gitignorePath}");
                Console.WriteLine(" - Includes: Python cache, venv, IDE files, OS files");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[code.csx] Create .gitignore error: {ex.Message}");
            }
            break;

        case "create_env":
        case "env_template":
            try
            {
                Console.WriteLine(" - Creating .env template file...");
                
                string workspacePath = GetCurrentWorkspacePath();
                if (string.IsNullOrEmpty(workspacePath))
                {
                    Console.WriteLine(" - ERROR: Could not determine workspace path");
                    break;
                }

                string envPath = Path.Combine(workspacePath, ".env");
                
                string envContent = @"# Environment Variables Template
# Copy this to .env and fill in your values

# Database
DATABASE_URL=
DB_HOST=localhost
DB_PORT=5432
DB_NAME=
DB_USER=
DB_PASSWORD=

# API Keys
API_KEY=
SECRET_KEY=

# Application
DEBUG=True
ENVIRONMENT=development
PORT=8000

# External Services
REDIS_URL=
CELERY_BROKER_URL=";

                File.WriteAllText(envPath, envContent);
                Console.WriteLine($" - ✓ .env template created at: {envPath}");
                Console.WriteLine(" - TIP: Fill in your actual values and never commit to Git");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[code.csx] Create .env error: {ex.Message}");
            }
            break;

        case "create_readme":
        case "readme":
            try
            {
                Console.WriteLine(" - Creating README.md template...");
                
                string workspacePath = GetCurrentWorkspacePath();
                if (string.IsNullOrEmpty(workspacePath))
                {
                    Console.WriteLine(" - ERROR: Could not determine workspace path");
                    break;
                }

                string projectName = Path.GetFileName(workspacePath) ?? "Project";
                string readmePath = Path.Combine(workspacePath, "README.md");
                
                string readmeContent = $@"# {projectName}

## Description
Brief description of your project.

## Installation

### Prerequisites
- Python 3.8+
- pip or uv package manager

### Setup
```bash
# Clone the repository
git clone <repository-url>
cd {projectName}

# Create virtual environment
python -m venv .venv

# Activate virtual environment
# Windows:
.venv\Scripts\activate
# Linux/Mac:
source .venv/bin/activate

# Install dependencies
pip install -r requirements.txt
```

## Usage
```bash
python main.py
```

## Features
- Feature 1
- Feature 2
- Feature 3

## Configuration
Configure the application by creating a `.env` file based on `.env.example`.

## Testing
```bash
pytest
```

## Contributing
Contributions are welcome! Please feel free to submit a Pull Request.

## License
MIT License
";

                File.WriteAllText(readmePath, readmeContent);
                Console.WriteLine($" - ✓ README.md created at: {readmePath}");
                Console.WriteLine(" - TIP: Edit the template to match your project");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[code.csx] Create README error: {ex.Message}");
            }
            break;

        // ==================== MAINTENANCE & CLEANUP ====================

        case "clear_pycache":
        case "clean_cache":
            try
            {
                Console.WriteLine(" - Cleaning __pycache__ directories...");
                
                string workspacePath = GetCurrentWorkspacePath();
                if (string.IsNullOrEmpty(workspacePath))
                {
                    Console.WriteLine(" - ERROR: Could not determine workspace path");
                    break;
                }

                int deletedCount = 0;
                string[] pycacheDirs = Directory.GetDirectories(workspacePath, "__pycache__", SearchOption.AllDirectories);
                
                foreach (string dir in pycacheDirs)
                {
                    try
                    {
                        Directory.Delete(dir, recursive: true);
                        deletedCount++;
                        Console.WriteLine($" - Deleted: {dir}");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($" - WARNING: Could not delete {dir}: {ex.Message}");
                    }
                }

                Console.WriteLine($" - ✓ Cleaned {deletedCount} __pycache__ director{(deletedCount == 1 ? "y" : "ies")}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[code.csx] Clear cache error: {ex.Message}");
            }
            break;

        case "clear_terminal":
        case "cls":
            try
            {
                Console.WriteLine(" - Clearing terminal...");
                
                if (!OpenTerminal()) break;
                
                string clearCmd = Environment.OSVersion.Platform == PlatformID.Win32NT ? "cls" : "clear";
                
                if (ExecuteTerminalCommand(clearCmd, 300))
                {
                    Console.WriteLine(" - ✓ Terminal cleared");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[code.csx] Clear terminal error: {ex.Message}");
            }
            break;

        // ==================== QUICK ACCESS COMMANDS ====================

        case "open_settings":
            try
            {
                Console.WriteLine(" - Opening VS Code settings...");
                
                if (SendKeysWithRetry("^,", 500))
                {
                    Console.WriteLine(" - ✓ Settings opened");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[code.csx] Open settings error: {ex.Message}");
            }
            break;

        case "open_extensions":
            try
            {
                Console.WriteLine(" - Opening Extensions panel...");
                
                if (SendKeysWithRetry("^+x", 500))
                {
                    Console.WriteLine(" - ✓ Extensions panel opened");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[code.csx] Open extensions error: {ex.Message}");
            }
            break;

        case "search_files":
        case "global_search":
            try
            {
                Console.WriteLine(" - Opening global search...");
                
                if (SendKeysWithRetry("^+f", 500))
                {
                    Console.WriteLine(" - ✓ Global search opened");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[code.csx] Global search error: {ex.Message}");
            }
            break;

        case "go_to_line":
            try
            {
                Console.WriteLine(" - Opening Go to Line dialog...");
                
                if (SendKeysWithRetry("^g", 300))
                {
                    Console.WriteLine(" - ✓ Go to Line dialog opened");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[code.csx] Go to line error: {ex.Message}");
            }
            break;

        case "format_document":
            try
            {
                Console.WriteLine(" - Formatting document...");
                
                if (SendKeysWithRetry("+%f", 800))
                {
                    Console.WriteLine(" - ✓ Document formatted");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[code.csx] Format document error: {ex.Message}");
            }
            break;

        case "toggle_comment":
            try
            {
                Console.WriteLine(" - Toggling comment...");
                
                if (SendKeysWithRetry("^/", 300))
                {
                    Console.WriteLine(" - ✓ Comment toggled");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[code.csx] Toggle comment error: {ex.Message}");
            }
            break;

        // ==================== HELP & INFORMATION ====================

        case "help":
        case "list_commands":
            try
            {
                Console.WriteLine("\n========== VS CODE PYTHON DEVELOPMENT AUTOMATION ==========");
                Console.WriteLine("\n📦 ENVIRONMENT MANAGEMENT:");
                Console.WriteLine("  • create_venv          - Create Python virtual environment (.venv)");
                Console.WriteLine("  • activate_venv        - Activate virtual environment");
                Console.WriteLine("  • deactivate_venv      - Deactivate virtual environment");
                Console.WriteLine("  • select_interpreter   - Open Python interpreter selector");
                
                Console.WriteLine("\n📚 PACKAGE MANAGEMENT:");
                Console.WriteLine("  • install_package:<pkg>        - Install Python package");
                Console.WriteLine("  • install_requirements[:uv]    - Install from requirements.txt");
                Console.WriteLine("  • upgrade_package:<pkg>        - Upgrade package");
                Console.WriteLine("  • uninstall_package:<pkg>      - Uninstall package");
                Console.WriteLine("  • list_packages                - List installed packages");
                Console.WriteLine("  • outdated_packages            - Check outdated packages");
                Console.WriteLine("  • freeze_requirements          - Save to requirements.txt");
                
                Console.WriteLine("\n🚀 RUNNING CODE:");
                Console.WriteLine("  • run_file[:<filename>]        - Run Python file");
                Console.WriteLine("  • run_flask[:<app>]            - Start Flask server");
                Console.WriteLine("  • run_django                   - Start Django server");
                Console.WriteLine("  • run_fastapi[:<module>]       - Start FastAPI with uvicorn");
                Console.WriteLine("  • run_streamlit[:<file>]       - Run Streamlit app");
                Console.WriteLine("  • http_server[:<port>]         - Start HTTP server");
                Console.WriteLine("  • run_jupyter                  - Start Jupyter Notebook");
                
                Console.WriteLine("\n🐛 DEBUGGING & TESTING:");
                Console.WriteLine("  • run_debug            - Start debugger (F5)");
                Console.WriteLine("  • run_pytest[:<path>]  - Run pytest");
                Console.WriteLine("  • run_unittest         - Run unittest");
                
                Console.WriteLine("\n📁 GIT OPERATIONS:");
                Console.WriteLine("  • git_init             - Initialize Git repository");
                Console.WriteLine("  • git_status           - Check Git status");
                Console.WriteLine("  • git_add_all          - Stage all changes");
                Console.WriteLine("  • git_commit:<msg>     - Commit with message");
                Console.WriteLine("  • git_push[:<branch>]  - Push to remote");
                Console.WriteLine("  • git_pull[:<branch>]  - Pull from remote");
                Console.WriteLine("  • git_log              - View commit history");
                
                Console.WriteLine("\n✨ FILE CREATION (INTELLIGENT):");
                Console.WriteLine("  • create_file                  - Create new file (asks for name & extension)");
                Console.WriteLine("  • create_file:python           - Create Python file (asks name only)");
                Console.WriteLine("  • create_file:javascript       - Create JavaScript file");
                Console.WriteLine("  • create_file:html|css|json|md - Create respective file type");
                
                Console.WriteLine("\n⚙️ PROJECT SETUP:");
                Console.WriteLine("  • create_gitignore     - Create Python .gitignore");
                Console.WriteLine("  • create_env           - Create .env template");
                Console.WriteLine("  • create_readme        - Create README.md template");
                
                Console.WriteLine("\n🧹 MAINTENANCE:");
                Console.WriteLine("  • clear_pycache        - Delete all __pycache__ folders");
                Console.WriteLine("  • clear_terminal       - Clear terminal output");
                
                Console.WriteLine("\n🔧 QUICK ACCESS:");
                Console.WriteLine("  • open_settings        - Open VS Code settings");
                Console.WriteLine("  • open_extensions      - Open Extensions panel");
                Console.WriteLine("  • search_files         - Global file search");
                Console.WriteLine("  • go_to_line           - Go to line dialog");
                Console.WriteLine("  • format_document      - Format current file");
                Console.WriteLine("  • toggle_comment       - Toggle line comment");
                
                Console.WriteLine("\n==========================================================\n");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[code.csx] Help display error: {ex.Message}");
            }
            break;

        default:
            Console.WriteLine($"[code.csx] ERROR: Unknown action '{actionName}'");
            Console.WriteLine(" - Type 'help' or 'list_commands' to see all available actions");
            Console.WriteLine(" - Common actions: create_venv, install_package, run_file, create_file");
            break;
    }

    Console.WriteLine("\n✓ Automation task complete.");
}
catch (Exception ex)
{
    Console.WriteLine($"\n[code.csx] CRITICAL ERROR executing action: {ex.Message}");
    Console.WriteLine($"   Stack trace: {ex.StackTrace}");
}
