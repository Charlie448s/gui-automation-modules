// Blender.csx — Robust Blender automation module
// Version 1
// Compatible with ModuleManager.cs
// Purpose: Initialize Blender setup and automate common UI tasks

using System;
using System.IO;
using System.Threading;
using System.Diagnostics;
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

const int MAX_RETRIES = 3;
const int BASE_DELAY = 200;

// ----------------- Win32 Helpers -----------------
static class Win32
{
    [System.Runtime.InteropServices.DllImport("user32.dll")]
    public static extern bool SetForegroundWindow(IntPtr hWnd);

    [System.Runtime.InteropServices.DllImport("user32.dll")]
    public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
}

// ----------------- Utility Functions -----------------
bool FocusWindow()
{
    try
    {
        var hwnd = (IntPtr)AppContext.Window.Current.NativeWindowHandle;
        if (hwnd == IntPtr.Zero) return false;

        Win32.ShowWindow(hwnd, 9); // SW_RESTORE
        Thread.Sleep(300);
        Win32.SetForegroundWindow(hwnd);
        Thread.Sleep(300);
        return true;
    }
    catch { return false; }
}

bool SendKeysSafe(string keys, int delay = BASE_DELAY)
{
    try
    {
        SendKeys.SendWait(keys);
        Thread.Sleep(delay);
        return true;
    }
    catch { return false; }
}

// ----------------- Action Parsing -----------------
string actionName = Action;
string actionParam = string.Empty;
int colon = Action.IndexOf(':');
if (colon >= 0)
{
    actionName = Action.Substring(0, colon).Trim();
    actionParam = Action.Substring(colon + 1).Trim();
}

Console.WriteLine("✓ Blender Automation Module Loaded!");
Console.WriteLine($" - Process: {AppContext.Window.Current.ProcessId}");
Console.WriteLine($" - Action : {actionName}");

if (!FocusWindow())
{
    Console.WriteLine(" - ERROR: Could not focus Blender window.");
    return;
}

// ----------------- Main Automation Logic -----------------
try
{
    switch (actionName.ToLower())
    {
        // --- INITIALIZATION FEATURE ---
        case "init":
        case "initialize":
            Console.WriteLine(" - Initializing Blender startup setup...");
            Console.WriteLine(" - Choosing 'General' preset...");
            SendKeysSafe("{ENTER}", 600); // confirm General

            Thread.Sleep(1000);
            Console.WriteLine(" - Setting right-click as select...");
            SendKeysSafe("^,"); // open Preferences
            Thread.Sleep(600);
            SendKeysSafe("select with right{ENTER}");
            Thread.Sleep(600);
            SendKeysSafe("%{F4}"); // close preferences
            Console.WriteLine(" - Setup complete (General + Right Click Select)");
            break;

        // --- COMMON UI ACTIONS ---
        case "toggle_overlay":
            Console.WriteLine(" - Toggling Overlays (Viewport)... (Shift + Alt + Z)");
            SendKeysSafe("+%z", 200);
            break;

        case "toggle_gizmo":
            Console.WriteLine(" - Toggling Gizmo visibility...");
            SendKeysSafe("^`", 200);
            break;

        case "open_preferences":
            Console.WriteLine(" - Opening Preferences window...");
            SendKeysSafe("^,", 200);
            break;

        case "save_project":
            Console.WriteLine(" - Saving Blender project...");
            SendKeysSafe("^s", 300);
            break;

        case "render_image":
            Console.WriteLine(" - Rendering current frame (F12)...");
            SendKeysSafe("{F12}", 300);
            break;

        case "render_animation":
            Console.WriteLine(" - Rendering full animation (Ctrl + F12)...");
            SendKeysSafe("^({F12})", 300);
            break;

        case "quick_material":
            Console.WriteLine(" - Applying basic material setup...");
            SendKeysSafe("n", 300); // open sidebar
            Thread.Sleep(300);
            SendKeysSafe("{TAB}", 300); // switch to edit mode
            Thread.Sleep(300);
            SendKeysSafe("a", 200); // select all
            SendKeysSafe("u", 200); // unwrap UV
            Thread.Sleep(300);
            SendKeysSafe("shift+a", 200);
            SendKeysSafe("material{ENTER}", 400);
            break;

        case "quick_light":
            Console.WriteLine(" - Adding light to scene...");
            SendKeysSafe("shift+a", 300);
            SendKeysSafe("light{ENTER}", 300);
            SendKeysSafe("{DOWN}{DOWN}{ENTER}", 300);
            break;

        case "quick_camera":
            Console.WriteLine(" - Adding camera to scene...");
            SendKeysSafe("shift+a", 300);
            SendKeysSafe("camera{ENTER}", 300);
            break;

        case "list_ui":
            Console.WriteLine(" - Common Blender UI automation targets:");
            Console.WriteLine("   init, toggle_overlay, toggle_gizmo, open_preferences, save_project,");
            Console.WriteLine("   render_image, render_animation, quick_material, quick_light, quick_camera");
            break;

        default:
            Console.WriteLine($" - Unknown action '{actionName}'. Use 'list_ui' to view all actions.");
            break;
    }
}
catch (Exception ex)
{
    Console.WriteLine($"[blender.csx] Error executing {actionName}: {ex.Message}");
}
