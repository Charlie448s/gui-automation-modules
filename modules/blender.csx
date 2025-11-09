// blender.csx
// Robust Blender automation module with focus, launch, and initialization features
// Compatible with ModuleManager.cs
// Version 2.0 — GPT-5 Edition

using System;
using System.IO;
using System.Diagnostics;
using System.Threading;
using System.Linq;
using System.Windows.Forms;

// -------------------------------
// Globals provided by ModuleManager
// -------------------------------
//   AppContext  -> ApplicationContext
//   Action      -> string
// -------------------------------

if (AppContext == null) throw new Exception("AppContext is null.");
if (string.IsNullOrWhiteSpace(Action)) throw new Exception("No Action provided.");

Console.WriteLine("✓ Blender Automation Module Loaded!");
Console.WriteLine($" - Process ID : {AppContext.Window.Current.ProcessId}");
Console.WriteLine($" - Action     : {Action}");

// ----------- CONFIG -----------
string blenderSteamPath   = @"C:\Program Files (x86)\Steam\steamapps\common\Blender\blender.exe";
string blenderFoundationPath = @"C:\Program Files\Blender Foundation\Blender 4.0\blender.exe";
string blenderExecutable = File.Exists(blenderFoundationPath)
    ? blenderFoundationPath
    : (File.Exists(blenderSteamPath) ? blenderSteamPath : string.Empty);

if (string.IsNullOrEmpty(blenderExecutable))
{
    Console.WriteLine("[blender.csx] ERROR: Blender executable not found. Please check installation path.");
    return;
}

// ----------- UTILS -----------

static class Win32
{
    [System.Runtime.InteropServices.DllImport("user32.dll")]
    public static extern bool SetForegroundWindow(IntPtr hWnd);

    [System.Runtime.InteropServices.DllImport("user32.dll")]
    public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
}

bool FocusBlender(int retries = 3)
{
    for (int attempt = 1; attempt <= retries; attempt++)
    {
        try
        {
            var blenderProc = Process.GetProcessesByName("blender").FirstOrDefault();
            if (blenderProc == null)
            {
                Console.WriteLine(" - Blender not running, launching...");
                Process.Start(blenderExecutable);
                Thread.Sleep(5000);
                continue;
            }

            IntPtr hwnd = blenderProc.MainWindowHandle;
            if (hwnd == IntPtr.Zero)
            {
                Console.WriteLine(" - Invalid window handle, retrying...");
                Thread.Sleep(300 * attempt);
                continue;
            }

            Win32.ShowWindow(hwnd, 9); // SW_RESTORE
            Thread.Sleep(200);
            if (Win32.SetForegroundWindow(hwnd))
            {
                Console.WriteLine("✓ Blender focused successfully");
                return true;
            }

            Console.WriteLine(" - Failed to set foreground, retrying...");
            Thread.Sleep(500 * attempt);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"[blender.csx] Focus error (attempt {attempt}): {ex.Message}");
            Thread.Sleep(500 * attempt);
        }
    }
    return false;
}

bool LaunchBlenderWithScript(string scriptPath)
{
    try
    {
        var psi = new ProcessStartInfo
        {
            FileName = blenderExecutable,
            Arguments = $"--python \"{scriptPath}\"",
            UseShellExecute = false,
            CreateNoWindow = false
        };
        Process.Start(psi);
        Console.WriteLine(" - Blender launched with automation script.");
        return true;
    }
    catch (Exception ex)
    {
        Console.WriteLine($"[blender.csx] Launch error: {ex.Message}");
        return false;
    }
}

// ----------- ACTION PARSING -----------
string actionName = Action;
string actionParam = string.Empty;
int colon = Action.IndexOf(':');
if (colon >= 0)
{
    actionName = Action.Substring(0, colon).Trim();
    actionParam = Action.Substring(colon + 1).Trim();
}

// ----------- MAIN AUTOMATION -----------
try
{
    switch (actionName.ToLower())
    {
        // ✅ Focus existing Blender window or launch new one
        case "focus":
        case "launch":
            if (FocusBlender())
                Console.WriteLine(" - Blender window focused or launched successfully.");
            else
                Console.WriteLine(" - ERROR: Could not focus or launch Blender.");
            break;

        // ✅ Initialize Blender with right-click, overlays, and base scene
        case "init":
        case "initialize":
        case "setup":
            string tempScriptPath = Path.Combine(Path.GetTempPath(), "blender_auto_init.py");
            string pythonScript = @"
import bpy

print('--- Blender Automation Init ---')

# 1️⃣ Reset to General template
bpy.ops.wm.read_homefile(app_template='')

# 2️⃣ Set right-click select
prefs = bpy.context.preferences
prefs.input.select_mouse = 'RIGHT'
prefs.view.use_zoom_to_mouse = True

# 3️⃣ Enable overlays
for area in bpy.context.screen.areas:
    if area.type == 'VIEW_3D':
        space = next(s for s in area.spaces if s.type == 'VIEW_3D')
        space.overlay.show_cursor = True
        space.overlay.show_stats = True
        space.overlay.show_object_origins = True
        space.shading.show_xray = True
        space.shading.type = 'SOLID'
        break

# 4️⃣ Create basic scene
bpy.ops.object.select_all(action='SELECT')
bpy.ops.object.delete(use_global=False)

bpy.ops.mesh.primitive_cube_add(size=2, location=(0, 1, 1))
bpy.ops.mesh.primitive_plane_add(size=10, location=(0, 0, 0))
bpy.ops.object.camera_add(location=(6, -6, 5), rotation=(1.1, 0, 0.78))
bpy.ops.object.light_add(type='SUN', radius=1, location=(4, -4, 6))

# 5️⃣ Save setup
bpy.ops.wm.save_mainfile(filepath=bpy.path.abspath('//auto_init.blend'))

print('--- Blender Automation Completed ---')
";
            File.WriteAllText(tempScriptPath, pythonScript);

            Console.WriteLine(" - Launching Blender for initialization...");
            if (!LaunchBlenderWithScript(tempScriptPath))
            {
                Console.WriteLine(" - ERROR: Failed to start Blender with script.");
                break;
            }
            break;

        // 🧱 Build a quick structure (walls + roof)
        case "build_structure":
            string structureScript = Path.Combine(Path.GetTempPath(), "blender_build_structure.py");
            string structurePy = @"
import bpy

print('--- Blender Build Structure Automation ---')

bpy.ops.object.select_all(action='SELECT')
bpy.ops.object.delete(use_global=False)

# Base floor
bpy.ops.mesh.primitive_plane_add(size=10, location=(0, 0, 0))

# 4 walls
for x in (-5, 5):
    bpy.ops.mesh.primitive_cube_add(size=1, location=(x, 0, 2.5))
    bpy.context.object.scale[0] = 0.1
    bpy.context.object.scale[1] = 5
    bpy.context.object.scale[2] = 2.5

for y in (-5, 5):
    bpy.ops.mesh.primitive_cube_add(size=1, location=(0, y, 2.5))
    bpy.context.object.scale[0] = 5
    bpy.context.object.scale[1] = 0.1
    bpy.context.object.scale[2] = 2.5

# Roof
bpy.ops.mesh.primitive_plane_add(size=10, location=(0, 0, 5))

# Add Sun light
bpy.ops.object.light_add(type='SUN', location=(8, -8, 10))

# Save
bpy.ops.wm.save_mainfile(filepath=bpy.path.abspath('//auto_structure.blend'))

print('--- Structure Build Complete ---')
";
            File.WriteAllText(structureScript, structurePy);

            Console.WriteLine(" - Building structure in Blender...");
            if (!LaunchBlenderWithScript(structureScript))
            {
                Console.WriteLine(" - ERROR: Failed to build structure.");
                break;
            }
            break;

        default:
            Console.WriteLine($"[blender.csx] ERROR: Unknown action '{actionName}'");
            Console.WriteLine(" - Available actions: focus, init, build_structure");
            break;
    }

    Console.WriteLine(" - Automation task complete.");
}
catch (Exception ex)
{
    Console.WriteLine($"[blender.csx] CRITICAL ERROR executing action: {ex.Message}");
    Console.WriteLine(ex.StackTrace);
}
