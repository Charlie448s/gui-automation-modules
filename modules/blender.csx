using System;
using System.Diagnostics;
using System.IO;
using System.Threading;
using System.Windows.Forms;

if (AppContext == null) throw new Exception("AppContext is null.");
if (string.IsNullOrWhiteSpace(Action)) throw new Exception("No Action provided.");

Console.WriteLine("✓ Blender Automation Module Loaded!");
Console.WriteLine($" - Process ID: {AppContext.Window.Current.ProcessId}");
Console.WriteLine($" - Action: {Action}");

string blenderPath = @"C:\Program Files\Blender Foundation\Blender 4.0\blender.exe"; // adjust as needed
string tempScriptPath = Path.Combine(Path.GetTempPath(), "blender_auto.py");

string scriptContent = @"
import bpy

print('--- Blender Automation Initialized ---')

# 1️⃣ Reset to General startup
bpy.ops.wm.read_homefile(app_template='')  # loads General template

# 2️⃣ Ensure right-click select is enabled
prefs = bpy.context.preferences
prefs.input.select_mouse = 'RIGHT'
prefs.view.use_zoom_to_mouse = True

# 3️⃣ Enable useful overlays and tools
area = next(a for a in bpy.context.screen.areas if a.type == 'VIEW_3D')
space = next(s for s in area.spaces if s.type == 'VIEW_3D')
space.overlay.show_cursor = True
space.overlay.show_stats = True
space.overlay.show_object_origins = True
space.shading.show_xray = True
space.shading.type = 'SOLID'

# 4️⃣ Create base layout: Cube, Plane, and Camera properly positioned
bpy.ops.object.select_all(action='SELECT')
bpy.ops.object.delete(use_global=False)

bpy.ops.mesh.primitive_cube_add(size=2, location=(0, 1, 1))
bpy.ops.mesh.primitive_plane_add(size=10, location=(0, 0, 0))
bpy.ops.object.camera_add(location=(6, -6, 5), rotation=(1.1, 0, 0.78))
bpy.ops.object.light_add(type='SUN', radius=1, location=(4, -4, 6))

# 5️⃣ Enable screencast keys (if installed)
try:
    bpy.ops.preferences.addon_enable(module='space_view3d_screencast_keys')
except Exception:
    print('Screencast keys not available.')

# 6️⃣ Save setup
bpy.ops.wm.save_mainfile(filepath=bpy.path.abspath('//auto_init.blend'))

print('--- Blender Automation Complete ---')
";

File.WriteAllText(tempScriptPath, scriptContent);

try
{
    Console.WriteLine(" - Launching Blender with automation script...");
    var process = new Process();
    process.StartInfo.FileName = blenderPath;
    process.StartInfo.Arguments = $"--python \"{tempScriptPath}\"";
    process.StartInfo.UseShellExecute = false;
    process.StartInfo.CreateNoWindow = false;
    process.Start();

    Console.WriteLine(" - Blender launched and executing setup script.");
}
catch (Exception ex)
{
    Console.WriteLine($"[blender.csx] ERROR: Failed to launch Blender - {ex.Message}");
}

Console.WriteLine(" - Automation task complete.");
