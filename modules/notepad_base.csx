// notepad.csx
//akshay
//sfsdf2
using System;
using System.Threading;
using System.Windows.Forms;
using System.Windows.Automation;
using System.IO;
using System.Text;

if (AppContext == null) throw new Exception("AppContext is null inside module.");
if (string.IsNullOrWhiteSpace(Action)) throw new Exception("No Action provided.");

void Focus()
{
    try
    {
        var hwnd = (IntPtr)AppContext.Window.Current.NativeWindowHandle;
        if (hwnd != IntPtr.Zero) Win32.SetForegroundWindow(hwnd);
        Thread.Sleep(200);
    }
    catch { }
}

void Send(string keys, int delay = 150)
{
    SendKeys.SendWait(keys);
    Thread.Sleep(delay);
}

static class Win32
{
    [System.Runtime.InteropServices.DllImport("user32.dll")]
    public static extern bool SetForegroundWindow(IntPtr hWnd);
}

string actionName = Action;
string actionParam = "";
int colon = Action.IndexOf(':');
if (colon >= 0)
{
    actionName = Action.Substring(0, colon).Trim();
    actionParam = Action.Substring(colon + 1).Trim();
}

Console.WriteLine($"[notepad.csx] Action='{actionName}', Param='{actionParam}'");
Focus();

try
{
    switch (actionName.ToLower())
    {
        case "new_file":
            // Ctrl+N
            Send("^n", 200);
            break;

        case "type":
            if (!string.IsNullOrEmpty(actionParam))
            {
                Send(actionParam, 50);
            }
            else
            {
                Console.WriteLine("[notepad.csx] type action requires a parameter.");
            }
            break;

        case "save_as":
            // Ctrl+Shift+S triggers Save As; then paste path and Enter
            Send("^(+s)", 300);
            Thread.Sleep(300);
            // fallback to Alt+F, A
            Send("%(f)", 200);
            Thread.Sleep(120);
            Send("a", 400);
            Thread.Sleep(600);
            if (!string.IsNullOrEmpty(actionParam))
            {
                // send the path
                Send(actionParam, 300);
                Send("{ENTER}", 300);
            }
            else
            {
                Console.WriteLine("[notepad.csx] save_as requires full path parameter.");
            }
            break;

        case "save_with_name":
            // Format: save_with_name:filename|extension
            // Example: save_with_name:myfile|txt or save_with_name:data|csv
            if (!string.IsNullOrEmpty(actionParam))
            {
                string[] parts = actionParam.Split('|');
                if (parts.Length == 2)
                {
                    string filename = parts[0].Trim();
                    string extension = parts[1].Trim().TrimStart('.');
                    string fullPath = $"{filename}.{extension}";
                    
                    Console.WriteLine($"[notepad.csx] Saving as: {fullPath}");
                    
                    // Open Save As dialog
                    Send("^(+s)", 300);
                    Thread.Sleep(300);
                    // Fallback to Alt+F, A
                    Send("%(f)", 200);
                    Thread.Sleep(120);
                    Send("a", 400);
                    Thread.Sleep(600);
                    
                    // Send the full path
                    Send(fullPath, 300);
                    Send("{ENTER}", 300);
                }
                else
                {
                    Console.WriteLine("[notepad.csx] save_with_name requires format: filename|extension");
                }
            }
            else
            {
                Console.WriteLine("[notepad.csx] save_with_name requires filename|extension parameter.");
            }
            break;

        case "convert_csv_to_tsv":
            // Format: convert_csv_to_tsv:input_path|output_path
            // Example: convert_csv_to_tsv:C:\data.csv|C:\data.tsv
            if (!string.IsNullOrEmpty(actionParam))
            {
                string[] paths = actionParam.Split('|');
                if (paths.Length == 2)
                {
                    string inputPath = paths[0].Trim();
                    string outputPath = paths[1].Trim();
                    
                    if (File.Exists(inputPath))
                    {
                        Console.WriteLine($"[notepad.csx] Converting CSV to TSV: {inputPath} -> {outputPath}");
                        
                        string content = File.ReadAllText(inputPath);
                        // Replace commas with tabs
                        string converted = content.Replace(",", "\t");
                        File.WriteAllText(outputPath, converted);
                        
                        Console.WriteLine($"[notepad.csx] Conversion complete: {outputPath}");
                    }
                    else
                    {
                        Console.WriteLine($"[notepad.csx] Input file not found: {inputPath}");
                    }
                }
                else
                {
                    Console.WriteLine("[notepad.csx] convert_csv_to_tsv requires format: input_path|output_path");
                }
            }
            else
            {
                Console.WriteLine("[notepad.csx] convert_csv_to_tsv requires input_path|output_path parameter.");
            }
            break;

        case "convert_tsv_to_csv":
            // Format: convert_tsv_to_csv:input_path|output_path
            // Example: convert_tsv_to_csv:C:\data.tsv|C:\data.csv
            if (!string.IsNullOrEmpty(actionParam))
            {
                string[] paths = actionParam.Split('|');
                if (paths.Length == 2)
                {
                    string inputPath = paths[0].Trim();
                    string outputPath = paths[1].Trim();
                    
                    if (File.Exists(inputPath))
                    {
                        Console.WriteLine($"[notepad.csx] Converting TSV to CSV: {inputPath} -> {outputPath}");
                        
                        string content = File.ReadAllText(inputPath);
                        // Replace tabs with commas
                        string converted = content.Replace("\t", ",");
                        File.WriteAllText(outputPath, converted);
                        
                        Console.WriteLine($"[notepad.csx] Conversion complete: {outputPath}");
                    }
                    else
                    {
                        Console.WriteLine($"[notepad.csx] Input file not found: {inputPath}");
                    }
                }
                else
                {
                    Console.WriteLine("[notepad.csx] convert_tsv_to_csv requires format: input_path|output_path");
                }
            }
            else
            {
                Console.WriteLine("[notepad.csx] convert_tsv_to_csv requires input_path|output_path parameter.");
            }
            break;

        case "close":
            // Alt+F4
            Send("%{F4}", 200);
            break;

        default:
            Console.WriteLine($"[notepad.csx] Unknown action: {actionName}");
            break;
    }
}
catch (Exception ex)
{
    Console.WriteLine($"[notepad.csx] Error executing action: {ex.Message}");
}
