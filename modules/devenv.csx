// C# Automation Script for Visual Studio
// Process Name: devenv

Console.WriteLine("✅ Visual Studio Module Initialized.");
Thread.Sleep(500); // Give VS time to react

switch (Action?.ToLower())
{
    case "show_solution_explorer":
        Console.WriteLine("   - Executing action: show_solution_explorer");
        System.Windows.Forms.SendKeys.SendWait("^%l"); // Ctrl+Alt+L
        break;
        
    case "show_error_list":
        Console.WriteLine("   - Executing action: show_error_list");
        System.Windows.Forms.SendKeys.SendWait(@"^(\e)"); // Ctrl+\, E
        break;
        
    case "open_search":
        Console.WriteLine("   - Executing action: open_search");
        System.Windows.Forms.SendKeys.SendWait("^t"); // Ctrl+T
        break;

    default:
        Console.WriteLine($"   - Unknown or missing action: '{Action}'");
        Console.WriteLine("   - Available actions for Visual Studio: show_solution_explorer, show_error_list, open_search");
        break;
}
