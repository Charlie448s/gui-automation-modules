// C# Automation Script for Visual Studio
// Process Name: devenv

Console.WriteLine("✅ Visual Studio Module Initialized.");

// Add a small delay to ensure VS is ready for input.
Thread.Sleep(500); 

// Use a switch statement to check the 'Action' variable passed from the agent.
switch (Action?.ToLower())
{
    case "show_solution_explorer":
        Console.WriteLine("   - Executing action: show_solution_explorer");
        // Shortcut: Ctrl+Alt+L
        System.Windows.Forms.SendKeys.SendWait("^%l"); 
        break;
        
    case "show_error_list":
        Console.WriteLine("   - Executing action: show_error_list");
        // Shortcut: Ctrl+\ followed by E
        System.Windows.Forms.SendKeys.SendWait(@"^(\e)"); 
        break;
        
    case "open_search":
        Console.WriteLine("   - Executing action: open_search");
        // Shortcut: Ctrl+T
        System.Windows.Forms.SendKeys.SendWait("^t");
        break;

    // This is the default case if the action is unknown.
    default:
        Console.WriteLine($"   - Unknown or missing action: '{Action}'");
        Console.WriteLine("   - Available actions: show_solution_explorer, show_error_list, open_search");
        break;
}

Console.WriteLine("   - Action complete.");
