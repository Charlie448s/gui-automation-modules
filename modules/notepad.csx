// C# Automation Script for Notepad
// Process Name: notepad

Console.WriteLine("✅ Notepad Module Initialized.");

// The 'Action' global variable holds the command from the agent.
// We use ?.ToLower() for safety and case-insensitivity.
switch (Action?.ToLower())
{
    case "write_greeting":
        Console.WriteLine("   - Executing action: write_greeting");
        Thread.Sleep(200);
        System.Windows.Forms.SendKeys.SendWait("Hello from an interactive module!");
        System.Windows.Forms.SendKeys.SendWait("{ENTER}");
        break;

    case "save_as":
        Console.WriteLine("   - Executing action: save_as");
        Thread.Sleep(200);
        System.Windows.Forms.SendKeys.SendWait("^s"); // Ctrl+S
        Thread.Sleep(300);
        System.Windows.Forms.SendKeys.SendWait("MyNewFile.txt");
        break;
        
    default:
        Console.WriteLine($"   - Unknown or missing action: '{Action}'");
        Console.WriteLine("   - Available actions for Notepad: write_greeting, save_as");
        break;
}
