// C# Automation Script for Notepad
// Process Name: notepad

// The 'AppContext' global is passed in from the main program.
Console.WriteLine("✅ Notepad Automation Module Loaded!");
Console.WriteLine($"   - App Name: {AppContext.Name}");
Console.WriteLine($"   - Process ID: {AppContext.ProcessId}");

// Give a moment for the window to be fully focused.
Thread.Sleep(500);

try
{
    string message = "Hello from the GUI Automation Agent! This text was typed automatically.";
    Console.WriteLine($"   - Automating: Typing a message...");

    // Simulate typing the message character by character.
    System.Windows.Forms.SendKeys.SendWait(message);

    // Add two new lines.
    System.Windows.Forms.SendKeys.SendWait("{ENTER}{ENTER}");
    Thread.Sleep(300);

    Console.WriteLine($"   - Automating: Opening the 'Save As' dialog...");
    // Simulate pressing Ctrl+S to open the "Save" or "Save As" dialog.
    System.Windows.Forms.SendKeys.SendWait("^s"); 
    Thread.Sleep(500);

    Console.WriteLine($"   - Automating: Typing a default file name...");
    System.Windows.Forms.SendKeys.SendWait("AutomatedFile.txt");
    Thread.Sleep(300);

    // The script stops here. You could extend it to press Enter to save.
    // For now, we leave the "Save As" dialog open.
    Console.WriteLine("   - Notepad automation task complete.");
}
catch (Exception e)
{
    Console.WriteLine($"   - Automation Error: {e.Message}");
}
