// This is an executable C# script for Visual Studio Code.
// It will be run by the GuiAutomationAgent.

// The 'AppContext' global variable is passed in from the main program.
// It contains information about the VS Code process and window.

Console.WriteLine("✅ VS Code Automation Module Loaded!");
Console.WriteLine($"   - App Name: {AppContext.Name}");
Console.WriteLine($"   - Process ID: {AppContext.ProcessId}");

// --- Example Automation Logic ---
// This example will open the command palette and search for the GitLens view.

try
{
    Console.WriteLine("   - Automating: Opening Command Palette...");

    // Use SendKeys to simulate keyboard shortcuts.
    // Make sure the window is focused before sending keys.
    // Note: SendKeys can be unreliable. UIAutomation is a more robust alternative for complex tasks.

    System.Windows.Forms.SendKeys.SendWait("^+p"); // Ctrl+Shift+P
    Thread.Sleep(500); // Wait for the palette to open

    Console.WriteLine("   - Automating: Typing 'View: Show GitLens'...");
    System.Windows.Forms.SendKeys.SendWait("View: Show GitLens");
    Thread.Sleep(500);

    System.Windows.Forms.SendKeys.SendWait("{ENTER}"); // Press Enter

    Console.WriteLine("   - Automation task complete.");
}
catch (Exception e)
{
    Console.WriteLine($"   - Automation Error: {e.Message}");
}
