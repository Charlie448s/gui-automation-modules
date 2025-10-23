// C# Automation Script for Visual Studio
// Process Name: devenv

Console.WriteLine("✅ Visual Studio Automation Module Loaded!");
Console.WriteLine($"   - App Name: {AppContext.Name}");
Console.WriteLine($"   - Process ID: {AppContext.ProcessId}");

// Allow time for Visual Studio to fully render and gain focus.
Thread.Sleep(1000);

try
{
    Console.WriteLine("   - Automating: Ensuring Solution Explorer is visible...");
    // Simulate pressing Ctrl+Alt+L to toggle the Solution Explorer window.
    // This is a reliable way to make sure it's open.
    System.Windows.Forms.SendKeys.SendWait("^%l"); // ^ = Ctrl, % = Alt
    Thread.Sleep(500);

    Console.WriteLine("   - Automating: Opening the Error List window...");
    // The shortcut for Error List is Ctrl+\ followed by E.
    // SendKeys handles sequences like this by enclosing the second part in parentheses.
    System.Windows.Forms.SendKeys.SendWait(@"^(\e)");
    Thread.Sleep(500);
    
    Console.WriteLine("   - Automating: Opening the 'Go To All' search bar...");
    // Simulate pressing Ctrl+T to open the main search/navigation bar.
    System.Windows.Forms.SendKeys.SendWait("^t");
    Thread.Sleep(500);
    
    Console.WriteLine("   - Automating: Closing the search bar...");
    // Press Escape to close the dialog/search bar.
    System.Windows.Forms.SendKeys.SendWait("{ESC}");
    
    Console.WriteLine("   - Visual Studio automation task complete.");
}
catch (Exception e)
{
    Console.WriteLine($"   - Automation Error: {e.Message}");
}
