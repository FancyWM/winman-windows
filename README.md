# WinMan (Windows/Win32 implementation) <img src="https://raw.githubusercontent.com/veselink1/winman/master/Resources/Icon.png" alt="Logo" width="96">

### WinMan - a window management library for .NET with the following features:
 - Platform-agnostic design
 - Event-based API
 - Designed to be used by window management software

### See the [Documentation](https://veselink1.github.io/winman-windows/WinMan.Windows.html).

---
### WinMan is a dependency of [FancyWM](https://www.microsoft.com/en-us/p/fancywm/9p1741lkhqs9).


## Example 1
Prints added, removed and initially present windows over the course of 10 seconds.
```csharp
using System;

using WinMan;
using WinMan.Windows;

void TestWinMan()
{
    using IWorkspace workspace = new Win32Workspace();
    workspace.WindowManaging += (s, e) =>
    {
        Console.WriteLine($"Window {e.Source} initially present on the workspace!");
    };
    workspace.WindowAdded += (s, e) =>
    {
        Console.WriteLine($"Window {e.Source} added to the workspace!");
    };
    workspace.WindowRemoved += (s, e) =>
    {
        Console.WriteLine($"Window {e.Source} removed from the workspace!");
    };
    workspace.Open();
    Thread.Sleep(100000);
}

```

## Example 2
Listens to changes to an empty notepad window.
```csharp
using System;

using WinMan;
using WinMan.Windows;

void TestWinMan()
{
    using IWorkspace workspace = new Win32Workspace();
    workspace.Open();
    var notepad = workspace.GetSnapshot().First(x => x.Title == "Notepad -- Untitled");
    notepad.StateChanged += (s, e) =>
    {
        Console.WriteLine($"Notepad is now {e.NewState}!");
    };
    Thread.Sleep(100000);
}

```
