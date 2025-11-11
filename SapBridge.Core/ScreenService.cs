using Serilog;
using SapBridge.Core.Models;

namespace SapBridge.Core;

public class ScreenService
{
    private readonly ILogger _logger;
    private readonly ComIntrospector _introspector;

    public ScreenService(ILogger logger, ComIntrospector introspector)
    {
        _logger = logger;
        _introspector = introspector;
    }

    public ScreenState GetScreenState(dynamic session)
    {
        var screenState = new ScreenState
        {
            Transaction = GetSafeProperty(session.Info, "Transaction") ?? "",
            ScreenNumber = GetSafeProperty(session.Info, "ScreenNumber") ?? "",
            ProgramName = GetSafeProperty(session.Info, "Program") ?? "",
        };

        // Get title bar
        try
        {
            var mainWindow = session.FindById("wnd[0]");
            screenState.TitleBar = GetSafeProperty(mainWindow, "Text") ?? "";
        }
        catch { }

        // Get status bar
        try
        {
            var statusbar = session.FindById("wnd[0]/sbar");
            screenState.StatusBar = new StatusBarInfo
            {
                Text = GetSafeProperty(statusbar, "Text") ?? "",
                Type = GetStatusBarType(statusbar)
            };

            // Get all messages
            try
            {
                int messageCount = statusbar.MessageCount;
                for (int i = 0; i < messageCount; i++)
                {
                    string message = statusbar.GetMessage(i);
                    screenState.StatusBar.Messages.Add(message);
                }
            }
            catch { }
        }
        catch { }

        // Get all objects in the main window
        try
        {
            var mainWindow = session.FindById("wnd[0]");
            TraverseObjects(mainWindow, "wnd[0]", screenState.Objects);
        }
        catch (Exception ex)
        {
            _logger.Error(ex, "Error traversing screen objects");
        }

        return screenState;
    }

    private void TraverseObjects(dynamic parent, string parentPath, List<ObjectInfo> objects, int depth = 0)
    {
        // Limit recursion depth
        if (depth > 10) return;

        try
        {
            var children = parent.Children;
            if (children == null) return;

            for (int i = 0; i < children.Count; i++)
            {
                try
                {
                    var child = children.ElementAt(i);
                    string childId = GetSafeProperty(child, "Id") ?? $"{parentPath}/child[{i}]";
                    string childType = GetSafeProperty(child, "Type") ?? "Unknown";

                    // Introspect this object
                    var objectInfo = _introspector.IntrospectObject(child, childId);
                    objects.Add(objectInfo);

                    // Recursively traverse if it's a container
                    if (IsContainer(childType))
                    {
                        TraverseObjects(child, childId, objects, depth + 1);
                    }
                }
                catch (Exception ex)
                {
                    _logger.Debug(ex, $"Error traversing child {i} of {parentPath}");
                }
            }
        }
        catch (Exception ex)
        {
            _logger.Debug(ex, $"Error getting children of {parentPath}");
        }
    }

    private bool IsContainer(string type)
    {
        var containerTypes = new[] 
        { 
            "GuiUserArea", "GuiSimpleContainer", "GuiScrollContainer",
            "GuiSplitterContainer", "GuiBox", "GuiTab", "GuiTabStrip",
            "GuiShell", "GuiCustomControl"
        };
        
        return containerTypes.Contains(type);
    }

    private string GetStatusBarType(dynamic statusbar)
    {
        try
        {
            string messageType = statusbar.MessageType;
            return messageType?.ToLower() switch
            {
                "s" => "success",
                "e" => "error",
                "w" => "warning",
                "i" or "a" => "info",
                _ => "info"
            };
        }
        catch
        {
            return "info";
        }
    }

    private string? GetSafeProperty(dynamic obj, string propertyName)
    {
        try
        {
            var value = obj.GetType().InvokeMember(
                propertyName,
                System.Reflection.BindingFlags.GetProperty,
                null,
                obj,
                null
            );
            return value?.ToString();
        }
        catch
        {
            return null;
        }
    }
}

