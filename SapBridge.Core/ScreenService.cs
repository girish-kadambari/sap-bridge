using System.Reflection;
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

    public ScreenState GetScreenState(object session)
    {
        object info = GetProperty(session, "Info")!;
        
        var screenState = new ScreenState
        {
            Transaction = GetSafeProperty(info, "Transaction") ?? "",
            ScreenNumber = GetSafeProperty(info, "ScreenNumber") ?? "",
            ProgramName = GetSafeProperty(info, "Program") ?? "",
        };

        // Get title bar
        try
        {
            var mainWindow = InvokeMethod(session, "FindById", "wnd[0]");
            if (mainWindow != null)
                screenState.TitleBar = GetSafeProperty(mainWindow, "Text") ?? "";
        }
        catch { }

        // Get status bar
        try
        {
            var statusbar = InvokeMethod(session, "FindById", "wnd[0]/sbar");
            if (statusbar != null)
            {
                screenState.StatusBar = new StatusBarInfo
                {
                    Text = GetSafeProperty(statusbar, "Text") ?? "",
                    Type = GetStatusBarType(statusbar)
                };

                // Get all messages
                try
                {
                    int messageCount = (int)(GetProperty(statusbar, "MessageCount") ?? 0);
                    for (int i = 0; i < messageCount; i++)
                    {
                        string? message = InvokeMethod(statusbar, "GetMessage", i) as string;
                        if (message != null)
                            screenState.StatusBar.Messages.Add(message);
                    }
                }
                catch { }
            }
        }
        catch { }

        // Get all objects in the main window
        try
        {
            var mainWindow = InvokeMethod(session, "FindById", "wnd[0]");
            if (mainWindow != null)
                TraverseObjects(mainWindow, "wnd[0]", screenState.Objects);
        }
        catch (Exception ex)
        {
            _logger.Error(ex, "Error traversing screen objects");
        }

        return screenState;
    }

    private void TraverseObjects(object parent, string parentPath, List<ObjectInfo> objects, int depth = 0)
    {
        // Limit recursion depth
        if (depth > 10) return;

        try
        {
            var children = GetProperty(parent, "Children");
            if (children == null) return;

            int childCount = (int)(GetProperty(children, "Count") ?? 0);
            for (int i = 0; i < childCount; i++)
            {
                try
                {
                    var child = InvokeMethod(children, "ElementAt", i);
                    if (child == null) continue;
                    
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

    private string GetStatusBarType(object statusbar)
    {
        try
        {
            string? messageType = GetSafeProperty(statusbar, "MessageType");
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

    private string? GetSafeProperty(object obj, string propertyName)
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

    // Reflection helpers
    private static object? InvokeMethod(object obj, string methodName, params object[] parameters)
    {
        Type type = obj.GetType();
        return type.InvokeMember(
            methodName,
            BindingFlags.InvokeMethod,
            null,
            obj,
            parameters
        );
    }

    private static object? GetProperty(object obj, string propertyName)
    {
        Type type = obj.GetType();
        return type.InvokeMember(
            propertyName,
            BindingFlags.GetProperty,
            null,
            obj,
            null
        );
    }
}

