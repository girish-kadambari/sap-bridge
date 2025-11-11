using System.Reflection;
using System.Text.Json;
using Serilog;
using SapBridge.Core.Models;
using PropertyInfo = SapBridge.Core.Models.PropertyInfo;

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

        // Get all objects in the main window - USE OPTIMIZED GetObjectTree
        // FIX: Ensure we fall back to slow traversal if GetObjectTree returns 0 objects
        // Previous issue: TryGetObjectTreeOptimized returned true even with 0 objects,
        // preventing the fallback from ever running
        try
        {
            _logger.Debug("Starting object retrieval...");
            int objectCountBefore = screenState.Objects.Count;
            
            // Try the fast path first (GetObjectTree - 10-100x faster!)
            if (TryGetObjectTreeOptimized(session, screenState.Objects))
            {
                int retrieved = screenState.Objects.Count - objectCountBefore;
                _logger.Information($"Used optimized GetObjectTree method - retrieved {retrieved} objects");
            }
            else
            {
                // Fallback to slow traversal if GetObjectTree not available
                _logger.Warning("GetObjectTree not available or returned 0 objects, falling back to slow traversal");
                var mainWindow = InvokeMethod(session, "FindById", "wnd[0]");
                if (mainWindow != null)
                {
                    _logger.Debug("Starting slow object traversal...");
                    TraverseObjects(mainWindow, "wnd[0]", screenState.Objects);
                    int retrieved = screenState.Objects.Count - objectCountBefore;
                    _logger.Information($"Slow traversal retrieved {retrieved} objects");
                }
                else
                {
                    _logger.Error("Could not find main window wnd[0]");
                }
            }
            
            _logger.Information($"Total objects in screen state: {screenState.Objects.Count}");
        }
        catch (Exception ex)
        {
            _logger.Error(ex, "Error getting screen objects");
        }

        return screenState;
    }

    /// <summary>
    /// Optimized object retrieval using GetObjectTree (SAP GUI 7.70 patch 3+)
    /// Returns entire object tree in single JSON call - 10-100x faster than individual COM calls
    /// </summary>
    private bool TryGetObjectTreeOptimized(object session, List<ObjectInfo> objects)
    {
        try
        {
            // Define properties we want to retrieve
            // Using property names (can also use Magic DispIDs for even better performance)
            string[] properties = new string[]
            {
                "Id",           // Object path/ID
                "Type",         // GuiTextField, GuiButton, etc.
                "SubType",      // For shell components: Grid, Tree, Table, etc.
                "Name",         // Field name (e.g., RSYST-MANDT)
                "Text",         // Current text value
                "Tooltip",      // Tooltip text
                "Changeable",   // Whether field is editable
                "Modified",     // Whether field was modified
                "Left",         // Position
                "Top",
                "Width",
                "Height",
                "MaxLength"     // For text fields
            };

            // Call GetObjectTree - returns JSON string
            // Empty string for ID = get full tree from wnd[0]
            object? result = InvokeMethod(session, "GetObjectTree", "", properties);
            
            if (result == null)
            {
                _logger.Debug("GetObjectTree returned null");
                return false;
            }

            string json = result.ToString() ?? "";
            if (string.IsNullOrEmpty(json))
            {
                _logger.Debug("GetObjectTree returned empty string");
                return false;
            }

            // Parse JSON into objects
            int initialCount = objects.Count;
            ParseObjectTreeJson(json, objects);
            
            int retrievedCount = objects.Count - initialCount;
            _logger.Information($"GetObjectTree retrieved {retrievedCount} objects");
            
            // Return false if no objects were retrieved, so we fall back to slow traversal
            if (retrievedCount == 0)
            {
                _logger.Warning("GetObjectTree returned 0 objects, falling back to slow method");
                return false;
            }
            
            return true;
        }
        catch (Exception ex)
        {
            // GetObjectTree not available (SAP GUI < 7.70 patch 3)
            // or other error - fall back to slow method
            _logger.Debug(ex, "GetObjectTree not available or failed");
            return false;
        }
    }

    /// <summary>
    /// Parse JSON from GetObjectTree into ObjectInfo list
    /// JSON structure: {"children": [{"properties": {...}, "children": [...]}]}
    /// </summary>
    private void ParseObjectTreeJson(string json, List<ObjectInfo> objects)
    {
        try
        {
            _logger.Information($"Parsing GetObjectTree JSON (length: {json.Length} chars)");
            
            // Log JSON preview for debugging
            string preview = json.Length > 500 ? json.Substring(0, 500) + "..." : json;
            _logger.Debug($"JSON preview: {preview}");
            
            using JsonDocument doc = JsonDocument.Parse(json);
            JsonElement root = doc.RootElement;

            _logger.Information($"JSON root type: {root.ValueKind}");
            
            int beforeCount = objects.Count;
            
            // Parse the JSON tree recursively
            ParseJsonElement(root, objects, 0);
            
            int parsedCount = objects.Count - beforeCount;
            _logger.Information($"Successfully parsed {parsedCount} objects from GetObjectTree JSON");
        }
        catch (Exception ex)
        {
            _logger.Error(ex, "Error parsing GetObjectTree JSON");
            throw;
        }
    }

    /// <summary>
    /// Recursively parse JSON elements into ObjectInfo
    /// Structure: {"properties": {...}, "children": [...]} or {"children": [...]}
    /// </summary>
    private void ParseJsonElement(JsonElement element, List<ObjectInfo> objects, int depth)
    {
        // Limit recursion depth for safety
        if (depth > 20)
        {
            _logger.Warning($"Max recursion depth reached at depth {depth}");
            return;
        }

        string indent = new string(' ', depth * 2);
        _logger.Debug($"{indent}ParseJsonElement: depth={depth}, type={element.ValueKind}");

        // If it's an array, process each item
        if (element.ValueKind == JsonValueKind.Array)
        {
            int index = 0;
            foreach (JsonElement item in element.EnumerateArray())
            {
                _logger.Debug($"{indent}  Processing array item {index}");
                ParseJsonElement(item, objects, depth + 1);
                index++;
            }
            return;
        }

        // If it's an object, check for "properties" and "children"
        if (element.ValueKind == JsonValueKind.Object)
        {
            JsonElement? propertiesElement = null;
            JsonElement? childrenElement = null;

            // First pass: identify "properties" and "children" fields
            foreach (JsonProperty jsonProp in element.EnumerateObject())
            {
                _logger.Debug($"{indent}  Found property: {jsonProp.Name} (type: {jsonProp.Value.ValueKind})");
                
                if (jsonProp.Name.Equals("properties", StringComparison.OrdinalIgnoreCase) && 
                    jsonProp.Value.ValueKind == JsonValueKind.Object)
                {
                    propertiesElement = jsonProp.Value;
                    _logger.Debug($"{indent}    -> This is the properties object");
                }
                else if (jsonProp.Name.Equals("children", StringComparison.OrdinalIgnoreCase) && 
                         jsonProp.Value.ValueKind == JsonValueKind.Array)
                {
                    childrenElement = jsonProp.Value;
                    int childCount = 0;
                    foreach (var _ in jsonProp.Value.EnumerateArray()) childCount++;
                    _logger.Debug($"{indent}    -> This is the children array ({childCount} children)");
                }
            }

            // If we have a "properties" object, parse it into ObjectInfo
            if (propertiesElement.HasValue)
            {
                var objInfo = new ObjectInfo();
                var properties = new Dictionary<string, PropertyInfo>();

                foreach (JsonProperty prop in propertiesElement.Value.EnumerateObject())
                {
                    string propName = prop.Name;
                    JsonElement value = prop.Value;
                    string stringValue = value.ValueKind == JsonValueKind.String ? value.GetString() ?? "" : "";

                    // Map key properties to ObjectInfo fields
                    switch (propName)
                    {
                        case "Id":
                            objInfo.Path = stringValue;
                            break;
                        case "Type":
                            objInfo.Type = stringValue;
                            break;
                        case "SubType":
                            objInfo.SubType = stringValue;
                            break;
                        case "Name":
                            objInfo.Name = stringValue;
                            break;
                        case "Text":
                            objInfo.Text = stringValue;
                            break;
                        case "Tooltip":
                            objInfo.Label = stringValue;
                            break;
                        default:
                            // Store all properties in the Properties dictionary
                            properties[propName] = new PropertyInfo
                            {
                                Type = value.ValueKind.ToString(),
                                Value = GetJsonValue(value),
                                Readable = true,
                                Writable = propName.Equals("Changeable", StringComparison.OrdinalIgnoreCase) && 
                                          stringValue.Equals("true", StringComparison.OrdinalIgnoreCase)
                            };
                            break;
                    }
                }

                // Add object if it has a valid ID/Path
                if (!string.IsNullOrEmpty(objInfo.Path))
                {
                    objInfo.Properties = properties;
                    objects.Add(objInfo);
                    
                    // Log with SubType if present
                    string typeInfo = string.IsNullOrEmpty(objInfo.SubType) 
                        ? $"Type: {objInfo.Type}" 
                        : $"Type: {objInfo.Type}, SubType: {objInfo.SubType}";
                    _logger.Information($"{indent}✓ Added object: {objInfo.Path} ({typeInfo})");
                }
                else
                {
                    _logger.Warning($"{indent}✗ Skipped object without Path/Id - Type: {objInfo.Type}, Name: {objInfo.Name}");
                }
            }
            else
            {
                _logger.Debug($"{indent}  No 'properties' object found at this level");
            }

            // Recursively process children array if it exists
            // BUT: Skip children for Table components to avoid capturing thousands of cells
            if (childrenElement.HasValue)
            {
                bool isTable = IsTableComponent(element);
                
                if (isTable)
                {
                    _logger.Debug($"{indent}  ⚠️  Skipping children of Table component (to avoid cell explosion)");
                }
                else
                {
                    _logger.Debug($"{indent}  Processing children array...");
                    ParseJsonElement(childrenElement.Value, objects, depth + 1);
                }
            }
            else
            {
                _logger.Debug($"{indent}  No 'children' array found");
            }
        }
    }

    /// <summary>
    /// Check if a JSON element represents a Table component
    /// Tables have thousands of cell children - we should skip them
    /// </summary>
    private bool IsTableComponent(JsonElement element)
    {
        if (element.ValueKind != JsonValueKind.Object)
            return false;

        try
        {
            foreach (JsonProperty prop in element.EnumerateObject())
            {
                if (prop.Name.Equals("properties", StringComparison.OrdinalIgnoreCase) && 
                    prop.Value.ValueKind == JsonValueKind.Object)
                {
                    foreach (JsonProperty innerProp in prop.Value.EnumerateObject())
                    {
                        // Check Type property
                        if (innerProp.Name.Equals("Type", StringComparison.OrdinalIgnoreCase))
                        {
                            string? typeValue = innerProp.Value.GetString();
                            if (typeValue != null && typeValue.Contains("Table", StringComparison.OrdinalIgnoreCase))
                            {
                                return true;
                            }
                        }
                        
                        // Check SubType property
                        if (innerProp.Name.Equals("SubType", StringComparison.OrdinalIgnoreCase))
                        {
                            string? subTypeValue = innerProp.Value.GetString();
                            if (subTypeValue != null && subTypeValue.Equals("Table", StringComparison.OrdinalIgnoreCase))
                            {
                                return true;
                            }
                        }
                    }
                }
            }
        }
        catch
        {
            // If any error, assume not a table
            return false;
        }

        return false;
    }

    /// <summary>
    /// Extract value from JSON element based on its type
    /// </summary>
    private object GetJsonValue(JsonElement element)
    {
        return element.ValueKind switch
        {
            JsonValueKind.String => element.GetString() ?? "",
            JsonValueKind.Number => element.TryGetInt32(out int intVal) ? intVal : element.GetDouble(),
            JsonValueKind.True => true,
            JsonValueKind.False => false,
            JsonValueKind.Null => "",
            _ => element.ToString()
        };
    }

    private void TraverseObjects(object parent, string parentPath, List<ObjectInfo> objects, int depth = 0)
    {
        // Limit recursion depth
        if (depth > 10)
        {
            _logger.Debug($"Max recursion depth reached at {parentPath}");
            return;
        }

        try
        {
            var children = GetProperty(parent, "Children");
            if (children == null)
            {
                _logger.Debug($"No children property found for {parentPath}");
                return;
            }

            int childCount = (int)(GetProperty(children, "Count") ?? 0);
            _logger.Debug($"Traversing {childCount} children of {parentPath} (depth: {depth})");
            
            for (int i = 0; i < childCount; i++)
            {
                try
                {
                    var child = InvokeMethod(children, "ElementAt", i);
                    if (child == null)
                    {
                        _logger.Debug($"Child {i} of {parentPath} is null");
                        continue;
                    }
                    
                    string childId = GetSafeProperty(child, "Id") ?? $"{parentPath}/child[{i}]";
                    string childType = GetSafeProperty(child, "Type") ?? "Unknown";
                    string childSubType = GetSafeProperty(child, "SubType") ?? "";

                    _logger.Debug($"Found object: {childId} (Type: {childType}, SubType: {childSubType})");

                    // Introspect this object
                    var objectInfo = _introspector.IntrospectObject(child, childId);
                    objects.Add(objectInfo);

                    // Check if this is a table - if so, skip its children to avoid cell explosion
                    bool isTable = childType.Contains("Table", StringComparison.OrdinalIgnoreCase) || 
                                   childSubType.Equals("Table", StringComparison.OrdinalIgnoreCase);
                    
                    if (isTable)
                    {
                        _logger.Debug($"⚠️  Skipping children of Table: {childId} (to avoid cell explosion)");
                    }
                    // Recursively traverse if it's a container (but not a table)
                    else if (IsContainer(childType))
                    {
                        _logger.Debug($"Recursing into container: {childId}");
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

    /// <summary>
    /// Capture screenshot of SAP GUI window using Windows API
    /// </summary>
    public byte[] CaptureScreenshot(object session)
    {
        try
        {
            // Get main window from SAP session
            var mainWindow = InvokeMethod(session, "FindById", "wnd[0]");
            if (mainWindow == null)
            {
                _logger.Warning("Main window not found for screenshot");
                return Array.Empty<byte>();
            }

            // Get window handle from SAP GUI window
            var handleObj = GetProperty(mainWindow, "Handle");
            if (handleObj == null)
            {
                _logger.Warning("Window handle not available");
                return Array.Empty<byte>();
            }

            // Convert handle to IntPtr
            IntPtr windowHandle;
            if (handleObj is int intHandle)
            {
                windowHandle = new IntPtr(intHandle);
            }
            else if (handleObj is long longHandle)
            {
                windowHandle = new IntPtr(longHandle);
            }
            else
            {
                windowHandle = new IntPtr(Convert.ToInt32(handleObj));
            }

            // Capture screenshot using Windows API
            var screenshot = CaptureWindowToByteArray(windowHandle);
            
            _logger.Information($"Screenshot captured successfully, size: {screenshot.Length} bytes");
            
            return screenshot;
        }
        catch (Exception ex)
        {
            _logger.Error(ex, "Error capturing screenshot");
            return Array.Empty<byte>();
        }
    }

    /// <summary>
    /// Capture window screenshot using Windows GDI API
    /// </summary>
    private byte[] CaptureWindowToByteArray(IntPtr windowHandle)
    {
        try
        {
            // Get window rectangle
            if (!NativeMethods.GetWindowRect(windowHandle, out NativeMethods.RECT rect))
            {
                _logger.Warning("Failed to get window rectangle");
                return Array.Empty<byte>();
            }

            int width = rect.Right - rect.Left;
            int height = rect.Bottom - rect.Top;

            if (width <= 0 || height <= 0)
            {
                _logger.Warning($"Invalid window dimensions: {width}x{height}");
                return Array.Empty<byte>();
            }

            // Get device context
            IntPtr hdcSrc = NativeMethods.GetWindowDC(windowHandle);
            IntPtr hdcDest = NativeMethods.CreateCompatibleDC(hdcSrc);
            IntPtr hBitmap = NativeMethods.CreateCompatibleBitmap(hdcSrc, width, height);
            IntPtr hOld = NativeMethods.SelectObject(hdcDest, hBitmap);

            // Copy window content
            NativeMethods.BitBlt(hdcDest, 0, 0, width, height, hdcSrc, 0, 0, NativeMethods.SRCCOPY);

            // Cleanup device contexts
            NativeMethods.SelectObject(hdcDest, hOld);
            NativeMethods.DeleteDC(hdcDest);
            NativeMethods.ReleaseDC(windowHandle, hdcSrc);

            // Convert to byte array
            using var bitmap = System.Drawing.Image.FromHbitmap(hBitmap);
            using var stream = new MemoryStream();
            bitmap.Save(stream, System.Drawing.Imaging.ImageFormat.Png);
            
            // Cleanup bitmap
            NativeMethods.DeleteObject(hBitmap);

            return stream.ToArray();
        }
        catch (Exception ex)
        {
            _logger.Error(ex, "Error in CaptureWindowToByteArray");
            return Array.Empty<byte>();
        }
    }

    /// <summary>
    /// Windows API methods for screen capture
    /// </summary>
    private static class NativeMethods
    {
        public const int SRCCOPY = 0x00CC0020;

        [System.Runtime.InteropServices.StructLayout(System.Runtime.InteropServices.LayoutKind.Sequential)]
        public struct RECT
        {
            public int Left;
            public int Top;
            public int Right;
            public int Bottom;
        }

        [System.Runtime.InteropServices.DllImport("user32.dll")]
        public static extern IntPtr GetWindowDC(IntPtr hWnd);

        [System.Runtime.InteropServices.DllImport("user32.dll")]
        public static extern int ReleaseDC(IntPtr hWnd, IntPtr hDC);

        [System.Runtime.InteropServices.DllImport("user32.dll")]
        [return: System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.Bool)]
        public static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);

        [System.Runtime.InteropServices.DllImport("gdi32.dll")]
        public static extern IntPtr CreateCompatibleDC(IntPtr hdc);

        [System.Runtime.InteropServices.DllImport("gdi32.dll")]
        public static extern IntPtr CreateCompatibleBitmap(IntPtr hdc, int nWidth, int nHeight);

        [System.Runtime.InteropServices.DllImport("gdi32.dll")]
        public static extern IntPtr SelectObject(IntPtr hdc, IntPtr hgdiobj);

        [System.Runtime.InteropServices.DllImport("gdi32.dll")]
        [return: System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.Bool)]
        public static extern bool BitBlt(IntPtr hdcDest, int nXDest, int nYDest, int nWidth, int nHeight, 
            IntPtr hdcSrc, int nXSrc, int nYSrc, int dwRop);

        [System.Runtime.InteropServices.DllImport("gdi32.dll")]
        public static extern bool DeleteDC(IntPtr hdc);

        [System.Runtime.InteropServices.DllImport("gdi32.dll")]
        public static extern bool DeleteObject(IntPtr hObject);
    }
}

