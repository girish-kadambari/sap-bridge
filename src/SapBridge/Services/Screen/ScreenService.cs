using System.Text.Json;
using SapBridge.Models;
using SapBridge.Repositories;
using SapBridge.Utils;
using Serilog;
using ILogger = Serilog.ILogger;

namespace SapBridge.Services.Screen;

/// <summary>
/// Service for SAP GUI screen operations.
/// </summary>
public class ScreenService : IScreenService
{
    private readonly ILogger _logger;
    private readonly ISapGuiRepository _repository;

    public ScreenService(ILogger logger, ISapGuiRepository repository)
    {
        _logger = logger ?? throw new ArgumentNullException(nameof(logger));
        _repository = repository ?? throw new ArgumentNullException(nameof(repository));
    }

    /// <inheritdoc/>
    public async Task<ScreenState> GetScreenStateAsync(string sessionId)
    {
        await Task.CompletedTask;

        try
        {
            _logger.Debug("Getting screen state for session {SessionId}", sessionId);

            var session = _repository.GetSession(sessionId);
            var activeWindow = _repository.GetObjectProperty(session, "ActiveWindow");

            if (activeWindow == null)
            {
                throw new InvalidOperationException("No active window found");
            }

            var screenState = new ScreenState
            {
                Transaction = await GetCurrentTransactionAsync(sessionId),
                ScreenNumber = await GetScreenNumberAsync(sessionId),
                ProgramName = await GetProgramNameAsync(sessionId),
                TitleBar = await GetTitleBarTextAsync(sessionId),
                StatusBar = await GetStatusBarAsync(sessionId),
                Objects = await GetScreenObjectsAsync(sessionId),
                CapturedAt = DateTime.UtcNow
            };

            return screenState;
        }
        catch (Exception ex)
        {
            _logger.Error(ex, "Error getting screen state");
            var mapped = ComExceptionMapper.MapException(ex, "getting screen state");
            throw new InvalidOperationException(mapped.Message, ex);
        }
    }

    /// <inheritdoc/>
    public async Task<string> GetCurrentTransactionAsync(string sessionId)
    {
        await Task.CompletedTask;

        try
        {
            var session = _repository.GetSession(sessionId);
            var info = _repository.GetObjectProperty(session, "Info");
            var transaction = _repository.GetObjectProperty(info, "Transaction")?.ToString() ?? string.Empty;
            
            return transaction;
        }
        catch (Exception ex)
        {
            _logger.Debug(ex, "Could not get current transaction");
            return string.Empty;
        }
    }

    /// <inheritdoc/>
    public async Task<string> GetTitleBarTextAsync(string sessionId)
    {
        await Task.CompletedTask;

        try
        {
            var session = _repository.GetSession(sessionId);
            var activeWindow = _repository.GetObjectProperty(session, "ActiveWindow");
            var titleBar = _repository.GetObjectProperty(activeWindow, "Text")?.ToString() ?? string.Empty;
            
            return titleBar;
        }
        catch (Exception ex)
        {
            _logger.Debug(ex, "Could not get title bar text");
            return string.Empty;
        }
    }

    /// <inheritdoc/>
    public async Task<StatusBarInfo> GetStatusBarAsync(string sessionId)
    {
        await Task.CompletedTask;

        try
        {
            var session = _repository.GetSession(sessionId);
            var activeWindow = _repository.GetObjectProperty(session, "ActiveWindow");
            
            var statusBarInfo = new StatusBarInfo();

            try
            {
                // Try to get status bar text
                var statusBar = _repository.FindObjectById(activeWindow, "sbar");
                if (statusBar != null)
                {
                    statusBarInfo.Text = _repository.GetObjectProperty(statusBar, "Text")?.ToString() ?? string.Empty;
                    
                    // Try to determine message type from properties
                    var messageType = _repository.GetObjectProperty(statusBar, "MessageType")?.ToString();
                    if (!string.IsNullOrEmpty(messageType))
                    {
                        statusBarInfo.Type = messageType.ToLower();
                    }
                }
            }
            catch (Exception ex)
            {
                _logger.Debug(ex, "Could not get status bar details");
            }

            return statusBarInfo;
        }
        catch (Exception ex)
        {
            _logger.Debug(ex, "Could not get status bar");
            return new StatusBarInfo();
        }
    }

    /// <inheritdoc/>
    public async Task<List<ObjectInfo>> GetScreenObjectsAsync(string sessionId)
    {
        await Task.CompletedTask;

        try
        {
            _logger.Information("Getting screen objects for session {SessionId}", sessionId);

            var session = _repository.GetSession(sessionId);
            
            var objects = new List<ObjectInfo>();
            
            // Use GetObjectTree (SAP GUI 7.70 patch 3+)
            // This is 10-100x faster than COM traversal
            bool success = GetObjectTreeOptimized(session, objects);
            
            if (!success)
            {
                _logger.Warning("GetObjectTree not available or failed - returning empty list");
                _logger.Warning("GetObjectTree requires SAP GUI 7.70 patch 3 or higher");
            }
            
            _logger.Information("Found {Count} objects on screen", objects.Count);
            return objects;
        }
        catch (Exception ex)
        {
            _logger.Error(ex, "Error getting screen objects");
            return new List<ObjectInfo>();
        }
    }

    /// <summary>
    /// Optimized object retrieval using GetObjectTree (SAP GUI 7.70 patch 3+).
    /// Returns entire object tree in single JSON call - 10-100x faster than individual COM calls.
    /// </summary>
    private bool GetObjectTreeOptimized(object session, List<ObjectInfo> objects)
    {
        try
        {
            _logger.Information("Attempting GetObjectTree (fast path)...");

            // Define properties we want to retrieve
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
                "ScreenLeft",   // Screen position
                "ScreenTop",
                "MaxLength"     // For text fields
            };

            // Call GetObjectTree - returns JSON string
            // Empty string for ID = get full tree from wnd[0]
            object? result = _repository.InvokeObjectMethod(session, "GetObjectTree", "", properties);
            
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
            _logger.Information("GetObjectTree retrieved {Count} objects", retrievedCount);
            
            // Return false if no objects were retrieved, so we fall back to slow traversal
            return retrievedCount > 0;
        }
        catch (Exception ex)
        {
            // GetObjectTree not available (SAP GUI < 7.70 patch 3) or other error
            _logger.Debug(ex, "GetObjectTree not available or failed, will use fallback");
            return false;
        }
    }

    /// <summary>
    /// Parse JSON from GetObjectTree into ObjectInfo list.
    /// JSON structure: {"children": [{"properties": {...}, "children": [...]}]}
    /// </summary>
    private void ParseObjectTreeJson(string json, List<ObjectInfo> objects)
    {
        try
        {
            _logger.Information("Parsing GetObjectTree JSON (length: {Length} chars)", json.Length);
            
            using JsonDocument doc = JsonDocument.Parse(json);
            JsonElement root = doc.RootElement;

            _logger.Debug("JSON root type: {Type}", root.ValueKind);
            
            int beforeCount = objects.Count;
            
            // Parse the JSON tree recursively
            ParseJsonElement(root, objects);
            
            int parsedCount = objects.Count - beforeCount;
            _logger.Information("Successfully parsed {Count} objects from GetObjectTree JSON", parsedCount);
        }
        catch (Exception ex)
        {
            _logger.Error(ex, "Error parsing GetObjectTree JSON");
            throw;
        }
    }

    /// <summary>
    /// Recursively parse JSON elements into ObjectInfo.
    /// Structure: {"properties": {...}, "children": [...]} or {"children": [...]}
    /// </summary>
    private void ParseJsonElement(JsonElement element, List<ObjectInfo> objects)
    {
        // If it's an array, process each item
        if (element.ValueKind == JsonValueKind.Array)
        {
            foreach (JsonElement item in element.EnumerateArray())
            {
                ParseJsonElement(item, objects);
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
                if (jsonProp.Name.Equals("properties", StringComparison.OrdinalIgnoreCase) && 
                    jsonProp.Value.ValueKind == JsonValueKind.Object)
                {
                    propertiesElement = jsonProp.Value;
                }
                else if (jsonProp.Name.Equals("children", StringComparison.OrdinalIgnoreCase) && 
                         jsonProp.Value.ValueKind == JsonValueKind.Array)
                {
                    childrenElement = jsonProp.Value;
                }
            }

            // If we have a "properties" object, parse it into ObjectInfo
            if (propertiesElement.HasValue)
            {
                var objInfo = ExtractObjectInfoFromJson(propertiesElement.Value);
                
                // Add object if it has a valid ID/Path
                if (objInfo != null && !string.IsNullOrEmpty(objInfo.Path))
                {
                    objects.Add(objInfo);
                    _logger.Debug("Added object: {Path} (Type: {Type}, SubType: {SubType})", 
                        objInfo.Path, objInfo.Type, objInfo.SubType ?? "N/A");
                }
            }

            // Recursively process children array if it exists
            // BUT: Skip children for Table components to avoid capturing thousands of cells
            if (childrenElement.HasValue)
            {
                bool isTable = IsTableComponent(element);
                
                if (isTable)
                {
                    _logger.Debug("Skipping children of Table component (to avoid cell explosion)");
                }
                else
                {
                    ParseJsonElement(childrenElement.Value, objects);
                }
            }
        }
    }

    /// <summary>
    /// Extract ObjectInfo from a JSON properties element.
    /// </summary>
    private ObjectInfo? ExtractObjectInfoFromJson(JsonElement propertiesElement)
    {
        try
        {
            var objInfo = new ObjectInfo();
            var properties = new Dictionary<string, PropertyValue>();

            foreach (JsonProperty prop in propertiesElement.EnumerateObject())
            {
                string propName = prop.Name;
                JsonElement value = prop.Value;
                string stringValue = GetJsonStringValue(value);

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
                    case "Changeable":
                        objInfo.Changeable = stringValue.Equals("true", StringComparison.OrdinalIgnoreCase);
                        break;
                    case "Modified":
                        objInfo.Modified = stringValue.Equals("true", StringComparison.OrdinalIgnoreCase);
                        break;
                    case "ScreenLeft":
                    case "ScreenTop":
                    case "Width":
                    case "Height":
                        // Handle bounds properties
                        if (objInfo.Bounds == null)
                        {
                            objInfo.Bounds = new ObjectBounds();
                        }
                        SetBoundsProperty(objInfo.Bounds, propName, value);
                        break;
                    default:
                        // Store all other properties in the Properties dictionary
                        properties[propName] = new PropertyValue
                        {
                            Type = value.ValueKind.ToString(),
                            Value = GetJsonValue(value)?.ToString(),
                            Readable = true,
                            Writable = propName.Equals("Changeable", StringComparison.OrdinalIgnoreCase) && 
                                      stringValue.Equals("true", StringComparison.OrdinalIgnoreCase)
                        };
                        break;
                }
            }

            objInfo.Properties = properties;
            return objInfo;
        }
        catch (Exception ex)
        {
            _logger.Debug(ex, "Error extracting ObjectInfo from JSON");
            return null;
        }
    }

    /// <summary>
    /// Check if a JSON element represents a Table component.
    /// Tables have thousands of cell children - we should skip them.
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
    /// Get string value from JSON element.
    /// </summary>
    private string GetJsonStringValue(JsonElement element)
    {
        return element.ValueKind == JsonValueKind.String ? element.GetString() ?? "" : "";
    }

    /// <summary>
    /// Extract value from JSON element based on its type.
    /// </summary>
    private object? GetJsonValue(JsonElement element)
    {
        return element.ValueKind switch
        {
            JsonValueKind.String => element.GetString() ?? "",
            JsonValueKind.Number => element.TryGetInt32(out int intVal) ? intVal : element.GetDouble(),
            JsonValueKind.True => true,
            JsonValueKind.False => false,
            JsonValueKind.Null => null,
            _ => element.ToString()
        };
    }

    /// <summary>
    /// Set bounds property from JSON value.
    /// </summary>
    private void SetBoundsProperty(ObjectBounds bounds, string propertyName, JsonElement value)
    {
        if (value.TryGetInt32(out int intValue))
        {
            switch (propertyName)
            {
                case "ScreenLeft":
                    bounds.ScreenLeft = intValue;
                    break;
                case "ScreenTop":
                    bounds.ScreenTop = intValue;
                    break;
                case "Width":
                    bounds.Width = intValue;
                    break;
                case "Height":
                    bounds.Height = intValue;
                    break;
            }
        }
    }

    /// <inheritdoc/>
    public async Task<bool> ObjectExistsAsync(string sessionId, string objectId)
    {
        await Task.CompletedTask;

        try
        {
            var session = _repository.GetSession(sessionId);
            var activeWindow = _repository.GetObjectProperty(session, "ActiveWindow");
            var obj = _repository.FindObjectById(activeWindow, objectId);
            
            return obj != null;
        }
        catch (Exception ex)
        {
            _logger.Debug(ex, "Object {ObjectId} does not exist", objectId);
            return false;
        }
    }

    /// <inheritdoc/>
    public async Task<bool> WaitForObjectAsync(string sessionId, string objectId, int timeoutMs = 5000)
    {
        _logger.Debug("Waiting for object {ObjectId} with timeout {Timeout}ms", objectId, timeoutMs);

        var startTime = DateTime.UtcNow;
        var timeout = TimeSpan.FromMilliseconds(timeoutMs);

        while (DateTime.UtcNow - startTime < timeout)
        {
            if (await ObjectExistsAsync(sessionId, objectId))
            {
                _logger.Debug("Object {ObjectId} appeared", objectId);
                return true;
            }

            await Task.Delay(100); // Check every 100ms
        }

        _logger.Debug("Timeout waiting for object {ObjectId}", objectId);
        return false;
    }

    /// <inheritdoc/>
    public async Task<string> GetScreenNumberAsync(string sessionId)
    {
        await Task.CompletedTask;

        try
        {
            var session = _repository.GetSession(sessionId);
            var info = _repository.GetObjectProperty(session, "Info");
            var screenNumber = _repository.GetObjectProperty(info, "ScreenNumber")?.ToString() ?? string.Empty;
            
            return screenNumber;
        }
        catch (Exception ex)
        {
            _logger.Debug(ex, "Could not get screen number");
            return string.Empty;
        }
    }

    /// <inheritdoc/>
    public async Task<string> GetProgramNameAsync(string sessionId)
    {
        await Task.CompletedTask;

        try
        {
            var session = _repository.GetSession(sessionId);
            var info = _repository.GetObjectProperty(session, "Info");
            var programName = _repository.GetObjectProperty(info, "Program")?.ToString() ?? string.Empty;
            
            return programName;
        }
        catch (Exception ex)
        {
            _logger.Debug(ex, "Could not get program name");
            return string.Empty;
        }
    }

    /// <inheritdoc/>
    public async Task RefreshScreenAsync(string sessionId)
    {
        await Task.CompletedTask;

        try
        {
            _logger.Information("Refreshing screen for session {SessionId}", sessionId);

            var session = _repository.GetSession(sessionId);
            var activeWindow = _repository.GetObjectProperty(session, "ActiveWindow");
            
            // Send F5 or refresh command
            _repository.InvokeObjectMethod(activeWindow, "SendVKey", 5); // F5
        }
        catch (Exception ex)
        {
            _logger.Error(ex, "Error refreshing screen");
            throw;
        }
    }
}

