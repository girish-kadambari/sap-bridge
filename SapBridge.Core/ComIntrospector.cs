using System.Runtime.InteropServices;
using Serilog;
using SapBridge.Core.Models;

namespace SapBridge.Core;

public class ComIntrospector
{
    private readonly ILogger _logger;

    public ComIntrospector(ILogger logger)
    {
        _logger = logger;
    }

    public ObjectInfo IntrospectObject(dynamic comObject, string path)
    {
        var objectInfo = new ObjectInfo
        {
            Path = path
        };

        try
        {
            // Get basic properties
            objectInfo.Type = GetSafeProperty(comObject, "Type") ?? "Unknown";
            objectInfo.Name = GetSafeProperty(comObject, "Name") ?? "";
            objectInfo.Text = GetSafeProperty(comObject, "Text") ?? "";
            
            // Try to get label from various sources
            objectInfo.Label = GetSafeProperty(comObject, "Label") 
                ?? GetSafeProperty(comObject, "Tooltip")
                ?? objectInfo.Text;

            // Discover methods and properties
            DiscoverCapabilities(comObject, objectInfo);

            // Get children
            try
            {
                var children = comObject.Children;
                if (children != null)
                {
                    for (int i = 0; i < children.Count; i++)
                    {
                        var child = children.ElementAt(i);
                        string childId = GetSafeProperty(child, "Id") ?? $"{path}/child[{i}]";
                        objectInfo.Children.Add(childId);
                    }
                }
            }
            catch
            {
                // Object may not have children
            }
        }
        catch (Exception ex)
        {
            _logger.Warning(ex, $"Error introspecting object: {path}");
        }

        return objectInfo;
    }

    private void DiscoverCapabilities(dynamic comObject, ObjectInfo objectInfo)
    {
        // Common SAP GUI methods based on object type
        var commonMethods = new Dictionary<string, List<string>>
        {
            ["GuiTextField"] = new() { "SetFocus", "SetText" },
            ["GuiCTextField"] = new() { "SetFocus", "SetText" },
            ["GuiPasswordField"] = new() { "SetFocus", "SetText" },
            ["GuiButton"] = new() { "Press", "SetFocus" },
            ["GuiCheckBox"] = new() { "SetFocus", "Selected" },
            ["GuiRadioButton"] = new() { "SetFocus", "Select" },
            ["GuiComboBox"] = new() { "SetFocus", "Key", "Value" },
            ["GuiGridView"] = new() { "GetCellValue", "SetCellValue", "SelectedRows", "RowCount" },
            ["GuiTableControl"] = new() { "GetCell", "SetCell", "RowCount", "ColumnCount" },
            ["GuiTree"] = new() { "ExpandNode", "CollapseNode", "SelectNode", "DoubleClickNode" },
            ["GuiTab"] = new() { "Select" },
            ["GuiMenu"] = new() { "Select" },
            ["GuiToolbarControl"] = new() { "PressButton", "SetButtonState" },
            ["GuiStatusbar"] = new() { "GetText", "GetMessageType" },
        };

        // Add known methods for this type
        if (commonMethods.TryGetValue(objectInfo.Type, out var methods))
        {
            objectInfo.Methods.AddRange(methods);
        }

        // Common properties to check
        var propertiesToCheck = new[] 
        {
            "Text", "Name", "Id", "Type", "Tooltip", "Label",
            "MaxLength", "Modified", "Changeable", "Left", "Top",
            "Width", "Height", "Key", "Value", "ScreenLeft", "ScreenTop"
        };

        foreach (var propName in propertiesToCheck)
        {
            try
            {
                var value = GetSafeProperty(comObject, propName);
                if (value != null)
                {
                    objectInfo.Properties[propName] = new PropertyInfo
                    {
                        Type = value.GetType().Name,
                        Value = value,
                        Readable = true,
                        Writable = IsPropertyWritable(comObject, propName)
                    };
                }
            }
            catch
            {
                // Property doesn't exist
            }
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

    private bool IsPropertyWritable(dynamic obj, string propertyName)
    {
        try
        {
            var property = obj.GetType().GetProperty(propertyName);
            return property?.CanWrite ?? false;
        }
        catch
        {
            return false;
        }
    }
}

