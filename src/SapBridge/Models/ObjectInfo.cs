namespace SapBridge.Models;

/// <summary>
/// Represents information about a SAP GUI object on screen.
/// </summary>
public class ObjectInfo
{
    /// <summary>
    /// Full path to the object (e.g., "wnd[0]/usr/txtField").
    /// </summary>
    public string Path { get; set; } = string.Empty;

    /// <summary>
    /// Object type (e.g., "GuiTextField", "GuiButton", "GuiGridView").
    /// </summary>
    public string Type { get; set; } = string.Empty;

    /// <summary>
    /// Object subtype for shell components (e.g., "Grid", "Tree", "Table").
    /// </summary>
    public string? SubType { get; set; }

    /// <summary>
    /// Object name/ID.
    /// </summary>
    public string Name { get; set; } = string.Empty;

    /// <summary>
    /// Object text content.
    /// </summary>
    public string? Text { get; set; }

    /// <summary>
    /// Object label or tooltip.
    /// </summary>
    public string? Label { get; set; }

    /// <summary>
    /// Whether the object is changeable/editable.
    /// </summary>
    public bool Changeable { get; set; }

    /// <summary>
    /// Whether the object has been modified.
    /// </summary>
    public bool Modified { get; set; }

    /// <summary>
    /// Object position and size.
    /// </summary>
    public ObjectBounds? Bounds { get; set; }

    /// <summary>
    /// List of available methods on this object.
    /// </summary>
    public List<string> Methods { get; set; } = new();

    /// <summary>
    /// Dictionary of properties with their values and metadata.
    /// </summary>
    public Dictionary<string, PropertyValue> Properties { get; set; } = new();

    /// <summary>
    /// Child object paths (for containers).
    /// </summary>
    public List<string> Children { get; set; } = new();
}

/// <summary>
/// Represents the bounds (position and size) of a SAP GUI object.
/// </summary>
public class ObjectBounds
{
    public int Left { get; set; }
    public int Top { get; set; }
    public int Width { get; set; }
    public int Height { get; set; }
    public int ScreenLeft { get; set; }
    public int ScreenTop { get; set; }
}

/// <summary>
/// Represents a property value with metadata.
/// </summary>
public class PropertyValue
{
    public string? Value { get; set; }
    public string Type { get; set; } = string.Empty;
    public bool Readable { get; set; }
    public bool Writable { get; set; }
}

