namespace SapBridge.Models;

/// <summary>
/// Represents the current state of a SAP GUI screen.
/// </summary>
public class ScreenState
{
    /// <summary>
    /// Current transaction code.
    /// </summary>
    public string Transaction { get; set; } = string.Empty;

    /// <summary>
    /// Screen number.
    /// </summary>
    public string ScreenNumber { get; set; } = string.Empty;

    /// <summary>
    /// Program name.
    /// </summary>
    public string ProgramName { get; set; } = string.Empty;

    /// <summary>
    /// Title bar text.
    /// </summary>
    public string TitleBar { get; set; } = string.Empty;

    /// <summary>
    /// Status bar information.
    /// </summary>
    public StatusBarInfo StatusBar { get; set; } = new();

    /// <summary>
    /// All objects on the current screen.
    /// </summary>
    public List<ObjectInfo> Objects { get; set; } = new();

    /// <summary>
    /// Timestamp when the screen state was captured (UTC).
    /// </summary>
    public DateTime CapturedAt { get; set; } = DateTime.UtcNow;
}

/// <summary>
/// Represents information from the SAP status bar.
/// </summary>
public class StatusBarInfo
{
    /// <summary>
    /// Status bar text message.
    /// </summary>
    public string Text { get; set; } = string.Empty;

    /// <summary>
    /// Status bar message type (e.g., "info", "warning", "error", "success").
    /// </summary>
    public string Type { get; set; } = "info";

    /// <summary>
    /// All messages in the status bar.
    /// </summary>
    public List<string> Messages { get; set; } = new();

    /// <summary>
    /// Whether the status bar indicates an error.
    /// </summary>
    public bool HasError => Type.Equals("error", StringComparison.OrdinalIgnoreCase);

    /// <summary>
    /// Whether the status bar indicates a warning.
    /// </summary>
    public bool HasWarning => Type.Equals("warning", StringComparison.OrdinalIgnoreCase);
}

