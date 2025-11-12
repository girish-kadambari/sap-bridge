namespace SapBridge.Models;

/// <summary>
/// Represents a coordinate point on the screen.
/// </summary>
public class ScreenPoint
{
    /// <summary>
    /// X coordinate (horizontal position).
    /// </summary>
    public int X { get; set; }

    /// <summary>
    /// Y coordinate (vertical position).
    /// </summary>
    public int Y { get; set; }

    public ScreenPoint() { }

    public ScreenPoint(int x, int y)
    {
        X = x;
        Y = y;
    }
}

/// <summary>
/// Represents a rectangular area on the screen.
/// </summary>
public class ScreenRectangle
{
    /// <summary>
    /// X coordinate of the top-left corner.
    /// </summary>
    public int X { get; set; }

    /// <summary>
    /// Y coordinate of the top-left corner.
    /// </summary>
    public int Y { get; set; }

    /// <summary>
    /// Width of the rectangle.
    /// </summary>
    public int Width { get; set; }

    /// <summary>
    /// Height of the rectangle.
    /// </summary>
    public int Height { get; set; }

    /// <summary>
    /// Right edge coordinate (X + Width).
    /// </summary>
    public int Right => X + Width;

    /// <summary>
    /// Bottom edge coordinate (Y + Height).
    /// </summary>
    public int Bottom => Y + Height;

    /// <summary>
    /// Center point of the rectangle.
    /// </summary>
    public ScreenPoint Center => new ScreenPoint(X + Width / 2, Y + Height / 2);
}

/// <summary>
/// Represents a screenshot capture result.
/// </summary>
public class Screenshot
{
    /// <summary>
    /// Screenshot data as base64-encoded PNG.
    /// </summary>
    public string ImageBase64 { get; set; } = string.Empty;

    /// <summary>
    /// Width of the screenshot in pixels.
    /// </summary>
    public int Width { get; set; }

    /// <summary>
    /// Height of the screenshot in pixels.
    /// </summary>
    public int Height { get; set; }

    /// <summary>
    /// Area that was captured.
    /// </summary>
    public ScreenRectangle? CapturedArea { get; set; }

    /// <summary>
    /// Timestamp when screenshot was taken (UTC).
    /// </summary>
    public DateTime CapturedAt { get; set; } = DateTime.UtcNow;

    /// <summary>
    /// Image format (PNG, JPEG, etc.).
    /// </summary>
    public string Format { get; set; } = "PNG";
}

/// <summary>
/// Mouse button types.
/// </summary>
public enum MouseButton
{
    Left,
    Right,
    Middle
}

/// <summary>
/// Keyboard key modifiers.
/// </summary>
[Flags]
public enum KeyModifier
{
    None = 0,
    Shift = 1,
    Control = 2,
    Alt = 4,
    Windows = 8
}

/// <summary>
/// Robot action result.
/// </summary>
public class RobotActionResult
{
    /// <summary>
    /// Whether the action was successful.
    /// </summary>
    public bool Success { get; set; }

    /// <summary>
    /// Error message if action failed.
    /// </summary>
    public string? ErrorMessage { get; set; }

    /// <summary>
    /// Action execution time in milliseconds.
    /// </summary>
    public int ExecutionTimeMs { get; set; }

    /// <summary>
    /// Creates a success result.
    /// </summary>
    public static RobotActionResult SuccessResult(int executionTimeMs = 0)
    {
        return new RobotActionResult
        {
            Success = true,
            ExecutionTimeMs = executionTimeMs
        };
    }

    /// <summary>
    /// Creates a failure result.
    /// </summary>
    public static RobotActionResult FailureResult(string errorMessage, int executionTimeMs = 0)
    {
        return new RobotActionResult
        {
            Success = false,
            ErrorMessage = errorMessage,
            ExecutionTimeMs = executionTimeMs
        };
    }
}

/// <summary>
/// Special keys for keyboard input.
/// </summary>
public enum SpecialKey
{
    Enter,
    Tab,
    Escape,
    Backspace,
    Delete,
    Home,
    End,
    PageUp,
    PageDown,
    ArrowUp,
    ArrowDown,
    ArrowLeft,
    ArrowRight,
    F1,
    F2,
    F3,
    F4,
    F5,
    F6,
    F7,
    F8,
    F9,
    F10,
    F11,
    F12
}

