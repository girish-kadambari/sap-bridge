using SapBridge.Models;

namespace SapBridge.Services.Vision;

/// <summary>
/// Service interface for vision and robot actions (screenshots, mouse, keyboard).
/// Enables coordinate-based interactions with SAP GUI.
/// </summary>
public interface IVisionService
{
    /// <summary>
    /// Captures a screenshot of the entire screen.
    /// </summary>
    /// <returns>Screenshot data.</returns>
    Task<Screenshot> CaptureScreenAsync();

    /// <summary>
    /// Captures a screenshot of a specific area.
    /// </summary>
    /// <param name="area">Area to capture.</param>
    /// <returns>Screenshot data.</returns>
    Task<Screenshot> CaptureAreaAsync(ScreenRectangle area);

    /// <summary>
    /// Captures a screenshot of a SAP GUI window.
    /// </summary>
    /// <param name="sessionId">The session ID.</param>
    /// <returns>Screenshot data.</returns>
    Task<Screenshot> CaptureWindowAsync(string sessionId);

    /// <summary>
    /// Moves the mouse to a specific screen position.
    /// </summary>
    /// <param name="point">Target position.</param>
    /// <returns>Action result.</returns>
    Task<RobotActionResult> MoveMouseAsync(ScreenPoint point);

    /// <summary>
    /// Clicks the mouse at a specific position.
    /// </summary>
    /// <param name="point">Position to click.</param>
    /// <param name="button">Mouse button to click.</param>
    /// <returns>Action result.</returns>
    Task<RobotActionResult> ClickAsync(ScreenPoint point, MouseButton button = MouseButton.Left);

    /// <summary>
    /// Double-clicks the mouse at a specific position.
    /// </summary>
    /// <param name="point">Position to double-click.</param>
    /// <param name="button">Mouse button to double-click.</param>
    /// <returns>Action result.</returns>
    Task<RobotActionResult> DoubleClickAsync(ScreenPoint point, MouseButton button = MouseButton.Left);

    /// <summary>
    /// Right-clicks at a specific position.
    /// </summary>
    /// <param name="point">Position to right-click.</param>
    /// <returns>Action result.</returns>
    Task<RobotActionResult> RightClickAsync(ScreenPoint point);

    /// <summary>
    /// Drags the mouse from one position to another.
    /// </summary>
    /// <param name="startPoint">Starting position.</param>
    /// <param name="endPoint">Ending position.</param>
    /// <param name="button">Mouse button to hold during drag.</param>
    /// <returns>Action result.</returns>
    Task<RobotActionResult> DragAsync(ScreenPoint startPoint, ScreenPoint endPoint, MouseButton button = MouseButton.Left);

    /// <summary>
    /// Types text using keyboard input.
    /// </summary>
    /// <param name="text">Text to type.</param>
    /// <param name="delayMs">Delay between keystrokes in milliseconds.</param>
    /// <returns>Action result.</returns>
    Task<RobotActionResult> TypeTextAsync(string text, int delayMs = 0);

    /// <summary>
    /// Presses a special key (Enter, Tab, etc.).
    /// </summary>
    /// <param name="key">Special key to press.</param>
    /// <param name="modifiers">Key modifiers (Shift, Ctrl, Alt).</param>
    /// <returns>Action result.</returns>
    Task<RobotActionResult> PressKeyAsync(SpecialKey key, KeyModifier modifiers = KeyModifier.None);

    /// <summary>
    /// Presses a key combination (e.g., Ctrl+C).
    /// </summary>
    /// <param name="key">Main key.</param>
    /// <param name="modifiers">Modifier keys.</param>
    /// <returns>Action result.</returns>
    Task<RobotActionResult> PressKeyCombinationAsync(string key, KeyModifier modifiers);

    /// <summary>
    /// Gets the current mouse cursor position.
    /// </summary>
    /// <returns>Current cursor position.</returns>
    Task<ScreenPoint> GetMousePositionAsync();

    /// <summary>
    /// Gets the screen bounds of a SAP GUI window.
    /// </summary>
    /// <param name="sessionId">The session ID.</param>
    /// <returns>Window bounds.</returns>
    Task<ScreenRectangle> GetWindowBoundsAsync(string sessionId);

    /// <summary>
    /// Waits for a specified duration (for timing between actions).
    /// </summary>
    /// <param name="milliseconds">Time to wait in milliseconds.</param>
    Task WaitAsync(int milliseconds);
}

