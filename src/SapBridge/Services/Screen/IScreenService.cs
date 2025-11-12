using SapBridge.Models;

namespace SapBridge.Services.Screen;

/// <summary>
/// Service interface for SAP GUI screen operations.
/// </summary>
public interface IScreenService
{
    /// <summary>
    /// Gets the current screen state.
    /// </summary>
    /// <param name="sessionId">The session ID.</param>
    /// <returns>Current screen state.</returns>
    Task<ScreenState> GetScreenStateAsync(string sessionId);

    /// <summary>
    /// Gets the current transaction code.
    /// </summary>
    /// <param name="sessionId">The session ID.</param>
    /// <returns>Transaction code.</returns>
    Task<string> GetCurrentTransactionAsync(string sessionId);

    /// <summary>
    /// Gets the title bar text.
    /// </summary>
    /// <param name="sessionId">The session ID.</param>
    /// <returns>Title bar text.</returns>
    Task<string> GetTitleBarTextAsync(string sessionId);

    /// <summary>
    /// Gets the status bar information.
    /// </summary>
    /// <param name="sessionId">The session ID.</param>
    /// <returns>Status bar information.</returns>
    Task<StatusBarInfo> GetStatusBarAsync(string sessionId);

    /// <summary>
    /// Gets all objects on the current screen.
    /// </summary>
    /// <param name="sessionId">The session ID.</param>
    /// <returns>List of objects on screen.</returns>
    Task<List<ObjectInfo>> GetScreenObjectsAsync(string sessionId);

    /// <summary>
    /// Checks if a specific object exists on the current screen.
    /// </summary>
    /// <param name="sessionId">The session ID.</param>
    /// <param name="objectId">The object ID to check.</param>
    /// <returns>True if object exists.</returns>
    Task<bool> ObjectExistsAsync(string sessionId, string objectId);

    /// <summary>
    /// Waits for a specific object to appear on screen.
    /// </summary>
    /// <param name="sessionId">The session ID.</param>
    /// <param name="objectId">The object ID to wait for.</param>
    /// <param name="timeoutMs">Timeout in milliseconds.</param>
    /// <returns>True if object appeared within timeout.</returns>
    Task<bool> WaitForObjectAsync(string sessionId, string objectId, int timeoutMs = 5000);

    /// <summary>
    /// Gets the screen number.
    /// </summary>
    /// <param name="sessionId">The session ID.</param>
    /// <returns>Screen number.</returns>
    Task<string> GetScreenNumberAsync(string sessionId);

    /// <summary>
    /// Gets the program name.
    /// </summary>
    /// <param name="sessionId">The session ID.</param>
    /// <returns>Program name.</returns>
    Task<string> GetProgramNameAsync(string sessionId);

    /// <summary>
    /// Refreshes the current screen.
    /// </summary>
    /// <param name="sessionId">The session ID.</param>
    Task RefreshScreenAsync(string sessionId);
}

