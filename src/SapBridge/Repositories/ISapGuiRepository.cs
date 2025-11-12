namespace SapBridge.Repositories;

/// <summary>
/// Repository interface for SAP GUI COM interactions.
/// Abstracts all direct COM access using reflection-based helpers.
/// </summary>
public interface ISapGuiRepository
{
    /// <summary>
    /// Gets the SAP GUI scripting application object.
    /// </summary>
    /// <returns>The SAP GUI application COM object.</returns>
    /// <exception cref="InvalidOperationException">Thrown when SAP GUI is not available.</exception>
    object GetSapGuiApplication();

    /// <summary>
    /// Gets a session by its ID.
    /// </summary>
    /// <param name="sessionId">The session ID.</param>
    /// <returns>The session COM object.</returns>
    /// <exception cref="KeyNotFoundException">Thrown when session ID is not found.</exception>
    object GetSession(string sessionId);

    /// <summary>
    /// Stores a session with the given ID.
    /// </summary>
    /// <param name="sessionId">The session ID.</param>
    /// <param name="session">The session COM object.</param>
    void StoreSession(string sessionId, object session);

    /// <summary>
    /// Removes a session by its ID.
    /// </summary>
    /// <param name="sessionId">The session ID.</param>
    /// <returns>True if the session was removed, false if not found.</returns>
    bool RemoveSession(string sessionId);

    /// <summary>
    /// Gets all active session IDs.
    /// </summary>
    /// <returns>List of session IDs.</returns>
    List<string> GetAllSessionIds();

    /// <summary>
    /// Finds a SAP GUI object by its path.
    /// </summary>
    /// <param name="session">The session COM object.</param>
    /// <param name="path">The object path (e.g., "wnd[0]/usr/txtField").</param>
    /// <returns>The found object or null if not found.</returns>
    object? FindObjectById(object session, string path);

    /// <summary>
    /// Gets a property value from a COM object.
    /// </summary>
    /// <param name="obj">The COM object.</param>
    /// <param name="propertyName">The property name.</param>
    /// <returns>The property value.</returns>
    object? GetObjectProperty(object obj, string propertyName);

    /// <summary>
    /// Sets a property value on a COM object.
    /// </summary>
    /// <param name="obj">The COM object.</param>
    /// <param name="propertyName">The property name.</param>
    /// <param name="value">The value to set.</param>
    void SetObjectProperty(object obj, string propertyName, object value);

    /// <summary>
    /// Invokes a method on a COM object.
    /// </summary>
    /// <param name="obj">The COM object.</param>
    /// <param name="methodName">The method name.</param>
    /// <param name="args">The method arguments.</param>
    /// <returns>The method return value.</returns>
    object? InvokeObjectMethod(object obj, string methodName, params object[] args);

    /// <summary>
    /// Checks if a session is still alive and valid.
    /// </summary>
    /// <param name="sessionId">The session ID.</param>
    /// <returns>True if the session is alive, false otherwise.</returns>
    bool IsSessionAlive(string sessionId);
}

