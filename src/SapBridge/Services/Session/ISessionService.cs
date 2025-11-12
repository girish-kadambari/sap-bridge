using SapBridge.Models;

namespace SapBridge.Services.Session;

/// <summary>
/// Service interface for SAP GUI session management.
/// </summary>
public interface ISessionService
{
    /// <summary>
    /// Connects to a SAP GUI session.
    /// </summary>
    /// <param name="sessionId">The session ID to use.</param>
    /// <param name="server">Server address or connection name (null = use existing session).</param>
    /// <param name="systemNumber">SAP system number (e.g., "00").</param>
    /// <param name="client">SAP client number (e.g., "100").</param>
    /// <returns>Session information.</returns>
    Task<SessionInfo> ConnectAsync(string sessionId, string? server = null, string systemNumber = "00", string client = "100");

    /// <summary>
    /// Gets information about a session.
    /// </summary>
    /// <param name="sessionId">The session ID.</param>
    /// <returns>Session information.</returns>
    Task<SessionInfo> GetSessionInfoAsync(string sessionId);

    /// <summary>
    /// Checks if a session is healthy and responsive.
    /// </summary>
    /// <param name="sessionId">The session ID.</param>
    /// <returns>True if session is healthy.</returns>
    Task<bool> IsSessionHealthyAsync(string sessionId);

    /// <summary>
    /// Gets all available sessions.
    /// </summary>
    /// <returns>List of session information.</returns>
    Task<List<SessionInfo>> GetAllSessionsAsync();

    /// <summary>
    /// Finds an object by its ID in the session.
    /// </summary>
    /// <param name="sessionId">The session ID.</param>
    /// <param name="objectId">The object ID/path.</param>
    /// <returns>Object information.</returns>
    Task<ObjectInfo> FindObjectByIdAsync(string sessionId, string objectId);

    /// <summary>
    /// Gets a property value from a session object.
    /// </summary>
    /// <param name="sessionId">The session ID.</param>
    /// <param name="objectId">The object ID/path.</param>
    /// <param name="propertyName">The property name.</param>
    /// <returns>Property value.</returns>
    Task<object?> GetObjectPropertyAsync(string sessionId, string objectId, string propertyName);

    /// <summary>
    /// Sets a property value on a session object.
    /// </summary>
    /// <param name="sessionId">The session ID.</param>
    /// <param name="objectId">The object ID/path.</param>
    /// <param name="propertyName">The property name.</param>
    /// <param name="value">The value to set.</param>
    Task SetObjectPropertyAsync(string sessionId, string objectId, string propertyName, object value);

    /// <summary>
    /// Invokes a method on a session object.
    /// </summary>
    /// <param name="sessionId">The session ID.</param>
    /// <param name="objectId">The object ID/path.</param>
    /// <param name="methodName">The method name.</param>
    /// <param name="parameters">Method parameters.</param>
    /// <returns>Method return value.</returns>
    Task<object?> InvokeObjectMethodAsync(string sessionId, string objectId, string methodName, params object[] parameters);

    /// <summary>
    /// Starts a transaction.
    /// </summary>
    /// <param name="sessionId">The session ID.</param>
    /// <param name="transactionCode">Transaction code to start (e.g., "VA01").</param>
    Task StartTransactionAsync(string sessionId, string transactionCode);

    /// <summary>
    /// Ends the current transaction.
    /// </summary>
    /// <param name="sessionId">The session ID.</param>
    Task EndTransactionAsync(string sessionId);

    /// <summary>
    /// Sends a virtual key press (e.g., Enter, F3).
    /// </summary>
    /// <param name="sessionId">The session ID.</param>
    /// <param name="virtualKey">Virtual key code (e.g., 0 for Enter, 3 for F3).</param>
    Task SendVKeyAsync(string sessionId, int virtualKey);
}

