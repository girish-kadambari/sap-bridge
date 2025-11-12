namespace SapBridge.Models;

/// <summary>
/// Represents information about a SAP GUI session.
/// </summary>
public class SessionInfo
{
    /// <summary>
    /// Unique session identifier assigned by the bridge.
    /// </summary>
    public string SessionId { get; set; } = string.Empty;

    /// <summary>
    /// SAP system name (e.g., "S23", "DEV").
    /// </summary>
    public string SystemName { get; set; } = string.Empty;

    /// <summary>
    /// SAP client number (e.g., "100", "800").
    /// </summary>
    public string Client { get; set; } = string.Empty;

    /// <summary>
    /// Logged in user name.
    /// </summary>
    public string User { get; set; } = string.Empty;

    /// <summary>
    /// Session language (e.g., "EN", "DE").
    /// </summary>
    public string Language { get; set; } = string.Empty;

    /// <summary>
    /// Transaction code currently displayed.
    /// </summary>
    public string? CurrentTransaction { get; set; }

    /// <summary>
    /// Connection timestamp (UTC).
    /// </summary>
    public DateTime ConnectedAt { get; set; }

    /// <summary>
    /// Last activity timestamp (UTC).
    /// </summary>
    public DateTime LastActivityAt { get; set; }

    /// <summary>
    /// Session state.
    /// </summary>
    public SessionState State { get; set; } = SessionState.Connected;

    /// <summary>
    /// Whether the session is currently connected.
    /// </summary>
    public bool IsConnected { get; set; } = true;
}

/// <summary>
/// Possible states of a SAP GUI session.
/// </summary>
public enum SessionState
{
    /// <summary>
    /// Session is active and connected.
    /// </summary>
    Connected,

    /// <summary>
    /// Session is disconnected.
    /// </summary>
    Disconnected,

    /// <summary>
    /// Session encountered an error.
    /// </summary>
    Error,

    /// <summary>
    /// Session has expired due to inactivity.
    /// </summary>
    Expired
}

