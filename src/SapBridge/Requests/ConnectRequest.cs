using System.ComponentModel.DataAnnotations;

namespace SapBridge.Requests;

/// <summary>
/// Request to connect to SAP GUI.
/// Supports three connection modes:
/// 1. Existing session (Server = null)
/// 2. Saved connection (Server = connection name)
/// 3. Direct connection (Server + SystemNumber + Client)
/// </summary>
public class ConnectRequest
{
    /// <summary>
    /// Server IP address (e.g., "172.21.72.22") or saved connection name.
    /// Leave null/empty to use existing active session.
    /// </summary>
    public string? Server { get; set; }

    /// <summary>
    /// SAP system number (e.g., "00").
    /// Default: "00"
    /// </summary>
    [StringLength(2, MinimumLength = 2)]
    public string SystemNumber { get; set; } = "00";

    /// <summary>
    /// SAP client number (e.g., "100", "800").
    /// Default: "100"
    /// </summary>
    [StringLength(3, MinimumLength = 3)]
    public string Client { get; set; } = "100";
}

