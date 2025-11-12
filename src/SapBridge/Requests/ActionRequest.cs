using System.ComponentModel.DataAnnotations;

namespace SapBridge.Requests;

/// <summary>
/// Request to execute an action on a SAP GUI object.
/// </summary>
public class ActionRequest
{
    /// <summary>
    /// Session ID.
    /// </summary>
    [Required]
    public string SessionId { get; set; } = string.Empty;

    /// <summary>
    /// Full path to the SAP GUI object (e.g., "wnd[0]/usr/txtField").
    /// </summary>
    [Required]
    public string ObjectPath { get; set; } = string.Empty;

    /// <summary>
    /// Method to invoke (e.g., "SetText", "Press", "SendVKey").
    /// </summary>
    [Required]
    public string Method { get; set; } = string.Empty;

    /// <summary>
    /// Arguments to pass to the method.
    /// </summary>
    public List<object> Args { get; set; } = new();
}

