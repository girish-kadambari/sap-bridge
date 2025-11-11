using Microsoft.AspNetCore.Mvc;
using System.Runtime.Versioning;
using SapBridge.Core;
using SapBridge.Core.Models;

namespace SapBridge.Api.Controllers;

[ApiController]
[Route("api/session")]
[SupportedOSPlatform("windows")]
public class SessionController : ControllerBase
{
    private readonly SapGuiConnector _connector;
    private readonly Serilog.ILogger _logger;

    public SessionController(SapGuiConnector connector, Serilog.ILogger logger)
    {
        _connector = connector;
        _logger = logger;
    }

    [HttpPost("connect")]
    public IActionResult Connect([FromBody] ConnectRequest request)
    {
        try
        {
            var sessionInfo = _connector.Connect(
                request.Server ?? "",
                request.SystemNumber ?? "00",
                request.Client ?? "100"
            );

            return Ok(sessionInfo);
        }
        catch (Exception ex)
        {
            _logger.Error(ex, "Error connecting to SAP");
            return StatusCode(500, new { error = ex.Message });
        }
    }

    [HttpDelete("{sessionId}")]
    public IActionResult Disconnect(string sessionId)
    {
        try
        {
            _connector.Disconnect(sessionId);
            return Ok(new { message = "Session disconnected" });
        }
        catch (Exception ex)
        {
            _logger.Error(ex, $"Error disconnecting session: {sessionId}");
            return StatusCode(500, new { error = ex.Message });
        }
    }

    [HttpGet("active")]
    public IActionResult GetActiveSessions()
    {
        var sessions = _connector.GetActiveSessions();
        return Ok(new { sessions });
    }
}

/// <summary>
/// Request to connect to SAP. All parameters are optional.
/// - No parameters: Use existing active SAP session
/// - Server only: Connect by saved connection name (e.g., "SAP Server Latest")
/// - Server + SystemNumber + Client: Connect by server details
/// </summary>
public class ConnectRequest
{
    /// <summary>
    /// Server IP (e.g., "172.21.72.22") or saved connection name (e.g., "SAP Server Latest")
    /// Leave empty to use existing active session
    /// </summary>
    public string? Server { get; set; }
    
    /// <summary>
    /// SAP system number (e.g., "00"). Default: "00"
    /// </summary>
    public string? SystemNumber { get; set; } = "00";
    
    /// <summary>
    /// SAP client (e.g., "100"). Default: "100"
    /// </summary>
    public string? Client { get; set; } = "100";
}

