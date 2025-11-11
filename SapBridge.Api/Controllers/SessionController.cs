using Microsoft.AspNetCore.Mvc;
using Serilog;
using SapBridge.Core;
using SapBridge.Core.Models;

namespace SapBridge.Api.Controllers;

[ApiController]
[Route("api/session")]
public class SessionController : ControllerBase
{
    private readonly SapGuiConnector _connector;
    private readonly ILogger _logger;

    public SessionController(SapGuiConnector connector, ILogger logger)
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

public class ConnectRequest
{
    public string? Server { get; set; }
    public string? SystemNumber { get; set; }
    public string? Client { get; set; }
}

