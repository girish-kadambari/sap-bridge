using Microsoft.AspNetCore.Mvc;
using System.Runtime.Versioning;
using SapBridge.Core;

namespace SapBridge.Api.Controllers;

[ApiController]
[Route("api/screen")]
[SupportedOSPlatform("windows")]
public class ScreenController : ControllerBase
{
    private readonly SapGuiConnector _connector;
    private readonly ScreenService _screenService;
    private readonly Serilog.ILogger _logger;

    public ScreenController(
        SapGuiConnector connector,
        ScreenService screenService,
        Serilog.ILogger logger)
    {
        _connector = connector;
        _screenService = screenService;
        _logger = logger;
    }

    [HttpGet("state/{sessionId}")]
    public IActionResult GetScreenState(string sessionId)
    {
        try
        {
            var session = _connector.GetSession(sessionId);
            var screenState = _screenService.GetScreenState(session);
            screenState.SessionId = sessionId;
            
            return Ok(screenState);
        }
        catch (Exception ex)
        {
            _logger.Error(ex, $"Error getting screen state for session: {sessionId}");
            return StatusCode(500, new { error = ex.Message });
        }
    }

    [HttpPost("query")]
    public IActionResult QueryObjects([FromBody] QueryRequest request)
    {
        try
        {
            var session = _connector.GetSession(request.SessionId);
            var screenState = _screenService.GetScreenState(session);

            // Filter objects based on query
            var matchingObjects = screenState.Objects.Where(obj =>
            {
                bool matches = true;

                if (!string.IsNullOrEmpty(request.Type))
                    matches = matches && obj.Type == request.Type;

                if (!string.IsNullOrEmpty(request.Name))
                    matches = matches && (obj.Name?.Contains(request.Name, StringComparison.OrdinalIgnoreCase) ?? false);

                if (!string.IsNullOrEmpty(request.Text))
                    matches = matches && (obj.Text?.Contains(request.Text, StringComparison.OrdinalIgnoreCase) ?? false);

                return matches;
            }).ToList();

            return Ok(new { objects = matchingObjects });
        }
        catch (Exception ex)
        {
            _logger.Error(ex, "Error querying objects");
            return StatusCode(500, new { error = ex.Message });
        }
    }

    /// <summary>
    /// Capture screenshot of SAP GUI window
    /// </summary>
    [HttpGet("screenshot/{sessionId}")]
    [ProducesResponseType(typeof(FileContentResult), 200)]
    [ProducesResponseType(404)]
    [ProducesResponseType(500)]
    public IActionResult CaptureScreenshot(string sessionId)
    {
        try
        {
            _logger.Information($"Capturing screenshot for session: {sessionId}");

            var session = _connector.GetSession(sessionId);
            if (session == null)
            {
                return NotFound(new { error = $"Session not found: {sessionId}" });
            }

            var imageData = _screenService.CaptureScreenshot(session);

            if (imageData == null || imageData.Length == 0)
            {
                return NotFound(new { error = "Failed to capture screenshot" });
            }

            _logger.Information($"Screenshot captured: {imageData.Length} bytes");

            return File(imageData, "image/png", $"sap_screenshot_{sessionId}.png");
        }
        catch (Exception ex)
        {
            _logger.Error(ex, $"Error capturing screenshot for session {sessionId}");
            return StatusCode(500, new { error = ex.Message });
        }
    }
}

public class QueryRequest
{
    public string SessionId { get; set; } = string.Empty;
    public string? Type { get; set; }
    public string? Name { get; set; }
    public string? Text { get; set; }
}

