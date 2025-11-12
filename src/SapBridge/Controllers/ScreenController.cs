using Microsoft.AspNetCore.Mvc;
using SapBridge.Services.Screen;
using Serilog;

namespace SapBridge.Controllers;

/// <summary>
/// Controller for SAP GUI screen operations.
/// </summary>
[ApiController]
[Route("api/[controller]")]
public class ScreenController : ControllerBase
{
    private readonly ILogger _logger;
    private readonly IScreenService _screenService;

    public ScreenController(ILogger logger, IScreenService screenService)
    {
        _logger = logger;
        _screenService = screenService;
    }

    /// <summary>
    /// Gets the current screen state.
    /// </summary>
    /// <param name="sessionId">The session ID.</param>
    [HttpGet("{sessionId}/state")]
    public async Task<IActionResult> GetScreenState(string sessionId)
    {
        try
        {
            var screenState = await _screenService.GetScreenStateAsync(sessionId);
            return Ok(screenState);
        }
        catch (Exception ex)
        {
            _logger.Error(ex, "Error getting screen state");
            return StatusCode(500, new { Error = ex.Message });
        }
    }

    /// <summary>
    /// Gets the current transaction code.
    /// </summary>
    /// <param name="sessionId">The session ID.</param>
    [HttpGet("{sessionId}/transaction")]
    public async Task<IActionResult> GetCurrentTransaction(string sessionId)
    {
        try
        {
            var transaction = await _screenService.GetCurrentTransactionAsync(sessionId);
            return Ok(new { Transaction = transaction });
        }
        catch (Exception ex)
        {
            _logger.Error(ex, "Error getting current transaction");
            return StatusCode(500, new { Error = ex.Message });
        }
    }

    /// <summary>
    /// Gets the title bar text.
    /// </summary>
    /// <param name="sessionId">The session ID.</param>
    [HttpGet("{sessionId}/titlebar")]
    public async Task<IActionResult> GetTitleBar(string sessionId)
    {
        try
        {
            var titleBar = await _screenService.GetTitleBarTextAsync(sessionId);
            return Ok(new { TitleBar = titleBar });
        }
        catch (Exception ex)
        {
            _logger.Error(ex, "Error getting title bar");
            return StatusCode(500, new { Error = ex.Message });
        }
    }

    /// <summary>
    /// Gets the status bar information.
    /// </summary>
    /// <param name="sessionId">The session ID.</param>
    [HttpGet("{sessionId}/statusbar")]
    public async Task<IActionResult> GetStatusBar(string sessionId)
    {
        try
        {
            var statusBar = await _screenService.GetStatusBarAsync(sessionId);
            return Ok(statusBar);
        }
        catch (Exception ex)
        {
            _logger.Error(ex, "Error getting status bar");
            return StatusCode(500, new { Error = ex.Message });
        }
    }

    /// <summary>
    /// Gets all objects on the current screen.
    /// </summary>
    /// <param name="sessionId">The session ID.</param>
    [HttpGet("{sessionId}/objects")]
    public async Task<IActionResult> GetScreenObjects(string sessionId)
    {
        try
        {
            var objects = await _screenService.GetScreenObjectsAsync(sessionId);
            return Ok(objects);
        }
        catch (Exception ex)
        {
            _logger.Error(ex, "Error getting screen objects");
            return StatusCode(500, new { Error = ex.Message });
        }
    }

    /// <summary>
    /// Checks if a specific object exists on the current screen.
    /// </summary>
    /// <param name="sessionId">The session ID.</param>
    /// <param name="objectId">The object ID to check.</param>
    [HttpGet("{sessionId}/objects/{objectId}/exists")]
    public async Task<IActionResult> CheckObjectExists(string sessionId, string objectId)
    {
        try
        {
            var exists = await _screenService.ObjectExistsAsync(sessionId, objectId);
            return Ok(new { ObjectId = objectId, Exists = exists });
        }
        catch (Exception ex)
        {
            _logger.Error(ex, "Error checking object existence");
            return StatusCode(500, new { Error = ex.Message });
        }
    }

    /// <summary>
    /// Waits for a specific object to appear on screen.
    /// </summary>
    /// <param name="sessionId">The session ID.</param>
    /// <param name="objectId">The object ID to wait for.</param>
    /// <param name="timeoutMs">Timeout in milliseconds (default: 5000).</param>
    [HttpPost("{sessionId}/objects/{objectId}/wait")]
    public async Task<IActionResult> WaitForObject(string sessionId, string objectId, [FromQuery] int timeoutMs = 5000)
    {
        try
        {
            var appeared = await _screenService.WaitForObjectAsync(sessionId, objectId, timeoutMs);
            
            if (appeared)
            {
                return Ok(new { Message = $"Object {objectId} appeared", Appeared = true });
            }
            else
            {
                return StatusCode(408, new { Message = $"Timeout waiting for object {objectId}", Appeared = false });
            }
        }
        catch (Exception ex)
        {
            _logger.Error(ex, "Error waiting for object");
            return StatusCode(500, new { Error = ex.Message });
        }
    }

    /// <summary>
    /// Gets the screen number.
    /// </summary>
    /// <param name="sessionId">The session ID.</param>
    [HttpGet("{sessionId}/number")]
    public async Task<IActionResult> GetScreenNumber(string sessionId)
    {
        try
        {
            var screenNumber = await _screenService.GetScreenNumberAsync(sessionId);
            return Ok(new { ScreenNumber = screenNumber });
        }
        catch (Exception ex)
        {
            _logger.Error(ex, "Error getting screen number");
            return StatusCode(500, new { Error = ex.Message });
        }
    }

    /// <summary>
    /// Gets the program name.
    /// </summary>
    /// <param name="sessionId">The session ID.</param>
    [HttpGet("{sessionId}/program")]
    public async Task<IActionResult> GetProgramName(string sessionId)
    {
        try
        {
            var programName = await _screenService.GetProgramNameAsync(sessionId);
            return Ok(new { ProgramName = programName });
        }
        catch (Exception ex)
        {
            _logger.Error(ex, "Error getting program name");
            return StatusCode(500, new { Error = ex.Message });
        }
    }

    /// <summary>
    /// Refreshes the current screen.
    /// </summary>
    /// <param name="sessionId">The session ID.</param>
    [HttpPost("{sessionId}/refresh")]
    public async Task<IActionResult> RefreshScreen(string sessionId)
    {
        try
        {
            await _screenService.RefreshScreenAsync(sessionId);
            return Ok(new { Message = "Screen refreshed" });
        }
        catch (Exception ex)
        {
            _logger.Error(ex, "Error refreshing screen");
            return StatusCode(500, new { Error = ex.Message });
        }
    }
}

