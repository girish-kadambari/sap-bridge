using Microsoft.AspNetCore.Mvc;
using SapBridge.Models;
using SapBridge.Services.Vision;
using Serilog;
using ILogger = Serilog.ILogger;

namespace SapBridge.Controllers;

/// <summary>
/// Controller for vision and robot actions (screenshots, mouse, keyboard).
/// </summary>
[ApiController]
[Route("api/[controller]")]
public class VisionController : ControllerBase
{
    private readonly ILogger _logger;
    private readonly IVisionService _visionService;

    public VisionController(ILogger logger, IVisionService visionService)
    {
        _logger = logger;
        _visionService = visionService;
    }

    /// <summary>
    /// Captures a screenshot of the entire screen.
    /// </summary>
    [HttpGet("screenshot/screen")]
    public async Task<IActionResult> CaptureScreen()
    {
        try
        {
            var screenshot = await _visionService.CaptureScreenAsync();
            return Ok(screenshot);
        }
        catch (Exception ex)
        {
            _logger.Error(ex, "Error capturing screen");
            return StatusCode(500, new { Error = ex.Message });
        }
    }

    /// <summary>
    /// Captures a screenshot of a SAP GUI window.
    /// </summary>
    [HttpGet("screenshot/window/{sessionId}")]
    public async Task<IActionResult> CaptureWindow(string sessionId)
    {
        try
        {
            var screenshot = await _visionService.CaptureWindowAsync(sessionId);
            return Ok(screenshot);
        }
        catch (Exception ex)
        {
            _logger.Error(ex, "Error capturing window");
            return StatusCode(500, new { Error = ex.Message });
        }
    }

    /// <summary>
    /// Captures a screenshot of a specific area.
    /// </summary>
    [HttpPost("screenshot/area")]
    public async Task<IActionResult> CaptureArea([FromBody] ScreenRectangle area)
    {
        try
        {
            var screenshot = await _visionService.CaptureAreaAsync(area);
            return Ok(screenshot);
        }
        catch (Exception ex)
        {
            _logger.Error(ex, "Error capturing area");
            return StatusCode(500, new { Error = ex.Message });
        }
    }

    /// <summary>
    /// Moves the mouse to a position.
    /// </summary>
    [HttpPost("mouse/move")]
    public async Task<IActionResult> MoveMouse([FromBody] ScreenPoint point)
    {
        try
        {
            var result = await _visionService.MoveMouseAsync(point);
            return Ok(result);
        }
        catch (Exception ex)
        {
            _logger.Error(ex, "Error moving mouse");
            return StatusCode(500, new { Error = ex.Message });
        }
    }

    /// <summary>
    /// Clicks at a position.
    /// </summary>
    [HttpPost("mouse/click")]
    public async Task<IActionResult> Click(
        [FromBody] ScreenPoint point,
        [FromQuery] MouseButton button = MouseButton.Left)
    {
        try
        {
            var result = await _visionService.ClickAsync(point, button);
            return Ok(result);
        }
        catch (Exception ex)
        {
            _logger.Error(ex, "Error clicking");
            return StatusCode(500, new { Error = ex.Message });
        }
    }

    /// <summary>
    /// Double-clicks at a position.
    /// </summary>
    [HttpPost("mouse/double-click")]
    public async Task<IActionResult> DoubleClick(
        [FromBody] ScreenPoint point,
        [FromQuery] MouseButton button = MouseButton.Left)
    {
        try
        {
            var result = await _visionService.DoubleClickAsync(point, button);
            return Ok(result);
        }
        catch (Exception ex)
        {
            _logger.Error(ex, "Error double-clicking");
            return StatusCode(500, new { Error = ex.Message });
        }
    }

    /// <summary>
    /// Right-clicks at a position.
    /// </summary>
    [HttpPost("mouse/right-click")]
    public async Task<IActionResult> RightClick([FromBody] ScreenPoint point)
    {
        try
        {
            var result = await _visionService.RightClickAsync(point);
            return Ok(result);
        }
        catch (Exception ex)
        {
            _logger.Error(ex, "Error right-clicking");
            return StatusCode(500, new { Error = ex.Message });
        }
    }

    /// <summary>
    /// Drags from one position to another.
    /// </summary>
    [HttpPost("mouse/drag")]
    public async Task<IActionResult> Drag(
        [FromQuery] int startX,
        [FromQuery] int startY,
        [FromQuery] int endX,
        [FromQuery] int endY,
        [FromQuery] MouseButton button = MouseButton.Left)
    {
        try
        {
            var startPoint = new ScreenPoint(startX, startY);
            var endPoint = new ScreenPoint(endX, endY);
            var result = await _visionService.DragAsync(startPoint, endPoint, button);
            return Ok(result);
        }
        catch (Exception ex)
        {
            _logger.Error(ex, "Error dragging");
            return StatusCode(500, new { Error = ex.Message });
        }
    }

    /// <summary>
    /// Types text.
    /// </summary>
    [HttpPost("keyboard/type")]
    public async Task<IActionResult> TypeText(
        [FromBody] string text,
        [FromQuery] int delayMs = 0)
    {
        try
        {
            var result = await _visionService.TypeTextAsync(text, delayMs);
            return Ok(result);
        }
        catch (Exception ex)
        {
            _logger.Error(ex, "Error typing text");
            return StatusCode(500, new { Error = ex.Message });
        }
    }

    /// <summary>
    /// Presses a special key.
    /// </summary>
    [HttpPost("keyboard/press-key")]
    public async Task<IActionResult> PressKey(
        [FromQuery] SpecialKey key,
        [FromQuery] KeyModifier modifiers = KeyModifier.None)
    {
        try
        {
            var result = await _visionService.PressKeyAsync(key, modifiers);
            return Ok(result);
        }
        catch (Exception ex)
        {
            _logger.Error(ex, "Error pressing key");
            return StatusCode(500, new { Error = ex.Message });
        }
    }

    /// <summary>
    /// Presses a key combination.
    /// </summary>
    [HttpPost("keyboard/key-combination")]
    public async Task<IActionResult> PressKeyCombination(
        [FromQuery] string key,
        [FromQuery] KeyModifier modifiers)
    {
        try
        {
            var result = await _visionService.PressKeyCombinationAsync(key, modifiers);
            return Ok(result);
        }
        catch (Exception ex)
        {
            _logger.Error(ex, "Error pressing key combination");
            return StatusCode(500, new { Error = ex.Message });
        }
    }

    /// <summary>
    /// Gets the current mouse position.
    /// </summary>
    [HttpGet("mouse/position")]
    public async Task<IActionResult> GetMousePosition()
    {
        try
        {
            var position = await _visionService.GetMousePositionAsync();
            return Ok(position);
        }
        catch (Exception ex)
        {
            _logger.Error(ex, "Error getting mouse position");
            return StatusCode(500, new { Error = ex.Message });
        }
    }

    /// <summary>
    /// Gets the window bounds.
    /// </summary>
    [HttpGet("window/{sessionId}/bounds")]
    public async Task<IActionResult> GetWindowBounds(string sessionId)
    {
        try
        {
            var bounds = await _visionService.GetWindowBoundsAsync(sessionId);
            return Ok(bounds);
        }
        catch (Exception ex)
        {
            _logger.Error(ex, "Error getting window bounds");
            return StatusCode(500, new { Error = ex.Message });
        }
    }

    /// <summary>
    /// Waits for a specified duration.
    /// </summary>
    [HttpPost("wait")]
    public async Task<IActionResult> Wait([FromQuery] int milliseconds)
    {
        try
        {
            await _visionService.WaitAsync(milliseconds);
            return Ok(new { Message = $"Waited {milliseconds}ms" });
        }
        catch (Exception ex)
        {
            _logger.Error(ex, "Error waiting");
            return StatusCode(500, new { Error = ex.Message });
        }
    }
}

