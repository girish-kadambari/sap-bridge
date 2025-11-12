using System.Diagnostics;
using SapBridge.Models;
using SapBridge.Repositories;
using SapBridge.Utils;
using Serilog;

namespace SapBridge.Services.Vision;

/// <summary>
/// Main service for vision and robot actions.
/// Provides screenshot capture and coordinate-based interactions.
/// </summary>
public class VisionService : IVisionService
{
    private readonly ILogger _logger;
    private readonly ISapGuiRepository _repository;
    private readonly ScreenshotCapture _screenshotCapture;
    private readonly RobotActionExecutor _robotExecutor;

    public VisionService(
        ILogger logger,
        ISapGuiRepository repository)
    {
        _logger = logger ?? throw new ArgumentNullException(nameof(logger));
        _repository = repository ?? throw new ArgumentNullException(nameof(repository));
        
        _screenshotCapture = new ScreenshotCapture(_logger);
        _robotExecutor = new RobotActionExecutor(_logger);
    }

    /// <inheritdoc/>
    public async Task<Screenshot> CaptureScreenAsync()
    {
        await Task.CompletedTask;

        try
        {
            _logger.Information("Capturing full screen");
            return _screenshotCapture.CaptureScreen();
        }
        catch (Exception ex)
        {
            _logger.Error(ex, "Error capturing screen");
            var mapped = ComExceptionMapper.MapException(ex, "capturing screen");
            throw new InvalidOperationException(mapped.Message, ex);
        }
    }

    /// <inheritdoc/>
    public async Task<Screenshot> CaptureAreaAsync(ScreenRectangle area)
    {
        await Task.CompletedTask;

        try
        {
            _logger.Information("Capturing area: X={X}, Y={Y}, W={Width}, H={Height}", 
                area.X, area.Y, area.Width, area.Height);
            return _screenshotCapture.CaptureArea(area);
        }
        catch (Exception ex)
        {
            _logger.Error(ex, "Error capturing area");
            var mapped = ComExceptionMapper.MapException(ex, "capturing area");
            throw new InvalidOperationException(mapped.Message, ex);
        }
    }

    /// <inheritdoc/>
    public async Task<Screenshot> CaptureWindowAsync(string sessionId)
    {
        await Task.CompletedTask;

        try
        {
            _logger.Information("Capturing SAP GUI window for session {SessionId}", sessionId);

            var session = _repository.GetSession(sessionId);
            var windowHandle = GetWindowHandle(session);

            if (windowHandle == IntPtr.Zero)
            {
                _logger.Warning("Could not get window handle, capturing full screen");
                return _screenshotCapture.CaptureScreen();
            }

            return _screenshotCapture.CaptureWindow(windowHandle);
        }
        catch (Exception ex)
        {
            _logger.Error(ex, "Error capturing window");
            throw;
        }
    }

    /// <inheritdoc/>
    public async Task<RobotActionResult> MoveMouseAsync(ScreenPoint point)
    {
        await Task.CompletedTask;
        var stopwatch = Stopwatch.StartNew();

        try
        {
            _logger.Information("Moving mouse to ({X}, {Y})", point.X, point.Y);
            _robotExecutor.MoveMouse(point);
            stopwatch.Stop();
            return RobotActionResult.SuccessResult((int)stopwatch.ElapsedMilliseconds);
        }
        catch (Exception ex)
        {
            stopwatch.Stop();
            _logger.Error(ex, "Error moving mouse");
            return RobotActionResult.FailureResult(ex.Message, (int)stopwatch.ElapsedMilliseconds);
        }
    }

    /// <inheritdoc/>
    public async Task<RobotActionResult> ClickAsync(ScreenPoint point, MouseButton button = MouseButton.Left)
    {
        await Task.CompletedTask;
        var stopwatch = Stopwatch.StartNew();

        try
        {
            _logger.Information("Clicking {Button} at ({X}, {Y})", button, point.X, point.Y);
            _robotExecutor.ClickAt(point, button);
            stopwatch.Stop();
            return RobotActionResult.SuccessResult((int)stopwatch.ElapsedMilliseconds);
        }
        catch (Exception ex)
        {
            stopwatch.Stop();
            _logger.Error(ex, "Error clicking");
            return RobotActionResult.FailureResult(ex.Message, (int)stopwatch.ElapsedMilliseconds);
        }
    }

    /// <inheritdoc/>
    public async Task<RobotActionResult> DoubleClickAsync(ScreenPoint point, MouseButton button = MouseButton.Left)
    {
        await Task.CompletedTask;
        var stopwatch = Stopwatch.StartNew();

        try
        {
            _logger.Information("Double-clicking {Button} at ({X}, {Y})", button, point.X, point.Y);
            _robotExecutor.DoubleClickAt(point, button);
            stopwatch.Stop();
            return RobotActionResult.SuccessResult((int)stopwatch.ElapsedMilliseconds);
        }
        catch (Exception ex)
        {
            stopwatch.Stop();
            _logger.Error(ex, "Error double-clicking");
            return RobotActionResult.FailureResult(ex.Message, (int)stopwatch.ElapsedMilliseconds);
        }
    }

    /// <inheritdoc/>
    public async Task<RobotActionResult> RightClickAsync(ScreenPoint point)
    {
        return await ClickAsync(point, MouseButton.Right);
    }

    /// <inheritdoc/>
    public async Task<RobotActionResult> DragAsync(ScreenPoint startPoint, ScreenPoint endPoint, MouseButton button = MouseButton.Left)
    {
        await Task.CompletedTask;
        var stopwatch = Stopwatch.StartNew();

        try
        {
            _logger.Information("Dragging from ({X1}, {Y1}) to ({X2}, {Y2})", 
                startPoint.X, startPoint.Y, endPoint.X, endPoint.Y);
            _robotExecutor.Drag(startPoint, endPoint, button);
            stopwatch.Stop();
            return RobotActionResult.SuccessResult((int)stopwatch.ElapsedMilliseconds);
        }
        catch (Exception ex)
        {
            stopwatch.Stop();
            _logger.Error(ex, "Error dragging");
            return RobotActionResult.FailureResult(ex.Message, (int)stopwatch.ElapsedMilliseconds);
        }
    }

    /// <inheritdoc/>
    public async Task<RobotActionResult> TypeTextAsync(string text, int delayMs = 0)
    {
        await Task.CompletedTask;
        var stopwatch = Stopwatch.StartNew();

        try
        {
            _logger.Information("Typing text: {Length} characters, delay={Delay}ms", text.Length, delayMs);
            _robotExecutor.TypeText(text, delayMs);
            stopwatch.Stop();
            return RobotActionResult.SuccessResult((int)stopwatch.ElapsedMilliseconds);
        }
        catch (Exception ex)
        {
            stopwatch.Stop();
            _logger.Error(ex, "Error typing text");
            return RobotActionResult.FailureResult(ex.Message, (int)stopwatch.ElapsedMilliseconds);
        }
    }

    /// <inheritdoc/>
    public async Task<RobotActionResult> PressKeyAsync(SpecialKey key, KeyModifier modifiers = KeyModifier.None)
    {
        await Task.CompletedTask;
        var stopwatch = Stopwatch.StartNew();

        try
        {
            _logger.Information("Pressing key {Key} with modifiers {Modifiers}", key, modifiers);
            _robotExecutor.PressSpecialKey(key, modifiers);
            stopwatch.Stop();
            return RobotActionResult.SuccessResult((int)stopwatch.ElapsedMilliseconds);
        }
        catch (Exception ex)
        {
            stopwatch.Stop();
            _logger.Error(ex, "Error pressing key");
            return RobotActionResult.FailureResult(ex.Message, (int)stopwatch.ElapsedMilliseconds);
        }
    }

    /// <inheritdoc/>
    public async Task<RobotActionResult> PressKeyCombinationAsync(string key, KeyModifier modifiers)
    {
        await Task.CompletedTask;
        var stopwatch = Stopwatch.StartNew();

        try
        {
            _logger.Information("Pressing key combination: {Modifiers}+{Key}", modifiers, key);

            if (string.IsNullOrEmpty(key) || key.Length != 1)
            {
                throw new ArgumentException("Key must be a single character", nameof(key));
            }

            // Convert char to virtual key code
            var vkCode = (ushort)char.ToUpper(key[0]);
            _robotExecutor.PressKey(vkCode, modifiers);

            stopwatch.Stop();
            return RobotActionResult.SuccessResult((int)stopwatch.ElapsedMilliseconds);
        }
        catch (Exception ex)
        {
            stopwatch.Stop();
            _logger.Error(ex, "Error pressing key combination");
            return RobotActionResult.FailureResult(ex.Message, (int)stopwatch.ElapsedMilliseconds);
        }
    }

    /// <inheritdoc/>
    public async Task<ScreenPoint> GetMousePositionAsync()
    {
        await Task.CompletedTask;

        try
        {
            return _robotExecutor.GetMousePosition();
        }
        catch (Exception ex)
        {
            _logger.Error(ex, "Error getting mouse position");
            return new ScreenPoint(0, 0);
        }
    }

    /// <inheritdoc/>
    public async Task<ScreenRectangle> GetWindowBoundsAsync(string sessionId)
    {
        await Task.CompletedTask;

        try
        {
            var session = _repository.GetSession(sessionId);
            var windowHandle = GetWindowHandle(session);

            if (windowHandle == IntPtr.Zero)
            {
                _logger.Warning("Could not get window handle, returning screen bounds");
                return _screenshotCapture.GetScreenBounds();
            }

            return _screenshotCapture.GetWindowBounds(windowHandle);
        }
        catch (Exception ex)
        {
            _logger.Error(ex, "Error getting window bounds");
            return _screenshotCapture.GetScreenBounds();
        }
    }

    /// <inheritdoc/>
    public async Task WaitAsync(int milliseconds)
    {
        _logger.Debug("Waiting {Milliseconds}ms", milliseconds);
        await Task.Delay(milliseconds);
    }

    /// <summary>
    /// Gets the window handle from a SAP GUI session.
    /// </summary>
    private IntPtr GetWindowHandle(object session)
    {
        try
        {
            // Try to get the window handle from the session
            var mainWindow = _repository.GetObjectProperty(session, "ActiveWindow");
            if (mainWindow != null)
            {
                var handle = _repository.GetObjectProperty(mainWindow, "Handle");
                if (handle != null)
                {
                    return new IntPtr(Convert.ToInt64(handle));
                }
            }

            _logger.Debug("Could not get window handle from session");
            return IntPtr.Zero;
        }
        catch (Exception ex)
        {
            _logger.Debug(ex, "Error getting window handle");
            return IntPtr.Zero;
        }
    }
}

