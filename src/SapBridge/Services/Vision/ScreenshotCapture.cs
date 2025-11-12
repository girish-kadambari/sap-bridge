using System.Drawing;
using System.Drawing.Imaging;
using System.Runtime.InteropServices;
using SapBridge.Models;
using Serilog;
using ILogger = Serilog.ILogger;

namespace SapBridge.Services.Vision;

/// <summary>
/// Captures screenshots using Windows GDI.
/// </summary>
public class ScreenshotCapture
{
    private readonly ILogger _logger;

    public ScreenshotCapture(ILogger logger)
    {
        _logger = logger ?? throw new ArgumentNullException(nameof(logger));
    }

    /// <summary>
    /// Captures the entire primary screen.
    /// </summary>
    public Screenshot CaptureScreen()
    {
        try
        {
            var bounds = GetScreenBounds();
            return CaptureArea(bounds);
        }
        catch (Exception ex)
        {
            _logger.Error(ex, "Error capturing screen");
            throw;
        }
    }

    /// <summary>
    /// Captures a specific area of the screen.
    /// </summary>
    public Screenshot CaptureArea(ScreenRectangle area)
    {
        try
        {
            _logger.Debug("Capturing area: X={X}, Y={Y}, W={Width}, H={Height}", 
                area.X, area.Y, area.Width, area.Height);

            using var bitmap = new Bitmap(area.Width, area.Height);
            using var graphics = Graphics.FromImage(bitmap);
            
            graphics.CopyFromScreen(
                area.X, 
                area.Y, 
                0, 
                0, 
                new Size(area.Width, area.Height));

            var base64 = ConvertToBase64(bitmap);

            var screenshot = new Screenshot
            {
                ImageBase64 = base64,
                Width = area.Width,
                Height = area.Height,
                CapturedArea = area,
                CapturedAt = DateTime.UtcNow,
                Format = "PNG"
            };

            _logger.Information("Captured screenshot: {Width}x{Height}, {Size}KB", 
                screenshot.Width, screenshot.Height, base64.Length / 1024);

            return screenshot;
        }
        catch (Exception ex)
        {
            _logger.Error(ex, "Error capturing area");
            throw;
        }
    }

    /// <summary>
    /// Captures a window by its handle.
    /// </summary>
    public Screenshot CaptureWindow(IntPtr windowHandle)
    {
        try
        {
            var bounds = GetWindowBounds(windowHandle);
            return CaptureArea(bounds);
        }
        catch (Exception ex)
        {
            _logger.Error(ex, "Error capturing window");
            throw;
        }
    }

    /// <summary>
    /// Gets the bounds of the primary screen.
    /// </summary>
    public ScreenRectangle GetScreenBounds()
    {
        return new ScreenRectangle
        {
            X = 0,
            Y = 0,
            Width = GetSystemMetrics(SM_CXSCREEN),
            Height = GetSystemMetrics(SM_CYSCREEN)
        };
    }

    /// <summary>
    /// Gets the bounds of a window.
    /// </summary>
    public ScreenRectangle GetWindowBounds(IntPtr windowHandle)
    {
        try
        {
            if (GetWindowRect(windowHandle, out RECT rect))
            {
                return new ScreenRectangle
                {
                    X = rect.Left,
                    Y = rect.Top,
                    Width = rect.Right - rect.Left,
                    Height = rect.Bottom - rect.Top
                };
            }

            _logger.Warning("Could not get window bounds, using screen bounds");
            return GetScreenBounds();
        }
        catch (Exception ex)
        {
            _logger.Warning(ex, "Error getting window bounds");
            return GetScreenBounds();
        }
    }

    /// <summary>
    /// Converts a bitmap to base64-encoded PNG.
    /// </summary>
    private string ConvertToBase64(Bitmap bitmap)
    {
        using var stream = new MemoryStream();
        bitmap.Save(stream, ImageFormat.Png);
        return Convert.ToBase64String(stream.ToArray());
    }

    #region Windows API

    private const int SM_CXSCREEN = 0;
    private const int SM_CYSCREEN = 1;

    [DllImport("user32.dll")]
    private static extern int GetSystemMetrics(int nIndex);

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);

    [StructLayout(LayoutKind.Sequential)]
    private struct RECT
    {
        public int Left;
        public int Top;
        public int Right;
        public int Bottom;
    }

    #endregion
}

