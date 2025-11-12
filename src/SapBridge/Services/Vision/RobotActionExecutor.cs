using System.Runtime.InteropServices;
using SapBridge.Models;
using Serilog;
using ILogger = Serilog.ILogger;

namespace SapBridge.Services.Vision;

/// <summary>
/// Executes robot actions (mouse and keyboard) using Windows SendInput API.
/// </summary>
public class RobotActionExecutor
{
    private readonly ILogger _logger;

    public RobotActionExecutor(ILogger logger)
    {
        _logger = logger ?? throw new ArgumentNullException(nameof(logger));
    }

    /// <summary>
    /// Moves the mouse cursor to a specific position.
    /// </summary>
    public void MoveMouse(ScreenPoint point)
    {
        try
        {
            SetCursorPos(point.X, point.Y);
            _logger.Debug("Moved mouse to ({X}, {Y})", point.X, point.Y);
        }
        catch (Exception ex)
        {
            _logger.Error(ex, "Error moving mouse");
            throw;
        }
    }

    /// <summary>
    /// Clicks the mouse at the current position.
    /// </summary>
    public void Click(MouseButton button = MouseButton.Left)
    {
        try
        {
            SendMouseEvent(button, true);  // Mouse down
            Thread.Sleep(50);
            SendMouseEvent(button, false); // Mouse up
            _logger.Debug("Clicked {Button} button", button);
        }
        catch (Exception ex)
        {
            _logger.Error(ex, "Error clicking mouse");
            throw;
        }
    }

    /// <summary>
    /// Clicks the mouse at a specific position.
    /// </summary>
    public void ClickAt(ScreenPoint point, MouseButton button = MouseButton.Left)
    {
        MoveMouse(point);
        Thread.Sleep(100); // Small delay after moving
        Click(button);
    }

    /// <summary>
    /// Double-clicks the mouse.
    /// </summary>
    public void DoubleClick(MouseButton button = MouseButton.Left)
    {
        Click(button);
        Thread.Sleep(50);
        Click(button);
        _logger.Debug("Double-clicked {Button} button", button);
    }

    /// <summary>
    /// Double-clicks at a specific position.
    /// </summary>
    public void DoubleClickAt(ScreenPoint point, MouseButton button = MouseButton.Left)
    {
        MoveMouse(point);
        Thread.Sleep(100);
        DoubleClick(button);
    }

    /// <summary>
    /// Drags from one position to another.
    /// </summary>
    public void Drag(ScreenPoint startPoint, ScreenPoint endPoint, MouseButton button = MouseButton.Left)
    {
        try
        {
            MoveMouse(startPoint);
            Thread.Sleep(100);

            SendMouseEvent(button, true); // Mouse down
            Thread.Sleep(100);

            // Move in steps for smoother drag
            int steps = 10;
            for (int i = 1; i <= steps; i++)
            {
                int x = startPoint.X + (endPoint.X - startPoint.X) * i / steps;
                int y = startPoint.Y + (endPoint.Y - startPoint.Y) * i / steps;
                MoveMouse(new ScreenPoint(x, y));
                Thread.Sleep(10);
            }

            Thread.Sleep(100);
            SendMouseEvent(button, false); // Mouse up

            _logger.Debug("Dragged from ({X1}, {Y1}) to ({X2}, {Y2})", 
                startPoint.X, startPoint.Y, endPoint.X, endPoint.Y);
        }
        catch (Exception ex)
        {
            _logger.Error(ex, "Error dragging mouse");
            throw;
        }
    }

    /// <summary>
    /// Types text character by character.
    /// </summary>
    public void TypeText(string text, int delayMs = 0)
    {
        try
        {
            foreach (char c in text)
            {
                SendKeyPress(c);
                if (delayMs > 0)
                {
                    Thread.Sleep(delayMs);
                }
            }
            _logger.Debug("Typed text: {TextLength} characters", text.Length);
        }
        catch (Exception ex)
        {
            _logger.Error(ex, "Error typing text");
            throw;
        }
    }

    /// <summary>
    /// Presses a special key.
    /// </summary>
    public void PressSpecialKey(SpecialKey key, KeyModifier modifiers = KeyModifier.None)
    {
        try
        {
            var vkCode = GetVirtualKeyCode(key);
            PressKey(vkCode, modifiers);
            _logger.Debug("Pressed {Key} with modifiers {Modifiers}", key, modifiers);
        }
        catch (Exception ex)
        {
            _logger.Error(ex, "Error pressing special key");
            throw;
        }
    }

    /// <summary>
    /// Presses a key with modifiers.
    /// </summary>
    public void PressKey(ushort vkCode, KeyModifier modifiers = KeyModifier.None)
    {
        try
        {
            // Press modifiers
            if (modifiers.HasFlag(KeyModifier.Control))
                SendKeyEvent(VK_CONTROL, true);
            if (modifiers.HasFlag(KeyModifier.Shift))
                SendKeyEvent(VK_SHIFT, true);
            if (modifiers.HasFlag(KeyModifier.Alt))
                SendKeyEvent(VK_MENU, true);
            if (modifiers.HasFlag(KeyModifier.Windows))
                SendKeyEvent(VK_LWIN, true);

            Thread.Sleep(50);

            // Press main key
            SendKeyEvent(vkCode, true);
            Thread.Sleep(50);
            SendKeyEvent(vkCode, false);

            Thread.Sleep(50);

            // Release modifiers
            if (modifiers.HasFlag(KeyModifier.Windows))
                SendKeyEvent(VK_LWIN, false);
            if (modifiers.HasFlag(KeyModifier.Alt))
                SendKeyEvent(VK_MENU, false);
            if (modifiers.HasFlag(KeyModifier.Shift))
                SendKeyEvent(VK_SHIFT, false);
            if (modifiers.HasFlag(KeyModifier.Control))
                SendKeyEvent(VK_CONTROL, false);
        }
        catch (Exception ex)
        {
            _logger.Error(ex, "Error pressing key");
            throw;
        }
    }

    /// <summary>
    /// Gets the current mouse cursor position.
    /// </summary>
    public ScreenPoint GetMousePosition()
    {
        if (GetCursorPos(out POINT point))
        {
            return new ScreenPoint(point.X, point.Y);
        }
        return new ScreenPoint(0, 0);
    }

    #region Private Methods

    private void SendMouseEvent(MouseButton button, bool down)
    {
        uint dwFlags = button switch
        {
            MouseButton.Left => down ? MOUSEEVENTF_LEFTDOWN : MOUSEEVENTF_LEFTUP,
            MouseButton.Right => down ? MOUSEEVENTF_RIGHTDOWN : MOUSEEVENTF_RIGHTUP,
            MouseButton.Middle => down ? MOUSEEVENTF_MIDDLEDOWN : MOUSEEVENTF_MIDDLEUP,
            _ => MOUSEEVENTF_LEFTDOWN
        };

        mouse_event(dwFlags, 0, 0, 0, 0);
    }

    private void SendKeyPress(char c)
    {
        keybd_event((byte)VkKeyScan(c), 0, 0, 0);
        Thread.Sleep(10);
        keybd_event((byte)VkKeyScan(c), 0, KEYEVENTF_KEYUP, 0);
    }

    private void SendKeyEvent(ushort vkCode, bool down)
    {
        uint flags = down ? 0 : KEYEVENTF_KEYUP;
        keybd_event((byte)vkCode, 0, flags, 0);
    }

    private ushort GetVirtualKeyCode(SpecialKey key)
    {
        return key switch
        {
            SpecialKey.Enter => VK_RETURN,
            SpecialKey.Tab => VK_TAB,
            SpecialKey.Escape => VK_ESCAPE,
            SpecialKey.Backspace => VK_BACK,
            SpecialKey.Delete => VK_DELETE,
            SpecialKey.Home => VK_HOME,
            SpecialKey.End => VK_END,
            SpecialKey.PageUp => VK_PRIOR,
            SpecialKey.PageDown => VK_NEXT,
            SpecialKey.ArrowUp => VK_UP,
            SpecialKey.ArrowDown => VK_DOWN,
            SpecialKey.ArrowLeft => VK_LEFT,
            SpecialKey.ArrowRight => VK_RIGHT,
            SpecialKey.F1 => VK_F1,
            SpecialKey.F2 => VK_F2,
            SpecialKey.F3 => VK_F3,
            SpecialKey.F4 => VK_F4,
            SpecialKey.F5 => VK_F5,
            SpecialKey.F6 => VK_F6,
            SpecialKey.F7 => VK_F7,
            SpecialKey.F8 => VK_F8,
            SpecialKey.F9 => VK_F9,
            SpecialKey.F10 => VK_F10,
            SpecialKey.F11 => VK_F11,
            SpecialKey.F12 => VK_F12,
            _ => 0
        };
    }

    #endregion

    #region Windows API

    [DllImport("user32.dll")]
    private static extern bool SetCursorPos(int x, int y);

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool GetCursorPos(out POINT lpPoint);

    [DllImport("user32.dll")]
    private static extern void mouse_event(uint dwFlags, int dx, int dy, uint dwData, int dwExtraInfo);

    [DllImport("user32.dll")]
    private static extern void keybd_event(byte bVk, byte bScan, uint dwFlags, int dwExtraInfo);

    [DllImport("user32.dll")]
    private static extern short VkKeyScan(char ch);

    [StructLayout(LayoutKind.Sequential)]
    private struct POINT
    {
        public int X;
        public int Y;
    }

    // Mouse event flags
    private const uint MOUSEEVENTF_LEFTDOWN = 0x0002;
    private const uint MOUSEEVENTF_LEFTUP = 0x0004;
    private const uint MOUSEEVENTF_RIGHTDOWN = 0x0008;
    private const uint MOUSEEVENTF_RIGHTUP = 0x0010;
    private const uint MOUSEEVENTF_MIDDLEDOWN = 0x0020;
    private const uint MOUSEEVENTF_MIDDLEUP = 0x0040;

    // Keyboard event flags
    private const uint KEYEVENTF_KEYUP = 0x0002;

    // Virtual key codes
    private const ushort VK_RETURN = 0x0D;
    private const ushort VK_TAB = 0x09;
    private const ushort VK_ESCAPE = 0x1B;
    private const ushort VK_BACK = 0x08;
    private const ushort VK_DELETE = 0x2E;
    private const ushort VK_HOME = 0x24;
    private const ushort VK_END = 0x23;
    private const ushort VK_PRIOR = 0x21; // Page Up
    private const ushort VK_NEXT = 0x22;  // Page Down
    private const ushort VK_UP = 0x26;
    private const ushort VK_DOWN = 0x28;
    private const ushort VK_LEFT = 0x25;
    private const ushort VK_RIGHT = 0x27;
    private const ushort VK_F1 = 0x70;
    private const ushort VK_F2 = 0x71;
    private const ushort VK_F3 = 0x72;
    private const ushort VK_F4 = 0x73;
    private const ushort VK_F5 = 0x74;
    private const ushort VK_F6 = 0x75;
    private const ushort VK_F7 = 0x76;
    private const ushort VK_F8 = 0x77;
    private const ushort VK_F9 = 0x78;
    private const ushort VK_F10 = 0x79;
    private const ushort VK_F11 = 0x7A;
    private const ushort VK_F12 = 0x7B;
    private const ushort VK_SHIFT = 0x10;
    private const ushort VK_CONTROL = 0x11;
    private const ushort VK_MENU = 0x12;    // Alt key
    private const ushort VK_LWIN = 0x5B;    // Left Windows key

    #endregion
}

