using System.Runtime.InteropServices;
using System.Runtime.Versioning;
using Serilog;
using SapBridge.Utils;
using ILogger = Serilog.ILogger;

namespace SapBridge.Repositories;

/// <summary>
/// Repository implementation for SAP GUI COM interactions.
/// Uses ComReflectionHelper for all COM operations to avoid type library dependencies.
/// </summary>
[SupportedOSPlatform("windows")]
public class SapGuiRepository : ISapGuiRepository
{
    private readonly ILogger _logger;
    private object? _sapGuiApp;
    private readonly Dictionary<string, object> _sessions = new();
    private readonly object _lock = new();

    public SapGuiRepository(ILogger logger)
    {
        _logger = logger ?? throw new ArgumentNullException(nameof(logger));
    }

    /// <inheritdoc/>
    public object GetSapGuiApplication()
    {
        if (_sapGuiApp != null)
            return _sapGuiApp;

        lock (_lock)
        {
            if (_sapGuiApp != null)
                return _sapGuiApp;

            try
            {
                _logger.Information("Binding to SAP GUI via ROT...");
                
                // Bind to SAP GUI via Running Object Table
                object sapGuiAuto = Marshal.BindToMoniker("SAPGUI");
                
                if (sapGuiAuto == null)
                {
                    throw new InvalidOperationException(
                        "SAP GUI not found. Ensure:\n" +
                        "1. SAP GUI is installed\n" +
                        "2. SAP Logon is running\n" +
                        "3. SAP GUI Scripting is enabled in Options → Accessibility & Scripting → Scripting");
                }

                // Get scripting engine
                _sapGuiApp = ComReflectionHelper.InvokeMethod(sapGuiAuto, "GetScriptingEngine");
                
                if (_sapGuiApp == null)
                {
                    throw new InvalidOperationException("Failed to get SAP GUI scripting engine.");
                }

                _logger.Information("Successfully connected to SAP GUI scripting engine.");
                return _sapGuiApp;
            }
            catch (COMException ex)
            {
                _logger.Error(ex, "COM error while connecting to SAP GUI");
                var mapped = ComExceptionMapper.MapException(ex, "connecting to SAP GUI");
                throw new InvalidOperationException(mapped.Message, ex);
            }
        }
    }

    /// <inheritdoc/>
    public object GetSession(string sessionId)
    {
        if (string.IsNullOrWhiteSpace(sessionId))
            throw new ArgumentException("Session ID cannot be null or empty.", nameof(sessionId));

        lock (_lock)
        {
            if (!_sessions.TryGetValue(sessionId, out var session))
            {
                throw new KeyNotFoundException($"Session '{sessionId}' not found.");
            }
            return session;
        }
    }

    /// <inheritdoc/>
    public void StoreSession(string sessionId, object session)
    {
        if (string.IsNullOrWhiteSpace(sessionId))
            throw new ArgumentException("Session ID cannot be null or empty.", nameof(sessionId));
        
        if (session == null)
            throw new ArgumentNullException(nameof(session));

        lock (_lock)
        {
            _sessions[sessionId] = session;
            _logger.Information("Session {SessionId} stored.", sessionId);
        }
    }

    /// <inheritdoc/>
    public bool RemoveSession(string sessionId)
    {
        if (string.IsNullOrWhiteSpace(sessionId))
            return false;

        lock (_lock)
        {
            var removed = _sessions.Remove(sessionId);
            if (removed)
            {
                _logger.Information("Session {SessionId} removed.", sessionId);
            }
            return removed;
        }
    }

    /// <inheritdoc/>
    public List<string> GetAllSessionIds()
    {
        lock (_lock)
        {
            return new List<string>(_sessions.Keys);
        }
    }

    /// <inheritdoc/>
    public object? FindObjectById(object session, string path)
    {
        if (session == null)
            throw new ArgumentNullException(nameof(session));
        
        if (string.IsNullOrWhiteSpace(path))
            throw new ArgumentException("Object path cannot be null or empty.", nameof(path));

        try
        {
            return ComReflectionHelper.InvokeMethod(session, "FindById", path);
        }
        catch (COMException ex)
        {
            _logger.Warning(ex, "Failed to find object by ID: {Path}", path);
            
            if (ComExceptionMapper.IsObjectNotFoundException(ex))
            {
                return null;
            }
            throw;
        }
    }

    /// <inheritdoc/>
    public object? GetObjectProperty(object obj, string propertyName)
    {
        if (obj == null)
            throw new ArgumentNullException(nameof(obj));

        try
        {
            return ComReflectionHelper.GetProperty(obj, propertyName);
        }
        catch (COMException ex)
        {
            _logger.Debug(ex, "Failed to get property {PropertyName}", propertyName);
            return null;
        }
    }

    /// <inheritdoc/>
    public void SetObjectProperty(object obj, string propertyName, object value)
    {
        if (obj == null)
            throw new ArgumentNullException(nameof(obj));

        try
        {
            ComReflectionHelper.SetProperty(obj, propertyName, value);
        }
        catch (COMException ex)
        {
            _logger.Error(ex, "Failed to set property {PropertyName}", propertyName);
            throw;
        }
    }

    /// <inheritdoc/>
    public object? InvokeObjectMethod(object obj, string methodName, params object[] args)
    {
        if (obj == null)
            throw new ArgumentNullException(nameof(obj));

        try
        {
            return ComReflectionHelper.InvokeMethod(obj, methodName, args);
        }
        catch (COMException ex)
        {
            _logger.Error(ex, "Failed to invoke method {MethodName}", methodName);
            throw;
        }
    }

    /// <inheritdoc/>
    public bool IsSessionAlive(string sessionId)
    {
        if (string.IsNullOrWhiteSpace(sessionId))
            return false;

        try
        {
            var session = GetSession(sessionId);
            
            // Try to access a basic property to verify the session is still valid
            var info = ComReflectionHelper.GetProperty(session, "Info");
            return info != null;
        }
        catch
        {
            return false;
        }
    }
}

