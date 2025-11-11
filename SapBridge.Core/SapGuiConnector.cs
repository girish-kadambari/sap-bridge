using System.Runtime.InteropServices;
using System.Runtime.Versioning;
using System.Reflection;
using Serilog;
using SapBridge.Core.Models;

namespace SapBridge.Core;

[SupportedOSPlatform("windows")]
public class SapGuiConnector
{
    private readonly ILogger _logger;
    private object? _sapGuiApp;
    private readonly Dictionary<string, object> _sessions = new();

    public SapGuiConnector(ILogger logger)
    {
        _logger = logger;
    }

    /// <summary>
    /// Connects to SAP GUI. Supports three scenarios:
    /// 1. Use existing active session (server=null or empty)
    /// 2. Open by connection name (server="SAP Server Latest")
    /// 3. Open by connection details (server="172.21.72.22", systemNumber="00", client="100")
    /// </summary>
    public SessionInfo Connect(string? server, string? systemNumber, string? client)
    {
        try
        {
            _logger.Information("Connecting to SAP GUI...");
            
            // Step 1: Get SAP GUI Scripting Engine from running instance
            InitializeSapGuiApp();

            // Step 2: Get or open connection
            object session;

            if (string.IsNullOrWhiteSpace(server))
            {
                // Scenario 1: Use existing active session
                _logger.Information("No server specified, using existing active session");
                session = GetActiveSession();
            }
            else
            {
                // Scenario 2 or 3: Open new connection
                object connection = OpenConnection(server, systemNumber, client);
                session = GetSessionFromConnection(connection);
            }

            // Step 3: Store session and return info
            string sessionId = Guid.NewGuid().ToString();
            _sessions[sessionId] = session;

            var sessionInfo = ExtractSessionInfo(session, sessionId);
            
            _logger.Information($"Successfully connected - Session: {sessionId}, System: {sessionInfo.SystemName}, Client: {sessionInfo.Client}");
            
            return sessionInfo;
        }
        catch (COMException ex)
        {
            _logger.Error(ex, "COM Exception while connecting to SAP");
            throw new Exception($"SAP COM Error: {ex.Message}. Ensure SAP GUI scripting is enabled.", ex);
        }
        catch (Exception ex)
        {
            _logger.Error(ex, "Error connecting to SAP");
            throw;
        }
    }

    private void InitializeSapGuiApp()
    {
        if (_sapGuiApp != null)
        {
            _logger.Information("SAP GUI App already initialized");
            return;
        }

        try
        {
            // Bind to running SAP GUI instance via ROT (Running Object Table)
            _logger.Information("Binding to SAP GUI via ROT...");
            object sapGuiAuto = Marshal.BindToMoniker("SAPGUI");
            
            if (sapGuiAuto == null)
            {
                throw new Exception(
                    "SAP GUI not found. Please ensure:\n" +
                    "1. SAP GUI is installed\n" +
                    "2. SAP Logon is running\n" +
                    "3. SAP GUI Scripting is enabled in Options → Accessibility & Scripting → Scripting"
                );
            }

            // Use reflection to call GetScriptingEngine() - avoids COM type library issues
            _sapGuiApp = InvokeMethod(sapGuiAuto, "GetScriptingEngine");
            
            if (_sapGuiApp == null)
            {
                throw new Exception(
                    "Could not get SAP GUI Scripting Engine. Please enable scripting:\n" +
                    "SAP Logon → Options → Accessibility & Scripting → Enable scripting"
                );
            }

            _logger.Information("SAP GUI Scripting Engine initialized successfully");
        }
        catch (Exception ex) when (ex.Message.Contains("0x800401E3") || ex.Message.Contains("MK_E_UNAVAILABLE"))
        {
            throw new Exception(
                "SAP GUI is not running. Please:\n" +
                "1. Start SAP Logon/SAP GUI\n" +
                "2. Ensure scripting is enabled\n" +
                "3. Try connecting again",
                ex
            );
        }
    }

    private object GetActiveSession()
    {
        object connections = GetProperty(_sapGuiApp!, "Connections");
        int connectionCount = (int)GetProperty(connections, "Count")!;
        
        if (connectionCount == 0)
        {
            throw new Exception(
                "No active SAP connections found. Please:\n" +
                "1. Open SAP Logon\n" +
                "2. Connect to an SAP system\n" +
                "3. Or provide server details to open a new connection"
            );
        }

        object connection = InvokeMethod(connections, "Item", 0)!;
        object children = GetProperty(connection, "Children")!;
        int childCount = (int)GetProperty(children, "Count")!;
        
        if (childCount == 0)
        {
            throw new Exception("No active session found in the connection");
        }

        object session = InvokeMethod(children, "Item", 0)!;
        _logger.Information("Using existing active session");
        
        return session;
    }

    private object OpenConnection(string server, string? systemNumber, string? client)
    {
        _logger.Information($"Opening connection to: {server}");

        try
        {
            // Try opening by connection name first (from SAP Logon saved connections)
            // Example: "SAP Server Latest"
            _logger.Information($"Attempting to open by connection name: {server}");
            return InvokeMethod(_sapGuiApp!, "OpenConnection", server, true)!;
        }
        catch (Exception ex1)
        {
            _logger.Information($"Connection name failed: {ex1.Message}, trying connection string...");

            try
            {
                // Build connection string based on available parameters
                string connectionString = BuildConnectionString(server, systemNumber);
                _logger.Information($"Opening connection with string: {connectionString}");
                return InvokeMethod(_sapGuiApp!, "OpenConnection", connectionString, true)!;
            }
            catch (Exception ex2)
            {
                _logger.Error(ex2, "Failed to open connection");
                throw new Exception(
                    $"Failed to connect to SAP server '{server}'. Tried:\n" +
                    $"1. Connection by name: {ex1.Message}\n" +
                    $"2. Connection by string: {ex2.Message}\n\n" +
                    $"Please ensure:\n" +
                    $"- Server is accessible\n" +
                    $"- Connection name exists in SAP Logon\n" +
                    $"- Or provide correct server IP and system number",
                    ex2
                );
            }
        }
    }

    private string BuildConnectionString(string server, string? systemNumber)
    {
        // If server already looks like a connection string, use it as-is
        if (server.StartsWith("/H/") || server.StartsWith("/M/"))
        {
            return server;
        }

        // Build connection string: /H/host/S/service or /H/host/S/instance
        string sysNum = string.IsNullOrWhiteSpace(systemNumber) ? "00" : systemNumber;
        
        // SAP uses port 32XX where XX is the system number
        // For system 00, it's 3200
        int sapPort = 3200 + int.Parse(sysNum);
        
        return $"/H/{server}/S/{sapPort}";
    }

    private object GetSessionFromConnection(object connection)
    {
        object children = GetProperty(connection, "Children")!;
        int childCount = (int)GetProperty(children, "Count")!;
        
        if (childCount == 0)
        {
            throw new Exception("Connection opened but no session available");
        }

        return InvokeMethod(children, "Item", 0)!;
    }

    private SessionInfo ExtractSessionInfo(object session, string sessionId)
    {
        try
        {
            object info = GetProperty(session, "Info")!;
            
            return new SessionInfo
            {
                SessionId = sessionId,
                SystemName = GetProperty(info, "SystemName") as string ?? "Unknown",
                Client = GetProperty(info, "Client") as string ?? "000",
                User = GetProperty(info, "User") as string ?? "Unknown",
                Language = GetProperty(info, "Language") as string ?? "EN",
                ConnectedAt = DateTime.UtcNow
            };
        }
        catch (Exception ex)
        {
            _logger.Warning(ex, "Could not extract all session info, using defaults");
            return new SessionInfo
            {
                SessionId = sessionId,
                SystemName = "SAP",
                Client = "000",
                User = "Unknown",
                Language = "EN",
                ConnectedAt = DateTime.UtcNow
            };
        }
    }

    public object GetSession(string sessionId)
    {
        if (!_sessions.TryGetValue(sessionId, out var session))
            throw new Exception($"Session not found: {sessionId}");
            
        return session;
    }

    public void Disconnect(string sessionId)
    {
        if (_sessions.TryGetValue(sessionId, out var session))
        {
            try
            {
                var connection = session.Parent;
                connection.CloseSession(session.Id);
                _sessions.Remove(sessionId);
                _logger.Information($"Disconnected session: {sessionId}");
            }
            catch (Exception ex)
            {
                _logger.Error(ex, $"Error disconnecting session: {sessionId}");
            }
        }
    }

    public List<string> GetActiveSessions()
    {
        return _sessions.Keys.ToList();
    }

    // Reflection helper methods to avoid COM type library issues
    private static object? InvokeMethod(object obj, string methodName, params object[] parameters)
    {
        Type type = obj.GetType();
        return type.InvokeMember(
            methodName,
            BindingFlags.InvokeMethod,
            null,
            obj,
            parameters
        );
    }

    private static object? GetProperty(object obj, string propertyName)
    {
        Type type = obj.GetType();
        return type.InvokeMember(
            propertyName,
            BindingFlags.GetProperty,
            null,
            obj,
            null
        );
    }
}

