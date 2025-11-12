using SapBridge.Models;
using SapBridge.Repositories;
using SapBridge.Utils;
using Serilog;
using ILogger = Serilog.ILogger;

namespace SapBridge.Services.Session;

/// <summary>
/// Service for SAP GUI session management.
/// </summary>
public class SessionService : ISessionService
{
    private readonly ILogger _logger;
    private readonly ISapGuiRepository _repository;

    public SessionService(ILogger logger, ISapGuiRepository repository)
    {
        _logger = logger ?? throw new ArgumentNullException(nameof(logger));
        _repository = repository ?? throw new ArgumentNullException(nameof(repository));
    }

    /// <inheritdoc/>
    public async Task<SessionInfo> ConnectAsync(string sessionId, string? server = null, string systemNumber = "00", string client = "100")
    {
        await Task.CompletedTask;

        try
        {
            _logger.Information("Connecting to session {SessionId} (Server: {Server}, Client: {Client})", 
                sessionId, server ?? "existing", client);

            // Get SAP GUI application
            var sapGuiApp = _repository.GetSapGuiApplication();
            
            // If server is provided, create a new connection
            if (!string.IsNullOrEmpty(server))
            {
                return await ConnectToNewServerAsync(sessionId, server, systemNumber, client, sapGuiApp);
            }
            
            // Get connections collection
            var connections = _repository.GetObjectProperty(sapGuiApp, "Connections");
            if (connections == null)
            {
                throw new InvalidOperationException("No SAP GUI connections found. Please open SAP GUI and log in.");
            }
            
            // Get connection count
            int connectionCount = Convert.ToInt32(_repository.GetObjectProperty(connections, "Count") ?? 0);
            if (connectionCount == 0)
            {
                throw new InvalidOperationException("No active SAP GUI connections. Please open SAP GUI and log in.");
            }
            
            // Get first connection (index 0)
            var connection = _repository.InvokeObjectMethod(connections, "Item", 0);
            if (connection == null)
            {
                throw new InvalidOperationException("Could not access SAP GUI connection.");
            }
            
            // Get sessions from connection
            var sessions = _repository.GetObjectProperty(connection, "Sessions");
            if (sessions == null)
            {
                throw new InvalidOperationException("No sessions found in connection.");
            }
            
            // Get session count
            int sessionCount = Convert.ToInt32(_repository.GetObjectProperty(sessions, "Count") ?? 0);
            if (sessionCount == 0)
            {
                throw new InvalidOperationException("No active sessions. Please open a session in SAP GUI.");
            }
            
            // Get the requested session (default to first session if "0" is requested)
            int sessionIndex = int.TryParse(sessionId, out int idx) ? idx : 0;
            if (sessionIndex >= sessionCount)
            {
                throw new InvalidOperationException($"Session index {sessionIndex} out of range. Available sessions: 0-{sessionCount - 1}");
            }
            
            var session = _repository.InvokeObjectMethod(sessions, "Item", sessionIndex);
            if (session == null)
            {
                throw new InvalidOperationException($"Could not access session {sessionIndex}.");
            }
            
            // Store the session for future use
            _repository.StoreSession(sessionId, session);
            
            var sessionInfo = ExtractSessionInfo(session, sessionId);

            _logger.Information("Connected to session {SessionId} successfully", sessionId);
            return sessionInfo;
        }
        catch (Exception ex)
        {
            _logger.Error(ex, "Error connecting to session {SessionId}", sessionId);
            var mapped = ComExceptionMapper.MapException(ex, $"connecting to session {sessionId}");
            throw new InvalidOperationException(mapped.Message, ex);
        }
    }

    /// <summary>
    /// Connects to a new SAP server.
    /// </summary>
    private async Task<SessionInfo> ConnectToNewServerAsync(
        string sessionId, 
        string server, 
        string systemNumber, 
        string client, 
        object sapGuiApp)
    {
        await Task.CompletedTask;
        
        _logger.Information("Opening new connection to server: {Server}", server);
        
        // Build connection string
        // Format 1: /H/server/S/sysnr (for direct connection)
        // Format 2: connection_name (for saved connections in SAP Logon)
        string connectionString;
        
        // Check if server looks like an IP or hostname (contains . or numbers)
        if (server.Contains(".") || server.Any(char.IsDigit))
        {
            // Direct connection: /H/server/S/sysnr/M/mandant
            connectionString = $"/H/{server}/S/{systemNumber}/M/{client}";
        }
        else
        {
            // Saved connection name
            connectionString = server;
        }
        
        _logger.Debug("Connection string: {ConnectionString}", connectionString);
        
        // Open connection
        var connection = _repository.InvokeObjectMethod(sapGuiApp, "OpenConnection", connectionString, true);
        
        if (connection == null)
        {
            throw new InvalidOperationException($"Failed to connect to server: {server}");
        }
        
        _logger.Debug("Connection established, getting sessions...");
        
        // Get sessions from the new connection
        var sessions = _repository.GetObjectProperty(connection, "Sessions");
        if (sessions == null)
        {
            throw new InvalidOperationException("No sessions available after connection.");
        }
        
        // Get session count
        int sessionCount = Convert.ToInt32(_repository.GetObjectProperty(sessions, "Count") ?? 0);
        if (sessionCount == 0)
        {
            throw new InvalidOperationException("Connection established but no session was created.");
        }
        
        // Get first session (newly created)
        var session = _repository.InvokeObjectMethod(sessions, "Item", 0);
        if (session == null)
        {
            throw new InvalidOperationException("Could not access the new session.");
        }
        
        // Store the session for future use
        _repository.StoreSession(sessionId, session);
        
        var sessionInfo = ExtractSessionInfo(session, sessionId);
        
        _logger.Information("Successfully connected to new server and created session {SessionId}", sessionId);
        
        return sessionInfo;
    }

    /// <inheritdoc/>
    public async Task<SessionInfo> GetSessionInfoAsync(string sessionId)
    {
        await Task.CompletedTask;

        try
        {
            _logger.Debug("Getting session info for {SessionId}", sessionId);

            var session = _repository.GetSession(sessionId);
            return ExtractSessionInfo(session, sessionId);
        }
        catch (Exception ex)
        {
            _logger.Error(ex, "Error getting session info for {SessionId}", sessionId);
            throw;
        }
    }

    /// <inheritdoc/>
    public async Task<bool> IsSessionHealthyAsync(string sessionId)
    {
        await Task.CompletedTask;

        try
        {
            var session = _repository.GetSession(sessionId);
            return _repository.IsSessionAlive(sessionId);
        }
        catch (Exception ex)
        {
            _logger.Debug(ex, "Session {SessionId} health check failed", sessionId);
            return false;
        }
    }

    /// <inheritdoc/>
    public async Task<List<SessionInfo>> GetAllSessionsAsync()
    {
        await Task.CompletedTask;

        try
        {
            _logger.Debug("Getting all sessions");

            var sessions = new List<SessionInfo>();
            // Note: This would require implementing a session tracking mechanism
            // For now, return empty list as we connect to sessions by ID
            
            return sessions;
        }
        catch (Exception ex)
        {
            _logger.Error(ex, "Error getting all sessions");
            throw;
        }
    }

    /// <inheritdoc/>
    public async Task<ObjectInfo> FindObjectByIdAsync(string sessionId, string objectId)
    {
        await Task.CompletedTask;

        try
        {
            _logger.Debug("Finding object {ObjectId} in session {SessionId}", objectId, sessionId);

            var session = _repository.GetSession(sessionId);
            var obj = _repository.FindObjectById(session, objectId);

            if (obj == null)
            {
                throw new InvalidOperationException($"Object not found: {objectId}");
            }

            var objectInfo = new ObjectInfo
            {
                Path = objectId,
                Type = _repository.GetObjectProperty(obj, "Type")?.ToString() ?? "Unknown",
                Name = _repository.GetObjectProperty(obj, "Name")?.ToString() ?? objectId,
                Text = _repository.GetObjectProperty(obj, "Text")?.ToString()
            };

            return objectInfo;
        }
        catch (Exception ex)
        {
            _logger.Error(ex, "Error finding object {ObjectId}", objectId);
            throw;
        }
    }

    /// <inheritdoc/>
    public async Task<object?> GetObjectPropertyAsync(string sessionId, string objectId, string propertyName)
    {
        await Task.CompletedTask;

        try
        {
            var session = _repository.GetSession(sessionId);
            var obj = _repository.FindObjectById(session, objectId);

            if (obj == null)
            {
                throw new InvalidOperationException($"Object not found: {objectId}");
            }

            return _repository.GetObjectProperty(obj, propertyName);
        }
        catch (Exception ex)
        {
            _logger.Error(ex, "Error getting property {Property} from object {ObjectId}", propertyName, objectId);
            throw;
        }
    }

    /// <inheritdoc/>
    public async Task SetObjectPropertyAsync(string sessionId, string objectId, string propertyName, object value)
    {
        await Task.CompletedTask;

        try
        {
            _logger.Information("Setting property {Property} on object {ObjectId} to '{Value}'", 
                propertyName, objectId, value);

            var session = _repository.GetSession(sessionId);
            var obj = _repository.FindObjectById(session, objectId);

            if (obj == null)
            {
                throw new InvalidOperationException($"Object not found: {objectId}");
            }

            _repository.SetObjectProperty(obj, propertyName, value);
        }
        catch (Exception ex)
        {
            _logger.Error(ex, "Error setting property {Property} on object {ObjectId}", propertyName, objectId);
            throw;
        }
    }

    /// <inheritdoc/>
    public async Task<object?> InvokeObjectMethodAsync(string sessionId, string objectId, string methodName, params object[] parameters)
    {
        await Task.CompletedTask;

        try
        {
            _logger.Debug("Invoking method {Method} on object {ObjectId}", methodName, objectId);

            var session = _repository.GetSession(sessionId);
            var obj = _repository.FindObjectById(session, objectId);

            if (obj == null)
            {
                throw new InvalidOperationException($"Object not found: {objectId}");
            }

            return _repository.InvokeObjectMethod(obj, methodName, parameters);
        }
        catch (Exception ex)
        {
            _logger.Error(ex, "Error invoking method {Method} on object {ObjectId}", methodName, objectId);
            throw;
        }
    }

    /// <inheritdoc/>
    public async Task StartTransactionAsync(string sessionId, string transactionCode)
    {
        await Task.CompletedTask;

        try
        {
            _logger.Information("Starting transaction {Transaction} in session {SessionId}", 
                transactionCode, sessionId);

            var session = _repository.GetSession(sessionId);
            _repository.InvokeObjectMethod(session, "StartTransaction", transactionCode);
        }
        catch (Exception ex)
        {
            _logger.Error(ex, "Error starting transaction {Transaction}", transactionCode);
            throw;
        }
    }

    /// <inheritdoc/>
    public async Task EndTransactionAsync(string sessionId)
    {
        await Task.CompletedTask;

        try
        {
            _logger.Information("Ending transaction in session {SessionId}", sessionId);

            var session = _repository.GetSession(sessionId);
            _repository.InvokeObjectMethod(session, "EndTransaction");
        }
        catch (Exception ex)
        {
            _logger.Error(ex, "Error ending transaction");
            throw;
        }
    }

    /// <inheritdoc/>
    public async Task SendVKeyAsync(string sessionId, int virtualKey)
    {
        await Task.CompletedTask;

        try
        {
            _logger.Debug("Sending virtual key {VKey} to session {SessionId}", virtualKey, sessionId);

            var session = _repository.GetSession(sessionId);
            var activeWindow = _repository.GetObjectProperty(session, "ActiveWindow");
            
            if (activeWindow != null)
            {
                _repository.InvokeObjectMethod(activeWindow, "SendVKey", virtualKey);
            }
        }
        catch (Exception ex)
        {
            _logger.Error(ex, "Error sending virtual key {VKey}", virtualKey);
            throw;
        }
    }

    /// <summary>
    /// Extracts session information from a COM session object.
    /// </summary>
    private SessionInfo ExtractSessionInfo(object session, string sessionId)
    {
        try
        {
            var info = _repository.GetObjectProperty(session, "Info");
            
            return new SessionInfo
            {
                SessionId = sessionId,
                SystemName = _repository.GetObjectProperty(info, "SystemName")?.ToString() ?? string.Empty,
                Client = _repository.GetObjectProperty(info, "Client")?.ToString() ?? string.Empty,
                User = _repository.GetObjectProperty(info, "User")?.ToString() ?? string.Empty,
                Language = _repository.GetObjectProperty(info, "Language")?.ToString() ?? string.Empty,
                CurrentTransaction = _repository.GetObjectProperty(info, "Transaction")?.ToString(),
                ConnectedAt = DateTime.UtcNow,
                LastActivityAt = DateTime.UtcNow,
                State = SessionState.Connected,
                IsConnected = true
            };
        }
        catch (Exception ex)
        {
            _logger.Warning(ex, "Could not extract full session info, returning basic info");
            
            return new SessionInfo
            {
                SessionId = sessionId,
                ConnectedAt = DateTime.UtcNow,
                LastActivityAt = DateTime.UtcNow,
                State = SessionState.Connected,
                IsConnected = true
            };
        }
    }
}

