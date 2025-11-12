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
    public async Task<SessionInfo> ConnectAsync(string sessionId)
    {
        await Task.CompletedTask;

        try
        {
            _logger.Information("Connecting to session {SessionId}", sessionId);

            var session = _repository.GetSession(sessionId);
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

