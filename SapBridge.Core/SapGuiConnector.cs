using System.Runtime.InteropServices;
using Serilog;
using SapBridge.Core.Models;

namespace SapBridge.Core;

public class SapGuiConnector
{
    private readonly ILogger _logger;
    private dynamic? _sapGuiApp;
    private readonly Dictionary<string, dynamic> _sessions = new();

    public SapGuiConnector(ILogger logger)
    {
        _logger = logger;
    }

    public SessionInfo Connect(string server, string systemNumber, string client)
    {
        try
        {
            _logger.Information("Connecting to SAP GUI...");
            
            // Get SAP GUI Scripting Engine
            Type sapGuiType = Type.GetTypeFromProgID("SAPGUI")
                ?? throw new Exception("SAP GUI not found. Please ensure SAP GUI is installed.");
            
            dynamic sapGui = Activator.CreateInstance(sapGuiType)!;
            _sapGuiApp = sapGui.GetScriptingEngine;

            if (_sapGuiApp == null)
                throw new Exception("Could not get SAP GUI Scripting Engine");

            _logger.Information("SAP GUI Scripting Engine obtained");

            // Open connection
            dynamic connection;
            if (!string.IsNullOrEmpty(server))
            {
                string connectionString = $"/H/{server}/S/{systemNumber}";
                _logger.Information($"Opening connection: {connectionString}");
                connection = _sapGuiApp.OpenConnection(connectionString, true);
            }
            else
            {
                // Use existing connection
                if (_sapGuiApp.Connections.Count == 0)
                    throw new Exception("No SAP GUI connections available");
                    
                connection = _sapGuiApp.Connections.ElementAt(0);
            }

            // Get session
            dynamic session = connection.Children(0);
            
            string sessionId = Guid.NewGuid().ToString();
            _sessions[sessionId] = session;

            _logger.Information($"Connected to SAP - Session ID: {sessionId}");

            return new SessionInfo
            {
                SessionId = sessionId,
                SystemName = session.Info.SystemName,
                Client = session.Info.Client,
                User = session.Info.User,
                Language = session.Info.Language,
                ConnectedAt = DateTime.UtcNow
            };
        }
        catch (COMException ex)
        {
            _logger.Error(ex, "COM Exception while connecting to SAP");
            throw new Exception($"COM Error: {ex.Message}", ex);
        }
        catch (Exception ex)
        {
            _logger.Error(ex, "Error connecting to SAP");
            throw;
        }
    }

    public dynamic GetSession(string sessionId)
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
}

