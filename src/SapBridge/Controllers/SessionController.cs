using Microsoft.AspNetCore.Mvc;
using SapBridge.Requests;
using SapBridge.Services.Session;
using Serilog;

namespace SapBridge.Controllers;

/// <summary>
/// Controller for SAP GUI session management.
/// </summary>
[ApiController]
[Route("api/[controller]")]
public class SessionController : ControllerBase
{
    private readonly ILogger _logger;
    private readonly ISessionService _sessionService;

    public SessionController(ILogger logger, ISessionService sessionService)
    {
        _logger = logger;
        _sessionService = sessionService;
    }

    /// <summary>
    /// Connects to a SAP GUI session.
    /// </summary>
    /// <param name="request">Connection parameters.</param>
    /// <returns>Session information.</returns>
    [HttpPost("connect")]
    public async Task<IActionResult> Connect([FromBody] ConnectRequest request)
    {
        try
        {
            var sessionInfo = await _sessionService.ConnectAsync(request.SessionId);
            return Ok(sessionInfo);
        }
        catch (Exception ex)
        {
            _logger.Error(ex, "Error connecting to session");
            return BadRequest(new { Error = ex.Message });
        }
    }

    /// <summary>
    /// Gets information about a session.
    /// </summary>
    /// <param name="sessionId">The session ID.</param>
    [HttpGet("{sessionId}")]
    public async Task<IActionResult> GetSessionInfo(string sessionId)
    {
        try
        {
            var sessionInfo = await _sessionService.GetSessionInfoAsync(sessionId);
            return Ok(sessionInfo);
        }
        catch (Exception ex)
        {
            _logger.Error(ex, "Error getting session info");
            return NotFound(new { Error = ex.Message });
        }
    }

    /// <summary>
    /// Checks if a session is healthy.
    /// </summary>
    /// <param name="sessionId">The session ID.</param>
    [HttpGet("{sessionId}/health")]
    public async Task<IActionResult> CheckHealth(string sessionId)
    {
        try
        {
            var isHealthy = await _sessionService.IsSessionHealthyAsync(sessionId);

            return Ok(new
            {
                SessionId = sessionId,
                IsHealthy = isHealthy,
                CheckedAt = DateTime.UtcNow
            });
        }
        catch (Exception ex)
        {
            _logger.Error(ex, "Error checking session health");
            return BadRequest(new { Error = ex.Message });
        }
    }

    /// <summary>
    /// Finds an object by ID in the session.
    /// </summary>
    /// <param name="sessionId">The session ID.</param>
    /// <param name="objectId">The object ID/path.</param>
    [HttpGet("{sessionId}/objects/{objectId}")]
    public async Task<IActionResult> FindObject(string sessionId, string objectId)
    {
        try
        {
            var objectInfo = await _sessionService.FindObjectByIdAsync(sessionId, objectId);
            return Ok(objectInfo);
        }
        catch (Exception ex)
        {
            _logger.Error(ex, "Error finding object");
            return BadRequest(new { Error = ex.Message });
        }
    }

    /// <summary>
    /// Starts a transaction.
    /// </summary>
    /// <param name="sessionId">The session ID.</param>
    /// <param name="transactionCode">Transaction code (e.g., "VA01").</param>
    [HttpPost("{sessionId}/transaction/start")]
    public async Task<IActionResult> StartTransaction(string sessionId, [FromQuery] string transactionCode)
    {
        try
        {
            await _sessionService.StartTransactionAsync(sessionId, transactionCode);
            return Ok(new { Message = $"Transaction {transactionCode} started" });
        }
        catch (Exception ex)
        {
            _logger.Error(ex, "Error starting transaction");
            return StatusCode(500, new { Error = ex.Message });
        }
    }

    /// <summary>
    /// Ends the current transaction.
    /// </summary>
    /// <param name="sessionId">The session ID.</param>
    [HttpPost("{sessionId}/transaction/end")]
    public async Task<IActionResult> EndTransaction(string sessionId)
    {
        try
        {
            await _sessionService.EndTransactionAsync(sessionId);
            return Ok(new { Message = "Transaction ended" });
        }
        catch (Exception ex)
        {
            _logger.Error(ex, "Error ending transaction");
            return StatusCode(500, new { Error = ex.Message });
        }
    }

    /// <summary>
    /// Sends a virtual key press.
    /// </summary>
    /// <param name="sessionId">The session ID.</param>
    /// <param name="virtualKey">Virtual key code (0=Enter, 3=F3, 8=F8, etc.).</param>
    [HttpPost("{sessionId}/vkey")]
    public async Task<IActionResult> SendVKey(string sessionId, [FromQuery] int virtualKey)
    {
        try
        {
            await _sessionService.SendVKeyAsync(sessionId, virtualKey);
            return Ok(new { Message = $"Sent virtual key {virtualKey}" });
        }
        catch (Exception ex)
        {
            _logger.Error(ex, "Error sending virtual key");
            return StatusCode(500, new { Error = ex.Message });
        }
    }
}

