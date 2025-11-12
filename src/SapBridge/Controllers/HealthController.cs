using Microsoft.AspNetCore.Mvc;

namespace SapBridge.Controllers;

/// <summary>
/// Health check endpoint for monitoring.
/// </summary>
[ApiController]
[Route("api/[controller]")]
public class HealthController : ControllerBase
{
    /// <summary>
    /// Basic health check endpoint.
    /// </summary>
    /// <returns>Health status.</returns>
    [HttpGet]
    public IActionResult GetHealth()
    {
        return Ok(new
        {
            Status = "healthy",
            Service = "SAP Bridge",
            Version = "1.0.0",
            Timestamp = DateTime.UtcNow
        });
    }

    /// <summary>
    /// Ping endpoint for connectivity testing.
    /// </summary>
    [HttpGet("ping")]
    public IActionResult Ping()
    {
        return Ok(new { Message = "pong", Timestamp = DateTime.UtcNow });
    }
}

