using Microsoft.AspNetCore.Mvc;
using System.Runtime.Versioning;
using SapBridge.Core;

namespace SapBridge.Api.Controllers;

[ApiController]
[Route("api/health")]
[SupportedOSPlatform("windows")]
public class HealthController : ControllerBase
{
    private readonly SapGuiConnector _connector;

    public HealthController(SapGuiConnector connector)
    {
        _connector = connector;
    }

    [HttpGet]
    public IActionResult GetHealth()
    {
        return Ok(new
        {
            status = "healthy",
            sapGuiVersion = "8.0",
            activeSessions = _connector.GetActiveSessions().Count,
            timestamp = DateTime.UtcNow
        });
    }
}

