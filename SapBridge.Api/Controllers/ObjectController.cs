using Microsoft.AspNetCore.Mvc;
using System.Runtime.Versioning;
using SapBridge.Core;
using SapBridge.Core.Models;

namespace SapBridge.Api.Controllers;

[ApiController]
[Route("api/object")]
[SupportedOSPlatform("windows")]
public class ObjectController : ControllerBase
{
    private readonly SapGuiConnector _connector;
    private readonly ActionExecutor _actionExecutor;
    private readonly ComIntrospector _introspector;
    private readonly Serilog.ILogger _logger;

    public ObjectController(
        SapGuiConnector connector,
        ActionExecutor actionExecutor,
        ComIntrospector introspector,
        Serilog.ILogger logger)
    {
        _connector = connector;
        _actionExecutor = actionExecutor;
        _introspector = introspector;
        _logger = logger;
    }

    [HttpGet("{sessionId}/introspect")]
    public IActionResult IntrospectObject(string sessionId, [FromQuery] string path)
    {
        try
        {
            var session = _connector.GetSession(sessionId);
            var obj = session.FindById(path);
            
            if (obj == null)
                return NotFound(new { error = $"Object not found: {path}" });

            var objectInfo = _introspector.IntrospectObject(obj, path);
            return Ok(objectInfo);
        }
        catch (Exception ex)
        {
            _logger.Error(ex, $"Error introspecting object: {path}");
            return StatusCode(500, new { error = ex.Message });
        }
    }

    [HttpPost("invoke")]
    public IActionResult InvokeMethod([FromBody] ActionRequest request)
    {
        try
        {
            var session = _connector.GetSession(request.SessionId);
            var result = _actionExecutor.ExecuteAction(session, request);
            
            if (result.Success)
                return Ok(result);
            else
                return BadRequest(result);
        }
        catch (Exception ex)
        {
            _logger.Error(ex, "Error invoking method");
            return StatusCode(500, new { error = ex.Message });
        }
    }
}

