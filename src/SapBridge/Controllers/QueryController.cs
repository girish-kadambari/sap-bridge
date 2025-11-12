using Microsoft.AspNetCore.Mvc;
using SapBridge.Models.Query;
using SapBridge.Services.Query;
using Serilog;

namespace SapBridge.Controllers;

/// <summary>
/// Controller for unified query execution across Grid/Table/Tree objects.
/// </summary>
[ApiController]
[Route("api/[controller]")]
public class QueryController : ControllerBase
{
    private readonly ILogger _logger;
    private readonly IQueryEngine _queryEngine;

    public QueryController(ILogger logger, IQueryEngine queryEngine)
    {
        _logger = logger;
        _queryEngine = queryEngine;
    }

    /// <summary>
    /// Executes a query against a SAP GUI object.
    /// </summary>
    /// <param name="sessionId">The session ID.</param>
    /// <param name="query">The query to execute.</param>
    /// <returns>Query result with matches.</returns>
    [HttpPost("{sessionId}/execute")]
    public async Task<IActionResult> ExecuteQuery(string sessionId, [FromBody] SapQuery query)
    {
        try
        {
            _logger.Information("Executing query: Type={Type}, Action={Action}, ObjectPath={Path}", 
                query.Type, query.Action, query.ObjectPath);

            var result = await _queryEngine.ExecuteAsync(sessionId, query);

            if (!result.Success)
            {
                return BadRequest(new
                {
                    Error = result.ErrorMessage,
                    ExecutionTimeMs = result.ExecutionTimeMs
                });
            }

            return Ok(result);
        }
        catch (Exception ex)
        {
            _logger.Error(ex, "Error executing query");
            return StatusCode(500, new { Error = ex.Message });
        }
    }

    /// <summary>
    /// Finds the first matching item.
    /// </summary>
    /// <param name="sessionId">The session ID.</param>
    /// <param name="query">The query to execute.</param>
    [HttpPost("{sessionId}/find-first")]
    public async Task<IActionResult> FindFirst(string sessionId, [FromBody] SapQuery query)
    {
        try
        {
            var match = await _queryEngine.FindFirstAsync(sessionId, query);
            
            if (match == null)
            {
                return NotFound(new { Message = "No matches found" });
            }

            return Ok(match);
        }
        catch (Exception ex)
        {
            _logger.Error(ex, "Error finding first match");
            return StatusCode(500, new { Error = ex.Message });
        }
    }

    /// <summary>
    /// Finds the last matching item.
    /// </summary>
    /// <param name="sessionId">The session ID.</param>
    /// <param name="query">The query to execute.</param>
    [HttpPost("{sessionId}/find-last")]
    public async Task<IActionResult> FindLast(string sessionId, [FromBody] SapQuery query)
    {
        try
        {
            var match = await _queryEngine.FindLastAsync(sessionId, query);
            
            if (match == null)
            {
                return NotFound(new { Message = "No matches found" });
            }

            return Ok(match);
        }
        catch (Exception ex)
        {
            _logger.Error(ex, "Error finding last match");
            return StatusCode(500, new { Error = ex.Message });
        }
    }

    /// <summary>
    /// Counts matching items.
    /// </summary>
    /// <param name="sessionId">The session ID.</param>
    /// <param name="query">The query to execute.</param>
    [HttpPost("{sessionId}/count")]
    public async Task<IActionResult> Count(string sessionId, [FromBody] SapQuery query)
    {
        try
        {
            var count = await _queryEngine.CountAsync(sessionId, query);
            return Ok(new { Count = count });
        }
        catch (Exception ex)
        {
            _logger.Error(ex, "Error counting matches");
            return StatusCode(500, new { Error = ex.Message });
        }
    }
}

