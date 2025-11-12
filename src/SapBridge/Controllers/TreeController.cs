using Microsoft.AspNetCore.Mvc;
using SapBridge.Models.Query;
using SapBridge.Services.Tree;
using Serilog;
using ILogger = Serilog.ILogger;

namespace SapBridge.Controllers;

/// <summary>
/// Controller for SAP GUI Tree operations.
/// </summary>
[ApiController]
[Route("api/[controller]")]
public class TreeController : ControllerBase
{
    private readonly ILogger _logger;
    private readonly ITreeService _treeService;

    public TreeController(ILogger logger, ITreeService treeService)
    {
        _logger = logger;
        _treeService = treeService;
    }

    /// <summary>
    /// Gets the complete tree structure.
    /// </summary>
    [HttpGet("{sessionId}/data")]
    public async Task<IActionResult> GetTreeData(
        string sessionId,
        [FromQuery] string treePath,
        [FromQuery] bool expandAll = false)
    {
        try
        {
            var data = await _treeService.GetTreeDataAsync(sessionId, treePath, expandAll);
            return Ok(data);
        }
        catch (Exception ex)
        {
            _logger.Error(ex, "Error getting tree data");
            return StatusCode(500, new { Error = ex.Message });
        }
    }

    /// <summary>
    /// Gets a specific node.
    /// </summary>
    [HttpGet("{sessionId}/nodes/{nodeKey}")]
    public async Task<IActionResult> GetNode(
        string sessionId,
        string nodeKey,
        [FromQuery] string treePath)
    {
        try
        {
            var node = await _treeService.GetNodeAsync(sessionId, treePath, nodeKey);
            return Ok(node);
        }
        catch (Exception ex)
        {
            _logger.Error(ex, "Error getting tree node");
            return StatusCode(500, new { Error = ex.Message });
        }
    }

    /// <summary>
    /// Gets root nodes.
    /// </summary>
    [HttpGet("{sessionId}/roots")]
    public async Task<IActionResult> GetRootNodes(
        string sessionId,
        [FromQuery] string treePath)
    {
        try
        {
            var roots = await _treeService.GetRootNodesAsync(sessionId, treePath);
            return Ok(roots);
        }
        catch (Exception ex)
        {
            _logger.Error(ex, "Error getting root nodes");
            return StatusCode(500, new { Error = ex.Message });
        }
    }

    /// <summary>
    /// Gets children of a node.
    /// </summary>
    [HttpGet("{sessionId}/nodes/{parentKey}/children")]
    public async Task<IActionResult> GetChildren(
        string sessionId,
        string parentKey,
        [FromQuery] string treePath)
    {
        try
        {
            var children = await _treeService.GetChildrenAsync(sessionId, treePath, parentKey);
            return Ok(children);
        }
        catch (Exception ex)
        {
            _logger.Error(ex, "Error getting node children");
            return StatusCode(500, new { Error = ex.Message });
        }
    }

    /// <summary>
    /// Expands a node.
    /// </summary>
    [HttpPost("{sessionId}/nodes/{nodeKey}/expand")]
    public async Task<IActionResult> ExpandNode(
        string sessionId,
        string nodeKey,
        [FromQuery] string treePath)
    {
        try
        {
            await _treeService.ExpandNodeAsync(sessionId, treePath, nodeKey);
            return Ok(new { Message = "Node expanded successfully" });
        }
        catch (Exception ex)
        {
            _logger.Error(ex, "Error expanding node");
            return StatusCode(500, new { Error = ex.Message });
        }
    }

    /// <summary>
    /// Collapses a node.
    /// </summary>
    [HttpPost("{sessionId}/nodes/{nodeKey}/collapse")]
    public async Task<IActionResult> CollapseNode(
        string sessionId,
        string nodeKey,
        [FromQuery] string treePath)
    {
        try
        {
            await _treeService.CollapseNodeAsync(sessionId, treePath, nodeKey);
            return Ok(new { Message = "Node collapsed successfully" });
        }
        catch (Exception ex)
        {
            _logger.Error(ex, "Error collapsing node");
            return StatusCode(500, new { Error = ex.Message });
        }
    }

    /// <summary>
    /// Selects a node.
    /// </summary>
    [HttpPost("{sessionId}/nodes/{nodeKey}/select")]
    public async Task<IActionResult> SelectNode(
        string sessionId,
        string nodeKey,
        [FromQuery] string treePath)
    {
        try
        {
            await _treeService.SelectNodeAsync(sessionId, treePath, nodeKey);
            return Ok(new { Message = "Node selected successfully" });
        }
        catch (Exception ex)
        {
            _logger.Error(ex, "Error selecting node");
            return StatusCode(500, new { Error = ex.Message });
        }
    }

    /// <summary>
    /// Double-clicks a node.
    /// </summary>
    [HttpPost("{sessionId}/nodes/{nodeKey}/double-click")]
    public async Task<IActionResult> DoubleClickNode(
        string sessionId,
        string nodeKey,
        [FromQuery] string treePath)
    {
        try
        {
            await _treeService.DoubleClickNodeAsync(sessionId, treePath, nodeKey);
            return Ok(new { Message = "Node double-clicked successfully" });
        }
        catch (Exception ex)
        {
            _logger.Error(ex, "Error double-clicking node");
            return StatusCode(500, new { Error = ex.Message });
        }
    }

    /// <summary>
    /// Gets the selected node.
    /// </summary>
    [HttpGet("{sessionId}/selected-node")]
    public async Task<IActionResult> GetSelectedNode(
        string sessionId,
        [FromQuery] string treePath)
    {
        try
        {
            var nodeKey = await _treeService.GetSelectedNodeKeyAsync(sessionId, treePath);
            return Ok(new { SelectedNodeKey = nodeKey });
        }
        catch (Exception ex)
        {
            _logger.Error(ex, "Error getting selected node");
            return StatusCode(500, new { Error = ex.Message });
        }
    }

    /// <summary>
    /// Finds nodes matching conditions.
    /// </summary>
    [HttpPost("{sessionId}/find")]
    public async Task<IActionResult> FindNodes(
        string sessionId,
        [FromQuery] string treePath,
        [FromBody] List<QueryCondition> conditions,
        [FromQuery] bool searchInText = true)
    {
        try
        {
            var matches = await _treeService.FindNodesAsync(sessionId, treePath, conditions, searchInText);
            return Ok(matches);
        }
        catch (Exception ex)
        {
            _logger.Error(ex, "Error finding nodes");
            return StatusCode(500, new { Error = ex.Message });
        }
    }

    /// <summary>
    /// Searches for a node by text.
    /// </summary>
    [HttpGet("{sessionId}/search")]
    public async Task<IActionResult> SearchByText(
        string sessionId,
        [FromQuery] string treePath,
        [FromQuery] string searchText,
        [FromQuery] bool caseSensitive = false)
    {
        try
        {
            var result = await _treeService.SearchNodeByTextAsync(sessionId, treePath, searchText, caseSensitive);
            return Ok(result);
        }
        catch (Exception ex)
        {
            _logger.Error(ex, "Error searching nodes");
            return StatusCode(500, new { Error = ex.Message });
        }
    }

    /// <summary>
    /// Gets the path to a node.
    /// </summary>
    [HttpGet("{sessionId}/nodes/{nodeKey}/path")]
    public async Task<IActionResult> GetNodePath(
        string sessionId,
        string nodeKey,
        [FromQuery] string treePath)
    {
        try
        {
            var path = await _treeService.GetNodePathAsync(sessionId, treePath, nodeKey);
            return Ok(new { Path = path });
        }
        catch (Exception ex)
        {
            _logger.Error(ex, "Error getting node path");
            return StatusCode(500, new { Error = ex.Message });
        }
    }
}

