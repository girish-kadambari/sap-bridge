using Microsoft.AspNetCore.Mvc;
using SapBridge.Models;
using SapBridge.Models.Query;
using SapBridge.Services.Grid;
using Serilog;
using ILogger = Serilog.ILogger;

namespace SapBridge.Controllers;

/// <summary>
/// Controller for SAP GUI Grid operations.
/// </summary>
[ApiController]
[Route("api/[controller]")]
public class GridController : ControllerBase
{
    private readonly ILogger _logger;
    private readonly IGridService _gridService;

    public GridController(ILogger logger, IGridService gridService)
    {
        _logger = logger;
        _gridService = gridService;
    }

    /// <summary>
    /// Gets all data from a grid with optional pagination.
    /// </summary>
    /// <param name="sessionId">The session ID.</param>
    /// <param name="gridPath">Path to the grid object.</param>
    /// <param name="startRow">Starting row index (0-based).</param>
    /// <param name="rowCount">Number of rows to retrieve.</param>
    [HttpGet("{sessionId}/data")]
    public async Task<IActionResult> GetAllData(
        string sessionId,
        [FromQuery] string gridPath,
        [FromQuery] int startRow = 0,
        [FromQuery] int? rowCount = null)
    {
        try
        {
            var data = await _gridService.GetAllDataAsync(sessionId, gridPath, startRow, rowCount);
            return Ok(data);
        }
        catch (Exception ex)
        {
            _logger.Error(ex, "Error getting grid data");
            return StatusCode(500, new { Error = ex.Message });
        }
    }

    /// <summary>
    /// Gets a specific row from the grid.
    /// </summary>
    [HttpGet("{sessionId}/rows/{rowIndex}")]
    public async Task<IActionResult> GetRow(
        string sessionId,
        int rowIndex,
        [FromQuery] string gridPath)
    {
        try
        {
            var row = await _gridService.GetRowAsync(sessionId, gridPath, rowIndex);
            return Ok(row);
        }
        catch (Exception ex)
        {
            _logger.Error(ex, "Error getting grid row");
            return StatusCode(500, new { Error = ex.Message });
        }
    }

    /// <summary>
    /// Gets a specific cell value.
    /// </summary>
    [HttpGet("{sessionId}/cells")]
    public async Task<IActionResult> GetCell(
        string sessionId,
        [FromQuery] string gridPath,
        [FromQuery] int row,
        [FromQuery] string column)
    {
        try
        {
            var cell = await _gridService.GetCellAsync(sessionId, gridPath, row, column);
            return Ok(cell);
        }
        catch (Exception ex)
        {
            _logger.Error(ex, "Error getting grid cell");
            return StatusCode(500, new { Error = ex.Message });
        }
    }

    /// <summary>
    /// Sets a cell value.
    /// </summary>
    [HttpPut("{sessionId}/cells")]
    public async Task<IActionResult> SetCell(
        string sessionId,
        [FromQuery] string gridPath,
        [FromQuery] int row,
        [FromQuery] string column,
        [FromBody] string value)
    {
        try
        {
            await _gridService.SetCellAsync(sessionId, gridPath, row, column, value);
            return Ok(new { Message = "Cell updated successfully" });
        }
        catch (Exception ex)
        {
            _logger.Error(ex, "Error setting grid cell");
            return StatusCode(500, new { Error = ex.Message });
        }
    }

    /// <summary>
    /// Finds rows matching conditions.
    /// </summary>
    [HttpPost("{sessionId}/find")]
    public async Task<IActionResult> FindRows(
        string sessionId,
        [FromQuery] string gridPath,
        [FromBody] List<QueryCondition> conditions)
    {
        try
        {
            var matches = await _gridService.FindRowsAsync(sessionId, gridPath, conditions);
            return Ok(matches);
        }
        catch (Exception ex)
        {
            _logger.Error(ex, "Error finding grid rows");
            return StatusCode(500, new { Error = ex.Message });
        }
    }

    /// <summary>
    /// Gets the first empty row index.
    /// </summary>
    [HttpGet("{sessionId}/empty-row")]
    public async Task<IActionResult> GetFirstEmptyRow(
        string sessionId,
        [FromQuery] string gridPath,
        [FromQuery] List<string>? columns = null)
    {
        try
        {
            var index = await _gridService.GetFirstEmptyRowIndexAsync(sessionId, gridPath, columns);
            return Ok(new { EmptyRowIndex = index });
        }
        catch (Exception ex)
        {
            _logger.Error(ex, "Error getting first empty row");
            return StatusCode(500, new { Error = ex.Message });
        }
    }

    /// <summary>
    /// Gets the row count.
    /// </summary>
    [HttpGet("{sessionId}/row-count")]
    public async Task<IActionResult> GetRowCount(
        string sessionId,
        [FromQuery] string gridPath)
    {
        try
        {
            var count = await _gridService.GetRowCountAsync(sessionId, gridPath);
            return Ok(new { RowCount = count });
        }
        catch (Exception ex)
        {
            _logger.Error(ex, "Error getting row count");
            return StatusCode(500, new { Error = ex.Message });
        }
    }

    /// <summary>
    /// Gets column definitions.
    /// </summary>
    [HttpGet("{sessionId}/columns")]
    public async Task<IActionResult> GetColumns(
        string sessionId,
        [FromQuery] string gridPath)
    {
        try
        {
            var columns = await _gridService.GetColumnsAsync(sessionId, gridPath);
            return Ok(columns);
        }
        catch (Exception ex)
        {
            _logger.Error(ex, "Error getting columns");
            return StatusCode(500, new { Error = ex.Message });
        }
    }

    /// <summary>
    /// Selects rows.
    /// </summary>
    [HttpPost("{sessionId}/select-rows")]
    public async Task<IActionResult> SelectRows(
        string sessionId,
        [FromQuery] string gridPath,
        [FromBody] List<int> rowIndices)
    {
        try
        {
            await _gridService.SelectRowsAsync(sessionId, gridPath, rowIndices);
            return Ok(new { Message = "Rows selected successfully" });
        }
        catch (Exception ex)
        {
            _logger.Error(ex, "Error selecting rows");
            return StatusCode(500, new { Error = ex.Message });
        }
    }

    /// <summary>
    /// Scrolls to a row.
    /// </summary>
    [HttpPost("{sessionId}/scroll")]
    public async Task<IActionResult> ScrollToRow(
        string sessionId,
        [FromQuery] string gridPath,
        [FromQuery] int rowIndex)
    {
        try
        {
            await _gridService.ScrollToRowAsync(sessionId, gridPath, rowIndex);
            return Ok(new { Message = "Scrolled successfully" });
        }
        catch (Exception ex)
        {
            _logger.Error(ex, "Error scrolling");
            return StatusCode(500, new { Error = ex.Message });
        }
    }
}

