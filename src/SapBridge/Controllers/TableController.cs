using Microsoft.AspNetCore.Mvc;
using SapBridge.Models.Query;
using SapBridge.Services.Table;
using Serilog;

namespace SapBridge.Controllers;

/// <summary>
/// Controller for SAP GUI Table operations.
/// </summary>
[ApiController]
[Route("api/[controller]")]
public class TableController : ControllerBase
{
    private readonly ILogger _logger;
    private readonly ITableService _tableService;

    public TableController(ILogger logger, ITableService tableService)
    {
        _logger = logger;
        _tableService = tableService;
    }

    /// <summary>
    /// Gets all data from a table with optional pagination.
    /// </summary>
    [HttpGet("{sessionId}/data")]
    public async Task<IActionResult> GetAllData(
        string sessionId,
        [FromQuery] string tablePath,
        [FromQuery] int startRow = 0,
        [FromQuery] int? rowCount = null)
    {
        try
        {
            var data = await _tableService.GetAllDataAsync(sessionId, tablePath, startRow, rowCount);
            return Ok(data);
        }
        catch (Exception ex)
        {
            _logger.Error(ex, "Error getting table data");
            return StatusCode(500, new { Error = ex.Message });
        }
    }

    /// <summary>
    /// Gets a specific row from the table.
    /// </summary>
    [HttpGet("{sessionId}/rows/{rowIndex}")]
    public async Task<IActionResult> GetRow(
        string sessionId,
        int rowIndex,
        [FromQuery] string tablePath)
    {
        try
        {
            var row = await _tableService.GetRowAsync(sessionId, tablePath, rowIndex);
            return Ok(row);
        }
        catch (Exception ex)
        {
            _logger.Error(ex, "Error getting table row");
            return StatusCode(500, new { Error = ex.Message });
        }
    }

    /// <summary>
    /// Gets a specific cell value.
    /// </summary>
    [HttpGet("{sessionId}/cells")]
    public async Task<IActionResult> GetCell(
        string sessionId,
        [FromQuery] string tablePath,
        [FromQuery] int row,
        [FromQuery] string column)
    {
        try
        {
            var cell = await _tableService.GetCellAsync(sessionId, tablePath, row, column);
            return Ok(cell);
        }
        catch (Exception ex)
        {
            _logger.Error(ex, "Error getting table cell");
            return StatusCode(500, new { Error = ex.Message });
        }
    }

    /// <summary>
    /// Sets a cell value.
    /// </summary>
    [HttpPut("{sessionId}/cells")]
    public async Task<IActionResult> SetCell(
        string sessionId,
        [FromQuery] string tablePath,
        [FromQuery] int row,
        [FromQuery] string column,
        [FromBody] string value)
    {
        try
        {
            await _tableService.SetCellAsync(sessionId, tablePath, row, column, value);
            return Ok(new { Message = "Cell updated successfully" });
        }
        catch (Exception ex)
        {
            _logger.Error(ex, "Error setting table cell");
            return StatusCode(500, new { Error = ex.Message });
        }
    }

    /// <summary>
    /// Finds rows matching conditions.
    /// </summary>
    [HttpPost("{sessionId}/find")]
    public async Task<IActionResult> FindRows(
        string sessionId,
        [FromQuery] string tablePath,
        [FromBody] List<QueryCondition> conditions)
    {
        try
        {
            var matches = await _tableService.FindRowsAsync(sessionId, tablePath, conditions);
            return Ok(matches);
        }
        catch (Exception ex)
        {
            _logger.Error(ex, "Error finding table rows");
            return StatusCode(500, new { Error = ex.Message });
        }
    }

    /// <summary>
    /// Gets the first empty row index.
    /// </summary>
    [HttpGet("{sessionId}/empty-row")]
    public async Task<IActionResult> GetFirstEmptyRow(
        string sessionId,
        [FromQuery] string tablePath,
        [FromQuery] List<string>? columns = null)
    {
        try
        {
            var index = await _tableService.GetFirstEmptyRowIndexAsync(sessionId, tablePath, columns);
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
        [FromQuery] string tablePath)
    {
        try
        {
            var count = await _tableService.GetRowCountAsync(sessionId, tablePath);
            return Ok(new { RowCount = count });
        }
        catch (Exception ex)
        {
            _logger.Error(ex, "Error getting row count");
            return StatusCode(500, new { Error = ex.Message });
        }
    }

    /// <summary>
    /// Gets column names.
    /// </summary>
    [HttpGet("{sessionId}/columns")]
    public async Task<IActionResult> GetColumns(
        string sessionId,
        [FromQuery] string tablePath)
    {
        try
        {
            var columns = await _tableService.GetColumnNamesAsync(sessionId, tablePath);
            return Ok(columns);
        }
        catch (Exception ex)
        {
            _logger.Error(ex, "Error getting columns");
            return StatusCode(500, new { Error = ex.Message });
        }
    }

    /// <summary>
    /// Selects a row.
    /// </summary>
    [HttpPost("{sessionId}/select-row")]
    public async Task<IActionResult> SelectRow(
        string sessionId,
        [FromQuery] string tablePath,
        [FromQuery] int rowIndex)
    {
        try
        {
            await _tableService.SelectRowAsync(sessionId, tablePath, rowIndex);
            return Ok(new { Message = "Row selected successfully" });
        }
        catch (Exception ex)
        {
            _logger.Error(ex, "Error selecting row");
            return StatusCode(500, new { Error = ex.Message });
        }
    }

    /// <summary>
    /// Gets the selected row.
    /// </summary>
    [HttpGet("{sessionId}/selected-row")]
    public async Task<IActionResult> GetSelectedRow(
        string sessionId,
        [FromQuery] string tablePath)
    {
        try
        {
            var selectedRow = await _tableService.GetSelectedRowAsync(sessionId, tablePath);
            return Ok(new { SelectedRowIndex = selectedRow });
        }
        catch (Exception ex)
        {
            _logger.Error(ex, "Error getting selected row");
            return StatusCode(500, new { Error = ex.Message });
        }
    }
}

