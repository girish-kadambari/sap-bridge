using SapBridge.Models;
using SapBridge.Models.Query;

namespace SapBridge.Services.Grid;

/// <summary>
/// Service interface for SAP GUI Grid (GuiGridView) operations.
/// Supports data extraction, navigation, selection, and query-based filtering.
/// </summary>
public interface IGridService
{
    /// <summary>
    /// Gets all data from a grid with optional pagination.
    /// </summary>
    /// <param name="sessionId">The session ID.</param>
    /// <param name="gridPath">Path to the grid object.</param>
    /// <param name="startRow">Starting row index (0-based).</param>
    /// <param name="rowCount">Number of rows to retrieve (null for all).</param>
    /// <returns>Grid data with columns and rows.</returns>
    Task<GridData> GetAllDataAsync(string sessionId, string gridPath, int startRow = 0, int? rowCount = null);

    /// <summary>
    /// Gets a specific row from the grid.
    /// </summary>
    /// <param name="sessionId">The session ID.</param>
    /// <param name="gridPath">Path to the grid object.</param>
    /// <param name="rowIndex">Row index (0-based).</param>
    /// <returns>The grid row data.</returns>
    Task<GridRow> GetRowAsync(string sessionId, string gridPath, int rowIndex);

    /// <summary>
    /// Gets a specific cell value.
    /// </summary>
    /// <param name="sessionId">The session ID.</param>
    /// <param name="gridPath">Path to the grid object.</param>
    /// <param name="row">Row index.</param>
    /// <param name="column">Column name.</param>
    /// <returns>The cell data.</returns>
    Task<GridCell> GetCellAsync(string sessionId, string gridPath, int row, string column);

    /// <summary>
    /// Sets a cell value.
    /// </summary>
    /// <param name="sessionId">The session ID.</param>
    /// <param name="gridPath">Path to the grid object.</param>
    /// <param name="row">Row index.</param>
    /// <param name="column">Column name.</param>
    /// <param name="value">Value to set.</param>
    Task SetCellAsync(string sessionId, string gridPath, int row, string column, string value);

    /// <summary>
    /// Finds rows matching query conditions.
    /// </summary>
    /// <param name="sessionId">The session ID.</param>
    /// <param name="gridPath">Path to the grid object.</param>
    /// <param name="conditions">Query conditions to match.</param>
    /// <returns>List of matching rows as query matches.</returns>
    Task<List<QueryMatch>> FindRowsAsync(string sessionId, string gridPath, List<QueryCondition> conditions);

    /// <summary>
    /// Finds the first row matching query conditions.
    /// </summary>
    /// <param name="sessionId">The session ID.</param>
    /// <param name="gridPath">Path to the grid object.</param>
    /// <param name="conditions">Query conditions to match.</param>
    /// <returns>First matching row or null.</returns>
    Task<QueryMatch?> FindFirstRowAsync(string sessionId, string gridPath, List<QueryCondition> conditions);

    /// <summary>
    /// Finds the last row matching query conditions.
    /// </summary>
    /// <param name="sessionId">The session ID.</param>
    /// <param name="gridPath">Path to the grid object.</param>
    /// <param name="conditions">Query conditions to match.</param>
    /// <returns>Last matching row or null.</returns>
    Task<QueryMatch?> FindLastRowAsync(string sessionId, string gridPath, List<QueryCondition> conditions);

    /// <summary>
    /// Gets the index of the first empty row (all specified columns are empty).
    /// </summary>
    /// <param name="sessionId">The session ID.</param>
    /// <param name="gridPath">Path to the grid object.</param>
    /// <param name="columns">Columns to check for emptiness (null checks all).</param>
    /// <returns>Index of first empty row or -1 if none found.</returns>
    Task<int> GetFirstEmptyRowIndexAsync(string sessionId, string gridPath, List<string>? columns = null);

    /// <summary>
    /// Gets the total row count in the grid.
    /// </summary>
    /// <param name="sessionId">The session ID.</param>
    /// <param name="gridPath">Path to the grid object.</param>
    /// <returns>Total number of rows.</returns>
    Task<int> GetRowCountAsync(string sessionId, string gridPath);

    /// <summary>
    /// Gets column definitions from the grid.
    /// </summary>
    /// <param name="sessionId">The session ID.</param>
    /// <param name="gridPath">Path to the grid object.</param>
    /// <returns>List of grid columns.</returns>
    Task<List<GridColumn>> GetColumnsAsync(string sessionId, string gridPath);

    /// <summary>
    /// Selects rows in the grid by their indices.
    /// </summary>
    /// <param name="sessionId">The session ID.</param>
    /// <param name="gridPath">Path to the grid object.</param>
    /// <param name="rowIndices">Indices of rows to select.</param>
    Task SelectRowsAsync(string sessionId, string gridPath, List<int> rowIndices);

    /// <summary>
    /// Gets currently selected row indices.
    /// </summary>
    /// <param name="sessionId">The session ID.</param>
    /// <param name="gridPath">Path to the grid object.</param>
    /// <returns>List of selected row indices.</returns>
    Task<List<int>> GetSelectedRowsAsync(string sessionId, string gridPath);

    /// <summary>
    /// Scrolls the grid to a specific row.
    /// </summary>
    /// <param name="sessionId">The session ID.</param>
    /// <param name="gridPath">Path to the grid object.</param>
    /// <param name="rowIndex">Row index to scroll to.</param>
    Task ScrollToRowAsync(string sessionId, string gridPath, int rowIndex);

    /// <summary>
    /// Executes a complete query against the grid.
    /// </summary>
    /// <param name="sessionId">The session ID.</param>
    /// <param name="query">The query to execute.</param>
    /// <returns>Query result with matches.</returns>
    Task<QueryResult> ExecuteQueryAsync(string sessionId, SapQuery query);
}

