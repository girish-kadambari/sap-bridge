using SapBridge.Models;
using SapBridge.Models.Query;

namespace SapBridge.Services.Table;

/// <summary>
/// Service interface for SAP GUI Table (GuiTableControl) operations.
/// Supports data extraction, row selection, and query-based filtering.
/// </summary>
public interface ITableService
{
    /// <summary>
    /// Gets all data from a table with optional pagination.
    /// </summary>
    /// <param name="sessionId">The session ID.</param>
    /// <param name="tablePath">Path to the table object.</param>
    /// <param name="startRow">Starting row index (0-based).</param>
    /// <param name="rowCount">Number of rows to retrieve (null for all).</param>
    /// <returns>Table data with columns and rows.</returns>
    Task<TableData> GetAllDataAsync(string sessionId, string tablePath, int startRow = 0, int? rowCount = null);

    /// <summary>
    /// Gets a specific row from the table.
    /// </summary>
    /// <param name="sessionId">The session ID.</param>
    /// <param name="tablePath">Path to the table object.</param>
    /// <param name="rowIndex">Row index (0-based).</param>
    /// <returns>The table row data.</returns>
    Task<TableRow> GetRowAsync(string sessionId, string tablePath, int rowIndex);

    /// <summary>
    /// Gets a specific cell value.
    /// </summary>
    /// <param name="sessionId">The session ID.</param>
    /// <param name="tablePath">Path to the table object.</param>
    /// <param name="row">Row index.</param>
    /// <param name="column">Column name.</param>
    /// <returns>The cell data.</returns>
    Task<TableCell> GetCellAsync(string sessionId, string tablePath, int row, string column);

    /// <summary>
    /// Sets a cell value.
    /// </summary>
    /// <param name="sessionId">The session ID.</param>
    /// <param name="tablePath">Path to the table object.</param>
    /// <param name="row">Row index.</param>
    /// <param name="column">Column name.</param>
    /// <param name="value">Value to set.</param>
    Task SetCellAsync(string sessionId, string tablePath, int row, string column, string value);

    /// <summary>
    /// Finds rows matching query conditions.
    /// </summary>
    /// <param name="sessionId">The session ID.</param>
    /// <param name="tablePath">Path to the table object.</param>
    /// <param name="conditions">Query conditions to match.</param>
    /// <returns>List of matching rows as query matches.</returns>
    Task<List<QueryMatch>> FindRowsAsync(string sessionId, string tablePath, List<QueryCondition> conditions);

    /// <summary>
    /// Finds the first row matching query conditions.
    /// </summary>
    /// <param name="sessionId">The session ID.</param>
    /// <param name="tablePath">Path to the table object.</param>
    /// <param name="conditions">Query conditions to match.</param>
    /// <returns>First matching row or null.</returns>
    Task<QueryMatch?> FindFirstRowAsync(string sessionId, string tablePath, List<QueryCondition> conditions);

    /// <summary>
    /// Finds the last row matching query conditions.
    /// </summary>
    /// <param name="sessionId">The session ID.</param>
    /// <param name="tablePath">Path to the table object.</param>
    /// <param name="conditions">Query conditions to match.</param>
    /// <returns>Last matching row or null.</returns>
    Task<QueryMatch?> FindLastRowAsync(string sessionId, string tablePath, List<QueryCondition> conditions);

    /// <summary>
    /// Gets the index of the first empty row (all specified columns are empty).
    /// </summary>
    /// <param name="sessionId">The session ID.</param>
    /// <param name="tablePath">Path to the table object.</param>
    /// <param name="columns">Columns to check for emptiness (null checks all).</param>
    /// <returns>Index of first empty row or -1 if none found.</returns>
    Task<int> GetFirstEmptyRowIndexAsync(string sessionId, string tablePath, List<string>? columns = null);

    /// <summary>
    /// Gets the total row count in the table.
    /// </summary>
    /// <param name="sessionId">The session ID.</param>
    /// <param name="tablePath">Path to the table object.</param>
    /// <returns>Total number of rows.</returns>
    Task<int> GetRowCountAsync(string sessionId, string tablePath);

    /// <summary>
    /// Gets column names from the table.
    /// </summary>
    /// <param name="sessionId">The session ID.</param>
    /// <param name="tablePath">Path to the table object.</param>
    /// <returns>List of column names.</returns>
    Task<List<string>> GetColumnNamesAsync(string sessionId, string tablePath);

    /// <summary>
    /// Selects a specific row in the table.
    /// </summary>
    /// <param name="sessionId">The session ID.</param>
    /// <param name="tablePath">Path to the table object.</param>
    /// <param name="rowIndex">Index of row to select.</param>
    Task SelectRowAsync(string sessionId, string tablePath, int rowIndex);

    /// <summary>
    /// Gets the currently selected row index.
    /// </summary>
    /// <param name="sessionId">The session ID.</param>
    /// <param name="tablePath">Path to the table object.</param>
    /// <returns>Selected row index or -1 if none selected.</returns>
    Task<int> GetSelectedRowAsync(string sessionId, string tablePath);

    /// <summary>
    /// Executes a complete query against the table.
    /// </summary>
    /// <param name="sessionId">The session ID.</param>
    /// <param name="query">The query to execute.</param>
    /// <returns>Query result with matches.</returns>
    Task<QueryResult> ExecuteQueryAsync(string sessionId, SapQuery query);
}
