using SapBridge.Models;
using SapBridge.Repositories;
using SapBridge.Utils;
using Serilog;

namespace SapBridge.Services.Table;

/// <summary>
/// Extracts data from SAP GUI Table (GuiTableControl) objects.
/// Handles column definitions, row data, and cell values.
/// </summary>
public class TableDataExtractor
{
    private readonly ILogger _logger;
    private readonly ISapGuiRepository _repository;

    public TableDataExtractor(ILogger logger, ISapGuiRepository repository)
    {
        _logger = logger ?? throw new ArgumentNullException(nameof(logger));
        _repository = repository ?? throw new ArgumentNullException(nameof(repository));
    }

    /// <summary>
    /// Extracts all data from a table.
    /// </summary>
    /// <param name="table">The table COM object.</param>
    /// <param name="tablePath">Path to the table.</param>
    /// <param name="startRow">Starting row index.</param>
    /// <param name="rowCount">Number of rows to extract (null for all).</param>
    /// <returns>Complete table data.</returns>
    public TableData ExtractData(object table, string tablePath, int startRow = 0, int? rowCount = null)
    {
        var tableData = new TableData
        {
            TablePath = tablePath,
            CapturedAt = DateTime.UtcNow
        };

        try
        {
            // Get row and column counts
            tableData.TotalRows = GetRowCount(table);
            tableData.TotalColumns = GetColumnCount(table);

            _logger.Debug("Table has {RowCount} rows and {ColumnCount} columns", 
                tableData.TotalRows, tableData.TotalColumns);

            // Extract column names
            tableData.ColumnNames = ExtractColumnNames(table);
            
            // Build column map
            for (int i = 0; i < tableData.ColumnNames.Count; i++)
            {
                tableData.ColumnMap[tableData.ColumnNames[i]] = i;
            }

            // Determine how many rows to extract
            int endRow = rowCount.HasValue 
                ? Math.Min(startRow + rowCount.Value, tableData.TotalRows)
                : tableData.TotalRows;

            // Extract row data
            for (int i = startRow; i < endRow; i++)
            {
                var row = ExtractRow(table, i, tableData.ColumnNames);
                if (row != null)
                {
                    tableData.Rows.Add(row);
                }
            }

            _logger.Information("Extracted {RowsExtracted} rows from table", tableData.Rows.Count);
        }
        catch (Exception ex)
        {
            _logger.Error(ex, "Error extracting table data from {TablePath}", tablePath);
            throw;
        }

        return tableData;
    }

    /// <summary>
    /// Extracts a single row from the table.
    /// </summary>
    public TableRow? ExtractRow(object table, int rowIndex, List<string>? columnNames = null)
    {
        try
        {
            // Get column names if not provided
            if (columnNames == null || columnNames.Count == 0)
            {
                columnNames = ExtractColumnNames(table);
            }

            var row = new TableRow
            {
                Index = rowIndex,
                Cells = new Dictionary<string, object?>()
            };

            // Extract cell values for each column
            foreach (var columnName in columnNames)
            {
                var cellValue = GetCellValue(table, rowIndex, columnName);
                row.Cells[columnName] = cellValue;
            }

            // Try to get row key
            try
            {
                var key = _repository.GetObjectProperty(table, $"GetAbsoluteRow({rowIndex})");
                row.Key = key?.ToString();
            }
            catch
            {
                // Row key not available, not critical
            }

            return row;
        }
        catch (Exception ex)
        {
            _logger.Warning(ex, "Error extracting row {RowIndex}", rowIndex);
            return null;
        }
    }

    /// <summary>
    /// Extracts column names from the table.
    /// </summary>
    public List<string> ExtractColumnNames(object table)
    {
        var columnNames = new List<string>();

        try
        {
            var columnCount = GetColumnCount(table);

            for (int i = 0; i < columnCount; i++)
            {
                // Try to get column title/name
                var columnName = ComReflectionHelper.GetPropertySafe<string>(table, $"GetColumnTitle({i})");
                
                if (string.IsNullOrEmpty(columnName))
                {
                    // Try alternative property names
                    columnName = ComReflectionHelper.GetPropertySafe<string>(table, $"ColumnName[{i}]");
                }
                
                if (string.IsNullOrEmpty(columnName))
                {
                    columnName = $"Column{i}";
                }

                columnNames.Add(columnName);
            }

            _logger.Debug("Extracted {ColumnCount} column names", columnNames.Count);
        }
        catch (Exception ex)
        {
            _logger.Error(ex, "Error extracting column names");
        }

        return columnNames;
    }

    /// <summary>
    /// Gets the value of a specific cell.
    /// </summary>
    public object? GetCellValue(object table, int row, string column)
    {
        try
        {
            var value = _repository.InvokeObjectMethod(table, "GetCellValue", row, column);
            return value;
        }
        catch (Exception ex)
        {
            _logger.Debug(ex, "Could not get cell value at row {Row}, column {Column}", row, column);
            return null;
        }
    }

    /// <summary>
    /// Sets the value of a specific cell.
    /// </summary>
    public void SetCellValue(object table, int row, string column, string value)
    {
        try
        {
            _repository.InvokeObjectMethod(table, "SetCellValue", row, column, value);
            _logger.Debug("Set cell value at row {Row}, column {Column} to '{Value}'", row, column, value);
        }
        catch (Exception ex)
        {
            _logger.Error(ex, "Error setting cell value at row {Row}, column {Column}", row, column);
            throw;
        }
    }

    /// <summary>
    /// Gets the total number of rows in the table.
    /// </summary>
    public int GetRowCount(object table)
    {
        try
        {
            // Try RowCount property
            var rowCount = _repository.GetObjectProperty(table, "RowCount");
            if (rowCount != null)
            {
                return Convert.ToInt32(rowCount);
            }

            // Try VisibleRowCount as fallback
            var visibleRowCount = _repository.GetObjectProperty(table, "VisibleRowCount");
            if (visibleRowCount != null)
            {
                return Convert.ToInt32(visibleRowCount);
            }

            return 0;
        }
        catch (Exception ex)
        {
            _logger.Warning(ex, "Could not get row count, returning 0");
            return 0;
        }
    }

    /// <summary>
    /// Gets the total number of columns in the table.
    /// </summary>
    public int GetColumnCount(object table)
    {
        try
        {
            var columnCount = _repository.GetObjectProperty(table, "ColumnCount");
            if (columnCount != null)
            {
                return Convert.ToInt32(columnCount);
            }

            // Try alternative property
            var columns = _repository.GetObjectProperty(table, "Columns");
            if (columns != null)
            {
                var count = _repository.GetObjectProperty(columns, "Count");
                if (count != null)
                {
                    return Convert.ToInt32(count);
                }
            }

            return 0;
        }
        catch (Exception ex)
        {
            _logger.Warning(ex, "Could not get column count, returning 0");
            return 0;
        }
    }

    /// <summary>
    /// Checks if a row is empty (all cells are empty).
    /// </summary>
    public bool IsRowEmpty(object table, int rowIndex, List<string>? columnsToCheck = null)
    {
        try
        {
            // Get column names if not specified
            if (columnsToCheck == null || columnsToCheck.Count == 0)
            {
                columnsToCheck = ExtractColumnNames(table);
            }

            // Check each column
            foreach (var column in columnsToCheck)
            {
                var value = GetCellValue(table, rowIndex, column);
                if (value != null && !string.IsNullOrWhiteSpace(value.ToString()))
                {
                    return false; // Found non-empty cell
                }
            }

            return true; // All cells are empty
        }
        catch (Exception ex)
        {
            _logger.Warning(ex, "Error checking if row {RowIndex} is empty", rowIndex);
            return false;
        }
    }

    /// <summary>
    /// Gets the currently selected row index.
    /// </summary>
    public int GetSelectedRow(object table)
    {
        try
        {
            var selectedRow = _repository.GetObjectProperty(table, "CurrentCellRow");
            if (selectedRow != null)
            {
                return Convert.ToInt32(selectedRow);
            }

            return -1;
        }
        catch (Exception ex)
        {
            _logger.Debug(ex, "Could not get selected row, returning -1");
            return -1;
        }
    }

    /// <summary>
    /// Selects a row in the table.
    /// </summary>
    public void SelectRow(object table, int rowIndex)
    {
        try
        {
            // Try setting CurrentCellRow property
            _repository.SetObjectProperty(table, "CurrentCellRow", rowIndex);
            _logger.Debug("Selected row {RowIndex}", rowIndex);
        }
        catch
        {
            try
            {
                // Try alternative method - select cell in first column
                var columnNames = ExtractColumnNames(table);
                if (columnNames.Count > 0)
                {
                    _repository.InvokeObjectMethod(table, "SetCurrentCell", rowIndex, columnNames[0]);
                    _logger.Debug("Selected row {RowIndex} via SetCurrentCell", rowIndex);
                }
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Error selecting row {RowIndex}", rowIndex);
                throw;
            }
        }
    }
}
