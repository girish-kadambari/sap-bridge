using SapBridge.Models;
using SapBridge.Repositories;
using SapBridge.Utils;
using Serilog;
using ILogger = Serilog.ILogger;

namespace SapBridge.Services.Grid;

/// <summary>
/// Extracts data from SAP GUI Grid (GuiGridView) objects.
/// Handles column definitions, row data, and cell values.
/// </summary>
public class GridDataExtractor
{
    private readonly ILogger _logger;
    private readonly ISapGuiRepository _repository;

    public GridDataExtractor(ILogger logger, ISapGuiRepository repository)
    {
        _logger = logger ?? throw new ArgumentNullException(nameof(logger));
        _repository = repository ?? throw new ArgumentNullException(nameof(repository));
    }

    /// <summary>
    /// Extracts all data from a grid.
    /// </summary>
    /// <param name="grid">The grid COM object.</param>
    /// <param name="gridPath">Path to the grid.</param>
    /// <param name="startRow">Starting row index.</param>
    /// <param name="rowCount">Number of rows to extract (null for all).</param>
    /// <returns>Complete grid data.</returns>
    public GridData ExtractData(object grid, string gridPath, int startRow = 0, int? rowCount = null)
    {
        var gridData = new GridData
        {
            GridPath = gridPath,
            CapturedAt = DateTime.UtcNow
        };

        try
        {
            // Get row and column counts
            gridData.TotalRows = GetRowCount(grid);
            gridData.TotalColumns = GetColumnCount(grid);

            _logger.Debug("Grid has {RowCount} rows and {ColumnCount} columns", 
                gridData.TotalRows, gridData.TotalColumns);

            // Extract column definitions
            gridData.Columns = ExtractColumns(grid);

            // Determine how many rows to extract
            int endRow = rowCount.HasValue 
                ? Math.Min(startRow + rowCount.Value, gridData.TotalRows)
                : gridData.TotalRows;

            // Extract row data
            for (int i = startRow; i < endRow; i++)
            {
                var row = ExtractRow(grid, i, gridData.Columns);
                if (row != null)
                {
                    gridData.Rows.Add(row);
                }
            }

            _logger.Information("Extracted {RowsExtracted} rows from grid", gridData.Rows.Count);
        }
        catch (Exception ex)
        {
            _logger.Error(ex, "Error extracting grid data from {GridPath}", gridPath);
            throw;
        }

        return gridData;
    }

    /// <summary>
    /// Extracts a single row from the grid.
    /// </summary>
    public GridRow? ExtractRow(object grid, int rowIndex, List<GridColumn>? columns = null)
    {
        try
        {
            // Get columns if not provided
            if (columns == null || columns.Count == 0)
            {
                columns = ExtractColumns(grid);
            }

            var row = new GridRow
            {
                Index = rowIndex,
                Cells = new Dictionary<string, object?>()
            };

            // Extract cell values for each column
            foreach (var column in columns)
            {
                var cellValue = GetCellValue(grid, rowIndex, column.Name);
                row.Cells[column.Name] = cellValue;
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
    /// Extracts column definitions from the grid.
    /// </summary>
    public List<GridColumn> ExtractColumns(object grid)
    {
        var columns = new List<GridColumn>();

        try
        {
            var columnCount = GetColumnCount(grid);

            for (int i = 0; i < columnCount; i++)
            {
                var column = ExtractColumn(grid, i);
                if (column != null)
                {
                    columns.Add(column);
                }
            }

            _logger.Debug("Extracted {ColumnCount} columns", columns.Count);
        }
        catch (Exception ex)
        {
            _logger.Error(ex, "Error extracting columns");
        }

        return columns;
    }

    /// <summary>
    /// Extracts a single column definition.
    /// </summary>
    private GridColumn? ExtractColumn(object grid, int columnIndex)
    {
        try
        {
            // Get column name
            var columnName = ComReflectionHelper.GetPropertySafe<string>(grid, $"ColumnName[{columnIndex}]") 
                ?? $"Column{columnIndex}";

            // Get column title
            var columnTitle = ComReflectionHelper.GetPropertySafe<string>(grid, $"ColumnTitle[{columnIndex}]") 
                ?? columnName;

            // Get column width
            var columnWidth = ComReflectionHelper.GetPropertySafe<int>(grid, $"ColumnWidth[{columnIndex}]");

            return new GridColumn
            {
                Name = columnName,
                Title = columnTitle,
                Index = columnIndex,
                Width = columnWidth,
                Visible = true // TODO: Check if column is visible
            };
        }
        catch (Exception ex)
        {
            _logger.Debug(ex, "Could not extract column {ColumnIndex}", columnIndex);
            return null;
        }
    }

    /// <summary>
    /// Gets the value of a specific cell.
    /// </summary>
    public object? GetCellValue(object grid, int row, string column)
    {
        try
        {
            var value = _repository.InvokeObjectMethod(grid, "GetCellValue", row, column);
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
    public void SetCellValue(object grid, int row, string column, string value)
    {
        try
        {
            _repository.InvokeObjectMethod(grid, "SetCellValue", row, column, value);
            _logger.Debug("Set cell value at row {Row}, column {Column} to '{Value}'", row, column, value);
        }
        catch (Exception ex)
        {
            _logger.Error(ex, "Error setting cell value at row {Row}, column {Column}", row, column);
            throw;
        }
    }

    /// <summary>
    /// Gets the total number of rows in the grid.
    /// </summary>
    public int GetRowCount(object grid)
    {
        try
        {
            var rowCount = _repository.GetObjectProperty(grid, "RowCount");
            return rowCount != null ? Convert.ToInt32(rowCount) : 0;
        }
        catch (Exception ex)
        {
            _logger.Warning(ex, "Could not get row count, returning 0");
            return 0;
        }
    }

    /// <summary>
    /// Gets the total number of columns in the grid.
    /// </summary>
    public int GetColumnCount(object grid)
    {
        try
        {
            var columnCount = _repository.GetObjectProperty(grid, "ColumnCount");
            return columnCount != null ? Convert.ToInt32(columnCount) : 0;
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
    public bool IsRowEmpty(object grid, int rowIndex, List<string>? columnsToCheck = null)
    {
        try
        {
            // Get columns if not specified
            if (columnsToCheck == null || columnsToCheck.Count == 0)
            {
                var allColumns = ExtractColumns(grid);
                columnsToCheck = allColumns.Select(c => c.Name).ToList();
            }

            // Check each column
            foreach (var column in columnsToCheck)
            {
                var value = GetCellValue(grid, rowIndex, column);
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
}

