using SapBridge.Repositories;
using SapBridge.Utils;
using Serilog;
using ILogger = Serilog.ILogger;

namespace SapBridge.Services.Grid;

/// <summary>
/// Handles navigation and selection operations on SAP GUI Grids.
/// Manages scrolling, row selection, and focus.
/// </summary>
public class GridNavigator
{
    private readonly ILogger _logger;
    private readonly ISapGuiRepository _repository;

    public GridNavigator(ILogger logger, ISapGuiRepository repository)
    {
        _logger = logger ?? throw new ArgumentNullException(nameof(logger));
        _repository = repository ?? throw new ArgumentNullException(nameof(repository));
    }

    /// <summary>
    /// Selects specific rows in the grid.
    /// </summary>
    /// <param name="grid">The grid COM object.</param>
    /// <param name="rowIndices">Indices of rows to select.</param>
    public void SelectRows(object grid, List<int> rowIndices)
    {
        if (rowIndices == null || rowIndices.Count == 0)
        {
            _logger.Debug("No rows to select");
            return;
        }

        try
        {
            // Clear existing selection first
            ClearSelection(grid);

            // Select each row
            foreach (var rowIndex in rowIndices)
            {
                try
                {
                    _repository.InvokeObjectMethod(grid, "SelectRow", rowIndex);
                    _logger.Debug("Selected row {RowIndex}", rowIndex);
                }
                catch (Exception ex)
                {
                    _logger.Warning(ex, "Could not select row {RowIndex}", rowIndex);
                }
            }

            _logger.Information("Selected {Count} rows", rowIndices.Count);
        }
        catch (Exception ex)
        {
            _logger.Error(ex, "Error selecting rows");
            throw;
        }
    }

    /// <summary>
    /// Clears all row selections in the grid.
    /// </summary>
    public void ClearSelection(object grid)
    {
        try
        {
            _repository.InvokeObjectMethod(grid, "ClearSelection");
            _logger.Debug("Cleared grid selection");
        }
        catch (Exception ex)
        {
            _logger.Debug(ex, "Could not clear selection");
        }
    }

    /// <summary>
    /// Gets the indices of currently selected rows.
    /// </summary>
    /// <param name="grid">The grid COM object.</param>
    /// <returns>List of selected row indices.</returns>
    public List<int> GetSelectedRows(object grid)
    {
        var selectedRows = new List<int>();

        try
        {
            var selection = _repository.GetObjectProperty(grid, "SelectedRows");
            if (selection != null)
            {
                // Try to get selection as a collection
                var count = ComReflectionHelper.GetCollectionCount(selection);
                for (int i = 0; i < count; i++)
                {
                    var item = ComReflectionHelper.GetCollectionItem(selection, i);
                    if (item != null)
                    {
                        selectedRows.Add(Convert.ToInt32(item));
                    }
                }
            }

            _logger.Debug("Retrieved {Count} selected rows", selectedRows.Count);
        }
        catch (Exception ex)
        {
            _logger.Warning(ex, "Could not get selected rows");
        }

        return selectedRows;
    }

    /// <summary>
    /// Scrolls the grid to make a specific row visible.
    /// </summary>
    /// <param name="grid">The grid COM object.</param>
    /// <param name="rowIndex">Row index to scroll to.</param>
    public void ScrollToRow(object grid, int rowIndex)
    {
        try
        {
            // Set the first visible row
            _repository.SetObjectProperty(grid, "FirstVisibleRow", rowIndex);
            _logger.Debug("Scrolled to row {RowIndex}", rowIndex);
        }
        catch (Exception ex)
        {
            _logger.Warning(ex, "Could not scroll to row {RowIndex}, trying alternative method", rowIndex);
            
            // Try alternative method
            try
            {
                _repository.InvokeObjectMethod(grid, "SetCurrentCell", rowIndex, 0);
                _logger.Debug("Scrolled using SetCurrentCell method");
            }
            catch (Exception ex2)
            {
                _logger.Error(ex2, "Could not scroll to row using any method");
                throw;
            }
        }
    }

    /// <summary>
    /// Gets the index of the first visible row in the grid.
    /// </summary>
    /// <param name="grid">The grid COM object.</param>
    /// <returns>First visible row index.</returns>
    public int GetFirstVisibleRow(object grid)
    {
        try
        {
            var firstVisible = _repository.GetObjectProperty(grid, "FirstVisibleRow");
            return firstVisible != null ? Convert.ToInt32(firstVisible) : 0;
        }
        catch (Exception ex)
        {
            _logger.Debug(ex, "Could not get first visible row");
            return 0;
        }
    }

    /// <summary>
    /// Gets the number of rows that can be displayed at once (visible rows).
    /// </summary>
    /// <param name="grid">The grid COM object.</param>
    /// <returns>Number of visible rows.</returns>
    public int GetVisibleRowCount(object grid)
    {
        try
        {
            var visibleCount = _repository.GetObjectProperty(grid, "VisibleRowCount");
            return visibleCount != null ? Convert.ToInt32(visibleCount) : 0;
        }
        catch (Exception ex)
        {
            _logger.Debug(ex, "Could not get visible row count");
            return 0;
        }
    }

    /// <summary>
    /// Sets focus on a specific cell in the grid.
    /// </summary>
    /// <param name="grid">The grid COM object.</param>
    /// <param name="row">Row index.</param>
    /// <param name="column">Column name or index.</param>
    public void SetCurrentCell(object grid, int row, object column)
    {
        try
        {
            _repository.InvokeObjectMethod(grid, "SetCurrentCell", row, column);
            _logger.Debug("Set current cell to row {Row}, column {Column}", row, column);
        }
        catch (Exception ex)
        {
            _logger.Error(ex, "Error setting current cell");
            throw;
        }
    }

    /// <summary>
    /// Gets the current cell position.
    /// </summary>
    /// <param name="grid">The grid COM object.</param>
    /// <returns>Tuple of (row, column) or null if not available.</returns>
    public (int Row, int Column)? GetCurrentCell(object grid)
    {
        try
        {
            var currentRow = _repository.GetObjectProperty(grid, "CurrentCellRow");
            var currentColumn = _repository.GetObjectProperty(grid, "CurrentCellColumn");

            if (currentRow != null && currentColumn != null)
            {
                return (Convert.ToInt32(currentRow), Convert.ToInt32(currentColumn));
            }
        }
        catch (Exception ex)
        {
            _logger.Debug(ex, "Could not get current cell");
        }

        return null;
    }

    /// <summary>
    /// Presses a button in a specific cell (if the cell contains a button).
    /// </summary>
    /// <param name="grid">The grid COM object.</param>
    /// <param name="row">Row index.</param>
    /// <param name="column">Column name.</param>
    public void PressCellButton(object grid, int row, string column)
    {
        try
        {
            _repository.InvokeObjectMethod(grid, "PressButton", row, column);
            _logger.Debug("Pressed button in cell ({Row}, {Column})", row, column);
        }
        catch (Exception ex)
        {
            _logger.Error(ex, "Error pressing cell button at ({Row}, {Column})", row, column);
            throw;
        }
    }

    /// <summary>
    /// Double-clicks a cell in the grid.
    /// </summary>
    /// <param name="grid">The grid COM object.</param>
    /// <param name="row">Row index.</param>
    /// <param name="column">Column name.</param>
    public void DoubleClickCell(object grid, int row, string column)
    {
        try
        {
            _repository.InvokeObjectMethod(grid, "DoubleClick", row, column);
            _logger.Debug("Double-clicked cell ({Row}, {Column})", row, column);
        }
        catch (Exception ex)
        {
            _logger.Error(ex, "Error double-clicking cell at ({Row}, {Column})", row, column);
            throw;
        }
    }
}

