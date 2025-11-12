using System.Diagnostics;
using SapBridge.Models;
using SapBridge.Models.Query;
using SapBridge.Repositories;
using SapBridge.Services.Query;
using SapBridge.Utils;
using Serilog;
using ILogger = Serilog.ILogger;

namespace SapBridge.Services.Grid;

/// <summary>
/// Main service for SAP GUI Grid operations.
/// Integrates extraction, navigation, and query capabilities.
/// </summary>
public class GridService : IGridService
{
    private readonly ILogger _logger;
    private readonly ISapGuiRepository _repository;
    private readonly GridDataExtractor _extractor;
    private readonly GridNavigator _navigator;
    private readonly ConditionEvaluator _conditionEvaluator;

    public GridService(
        ILogger logger,
        ISapGuiRepository repository,
        ConditionEvaluator conditionEvaluator)
    {
        _logger = logger ?? throw new ArgumentNullException(nameof(logger));
        _repository = repository ?? throw new ArgumentNullException(nameof(repository));
        _conditionEvaluator = conditionEvaluator ?? throw new ArgumentNullException(nameof(conditionEvaluator));
        
        _extractor = new GridDataExtractor(_logger, _repository);
        _navigator = new GridNavigator(_logger, _repository);
    }

    /// <inheritdoc/>
    public async Task<GridData> GetAllDataAsync(string sessionId, string gridPath, int startRow = 0, int? rowCount = null)
    {
        await Task.CompletedTask; // Make it async

        try
        {
            _logger.Information("Getting grid data from {GridPath}, startRow={StartRow}, rowCount={RowCount}", 
                gridPath, startRow, rowCount);

            var session = _repository.GetSession(sessionId);
            var grid = _repository.FindObjectById(session, gridPath);

            if (grid == null)
            {
                throw new InvalidOperationException($"Grid not found at path: {gridPath}");
            }

            return _extractor.ExtractData(grid, gridPath, startRow, rowCount);
        }
        catch (Exception ex)
        {
            _logger.Error(ex, "Error getting grid data");
            var mapped = ComExceptionMapper.MapException(ex, $"getting grid data from {gridPath}");
            throw new InvalidOperationException(mapped.Message, ex);
        }
    }

    /// <inheritdoc/>
    public async Task<GridRow> GetRowAsync(string sessionId, string gridPath, int rowIndex)
    {
        await Task.CompletedTask;

        try
        {
            var session = _repository.GetSession(sessionId);
            var grid = _repository.FindObjectById(session, gridPath);

            if (grid == null)
            {
                throw new InvalidOperationException($"Grid not found at path: {gridPath}");
            }

            var row = _extractor.ExtractRow(grid, rowIndex);
            if (row == null)
            {
                throw new InvalidOperationException($"Could not extract row {rowIndex}");
            }

            return row;
        }
        catch (Exception ex)
        {
            _logger.Error(ex, "Error getting row {RowIndex}", rowIndex);
            throw;
        }
    }

    /// <inheritdoc/>
    public async Task<GridCell> GetCellAsync(string sessionId, string gridPath, int row, string column)
    {
        await Task.CompletedTask;

        try
        {
            var session = _repository.GetSession(sessionId);
            var grid = _repository.FindObjectById(session, gridPath);

            if (grid == null)
            {
                throw new InvalidOperationException($"Grid not found at path: {gridPath}");
            }

            var value = _extractor.GetCellValue(grid, row, column);

            return new GridCell
            {
                Row = row,
                Column = column,
                Value = value
            };
        }
        catch (Exception ex)
        {
            _logger.Error(ex, "Error getting cell ({Row}, {Column})", row, column);
            throw;
        }
    }

    /// <inheritdoc/>
    public async Task SetCellAsync(string sessionId, string gridPath, int row, string column, string value)
    {
        await Task.CompletedTask;

        try
        {
            var session = _repository.GetSession(sessionId);
            var grid = _repository.FindObjectById(session, gridPath);

            if (grid == null)
            {
                throw new InvalidOperationException($"Grid not found at path: {gridPath}");
            }

            _extractor.SetCellValue(grid, row, column, value);
            _logger.Information("Set cell ({Row}, {Column}) to '{Value}'", row, column, value);
        }
        catch (Exception ex)
        {
            _logger.Error(ex, "Error setting cell ({Row}, {Column})", row, column);
            throw;
        }
    }

    /// <inheritdoc/>
    public async Task<List<QueryMatch>> FindRowsAsync(string sessionId, string gridPath, List<QueryCondition> conditions)
    {
        await Task.CompletedTask;

        try
        {
            var gridData = await GetAllDataAsync(sessionId, gridPath);
            var matches = new List<QueryMatch>();

            foreach (var row in gridData.Rows)
            {
                if (_conditionEvaluator.EvaluateConditions(conditions, row.Cells))
                {
                    matches.Add(new QueryMatch
                    {
                        Index = row.Index,
                        Data = row.Cells
                    });
                }
            }

            _logger.Information("Found {MatchCount} rows matching conditions", matches.Count);
            return matches;
        }
        catch (Exception ex)
        {
            _logger.Error(ex, "Error finding rows");
            throw;
        }
    }

    /// <inheritdoc/>
    public async Task<QueryMatch?> FindFirstRowAsync(string sessionId, string gridPath, List<QueryCondition> conditions)
    {
        var matches = await FindRowsAsync(sessionId, gridPath, conditions);
        return matches.FirstOrDefault();
    }

    /// <inheritdoc/>
    public async Task<QueryMatch?> FindLastRowAsync(string sessionId, string gridPath, List<QueryCondition> conditions)
    {
        var matches = await FindRowsAsync(sessionId, gridPath, conditions);
        return matches.LastOrDefault();
    }

    /// <inheritdoc/>
    public async Task<int> GetFirstEmptyRowIndexAsync(string sessionId, string gridPath, List<string>? columns = null)
    {
        await Task.CompletedTask;

        try
        {
            var session = _repository.GetSession(sessionId);
            var grid = _repository.FindObjectById(session, gridPath);

            if (grid == null)
            {
                throw new InvalidOperationException($"Grid not found at path: {gridPath}");
            }

            var rowCount = _extractor.GetRowCount(grid);

            for (int i = 0; i < rowCount; i++)
            {
                if (_extractor.IsRowEmpty(grid, i, columns))
                {
                    _logger.Information("First empty row found at index {Index}", i);
                    return i;
                }
            }

            _logger.Information("No empty row found");
            return -1;
        }
        catch (Exception ex)
        {
            _logger.Error(ex, "Error finding first empty row");
            throw;
        }
    }

    /// <inheritdoc/>
    public async Task<int> GetRowCountAsync(string sessionId, string gridPath)
    {
        await Task.CompletedTask;

        try
        {
            var session = _repository.GetSession(sessionId);
            var grid = _repository.FindObjectById(session, gridPath);

            if (grid == null)
            {
                throw new InvalidOperationException($"Grid not found at path: {gridPath}");
            }

            return _extractor.GetRowCount(grid);
        }
        catch (Exception ex)
        {
            _logger.Error(ex, "Error getting row count");
            throw;
        }
    }

    /// <inheritdoc/>
    public async Task<List<GridColumn>> GetColumnsAsync(string sessionId, string gridPath)
    {
        await Task.CompletedTask;

        try
        {
            var session = _repository.GetSession(sessionId);
            var grid = _repository.FindObjectById(session, gridPath);

            if (grid == null)
            {
                throw new InvalidOperationException($"Grid not found at path: {gridPath}");
            }

            return _extractor.ExtractColumns(grid);
        }
        catch (Exception ex)
        {
            _logger.Error(ex, "Error getting columns");
            throw;
        }
    }

    /// <inheritdoc/>
    public async Task SelectRowsAsync(string sessionId, string gridPath, List<int> rowIndices)
    {
        await Task.CompletedTask;

        try
        {
            var session = _repository.GetSession(sessionId);
            var grid = _repository.FindObjectById(session, gridPath);

            if (grid == null)
            {
                throw new InvalidOperationException($"Grid not found at path: {gridPath}");
            }

            _navigator.SelectRows(grid, rowIndices);
        }
        catch (Exception ex)
        {
            _logger.Error(ex, "Error selecting rows");
            throw;
        }
    }

    /// <inheritdoc/>
    public async Task<List<int>> GetSelectedRowsAsync(string sessionId, string gridPath)
    {
        await Task.CompletedTask;

        try
        {
            var session = _repository.GetSession(sessionId);
            var grid = _repository.FindObjectById(session, gridPath);

            if (grid == null)
            {
                throw new InvalidOperationException($"Grid not found at path: {gridPath}");
            }

            return _navigator.GetSelectedRows(grid);
        }
        catch (Exception ex)
        {
            _logger.Error(ex, "Error getting selected rows");
            throw;
        }
    }

    /// <inheritdoc/>
    public async Task ScrollToRowAsync(string sessionId, string gridPath, int rowIndex)
    {
        await Task.CompletedTask;

        try
        {
            var session = _repository.GetSession(sessionId);
            var grid = _repository.FindObjectById(session, gridPath);

            if (grid == null)
            {
                throw new InvalidOperationException($"Grid not found at path: {gridPath}");
            }

            _navigator.ScrollToRow(grid, rowIndex);
        }
        catch (Exception ex)
        {
            _logger.Error(ex, "Error scrolling to row {RowIndex}", rowIndex);
            throw;
        }
    }

    /// <inheritdoc/>
    public async Task<QueryResult> ExecuteQueryAsync(string sessionId, SapQuery query)
    {
        var stopwatch = Stopwatch.StartNew();

        try
        {
            _logger.Information("Executing grid query: Action={Action}, Conditions={ConditionCount}", 
                query.Action, query.Conditions.Count);

            List<QueryMatch> matches = query.Action switch
            {
                QueryAction.GetFirst => new List<QueryMatch>
                {
                    (await FindFirstRowAsync(sessionId, query.ObjectPath, query.Conditions))!
                }.Where(m => m != null).ToList(),
                
                QueryAction.GetLast => new List<QueryMatch>
                {
                    (await FindLastRowAsync(sessionId, query.ObjectPath, query.Conditions))!
                }.Where(m => m != null).ToList(),
                
                QueryAction.GetAll => await FindRowsAsync(sessionId, query.ObjectPath, query.Conditions),
                
                QueryAction.Count => await FindRowsAsync(sessionId, query.ObjectPath, query.Conditions),
                
                _ => throw new NotSupportedException($"Query action '{query.Action}' is not supported for grids.")
            };

            stopwatch.Stop();
            return QueryResult.SuccessResult(matches, (int)stopwatch.ElapsedMilliseconds);
        }
        catch (Exception ex)
        {
            stopwatch.Stop();
            _logger.Error(ex, "Error executing grid query");
            return QueryResult.FailureResult(ex.Message, (int)stopwatch.ElapsedMilliseconds);
        }
    }
}

