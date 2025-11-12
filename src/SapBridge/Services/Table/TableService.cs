using System.Diagnostics;
using SapBridge.Models;
using SapBridge.Models.Query;
using SapBridge.Repositories;
using SapBridge.Services.Query;
using SapBridge.Utils;
using Serilog;

namespace SapBridge.Services.Table;

/// <summary>
/// Main service for SAP GUI Table operations.
/// Integrates extraction and query capabilities.
/// </summary>
public class TableService : ITableService
{
    private readonly ILogger _logger;
    private readonly ISapGuiRepository _repository;
    private readonly TableDataExtractor _extractor;
    private readonly ConditionEvaluator _conditionEvaluator;

    public TableService(
        ILogger logger,
        ISapGuiRepository repository,
        ConditionEvaluator conditionEvaluator)
    {
        _logger = logger ?? throw new ArgumentNullException(nameof(logger));
        _repository = repository ?? throw new ArgumentNullException(nameof(repository));
        _conditionEvaluator = conditionEvaluator ?? throw new ArgumentNullException(nameof(conditionEvaluator));
        
        _extractor = new TableDataExtractor(_logger, _repository);
    }

    /// <inheritdoc/>
    public async Task<TableData> GetAllDataAsync(string sessionId, string tablePath, int startRow = 0, int? rowCount = null)
    {
        await Task.CompletedTask; // Make it async

        try
        {
            _logger.Information("Getting table data from {TablePath}, startRow={StartRow}, rowCount={RowCount}", 
                tablePath, startRow, rowCount);

            var session = _repository.GetSession(sessionId);
            var table = _repository.FindObjectById(session, tablePath);

            if (table == null)
            {
                throw new InvalidOperationException($"Table not found at path: {tablePath}");
            }

            return _extractor.ExtractData(table, tablePath, startRow, rowCount);
        }
        catch (Exception ex)
        {
            _logger.Error(ex, "Error getting table data");
            var mapped = ComExceptionMapper.MapException(ex, $"getting table data from {tablePath}");
            throw new InvalidOperationException(mapped.Message, ex);
        }
    }

    /// <inheritdoc/>
    public async Task<TableRow> GetRowAsync(string sessionId, string tablePath, int rowIndex)
    {
        await Task.CompletedTask;

        try
        {
            var session = _repository.GetSession(sessionId);
            var table = _repository.FindObjectById(session, tablePath);

            if (table == null)
            {
                throw new InvalidOperationException($"Table not found at path: {tablePath}");
            }

            var row = _extractor.ExtractRow(table, rowIndex);
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
    public async Task<TableCell> GetCellAsync(string sessionId, string tablePath, int row, string column)
    {
        await Task.CompletedTask;

        try
        {
            var session = _repository.GetSession(sessionId);
            var table = _repository.FindObjectById(session, tablePath);

            if (table == null)
            {
                throw new InvalidOperationException($"Table not found at path: {tablePath}");
            }

            var value = _extractor.GetCellValue(table, row, column);

            return new TableCell
            {
                Row = row,
                ColumnName = column,
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
    public async Task SetCellAsync(string sessionId, string tablePath, int row, string column, string value)
    {
        await Task.CompletedTask;

        try
        {
            var session = _repository.GetSession(sessionId);
            var table = _repository.FindObjectById(session, tablePath);

            if (table == null)
            {
                throw new InvalidOperationException($"Table not found at path: {tablePath}");
            }

            _extractor.SetCellValue(table, row, column, value);
            _logger.Information("Set cell ({Row}, {Column}) to '{Value}'", row, column, value);
        }
        catch (Exception ex)
        {
            _logger.Error(ex, "Error setting cell ({Row}, {Column})", row, column);
            throw;
        }
    }

    /// <inheritdoc/>
    public async Task<List<QueryMatch>> FindRowsAsync(string sessionId, string tablePath, List<QueryCondition> conditions)
    {
        await Task.CompletedTask;

        try
        {
            var tableData = await GetAllDataAsync(sessionId, tablePath);
            var matches = new List<QueryMatch>();

            foreach (var row in tableData.Rows)
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
    public async Task<QueryMatch?> FindFirstRowAsync(string sessionId, string tablePath, List<QueryCondition> conditions)
    {
        var matches = await FindRowsAsync(sessionId, tablePath, conditions);
        return matches.FirstOrDefault();
    }

    /// <inheritdoc/>
    public async Task<QueryMatch?> FindLastRowAsync(string sessionId, string tablePath, List<QueryCondition> conditions)
    {
        var matches = await FindRowsAsync(sessionId, tablePath, conditions);
        return matches.LastOrDefault();
    }

    /// <inheritdoc/>
    public async Task<int> GetFirstEmptyRowIndexAsync(string sessionId, string tablePath, List<string>? columns = null)
    {
        await Task.CompletedTask;

        try
        {
            var session = _repository.GetSession(sessionId);
            var table = _repository.FindObjectById(session, tablePath);

            if (table == null)
            {
                throw new InvalidOperationException($"Table not found at path: {tablePath}");
            }

            var rowCount = _extractor.GetRowCount(table);

            for (int i = 0; i < rowCount; i++)
            {
                if (_extractor.IsRowEmpty(table, i, columns))
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
    public async Task<int> GetRowCountAsync(string sessionId, string tablePath)
    {
        await Task.CompletedTask;

        try
        {
            var session = _repository.GetSession(sessionId);
            var table = _repository.FindObjectById(session, tablePath);

            if (table == null)
            {
                throw new InvalidOperationException($"Table not found at path: {tablePath}");
            }

            return _extractor.GetRowCount(table);
        }
        catch (Exception ex)
        {
            _logger.Error(ex, "Error getting row count");
            throw;
        }
    }

    /// <inheritdoc/>
    public async Task<List<string>> GetColumnNamesAsync(string sessionId, string tablePath)
    {
        await Task.CompletedTask;

        try
        {
            var session = _repository.GetSession(sessionId);
            var table = _repository.FindObjectById(session, tablePath);

            if (table == null)
            {
                throw new InvalidOperationException($"Table not found at path: {tablePath}");
            }

            return _extractor.ExtractColumnNames(table);
        }
        catch (Exception ex)
        {
            _logger.Error(ex, "Error getting column names");
            throw;
        }
    }

    /// <inheritdoc/>
    public async Task SelectRowAsync(string sessionId, string tablePath, int rowIndex)
    {
        await Task.CompletedTask;

        try
        {
            var session = _repository.GetSession(sessionId);
            var table = _repository.FindObjectById(session, tablePath);

            if (table == null)
            {
                throw new InvalidOperationException($"Table not found at path: {tablePath}");
            }

            _extractor.SelectRow(table, rowIndex);
        }
        catch (Exception ex)
        {
            _logger.Error(ex, "Error selecting row {RowIndex}", rowIndex);
            throw;
        }
    }

    /// <inheritdoc/>
    public async Task<int> GetSelectedRowAsync(string sessionId, string tablePath)
    {
        await Task.CompletedTask;

        try
        {
            var session = _repository.GetSession(sessionId);
            var table = _repository.FindObjectById(session, tablePath);

            if (table == null)
            {
                throw new InvalidOperationException($"Table not found at path: {tablePath}");
            }

            return _extractor.GetSelectedRow(table);
        }
        catch (Exception ex)
        {
            _logger.Error(ex, "Error getting selected row");
            throw;
        }
    }

    /// <inheritdoc/>
    public async Task<QueryResult> ExecuteQueryAsync(string sessionId, SapQuery query)
    {
        var stopwatch = Stopwatch.StartNew();

        try
        {
            _logger.Information("Executing table query: Action={Action}, Conditions={ConditionCount}", 
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
                
                _ => throw new NotSupportedException($"Query action '{query.Action}' is not supported for tables.")
            };

            stopwatch.Stop();
            return QueryResult.SuccessResult(matches, (int)stopwatch.ElapsedMilliseconds);
        }
        catch (Exception ex)
        {
            stopwatch.Stop();
            _logger.Error(ex, "Error executing table query");
            return QueryResult.FailureResult(ex.Message, (int)stopwatch.ElapsedMilliseconds);
        }
    }
}

