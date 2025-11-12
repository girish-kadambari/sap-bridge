using System.Diagnostics;
using SapBridge.Models.Query;
using SapBridge.Services.Grid;
using SapBridge.Services.Table;
using SapBridge.Services.Tree;
using Serilog;
using ILogger = Serilog.ILogger;

namespace SapBridge.Services.Query;

/// <summary>
/// Unified query engine for executing queries across grids, tables, and trees.
/// Routes queries to the appropriate service based on object type.
/// </summary>
public class QueryEngine : IQueryEngine
{
    private readonly ILogger _logger;
    private readonly QueryValidator _validator;
    private readonly IGridService? _gridService;
    private readonly ITableService? _tableService;
    private readonly ITreeService? _treeService;

    public QueryEngine(
        ILogger logger,
        QueryValidator validator,
        IGridService? gridService = null,
        ITableService? tableService = null,
        ITreeService? treeService = null)
    {
        _logger = logger ?? throw new ArgumentNullException(nameof(logger));
        _validator = validator ?? throw new ArgumentNullException(nameof(validator));
        _gridService = gridService;
        _tableService = tableService;
        _treeService = treeService;
    }

    /// <inheritdoc/>
    public async Task<QueryResult> ExecuteAsync(string sessionId, SapQuery query)
    {
        var stopwatch = Stopwatch.StartNew();

        try
        {
            _logger.Information("Executing query: Type={Type}, Action={Action}, Path={Path}", 
                query.Type, query.Action, query.ObjectPath);

            // Validate the query
            var validationResult = _validator.Validate(query);
            if (!validationResult.IsValid)
            {
                var errorMessage = validationResult.ErrorMessage ?? "Validation failed";
                _logger.Warning("Query validation failed: {Errors}", errorMessage);
                stopwatch.Stop();
                return QueryResult.FailureResult(
                    $"Query validation failed: {errorMessage}", 
                    (int)stopwatch.ElapsedMilliseconds);
            }

            // Route to the appropriate service
            QueryResult result = query.Type switch
            {
                ObjectType.Grid => await ExecuteGridQueryAsync(sessionId, query),
                ObjectType.Table => await ExecuteTableQueryAsync(sessionId, query),
                ObjectType.Tree => await ExecuteTreeQueryAsync(sessionId, query),
                _ => throw new NotSupportedException($"Query type '{query.Type}' is not supported.")
            };

            stopwatch.Stop();
            _logger.Information("Query executed in {ElapsedMs}ms, found {MatchCount} matches", 
                stopwatch.ElapsedMilliseconds, result.Matches.Count);

            return result;
        }
        catch (Exception ex)
        {
            stopwatch.Stop();
            _logger.Error(ex, "Error executing query");
            return QueryResult.FailureResult(ex.Message, (int)stopwatch.ElapsedMilliseconds);
        }
    }

    /// <inheritdoc/>
    public async Task<QueryMatch?> FindFirstAsync(string sessionId, SapQuery query)
    {
        query.Action = QueryAction.GetFirst;
        var result = await ExecuteAsync(sessionId, query);
        return result.Success ? result.Matches.FirstOrDefault() : null;
    }

    /// <inheritdoc/>
    public async Task<QueryMatch?> FindLastAsync(string sessionId, SapQuery query)
    {
        query.Action = QueryAction.GetLast;
        var result = await ExecuteAsync(sessionId, query);
        return result.Success ? result.Matches.LastOrDefault() : null;
    }

    /// <inheritdoc/>
    public async Task<int> CountAsync(string sessionId, SapQuery query)
    {
        query.Action = QueryAction.Count;
        var result = await ExecuteAsync(sessionId, query);
        return result.Success ? result.Matches.Count : 0;
    }

    /// <summary>
    /// Executes a query against a grid.
    /// </summary>
    private async Task<QueryResult> ExecuteGridQueryAsync(string sessionId, SapQuery query)
    {
        if (_gridService == null)
        {
            throw new InvalidOperationException("GridService is not available. Please ensure it is registered in DI.");
        }

        return await _gridService.ExecuteQueryAsync(sessionId, query);
    }

    /// <summary>
    /// Executes a query against a table.
    /// </summary>
    private async Task<QueryResult> ExecuteTableQueryAsync(string sessionId, SapQuery query)
    {
        if (_tableService == null)
        {
            throw new InvalidOperationException("TableService is not available. Please ensure it is registered in DI.");
        }

        return await _tableService.ExecuteQueryAsync(sessionId, query);
    }

    /// <summary>
    /// Executes a query against a tree.
    /// </summary>
    private async Task<QueryResult> ExecuteTreeQueryAsync(string sessionId, SapQuery query)
    {
        if (_treeService == null)
        {
            throw new InvalidOperationException("TreeService is not available. Please ensure it is registered in DI.");
        }

        return await _treeService.ExecuteQueryAsync(sessionId, query);
    }
}

