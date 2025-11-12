using SapBridge.Models.Query;

namespace SapBridge.Services.Query;

/// <summary>
/// Unified query engine interface for executing queries across grids, tables, and trees.
/// Provides a consistent API for complex conditional filtering.
/// </summary>
public interface IQueryEngine
{
    /// <summary>
    /// Executes a query against a SAP GUI object.
    /// </summary>
    /// <param name="sessionId">The session ID.</param>
    /// <param name="query">The query to execute.</param>
    /// <returns>Query result with matched items.</returns>
    /// <exception cref="ArgumentException">Thrown when query is invalid.</exception>
    /// <exception cref="InvalidOperationException">Thrown when query execution fails.</exception>
    Task<QueryResult> ExecuteAsync(string sessionId, SapQuery query);

    /// <summary>
    /// Finds the first object matching the query conditions.
    /// </summary>
    /// <param name="sessionId">The session ID.</param>
    /// <param name="query">The query to execute.</param>
    /// <returns>The first matched item or null if no matches.</returns>
    Task<QueryMatch?> FindFirstAsync(string sessionId, SapQuery query);

    /// <summary>
    /// Finds the last object matching the query conditions.
    /// </summary>
    /// <param name="sessionId">The session ID.</param>
    /// <param name="query">The query to execute.</param>
    /// <returns>The last matched item or null if no matches.</returns>
    Task<QueryMatch?> FindLastAsync(string sessionId, SapQuery query);

    /// <summary>
    /// Counts objects matching the query conditions.
    /// </summary>
    /// <param name="sessionId">The session ID.</param>
    /// <param name="query">The query to execute.</param>
    /// <returns>Count of matched items.</returns>
    Task<int> CountAsync(string sessionId, SapQuery query);
}

