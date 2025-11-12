namespace SapBridge.Models.Query;

/// <summary>
/// Represents the result of a query execution.
/// </summary>
public class QueryResult
{
    /// <summary>
    /// Whether the query executed successfully.
    /// </summary>
    public bool Success { get; set; }

    /// <summary>
    /// Total number of items that matched the query conditions.
    /// </summary>
    public int TotalMatches { get; set; }

    /// <summary>
    /// List of matched items with their data.
    /// </summary>
    public List<QueryMatch> Matches { get; set; } = new();

    /// <summary>
    /// Error message if the query failed.
    /// </summary>
    public string? Error { get; set; }

    /// <summary>
    /// Query execution time in milliseconds.
    /// </summary>
    public int ExecutionTimeMs { get; set; }

    /// <summary>
    /// Creates a successful query result.
    /// </summary>
    public static QueryResult SuccessResult(List<QueryMatch> matches, int executionTimeMs)
    {
        return new QueryResult
        {
            Success = true,
            TotalMatches = matches.Count,
            Matches = matches,
            ExecutionTimeMs = executionTimeMs
        };
    }

    /// <summary>
    /// Creates a failed query result.
    /// </summary>
    public static QueryResult FailureResult(string error, int executionTimeMs = 0)
    {
        return new QueryResult
        {
            Success = false,
            Error = error,
            ExecutionTimeMs = executionTimeMs
        };
    }
}

/// <summary>
/// Represents a single match from a query.
/// </summary>
public class QueryMatch
{
    /// <summary>
    /// Index of the matched item (row index for grids/tables, node index for trees).
    /// </summary>
    public int Index { get; set; }

    /// <summary>
    /// Key or identifier of the matched item (e.g., node key for trees).
    /// </summary>
    public string? Key { get; set; }

    /// <summary>
    /// Field values from the matched item.
    /// Key = field/column name, Value = field value.
    /// </summary>
    public Dictionary<string, object?> Data { get; set; } = new();

    /// <summary>
    /// Full path to the matched object (if applicable).
    /// </summary>
    public string? ObjectPath { get; set; }

    /// <summary>
    /// Additional metadata about the matched item.
    /// </summary>
    public Dictionary<string, object?>? Metadata { get; set; }
}

