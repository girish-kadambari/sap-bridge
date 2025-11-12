namespace SapBridge.Models;

/// <summary>
/// Represents the result of an action execution.
/// </summary>
public class ActionResult
{
    /// <summary>
    /// Whether the action executed successfully.
    /// </summary>
    public bool Success { get; set; }

    /// <summary>
    /// Result value returned by the action (if any).
    /// </summary>
    public object? Result { get; set; }

    /// <summary>
    /// Error message if the action failed.
    /// </summary>
    public string? Error { get; set; }

    /// <summary>
    /// Error category for classification.
    /// </summary>
    public string? ErrorCategory { get; set; }

    /// <summary>
    /// Suggestion for resolving the error.
    /// </summary>
    public string? Suggestion { get; set; }

    /// <summary>
    /// Execution time in milliseconds.
    /// </summary>
    public int ExecutionTimeMs { get; set; }

    /// <summary>
    /// Timestamp when the action was executed (UTC).
    /// </summary>
    public DateTime ExecutedAt { get; set; } = DateTime.UtcNow;

    /// <summary>
    /// Creates a successful action result.
    /// </summary>
    public static ActionResult SuccessResult(object? result = null, int executionTimeMs = 0)
    {
        return new ActionResult
        {
            Success = true,
            Result = result,
            ExecutionTimeMs = executionTimeMs
        };
    }

    /// <summary>
    /// Creates a failed action result.
    /// </summary>
    public static ActionResult FailureResult(string error, string? category = null, string? suggestion = null, int executionTimeMs = 0)
    {
        return new ActionResult
        {
            Success = false,
            Error = error,
            ErrorCategory = category,
            Suggestion = suggestion,
            ExecutionTimeMs = executionTimeMs
        };
    }
}

