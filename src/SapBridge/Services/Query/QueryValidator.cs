using SapBridge.Models.Query;
using Serilog;
using ILogger = Serilog.ILogger;

namespace SapBridge.Services.Query;

/// <summary>
/// Validates query requests before execution.
/// Ensures queries are well-formed and valid.
/// </summary>
public class QueryValidator
{
    private readonly ILogger _logger;

    public QueryValidator(ILogger logger)
    {
        _logger = logger ?? throw new ArgumentNullException(nameof(logger));
    }

    /// <summary>
    /// Validates a query request.
    /// </summary>
    /// <param name="query">The query to validate.</param>
    /// <returns>Validation result with any error messages.</returns>
    public ValidationResult Validate(SapQuery query)
    {
        if (query == null)
            return ValidationResult.Failure("Query cannot be null.");

        var errors = new List<string>();

        // Validate object path
        if (string.IsNullOrWhiteSpace(query.ObjectPath))
        {
            errors.Add("Object path is required.");
        }

        // Validate conditions
        if (query.Conditions != null && query.Conditions.Count > 0)
        {
            for (int i = 0; i < query.Conditions.Count; i++)
            {
                var condition = query.Conditions[i];
                var conditionErrors = ValidateCondition(condition, i);
                errors.AddRange(conditionErrors);
            }
        }

        // Validate options
        if (query.Options != null)
        {
            var optionErrors = ValidateOptions(query.Options);
            errors.AddRange(optionErrors);
        }

        if (errors.Count > 0)
        {
            _logger.Warning("Query validation failed: {Errors}", string.Join("; ", errors));
            return ValidationResult.Failure(string.Join(" ", errors));
        }

        return ValidationResult.Success();
    }

    /// <summary>
    /// Validates a single query condition.
    /// </summary>
    private List<string> ValidateCondition(QueryCondition condition, int index)
    {
        var errors = new List<string>();

        if (string.IsNullOrWhiteSpace(condition.Field))
        {
            errors.Add($"Condition {index}: Field name is required.");
        }

        // Validate that value-based operators have a value
        if (RequiresValue(condition.Operator) && condition.Value == null)
        {
            errors.Add($"Condition {index}: Operator '{condition.Operator}' requires a value.");
        }

        // Validate that non-value operators don't have a value
        if (!RequiresValue(condition.Operator) && condition.Value != null)
        {
            _logger.Debug("Condition {Index}: Operator '{Operator}' does not use value, but value was provided.", 
                index, condition.Operator);
        }

        return errors;
    }

    /// <summary>
    /// Validates query options.
    /// </summary>
    private List<string> ValidateOptions(QueryOptions options)
    {
        var errors = new List<string>();

        if (options.Limit.HasValue && options.Limit.Value < 0)
        {
            errors.Add("Limit must be a positive number.");
        }

        if (options.Skip.HasValue && options.Skip.Value < 0)
        {
            errors.Add("Skip must be a positive number.");
        }

        if (!options.IncludeAllFields && (options.Fields == null || options.Fields.Count == 0))
        {
            errors.Add("When IncludeAllFields is false, Fields list must be provided.");
        }

        return errors;
    }

    /// <summary>
    /// Checks if an operator requires a value.
    /// </summary>
    private bool RequiresValue(ConditionOperator op)
    {
        return op switch
        {
            ConditionOperator.IsEmpty => false,
            ConditionOperator.IsNotEmpty => false,
            ConditionOperator.IsNull => false,
            ConditionOperator.IsNotNull => false,
            _ => true
        };
    }
}

/// <summary>
/// Result of query validation.
/// </summary>
public class ValidationResult
{
    /// <summary>
    /// Whether the validation succeeded.
    /// </summary>
    public bool IsValid { get; set; }

    /// <summary>
    /// Error message if validation failed.
    /// </summary>
    public string? ErrorMessage { get; set; }

    /// <summary>
    /// Creates a successful validation result.
    /// </summary>
    public static ValidationResult Success()
    {
        return new ValidationResult { IsValid = true };
    }

    /// <summary>
    /// Creates a failed validation result.
    /// </summary>
    public static ValidationResult Failure(string errorMessage)
    {
        return new ValidationResult
        {
            IsValid = false,
            ErrorMessage = errorMessage
        };
    }
}

