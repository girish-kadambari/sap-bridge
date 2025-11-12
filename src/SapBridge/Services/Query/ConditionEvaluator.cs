using SapBridge.Models.Query;
using Serilog;

namespace SapBridge.Services.Query;

/// <summary>
/// Evaluates query conditions against data values.
/// Supports all comparison operators and type conversions.
/// </summary>
public class ConditionEvaluator
{
    private readonly ILogger _logger;

    public ConditionEvaluator(ILogger logger)
    {
        _logger = logger ?? throw new ArgumentNullException(nameof(logger));
    }

    /// <summary>
    /// Evaluates a single condition against a data value.
    /// </summary>
    /// <param name="condition">The condition to evaluate.</param>
    /// <param name="fieldValue">The actual field value.</param>
    /// <returns>True if the condition is met, false otherwise.</returns>
    public bool EvaluateCondition(QueryCondition condition, object? fieldValue)
    {
        if (condition == null)
            throw new ArgumentNullException(nameof(condition));

        try
        {
            return condition.Operator switch
            {
                ConditionOperator.Equals => AreEqual(fieldValue, condition.Value),
                ConditionOperator.NotEquals => !AreEqual(fieldValue, condition.Value),
                ConditionOperator.Contains => Contains(fieldValue, condition.Value),
                ConditionOperator.StartsWith => StartsWith(fieldValue, condition.Value),
                ConditionOperator.EndsWith => EndsWith(fieldValue, condition.Value),
                ConditionOperator.GreaterThan => IsGreaterThan(fieldValue, condition.Value),
                ConditionOperator.LessThan => IsLessThan(fieldValue, condition.Value),
                ConditionOperator.GreaterOrEqual => IsGreaterOrEqual(fieldValue, condition.Value),
                ConditionOperator.LessOrEqual => IsLessOrEqual(fieldValue, condition.Value),
                ConditionOperator.IsEmpty => IsEmpty(fieldValue),
                ConditionOperator.IsNotEmpty => !IsEmpty(fieldValue),
                ConditionOperator.IsNull => fieldValue == null,
                ConditionOperator.IsNotNull => fieldValue != null,
                _ => throw new NotSupportedException($"Operator '{condition.Operator}' is not supported.")
            };
        }
        catch (Exception ex)
        {
            _logger.Warning(ex, "Error evaluating condition for field {Field} with operator {Operator}", 
                condition.Field, condition.Operator);
            return false;
        }
    }

    /// <summary>
    /// Evaluates a list of conditions with logical operators.
    /// </summary>
    /// <param name="conditions">The list of conditions.</param>
    /// <param name="dataRow">Dictionary of field names to values.</param>
    /// <returns>True if all conditions are met according to logical operators.</returns>
    public bool EvaluateConditions(List<QueryCondition> conditions, Dictionary<string, object?> dataRow)
    {
        if (conditions == null || conditions.Count == 0)
            return true; // No conditions means match all

        if (dataRow == null)
            return false;

        bool result = true;
        bool isFirstCondition = true;

        for (int i = 0; i < conditions.Count; i++)
        {
            var condition = conditions[i];
            
            // Get the field value from data row
            dataRow.TryGetValue(condition.Field, out var fieldValue);
            
            // Evaluate this condition
            bool conditionResult = EvaluateCondition(condition, fieldValue);

            if (isFirstCondition)
            {
                result = conditionResult;
                isFirstCondition = false;
            }
            else
            {
                // Get the logical operator from the PREVIOUS condition
                var previousLogicalOp = i > 0 ? conditions[i - 1].LogicalOp : LogicalOperator.And;
                
                result = previousLogicalOp == LogicalOperator.And
                    ? result && conditionResult
                    : result || conditionResult;
            }
        }

        return result;
    }

    /// <summary>
    /// Checks if two values are equal with type conversion.
    /// </summary>
    private bool AreEqual(object? value1, object? value2)
    {
        if (value1 == null && value2 == null)
            return true;
        
        if (value1 == null || value2 == null)
            return false;

        // Try direct comparison first
        if (value1.Equals(value2))
            return true;

        // Try string comparison (case-insensitive)
        var str1 = value1.ToString() ?? "";
        var str2 = value2.ToString() ?? "";
        return str1.Equals(str2, StringComparison.OrdinalIgnoreCase);
    }

    /// <summary>
    /// Checks if value1 contains value2 (string operation).
    /// </summary>
    private bool Contains(object? value1, object? value2)
    {
        if (value1 == null || value2 == null)
            return false;

        var str1 = value1.ToString() ?? "";
        var str2 = value2.ToString() ?? "";
        return str1.Contains(str2, StringComparison.OrdinalIgnoreCase);
    }

    /// <summary>
    /// Checks if value1 starts with value2 (string operation).
    /// </summary>
    private bool StartsWith(object? value1, object? value2)
    {
        if (value1 == null || value2 == null)
            return false;

        var str1 = value1.ToString() ?? "";
        var str2 = value2.ToString() ?? "";
        return str1.StartsWith(str2, StringComparison.OrdinalIgnoreCase);
    }

    /// <summary>
    /// Checks if value1 ends with value2 (string operation).
    /// </summary>
    private bool EndsWith(object? value1, object? value2)
    {
        if (value1 == null || value2 == null)
            return false;

        var str1 = value1.ToString() ?? "";
        var str2 = value2.ToString() ?? "";
        return str1.EndsWith(str2, StringComparison.OrdinalIgnoreCase);
    }

    /// <summary>
    /// Checks if value1 is greater than value2 (numeric/date comparison).
    /// </summary>
    private bool IsGreaterThan(object? value1, object? value2)
    {
        return CompareValues(value1, value2) > 0;
    }

    /// <summary>
    /// Checks if value1 is less than value2 (numeric/date comparison).
    /// </summary>
    private bool IsLessThan(object? value1, object? value2)
    {
        return CompareValues(value1, value2) < 0;
    }

    /// <summary>
    /// Checks if value1 is greater than or equal to value2.
    /// </summary>
    private bool IsGreaterOrEqual(object? value1, object? value2)
    {
        return CompareValues(value1, value2) >= 0;
    }

    /// <summary>
    /// Checks if value1 is less than or equal to value2.
    /// </summary>
    private bool IsLessOrEqual(object? value1, object? value2)
    {
        return CompareValues(value1, value2) <= 0;
    }

    /// <summary>
    /// Compares two values numerically or lexicographically.
    /// </summary>
    /// <returns>-1 if value1 &lt; value2, 0 if equal, 1 if value1 &gt; value2.</returns>
    private int CompareValues(object? value1, object? value2)
    {
        if (value1 == null && value2 == null)
            return 0;
        if (value1 == null)
            return -1;
        if (value2 == null)
            return 1;

        // Try numeric comparison
        if (TryParseNumeric(value1, out double num1) && TryParseNumeric(value2, out double num2))
        {
            return num1.CompareTo(num2);
        }

        // Try date comparison
        if (DateTime.TryParse(value1.ToString(), out DateTime date1) && 
            DateTime.TryParse(value2.ToString(), out DateTime date2))
        {
            return date1.CompareTo(date2);
        }

        // Fall back to string comparison
        var str1 = value1.ToString() ?? "";
        var str2 = value2.ToString() ?? "";
        return string.Compare(str1, str2, StringComparison.OrdinalIgnoreCase);
    }

    /// <summary>
    /// Checks if a value is empty (null, empty string, or whitespace).
    /// </summary>
    private bool IsEmpty(object? value)
    {
        if (value == null)
            return true;

        var str = value.ToString();
        return string.IsNullOrWhiteSpace(str);
    }

    /// <summary>
    /// Tries to parse a value as a numeric type.
    /// </summary>
    private bool TryParseNumeric(object value, out double result)
    {
        result = 0;
        
        if (value is int intVal)
        {
            result = intVal;
            return true;
        }
        if (value is double doubleVal)
        {
            result = doubleVal;
            return true;
        }
        if (value is decimal decimalVal)
        {
            result = (double)decimalVal;
            return true;
        }
        if (value is float floatVal)
        {
            result = floatVal;
            return true;
        }

        return double.TryParse(value.ToString(), out result);
    }
}

