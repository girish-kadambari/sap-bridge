using System.ComponentModel.DataAnnotations;

namespace SapBridge.Models.Query;

/// <summary>
/// Represents a query for finding SAP GUI objects with conditional filtering.
/// Supports complex queries across grids, tables, and trees.
/// </summary>
public class SapQuery
{
    /// <summary>
    /// Base object path (grid, table, or tree path).
    /// </summary>
    [Required]
    public string ObjectPath { get; set; } = string.Empty;

    /// <summary>
    /// Type of object being queried.
    /// </summary>
    [Required]
    public ObjectType Type { get; set; }

    /// <summary>
    /// Action to perform on matched items.
    /// </summary>
    public QueryAction Action { get; set; } = QueryAction.GetFirst;

    /// <summary>
    /// List of conditions to match.
    /// </summary>
    public List<QueryCondition> Conditions { get; set; } = new();

    /// <summary>
    /// Optional query options.
    /// </summary>
    public QueryOptions? Options { get; set; }
}

/// <summary>
/// Type of SAP GUI object.
/// </summary>
public enum ObjectType
{
    Grid,
    Table,
    Tree
}

/// <summary>
/// Action to perform with query results.
/// </summary>
public enum QueryAction
{
    /// <summary>
    /// Get the first matching item.
    /// </summary>
    GetFirst,

    /// <summary>
    /// Get the last matching item.
    /// </summary>
    GetLast,

    /// <summary>
    /// Get all matching items.
    /// </summary>
    GetAll,

    /// <summary>
    /// Count matching items.
    /// </summary>
    Count,

    /// <summary>
    /// Select matching items (for grids/tables).
    /// </summary>
    Select,

    /// <summary>
    /// Extract data from matching items.
    /// </summary>
    Extract
}

/// <summary>
/// Represents a single filter condition.
/// </summary>
public class QueryCondition
{
    /// <summary>
    /// Field/column name to filter on.
    /// </summary>
    [Required]
    public string Field { get; set; } = string.Empty;

    /// <summary>
    /// Comparison operator.
    /// </summary>
    public ConditionOperator Operator { get; set; } = ConditionOperator.Equals;

    /// <summary>
    /// Value to compare against (null for IsEmpty/IsNotEmpty operators).
    /// </summary>
    public object? Value { get; set; }

    /// <summary>
    /// Logical operator to combine with next condition.
    /// </summary>
    public LogicalOperator LogicalOp { get; set; } = LogicalOperator.And;
}

/// <summary>
/// Comparison operators for query conditions.
/// </summary>
public enum ConditionOperator
{
    /// <summary>
    /// Field equals value (exact match).
    /// </summary>
    Equals,

    /// <summary>
    /// Field does not equal value.
    /// </summary>
    NotEquals,

    /// <summary>
    /// Field contains value (substring match).
    /// </summary>
    Contains,

    /// <summary>
    /// Field starts with value.
    /// </summary>
    StartsWith,

    /// <summary>
    /// Field ends with value.
    /// </summary>
    EndsWith,

    /// <summary>
    /// Field is greater than value.
    /// </summary>
    GreaterThan,

    /// <summary>
    /// Field is less than value.
    /// </summary>
    LessThan,

    /// <summary>
    /// Field is greater than or equal to value.
    /// </summary>
    GreaterOrEqual,

    /// <summary>
    /// Field is less than or equal to value.
    /// </summary>
    LessOrEqual,

    /// <summary>
    /// Field is empty or null.
    /// </summary>
    IsEmpty,

    /// <summary>
    /// Field is not empty or null.
    /// </summary>
    IsNotEmpty,

    /// <summary>
    /// Field is null.
    /// </summary>
    IsNull,

    /// <summary>
    /// Field is not null.
    /// </summary>
    IsNotNull
}

/// <summary>
/// Logical operators for combining conditions.
/// </summary>
public enum LogicalOperator
{
    /// <summary>
    /// Logical AND - both conditions must be true.
    /// </summary>
    And,

    /// <summary>
    /// Logical OR - at least one condition must be true.
    /// </summary>
    Or
}

/// <summary>
/// Optional query options for pagination and limits.
/// </summary>
public class QueryOptions
{
    /// <summary>
    /// Maximum number of results to return.
    /// </summary>
    public int? Limit { get; set; }

    /// <summary>
    /// Number of results to skip (for pagination).
    /// </summary>
    public int? Skip { get; set; }

    /// <summary>
    /// Whether to include all columns/fields in results.
    /// </summary>
    public bool IncludeAllFields { get; set; } = true;

    /// <summary>
    /// Specific fields to include (if IncludeAllFields is false).
    /// </summary>
    public List<string>? Fields { get; set; }
}

