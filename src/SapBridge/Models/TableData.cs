namespace SapBridge.Models;

/// <summary>
/// Represents data extracted from a SAP GUI Table (GuiTableControl).
/// </summary>
public class TableData
{
    /// <summary>
    /// Table object path.
    /// </summary>
    public string TablePath { get; set; } = string.Empty;

    /// <summary>
    /// Total number of rows in the table.
    /// </summary>
    public int TotalRows { get; set; }

    /// <summary>
    /// Total number of columns in the table.
    /// </summary>
    public int TotalColumns { get; set; }

    /// <summary>
    /// Column names mapped to their indices.
    /// </summary>
    public Dictionary<string, int> ColumnMap { get; set; } = new();

    /// <summary>
    /// Column names in order.
    /// </summary>
    public List<string> ColumnNames { get; set; } = new();

    /// <summary>
    /// Table rows with cell data.
    /// </summary>
    public List<TableRow> Rows { get; set; } = new();

    /// <summary>
    /// Timestamp when data was captured (UTC).
    /// </summary>
    public DateTime CapturedAt { get; set; } = DateTime.UtcNow;
}

/// <summary>
/// Represents a row in a SAP GUI Table.
/// </summary>
public class TableRow
{
    /// <summary>
    /// Row index (zero-based).
    /// </summary>
    public int Index { get; set; }

    /// <summary>
    /// Cell values keyed by column name.
    /// </summary>
    public Dictionary<string, object?> Cells { get; set; } = new();

    /// <summary>
    /// Whether the row is selected.
    /// </summary>
    public bool IsSelected { get; set; }

    /// <summary>
    /// Row key or identifier (if available).
    /// </summary>
    public string? Key { get; set; }
}

/// <summary>
/// Represents a single cell in a table.
/// </summary>
public class TableCell
{
    /// <summary>
    /// Row index.
    /// </summary>
    public int Row { get; set; }

    /// <summary>
    /// Column index.
    /// </summary>
    public int Column { get; set; }

    /// <summary>
    /// Column name.
    /// </summary>
    public string? ColumnName { get; set; }

    /// <summary>
    /// Cell value.
    /// </summary>
    public object? Value { get; set; }

    /// <summary>
    /// Whether the cell is changeable.
    /// </summary>
    public bool Changeable { get; set; }
}

