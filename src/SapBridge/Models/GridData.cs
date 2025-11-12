namespace SapBridge.Models;

/// <summary>
/// Represents data extracted from a SAP GUI Grid (GuiGridView).
/// </summary>
public class GridData
{
    /// <summary>
    /// Grid object path.
    /// </summary>
    public string GridPath { get; set; } = string.Empty;

    /// <summary>
    /// Total number of rows in the grid.
    /// </summary>
    public int TotalRows { get; set; }

    /// <summary>
    /// Total number of columns in the grid.
    /// </summary>
    public int TotalColumns { get; set; }

    /// <summary>
    /// Column definitions.
    /// </summary>
    public List<GridColumn> Columns { get; set; } = new();

    /// <summary>
    /// Grid rows with cell data.
    /// </summary>
    public List<GridRow> Rows { get; set; } = new();

    /// <summary>
    /// Currently selected row indices.
    /// </summary>
    public List<int> SelectedRows { get; set; } = new();

    /// <summary>
    /// Current vertical scroll position.
    /// </summary>
    public int? CurrentScrollPosition { get; set; }

    /// <summary>
    /// Timestamp when data was captured (UTC).
    /// </summary>
    public DateTime CapturedAt { get; set; } = DateTime.UtcNow;
}

/// <summary>
/// Represents a column in a SAP GUI Grid.
/// </summary>
public class GridColumn
{
    /// <summary>
    /// Column name/ID.
    /// </summary>
    public string Name { get; set; } = string.Empty;

    /// <summary>
    /// Column title/header text.
    /// </summary>
    public string Title { get; set; } = string.Empty;

    /// <summary>
    /// Column index (zero-based).
    /// </summary>
    public int Index { get; set; }

    /// <summary>
    /// Column width in pixels.
    /// </summary>
    public int Width { get; set; }

    /// <summary>
    /// Whether the column is visible.
    /// </summary>
    public bool Visible { get; set; } = true;

    /// <summary>
    /// Whether the column is editable.
    /// </summary>
    public bool Editable { get; set; }

    /// <summary>
    /// Column data type (Text, Number, Date, etc.).
    /// </summary>
    public string? DataType { get; set; }
}

/// <summary>
/// Represents a row in a SAP GUI Grid.
/// </summary>
public class GridRow
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
    /// Row metadata (colors, icons, etc.).
    /// </summary>
    public Dictionary<string, object?>? Metadata { get; set; }
}

/// <summary>
/// Represents a single cell in a grid.
/// </summary>
public class GridCell
{
    /// <summary>
    /// Row index.
    /// </summary>
    public int Row { get; set; }

    /// <summary>
    /// Column name.
    /// </summary>
    public string Column { get; set; } = string.Empty;

    /// <summary>
    /// Cell value.
    /// </summary>
    public object? Value { get; set; }

    /// <summary>
    /// Whether the cell is editable.
    /// </summary>
    public bool Editable { get; set; }

    /// <summary>
    /// Cell tooltip text.
    /// </summary>
    public string? Tooltip { get; set; }
}

