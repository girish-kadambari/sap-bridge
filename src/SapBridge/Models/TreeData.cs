namespace SapBridge.Models;

/// <summary>
/// Represents data extracted from a SAP GUI Tree (GuiTree).
/// </summary>
public class TreeData
{
    /// <summary>
    /// Tree object path.
    /// </summary>
    public string TreePath { get; set; } = string.Empty;

    /// <summary>
    /// Root nodes of the tree.
    /// </summary>
    public List<TreeNode> RootNodes { get; set; } = new();

    /// <summary>
    /// Total number of nodes in the tree.
    /// </summary>
    public int TotalNodes { get; set; }

    /// <summary>
    /// Timestamp when data was captured (UTC).
    /// </summary>
    public DateTime CapturedAt { get; set; } = DateTime.UtcNow;

    /// <summary>
    /// Whether the tree structure is fully loaded.
    /// </summary>
    public bool IsFullyLoaded { get; set; }
}

/// <summary>
/// Represents a node in a SAP GUI Tree.
/// </summary>
public class TreeNode
{
    /// <summary>
    /// Unique node key/ID.
    /// </summary>
    public string Key { get; set; } = string.Empty;

    /// <summary>
    /// Node text/label.
    /// </summary>
    public string Text { get; set; } = string.Empty;

    /// <summary>
    /// Parent node key (null for root nodes).
    /// </summary>
    public string? ParentKey { get; set; }

    /// <summary>
    /// Child nodes.
    /// </summary>
    public List<TreeNode> Children { get; set; } = new();

    /// <summary>
    /// Whether the node is expanded.
    /// </summary>
    public bool IsExpanded { get; set; }

    /// <summary>
    /// Whether the node is selected.
    /// </summary>
    public bool IsSelected { get; set; }

    /// <summary>
    /// Whether the node has children (even if not loaded).
    /// </summary>
    public bool HasChildren { get; set; }

    /// <summary>
    /// Node depth/level in the tree (0 for root).
    /// </summary>
    public int Level { get; set; }

    /// <summary>
    /// Node icon/image name (if available).
    /// </summary>
    public string? Icon { get; set; }

    /// <summary>
    /// Additional node properties/attributes.
    /// </summary>
    public Dictionary<string, object?> Properties { get; set; } = new();

    /// <summary>
    /// Whether the node is a leaf (has no children and cannot have children).
    /// </summary>
    public bool IsLeaf => !HasChildren && Children.Count == 0;
}

/// <summary>
/// Represents the result of a tree search operation.
/// </summary>
public class TreeSearchResult
{
    /// <summary>
    /// The found node.
    /// </summary>
    public TreeNode? Node { get; set; }

    /// <summary>
    /// The path from root to the found node (list of node keys).
    /// </summary>
    public List<string> Path { get; set; } = new();

    /// <summary>
    /// Whether the search was successful.
    /// </summary>
    public bool Found => Node != null;
}

