using SapBridge.Models;
using SapBridge.Models.Query;

namespace SapBridge.Services.Tree;

/// <summary>
/// Service interface for SAP GUI Tree (GuiTree) operations.
/// Supports tree traversal, node operations, and query-based filtering.
/// </summary>
public interface ITreeService
{
    /// <summary>
    /// Gets the complete tree structure.
    /// </summary>
    /// <param name="sessionId">The session ID.</param>
    /// <param name="treePath">Path to the tree object.</param>
    /// <param name="expandAll">Whether to expand all nodes before extracting.</param>
    /// <returns>Tree data with all nodes.</returns>
    Task<TreeData> GetTreeDataAsync(string sessionId, string treePath, bool expandAll = false);

    /// <summary>
    /// Gets a specific node by its key.
    /// </summary>
    /// <param name="sessionId">The session ID.</param>
    /// <param name="treePath">Path to the tree object.</param>
    /// <param name="nodeKey">The node key/ID.</param>
    /// <returns>The tree node.</returns>
    Task<TreeNode> GetNodeAsync(string sessionId, string treePath, string nodeKey);

    /// <summary>
    /// Gets all root nodes.
    /// </summary>
    /// <param name="sessionId">The session ID.</param>
    /// <param name="treePath">Path to the tree object.</param>
    /// <returns>List of root nodes.</returns>
    Task<List<TreeNode>> GetRootNodesAsync(string sessionId, string treePath);

    /// <summary>
    /// Gets children of a specific node.
    /// </summary>
    /// <param name="sessionId">The session ID.</param>
    /// <param name="treePath">Path to the tree object.</param>
    /// <param name="parentKey">Parent node key.</param>
    /// <returns>List of child nodes.</returns>
    Task<List<TreeNode>> GetChildrenAsync(string sessionId, string treePath, string parentKey);

    /// <summary>
    /// Expands a node to show its children.
    /// </summary>
    /// <param name="sessionId">The session ID.</param>
    /// <param name="treePath">Path to the tree object.</param>
    /// <param name="nodeKey">Node key to expand.</param>
    Task ExpandNodeAsync(string sessionId, string treePath, string nodeKey);

    /// <summary>
    /// Collapses a node to hide its children.
    /// </summary>
    /// <param name="sessionId">The session ID.</param>
    /// <param name="treePath">Path to the tree object.</param>
    /// <param name="nodeKey">Node key to collapse.</param>
    Task CollapseNodeAsync(string sessionId, string treePath, string nodeKey);

    /// <summary>
    /// Selects a node in the tree.
    /// </summary>
    /// <param name="sessionId">The session ID.</param>
    /// <param name="treePath">Path to the tree object.</param>
    /// <param name="nodeKey">Node key to select.</param>
    Task SelectNodeAsync(string sessionId, string treePath, string nodeKey);

    /// <summary>
    /// Gets the currently selected node key.
    /// </summary>
    /// <param name="sessionId">The session ID.</param>
    /// <param name="treePath">Path to the tree object.</param>
    /// <returns>Selected node key or null if none selected.</returns>
    Task<string?> GetSelectedNodeKeyAsync(string sessionId, string treePath);

    /// <summary>
    /// Finds nodes matching query conditions.
    /// </summary>
    /// <param name="sessionId">The session ID.</param>
    /// <param name="treePath">Path to the tree object.</param>
    /// <param name="conditions">Query conditions to match.</param>
    /// <param name="searchInText">Whether to search in node text.</param>
    /// <returns>List of matching nodes as query matches.</returns>
    Task<List<QueryMatch>> FindNodesAsync(string sessionId, string treePath, List<QueryCondition> conditions, bool searchInText = true);

    /// <summary>
    /// Finds the first node matching query conditions.
    /// </summary>
    /// <param name="sessionId">The session ID.</param>
    /// <param name="treePath">Path to the tree object.</param>
    /// <param name="conditions">Query conditions to match.</param>
    /// <param name="searchInText">Whether to search in node text.</param>
    /// <returns>First matching node or null.</returns>
    Task<QueryMatch?> FindFirstNodeAsync(string sessionId, string treePath, List<QueryCondition> conditions, bool searchInText = true);

    /// <summary>
    /// Searches for a node by text.
    /// </summary>
    /// <param name="sessionId">The session ID.</param>
    /// <param name="treePath">Path to the tree object.</param>
    /// <param name="searchText">Text to search for.</param>
    /// <param name="caseSensitive">Whether search is case-sensitive.</param>
    /// <returns>Search result with found node and path.</returns>
    Task<TreeSearchResult> SearchNodeByTextAsync(string sessionId, string treePath, string searchText, bool caseSensitive = false);

    /// <summary>
    /// Gets the path from root to a specific node.
    /// </summary>
    /// <param name="sessionId">The session ID.</param>
    /// <param name="treePath">Path to the tree object.</param>
    /// <param name="nodeKey">Target node key.</param>
    /// <returns>List of node keys from root to target.</returns>
    Task<List<string>> GetNodePathAsync(string sessionId, string treePath, string nodeKey);

    /// <summary>
    /// Double-clicks a node (typically to trigger an action).
    /// </summary>
    /// <param name="sessionId">The session ID.</param>
    /// <param name="treePath">Path to the tree object.</param>
    /// <param name="nodeKey">Node key to double-click.</param>
    Task DoubleClickNodeAsync(string sessionId, string treePath, string nodeKey);

    /// <summary>
    /// Executes a complete query against the tree.
    /// </summary>
    /// <param name="sessionId">The session ID.</param>
    /// <param name="query">The query to execute.</param>
    /// <returns>Query result with matches.</returns>
    Task<QueryResult> ExecuteQueryAsync(string sessionId, SapQuery query);
}

