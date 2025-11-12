using SapBridge.Models;
using SapBridge.Repositories;
using SapBridge.Utils;
using Serilog;
using ILogger = Serilog.ILogger;

namespace SapBridge.Services.Tree;

/// <summary>
/// Handles navigation and operations on SAP GUI Tree nodes.
/// Manages expansion, selection, and node interactions.
/// </summary>
public class TreeNavigator
{
    private readonly ILogger _logger;
    private readonly ISapGuiRepository _repository;

    public TreeNavigator(ILogger logger, ISapGuiRepository repository)
    {
        _logger = logger ?? throw new ArgumentNullException(nameof(logger));
        _repository = repository ?? throw new ArgumentNullException(nameof(repository));
    }

    /// <summary>
    /// Extracts a single tree node by key.
    /// </summary>
    /// <param name="tree">The tree COM object.</param>
    /// <param name="nodeKey">Node key to extract.</param>
    /// <param name="includeChildren">Whether to include child nodes.</param>
    /// <returns>The tree node.</returns>
    public TreeNode? ExtractNode(object tree, string nodeKey, bool includeChildren = false)
    {
        try
        {
            var node = new TreeNode
            {
                Key = nodeKey
            };

            // Get node text
            node.Text = GetNodeText(tree, nodeKey) ?? nodeKey;

            // Check if expanded
            node.IsExpanded = IsNodeExpanded(tree, nodeKey);

            // Check if has children
            node.HasChildren = HasChildren(tree, nodeKey);

            // Get parent key
            node.ParentKey = GetParentKey(tree, nodeKey);

            // Get node level
            node.Level = GetNodeLevel(tree, nodeKey);

            // Get node icon
            node.Icon = GetNodeIcon(tree, nodeKey);

            // Get child nodes if requested
            if (includeChildren && node.HasChildren)
            {
                var childKeys = GetChildKeys(tree, nodeKey);
                foreach (var childKey in childKeys)
                {
                    var childNode = ExtractNode(tree, childKey, includeChildren);
                    if (childNode != null)
                    {
                        node.Children.Add(childNode);
                    }
                }
            }

            return node;
        }
        catch (Exception ex)
        {
            _logger.Warning(ex, "Could not extract node {NodeKey}", nodeKey);
            return null;
        }
    }

    /// <summary>
    /// Gets all root node keys from the tree.
    /// </summary>
    /// <param name="tree">The tree COM object.</param>
    /// <returns>List of root node keys.</returns>
    public List<string> GetRootKeys(object tree)
    {
        var rootKeys = new List<string>();

        try
        {
            // Try to get root node keys collection
            var roots = _repository.InvokeObjectMethod(tree, "GetRootKeys");
            if (roots != null)
            {
                rootKeys = ConvertToStringList(roots);
            }

            if (rootKeys.Count == 0)
            {
                // Try alternative method - get all nodes at level 0
                var allNodes = GetAllNodeKeys(tree);
                foreach (var nodeKey in allNodes)
                {
                    if (GetNodeLevel(tree, nodeKey) == 0)
                    {
                        rootKeys.Add(nodeKey);
                    }
                }
            }

            _logger.Debug("Found {Count} root nodes", rootKeys.Count);
        }
        catch (Exception ex)
        {
            _logger.Warning(ex, "Could not get root keys");
        }

        return rootKeys;
    }

    /// <summary>
    /// Gets all child keys of a node.
    /// </summary>
    /// <param name="tree">The tree COM object.</param>
    /// <param name="parentKey">Parent node key.</param>
    /// <returns>List of child node keys.</returns>
    public List<string> GetChildKeys(object tree, string parentKey)
    {
        var childKeys = new List<string>();

        try
        {
            var children = _repository.InvokeObjectMethod(tree, "GetChildrenKeys", parentKey);
            if (children != null)
            {
                childKeys = ConvertToStringList(children);
            }

            _logger.Debug("Node {ParentKey} has {Count} children", parentKey, childKeys.Count);
        }
        catch (Exception ex)
        {
            _logger.Debug(ex, "Could not get children for node {ParentKey}", parentKey);
        }

        return childKeys;
    }

    /// <summary>
    /// Gets all node keys in the tree.
    /// </summary>
    /// <param name="tree">The tree COM object.</param>
    /// <returns>List of all node keys.</returns>
    public List<string> GetAllNodeKeys(object tree)
    {
        var nodeKeys = new List<string>();

        try
        {
            var allNodes = _repository.InvokeObjectMethod(tree, "GetAllNodeKeys");
            if (allNodes != null)
            {
                nodeKeys = ConvertToStringList(allNodes);
            }

            _logger.Debug("Tree has {Count} total nodes", nodeKeys.Count);
        }
        catch (Exception ex)
        {
            _logger.Warning(ex, "Could not get all node keys");
        }

        return nodeKeys;
    }

    /// <summary>
    /// Gets the text/label of a node.
    /// </summary>
    /// <param name="tree">The tree COM object.</param>
    /// <param name="nodeKey">Node key.</param>
    /// <returns>Node text.</returns>
    public string? GetNodeText(object tree, string nodeKey)
    {
        try
        {
            var text = _repository.InvokeObjectMethod(tree, "GetNodeText", nodeKey);
            return text?.ToString();
        }
        catch (Exception ex)
        {
            _logger.Debug(ex, "Could not get text for node {NodeKey}", nodeKey);
            return null;
        }
    }

    /// <summary>
    /// Checks if a node is expanded.
    /// </summary>
    /// <param name="tree">The tree COM object.</param>
    /// <param name="nodeKey">Node key.</param>
    /// <returns>True if expanded.</returns>
    public bool IsNodeExpanded(object tree, string nodeKey)
    {
        try
        {
            var expanded = _repository.InvokeObjectMethod(tree, "IsNodeExpanded", nodeKey);
            return expanded != null && Convert.ToBoolean(expanded);
        }
        catch (Exception ex)
        {
            _logger.Debug(ex, "Could not check if node {NodeKey} is expanded", nodeKey);
            return false;
        }
    }

    /// <summary>
    /// Checks if a node has children.
    /// </summary>
    /// <param name="tree">The tree COM object.</param>
    /// <param name="nodeKey">Node key.</param>
    /// <returns>True if has children.</returns>
    public bool HasChildren(object tree, string nodeKey)
    {
        try
        {
            var hasChildren = _repository.InvokeObjectMethod(tree, "HasChildren", nodeKey);
            if (hasChildren != null)
            {
                return Convert.ToBoolean(hasChildren);
            }

            // Alternative: check child count
            var childKeys = GetChildKeys(tree, nodeKey);
            return childKeys.Count > 0;
        }
        catch (Exception ex)
        {
            _logger.Debug(ex, "Could not check if node {NodeKey} has children", nodeKey);
            return false;
        }
    }

    /// <summary>
    /// Gets the parent key of a node.
    /// </summary>
    /// <param name="tree">The tree COM object.</param>
    /// <param name="nodeKey">Node key.</param>
    /// <returns>Parent node key or null if root.</returns>
    public string? GetParentKey(object tree, string nodeKey)
    {
        try
        {
            var parent = _repository.InvokeObjectMethod(tree, "GetParentKey", nodeKey);
            return parent?.ToString();
        }
        catch (Exception ex)
        {
            _logger.Debug(ex, "Could not get parent for node {NodeKey}", nodeKey);
            return null;
        }
    }

    /// <summary>
    /// Gets the level/depth of a node (0 for root).
    /// </summary>
    /// <param name="tree">The tree COM object.</param>
    /// <param name="nodeKey">Node key.</param>
    /// <returns>Node level.</returns>
    public int GetNodeLevel(object tree, string nodeKey)
    {
        try
        {
            var level = _repository.InvokeObjectMethod(tree, "GetNodeLevel", nodeKey);
            if (level != null)
            {
                return Convert.ToInt32(level);
            }

            // Calculate level by counting parents
            int calculatedLevel = 0;
            string? currentKey = nodeKey;
            while (currentKey != null)
            {
                currentKey = GetParentKey(tree, currentKey);
                if (currentKey != null)
                {
                    calculatedLevel++;
                }
            }

            return calculatedLevel;
        }
        catch (Exception ex)
        {
            _logger.Debug(ex, "Could not get level for node {NodeKey}", nodeKey);
            return 0;
        }
    }

    /// <summary>
    /// Gets the icon name for a node.
    /// </summary>
    /// <param name="tree">The tree COM object.</param>
    /// <param name="nodeKey">Node key.</param>
    /// <returns>Icon name or null.</returns>
    public string? GetNodeIcon(object tree, string nodeKey)
    {
        try
        {
            var icon = _repository.InvokeObjectMethod(tree, "GetNodeIcon", nodeKey);
            return icon?.ToString();
        }
        catch (Exception ex)
        {
            _logger.Debug(ex, "Could not get icon for node {NodeKey}", nodeKey);
            return null;
        }
    }

    /// <summary>
    /// Expands a node to show its children.
    /// </summary>
    /// <param name="tree">The tree COM object.</param>
    /// <param name="nodeKey">Node key to expand.</param>
    public void ExpandNode(object tree, string nodeKey)
    {
        try
        {
            _repository.InvokeObjectMethod(tree, "ExpandNode", nodeKey);
            _logger.Debug("Expanded node {NodeKey}", nodeKey);
        }
        catch (Exception ex)
        {
            _logger.Error(ex, "Error expanding node {NodeKey}", nodeKey);
            throw;
        }
    }

    /// <summary>
    /// Collapses a node to hide its children.
    /// </summary>
    /// <param name="tree">The tree COM object.</param>
    /// <param name="nodeKey">Node key to collapse.</param>
    public void CollapseNode(object tree, string nodeKey)
    {
        try
        {
            _repository.InvokeObjectMethod(tree, "CollapseNode", nodeKey);
            _logger.Debug("Collapsed node {NodeKey}", nodeKey);
        }
        catch (Exception ex)
        {
            _logger.Error(ex, "Error collapsing node {NodeKey}", nodeKey);
            throw;
        }
    }

    /// <summary>
    /// Expands all nodes in the tree.
    /// </summary>
    /// <param name="tree">The tree COM object.</param>
    public void ExpandAll(object tree)
    {
        try
        {
            _repository.InvokeObjectMethod(tree, "ExpandAll");
            _logger.Debug("Expanded all nodes");
        }
        catch (Exception ex)
        {
            _logger.Warning(ex, "Could not expand all nodes, trying manual expansion");
            
            // Manual expansion fallback
            var allKeys = GetAllNodeKeys(tree);
            foreach (var key in allKeys)
            {
                try
                {
                    if (HasChildren(tree, key))
                    {
                        ExpandNode(tree, key);
                    }
                }
                catch
                {
                    // Continue with other nodes
                }
            }
        }
    }

    /// <summary>
    /// Selects a node in the tree.
    /// </summary>
    /// <param name="tree">The tree COM object.</param>
    /// <param name="nodeKey">Node key to select.</param>
    public void SelectNode(object tree, string nodeKey)
    {
        try
        {
            _repository.InvokeObjectMethod(tree, "SelectNode", nodeKey);
            _logger.Debug("Selected node {NodeKey}", nodeKey);
        }
        catch (Exception ex)
        {
            _logger.Error(ex, "Error selecting node {NodeKey}", nodeKey);
            throw;
        }
    }

    /// <summary>
    /// Gets the currently selected node key.
    /// </summary>
    /// <param name="tree">The tree COM object.</param>
    /// <returns>Selected node key or null.</returns>
    public string? GetSelectedNodeKey(object tree)
    {
        try
        {
            var selected = _repository.GetObjectProperty(tree, "SelectedNode");
            return selected?.ToString();
        }
        catch (Exception ex)
        {
            _logger.Debug(ex, "Could not get selected node");
            return null;
        }
    }

    /// <summary>
    /// Double-clicks a node (typically to trigger an action).
    /// </summary>
    /// <param name="tree">The tree COM object.</param>
    /// <param name="nodeKey">Node key to double-click.</param>
    public void DoubleClickNode(object tree, string nodeKey)
    {
        try
        {
            _repository.InvokeObjectMethod(tree, "DoubleClickNode", nodeKey);
            _logger.Debug("Double-clicked node {NodeKey}", nodeKey);
        }
        catch (Exception ex)
        {
            _logger.Error(ex, "Error double-clicking node {NodeKey}", nodeKey);
            throw;
        }
    }

    /// <summary>
    /// Gets the path from root to a specific node.
    /// </summary>
    /// <param name="tree">The tree COM object.</param>
    /// <param name="nodeKey">Target node key.</param>
    /// <returns>List of node keys from root to target.</returns>
    public List<string> GetNodePath(object tree, string nodeKey)
    {
        var path = new List<string>();

        try
        {
            string? currentKey = nodeKey;
            while (currentKey != null)
            {
                path.Insert(0, currentKey); // Add to front
                currentKey = GetParentKey(tree, currentKey);
            }

            _logger.Debug("Path to node {NodeKey} has {Count} nodes", nodeKey, path.Count);
        }
        catch (Exception ex)
        {
            _logger.Warning(ex, "Could not get path to node {NodeKey}", nodeKey);
        }

        return path;
    }

    /// <summary>
    /// Converts a COM collection to a list of strings.
    /// </summary>
    private List<string> ConvertToStringList(object collection)
    {
        var list = new List<string>();

        try
        {
            var count = ComReflectionHelper.GetCollectionCount(collection);
            for (int i = 0; i < count; i++)
            {
                var item = ComReflectionHelper.GetCollectionItem(collection, i);
                if (item != null)
                {
                    list.Add(item.ToString() ?? string.Empty);
                }
            }
        }
        catch (Exception ex)
        {
            _logger.Debug(ex, "Could not convert collection to string list");
        }

        return list;
    }
}

