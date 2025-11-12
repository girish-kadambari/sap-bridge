using System.Diagnostics;
using SapBridge.Models;
using SapBridge.Models.Query;
using SapBridge.Repositories;
using SapBridge.Services.Query;
using SapBridge.Utils;
using Serilog;

namespace SapBridge.Services.Tree;

/// <summary>
/// Main service for SAP GUI Tree operations.
/// Integrates tree navigation and query capabilities.
/// </summary>
public class TreeService : ITreeService
{
    private readonly ILogger _logger;
    private readonly ISapGuiRepository _repository;
    private readonly TreeNavigator _navigator;
    private readonly ConditionEvaluator _conditionEvaluator;

    public TreeService(
        ILogger logger,
        ISapGuiRepository repository,
        ConditionEvaluator conditionEvaluator)
    {
        _logger = logger ?? throw new ArgumentNullException(nameof(logger));
        _repository = repository ?? throw new ArgumentNullException(nameof(repository));
        _conditionEvaluator = conditionEvaluator ?? throw new ArgumentNullException(nameof(conditionEvaluator));
        
        _navigator = new TreeNavigator(_logger, _repository);
    }

    /// <inheritdoc/>
    public async Task<TreeData> GetTreeDataAsync(string sessionId, string treePath, bool expandAll = false)
    {
        await Task.CompletedTask;

        try
        {
            _logger.Information("Getting tree data from {TreePath}, expandAll={ExpandAll}", 
                treePath, expandAll);

            var session = _repository.GetSession(sessionId);
            var tree = _repository.FindObjectById(session, treePath);

            if (tree == null)
            {
                throw new InvalidOperationException($"Tree not found at path: {treePath}");
            }

            // Expand all nodes if requested
            if (expandAll)
            {
                _navigator.ExpandAll(tree);
            }

            var treeData = new TreeData
            {
                TreePath = treePath,
                CapturedAt = DateTime.UtcNow,
                IsFullyLoaded = expandAll
            };

            // Get root nodes
            var rootKeys = _navigator.GetRootKeys(tree);
            foreach (var rootKey in rootKeys)
            {
                var node = _navigator.ExtractNode(tree, rootKey, includeChildren: true);
                if (node != null)
                {
                    treeData.RootNodes.Add(node);
                }
            }

            // Count total nodes
            treeData.TotalNodes = CountAllNodes(treeData.RootNodes);

            _logger.Information("Extracted tree with {RootCount} root nodes, {TotalCount} total nodes", 
                treeData.RootNodes.Count, treeData.TotalNodes);

            return treeData;
        }
        catch (Exception ex)
        {
            _logger.Error(ex, "Error getting tree data");
            var mapped = ComExceptionMapper.MapException(ex, $"getting tree data from {treePath}");
            throw new InvalidOperationException(mapped.Message, ex);
        }
    }

    /// <inheritdoc/>
    public async Task<TreeNode> GetNodeAsync(string sessionId, string treePath, string nodeKey)
    {
        await Task.CompletedTask;

        try
        {
            var session = _repository.GetSession(sessionId);
            var tree = _repository.FindObjectById(session, treePath);

            if (tree == null)
            {
                throw new InvalidOperationException($"Tree not found at path: {treePath}");
            }

            var node = _navigator.ExtractNode(tree, nodeKey, includeChildren: true);
            if (node == null)
            {
                throw new InvalidOperationException($"Node not found: {nodeKey}");
            }

            return node;
        }
        catch (Exception ex)
        {
            _logger.Error(ex, "Error getting node {NodeKey}", nodeKey);
            throw;
        }
    }

    /// <inheritdoc/>
    public async Task<List<TreeNode>> GetRootNodesAsync(string sessionId, string treePath)
    {
        await Task.CompletedTask;

        try
        {
            var session = _repository.GetSession(sessionId);
            var tree = _repository.FindObjectById(session, treePath);

            if (tree == null)
            {
                throw new InvalidOperationException($"Tree not found at path: {treePath}");
            }

            var rootNodes = new List<TreeNode>();
            var rootKeys = _navigator.GetRootKeys(tree);

            foreach (var rootKey in rootKeys)
            {
                var node = _navigator.ExtractNode(tree, rootKey);
                if (node != null)
                {
                    rootNodes.Add(node);
                }
            }

            return rootNodes;
        }
        catch (Exception ex)
        {
            _logger.Error(ex, "Error getting root nodes");
            throw;
        }
    }

    /// <inheritdoc/>
    public async Task<List<TreeNode>> GetChildrenAsync(string sessionId, string treePath, string parentKey)
    {
        await Task.CompletedTask;

        try
        {
            var session = _repository.GetSession(sessionId);
            var tree = _repository.FindObjectById(session, treePath);

            if (tree == null)
            {
                throw new InvalidOperationException($"Tree not found at path: {treePath}");
            }

            var children = new List<TreeNode>();
            var childKeys = _navigator.GetChildKeys(tree, parentKey);

            foreach (var childKey in childKeys)
            {
                var node = _navigator.ExtractNode(tree, childKey);
                if (node != null)
                {
                    children.Add(node);
                }
            }

            return children;
        }
        catch (Exception ex)
        {
            _logger.Error(ex, "Error getting children of node {ParentKey}", parentKey);
            throw;
        }
    }

    /// <inheritdoc/>
    public async Task ExpandNodeAsync(string sessionId, string treePath, string nodeKey)
    {
        await Task.CompletedTask;

        try
        {
            var session = _repository.GetSession(sessionId);
            var tree = _repository.FindObjectById(session, treePath);

            if (tree == null)
            {
                throw new InvalidOperationException($"Tree not found at path: {treePath}");
            }

            _navigator.ExpandNode(tree, nodeKey);
            _logger.Information("Expanded node {NodeKey}", nodeKey);
        }
        catch (Exception ex)
        {
            _logger.Error(ex, "Error expanding node {NodeKey}", nodeKey);
            throw;
        }
    }

    /// <inheritdoc/>
    public async Task CollapseNodeAsync(string sessionId, string treePath, string nodeKey)
    {
        await Task.CompletedTask;

        try
        {
            var session = _repository.GetSession(sessionId);
            var tree = _repository.FindObjectById(session, treePath);

            if (tree == null)
            {
                throw new InvalidOperationException($"Tree not found at path: {treePath}");
            }

            _navigator.CollapseNode(tree, nodeKey);
            _logger.Information("Collapsed node {NodeKey}", nodeKey);
        }
        catch (Exception ex)
        {
            _logger.Error(ex, "Error collapsing node {NodeKey}", nodeKey);
            throw;
        }
    }

    /// <inheritdoc/>
    public async Task SelectNodeAsync(string sessionId, string treePath, string nodeKey)
    {
        await Task.CompletedTask;

        try
        {
            var session = _repository.GetSession(sessionId);
            var tree = _repository.FindObjectById(session, treePath);

            if (tree == null)
            {
                throw new InvalidOperationException($"Tree not found at path: {treePath}");
            }

            _navigator.SelectNode(tree, nodeKey);
            _logger.Information("Selected node {NodeKey}", nodeKey);
        }
        catch (Exception ex)
        {
            _logger.Error(ex, "Error selecting node {NodeKey}", nodeKey);
            throw;
        }
    }

    /// <inheritdoc/>
    public async Task<string?> GetSelectedNodeKeyAsync(string sessionId, string treePath)
    {
        await Task.CompletedTask;

        try
        {
            var session = _repository.GetSession(sessionId);
            var tree = _repository.FindObjectById(session, treePath);

            if (tree == null)
            {
                throw new InvalidOperationException($"Tree not found at path: {treePath}");
            }

            return _navigator.GetSelectedNodeKey(tree);
        }
        catch (Exception ex)
        {
            _logger.Error(ex, "Error getting selected node");
            throw;
        }
    }

    /// <inheritdoc/>
    public async Task<List<QueryMatch>> FindNodesAsync(string sessionId, string treePath, List<QueryCondition> conditions, bool searchInText = true)
    {
        await Task.CompletedTask;

        try
        {
            var treeData = await GetTreeDataAsync(sessionId, treePath, expandAll: true);
            var matches = new List<QueryMatch>();

            // Flatten tree to list of all nodes
            var allNodes = FlattenTree(treeData.RootNodes);

            foreach (var node in allNodes)
            {
                // Build searchable data dictionary
                var data = new Dictionary<string, object?>
                {
                    { "Key", node.Key },
                    { "Text", node.Text },
                    { "Level", node.Level },
                    { "HasChildren", node.HasChildren },
                    { "IsExpanded", node.IsExpanded }
                };

                // Add node properties
                foreach (var prop in node.Properties)
                {
                    data[prop.Key] = prop.Value;
                }

                // Evaluate conditions
                if (_conditionEvaluator.EvaluateConditions(conditions, data))
                {
                    matches.Add(new QueryMatch
                    {
                        Index = matches.Count, // Use match index as order
                        Data = data
                    });
                }
            }

            _logger.Information("Found {MatchCount} nodes matching conditions", matches.Count);
            return matches;
        }
        catch (Exception ex)
        {
            _logger.Error(ex, "Error finding nodes");
            throw;
        }
    }

    /// <inheritdoc/>
    public async Task<QueryMatch?> FindFirstNodeAsync(string sessionId, string treePath, List<QueryCondition> conditions, bool searchInText = true)
    {
        var matches = await FindNodesAsync(sessionId, treePath, conditions, searchInText);
        return matches.FirstOrDefault();
    }

    /// <inheritdoc/>
    public async Task<TreeSearchResult> SearchNodeByTextAsync(string sessionId, string treePath, string searchText, bool caseSensitive = false)
    {
        await Task.CompletedTask;

        try
        {
            var treeData = await GetTreeDataAsync(sessionId, treePath, expandAll: true);
            var allNodes = FlattenTree(treeData.RootNodes);

            foreach (var node in allNodes)
            {
                bool matches = caseSensitive 
                    ? node.Text.Contains(searchText)
                    : node.Text.Contains(searchText, StringComparison.OrdinalIgnoreCase);

                if (matches)
                {
                    var path = GetNodePathKeys(node, allNodes);
                    return new TreeSearchResult
                    {
                        Node = node,
                        Path = path
                    };
                }
            }

            return new TreeSearchResult();
        }
        catch (Exception ex)
        {
            _logger.Error(ex, "Error searching for node by text '{SearchText}'", searchText);
            throw;
        }
    }

    /// <inheritdoc/>
    public async Task<List<string>> GetNodePathAsync(string sessionId, string treePath, string nodeKey)
    {
        await Task.CompletedTask;

        try
        {
            var session = _repository.GetSession(sessionId);
            var tree = _repository.FindObjectById(session, treePath);

            if (tree == null)
            {
                throw new InvalidOperationException($"Tree not found at path: {treePath}");
            }

            return _navigator.GetNodePath(tree, nodeKey);
        }
        catch (Exception ex)
        {
            _logger.Error(ex, "Error getting path to node {NodeKey}", nodeKey);
            throw;
        }
    }

    /// <inheritdoc/>
    public async Task DoubleClickNodeAsync(string sessionId, string treePath, string nodeKey)
    {
        await Task.CompletedTask;

        try
        {
            var session = _repository.GetSession(sessionId);
            var tree = _repository.FindObjectById(session, treePath);

            if (tree == null)
            {
                throw new InvalidOperationException($"Tree not found at path: {treePath}");
            }

            _navigator.DoubleClickNode(tree, nodeKey);
            _logger.Information("Double-clicked node {NodeKey}", nodeKey);
        }
        catch (Exception ex)
        {
            _logger.Error(ex, "Error double-clicking node {NodeKey}", nodeKey);
            throw;
        }
    }

    /// <inheritdoc/>
    public async Task<QueryResult> ExecuteQueryAsync(string sessionId, SapQuery query)
    {
        var stopwatch = Stopwatch.StartNew();

        try
        {
            _logger.Information("Executing tree query: Action={Action}, Conditions={ConditionCount}", 
                query.Action, query.Conditions.Count);

            List<QueryMatch> matches = query.Action switch
            {
                QueryAction.GetFirst => new List<QueryMatch>
                {
                    (await FindFirstNodeAsync(sessionId, query.ObjectPath, query.Conditions))!
                }.Where(m => m != null).ToList(),
                
                QueryAction.GetAll => await FindNodesAsync(sessionId, query.ObjectPath, query.Conditions),
                
                QueryAction.Count => await FindNodesAsync(sessionId, query.ObjectPath, query.Conditions),
                
                _ => throw new NotSupportedException($"Query action '{query.Action}' is not supported for trees.")
            };

            stopwatch.Stop();
            return QueryResult.SuccessResult(matches, (int)stopwatch.ElapsedMilliseconds);
        }
        catch (Exception ex)
        {
            stopwatch.Stop();
            _logger.Error(ex, "Error executing tree query");
            return QueryResult.FailureResult(ex.Message, (int)stopwatch.ElapsedMilliseconds);
        }
    }

    /// <summary>
    /// Counts all nodes in a tree structure.
    /// </summary>
    private int CountAllNodes(List<TreeNode> nodes)
    {
        int count = nodes.Count;
        foreach (var node in nodes)
        {
            count += CountAllNodes(node.Children);
        }
        return count;
    }

    /// <summary>
    /// Flattens a tree structure into a list of all nodes.
    /// </summary>
    private List<TreeNode> FlattenTree(List<TreeNode> nodes)
    {
        var flatList = new List<TreeNode>();
        foreach (var node in nodes)
        {
            flatList.Add(node);
            flatList.AddRange(FlattenTree(node.Children));
        }
        return flatList;
    }

    /// <summary>
    /// Gets the path of keys from root to a node.
    /// </summary>
    private List<string> GetNodePathKeys(TreeNode targetNode, List<TreeNode> allNodes)
    {
        var path = new List<string> { targetNode.Key };
        string? currentParent = targetNode.ParentKey;

        while (currentParent != null)
        {
            path.Insert(0, currentParent);
            var parentNode = allNodes.FirstOrDefault(n => n.Key == currentParent);
            currentParent = parentNode?.ParentKey;
        }

        return path;
    }
}

