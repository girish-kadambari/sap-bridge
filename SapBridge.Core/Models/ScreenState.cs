namespace SapBridge.Core.Models;

public class ScreenState
{
    public string SessionId { get; set; } = string.Empty;
    public string Transaction { get; set; } = string.Empty;
    public string ScreenNumber { get; set; } = string.Empty;
    public string ProgramName { get; set; } = string.Empty;
    public StatusBarInfo StatusBar { get; set; } = new();
    public string TitleBar { get; set; } = string.Empty;
    public List<ObjectInfo> Objects { get; set; } = new();
    public string? Screenshot { get; set; }
}

public class StatusBarInfo
{
    public string Text { get; set; } = string.Empty;
    public string Type { get; set; } = "info"; // info, success, warning, error
    public List<string> Messages { get; set} = new();
}

public class ObjectInfo
{
    public string Path { get; set; } = string.Empty;
    public string Type { get; set; } = string.Empty;
    public string Name { get; set; } = string.Empty;
    public string Text { get; set; } = string.Empty;
    public string Label { get; set; } = string.Empty;
    public List<string> Methods { get; set; } = new();
    public Dictionary<string, PropertyInfo> Properties { get; set; } = new();
    public List<string> Children { get; set; } = new();
}

public class PropertyInfo
{
    public string Type { get; set; } = string.Empty;
    public object? Value { get; set; }
    public bool Readable { get; set; }
    public bool Writable { get; set; }
}

public class SessionInfo
{
    public string SessionId { get; set; } = string.Empty;
    public string SystemName { get; set; } = string.Empty;
    public string Client { get; set; } = string.Empty;
    public string User { get; set; } = string.Empty;
    public string Language { get; set; } = string.Empty;
    public DateTime ConnectedAt { get; set; }
}

public class ActionRequest
{
    public string SessionId { get; set; } = string.Empty;
    public string ObjectPath { get; set; } = string.Empty;
    public string Method { get; set; } = string.Empty;
    public List<object> Args { get; set; } = new();
}

public class ActionResult
{
    public bool Success { get; set; }
    public object? Result { get; set; }
    public string? Error { get; set; }
    public int ExecutionTimeMs { get; set; }
}

