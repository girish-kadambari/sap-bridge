using System.Diagnostics;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text.Json;
using Serilog;
using SapBridge.Core.Models;

namespace SapBridge.Core;

public class ActionExecutor
{
    private readonly ILogger _logger;

    public ActionExecutor(ILogger logger)
    {
        _logger = logger;
    }

    public ActionResult ExecuteAction(object session, ActionRequest request)
    {
        var stopwatch = Stopwatch.StartNew();
        
        try
        {
            _logger.Information($"Executing {request.Method} on {request.ObjectPath}");

            // Find the object
            object? obj = FindById(session, request.ObjectPath);
            
            if (obj == null)
            {
                return new ActionResult
                {
                    Success = false,
                    Error = $"Object not found: {request.ObjectPath}",
                    ExecutionTimeMs = (int)stopwatch.ElapsedMilliseconds
                };
            }

            // Execute the method
            object? result = InvokeMethod(obj, request.Method, request.Args.ToArray());

            stopwatch.Stop();
            
            return new ActionResult
            {
                Success = true,
                Result = result,
                ExecutionTimeMs = (int)stopwatch.ElapsedMilliseconds
            };
        }
        catch (COMException ex)
        {
            stopwatch.Stop();
            _logger.Error(ex, $"COM Exception executing {request.Method}");
            
            return new ActionResult
            {
                Success = false,
                Error = $"COM Error: {ex.Message}",
                ExecutionTimeMs = (int)stopwatch.ElapsedMilliseconds
            };
        }
        catch (Exception ex)
        {
            stopwatch.Stop();
            _logger.Error(ex, $"Error executing {request.Method}");
            
            return new ActionResult
            {
                Success = false,
                Error = ex.Message,
                ExecutionTimeMs = (int)stopwatch.ElapsedMilliseconds
            };
        }
    }

    private static object? FindById(object session, string path)
    {
        return session.GetType().InvokeMember(
            "FindById",
            BindingFlags.InvokeMethod,
            null,
            session,
            new object[] { path }
        );
    }

    private object? InvokeMethod(object obj, string methodName, object[] args)
    {
        try
        {
            // Handle common SAP GUI methods
            return methodName.ToLower() switch
            {
                "settext" => InvokeSetText(obj, args),
                "setfocus" => InvokeSetFocus(obj),
                "press" => InvokePress(obj),
                "sendvkey" => InvokeSendVKey(obj, args),
                "select" => InvokeSelect(obj),
                "gettext" => InvokeGetText(obj),
                "getproperty" => GetProperty(obj, args),
                "setproperty" => SetProperty(obj, args),
                _ => InvokeGenericMethod(obj, methodName, args)
            };
        }
        catch (Exception ex)
        {
            _logger.Error(ex, $"Error invoking method: {methodName}");
            throw;
        }
    }

    private object? InvokeSetText(object obj, object[] args)
    {
        if (args.Length == 0)
            throw new ArgumentException("SetText requires a text argument");
        
        // Convert JsonElement to string if needed
        string text = ConvertToNativeType<string>(args[0]);
        
        // Use reflection to set Text property
        obj.GetType().InvokeMember(
            "Text",
            BindingFlags.SetProperty,
            null,
            obj,
            new[] { text }
        );
        return null;
    }

    private object? InvokeSetFocus(object obj)
    {
        // Use reflection to call SetFocus method
        return obj.GetType().InvokeMember(
            "SetFocus",
            BindingFlags.InvokeMethod,
            null,
            obj,
            null
        );
    }

    private object? InvokePress(object obj)
    {
        // Use reflection to call Press method
        return obj.GetType().InvokeMember(
            "Press",
            BindingFlags.InvokeMethod,
            null,
            obj,
            null
        );
    }

    private object? InvokeSendVKey(object obj, object[] args)
    {
        if (args.Length == 0)
            throw new ArgumentException("SendVKey requires a key code argument");
        
        // Convert JsonElement to int if needed
        int keyCode = ConvertToNativeType<int>(args[0]);
        
        // Use reflection to call SendVKey method
        return obj.GetType().InvokeMember(
            "SendVKey",
            BindingFlags.InvokeMethod,
            null,
            obj,
            new object[] { keyCode }
        );
    }

    private object? InvokeSelect(object obj)
    {
        // Use reflection to call Select method
        return obj.GetType().InvokeMember(
            "Select",
            BindingFlags.InvokeMethod,
            null,
            obj,
            null
        );
    }

    private object? InvokeGetText(object obj)
    {
        // Use reflection to get Text property
        return obj.GetType().InvokeMember(
            "Text",
            BindingFlags.GetProperty,
            null,
            obj,
            null
        )?.ToString();
    }

    private object? GetProperty(object obj, object[] args)
    {
        if (args.Length == 0)
            throw new ArgumentException("GetProperty requires a property name");
            
        string propertyName = ConvertToNativeType<string>(args[0]);
        
        return obj.GetType().InvokeMember(
            propertyName,
            BindingFlags.GetProperty,
            null,
            obj,
            null
        );
    }

    private object? SetProperty(object obj, object[] args)
    {
        if (args.Length < 2)
            throw new ArgumentException("SetProperty requires property name and value");
            
        string propertyName = ConvertToNativeType<string>(args[0]);
        object value = ConvertJsonElement(args[1]);
        
        obj.GetType().InvokeMember(
            propertyName,
            BindingFlags.SetProperty,
            null,
            obj,
            new[] { value }
        );
        
        return null;
    }

    private object? InvokeGenericMethod(object obj, string methodName, object[] args)
    {
        // Convert any JsonElement args to native types
        object[] convertedArgs = args.Select(ConvertJsonElement).ToArray();
        
        return obj.GetType().InvokeMember(
            methodName,
            BindingFlags.InvokeMethod,
            null,
            obj,
            convertedArgs
        );
    }

    /// <summary>
    /// Convert JsonElement or object to native type T
    /// Handles the case where args come from JSON deserialization
    /// </summary>
    private static T ConvertToNativeType<T>(object value)
    {
        if (value is JsonElement jsonElement)
        {
            return jsonElement.Deserialize<T>() ?? throw new InvalidCastException($"Cannot convert JsonElement to {typeof(T).Name}");
        }
        
        if (value is T typedValue)
        {
            return typedValue;
        }
        
        // Fallback to Convert
        return (T)Convert.ChangeType(value, typeof(T));
    }

    /// <summary>
    /// Convert JsonElement to appropriate .NET type
    /// </summary>
    private static object ConvertJsonElement(object value)
    {
        if (value is JsonElement jsonElement)
        {
            return jsonElement.ValueKind switch
            {
                JsonValueKind.String => jsonElement.GetString() ?? string.Empty,
                JsonValueKind.Number => jsonElement.TryGetInt32(out int intVal) ? intVal : jsonElement.GetDouble(),
                JsonValueKind.True => true,
                JsonValueKind.False => false,
                JsonValueKind.Null => null!,
                _ => jsonElement.ToString()
            };
        }
        
        return value;
    }
}

