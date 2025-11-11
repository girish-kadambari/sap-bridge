using System.Diagnostics;
using System.Reflection;
using System.Runtime.InteropServices;
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

    private object? InvokeSetText(dynamic obj, object[] args)
    {
        if (args.Length == 0)
            throw new ArgumentException("SetText requires a text argument");
            
        obj.Text = args[0].ToString();
        return null;
    }

    private object? InvokeSetFocus(dynamic obj)
    {
        obj.SetFocus();
        return null;
    }

    private object? InvokePress(dynamic obj)
    {
        obj.Press();
        return null;
    }

    private object? InvokeSendVKey(dynamic obj, object[] args)
    {
        if (args.Length == 0)
            throw new ArgumentException("SendVKey requires a key code argument");
            
        int keyCode = Convert.ToInt32(args[0]);
        obj.SendVKey(keyCode);
        return null;
    }

    private object? InvokeSelect(dynamic obj)
    {
        obj.Select();
        return null;
    }

    private object? InvokeGetText(dynamic obj)
    {
        return obj.Text?.ToString();
    }

    private object? GetProperty(dynamic obj, object[] args)
    {
        if (args.Length == 0)
            throw new ArgumentException("GetProperty requires a property name");
            
        string propertyName = args[0].ToString() ?? throw new ArgumentException("Property name cannot be null");
        
        return obj.GetType().InvokeMember(
            propertyName,
            System.Reflection.BindingFlags.GetProperty,
            null,
            obj,
            null
        );
    }

    private object? SetProperty(dynamic obj, object[] args)
    {
        if (args.Length < 2)
            throw new ArgumentException("SetProperty requires property name and value");
            
        string propertyName = args[0].ToString() ?? throw new ArgumentException("Property name cannot be null");
        object value = args[1];
        
        obj.GetType().InvokeMember(
            propertyName,
            System.Reflection.BindingFlags.SetProperty,
            null,
            obj,
            new[] { value }
        );
        
        return null;
    }

    private object? InvokeGenericMethod(dynamic obj, string methodName, object[] args)
    {
        return obj.GetType().InvokeMember(
            methodName,
            System.Reflection.BindingFlags.InvokeMethod,
            null,
            obj,
            args
        );
    }
}

