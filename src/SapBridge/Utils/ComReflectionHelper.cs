using System.Reflection;
using System.Runtime.InteropServices;

namespace SapBridge.Utils;

/// <summary>
/// Utility class for COM object interaction using reflection.
/// Avoids type library dependencies and provides type-safe wrappers.
/// </summary>
public static class ComReflectionHelper
{
    /// <summary>
    /// Gets a property value from a COM object using reflection.
    /// </summary>
    /// <param name="comObject">The COM object.</param>
    /// <param name="propertyName">The property name to retrieve.</param>
    /// <returns>The property value or null if not found.</returns>
    /// <exception cref="ArgumentNullException">Thrown when comObject is null.</exception>
    /// <exception cref="COMException">Thrown when COM operation fails.</exception>
    public static object? GetProperty(object comObject, string propertyName)
    {
        if (comObject == null)
            throw new ArgumentNullException(nameof(comObject));
        
        if (string.IsNullOrWhiteSpace(propertyName))
            throw new ArgumentException("Property name cannot be null or empty.", nameof(propertyName));

        try
        {
            return comObject.GetType().InvokeMember(
                propertyName,
                BindingFlags.GetProperty,
                null,
                comObject,
                null
            );
        }
        catch (TargetInvocationException ex) when (ex.InnerException is COMException)
        {
            throw new COMException($"Failed to get property '{propertyName}'", ex.InnerException);
        }
    }

    /// <summary>
    /// Gets a property value with type safety.
    /// </summary>
    /// <typeparam name="T">The expected return type.</typeparam>
    /// <param name="comObject">The COM object.</param>
    /// <param name="propertyName">The property name to retrieve.</param>
    /// <returns>The typed property value or default(T) if not found or conversion fails.</returns>
    public static T? GetPropertySafe<T>(object comObject, string propertyName)
    {
        try
        {
            var value = GetProperty(comObject, propertyName);
            if (value == null)
                return default;

            if (value is T typedValue)
                return typedValue;

            // Attempt conversion
            return (T)Convert.ChangeType(value, typeof(T));
        }
        catch
        {
            return default;
        }
    }

    /// <summary>
    /// Gets a property value as a string, handling null and empty values.
    /// </summary>
    /// <param name="comObject">The COM object.</param>
    /// <param name="propertyName">The property name to retrieve.</param>
    /// <param name="defaultValue">Default value if property is null or empty.</param>
    /// <returns>The property value as string.</returns>
    public static string GetPropertyAsString(object comObject, string propertyName, string defaultValue = "")
    {
        try
        {
            var value = GetProperty(comObject, propertyName);
            return value?.ToString() ?? defaultValue;
        }
        catch
        {
            return defaultValue;
        }
    }

    /// <summary>
    /// Sets a property value on a COM object using reflection.
    /// </summary>
    /// <param name="comObject">The COM object.</param>
    /// <param name="propertyName">The property name to set.</param>
    /// <param name="value">The value to set.</param>
    /// <exception cref="ArgumentNullException">Thrown when comObject is null.</exception>
    /// <exception cref="COMException">Thrown when COM operation fails.</exception>
    public static void SetProperty(object comObject, string propertyName, object value)
    {
        if (comObject == null)
            throw new ArgumentNullException(nameof(comObject));
        
        if (string.IsNullOrWhiteSpace(propertyName))
            throw new ArgumentException("Property name cannot be null or empty.", nameof(propertyName));

        try
        {
            comObject.GetType().InvokeMember(
                propertyName,
                BindingFlags.SetProperty,
                null,
                comObject,
                new[] { value }
            );
        }
        catch (TargetInvocationException ex) when (ex.InnerException is COMException)
        {
            throw new COMException($"Failed to set property '{propertyName}'", ex.InnerException);
        }
    }

    /// <summary>
    /// Invokes a method on a COM object using reflection.
    /// </summary>
    /// <param name="comObject">The COM object.</param>
    /// <param name="methodName">The method name to invoke.</param>
    /// <param name="args">The method arguments.</param>
    /// <returns>The method return value or null.</returns>
    /// <exception cref="ArgumentNullException">Thrown when comObject is null.</exception>
    /// <exception cref="COMException">Thrown when COM operation fails.</exception>
    public static object? InvokeMethod(object comObject, string methodName, params object[] args)
    {
        if (comObject == null)
            throw new ArgumentNullException(nameof(comObject));
        
        if (string.IsNullOrWhiteSpace(methodName))
            throw new ArgumentException("Method name cannot be null or empty.", nameof(methodName));

        try
        {
            return comObject.GetType().InvokeMember(
                methodName,
                BindingFlags.InvokeMethod,
                null,
                comObject,
                args
            );
        }
        catch (TargetInvocationException ex) when (ex.InnerException is COMException)
        {
            throw new COMException($"Failed to invoke method '{methodName}'", ex.InnerException);
        }
    }

    /// <summary>
    /// Attempts to invoke a method on a COM object, returning success status.
    /// </summary>
    /// <param name="comObject">The COM object.</param>
    /// <param name="methodName">The method name to invoke.</param>
    /// <param name="args">The method arguments.</param>
    /// <param name="result">The method return value if successful.</param>
    /// <returns>True if the method was invoked successfully, false otherwise.</returns>
    public static bool TryInvokeMethod(object comObject, string methodName, object[] args, out object? result)
    {
        result = null;
        
        if (comObject == null || string.IsNullOrWhiteSpace(methodName))
            return false;

        try
        {
            result = InvokeMethod(comObject, methodName, args);
            return true;
        }
        catch
        {
            return false;
        }
    }

    /// <summary>
    /// Checks if a COM object has a specific property.
    /// </summary>
    /// <param name="comObject">The COM object.</param>
    /// <param name="propertyName">The property name to check.</param>
    /// <returns>True if the property exists, false otherwise.</returns>
    public static bool HasProperty(object comObject, string propertyName)
    {
        if (comObject == null || string.IsNullOrWhiteSpace(propertyName))
            return false;

        try
        {
            GetProperty(comObject, propertyName);
            return true;
        }
        catch
        {
            return false;
        }
    }

    /// <summary>
    /// Checks if a COM object has a specific method.
    /// </summary>
    /// <param name="comObject">The COM object.</param>
    /// <param name="methodName">The method name to check.</param>
    /// <returns>True if the method exists, false otherwise.</returns>
    public static bool HasMethod(object comObject, string methodName)
    {
        if (comObject == null || string.IsNullOrWhiteSpace(methodName))
            return false;

        try
        {
            var type = comObject.GetType();
            var methods = type.GetMethods(BindingFlags.Public | BindingFlags.Instance);
            return methods.Any(m => m.Name.Equals(methodName, StringComparison.OrdinalIgnoreCase));
        }
        catch
        {
            return false;
        }
    }

    /// <summary>
    /// Safely releases a COM object.
    /// </summary>
    /// <param name="comObject">The COM object to release.</param>
    public static void ReleaseComObject(ref object? comObject)
    {
        if (comObject != null)
        {
            try
            {
                Marshal.ReleaseComObject(comObject);
            }
            catch
            {
                // Ignore errors during release
            }
            finally
            {
                comObject = null;
            }
        }
    }

    /// <summary>
    /// Gets the count of items in a COM collection.
    /// </summary>
    /// <param name="collection">The COM collection object.</param>
    /// <returns>The count of items, or 0 if the collection is null or has no Count property.</returns>
    public static int GetCollectionCount(object? collection)
    {
        if (collection == null)
            return 0;

        try
        {
            var count = GetProperty(collection, "Count");
            return count != null ? Convert.ToInt32(count) : 0;
        }
        catch
        {
            return 0;
        }
    }

    /// <summary>
    /// Gets an item from a COM collection by index.
    /// </summary>
    /// <param name="collection">The COM collection object.</param>
    /// <param name="index">The zero-based index.</param>
    /// <returns>The item at the specified index, or null if not found.</returns>
    public static object? GetCollectionItem(object collection, int index)
    {
        if (collection == null)
            return null;

        try
        {
            return InvokeMethod(collection, "Item", index);
        }
        catch
        {
            // Try ElementAt for some collection types
            try
            {
                return InvokeMethod(collection, "ElementAt", index);
            }
            catch
            {
                return null;
            }
        }
    }
}

