using System.Runtime.InteropServices;

namespace SapBridge.Utils;

/// <summary>
/// Maps COM exceptions to meaningful error messages for AI agents and users.
/// Provides categorization and context-aware error descriptions.
/// </summary>
public static class ComExceptionMapper
{
    /// <summary>
    /// Error categories for COM exceptions.
    /// </summary>
    public enum ErrorCategory
    {
        Connection,
        ObjectNotFound,
        Permission,
        InvalidOperation,
        SessionState,
        Timeout,
        Unknown
    }

    /// <summary>
    /// Mapped error information with category and user-friendly message.
    /// </summary>
    public class MappedError
    {
        public ErrorCategory Category { get; set; }
        public string Message { get; set; } = string.Empty;
        public string TechnicalDetails { get; set; } = string.Empty;
        public string Suggestion { get; set; } = string.Empty;
        public int ErrorCode { get; set; }
    }

    private static readonly Dictionary<int, (ErrorCategory Category, string Message, string Suggestion)> ErrorMappings = new()
    {
        // Connection Errors
        { unchecked((int)0x800401E4), (ErrorCategory.Connection, "SAP GUI not found. The SAP Logon application is not running.", "Please start SAP Logon and ensure it is running.") },
        { unchecked((int)0x80029C4A), (ErrorCategory.Connection, "SAP GUI scripting is disabled or type library not found.", "Enable SAP GUI scripting in Options → Accessibility & Scripting → Scripting.") },
        { unchecked((int)0x80010108), (ErrorCategory.Connection, "The SAP GUI connection has been disconnected.", "Reconnect to SAP GUI and try again.") },
        
        // Object Not Found Errors
        { unchecked((int)0x8002000E), (ErrorCategory.ObjectNotFound, "The specified SAP GUI object was not found.", "Verify the object path is correct and the object exists on the current screen.") },
        { unchecked((int)0x80020009), (ErrorCategory.ObjectNotFound, "The object does not have the requested property or method.", "Check the object type and verify it supports the requested operation.") },
        
        // Permission Errors
        { unchecked((int)0x80070005), (ErrorCategory.Permission, "Access denied. Insufficient permissions to perform this operation.", "Check user permissions and ensure SAP GUI scripting is allowed for your account.") },
        { unchecked((int)0x800706BA), (ErrorCategory.Permission, "The RPC server is unavailable.", "Check Windows permissions and firewall settings.") },
        
        // Invalid Operation Errors
        { unchecked((int)0x80020006), (ErrorCategory.InvalidOperation, "The object property or method cannot be used in the current state.", "Ensure the screen is in the correct state before performing this operation.") },
        { unchecked((int)0x8002000A), (ErrorCategory.InvalidOperation, "Invalid argument or parameter type.", "Check the data types and values of the parameters being passed.") },
        { unchecked((int)0x80020005), (ErrorCategory.InvalidOperation, "Type mismatch. The value type does not match the expected type.", "Verify parameter types match what the SAP GUI object expects.") },
        
        // Session State Errors
        { unchecked((int)0x800706BE), (ErrorCategory.SessionState, "The SAP session is no longer available or has timed out.", "Reconnect to SAP and start a new session.") },
        { unchecked((int)0x80010001), (ErrorCategory.SessionState, "SAP GUI has been closed or the session is invalid.", "Restart SAP GUI and reconnect.") },
        
        // Timeout Errors
        { unchecked((int)0x8001011F), (ErrorCategory.Timeout, "The operation timed out waiting for SAP GUI to respond.", "The SAP system may be slow. Try increasing the timeout or check SAP server performance.") },
    };

    /// <summary>
    /// Maps a COM exception to a user-friendly error with context and suggestions.
    /// </summary>
    /// <param name="exception">The COM exception to map.</param>
    /// <param name="context">Optional context about what operation was being performed.</param>
    /// <returns>Mapped error information with category and suggestions.</returns>
    public static MappedError MapException(COMException exception, string? context = null)
    {
        if (exception == null)
            throw new ArgumentNullException(nameof(exception));

        var errorCode = exception.ErrorCode;
        
        if (ErrorMappings.TryGetValue(errorCode, out var mapping))
        {
            return new MappedError
            {
                Category = mapping.Category,
                Message = BuildContextualMessage(mapping.Message, context),
                TechnicalDetails = $"COM Error 0x{errorCode:X8}: {exception.Message}",
                Suggestion = mapping.Suggestion,
                ErrorCode = errorCode
            };
        }

        // Unknown error
        return new MappedError
        {
            Category = ErrorCategory.Unknown,
            Message = BuildContextualMessage("An unexpected COM error occurred.", context),
            TechnicalDetails = $"COM Error 0x{errorCode:X8}: {exception.Message}",
            Suggestion = "Please check the SAP GUI connection and try again. If the problem persists, review the technical details.",
            ErrorCode = errorCode
        };
    }

    /// <summary>
    /// Maps a general exception to an error message.
    /// </summary>
    /// <param name="exception">The exception to map.</param>
    /// <param name="context">Optional context about what operation was being performed.</param>
    /// <returns>Mapped error information.</returns>
    public static MappedError MapException(Exception exception, string? context = null)
    {
        if (exception == null)
            throw new ArgumentNullException(nameof(exception));

        if (exception is COMException comException)
            return MapException(comException, context);

        return new MappedError
        {
            Category = ErrorCategory.Unknown,
            Message = BuildContextualMessage(exception.Message, context),
            TechnicalDetails = exception.ToString(),
            Suggestion = "Review the error details and ensure all prerequisites are met.",
            ErrorCode = exception.HResult
        };
    }

    /// <summary>
    /// Checks if an exception is related to object not found errors.
    /// </summary>
    /// <param name="exception">The exception to check.</param>
    /// <returns>True if the exception indicates an object was not found.</returns>
    public static bool IsObjectNotFoundException(Exception exception)
    {
        if (exception is COMException comEx)
        {
            var mapped = MapException(comEx);
            return mapped.Category == ErrorCategory.ObjectNotFound;
        }
        return false;
    }

    /// <summary>
    /// Checks if an exception is related to connection issues.
    /// </summary>
    /// <param name="exception">The exception to check.</param>
    /// <returns>True if the exception indicates a connection problem.</returns>
    public static bool IsConnectionException(Exception exception)
    {
        if (exception is COMException comEx)
        {
            var mapped = MapException(comEx);
            return mapped.Category == ErrorCategory.Connection;
        }
        return false;
    }

    /// <summary>
    /// Checks if an exception is related to session state issues.
    /// </summary>
    /// <param name="exception">The exception to check.</param>
    /// <returns>True if the exception indicates a session state problem.</returns>
    public static bool IsSessionStateException(Exception exception)
    {
        if (exception is COMException comEx)
        {
            var mapped = MapException(comEx);
            return mapped.Category == ErrorCategory.SessionState;
        }
        return false;
    }

    /// <summary>
    /// Gets a concise error message for logging.
    /// </summary>
    /// <param name="exception">The exception.</param>
    /// <returns>Concise error message.</returns>
    public static string GetConciseMessage(Exception exception)
    {
        var mapped = MapException(exception);
        return $"[{mapped.Category}] {mapped.Message}";
    }

    /// <summary>
    /// Gets a detailed error message for API responses.
    /// </summary>
    /// <param name="exception">The exception.</param>
    /// <param name="includeStackTrace">Whether to include stack trace.</param>
    /// <returns>Detailed error message.</returns>
    public static string GetDetailedMessage(Exception exception, bool includeStackTrace = false)
    {
        var mapped = MapException(exception);
        var message = $"{mapped.Message}\n\nSuggestion: {mapped.Suggestion}";
        
        if (includeStackTrace && exception.StackTrace != null)
        {
            message += $"\n\nStack Trace:\n{exception.StackTrace}";
        }
        
        return message;
    }

    private static string BuildContextualMessage(string baseMessage, string? context)
    {
        if (string.IsNullOrWhiteSpace(context))
            return baseMessage;

        return $"{baseMessage} (Context: {context})";
    }

    /// <summary>
    /// Creates a user-friendly error response object for API consumption.
    /// </summary>
    /// <param name="exception">The exception to convert.</param>
    /// <param name="context">Optional context.</param>
    /// <returns>Error response object.</returns>
    public static object CreateErrorResponse(Exception exception, string? context = null)
    {
        var mapped = MapException(exception, context);
        
        return new
        {
            error = mapped.Message,
            category = mapped.Category.ToString(),
            suggestion = mapped.Suggestion,
            technicalDetails = mapped.TechnicalDetails,
            errorCode = $"0x{mapped.ErrorCode:X8}"
        };
    }
}

