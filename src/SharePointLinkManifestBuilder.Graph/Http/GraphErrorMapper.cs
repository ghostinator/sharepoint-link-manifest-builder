using System.Net;
using System.Text.Json;
using SharePointLinkManifestBuilder.Core.Models;

namespace SharePointLinkManifestBuilder.Graph.Http;

/// <summary>
/// Turns an HTTP failure into a normalized <see cref="GraphError"/> with a plain-language
/// explanation.
/// <para>
/// This is the single place raw service output is interpreted. Nothing downstream ever sees a
/// response body, so no token, header, or payload fragment can reach a log or the UI through
/// an error path.
/// </para>
/// </summary>
public static class GraphErrorMapper
{
    /// <summary>Maps a status code, Graph error code and operation into a normalized error.</summary>
    /// <param name="statusCode">HTTP status code.</param>
    /// <param name="graphErrorCode">The <c>error.code</c> value, when the body supplied one.</param>
    /// <param name="graphMessage">The <c>error.message</c> value, used only for classification.</param>
    /// <param name="operation">A short description of what was attempted.</param>
    /// <param name="clientRequestId">Correlation ID sent with the request.</param>
    /// <param name="serviceRequestId">Request ID returned by the service.</param>
    public static GraphError Map(
        int statusCode,
        string? graphErrorCode,
        string? graphMessage,
        string operation,
        string? clientRequestId = null,
        string? serviceRequestId = null)
    {
        var (kind, message, action, retryable) =
            Classify(statusCode, graphErrorCode, graphMessage, operation);

        return new GraphError
        {
            Kind = kind,
            Message = message,
            StatusCode = statusCode,
            GraphErrorCode = graphErrorCode,
            ClientRequestId = clientRequestId,
            ServiceRequestId = serviceRequestId,
            IsRetryable = retryable,
            SuggestedAction = action,
        };
    }

    /// <summary>Maps a transport-level exception, where no HTTP response was received.</summary>
    public static GraphError MapException(Exception exception, string operation)
    {
        ArgumentNullException.ThrowIfNull(exception);

        return exception switch
        {
            // Order matters: TaskCanceledException derives from OperationCanceledException.
            // Since .NET 6 an HttpClient timeout surfaces as a TaskCanceledException whose
            // inner exception is a TimeoutException, which is how a timeout is told apart
            // from a genuine user cancellation.
            TaskCanceledException { InnerException: TimeoutException } => new GraphError
            {
                Kind = GraphErrorKind.NetworkOutage,
                Message = $"The request timed out while trying to {operation}.",
                IsRetryable = true,
                SuggestedAction = "Check your network connection and try again.",
            },

            OperationCanceledException => GraphError.Canceled(),

            HttpRequestException http => new GraphError
            {
                Kind = GraphErrorKind.NetworkOutage,
                Message = $"Could not reach Microsoft Graph while trying to {operation}.",
                StatusCode = http.StatusCode is { } code ? (int)code : null,
                IsRetryable = true,
                SuggestedAction = "Check your network connection, proxy settings and firewall, then try again.",
            },

            JsonException => new GraphError
            {
                Kind = GraphErrorKind.Unknown,
                Message = $"Microsoft Graph returned a response that could not be understood while trying to {operation}.",
                IsRetryable = true,
            },

            _ => new GraphError
            {
                Kind = GraphErrorKind.Unknown,

                // The exception type is safe to surface; its message may embed request details,
                // so it is deliberately not included.
                Message = $"An unexpected {exception.GetType().Name} occurred while trying to {operation}.",
                IsRetryable = false,
            },
        };
    }

    private static (GraphErrorKind Kind, string Message, string? Action, bool Retryable) Classify(
        int statusCode,
        string? code,
        string? graphMessage,
        string operation)
    {
        code ??= string.Empty;
        var text = graphMessage ?? string.Empty;

        // Sharing and policy decisions are the ones users most need explained precisely, so
        // they are matched on the Graph error code before falling back to the status code.
        if (Contains(code, "notAllowed") || Contains(text, "sharing is disabled")
            || Contains(text, "external sharing"))
        {
            return (GraphErrorKind.ExternalSharingDisabled,
                "Your organization's sharing policy does not allow this kind of link for this location.",
                "Choose a narrower audience, such as 'People in the organization', or ask an administrator "
                + "about the external sharing policy for this site.",
                false);
        }

        if (Contains(text, "anonymous") && (Contains(text, "disabled") || Contains(text, "not allowed")))
        {
            return (GraphErrorKind.AnonymousSharingDisabled,
                "Anonymous 'Anyone with the link' sharing is disabled for this location.",
                "Choose 'People in the organization' or 'Specific people' instead.",
                false);
        }

        return statusCode switch
        {
            400 when Contains(code, "invalidRequest") && Contains(text, "recipient") =>
                (GraphErrorKind.RecipientRejected,
                    "Microsoft 365 rejected one or more of the recipients.",
                    "Check the recipient addresses and whether guest sharing is permitted.", false),

            400 => (GraphErrorKind.UnsupportedLinkType,
                $"Microsoft Graph rejected the request to {operation} as invalid.",
                "Check the requested link type and audience; some combinations are not supported for this item.",
                false),

            401 => (GraphErrorKind.AuthenticationFailed,
                "Your sign-in is no longer valid.",
                "Sign in again from the Microsoft 365 connection settings.", false),

            403 when Contains(code, "accessDenied") && Contains(text, "consent") =>
                (GraphErrorKind.ConsentRequired,
                    "This action needs a permission that has not been granted.",
                    "Open Settings, then Microsoft 365 Connection, then Permissions, and request the missing consent.",
                    false),

            403 when Contains(text, "quota") || Contains(text, "not provisioned") =>
                (GraphErrorKind.UserDriveUnprovisioned,
                    "That user's OneDrive has not been set up yet.",
                    "The user must open OneDrive once so it is provisioned. This application never provisions it for them.",
                    false),

            403 => (GraphErrorKind.SharePointAccessDenied,
                $"You do not have permission to {operation}.",
                "Delegated access is limited to what your own account can already open. Ask a site owner for access.",
                false),

            404 when Contains(operation, "site") =>
                (GraphErrorKind.SiteNotFound, "That SharePoint site could not be found.",
                    "Check the URL, or that you have access to the site.", false),

            404 when Contains(operation, "drive") || Contains(operation, "library") =>
                (GraphErrorKind.LibraryNotFound, "That document library or drive could not be found.", null, false),

            404 when Contains(operation, "folder") =>
                (GraphErrorKind.FolderNotFound, "That folder could not be found.",
                    "It may have been moved or deleted since it was selected.", false),

            404 => (GraphErrorKind.FileDeletedDuringProcessing,
                "The item could not be found. It may have been deleted or moved while the job was running.",
                null, false),

            409 => (GraphErrorKind.NameConflict,
                "An item with that name already exists at the destination.", null, false),

            412 => (GraphErrorKind.ETagConflict,
                "The file changed in SharePoint after this application read it, so the write was refused.",
                "Re-run the job so the latest version is read before writing.", false),

            423 => (GraphErrorKind.ManifestWriteDenied,
                "The file is locked, most likely because it is open elsewhere.",
                "Close the file and try again.", true),

            429 => (GraphErrorKind.Throttled,
                "Microsoft 365 is asking this application to slow down.",
                "The job will wait and retry automatically. Lowering the concurrency setting can help.", true),

            >= 500 and < 600 => (GraphErrorKind.ServiceUnavailable,
                "Microsoft 365 is temporarily unavailable.",
                "The job will retry automatically.", true),

            _ => (GraphErrorKind.Unknown,
                $"Microsoft Graph returned HTTP {statusCode} while trying to {operation}.",
                null, statusCode >= 500),
        };
    }

    private static bool Contains(string haystack, string needle) =>
        haystack.Contains(needle, StringComparison.OrdinalIgnoreCase);

    /// <summary>
    /// Extracts <c>error.code</c> and <c>error.message</c> from a Graph error body. Returns
    /// nulls rather than throwing when the body is absent or malformed, because an error path
    /// must never fail in a second way.
    /// </summary>
    public static (string? Code, string? Message) TryReadErrorBody(string? body)
    {
        if (string.IsNullOrWhiteSpace(body))
        {
            return (null, null);
        }

        try
        {
            using var document = JsonDocument.Parse(body);
            if (!document.RootElement.TryGetProperty("error", out var error))
            {
                return (null, null);
            }

            var code = error.TryGetProperty("code", out var c) ? c.GetString() : null;
            var message = error.TryGetProperty("message", out var m) ? m.GetString() : null;
            return (code, message);
        }
        catch (JsonException)
        {
            return (null, null);
        }
    }

    /// <summary>Maps an <see cref="HttpStatusCode"/> convenience overload.</summary>
    public static GraphError Map(HttpStatusCode statusCode, string operation) =>
        Map((int)statusCode, null, null, operation);
}
