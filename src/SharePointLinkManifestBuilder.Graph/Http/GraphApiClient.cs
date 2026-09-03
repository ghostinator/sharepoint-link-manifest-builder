using System.Globalization;
using System.Net;
using System.Net.Http.Headers;
using System.Net.Http.Json;
using System.Runtime.CompilerServices;
using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;
using Microsoft.Extensions.Logging;
using SharePointLinkManifestBuilder.Core.Abstractions;
using SharePointLinkManifestBuilder.Core.Models;
using SharePointLinkManifestBuilder.Core.Resilience;
using SharePointLinkManifestBuilder.Core.Security;

namespace SharePointLinkManifestBuilder.Graph.Http;

/// <summary>
/// The single Microsoft Graph transport for the whole application.
/// <para>
/// Centralizing every call here is what makes the cross-cutting behaviour testable and
/// consistent: authentication, correlation IDs, pagination, throttling and retry, ETag
/// concurrency, error normalization and sanitized logging all happen in exactly one place.
/// </para>
/// <para>
/// Tests drive this class through a fake <see cref="HttpMessageHandler"/>, so the real
/// pagination loop and real retry policy are exercised rather than replaced by a mock.
/// </para>
/// </summary>
public sealed class GraphApiClient : IGraphApiClient
{
    /// <summary>Header Microsoft Graph uses for client-supplied correlation.</summary>
    public const string ClientRequestIdHeader = "client-request-id";

    /// <summary>Header Microsoft Graph returns identifying the request server-side.</summary>
    public const string ServiceRequestIdHeader = "request-id";

    internal static readonly JsonSerializerOptions SerializerOptions = new()
    {
        PropertyNameCaseInsensitive = true,
        DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull,
    };

    private readonly HttpClient _httpClient;
    private readonly IAuthenticationService _authentication;
    private readonly GraphClientContext _context;
    private readonly GraphRetryPolicy _retryPolicy;
    private readonly ILogger<GraphApiClient> _logger;
    private readonly TimeProvider _timeProvider;

    /// <summary>Creates the transport.</summary>
    /// <param name="httpClient">The HTTP client, supplied by <c>IHttpClientFactory</c>.</param>
    /// <param name="authentication">Supplies bearer tokens. Never asked for a token to log.</param>
    /// <param name="context">Current endpoint and scopes.</param>
    /// <param name="retryPolicy">Throttling and transient-failure policy.</param>
    /// <param name="logger">Structured logger. Everything written passes through redaction.</param>
    /// <param name="timeProvider">Clock, injected so retry waits are testable.</param>
    public GraphApiClient(
        HttpClient httpClient,
        IAuthenticationService authentication,
        GraphClientContext context,
        GraphRetryPolicy retryPolicy,
        ILogger<GraphApiClient> logger,
        TimeProvider? timeProvider = null)
    {
        _httpClient = httpClient ?? throw new ArgumentNullException(nameof(httpClient));
        _authentication = authentication ?? throw new ArgumentNullException(nameof(authentication));
        _context = context ?? throw new ArgumentNullException(nameof(context));
        _retryPolicy = retryPolicy ?? throw new ArgumentNullException(nameof(retryPolicy));
        _logger = logger ?? throw new ArgumentNullException(nameof(logger));
        _timeProvider = timeProvider ?? TimeProvider.System;
    }

    /// <inheritdoc />
    public Task<GraphResponse<T>> GetAsync<T>(string relativeUrl, CancellationToken cancellationToken = default) =>
        SendJsonAsync<T>(HttpMethod.Get, relativeUrl, null, null, $"read {Describe(relativeUrl)}", cancellationToken);

    /// <inheritdoc />
    public Task<GraphResponse<TResponse>> PostAsync<TResponse>(
        string relativeUrl,
        object body,
        CancellationToken cancellationToken = default) =>
        SendJsonAsync<TResponse>(
            HttpMethod.Post, relativeUrl, body, null, $"submit {Describe(relativeUrl)}", cancellationToken);

    /// <inheritdoc />
    public Task<GraphResponse<TResponse>> PatchAsync<TResponse>(
        string relativeUrl,
        object body,
        CancellationToken cancellationToken = default) =>
        SendJsonAsync<TResponse>(
            HttpMethod.Patch, relativeUrl, body, null, $"update {Describe(relativeUrl)}", cancellationToken);

    /// <inheritdoc />
    public async Task<GraphResponse<bool>> DeleteAsync(
        string relativeUrl,
        CancellationToken cancellationToken = default)
    {
        var response = await SendJsonAsync<JsonElement>(
            HttpMethod.Delete, relativeUrl, null, null, $"delete {Describe(relativeUrl)}", cancellationToken)
            .ConfigureAwait(false);

        return new GraphResponse<bool>
        {
            Succeeded = response.Succeeded,
            Value = response.Succeeded,
            StatusCode = response.StatusCode,
            Error = response.Error,
            ClientRequestId = response.ClientRequestId,
        };
    }

    /// <inheritdoc />
    public Task<GraphResponse<TResponse>> PutContentAsync<TResponse>(
        string relativeUrl,
        ReadOnlyMemory<byte> content,
        string contentType,
        string? ifMatchETag = null,
        CancellationToken cancellationToken = default) =>
        SendAsync<TResponse>(
            () =>
            {
                var request = new HttpRequestMessage(HttpMethod.Put, BuildUri(relativeUrl))
                {
                    Content = new ReadOnlyMemoryContent(content),
                };

                // Parse, not the constructor: MediaTypeHeaderValue(string) rejects anything
                // carrying parameters, so "text/plain; charset=utf-8" throws a FormatException.
                request.Content.Headers.ContentType = MediaTypeHeaderValue.Parse(contentType);

                // A conditional write is the only thing standing between this application and
                // silently clobbering a file somebody changed a moment ago.
                if (!string.IsNullOrEmpty(ifMatchETag))
                {
                    request.Headers.TryAddWithoutValidation("If-Match", ifMatchETag);
                }

                return request;
            },
            $"upload {Describe(relativeUrl)}",
            authenticate: true,
            cancellationToken);

    /// <inheritdoc />
    public async Task<GraphResponse<byte[]>> GetContentAsync(
        string relativeUrl,
        CancellationToken cancellationToken = default)
    {
        var operation = $"download {Describe(relativeUrl)}";

        return await ExecuteWithRetryAsync(
            () => new HttpRequestMessage(HttpMethod.Get, BuildUri(relativeUrl)),
            operation,
            authenticate: true,
            async (response, ct) =>
            {
                var bytes = await response.Content.ReadAsByteArrayAsync(ct).ConfigureAwait(false);
                return (bytes, (string?)response.Headers.ETag?.Tag);
            },
            cancellationToken).ConfigureAwait(false);
    }

    /// <inheritdoc />
    public async IAsyncEnumerable<T> EnumeratePagedAsync<T>(
        string relativeUrl,
        [EnumeratorCancellation] CancellationToken cancellationToken = default)
    {
        var next = relativeUrl;
        var pageNumber = 0;

        while (!string.IsNullOrEmpty(next))
        {
            cancellationToken.ThrowIfCancellationRequested();

            var page = await SendJsonAsync<GraphPage<T>>(
                HttpMethod.Get, next, null, null, $"list {Describe(relativeUrl)}", cancellationToken)
                .ConfigureAwait(false);

            if (!page.Succeeded || page.Value is null)
            {
                // Surfacing the failure as an exception is deliberate: silently ending the
                // sequence would look identical to "the folder is empty", and a job would then
                // report success while having enumerated nothing.
                throw new GraphOperationException(
                    page.Error ?? GraphErrorMapper.Map(page.StatusCode, null, null, "list items"));
            }

            pageNumber++;
            foreach (var item in page.Value.Value)
            {
                yield return item;
            }

            next = page.Value.NextLink;

            if (!string.IsNullOrEmpty(next) && _logger.IsEnabled(LogLevel.Debug))
            {
                _logger.LogDebug(
                    "Following Graph pagination to page {PageNumber} for {Operation}.",
                    pageNumber + 1,
                    Describe(relativeUrl));
            }
        }
    }

    /// <inheritdoc />
    public Task<GraphResponse<bool>> PutUploadSessionChunkAsync(
        Uri uploadUrl,
        ReadOnlyMemory<byte> chunk,
        ContentRangeHeaderValue range,
        long totalLength,
        CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(uploadUrl);
        ArgumentNullException.ThrowIfNull(range);

        return SendAsync<bool>(
            () =>
            {
                var request = new HttpRequestMessage(HttpMethod.Put, uploadUrl)
                {
                    Content = new ReadOnlyMemoryContent(chunk),
                };

                request.Content.Headers.ContentLength = chunk.Length;
                request.Content.Headers.ContentRange = range;
                return request;
            },
            "upload a manifest chunk",

            // An upload session URL is pre-authorized and carries its own credentials in the
            // query string. Attaching a bearer token is unnecessary and would leak it wider.
            authenticate: false,
            cancellationToken);
    }

    private async Task<GraphResponse<T>> SendJsonAsync<T>(
        HttpMethod method,
        string relativeUrl,
        object? body,
        string? ifMatchETag,
        string operation,
        CancellationToken cancellationToken) =>
        await SendAsync<T>(
            () =>
            {
                var request = new HttpRequestMessage(method, BuildUri(relativeUrl));

                if (body is not null)
                {
                    request.Content = JsonContent.Create(body, options: SerializerOptions);
                }

                if (!string.IsNullOrEmpty(ifMatchETag))
                {
                    request.Headers.TryAddWithoutValidation("If-Match", ifMatchETag);
                }

                return request;
            },
            operation,
            authenticate: true,
            cancellationToken).ConfigureAwait(false);

    private Task<GraphResponse<T>> SendAsync<T>(
        Func<HttpRequestMessage> requestFactory,
        string operation,
        bool authenticate,
        CancellationToken cancellationToken) =>
        ExecuteWithRetryAsync(
            requestFactory,
            operation,
            authenticate,
            async (response, ct) =>
            {
                var etag = response.Headers.ETag?.Tag;

                if (response.StatusCode == HttpStatusCode.NoContent || response.Content.Headers.ContentLength == 0)
                {
                    return (default(T), etag);
                }

                // bool is used as a "did it work" payload for calls with no meaningful body.
                if (typeof(T) == typeof(bool))
                {
                    return ((T)(object)true, etag);
                }

                var value = await response.Content
                    .ReadFromJsonAsync<T>(SerializerOptions, ct)
                    .ConfigureAwait(false);

                return (value, etag);
            },
            cancellationToken);

    private async Task<GraphResponse<T>> ExecuteWithRetryAsync<T>(
        Func<HttpRequestMessage> requestFactory,
        string operation,
        bool authenticate,
        Func<HttpResponseMessage, CancellationToken, Task<(T? Value, string? ETag)>> readPayload,
        CancellationToken cancellationToken)
    {
        var attempt = 0;

        while (true)
        {
            attempt++;
            cancellationToken.ThrowIfCancellationRequested();

            var correlationId = Guid.NewGuid().ToString("d");
            using var request = requestFactory();
            request.Headers.TryAddWithoutValidation(ClientRequestIdHeader, correlationId);

            if (authenticate)
            {
                var token = await _authentication
                    .GetAccessTokenAsync(_context.Scopes, cancellationToken)
                    .ConfigureAwait(false);

                if (string.IsNullOrEmpty(token))
                {
                    return Failure<T>(
                        new GraphError
                        {
                            Kind = GraphErrorKind.AuthenticationFailed,
                            Message = $"No valid sign-in is available, so the application could not {operation}.",
                            SuggestedAction = "Sign in again from the Microsoft 365 connection settings.",
                            ClientRequestId = correlationId,
                        },
                        correlationId);
                }

                request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
            }

            HttpResponseMessage? response = null;

            try
            {
                // The URL is logged with its query string stripped, so no parameter of any kind
                // can reach the log through this path. Redaction is skipped entirely when debug
                // logging is off, so it costs nothing on the hot path.
                if (_logger.IsEnabled(LogLevel.Debug))
                {
                    _logger.LogDebug(
                        "Graph {Method} {Url} (attempt {Attempt}, correlation {CorrelationId}).",
                        request.Method.Method,
                        SensitiveDataRedactor.RedactUrl(request.RequestUri?.ToString()),
                        attempt,
                        correlationId);
                }

                response = await _httpClient
                    .SendAsync(request, HttpCompletionOption.ResponseHeadersRead, cancellationToken)
                    .ConfigureAwait(false);

                var status = (int)response.StatusCode;
                var serviceRequestId = ReadHeader(response, ServiceRequestIdHeader);

                if (response.IsSuccessStatusCode || status == 207)
                {
                    var (value, etag) = await readPayload(response, cancellationToken).ConfigureAwait(false);

                    return new GraphResponse<T>
                    {
                        Succeeded = true,
                        Value = value,
                        StatusCode = status,
                        ETag = etag,
                        ClientRequestId = correlationId,
                    };
                }

                var body = await SafeReadBodyAsync(response, cancellationToken).ConfigureAwait(false);
                var (code, message) = GraphErrorMapper.TryReadErrorBody(body);

                var decision = _retryPolicy.Evaluate(attempt, status, ReadRetryAfter(response));

                if (!decision.ShouldRetry)
                {
                    var error = GraphErrorMapper.Map(
                        status, code, message, operation, correlationId, serviceRequestId);

                    _logger.LogWarning(
                        "Graph request failed: {Operation} returned HTTP {Status} ({GraphCode}). "
                        + "Correlation {CorrelationId}, service request {ServiceRequestId}.",
                        operation, status, code ?? "none", correlationId, serviceRequestId ?? "none");

                    return Failure<T>(error, correlationId);
                }

                // Guarded because the structured arguments box value types; on a throttled
                // run this path can be hit thousands of times.
                if (_logger.IsEnabled(LogLevel.Information))
                {
                    _logger.LogInformation(
                        "Graph request will be retried: {Operation} returned HTTP {Status}. {Reason} "
                        + "Waiting {Delay} before attempt {NextAttempt}.",
                        operation, status, decision.Reason, decision.Delay, attempt + 1);
                }

                await Task.Delay(decision.Delay, _timeProvider, cancellationToken).ConfigureAwait(false);
            }
            catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested)
            {
                throw;
            }
            catch (Exception ex) when (ex is HttpRequestException or TaskCanceledException or IOException)
            {
                var decision = _retryPolicy.EvaluateTransport(attempt, isTransient: true);

                if (!decision.ShouldRetry)
                {
                    _logger.LogWarning(
                        ex, "Graph request failed permanently: {Operation}. {Reason}", operation, decision.Reason);

                    return Failure<T>(GraphErrorMapper.MapException(ex, operation), correlationId);
                }

                if (_logger.IsEnabled(LogLevel.Information))
                {
                    _logger.LogInformation(
                        "Transient network failure during {Operation}; retrying in {Delay}.",
                        operation, decision.Delay);
                }

                await Task.Delay(decision.Delay, _timeProvider, cancellationToken).ConfigureAwait(false);
            }
            finally
            {
                response?.Dispose();
            }
        }
    }

    private static GraphResponse<T> Failure<T>(GraphError error, string correlationId) => new()
    {
        Succeeded = false,
        StatusCode = error.StatusCode ?? 0,
        Error = error,
        ClientRequestId = correlationId,
    };

    /// <summary>
    /// Reads an error body without ever letting a failure here mask the original failure.
    /// </summary>
    private static async Task<string?> SafeReadBodyAsync(
        HttpResponseMessage response,
        CancellationToken cancellationToken)
    {
        try
        {
            return await response.Content.ReadAsStringAsync(cancellationToken).ConfigureAwait(false);
        }
        catch (Exception) when (!cancellationToken.IsCancellationRequested)
        {
            return null;
        }
    }

    /// <summary>
    /// Reads <c>Retry-After</c>, which Graph may send either as a delay in seconds or as an
    /// HTTP date. Both forms are honoured.
    /// </summary>
    internal static TimeSpan? ReadRetryAfter(HttpResponseMessage response)
    {
        var retryAfter = response.Headers.RetryAfter;

        if (retryAfter is null)
        {
            return null;
        }

        if (retryAfter.Delta is { } delta)
        {
            return delta;
        }

        if (retryAfter.Date is { } date)
        {
            var wait = date - DateTimeOffset.UtcNow;
            return wait > TimeSpan.Zero ? wait : TimeSpan.Zero;
        }

        return null;
    }

    private static string? ReadHeader(HttpResponseMessage response, string name) =>
        response.Headers.TryGetValues(name, out var values) ? values.FirstOrDefault() : null;

    private Uri BuildUri(string relativeOrAbsolute)
    {
        // A nextLink is absolute; everything else is relative to the configured endpoint.
        //
        // The scheme is checked explicitly rather than relying on Uri.TryCreate with
        // UriKind.Absolute. On macOS and Linux "/me" IS a valid absolute URI: it parses as the
        // file scheme. Trusting TryCreate here turned every relative Graph path into
        // "file:///me" on those platforms while working correctly on Windows.
        if (relativeOrAbsolute.StartsWith("https://", StringComparison.OrdinalIgnoreCase)
            || relativeOrAbsolute.StartsWith("http://", StringComparison.OrdinalIgnoreCase))
        {
            return new Uri(relativeOrAbsolute, UriKind.Absolute);
        }

        var path = relativeOrAbsolute.StartsWith('/') ? relativeOrAbsolute : "/" + relativeOrAbsolute;
        return new Uri(_context.Endpoint + path, UriKind.Absolute);
    }

    /// <summary>
    /// Produces a short, non-identifying description of an operation for logs and error text.
    /// Concrete IDs are deliberately dropped: a log line should say "read a drive item", not
    /// name the tenant's item.
    /// </summary>
    private static string Describe(string relativeUrl)
    {
        var path = relativeUrl.Split('?')[0].Trim('/');
        var segments = path.Split('/', StringSplitOptions.RemoveEmptyEntries);

        if (segments.Length == 0)
        {
            return "a Microsoft Graph resource";
        }

        var noun = segments[0].ToLowerInvariant() switch
        {
            "sites" => "a SharePoint site",

            // A missing drive and a missing item inside a drive are different failures, so the
            // wording has to distinguish them for the error mapper.
            "drives" => path.Contains("/items/", StringComparison.OrdinalIgnoreCase)
                ? "a drive item"
                : "a drive",
            "me" => "your profile",
            "users" => "a user",
            "shares" => "a shared item",
            "applications" => "an application registration",
            "serviceprincipals" => "a service principal",
            "organization" => "the organization profile",
            "oauth2permissiongrants" => "a permission grant",
            _ => "a Microsoft Graph resource",
        };

        if (path.Contains("createLink", StringComparison.OrdinalIgnoreCase))
        {
            return "create a sharing link";
        }

        if (path.Contains("invite", StringComparison.OrdinalIgnoreCase))
        {
            return "grant access to recipients";
        }

        if (path.Contains("children", StringComparison.OrdinalIgnoreCase))
        {
            return "list the contents of a folder";
        }

        if (path.Contains("createUploadSession", StringComparison.OrdinalIgnoreCase))
        {
            return "start a manifest upload";
        }

        return noun;
    }
}

/// <summary>One page of an OData collection.</summary>
/// <typeparam name="T">Element type.</typeparam>
internal sealed record GraphPage<T>
{
    /// <summary>The items on this page.</summary>
    [JsonPropertyName("value")]
    public IReadOnlyList<T> Value { get; init; } = [];

    /// <summary>Absolute URL of the next page, or null when this is the last page.</summary>
    [JsonPropertyName("@odata.nextLink")]
    public string? NextLink { get; init; }
}

/// <summary>
/// Thrown when a streaming enumeration cannot continue. Streaming has no other way to report a
/// failure, and ending the sequence quietly would be indistinguishable from an empty folder.
/// </summary>
public sealed class GraphOperationException : Exception
{
    /// <summary>Creates the exception from a normalized error.</summary>
    public GraphOperationException(GraphError error)
        : base(error?.Message ?? "A Microsoft Graph operation failed.") =>
        Error = error ?? throw new ArgumentNullException(nameof(error));

    /// <summary>Creates the exception with a message only.</summary>
    public GraphOperationException(string message)
        : base(message) =>
        Error = new GraphError { Kind = GraphErrorKind.Unknown, Message = message };

    /// <summary>Creates the exception with a message and inner exception.</summary>
    public GraphOperationException(string message, Exception innerException)
        : base(message, innerException) =>
        Error = new GraphError { Kind = GraphErrorKind.Unknown, Message = message };

    /// <summary>Creates the exception with no detail.</summary>
    public GraphOperationException()
        : this("A Microsoft Graph operation failed.")
    {
    }

    /// <summary>The normalized error.</summary>
    public GraphError Error { get; }
}
