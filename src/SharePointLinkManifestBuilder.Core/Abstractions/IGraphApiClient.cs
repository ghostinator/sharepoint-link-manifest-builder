using System.Net.Http.Headers;
using SharePointLinkManifestBuilder.Core.Models;

namespace SharePointLinkManifestBuilder.Core.Abstractions;

/// <summary>A Microsoft Graph response, normalized and stripped of sensitive material.</summary>
/// <typeparam name="T">The deserialized payload type.</typeparam>
public sealed record GraphResponse<T>
{
    /// <summary>True when the request produced a usable payload.</summary>
    public required bool Succeeded { get; init; }

    /// <summary>The deserialized payload.</summary>
    public T? Value { get; init; }

    /// <summary>
    /// The HTTP status code. Product-significant for <c>createLink</c>, where 201 means a new
    /// link was created and 200 means an equivalent link already existed.
    /// </summary>
    public int StatusCode { get; init; }

    /// <summary>The response ETag, when present.</summary>
    public string? ETag { get; init; }

    /// <summary>The failure, when the request did not succeed.</summary>
    public GraphError? Error { get; init; }

    /// <summary>Correlation ID sent with the request.</summary>
    public string? ClientRequestId { get; init; }

    /// <summary>True when the service reported that the resource already existed (HTTP 200).</summary>
    public bool IsExistingResource => StatusCode == 200;

    /// <summary>True when the service created a new resource (HTTP 201).</summary>
    public bool IsCreatedResource => StatusCode == 201;

    /// <summary>True when only some sub-operations succeeded (HTTP 207).</summary>
    public bool IsPartialSuccess => StatusCode == 207;
}

/// <summary>
/// The single Microsoft Graph transport. Every Graph call in the application goes through
/// this interface, which centralizes authentication, correlation IDs, pagination, retry and
/// throttling, error normalization and sanitized logging.
/// <para>
/// No view model may reference this type's implementation or <c>HttpClient</c> directly.
/// </para>
/// </summary>
public interface IGraphApiClient
{
    /// <summary>Issues a GET and deserializes the payload.</summary>
    /// <typeparam name="T">Payload type.</typeparam>
    /// <param name="relativeUrl">Path relative to the Graph endpoint, for example <c>/me</c>.</param>
    /// <param name="cancellationToken">Cancellation token.</param>
    Task<GraphResponse<T>> GetAsync<T>(string relativeUrl, CancellationToken cancellationToken = default);

    /// <summary>Issues a POST with a JSON body.</summary>
    /// <typeparam name="TResponse">Response payload type.</typeparam>
    /// <param name="relativeUrl">Path relative to the Graph endpoint.</param>
    /// <param name="body">Body serialized as JSON.</param>
    /// <param name="cancellationToken">Cancellation token.</param>
    Task<GraphResponse<TResponse>> PostAsync<TResponse>(
        string relativeUrl,
        object body,
        CancellationToken cancellationToken = default);

    /// <summary>Issues a PATCH with a JSON body.</summary>
    /// <typeparam name="TResponse">Response payload type.</typeparam>
    /// <param name="relativeUrl">Path relative to the Graph endpoint.</param>
    /// <param name="body">Body serialized as JSON.</param>
    /// <param name="cancellationToken">Cancellation token.</param>
    Task<GraphResponse<TResponse>> PatchAsync<TResponse>(
        string relativeUrl,
        object body,
        CancellationToken cancellationToken = default);

    /// <summary>Issues a DELETE. Used only by the explicitly guarded registration deletion.</summary>
    Task<GraphResponse<bool>> DeleteAsync(string relativeUrl, CancellationToken cancellationToken = default);

    /// <summary>
    /// Uploads content, optionally guarded by an ETag. Supplying <paramref name="ifMatchETag"/>
    /// makes the write conditional, so a remotely modified file cannot be silently overwritten.
    /// </summary>
    /// <typeparam name="TResponse">Response payload type.</typeparam>
    /// <param name="relativeUrl">Path relative to the Graph endpoint.</param>
    /// <param name="content">Bytes to upload.</param>
    /// <param name="contentType">MIME type.</param>
    /// <param name="ifMatchETag">ETag for a conditional write, or null for unconditional.</param>
    /// <param name="cancellationToken">Cancellation token.</param>
    Task<GraphResponse<TResponse>> PutContentAsync<TResponse>(
        string relativeUrl,
        ReadOnlyMemory<byte> content,
        string contentType,
        string? ifMatchETag = null,
        CancellationToken cancellationToken = default);

    /// <summary>Downloads raw content, for reading an existing manifest.</summary>
    Task<GraphResponse<byte[]>> GetContentAsync(
        string relativeUrl,
        CancellationToken cancellationToken = default);

    /// <summary>
    /// Enumerates a paged collection, following <c>@odata.nextLink</c> until it is absent.
    /// Yields lazily so a large library never has to be held in memory at once.
    /// </summary>
    /// <typeparam name="T">Element type.</typeparam>
    /// <param name="relativeUrl">First page URL, relative to the Graph endpoint.</param>
    /// <param name="cancellationToken">Cancellation token, honoured between pages.</param>
    IAsyncEnumerable<T> EnumeratePagedAsync<T>(
        string relativeUrl,
        CancellationToken cancellationToken = default);

    /// <summary>
    /// Uploads to a pre-authorized upload-session URL. The session URL carries its own
    /// authorization, so no bearer token is attached.
    /// </summary>
    /// <param name="uploadUrl">Absolute upload session URL returned by Graph.</param>
    /// <param name="chunk">The bytes for this chunk.</param>
    /// <param name="range">Content range describing the chunk's position.</param>
    /// <param name="totalLength">Total length of the upload.</param>
    /// <param name="cancellationToken">Cancellation token.</param>
    Task<GraphResponse<bool>> PutUploadSessionChunkAsync(
        Uri uploadUrl,
        ReadOnlyMemory<byte> chunk,
        ContentRangeHeaderValue range,
        long totalLength,
        CancellationToken cancellationToken = default);
}
