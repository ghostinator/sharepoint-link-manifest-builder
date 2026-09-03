using System.Net.Http.Headers;
using System.Text;
using Microsoft.Extensions.Logging;
using SharePointLinkManifestBuilder.Core.Abstractions;
using SharePointLinkManifestBuilder.Core.Models;
using SharePointLinkManifestBuilder.Graph.Dto;
using SharePointLinkManifestBuilder.Graph.Http;

namespace SharePointLinkManifestBuilder.Graph.Services;

/// <summary>
/// Reads and writes manifest files in SharePoint and OneDrive.
/// <para>
/// Writes are conditional on an ETag wherever one is known, so a manifest that changed remotely
/// between the read and the write is refused rather than silently overwritten. The refusal
/// surfaces as a typed conflict for the caller's conflict policy to handle.
/// </para>
/// </summary>
public sealed class ManifestStorageService : IManifestStorageService
{
    /// <summary>
    /// Uploads at or below this size use a simple PUT. Graph's documented threshold for the
    /// simple upload API is 4 MiB.
    /// </summary>
    public const int SimpleUploadLimitBytes = 4 * 1024 * 1024;

    /// <summary>
    /// Chunk size for large uploads. Graph requires a multiple of 320 KiB; this is 10 x 320 KiB.
    /// </summary>
    public const int UploadChunkSizeBytes = 3_276_800;

    private static readonly UTF8Encoding Utf8NoBom = new(encoderShouldEmitUTF8Identifier: false);

    private readonly IGraphApiClient _client;
    private readonly ILogger<ManifestStorageService> _logger;

    /// <summary>Creates the service.</summary>
    public ManifestStorageService(IGraphApiClient client, ILogger<ManifestStorageService> logger)
    {
        _client = client ?? throw new ArgumentNullException(nameof(client));
        _logger = logger ?? throw new ArgumentNullException(nameof(logger));
    }

    /// <inheritdoc />
    public async Task<OperationResult<ExistingManifest?>> ReadManifestAsync(
        string driveId,
        string parentItemId,
        string fileName,
        CancellationToken cancellationToken = default)
    {
        var metadata = await _client
            .GetAsync<GraphDriveItemDto>(
                GraphPaths.ChildByName(driveId, parentItemId, fileName), cancellationToken)
            .ConfigureAwait(false);

        if (!metadata.Succeeded)
        {
            // "Not found" is the normal case for a first run, not a failure.
            if (metadata.StatusCode == 404)
            {
                return OperationResult<ExistingManifest?>.Success(null);
            }

            return OperationResult<ExistingManifest?>.Failure(
                metadata.Error ?? GraphErrorMapper.Map(metadata.StatusCode, null, null, "read an existing manifest"));
        }

        if (metadata.Value?.Id is null)
        {
            return OperationResult<ExistingManifest?>.Success(null);
        }

        var content = await _client
            .GetContentAsync(GraphPaths.ChildContentByName(driveId, parentItemId, fileName), cancellationToken)
            .ConfigureAwait(false);

        if (!content.Succeeded || content.Value is null)
        {
            return OperationResult<ExistingManifest?>.Failure(
                content.Error ?? GraphErrorMapper.Map(content.StatusCode, null, null, "download an existing manifest"));
        }

        return OperationResult<ExistingManifest?>.Success(new ExistingManifest
        {
            ItemId = metadata.Value.Id,
            Content = DecodeUtf8(content.Value),

            // The item ETag is preferred over the content tag: it changes on any modification,
            // which is exactly the condition a safe write needs to detect.
            ETag = metadata.Value.ETag,
            WebUrl = metadata.Value.WebUrl,
            LastModifiedUtc = metadata.Value.LastModifiedDateTime,
        });
    }

    /// <inheritdoc />
    public async Task<OperationResult<ManifestWriteResult>> WriteManifestAsync(
        string driveId,
        string parentItemId,
        string fileName,
        string content,
        ManifestFormats format,
        bool isMaster,
        int entryCount,
        string? ifMatchETag,
        CancellationToken cancellationToken = default)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(fileName);
        ArgumentNullException.ThrowIfNull(content);

        var bytes = Utf8NoBom.GetBytes(content);

        var result = bytes.Length <= SimpleUploadLimitBytes
            ? await WriteSmallAsync(driveId, parentItemId, fileName, bytes, ifMatchETag, cancellationToken)
                .ConfigureAwait(false)
            : await WriteLargeAsync(driveId, parentItemId, fileName, bytes, cancellationToken)
                .ConfigureAwait(false);

        if (!result.Succeeded)
        {
            return OperationResult<ManifestWriteResult>.Failure(result.Error!);
        }

        return OperationResult<ManifestWriteResult>.Success(new ManifestWriteResult
        {
            DisplayPath = fileName,
            WebUrl = result.Value,
            Format = format,
            IsMaster = isMaster,
            EntryCount = entryCount,
            Succeeded = true,
        });
    }

    private async Task<OperationResult<string?>> WriteSmallAsync(
        string driveId,
        string parentItemId,
        string fileName,
        byte[] bytes,
        string? ifMatchETag,
        CancellationToken cancellationToken)
    {
        var response = await _client
            .PutContentAsync<GraphDriveItemDto>(
                GraphPaths.ChildContentByName(driveId, parentItemId, fileName),
                bytes,
                ContentTypeFor(fileName),
                ifMatchETag,
                cancellationToken)
            .ConfigureAwait(false);

        if (response.Succeeded)
        {
            return OperationResult<string?>.Success(response.Value?.WebUrl);
        }

        // A 412 means the remote copy changed after it was read. Retrying blindly would either
        // fail identically or overwrite work that arrived in the meantime, so it is surfaced.
        if (response.StatusCode == 412)
        {
            _logger.LogWarning(
                "Manifest write refused: the file changed in SharePoint after it was read. "
                + "The existing manifest was left untouched.");

            return OperationResult<string?>.Failure(new GraphError
            {
                Kind = GraphErrorKind.ManifestConflict,
                Message = "The manifest changed in SharePoint after this application read it, so it was not written.",
                StatusCode = 412,
                SuggestedAction = "Run the job again so the latest version is read before writing.",
                IsRetryable = true,
            });
        }

        return OperationResult<string?>.Failure(
            response.Error ?? GraphErrorMapper.Map(response.StatusCode, null, null, "write a manifest"));
    }

    private async Task<OperationResult<string?>> WriteLargeAsync(
        string driveId,
        string parentItemId,
        string fileName,
        byte[] bytes,
        CancellationToken cancellationToken)
    {
        var sessionRequest = new CreateUploadSessionRequest
        {
            Item = new UploadSessionItemDto { ConflictBehavior = "replace", Name = fileName },
        };

        var session = await _client
            .PostAsync<GraphUploadSessionDto>(
                GraphPaths.CreateUploadSession(driveId, parentItemId, fileName),
                sessionRequest,
                cancellationToken)
            .ConfigureAwait(false);

        if (!session.Succeeded || string.IsNullOrEmpty(session.Value?.UploadUrl))
        {
            return OperationResult<string?>.Failure(
                session.Error ?? GraphErrorMapper.Map(session.StatusCode, null, null, "start a manifest upload"));
        }

        var uploadUri = new Uri(session.Value.UploadUrl, UriKind.Absolute);
        var total = bytes.LongLength;

        for (long offset = 0; offset < total; offset += UploadChunkSizeBytes)
        {
            cancellationToken.ThrowIfCancellationRequested();

            var length = (int)Math.Min(UploadChunkSizeBytes, total - offset);
            var chunk = new ReadOnlyMemory<byte>(bytes, (int)offset, length);

            var range = new ContentRangeHeaderValue(offset, offset + length - 1, total);

            var chunkResponse = await _client
                .PutUploadSessionChunkAsync(uploadUri, chunk, range, total, cancellationToken)
                .ConfigureAwait(false);

            if (!chunkResponse.Succeeded)
            {
                return OperationResult<string?>.Failure(
                    chunkResponse.Error
                    ?? GraphErrorMapper.Map(chunkResponse.StatusCode, null, null, "upload a manifest chunk"));
            }
        }

        _logger.LogInformation(
            "Uploaded a large manifest of {Bytes} bytes in {Chunks} chunk(s).",
            total,
            (total + UploadChunkSizeBytes - 1) / UploadChunkSizeBytes);

        return OperationResult<string?>.Success(null);
    }

    /// <summary>
    /// Decodes manifest bytes as UTF-8, tolerating a byte-order mark that another tool may have
    /// written even though this application never writes one.
    /// </summary>
    internal static string DecodeUtf8(byte[] bytes)
    {
        if (bytes.Length >= 3 && bytes[0] == 0xEF && bytes[1] == 0xBB && bytes[2] == 0xBF)
        {
            return Encoding.UTF8.GetString(bytes, 3, bytes.Length - 3);
        }

        return Encoding.UTF8.GetString(bytes);
    }

    /// <summary>Chooses the MIME type from the manifest file's extension.</summary>
    internal static string ContentTypeFor(string fileName) =>
        Path.GetExtension(fileName).ToLowerInvariant() switch
        {
            ".json" => "application/json; charset=utf-8",
            ".csv" => "text/csv; charset=utf-8",
            ".md" => "text/markdown; charset=utf-8",
            _ => "text/plain; charset=utf-8",
        };
}
