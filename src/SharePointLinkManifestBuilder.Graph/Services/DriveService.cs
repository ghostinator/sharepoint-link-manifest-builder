using System.Runtime.CompilerServices;
using Microsoft.Extensions.Logging;
using SharePointLinkManifestBuilder.Core.Abstractions;
using SharePointLinkManifestBuilder.Core.Models;
using SharePointLinkManifestBuilder.Core.Urls;
using SharePointLinkManifestBuilder.Graph.Dto;
using SharePointLinkManifestBuilder.Graph.Http;

namespace SharePointLinkManifestBuilder.Graph.Services;

/// <summary>Drive and drive-item access for SharePoint libraries and OneDrive.</summary>
public sealed class DriveService : IDriveService
{
    private readonly IGraphApiClient _client;
    private readonly ILogger<DriveService> _logger;

    /// <summary>Creates the service.</summary>
    public DriveService(IGraphApiClient client, ILogger<DriveService> logger)
    {
        _client = client ?? throw new ArgumentNullException(nameof(client));
        _logger = logger ?? throw new ArgumentNullException(nameof(logger));
    }

    /// <inheritdoc />
    public Task<OperationResult<DriveResource>> GetMyDriveAsync(CancellationToken cancellationToken = default) =>
        GetDriveFromPathAsync(GraphPaths.MyDrive(), "read your OneDrive", cancellationToken);

    /// <inheritdoc />
    public async Task<OperationResult<DriveResource>> GetUserDriveAsync(
        string userId,
        CancellationToken cancellationToken = default)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(userId);

        var result = await GetDriveFromPathAsync(
            GraphPaths.UserDrive(userId),
            "read that user's OneDrive",
            cancellationToken).ConfigureAwait(false);

        if (result.Succeeded)
        {
            return result;
        }

        _logger.LogInformation(
            "Could not open the requested user's OneDrive: {Kind}. This is an expected outcome when the "
            + "drive is unprovisioned or the signed-in user lacks access.",
            result.Error!.Kind);

        // A user's OneDrive being unavailable is a normal, expected outcome rather than a bug.
        // The distinction matters to the user, so it is reported specifically instead of being
        // collapsed into a generic access error. This application never provisions a drive.
        var error = result.Error!;

        return OperationResult<DriveResource>.Failure(error.Kind switch
        {
            GraphErrorKind.FileDeletedDuringProcessing or GraphErrorKind.LibraryNotFound => error with
            {
                Kind = GraphErrorKind.UserDriveUnprovisioned,
                Message = "That user does not have a OneDrive yet, or it has not been set up.",
                SuggestedAction =
                    "The user needs to open OneDrive once so it is created. This application does not "
                    + "provision OneDrive on someone's behalf.",
            },

            GraphErrorKind.SharePointAccessDenied => error with
            {
                Kind = GraphErrorKind.OneDriveAccessDenied,
                Message = "You do not have permission to read that user's OneDrive.",
                SuggestedAction =
                    "Delegated access is limited to what your own account can open. Administrator consent "
                    + "alone does not grant access to every user's OneDrive.",
            },

            _ => error,
        });
    }

    /// <inheritdoc />
    public Task<OperationResult<DriveResource>> GetDriveAsync(
        string driveId,
        CancellationToken cancellationToken = default) =>
        GetDriveFromPathAsync(GraphPaths.Drive(driveId), "read a drive", cancellationToken);

    /// <inheritdoc />
    public async Task<OperationResult<SharePointFolder>> GetRootFolderAsync(
        string driveId,
        CancellationToken cancellationToken = default)
    {
        var response = await _client
            .GetAsync<GraphDriveItemDto>(GraphPaths.DriveRoot(driveId), cancellationToken)
            .ConfigureAwait(false);

        if (!response.Succeeded || response.Value?.Id is null)
        {
            return OperationResult<SharePointFolder>.Failure(
                response.Error ?? GraphErrorMapper.Map(response.StatusCode, null, null, "read a drive root folder"));
        }

        return OperationResult<SharePointFolder>.Success(new SharePointFolder
        {
            DriveId = driveId,
            ItemId = response.Value.Id,
            Name = response.Value.Name ?? "root",
            RelativePath = string.Empty,
            WebUrl = response.Value.WebUrl,
            ChildCount = response.Value.Folder?.ChildCount,
            IsRoot = true,
        });
    }

    /// <inheritdoc />
    public async Task<OperationResult<SharePointFolder>> GetFolderByPathAsync(
        string driveId,
        string relativePath,
        CancellationToken cancellationToken = default)
    {
        var response = await _client
            .GetAsync<GraphDriveItemDto>(GraphPaths.ItemByPath(driveId, relativePath), cancellationToken)
            .ConfigureAwait(false);

        if (!response.Succeeded || response.Value?.Id is null)
        {
            return OperationResult<SharePointFolder>.Failure(
                response.Error ?? GraphErrorMapper.Map(response.StatusCode, null, null, "resolve a folder"));
        }

        if (response.Value.Folder is null)
        {
            return OperationResult<SharePointFolder>.Failure(new GraphError
            {
                Kind = GraphErrorKind.FolderNotFound,
                Message = $"'{response.Value.Name}' is a file, not a folder.",
                SuggestedAction = "Select a folder, a document library, or a site.",
            });
        }

        return OperationResult<SharePointFolder>.Success(new SharePointFolder
        {
            DriveId = driveId,
            ItemId = response.Value.Id,
            Name = response.Value.Name ?? "folder",
            RelativePath = NormalizeRelativePath(relativePath),
            WebUrl = response.Value.WebUrl,
            ChildCount = response.Value.Folder.ChildCount,
            IsRoot = string.IsNullOrEmpty(relativePath),
        });
    }

    /// <inheritdoc />
    public async Task<OperationResult<DiscoveredFile>> GetItemAsync(
        string driveId,
        string itemId,
        CancellationToken cancellationToken = default)
    {
        var response = await _client
            .GetAsync<GraphDriveItemDto>(GraphPaths.Item(driveId, itemId), cancellationToken)
            .ConfigureAwait(false);

        if (!response.Succeeded || response.Value?.Id is null)
        {
            return OperationResult<DiscoveredFile>.Failure(
                response.Error ?? GraphErrorMapper.Map(response.StatusCode, null, null, "read a drive item"));
        }

        return OperationResult<DiscoveredFile>.Success(MapItem(response.Value, driveId, string.Empty));
    }

    /// <inheritdoc />
    public async IAsyncEnumerable<DiscoveredFile> GetChildrenAsync(
        string driveId,
        string folderItemId,
        [EnumeratorCancellation] CancellationToken cancellationToken = default)
    {
        await foreach (var dto in _client
            .EnumeratePagedAsync<GraphDriveItemDto>(
                GraphPaths.Children(driveId, folderItemId), cancellationToken)
            .ConfigureAwait(false))
        {
            if (dto.Id is null)
            {
                continue;
            }

            yield return MapItem(dto, driveId, string.Empty);
        }
    }

    /// <inheritdoc />
    public async Task<OperationResult<IReadOnlyList<SharePointFolder>>> GetSubfoldersAsync(
        string driveId,
        string folderItemId,
        CancellationToken cancellationToken = default)
    {
        try
        {
            var folders = new List<SharePointFolder>();

            await foreach (var dto in _client
                .EnumeratePagedAsync<GraphDriveItemDto>(
                    GraphPaths.Children(driveId, folderItemId), cancellationToken)
                .ConfigureAwait(false))
            {
                if (dto.Id is null || dto.Folder is null)
                {
                    continue;
                }

                folders.Add(new SharePointFolder
                {
                    DriveId = driveId,
                    ItemId = dto.Id,
                    Name = dto.Name ?? "folder",
                    RelativePath = BuildRelativePath(dto),
                    WebUrl = dto.WebUrl,
                    ChildCount = dto.Folder.ChildCount,
                });
            }

            return OperationResult<IReadOnlyList<SharePointFolder>>.Success(folders);
        }
        catch (GraphOperationException ex)
        {
            return OperationResult<IReadOnlyList<SharePointFolder>>.Failure(ex.Error);
        }
    }

    /// <inheritdoc />
    public async Task<OperationResult<DiscoveredFile>> ResolveSharingUrlAsync(
        string sharingUrl,
        CancellationToken cancellationToken = default)
    {
        var parsed = SharePointUrlParser.Parse(sharingUrl);

        if (!parsed.IsValid)
        {
            return OperationResult<DiscoveredFile>.Failure(new GraphError
            {
                Kind = GraphErrorKind.InvalidUrl,
                Message = parsed.FailureReason ?? "That URL could not be understood.",
            });
        }

        var token = GraphShareTokenEncoder.Encode(sharingUrl);
        var response = await _client
            .GetAsync<GraphDriveItemDto>(GraphPaths.SharedItem(token), cancellationToken)
            .ConfigureAwait(false);

        if (!response.Succeeded || response.Value?.Id is null)
        {
            return OperationResult<DiscoveredFile>.Failure(
                response.Error ?? GraphErrorMapper.Map(response.StatusCode, null, null, "resolve a shared item"));
        }

        var driveId = response.Value.ParentReference?.DriveId;

        if (string.IsNullOrEmpty(driveId))
        {
            return OperationResult<DiscoveredFile>.Failure(new GraphError
            {
                Kind = GraphErrorKind.InvalidUrl,
                Message = "That shared item could not be traced back to a drive.",
            });
        }

        return OperationResult<DiscoveredFile>.Success(MapItem(response.Value, driveId, string.Empty));
    }

    private async Task<OperationResult<DriveResource>> GetDriveFromPathAsync(
        string path,
        string operation,
        CancellationToken cancellationToken)
    {
        var response = await _client.GetAsync<GraphDriveDto>(path, cancellationToken).ConfigureAwait(false);

        if (!response.Succeeded || response.Value?.Id is null)
        {
            return OperationResult<DriveResource>.Failure(
                response.Error ?? GraphErrorMapper.Map(response.StatusCode, null, null, operation));
        }

        var dto = response.Value;

        return OperationResult<DriveResource>.Success(new DriveResource
        {
            DriveId = dto.Id,
            Name = dto.Name ?? "OneDrive",
            DriveType = dto.DriveType,
            WebUrl = dto.WebUrl,
            OwnerDisplayName = dto.Owner?.User?.DisplayName,
            OwnerUserId = dto.Owner?.User?.Id,
        });
    }

    /// <summary>
    /// Maps a Graph drive item, classifying its kind. Classification order matters: a OneNote
    /// notebook carries both a folder facet and a package facet, and treating it as a folder
    /// would make a job descend into something that is not really a folder.
    /// </summary>
    /// <param name="dto">The Graph payload.</param>
    /// <param name="driveId">The owning drive.</param>
    /// <param name="relativePathPrefix">Path of the enumeration root, for relative paths.</param>
    internal static DiscoveredFile MapItem(GraphDriveItemDto dto, string driveId, string relativePathPrefix)
    {
        var kind = dto switch
        {
            { Package: not null } => DriveItemKind.Package,
            { RemoteItem: not null } => DriveItemKind.RemoteItem,
            { Folder: not null } => DriveItemKind.Folder,
            { File: not null } => DriveItemKind.File,
            _ => DriveItemKind.Unsupported,
        };

        var name = dto.Name ?? "(unnamed)";
        var parentPath = BuildRelativePath(dto);

        var relativePath = string.IsNullOrEmpty(relativePathPrefix)
            ? CombinePath(parentPath, name)
            : CombinePath(TrimPrefix(parentPath, relativePathPrefix), name);

        return new DiscoveredFile
        {
            DriveId = driveId,
            ItemId = dto.Id!,
            Name = name,
            RelativePath = relativePath,
            WebUrl = dto.WebUrl,
            Size = dto.Size,
            LastModifiedUtc = dto.LastModifiedDateTime,
            ETag = dto.ETag,
            ParentFolderItemId = dto.ParentReference?.Id,
            ParentRelativePath = string.IsNullOrEmpty(relativePathPrefix)
                ? parentPath
                : TrimPrefix(parentPath, relativePathPrefix),
            Kind = kind,
        };
    }

    /// <summary>
    /// Extracts the drive-relative folder path from a Graph parent reference, whose
    /// <c>path</c> takes the form <c>/drive/root:/Folder/Sub</c> or <c>/drives/{id}/root:/…</c>.
    /// </summary>
    internal static string BuildRelativePath(GraphDriveItemDto dto)
    {
        var path = dto.ParentReference?.Path;

        if (string.IsNullOrEmpty(path))
        {
            return string.Empty;
        }

        var marker = path.IndexOf("root:", StringComparison.OrdinalIgnoreCase);
        var tail = marker >= 0 ? path[(marker + "root:".Length)..] : path;

        return Uri.UnescapeDataString(tail).Trim('/');
    }

    private static string CombinePath(string folder, string name) =>
        string.IsNullOrEmpty(folder) ? name : $"{folder}/{name}";

    private static string TrimPrefix(string path, string prefix)
    {
        prefix = prefix.Trim('/');

        if (prefix.Length == 0)
        {
            return path;
        }

        if (string.Equals(path, prefix, StringComparison.OrdinalIgnoreCase))
        {
            return string.Empty;
        }

        return path.StartsWith(prefix + "/", StringComparison.OrdinalIgnoreCase)
            ? path[(prefix.Length + 1)..]
            : path;
    }

    private static string NormalizeRelativePath(string? path) =>
        string.IsNullOrWhiteSpace(path) ? string.Empty : path.Replace('\\', '/').Trim('/');
}
