using System.Runtime.CompilerServices;
using Microsoft.Extensions.Logging;
using SharePointLinkManifestBuilder.Core.Abstractions;
using SharePointLinkManifestBuilder.Core.Filtering;
using SharePointLinkManifestBuilder.Core.Models;
using SharePointLinkManifestBuilder.Core.Targets;

namespace SharePointLinkManifestBuilder.Core.Jobs;

/// <summary>
/// Walks a processing target and yields the files it contains.
/// <para>
/// The walk is streamed: files are produced as they are found rather than collected into a
/// list, so a library with hundreds of thousands of items does not have to fit in memory, and
/// the preview can start showing results immediately.
/// </para>
/// <para>
/// Skipped items are yielded too, carrying their reason. The preview needs to show a user
/// <em>why</em> something will not be processed; silently dropping items would make a job that
/// found nothing indistinguishable from a job that filtered everything out.
/// </para>
/// </summary>
public sealed class FileDiscoveryService : IFileDiscoveryService
{
    private readonly IDriveService _driveService;
    private readonly ILogger<FileDiscoveryService> _logger;

    /// <summary>Creates the service.</summary>
    public FileDiscoveryService(IDriveService driveService, ILogger<FileDiscoveryService> logger)
    {
        _driveService = driveService ?? throw new ArgumentNullException(nameof(driveService));
        _logger = logger ?? throw new ArgumentNullException(nameof(logger));
    }

    /// <inheritdoc />
    public async IAsyncEnumerable<DiscoveredFile> DiscoverAsync(
        ProcessingTarget target,
        FilterConfiguration filters,
        ManifestConfiguration manifestConfiguration,
        [EnumeratorCancellation] CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(target);
        ArgumentNullException.ThrowIfNull(filters);
        ArgumentNullException.ThrowIfNull(manifestConfiguration);

        if (string.IsNullOrEmpty(target.DriveId))
        {
            throw new InvalidOperationException(
                $"Target '{target.DisplayPath}' has not been resolved to a drive. Whole-site targets must be "
                + "expanded into one target per library during preflight.");
        }

        var evaluator = new FileFilterEvaluator(filters, manifestConfiguration);
        var startingFolderId = await ResolveStartingFolderAsync(target, cancellationToken).ConfigureAwait(false);

        if (startingFolderId is null)
        {
            yield break;
        }

        // An explicit stack rather than recursion: a deeply nested library must not be able to
        // overflow the stack, and an iterative walk is trivially cancellable between folders.
        var pending = new Stack<(string ItemId, string RelativePath)>();
        pending.Push((startingFolderId, string.Empty));

        var visitedFolders = new HashSet<string>(StringComparer.Ordinal);
        var isFirstFolder = true;

        while (pending.Count > 0)
        {
            cancellationToken.ThrowIfCancellationRequested();

            var (folderId, folderRelativePath) = pending.Pop();

            // Guards against a cycle introduced by a shortcut or a Graph anomaly. Without it a
            // malformed hierarchy could loop forever, hammering the tenant.
            if (!visitedFolders.Add(folderId))
            {
                _logger.LogWarning(
                    "Skipping a folder that has already been visited in this walk, which suggests a cycle.");
                continue;
            }

            var includeFilesHere = !isFirstFolder || target.IncludeDirectFiles;
            isFirstFolder = false;

            IReadOnlyList<DiscoveredFile> children = [];
            var folderUnreadable = false;

            try
            {
                children = await ReadChildrenAsync(target.DriveId, folderId, cancellationToken)
                    .ConfigureAwait(false);
            }
            catch (Exception ex) when (ex is not OperationCanceledException)
            {
                // One inaccessible subtree must not abort the whole job. C# forbids yielding
                // from a catch block, so the failure is recorded and reported just below.
                _logger.LogWarning(
                    ex,
                    "Could not list the contents of a folder; it will be reported as skipped and the walk continues.");

                folderUnreadable = true;
            }

            if (folderUnreadable)
            {
                // Reported as a skipped pseudo-item so an inaccessible subtree is visible in the
                // preview and the job report rather than silently absent.
                yield return new DiscoveredFile
                {
                    DriveId = target.DriveId,
                    ItemId = folderId,
                    Name = folderRelativePath.Length == 0 ? "(starting folder)" : folderRelativePath,
                    RelativePath = folderRelativePath,
                    ParentRelativePath = folderRelativePath,
                    Kind = DriveItemKind.Folder,
                    SkipReason = SkipReason.UnsupportedItemType,
                    TargetId = target.TargetId,
                };

                continue;
            }

            foreach (var child in children)
            {
                cancellationToken.ThrowIfCancellationRequested();

                var childRelativePath = Combine(folderRelativePath, child.Name);

                if (child.Kind == DriveItemKind.Folder)
                {
                    if (target.Recursive)
                    {
                        pending.Push((child.ItemId, childRelativePath));
                    }

                    continue;
                }

                if (!includeFilesHere)
                {
                    continue;
                }

                var normalized = child with
                {
                    RelativePath = childRelativePath,
                    ParentRelativePath = folderRelativePath,
                    ParentFolderItemId = folderId,
                    TargetId = target.TargetId,
                };

                var skipReason = evaluator.Evaluate(normalized);

                yield return skipReason == SkipReason.None
                    ? normalized
                    : normalized with { SkipReason = skipReason };
            }
        }
    }

    /// <summary>
    /// Materializes one folder's children. Buffering a single page-set is intentional: it keeps
    /// the iterator's try/catch around the network call, which C# does not permit around a
    /// <c>yield return</c>, while still bounding memory to one folder at a time.
    /// </summary>
    private async Task<IReadOnlyList<DiscoveredFile>> ReadChildrenAsync(
        string driveId,
        string folderItemId,
        CancellationToken cancellationToken)
    {
        var children = new List<DiscoveredFile>();

        await foreach (var child in _driveService
            .GetChildrenAsync(driveId, folderItemId, cancellationToken)
            .ConfigureAwait(false))
        {
            children.Add(child);
        }

        return children;
    }

    private async Task<string?> ResolveStartingFolderAsync(
        ProcessingTarget target,
        CancellationToken cancellationToken)
    {
        if (!string.IsNullOrEmpty(target.StartingFolderItemId))
        {
            return target.StartingFolderItemId;
        }

        var path = TargetPlanner.NormalizePath(target.StartingFolderRelativePath);

        var result = path.Length == 0
            ? await _driveService.GetRootFolderAsync(target.DriveId!, cancellationToken).ConfigureAwait(false)
            : await _driveService.GetFolderByPathAsync(target.DriveId!, path, cancellationToken)
                .ConfigureAwait(false);

        if (result.Succeeded)
        {
            return result.Value!.ItemId;
        }

        _logger.LogWarning(
            "Could not resolve the starting folder for target {Target}: {Reason}",
            target.DisplayPath,
            result.Error!.Message);

        return null;
    }

    private static string Combine(string folder, string name) =>
        string.IsNullOrEmpty(folder) ? name : $"{folder}/{name}";
}
