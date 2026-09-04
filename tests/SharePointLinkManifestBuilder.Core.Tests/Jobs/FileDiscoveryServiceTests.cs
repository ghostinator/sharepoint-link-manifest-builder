using System.Runtime.CompilerServices;
using Microsoft.Extensions.Logging.Abstractions;
using SharePointLinkManifestBuilder.Core.Abstractions;
using SharePointLinkManifestBuilder.Core.Jobs;
using SharePointLinkManifestBuilder.Core.Models;

namespace SharePointLinkManifestBuilder.Core.Tests.Jobs;

/// <summary>
/// An in-memory drive standing in for Microsoft Graph, so recursion, filtering and relative
/// paths can be verified without a network.
/// </summary>
internal sealed class FakeDriveService : IDriveService
{
    private readonly Dictionary<string, List<DiscoveredFile>> _childrenByFolder = new(StringComparer.Ordinal);
    private readonly HashSet<string> _unreadableFolders = new(StringComparer.Ordinal);

    /// <summary>Folders whose children were actually requested, in order.</summary>
    public List<string> RequestedFolders { get; } = [];

    /// <summary>Adds a folder and its direct children.</summary>
    public FakeDriveService WithFolder(string folderId, params DiscoveredFile[] children)
    {
        _childrenByFolder[folderId] = [.. children];
        return this;
    }

    /// <summary>Marks a folder as one that throws when enumerated.</summary>
    public FakeDriveService WithUnreadableFolder(string folderId)
    {
        _unreadableFolders.Add(folderId);
        return this;
    }

    /// <summary>Builds a file child.</summary>
    public static DiscoveredFile File(string name, string parentFolderId, string parentPath = "") => new()
    {
        DriveId = "drive-1",
        ItemId = "item-" + name,
        Name = name,
        RelativePath = name,
        ParentFolderItemId = parentFolderId,
        ParentRelativePath = parentPath,
        Kind = DriveItemKind.File,
        Size = 100,
        LastModifiedUtc = new DateTimeOffset(2026, 1, 1, 0, 0, 0, TimeSpan.Zero),
    };

    /// <summary>Builds a folder child.</summary>
    public static DiscoveredFile Folder(string name, string itemId, string parentFolderId) => new()
    {
        DriveId = "drive-1",
        ItemId = itemId,
        Name = name,
        RelativePath = name,
        ParentFolderItemId = parentFolderId,
        Kind = DriveItemKind.Folder,
    };

    public async IAsyncEnumerable<DiscoveredFile> GetChildrenAsync(
        string driveId,
        string folderItemId,
        [EnumeratorCancellation] CancellationToken cancellationToken = default)
    {
        RequestedFolders.Add(folderItemId);

        if (_unreadableFolders.Contains(folderItemId))
        {
            throw new InvalidOperationException("Access denied to this folder.");
        }

        await Task.Yield();

        foreach (var child in _childrenByFolder.GetValueOrDefault(folderItemId, []))
        {
            cancellationToken.ThrowIfCancellationRequested();
            yield return child;
        }
    }

    public Task<OperationResult<SharePointFolder>> GetRootFolderAsync(
        string driveId, CancellationToken cancellationToken = default) =>
        Task.FromResult(OperationResult<SharePointFolder>.Success(new SharePointFolder
        {
            DriveId = driveId,
            ItemId = "root",
            Name = "root",
            IsRoot = true,
        }));

    public Task<OperationResult<SharePointFolder>> GetFolderByPathAsync(
        string driveId, string relativePath, CancellationToken cancellationToken = default) =>
        Task.FromResult(OperationResult<SharePointFolder>.Success(new SharePointFolder
        {
            DriveId = driveId,
            ItemId = "folder-" + relativePath.Replace('/', '-'),
            Name = relativePath,
            RelativePath = relativePath,
        }));

    public Task<OperationResult<DriveResource>> GetMyDriveAsync(CancellationToken cancellationToken = default) =>
        throw new NotSupportedException();

    public Task<OperationResult<DriveResource>> GetUserDriveAsync(
        string userId, CancellationToken cancellationToken = default) =>
        throw new NotSupportedException();

    public Task<OperationResult<DriveResource>> GetDriveAsync(
        string driveId, CancellationToken cancellationToken = default) =>
        throw new NotSupportedException();

    public Task<OperationResult<DiscoveredFile>> GetItemAsync(
        string driveId, string itemId, CancellationToken cancellationToken = default) =>
        throw new NotSupportedException();

    public Task<OperationResult<IReadOnlyList<SharePointFolder>>> GetSubfoldersAsync(
        string driveId, string folderItemId, CancellationToken cancellationToken = default) =>
        throw new NotSupportedException();

    public Task<OperationResult<DiscoveredFile>> ResolveSharingUrlAsync(
        string sharingUrl, CancellationToken cancellationToken = default) =>
        throw new NotSupportedException();
}

/// <summary>Recursion, filtering, relative paths and resilience during the folder walk.</summary>
public sealed class FileDiscoveryServiceTests
{
    private static FileDiscoveryService Create(FakeDriveService drive) =>
        new(drive, NullLogger<FileDiscoveryService>.Instance);

    private static ProcessingTarget Target(bool recursive, bool includeDirectFiles = true) => new()
    {
        TargetId = "t1",
        SourceType = TargetSourceType.DocumentLibrary,
        TenantId = TestData.TenantId,
        DriveId = "drive-1",
        DriveName = "Documents",
        StartingFolderItemId = "root",
        Recursive = recursive,
        IncludeDirectFiles = includeDirectFiles,
    };

    private static async Task<List<DiscoveredFile>> CollectAsync(
        FileDiscoveryService service,
        ProcessingTarget target,
        FilterConfiguration? filters = null)
    {
        var results = new List<DiscoveredFile>();

        await foreach (var file in service.DiscoverAsync(
            target, filters ?? new FilterConfiguration(), new ManifestConfiguration()))
        {
            results.Add(file);
        }

        return results;
    }

    /// <summary>A non-recursive target processes only the folder's direct children.</summary>
    [Fact]
    public async Task NonRecursive_ProcessesOnlyDirectChildren()
    {
        var drive = new FakeDriveService()
            .WithFolder("root",
                FakeDriveService.File("a.docx", "root"),
                FakeDriveService.File("b.docx", "root"),
                FakeDriveService.Folder("Reports", "f-reports", "root"))
            .WithFolder("f-reports", FakeDriveService.File("deep.docx", "f-reports"));

        var results = await CollectAsync(Create(drive), Target(recursive: false));

        Assert.Equal(2, results.Count);
        Assert.Equal(["a.docx", "b.docx"], results.Select(r => r.Name));

        // The subfolder is never even enumerated.
        Assert.Equal(["root"], drive.RequestedFolders);
    }

    [Fact]
    public async Task Recursive_DescendsIntoEverySubfolder()
    {
        var drive = new FakeDriveService()
            .WithFolder("root",
                FakeDriveService.File("a.docx", "root"),
                FakeDriveService.Folder("Reports", "f-reports", "root"))
            .WithFolder("f-reports",
                FakeDriveService.File("b.docx", "f-reports"),
                FakeDriveService.Folder("Q1", "f-q1", "f-reports"))
            .WithFolder("f-q1", FakeDriveService.File("c.docx", "f-q1"));

        var results = await CollectAsync(Create(drive), Target(recursive: true));

        Assert.Equal(3, results.Count);
        Assert.Equal(["a.docx", "b.docx", "c.docx"], results.Select(r => r.Name).Order());
        Assert.Equal(3, drive.RequestedFolders.Count);
    }

    /// <summary>Relative paths accumulate as the walk descends.</summary>
    [Fact]
    public async Task Recursive_BuildsPathsRelativeToTheStartingFolder()
    {
        var drive = new FakeDriveService()
            .WithFolder("root", FakeDriveService.Folder("Reports", "f-reports", "root"))
            .WithFolder("f-reports", FakeDriveService.Folder("Q1", "f-q1", "f-reports"))
            .WithFolder("f-q1", FakeDriveService.File("summary.docx", "f-q1"));

        var results = await CollectAsync(Create(drive), Target(recursive: true));

        var file = Assert.Single(results);
        Assert.Equal("Reports/Q1/summary.docx", file.RelativePath);
        Assert.Equal("Reports/Q1", file.ParentRelativePath);
        Assert.Equal("f-q1", file.ParentFolderItemId);
    }

    /// <summary>Turning off direct files keeps the starting folder's own files out.</summary>
    [Fact]
    public async Task IncludeDirectFilesOff_SkipsTheStartingFoldersOwnFiles()
    {
        var drive = new FakeDriveService()
            .WithFolder("root",
                FakeDriveService.File("top.docx", "root"),
                FakeDriveService.Folder("Reports", "f-reports", "root"))
            .WithFolder("f-reports", FakeDriveService.File("nested.docx", "f-reports"));

        var results = await CollectAsync(
            Create(drive), Target(recursive: true, includeDirectFiles: false));

        Assert.Equal("nested.docx", Assert.Single(results).Name);
    }

    /// <summary>Skipped items are yielded with a reason so the preview can explain them.</summary>
    [Fact]
    public async Task SkippedItems_AreYieldedWithTheirReason()
    {
        var drive = new FakeDriveService().WithFolder("root",
            FakeDriveService.File("good.docx", "root"),
            FakeDriveService.File("~$locked.docx", "root"),
            FakeDriveService.File("_sharepoint-links.txt", "root"));

        var results = await CollectAsync(Create(drive), Target(recursive: false));

        Assert.Equal(3, results.Count);
        Assert.Equal(SkipReason.None, results.Single(r => r.Name == "good.docx").SkipReason);
        Assert.Equal(SkipReason.TemporaryFile, results.Single(r => r.Name == "~$locked.docx").SkipReason);

        // A job must never ingest the manifests it is about to write.
        Assert.Equal(
            SkipReason.GeneratedManifest,
            results.Single(r => r.Name == "_sharepoint-links.txt").SkipReason);
    }

    [Fact]
    public async Task Filters_AreAppliedDuringTheWalk()
    {
        var drive = new FakeDriveService().WithFolder("root",
            FakeDriveService.File("keep.docx", "root"),
            FakeDriveService.File("drop.pdf", "root"));

        var results = await CollectAsync(
            Create(drive),
            Target(recursive: false),
            new FilterConfiguration { IncludeExtensions = ["docx"] });

        Assert.Equal(SkipReason.None, results.Single(r => r.Name == "keep.docx").SkipReason);
        Assert.Equal(SkipReason.FilteredOut, results.Single(r => r.Name == "drop.pdf").SkipReason);
    }

    /// <summary>
    /// One inaccessible subtree must not abort the job. It is reported as a skipped item so it
    /// is visible in the preview rather than silently absent.
    /// </summary>
    [Fact]
    public async Task UnreadableSubfolder_IsReportedAndTheWalkContinues()
    {
        var drive = new FakeDriveService()
            .WithFolder("root",
                FakeDriveService.File("a.docx", "root"),
                FakeDriveService.Folder("Locked", "f-locked", "root"),
                FakeDriveService.Folder("Open", "f-open", "root"))
            .WithUnreadableFolder("f-locked")
            .WithFolder("f-open", FakeDriveService.File("b.docx", "f-open"));

        var results = await CollectAsync(Create(drive), Target(recursive: true));

        Assert.Contains(results, r => r.Name == "a.docx" && r.SkipReason == SkipReason.None);
        Assert.Contains(results, r => r.Name == "b.docx" && r.SkipReason == SkipReason.None);
        Assert.Contains(results, r => r.SkipReason == SkipReason.UnsupportedItemType);
    }

    /// <summary>A cycle introduced by a shortcut must not loop forever.</summary>
    [Fact]
    public async Task CyclicHierarchy_TerminatesInsteadOfLooping()
    {
        var drive = new FakeDriveService()
            .WithFolder("root", FakeDriveService.Folder("Loop", "f-loop", "root"))
            .WithFolder("f-loop",
                FakeDriveService.File("a.docx", "f-loop"),

                // Points back at the root, which without a guard would recurse forever.
                FakeDriveService.Folder("Back", "root", "f-loop"));

        var results = await CollectAsync(Create(drive), Target(recursive: true));

        Assert.Contains(results, r => r.Name == "a.docx");
        Assert.Equal(2, drive.RequestedFolders.Distinct(StringComparer.Ordinal).Count());
    }

    [Fact]
    public async Task EmptyFolder_YieldsNothing()
    {
        var drive = new FakeDriveService().WithFolder("root");

        Assert.Empty(await CollectAsync(Create(drive), Target(recursive: true)));
    }

    /// <summary>Cancellation must stop the walk promptly.</summary>
    [Fact]
    public async Task Cancellation_StopsTheWalk()
    {
        var drive = new FakeDriveService()
            .WithFolder("root",
                FakeDriveService.Folder("A", "f-a", "root"),
                FakeDriveService.Folder("B", "f-b", "root"))
            .WithFolder("f-a", FakeDriveService.File("a.docx", "f-a"))
            .WithFolder("f-b", FakeDriveService.File("b.docx", "f-b"));

        using var cancellation = new CancellationTokenSource();
        var service = Create(drive);
        var collected = new List<DiscoveredFile>();

        await Assert.ThrowsAnyAsync<OperationCanceledException>(async () =>
        {
            await foreach (var file in service.DiscoverAsync(
                Target(recursive: true), new FilterConfiguration(), new ManifestConfiguration(),
                cancellation.Token))
            {
                collected.Add(file);
                await cancellation.CancelAsync();
            }
        });

        Assert.Single(collected);
    }

    /// <summary>A whole-site target must be expanded before discovery, not silently mishandled.</summary>
    [Fact]
    public async Task UnresolvedTarget_ThrowsRatherThanSilentlyProducingNothing()
    {
        var target = Target(recursive: false) with { DriveId = null };
        var service = Create(new FakeDriveService());

        await Assert.ThrowsAsync<InvalidOperationException>(async () =>
        {
            await foreach (var _ in service.DiscoverAsync(
                target, new FilterConfiguration(), new ManifestConfiguration()))
            {
                // The exception is expected before the first item.
            }
        });
    }

    /// <summary>Every discovered file is stamped with the target it came from.</summary>
    [Fact]
    public async Task DiscoveredFiles_CarryTheirTargetId()
    {
        var drive = new FakeDriveService().WithFolder("root", FakeDriveService.File("a.docx", "root"));

        var results = await CollectAsync(Create(drive), Target(recursive: false));

        Assert.Equal("t1", Assert.Single(results).TargetId);
    }
}
