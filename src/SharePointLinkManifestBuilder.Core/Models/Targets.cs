namespace SharePointLinkManifestBuilder.Core.Models;

/// <summary>Where a processing target's files come from.</summary>
public enum TargetSourceType
{
    /// <summary>An entire SharePoint site (every accessible document library within it).</summary>
    SharePointSite = 0,

    /// <summary>A single SharePoint document library.</summary>
    DocumentLibrary = 1,

    /// <summary>A folder inside a SharePoint document library.</summary>
    SharePointFolder = 2,

    /// <summary>The signed-in user's own OneDrive, or a folder within it.</summary>
    MyOneDrive = 3,

    /// <summary>Another user's OneDrive, or a folder within it, where access is permitted.</summary>
    UserOneDrive = 4,
}

/// <summary>
/// One location a job will process. A job may mix any number of these across SharePoint and
/// OneDrive. Immutable: editing produces a new instance.
/// </summary>
public sealed record ProcessingTarget
{
    /// <summary>Stable identifier for this target within a job, used by the UI and reports.</summary>
    public string TargetId { get; init; } = Guid.NewGuid().ToString("n");

    /// <summary>Which source family this target belongs to.</summary>
    public required TargetSourceType SourceType { get; init; }

    /// <summary>Tenant this target belongs to. Guards against cross-tenant mixing.</summary>
    public required string TenantId { get; init; }

    /// <summary>Site display name, for SharePoint targets.</summary>
    public string? SiteName { get; init; }

    /// <summary>Site URL, for SharePoint targets.</summary>
    public string? SiteUrl { get; init; }

    /// <summary>Graph site ID, for SharePoint targets.</summary>
    public string? SiteId { get; init; }

    /// <summary>Owning user's display name, for a User OneDrive target.</summary>
    public string? UserDisplayName { get; init; }

    /// <summary>Owning user's ID, for a User OneDrive target.</summary>
    public string? UserId { get; init; }

    /// <summary>Friendly library or drive name.</summary>
    public string? DriveName { get; init; }

    /// <summary>
    /// Graph drive ID. Null only for a whole-site target, whose libraries are resolved during
    /// preflight and expanded into one concrete target per library.
    /// </summary>
    public string? DriveId { get; init; }

    /// <summary>Name of the folder processing starts from. Null means the drive root.</summary>
    public string? StartingFolderName { get; init; }

    /// <summary>Graph item ID of the starting folder. Null means the drive root.</summary>
    public string? StartingFolderItemId { get; init; }

    /// <summary>Starting folder path relative to the drive root, forward-slashed. Empty for root.</summary>
    public string StartingFolderRelativePath { get; init; } = string.Empty;

    /// <summary>Absolute URL of the starting location, for "Open in browser".</summary>
    public string? WebUrl { get; init; }

    /// <summary>True to descend into subfolders. Configurable per target.</summary>
    public bool Recursive { get; init; }

    /// <summary>True to include files directly inside the starting folder. Default true.</summary>
    public bool IncludeDirectFiles { get; init; } = true;

    /// <summary>False to keep the target in the job but skip it on this run.</summary>
    public bool IsEnabled { get; init; } = true;

    /// <summary>A friendly one-line description used throughout the UI in preference to any ID.</summary>
    public string DisplayPath => SourceType switch
    {
        TargetSourceType.SharePointSite => $"{SiteName ?? "SharePoint site"} (all libraries)",
        TargetSourceType.DocumentLibrary => $"{SiteName} / {DriveName}",
        TargetSourceType.SharePointFolder =>
            $"{SiteName} / {DriveName}/{StartingFolderRelativePath}".TrimEnd('/'),
        TargetSourceType.MyOneDrive =>
            string.IsNullOrEmpty(StartingFolderRelativePath)
                ? "My OneDrive"
                : $"My OneDrive/{StartingFolderRelativePath}",
        TargetSourceType.UserOneDrive =>
            string.IsNullOrEmpty(StartingFolderRelativePath)
                ? $"{UserDisplayName}'s OneDrive"
                : $"{UserDisplayName}'s OneDrive/{StartingFolderRelativePath}",
        _ => "Unknown target",
    };

    /// <summary>A human-readable source label used in manifest headers.</summary>
    public string SourceTypeLabel => SourceType switch
    {
        TargetSourceType.SharePointSite => "SharePoint Site",
        TargetSourceType.DocumentLibrary => "Document Library",
        TargetSourceType.SharePointFolder => "Document Library",
        TargetSourceType.MyOneDrive => "My OneDrive",
        TargetSourceType.UserOneDrive => "User OneDrive",
        _ => "Unknown",
    };

    /// <summary>
    /// True when the target names a concrete drive and can be enumerated directly. A whole-site
    /// target is not resolved until preflight expands it into one target per library.
    /// </summary>
    public bool IsResolved => !string.IsNullOrEmpty(DriveId);
}

/// <summary>How an overlapping pair of targets should be reconciled.</summary>
public enum OverlapResolution
{
    /// <summary>Keep the broader target and drop the narrower one. The default.</summary>
    KeepParent = 0,

    /// <summary>Keep the narrower target and drop the broader one.</summary>
    KeepChild = 1,

    /// <summary>Keep both targets, relying on file-level deduplication to process each file once.</summary>
    KeepBothDeduplicate = 2,
}

/// <summary>A detected overlap between two selected targets, surfaced before a job runs.</summary>
public sealed record TargetOverlap
{
    /// <summary>The broader target.</summary>
    public required ProcessingTarget Parent { get; init; }

    /// <summary>The target contained within <see cref="Parent"/>.</summary>
    public required ProcessingTarget Child { get; init; }

    /// <summary>Why these two overlap, in plain language.</summary>
    public required string Explanation { get; init; }

    /// <summary>True when the two targets are the same location selected twice.</summary>
    public bool IsExactDuplicate { get; init; }
}
