namespace SharePointLinkManifestBuilder.Core.Models;

/// <summary>
/// One file's record inside a manifest. Identity is (<see cref="DriveId"/>,
/// <see cref="ItemId"/>); the name and path are display metadata that may change between runs.
/// </summary>
public sealed record ManifestEntry
{
    /// <summary>File name including extension.</summary>
    public required string FileName { get; init; }

    /// <summary>Path relative to the manifest's scope, forward-slashed.</summary>
    public string RelativePath { get; init; } = string.Empty;

    /// <summary>The ordinary SharePoint or OneDrive URL of the item.</summary>
    public string? WebUrl { get; init; }

    /// <summary>The sharing URL produced for the item.</summary>
    public string? SharingUrl { get; init; }

    /// <summary>Drive containing the item. Half of the entry's identity.</summary>
    public required string DriveId { get; init; }

    /// <summary>Graph item ID. Half of the entry's identity.</summary>
    public required string ItemId { get; init; }

    /// <summary>Created, Reused or Existing.</summary>
    public required string Status { get; init; }

    /// <summary>When this entry was produced.</summary>
    public DateTimeOffset GeneratedUtc { get; init; } = DateTimeOffset.UtcNow;

    /// <summary>Graph error code, for a failed entry retained in a CSV or JSON report.</summary>
    public string? ErrorCode { get; init; }

    /// <summary>Human-readable failure message, for a failed entry.</summary>
    public string? ErrorMessage { get; init; }

    /// <summary>
    /// Set when a previous run recorded this file but the latest run did not encounter it, and
    /// the missing-entry policy is to mark rather than remove.
    /// </summary>
    public bool IsMissing { get; init; }

    /// <summary>The identity key used to match entries across runs.</summary>
    public string IdentityKey => $"{DriveId}|{ItemId}";

    /// <summary>Builds an entry from a successful link result.</summary>
    public static ManifestEntry FromResult(LinkResult result) => new()
    {
        FileName = result.File.Name,
        RelativePath = result.File.RelativePath,
        WebUrl = result.File.WebUrl,
        SharingUrl = result.SharingUrl,
        DriveId = result.File.DriveId,
        ItemId = result.File.ItemId,
        Status = result.ManifestStatus,
        GeneratedUtc = result.TimestampUtc,
        ErrorCode = result.Error?.GraphErrorCode,
        ErrorMessage = result.Error?.Message,
    };
}

/// <summary>The provenance block at the top of a manifest.</summary>
public sealed record ManifestHeader
{
    /// <summary>Manifest schema version, so a future reader can adapt.</summary>
    public string SchemaVersion { get; init; } = ManifestDefaults.SchemaVersion;

    /// <summary>Version of the application that produced the manifest.</summary>
    public required string ApplicationVersion { get; init; }

    /// <summary>The job that produced the manifest.</summary>
    public required string JobId { get; init; }

    /// <summary>When the manifest was generated.</summary>
    public DateTimeOffset GeneratedUtc { get; init; } = DateTimeOffset.UtcNow;

    /// <summary>Tenant display name, or the tenant ID when the name is unavailable.</summary>
    public required string TenantDisplayName { get; init; }

    /// <summary>Tenant ID.</summary>
    public required string TenantId { get; init; }

    /// <summary>SharePoint Site, Document Library, My OneDrive, or User OneDrive.</summary>
    public required string SourceType { get; init; }

    /// <summary>Site title, for SharePoint sources.</summary>
    public string? SiteName { get; init; }

    /// <summary>Site URL, for SharePoint sources.</summary>
    public string? SiteUrl { get; init; }

    /// <summary>Owning user's display name, for a User OneDrive source.</summary>
    public string? UserDisplayName { get; init; }

    /// <summary>Friendly document library or drive name.</summary>
    public string? LibraryOrDriveName { get; init; }

    /// <summary>Folder the manifest's scope starts at.</summary>
    public string StartingFolder { get; init; } = "/";

    /// <summary>Whether the scope included subfolders.</summary>
    public bool Recursive { get; init; }

    /// <summary>Requested link permission.</summary>
    public required string LinkPermission { get; init; }

    /// <summary>Requested link audience.</summary>
    public required string LinkAudience { get; init; }

    /// <summary>Files that produced an entry.</summary>
    public int SuccessfulFiles { get; init; }

    /// <summary>Existing links returned rather than created.</summary>
    public int ReusedLinks { get; init; }

    /// <summary>Files skipped.</summary>
    public int SkippedFiles { get; init; }

    /// <summary>Files that failed.</summary>
    public int FailedFiles { get; init; }
}

/// <summary>A complete manifest: its header and its entries.</summary>
public sealed record ManifestDocument
{
    /// <summary>Provenance block.</summary>
    public required ManifestHeader Header { get; init; }

    /// <summary>File records.</summary>
    public IReadOnlyList<ManifestEntry> Entries { get; init; } = [];

    /// <summary>True when this document was parsed from a manifest this application wrote.</summary>
    public bool WasGeneratedByThisApplication { get; init; } = true;
}

/// <summary>The result of attempting to parse an existing manifest before updating it.</summary>
public sealed record ManifestParseResult
{
    /// <summary>True when the content was recognised and parsed.</summary>
    public required bool Succeeded { get; init; }

    /// <summary>The parsed document, when parsing succeeded.</summary>
    public ManifestDocument? Document { get; init; }

    /// <summary>
    /// Why parsing failed. A failure is not an error condition: it means the destination holds
    /// a file this application did not write, so it must not be overwritten.
    /// </summary>
    public string? FailureReason { get; init; }

    /// <summary>Creates a successful parse result.</summary>
    public static ManifestParseResult Success(ManifestDocument document) =>
        new() { Succeeded = true, Document = document };

    /// <summary>Creates a failed parse result.</summary>
    public static ManifestParseResult Failure(string reason) =>
        new() { Succeeded = false, FailureReason = reason };
}
