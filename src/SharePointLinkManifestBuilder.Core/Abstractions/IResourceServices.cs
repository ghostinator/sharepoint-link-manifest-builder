using SharePointLinkManifestBuilder.Core.Models;

namespace SharePointLinkManifestBuilder.Core.Abstractions;

/// <summary>SharePoint site discovery and metadata.</summary>
public interface ISiteService
{
    /// <summary>
    /// Searches sites the signed-in user can see.
    /// <para>
    /// This does <em>not</em> return every site in the tenant: Graph returns what the search
    /// index exposes to this user. Callers must not present the result as exhaustive.
    /// </para>
    /// </summary>
    Task<OperationResult<IReadOnlyList<SharePointSite>>> SearchSitesAsync(
        string query,
        CancellationToken cancellationToken = default);

    /// <summary>Resolves a pasted SharePoint URL to a site. Never fetches the URL directly.</summary>
    Task<OperationResult<SharePointSite>> ResolveSiteByUrlAsync(
        string siteUrl,
        CancellationToken cancellationToken = default);

    /// <summary>Reads a site by its Graph ID.</summary>
    Task<OperationResult<SharePointSite>> GetSiteAsync(
        string siteId,
        CancellationToken cancellationToken = default);

    /// <summary>Reads the tenant root site.</summary>
    Task<OperationResult<SharePointSite>> GetRootSiteAsync(CancellationToken cancellationToken = default);

    /// <summary>Sites the signed-in user follows, used to populate Recent.</summary>
    Task<OperationResult<IReadOnlyList<SharePointSite>>> GetFollowedSitesAsync(
        CancellationToken cancellationToken = default);

    /// <summary>Lists the document libraries of a site.</summary>
    Task<OperationResult<IReadOnlyList<DriveResource>>> GetSiteDrivesAsync(
        string siteId,
        CancellationToken cancellationToken = default);
}

/// <summary>Drive and item resolution for SharePoint libraries and OneDrive.</summary>
public interface IDriveService
{
    /// <summary>Resolves the signed-in user's own OneDrive.</summary>
    Task<OperationResult<DriveResource>> GetMyDriveAsync(CancellationToken cancellationToken = default);

    /// <summary>
    /// Resolves another user's OneDrive.
    /// <para>
    /// May legitimately fail: the drive can be unprovisioned, access-denied, or blocked by
    /// policy. Administrator consent does not by itself grant access to every user's OneDrive,
    /// and this method never provisions a drive.
    /// </para>
    /// </summary>
    Task<OperationResult<DriveResource>> GetUserDriveAsync(
        string userId,
        CancellationToken cancellationToken = default);

    /// <summary>Reads a drive by ID.</summary>
    Task<OperationResult<DriveResource>> GetDriveAsync(
        string driveId,
        CancellationToken cancellationToken = default);

    /// <summary>Reads a drive's root folder.</summary>
    Task<OperationResult<SharePointFolder>> GetRootFolderAsync(
        string driveId,
        CancellationToken cancellationToken = default);

    /// <summary>Resolves a folder by its path relative to the drive root.</summary>
    Task<OperationResult<SharePointFolder>> GetFolderByPathAsync(
        string driveId,
        string relativePath,
        CancellationToken cancellationToken = default);

    /// <summary>Reads a single item by ID.</summary>
    Task<OperationResult<DiscoveredFile>> GetItemAsync(
        string driveId,
        string itemId,
        CancellationToken cancellationToken = default);

    /// <summary>
    /// Lists the immediate children of a folder, following pagination. Streams lazily so a
    /// large folder never has to be held in memory at once.
    /// </summary>
    IAsyncEnumerable<DiscoveredFile> GetChildrenAsync(
        string driveId,
        string folderItemId,
        CancellationToken cancellationToken = default);

    /// <summary>Lists only the subfolders of a folder, for lazy tree expansion.</summary>
    Task<OperationResult<IReadOnlyList<SharePointFolder>>> GetSubfoldersAsync(
        string driveId,
        string folderItemId,
        CancellationToken cancellationToken = default);

    /// <summary>Resolves a pasted sharing URL through the Graph shares endpoint.</summary>
    Task<OperationResult<DiscoveredFile>> ResolveSharingUrlAsync(
        string sharingUrl,
        CancellationToken cancellationToken = default);
}

/// <summary>Directory lookup for the User OneDrive people picker.</summary>
public interface IUserDirectoryService
{
    /// <summary>
    /// Searches the directory for users. Returns only what Graph returns for the signed-in
    /// user; it never reveals information the caller is not authorized to read.
    /// </summary>
    Task<OperationResult<IReadOnlyList<OneDriveUser>>> SearchUsersAsync(
        string query,
        CancellationToken cancellationToken = default);

    /// <summary>Reads a user by ID or user principal name.</summary>
    Task<OperationResult<OneDriveUser>> GetUserAsync(
        string userIdOrUpn,
        CancellationToken cancellationToken = default);
}

/// <summary>Creating and inspecting sharing links.</summary>
public interface ISharingLinkService
{
    /// <summary>
    /// Requests a sharing link for one file and reports what actually happened.
    /// <para>
    /// Distinguishes a newly created link (HTTP 201) from an existing equivalent link returned
    /// by Graph (HTTP 200), and never reports the latter as created.
    /// </para>
    /// </summary>
    Task<LinkResult> CreateOrGetLinkAsync(
        DiscoveredFile file,
        LinkConfiguration configuration,
        CancellationToken cancellationToken = default);

    /// <summary>
    /// Grants named recipients access using the Graph invite action, which is the only v1.0
    /// operation that accepts recipients. Handles <c>207 Multi-Status</c> partial success.
    /// </summary>
    Task<IReadOnlyList<RecipientResult>> InviteRecipientsAsync(
        DiscoveredFile file,
        LinkConfiguration configuration,
        CancellationToken cancellationToken = default);

    /// <summary>Reads existing sharing links on an item, to support skip-if-exists.</summary>
    Task<OperationResult<IReadOnlyList<ExistingSharingLink>>> GetExistingLinksAsync(
        DiscoveredFile file,
        CancellationToken cancellationToken = default);
}

/// <summary>A sharing link that already exists on an item.</summary>
public sealed record ExistingSharingLink
{
    /// <summary>Graph permission ID.</summary>
    public required string PermissionId { get; init; }

    /// <summary>Graph link type, for example <c>view</c>.</summary>
    public string? LinkType { get; init; }

    /// <summary>Graph link scope, for example <c>organization</c>.</summary>
    public string? Scope { get; init; }

    /// <summary>The link URL.</summary>
    public string? WebUrl { get; init; }

    /// <summary>Expiry, when set.</summary>
    public DateTimeOffset? ExpirationUtc { get; init; }

    /// <summary>True when this link matches the requested type and scope.</summary>
    public bool Matches(LinkConfiguration configuration) =>
        string.Equals(LinkType, configuration.GraphLinkType, StringComparison.OrdinalIgnoreCase)
        && string.Equals(Scope, configuration.GraphScope, StringComparison.OrdinalIgnoreCase);
}

/// <summary>Reading and writing manifest files in SharePoint and OneDrive.</summary>
public interface IManifestStorageService
{
    /// <summary>
    /// Reads an existing manifest, returning its content and ETag. A missing file is a normal
    /// outcome, not an error.
    /// </summary>
    Task<OperationResult<ExistingManifest?>> ReadManifestAsync(
        string driveId,
        string parentItemId,
        string fileName,
        CancellationToken cancellationToken = default);

    /// <summary>
    /// Writes a manifest. When <paramref name="ifMatchETag"/> is supplied the write is
    /// conditional and a concurrent remote change surfaces as an ETag conflict rather than
    /// silently overwriting.
    /// </summary>
    Task<OperationResult<ManifestWriteResult>> WriteManifestAsync(
        string driveId,
        string parentItemId,
        string fileName,
        string content,
        ManifestFormats format,
        bool isMaster,
        int entryCount,
        string? ifMatchETag,
        CancellationToken cancellationToken = default);
}

/// <summary>An existing manifest read from a drive.</summary>
public sealed record ExistingManifest
{
    /// <summary>Graph item ID of the file.</summary>
    public required string ItemId { get; init; }

    /// <summary>Decoded UTF-8 content.</summary>
    public required string Content { get; init; }

    /// <summary>ETag used to make the subsequent write conditional.</summary>
    public string? ETag { get; init; }

    /// <summary>Absolute URL of the file.</summary>
    public string? WebUrl { get; init; }

    /// <summary>Last modification timestamp.</summary>
    public DateTimeOffset? LastModifiedUtc { get; init; }
}
