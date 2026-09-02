namespace SharePointLinkManifestBuilder.Core.Models;

/// <summary>The kind of node shown in the unified resource tree.</summary>
public enum ResourceKind
{
    /// <summary>A SharePoint site.</summary>
    Site = 0,

    /// <summary>A document library or a OneDrive drive.</summary>
    Drive = 1,

    /// <summary>A folder inside a drive.</summary>
    Folder = 2,

    /// <summary>A file inside a drive.</summary>
    File = 3,

    /// <summary>A grouping node such as "SharePoint Sites" or "My OneDrive".</summary>
    Category = 4,

    /// <summary>A user whose OneDrive can be browsed.</summary>
    User = 5,
}

/// <summary>A SharePoint site as returned by Microsoft Graph.</summary>
public sealed record SharePointSite
{
    /// <summary>The composite Graph site ID (<c>hostname,siteCollectionId,siteId</c>).</summary>
    public required string SiteId { get; init; }

    /// <summary>Site title. Falls back to the name when a display name is absent.</summary>
    public required string DisplayName { get; init; }

    /// <summary>Absolute site URL.</summary>
    public string? WebUrl { get; init; }

    /// <summary>SharePoint hostname, useful when several are in play.</summary>
    public string? Hostname { get; init; }

    /// <summary>Description, when Graph returns one.</summary>
    public string? Description { get; init; }

    /// <summary>True when this is the tenant root site.</summary>
    public bool IsRootSite { get; init; }

    /// <summary>When the site was created, when Graph returns it.</summary>
    public DateTimeOffset? CreatedUtc { get; init; }
}

/// <summary>A document library or OneDrive drive.</summary>
public sealed record DriveResource
{
    /// <summary>Graph drive ID.</summary>
    public required string DriveId { get; init; }

    /// <summary>Friendly library or drive name, shown to the user in preference to any ID.</summary>
    public required string Name { get; init; }

    /// <summary>Graph <c>driveType</c>: <c>documentLibrary</c>, <c>business</c>, <c>personal</c>.</summary>
    public string? DriveType { get; init; }

    /// <summary>Absolute URL of the library or drive.</summary>
    public string? WebUrl { get; init; }

    /// <summary>Owning site ID, when this drive belongs to a SharePoint site.</summary>
    public string? SiteId { get; init; }

    /// <summary>Owning site name, for display.</summary>
    public string? SiteName { get; init; }

    /// <summary>Owner display name, when this is a personal OneDrive.</summary>
    public string? OwnerDisplayName { get; init; }

    /// <summary>Owning user ID, when this is a personal OneDrive.</summary>
    public string? OwnerUserId { get; init; }

    /// <summary>True when this drive is a user's OneDrive rather than a SharePoint library.</summary>
    public bool IsPersonal =>
        string.Equals(DriveType, "personal", StringComparison.OrdinalIgnoreCase) ||
        string.Equals(DriveType, "business", StringComparison.OrdinalIgnoreCase);
}

/// <summary>A user whose OneDrive may be browsable, subject to authorization.</summary>
public sealed record OneDriveUser
{
    /// <summary>Entra object ID.</summary>
    public required string UserId { get; init; }

    /// <summary>Display name as returned by Graph.</summary>
    public required string DisplayName { get; init; }

    /// <summary>User principal name as returned by Graph.</summary>
    public string? UserPrincipalName { get; init; }

    /// <summary>Job title, only when Graph returns it.</summary>
    public string? JobTitle { get; init; }
}

/// <summary>A folder within a drive.</summary>
public sealed record SharePointFolder
{
    /// <summary>Drive containing the folder.</summary>
    public required string DriveId { get; init; }

    /// <summary>Graph item ID of the folder.</summary>
    public required string ItemId { get; init; }

    /// <summary>Folder name.</summary>
    public required string Name { get; init; }

    /// <summary>Path relative to the drive root, using forward slashes. Empty for the root.</summary>
    public string RelativePath { get; init; } = string.Empty;

    /// <summary>Absolute folder URL.</summary>
    public string? WebUrl { get; init; }

    /// <summary>Number of direct children, when Graph reports it.</summary>
    public int? ChildCount { get; init; }

    /// <summary>True when this node is the drive root.</summary>
    public bool IsRoot { get; init; }
}
