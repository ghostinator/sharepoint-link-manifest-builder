using SharePointLinkManifestBuilder.Core.Models;

namespace SharePointLinkManifestBuilder.Core.Tests;

/// <summary>Builders for test fixtures. All values are synthetic.</summary>
internal static class TestData
{
    public const string TenantId = "11111111-1111-1111-1111-111111111111";
    public const string DriveA = "drive-a";
    public const string DriveB = "drive-b";
    public const string SiteA = "example.sharepoint.test,site-a,web-a";

    public static DiscoveredFile File(
        string name,
        string relativePath = "",
        string driveId = DriveA,
        string? itemId = null,
        DriveItemKind kind = DriveItemKind.File,
        long? size = 1024,
        DateTimeOffset? modified = null,
        string parentRelativePath = "",
        string? parentFolderItemId = "folder-1") => new()
        {
            DriveId = driveId,
            ItemId = itemId ?? "item-" + name,
            Name = name,
            RelativePath = string.IsNullOrEmpty(relativePath) ? name : relativePath,
            WebUrl = $"https://example.sharepoint.test/sites/A/Shared%20Documents/{Uri.EscapeDataString(name)}",
            Size = size,
            LastModifiedUtc = modified ?? new DateTimeOffset(2026, 1, 15, 12, 0, 0, TimeSpan.Zero),
            Kind = kind,
            ParentRelativePath = parentRelativePath,
            ParentFolderItemId = parentFolderItemId,
        };

    public static ProcessingTarget Target(
        TargetSourceType sourceType = TargetSourceType.SharePointFolder,
        string? driveId = DriveA,
        string startingPath = "",
        bool recursive = false,
        string? siteId = SiteA,
        string? targetId = null) => new()
        {
            TargetId = targetId ?? Guid.NewGuid().ToString("n"),
            SourceType = sourceType,
            TenantId = TenantId,
            SiteId = siteId,
            SiteName = "Example Site",
            SiteUrl = "https://example.sharepoint.test/sites/A",
            DriveId = driveId,
            DriveName = "Documents",
            StartingFolderRelativePath = startingPath,
            Recursive = recursive,
        };

    public static LinkResult Result(
        DiscoveredFile file,
        LinkResultStatus status = LinkResultStatus.Created,
        string? sharingUrl = "https://example.sharepoint.test/:w:/s/A/EXAMPLE") => new()
        {
            File = file,
            Status = status,
            SharingUrl = sharingUrl,
            TimestampUtc = new DateTimeOffset(2026, 2, 1, 9, 30, 0, TimeSpan.Zero),
        };

    public static JobConfiguration Job(
        IReadOnlyList<ProcessingTarget>? targets = null,
        LinkConfiguration? link = null,
        ManifestConfiguration? manifest = null) => new()
        {
            JobId = "job-0001",
            TenantId = TenantId,
            TenantDisplayName = "Example Organization",
            Targets = targets ?? [Target()],
            Link = link ?? new LinkConfiguration(),
            Manifest = manifest ?? new ManifestConfiguration(),
        };
}
