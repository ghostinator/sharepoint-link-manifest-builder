using SharePointLinkManifestBuilder.Core.Abstractions;
using SharePointLinkManifestBuilder.Core.Models;
using SharePointLinkManifestBuilder.Core.Security;

namespace SharePointLinkManifestBuilder.Core.Manifests;

/// <summary>
/// Writes the CSV manifest. Every field passes through <see cref="CsvSanitizer"/>, because file
/// names come from SharePoint and a name beginning with <c>=</c> would otherwise become a live
/// formula when the export is opened.
/// </summary>
public sealed class CsvManifestFormatter : IManifestFormatter
{
    /// <summary>The column order, which is part of the documented output contract.</summary>
    public static readonly IReadOnlyList<string> Columns =
    [
        "Filename", "RelativePath", "WebUrl", "SharingUrl", "Tenant", "SourceType",
        "Site", "LibraryOrDrive", "DriveId", "ItemId", "Status", "ErrorCode",
        "ErrorMessage", "Timestamp",
    ];

    /// <inheritdoc />
    public ManifestFormats Format => ManifestFormats.Csv;

    /// <inheritdoc />
    public string FileExtension => ".csv";

    /// <inheritdoc />
    public string Render(ManifestDocument document)
    {
        ArgumentNullException.ThrowIfNull(document);

        var header = document.Header;

        var rows = document.Entries.Select(entry => new string?[]
        {
            entry.FileName,
            entry.RelativePath,
            entry.WebUrl,
            entry.SharingUrl,
            header.TenantDisplayName,
            header.SourceType,
            header.SiteName,
            header.LibraryOrDriveName,
            entry.DriveId,
            entry.ItemId,
            entry.IsMissing ? entry.Status + " (not found on last run)" : entry.Status,
            entry.ErrorCode,
            entry.ErrorMessage,
            CsvSanitizer.FormatTimestamp(entry.GeneratedUtc),
        });

        return CsvSanitizer.BuildDocument(Columns, rows);
    }
}
