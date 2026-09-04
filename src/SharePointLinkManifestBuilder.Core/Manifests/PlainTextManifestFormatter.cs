using System.Globalization;
using System.Text;
using SharePointLinkManifestBuilder.Core.Abstractions;
using SharePointLinkManifestBuilder.Core.Models;

namespace SharePointLinkManifestBuilder.Core.Manifests;

/// <summary>
/// Writes the plain-text manifest format, which is the default and the only format this
/// application parses back for update mode.
/// <para>
/// Encoded as UTF-8 without a byte-order mark. A BOM would appear as stray characters to many
/// text consumers, and nothing in the pipeline requires one.
/// </para>
/// </summary>
public sealed class PlainTextManifestFormatter : IManifestFormatter
{
    /// <summary>The first line of every manifest, used to recognise the format.</summary>
    public const string DocumentHeader = "SharePoint Link Manifest";

    /// <summary>The separator written after each entry.</summary>
    public const string EntrySeparator = "---";

    /// <inheritdoc />
    public ManifestFormats Format => ManifestFormats.PlainText;

    /// <inheritdoc />
    public string FileExtension => ".txt";

    /// <inheritdoc />
    public string Render(ManifestDocument document)
    {
        ArgumentNullException.ThrowIfNull(document);

        var header = document.Header;
        var builder = new StringBuilder();

        builder.Append(DocumentHeader).Append('\n');
        AppendField(builder, "Schema Version", header.SchemaVersion);
        AppendField(builder, "Application Version", header.ApplicationVersion);
        AppendField(builder, "Job ID", header.JobId);
        AppendField(builder, "Generated", FormatTimestamp(header.GeneratedUtc));
        AppendField(builder, "Tenant", header.TenantDisplayName);
        AppendField(builder, "Tenant ID", header.TenantId);
        AppendField(builder, "Source Type", header.SourceType);

        // "When applicable" fields are omitted entirely rather than written empty, so a reader
        // never has to distinguish "absent" from "blank".
        AppendOptionalField(builder, "Site", header.SiteName);
        AppendOptionalField(builder, "Site URL", header.SiteUrl);
        AppendOptionalField(builder, "User", header.UserDisplayName);
        AppendOptionalField(builder, "Document Library or Drive", header.LibraryOrDriveName);

        AppendField(builder, "Starting Folder", string.IsNullOrEmpty(header.StartingFolder) ? "/" : header.StartingFolder);
        AppendField(builder, "Recursive", header.Recursive ? "Yes" : "No");
        AppendField(builder, "Link Permission", header.LinkPermission);
        AppendField(builder, "Link Audience", header.LinkAudience);
        AppendField(builder, "Successful Files", header.SuccessfulFiles.ToString(CultureInfo.InvariantCulture));
        AppendField(builder, "Reused Links", header.ReusedLinks.ToString(CultureInfo.InvariantCulture));
        AppendField(builder, "Skipped Files", header.SkippedFiles.ToString(CultureInfo.InvariantCulture));
        AppendField(builder, "Failed Files", header.FailedFiles.ToString(CultureInfo.InvariantCulture));

        builder.Append('\n');

        foreach (var entry in document.Entries)
        {
            AppendField(builder, "File", entry.FileName);
            AppendField(builder, "Relative Path", entry.RelativePath);
            AppendField(builder, "Web URL", entry.WebUrl ?? string.Empty);
            AppendField(builder, "Sharing Link", entry.SharingUrl ?? string.Empty);
            AppendField(builder, "Drive ID", entry.DriveId);
            AppendField(builder, "Item ID", entry.ItemId);
            AppendField(builder, "Status", entry.IsMissing ? entry.Status + " (not found on last run)" : entry.Status);
            AppendField(builder, "Generated", FormatTimestamp(entry.GeneratedUtc));
            builder.Append(EntrySeparator).Append('\n');
        }

        return builder.ToString();
    }

    /// <summary>Formats a timestamp as ISO 8601 in UTC.</summary>
    public static string FormatTimestamp(DateTimeOffset value) =>
        value.UtcDateTime.ToString("yyyy-MM-dd'T'HH:mm:ss'Z'", CultureInfo.InvariantCulture);

    /// <summary>
    /// Collapses line breaks in a value. A SharePoint file name cannot contain a newline, but
    /// treating that as guaranteed would let one malformed value corrupt the whole document.
    /// </summary>
    private static string SingleLine(string value) =>
        value.Replace("\r\n", " ", StringComparison.Ordinal)
             .Replace('\r', ' ')
             .Replace('\n', ' ');

    private static void AppendField(StringBuilder builder, string name, string value) =>
        builder.Append(name).Append(':').Append(' ').Append(SingleLine(value)).Append('\n');

    private static void AppendOptionalField(StringBuilder builder, string name, string? value)
    {
        if (!string.IsNullOrWhiteSpace(value))
        {
            AppendField(builder, name, value);
        }
    }
}
