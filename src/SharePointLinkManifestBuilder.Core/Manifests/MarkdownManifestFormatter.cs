using System.Globalization;
using System.Text;
using SharePointLinkManifestBuilder.Core.Abstractions;
using SharePointLinkManifestBuilder.Core.Models;
using SharePointLinkManifestBuilder.Core.Security;

namespace SharePointLinkManifestBuilder.Core.Manifests;

/// <summary>
/// Writes the Markdown manifest. All untrusted values are escaped so a file name cannot
/// restructure the document or forge a link target.
/// </summary>
public sealed class MarkdownManifestFormatter : IManifestFormatter
{
    /// <inheritdoc />
    public ManifestFormats Format => ManifestFormats.Markdown;

    /// <inheritdoc />
    public string FileExtension => ".md";

    /// <inheritdoc />
    public string Render(ManifestDocument document)
    {
        ArgumentNullException.ThrowIfNull(document);

        var header = document.Header;
        var builder = new StringBuilder();

        builder.Append("# SharePoint Link Manifest\n\n");

        builder.Append("| Field | Value |\n| --- | --- |\n");
        AppendRow(builder, "Schema Version", header.SchemaVersion);
        AppendRow(builder, "Application Version", header.ApplicationVersion);
        AppendRow(builder, "Job ID", header.JobId);
        AppendRow(builder, "Generated", PlainTextManifestFormatter.FormatTimestamp(header.GeneratedUtc));
        AppendRow(builder, "Tenant", header.TenantDisplayName);
        AppendRow(builder, "Tenant ID", header.TenantId);
        AppendRow(builder, "Source Type", header.SourceType);
        AppendOptionalRow(builder, "Site", header.SiteName);
        AppendOptionalRow(builder, "Site URL", header.SiteUrl);
        AppendOptionalRow(builder, "User", header.UserDisplayName);
        AppendOptionalRow(builder, "Document Library or Drive", header.LibraryOrDriveName);
        AppendRow(builder, "Starting Folder", string.IsNullOrEmpty(header.StartingFolder) ? "/" : header.StartingFolder);
        AppendRow(builder, "Recursive", header.Recursive ? "Yes" : "No");
        AppendRow(builder, "Link Permission", header.LinkPermission);
        AppendRow(builder, "Link Audience", header.LinkAudience);
        AppendRow(builder, "Successful Files", header.SuccessfulFiles.ToString(CultureInfo.InvariantCulture));
        AppendRow(builder, "Reused Links", header.ReusedLinks.ToString(CultureInfo.InvariantCulture));
        AppendRow(builder, "Skipped Files", header.SkippedFiles.ToString(CultureInfo.InvariantCulture));
        AppendRow(builder, "Failed Files", header.FailedFiles.ToString(CultureInfo.InvariantCulture));

        builder.Append("\n## Files\n\n");

        if (document.Entries.Count == 0)
        {
            builder.Append("_No files were successfully processed._\n");
            return builder.ToString();
        }

        builder.Append("| File | Relative Path | Sharing Link | Status | Generated |\n");
        builder.Append("| --- | --- | --- | --- | --- |\n");

        foreach (var entry in document.Entries)
        {
            var link = string.IsNullOrWhiteSpace(entry.SharingUrl)
                ? "_none_"
                : MarkdownEscaper.Link("Open", entry.SharingUrl);

            var status = entry.IsMissing ? entry.Status + " (not found on last run)" : entry.Status;

            builder
                .Append("| ").Append(MarkdownEscaper.EscapeTableCell(entry.FileName))
                .Append(" | ").Append(MarkdownEscaper.EscapeTableCell(entry.RelativePath))
                .Append(" | ").Append(link)
                .Append(" | ").Append(MarkdownEscaper.EscapeTableCell(status))
                .Append(" | ").Append(PlainTextManifestFormatter.FormatTimestamp(entry.GeneratedUtc))
                .Append(" |\n");
        }

        return builder.ToString();
    }

    private static void AppendRow(StringBuilder builder, string name, string value) =>
        builder.Append("| ").Append(MarkdownEscaper.EscapeTableCell(name))
               .Append(" | ").Append(MarkdownEscaper.EscapeTableCell(value)).Append(" |\n");

    private static void AppendOptionalRow(StringBuilder builder, string name, string? value)
    {
        if (!string.IsNullOrWhiteSpace(value))
        {
            AppendRow(builder, name, value);
        }
    }
}
