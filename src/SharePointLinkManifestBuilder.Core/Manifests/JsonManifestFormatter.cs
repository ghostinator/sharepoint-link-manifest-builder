using System.Text.Json;
using System.Text.Json.Serialization;
using SharePointLinkManifestBuilder.Core.Abstractions;
using SharePointLinkManifestBuilder.Core.Models;

namespace SharePointLinkManifestBuilder.Core.Manifests;

/// <summary>
/// Writes the JSON manifest against a documented, versioned schema with stable camelCase
/// property names, so downstream consumers can bind to it and future versions can extend it.
/// </summary>
public sealed class JsonManifestFormatter : IManifestFormatter
{
    /// <summary>The schema identifier written into every JSON manifest.</summary>
    public const string SchemaId =
        "https://example.invalid/PLACEHOLDER-SOURCE/schemas/sharepoint-link-manifest/v1.json";

    private static readonly JsonSerializerOptions SerializerOptions = new()
    {
        WriteIndented = true,
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull,

        // The strict default encoder escapes characters that could be dangerous if the JSON is
        // ever embedded in HTML. Manifests contain untrusted file names, so the strict default
        // is kept deliberately rather than relaxed for prettier output.
        Encoder = System.Text.Encodings.Web.JavaScriptEncoder.Default,
    };

    /// <inheritdoc />
    public ManifestFormats Format => ManifestFormats.Json;

    /// <inheritdoc />
    public string FileExtension => ".json";

    /// <inheritdoc />
    public string Render(ManifestDocument document)
    {
        ArgumentNullException.ThrowIfNull(document);

        var payload = new JsonManifest
        {
            Schema = SchemaId,
            SchemaVersion = document.Header.SchemaVersion,
            ApplicationVersion = document.Header.ApplicationVersion,
            JobId = document.Header.JobId,
            GeneratedUtc = document.Header.GeneratedUtc,
            Tenant = new JsonTenant
            {
                DisplayName = document.Header.TenantDisplayName,
                TenantId = document.Header.TenantId,
            },
            Source = new JsonSource
            {
                SourceType = document.Header.SourceType,
                SiteName = document.Header.SiteName,
                SiteUrl = document.Header.SiteUrl,
                UserDisplayName = document.Header.UserDisplayName,
                LibraryOrDriveName = document.Header.LibraryOrDriveName,
                StartingFolder = document.Header.StartingFolder,
                Recursive = document.Header.Recursive,
            },
            Link = new JsonLink
            {
                Permission = document.Header.LinkPermission,
                Audience = document.Header.LinkAudience,
            },
            Counts = new JsonCounts
            {
                SuccessfulFiles = document.Header.SuccessfulFiles,
                ReusedLinks = document.Header.ReusedLinks,
                SkippedFiles = document.Header.SkippedFiles,
                FailedFiles = document.Header.FailedFiles,
            },
            Files = document.Entries.Select(entry => new JsonEntry
            {
                FileName = entry.FileName,
                RelativePath = entry.RelativePath,
                WebUrl = entry.WebUrl,
                SharingUrl = entry.SharingUrl,
                DriveId = entry.DriveId,
                ItemId = entry.ItemId,
                Status = entry.Status,
                GeneratedUtc = entry.GeneratedUtc,
                ErrorCode = entry.ErrorCode,
                ErrorMessage = entry.ErrorMessage,
                IsMissing = entry.IsMissing ? true : null,
            }).ToArray(),
        };

        return JsonSerializer.Serialize(payload, SerializerOptions);
    }

    private sealed record JsonManifest
    {
        [JsonPropertyName("$schema")]
        public required string Schema { get; init; }

        public required string SchemaVersion { get; init; }

        public required string ApplicationVersion { get; init; }

        public required string JobId { get; init; }

        public required DateTimeOffset GeneratedUtc { get; init; }

        public required JsonTenant Tenant { get; init; }

        public required JsonSource Source { get; init; }

        public required JsonLink Link { get; init; }

        public required JsonCounts Counts { get; init; }

        public required IReadOnlyList<JsonEntry> Files { get; init; }
    }

    private sealed record JsonTenant
    {
        public required string DisplayName { get; init; }

        public required string TenantId { get; init; }
    }

    private sealed record JsonSource
    {
        public required string SourceType { get; init; }

        public string? SiteName { get; init; }

        public string? SiteUrl { get; init; }

        public string? UserDisplayName { get; init; }

        public string? LibraryOrDriveName { get; init; }

        public required string StartingFolder { get; init; }

        public required bool Recursive { get; init; }
    }

    private sealed record JsonLink
    {
        public required string Permission { get; init; }

        public required string Audience { get; init; }
    }

    private sealed record JsonCounts
    {
        public required int SuccessfulFiles { get; init; }

        public required int ReusedLinks { get; init; }

        public required int SkippedFiles { get; init; }

        public required int FailedFiles { get; init; }
    }

    private sealed record JsonEntry
    {
        public required string FileName { get; init; }

        public required string RelativePath { get; init; }

        public string? WebUrl { get; init; }

        public string? SharingUrl { get; init; }

        public required string DriveId { get; init; }

        public required string ItemId { get; init; }

        public required string Status { get; init; }

        public required DateTimeOffset GeneratedUtc { get; init; }

        public string? ErrorCode { get; init; }

        public string? ErrorMessage { get; init; }

        public bool? IsMissing { get; init; }
    }
}
