using System.Globalization;
using SharePointLinkManifestBuilder.Core.Abstractions;
using SharePointLinkManifestBuilder.Core.Models;

namespace SharePointLinkManifestBuilder.Core.Manifests;

/// <summary>
/// Reads a plain-text manifest previously written by this application, so it can be updated in
/// place rather than overwritten.
/// <para>
/// Parsing is defensive. Unknown fields are ignored, malformed entries are dropped rather than
/// throwing, and no value is ever interpreted as an instruction. A parse failure is a normal
/// outcome meaning "this application did not write this file", which the conflict policy uses
/// to avoid destroying someone else's content.
/// </para>
/// </summary>
public sealed class PlainTextManifestParser : IManifestParser
{
    /// <inheritdoc />
    public ManifestFormats Format => ManifestFormats.PlainText;

    /// <inheritdoc />
    public ManifestParseResult TryParse(string content)
    {
        if (string.IsNullOrWhiteSpace(content))
        {
            return ManifestParseResult.Failure("The file is empty.");
        }

        var lines = content.Replace("\r\n", "\n", StringComparison.Ordinal).Split('\n');

        if (!lines[0].Trim().Equals(PlainTextManifestFormatter.DocumentHeader, StringComparison.Ordinal))
        {
            return ManifestParseResult.Failure(
                "The file does not begin with the manifest header, so it was not written by this application.");
        }

        var headerFields = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        var index = 1;

        for (; index < lines.Length; index++)
        {
            var line = lines[index];
            if (string.IsNullOrWhiteSpace(line))
            {
                index++;
                break;
            }

            if (TrySplitField(line, out var key, out var value))
            {
                headerFields[key] = value;
            }
        }

        if (!headerFields.TryGetValue("Schema Version", out var schemaVersion))
        {
            return ManifestParseResult.Failure("The manifest has no schema version.");
        }

        // Reject a major version this build does not understand rather than silently
        // mis-parsing it and then overwriting the file.
        if (!IsSupportedSchemaVersion(schemaVersion))
        {
            return ManifestParseResult.Failure(
                $"Manifest schema version '{schemaVersion}' is newer than this application supports "
                + $"(supported: {ManifestDefaults.SchemaVersion}).");
        }

        var header = new ManifestHeader
        {
            SchemaVersion = schemaVersion,
            ApplicationVersion = Get(headerFields, "Application Version", "unknown"),
            JobId = Get(headerFields, "Job ID", "unknown"),
            GeneratedUtc = ParseTimestamp(Get(headerFields, "Generated", string.Empty)),
            TenantDisplayName = Get(headerFields, "Tenant", "unknown"),
            TenantId = Get(headerFields, "Tenant ID", "unknown"),
            SourceType = Get(headerFields, "Source Type", "unknown"),
            SiteName = GetOptional(headerFields, "Site"),
            SiteUrl = GetOptional(headerFields, "Site URL"),
            UserDisplayName = GetOptional(headerFields, "User"),
            LibraryOrDriveName = GetOptional(headerFields, "Document Library or Drive"),
            StartingFolder = Get(headerFields, "Starting Folder", "/"),
            Recursive = string.Equals(GetOptional(headerFields, "Recursive"), "Yes", StringComparison.OrdinalIgnoreCase),
            LinkPermission = Get(headerFields, "Link Permission", "unknown"),
            LinkAudience = Get(headerFields, "Link Audience", "unknown"),
            SuccessfulFiles = ParseInt(GetOptional(headerFields, "Successful Files")),
            ReusedLinks = ParseInt(GetOptional(headerFields, "Reused Links")),
            SkippedFiles = ParseInt(GetOptional(headerFields, "Skipped Files")),
            FailedFiles = ParseInt(GetOptional(headerFields, "Failed Files")),
        };

        var entries = new List<ManifestEntry>();
        var current = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        for (; index < lines.Length; index++)
        {
            var line = lines[index];

            if (line.Trim() == PlainTextManifestFormatter.EntrySeparator)
            {
                if (TryBuildEntry(current, out var entry))
                {
                    entries.Add(entry);
                }

                current.Clear();
                continue;
            }

            if (string.IsNullOrWhiteSpace(line))
            {
                continue;
            }

            if (TrySplitField(line, out var key, out var value))
            {
                current[key] = value;
            }
        }

        // Tolerate a final entry that was not followed by a separator.
        if (current.Count > 0 && TryBuildEntry(current, out var trailing))
        {
            entries.Add(trailing);
        }

        return ManifestParseResult.Success(new ManifestDocument
        {
            Header = header,
            Entries = entries,
            WasGeneratedByThisApplication = true,
        });
    }

    /// <summary>True when this build can safely read the given schema version.</summary>
    public static bool IsSupportedSchemaVersion(string? version)
    {
        if (string.IsNullOrWhiteSpace(version))
        {
            return false;
        }

        var supportedMajor = MajorOf(ManifestDefaults.SchemaVersion);
        var candidateMajor = MajorOf(version);

        return candidateMajor > 0 && candidateMajor <= supportedMajor;
    }

    private static int MajorOf(string version)
    {
        var head = version.Split('.')[0];
        return int.TryParse(head, NumberStyles.Integer, CultureInfo.InvariantCulture, out var major) ? major : 0;
    }

    private static bool TryBuildEntry(Dictionary<string, string> fields, out ManifestEntry entry)
    {
        entry = null!;

        // Identity is (driveId, itemId). An entry without both cannot be matched across runs,
        // so it is dropped rather than guessed at from the file name.
        if (!fields.TryGetValue("Drive ID", out var driveId) || string.IsNullOrWhiteSpace(driveId)
            || !fields.TryGetValue("Item ID", out var itemId) || string.IsNullOrWhiteSpace(itemId)
            || !fields.TryGetValue("File", out var fileName) || string.IsNullOrWhiteSpace(fileName))
        {
            return false;
        }

        var status = fields.GetValueOrDefault("Status", "Unknown");
        var isMissing = status.Contains("not found", StringComparison.OrdinalIgnoreCase);

        entry = new ManifestEntry
        {
            FileName = fileName,
            RelativePath = fields.GetValueOrDefault("Relative Path", string.Empty),
            WebUrl = NullIfEmpty(fields.GetValueOrDefault("Web URL")),
            SharingUrl = NullIfEmpty(fields.GetValueOrDefault("Sharing Link")),
            DriveId = driveId,
            ItemId = itemId,
            Status = isMissing ? status.Split('(')[0].Trim() : status,
            GeneratedUtc = ParseTimestamp(fields.GetValueOrDefault("Generated", string.Empty)),
            IsMissing = isMissing,
        };

        return true;
    }

    private static bool TrySplitField(string line, out string key, out string value)
    {
        key = string.Empty;
        value = string.Empty;

        var separator = line.IndexOf(':', StringComparison.Ordinal);
        if (separator <= 0)
        {
            return false;
        }

        key = line[..separator].Trim();
        value = line[(separator + 1)..].Trim();
        return key.Length > 0;
    }

    private static string Get(Dictionary<string, string> fields, string key, string fallback) =>
        fields.TryGetValue(key, out var value) && !string.IsNullOrWhiteSpace(value) ? value : fallback;

    private static string? GetOptional(Dictionary<string, string> fields, string key) =>
        fields.TryGetValue(key, out var value) && !string.IsNullOrWhiteSpace(value) ? value : null;

    private static string? NullIfEmpty(string? value) =>
        string.IsNullOrWhiteSpace(value) ? null : value;

    private static int ParseInt(string? value) =>
        int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out var result) ? result : 0;

    private static DateTimeOffset ParseTimestamp(string value) =>
        DateTimeOffset.TryParse(
            value,
            CultureInfo.InvariantCulture,
            DateTimeStyles.AdjustToUniversal | DateTimeStyles.AssumeUniversal,
            out var parsed)
            ? parsed
            : DateTimeOffset.MinValue;
}
