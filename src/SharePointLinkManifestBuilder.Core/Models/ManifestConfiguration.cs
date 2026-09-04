namespace SharePointLinkManifestBuilder.Core.Models;

/// <summary>Output formats a manifest can be written in.</summary>
[Flags]
public enum ManifestFormats
{
    /// <summary>No manifest output.</summary>
    None = 0,

    /// <summary>Plain UTF-8 text. Enabled by default.</summary>
    PlainText = 1,

    /// <summary>Markdown, with untrusted content escaped.</summary>
    Markdown = 2,

    /// <summary>CSV, protected against formula injection.</summary>
    Csv = 4,

    /// <summary>JSON with a documented, versioned schema.</summary>
    Json = 8,
}

/// <summary>What to do when a manifest already exists at the destination.</summary>
public enum ManifestConflictPolicy
{
    /// <summary>
    /// Merge into the existing manifest when this application generated it and it parses.
    /// Falls back to <see cref="CreateTimestampedVersion"/> otherwise. The default.
    /// </summary>
    UpdateSafely = 0,

    /// <summary>Overwrite the existing file unconditionally.</summary>
    Replace = 1,

    /// <summary>Write a new file with a timestamp suffix, leaving the existing file untouched.</summary>
    CreateTimestampedVersion = 2,

    /// <summary>Leave the existing file alone and record the manifest as skipped.</summary>
    Skip = 3,

    /// <summary>Treat an existing manifest as an error.</summary>
    Fail = 4,
}

/// <summary>What happens to manifest entries whose files were not seen on this run.</summary>
public enum MissingEntryPolicy
{
    /// <summary>Leave them in place, unchanged. The default: least destructive.</summary>
    Preserve = 0,

    /// <summary>Keep them but annotate them as not found on the latest run.</summary>
    Mark = 1,

    /// <summary>Remove them from the manifest.</summary>
    Remove = 2,
}

/// <summary>Where a combined master manifest for a multi-target job is written.</summary>
public sealed record MasterManifestDestination
{
    /// <summary>Drive that will receive the manifest.</summary>
    public required string DriveId { get; init; }

    /// <summary>Item ID of the destination folder.</summary>
    public required string FolderItemId { get; init; }

    /// <summary>Friendly description of the destination, shown before the job runs.</summary>
    public required string DisplayPath { get; init; }

    /// <summary>Absolute URL of the destination folder.</summary>
    public string? WebUrl { get; init; }
}

/// <summary>How manifests are produced and written for a job.</summary>
public sealed record ManifestConfiguration
{
    /// <summary>Write one manifest inside each folder that contained processed files.</summary>
    public bool WritePerFolderManifest { get; init; } = true;

    /// <summary>Write one manifest covering every successful file beneath the target.</summary>
    public bool WriteMasterManifest { get; init; }

    /// <summary>Formats to emit. Plain text is enabled by default.</summary>
    public ManifestFormats Formats { get; init; } = ManifestFormats.PlainText;

    /// <summary>File name for per-folder manifests, without an extension.</summary>
    public string PerFolderFileName { get; init; } = ManifestDefaults.PerFolderBaseName;

    /// <summary>File name for master manifests, without an extension.</summary>
    public string MasterFileName { get; init; } = ManifestDefaults.MasterBaseName;

    /// <summary>
    /// When true, a per-folder manifest also lists files from descendant folders. Off by
    /// default because it duplicates entries across nested manifests.
    /// </summary>
    public bool AggregateDescendantsInPerFolderManifest { get; init; }

    /// <summary>Behaviour when a manifest already exists.</summary>
    public ManifestConflictPolicy ConflictPolicy { get; init; } = ManifestConflictPolicy.UpdateSafely;

    /// <summary>What to do with entries whose files were not seen on this run.</summary>
    public MissingEntryPolicy MissingEntryPolicy { get; init; } = MissingEntryPolicy.Preserve;

    /// <summary>
    /// For a multi-target job: true writes one master manifest per target; false writes a
    /// single combined manifest to <see cref="CombinedMasterDestination"/>.
    /// </summary>
    public bool MasterManifestPerTarget { get; init; } = true;

    /// <summary>
    /// Destination for a single combined master manifest. Required when
    /// <see cref="MasterManifestPerTarget"/> is false, and for whole-site targets that span
    /// several libraries. Never guessed.
    /// </summary>
    public MasterManifestDestination? CombinedMasterDestination { get; init; }

    /// <summary>Every file name this configuration can generate, excluded from discovery.</summary>
    public IReadOnlyList<string> GeneratedFileNames()
    {
        var names = new List<string>();
        foreach (var extension in ManifestDefaults.ExtensionsFor(Formats))
        {
            names.Add(PerFolderFileName + extension);
            names.Add(MasterFileName + extension);
        }

        return names;
    }

    /// <summary>Validates the configuration, returning one message per problem.</summary>
    public IReadOnlyList<string> Validate()
    {
        var problems = new List<string>();

        if (!WritePerFolderManifest && !WriteMasterManifest)
        {
            problems.Add("No manifest will be written: enable per-folder manifests, master manifests, or both.");
        }

        if (Formats == ManifestFormats.None)
        {
            problems.Add("No manifest format is selected.");
        }

        if (WriteMasterManifest && !MasterManifestPerTarget && CombinedMasterDestination is null)
        {
            problems.Add("A combined master manifest needs a writable destination library and folder.");
        }

        foreach (var name in new[] { PerFolderFileName, MasterFileName })
        {
            if (string.IsNullOrWhiteSpace(name) || name.IndexOfAny(ManifestDefaults.InvalidNameChars) >= 0)
            {
                problems.Add($"'{name}' is not a valid manifest file name.");
            }
        }

        return problems;
    }
}

/// <summary>Manifest naming and format conventions.</summary>
public static class ManifestDefaults
{
    /// <summary>Default base name for a per-folder manifest.</summary>
    public const string PerFolderBaseName = "_sharepoint-links";

    /// <summary>Default base name for a master manifest.</summary>
    public const string MasterBaseName = "_sharepoint-links-master";

    /// <summary>The manifest schema version this build reads and writes.</summary>
    public const string SchemaVersion = "1.0";

    /// <summary>Characters that may not appear in a SharePoint or OneDrive file name.</summary>
    public static readonly char[] InvalidNameChars = ['"', '*', ':', '<', '>', '?', '/', '\\', '|'];

    /// <summary>The file extension used for each format.</summary>
    public static string ExtensionFor(ManifestFormats format) => format switch
    {
        ManifestFormats.PlainText => ".txt",
        ManifestFormats.Markdown => ".md",
        ManifestFormats.Csv => ".csv",
        ManifestFormats.Json => ".json",
        _ => ".txt",
    };

    /// <summary>Expands a flags value into the individual formats it selects.</summary>
    public static IEnumerable<ManifestFormats> Split(ManifestFormats formats)
    {
        if (formats.HasFlag(ManifestFormats.PlainText))
        {
            yield return ManifestFormats.PlainText;
        }

        if (formats.HasFlag(ManifestFormats.Markdown))
        {
            yield return ManifestFormats.Markdown;
        }

        if (formats.HasFlag(ManifestFormats.Csv))
        {
            yield return ManifestFormats.Csv;
        }

        if (formats.HasFlag(ManifestFormats.Json))
        {
            yield return ManifestFormats.Json;
        }
    }

    /// <summary>The extensions selected by a flags value.</summary>
    public static IEnumerable<string> ExtensionsFor(ManifestFormats formats) =>
        Split(formats).Select(ExtensionFor);

    /// <summary>
    /// True when a file name looks like a manifest this application generates. Used to keep a
    /// job from discovering and re-processing its own output.
    /// </summary>
    public static bool IsGeneratedManifestName(string fileName)
    {
        var name = Path.GetFileNameWithoutExtension(fileName);
        var extension = Path.GetExtension(fileName);

        if (!ExtensionsFor(ManifestFormats.PlainText | ManifestFormats.Markdown
                | ManifestFormats.Csv | ManifestFormats.Json)
            .Contains(extension, StringComparer.OrdinalIgnoreCase))
        {
            return false;
        }

        // Matches the base names and their timestamped variants, e.g. "_sharepoint-links-20260902-101500".
        return name.StartsWith(PerFolderBaseName, StringComparison.OrdinalIgnoreCase)
            || name.StartsWith(MasterBaseName, StringComparison.OrdinalIgnoreCase);
    }
}
