using System.Globalization;
using SharePointLinkManifestBuilder.Core.Abstractions;
using SharePointLinkManifestBuilder.Core.Models;

namespace SharePointLinkManifestBuilder.Core.Manifests;

/// <summary>What to do about a manifest that already exists at the destination.</summary>
public enum ManifestWriteAction
{
    /// <summary>Nothing is there; write a new file.</summary>
    CreateNew = 0,

    /// <summary>Merge with the existing document and write conditionally on its ETag.</summary>
    MergeAndReplace = 1,

    /// <summary>Overwrite unconditionally.</summary>
    OverwriteWithoutMerge = 2,

    /// <summary>Write alongside, under a timestamped name.</summary>
    WriteTimestampedCopy = 3,

    /// <summary>Do not write at all.</summary>
    Skip = 4,

    /// <summary>Treat the existing file as an error.</summary>
    Fail = 5,
}

/// <summary>The decision, the name to write, and a user-facing explanation.</summary>
/// <param name="Action">What to do.</param>
/// <param name="FileName">The name to write to, which may be timestamped.</param>
/// <param name="ExistingDocument">The parsed existing document, when it could be parsed.</param>
/// <param name="IfMatchETag">The ETag for a conditional write, or null.</param>
/// <param name="Explanation">Why this action was chosen.</param>
public readonly record struct ManifestConflictDecision(
    ManifestWriteAction Action,
    string FileName,
    ManifestDocument? ExistingDocument,
    string? IfMatchETag,
    string Explanation);

/// <summary>
/// Decides how to handle an existing manifest.
/// <para>
/// The governing rule is that a file this application did not write is never overwritten by
/// default. "Update safely" degrades to writing a timestamped copy whenever the existing file
/// cannot be recognised and parsed, so a user's own document at that path survives.
/// </para>
/// </summary>
public sealed class ManifestConflictResolver
{
    private readonly IManifestParser _parser;
    private readonly TimeProvider _timeProvider;

    /// <summary>Creates a resolver.</summary>
    /// <param name="parser">Parser used to recognise this application's own output.</param>
    /// <param name="timeProvider">Clock, injected so timestamped names are testable.</param>
    public ManifestConflictResolver(IManifestParser parser, TimeProvider? timeProvider = null)
    {
        _parser = parser ?? throw new ArgumentNullException(nameof(parser));
        _timeProvider = timeProvider ?? TimeProvider.System;
    }

    /// <summary>Decides what to do about a destination that may already hold a manifest.</summary>
    /// <param name="fileName">The manifest name that would be written.</param>
    /// <param name="existing">The existing file, or null when the destination is empty.</param>
    /// <param name="policy">The configured conflict policy.</param>
    public ManifestConflictDecision Resolve(
        string fileName,
        ExistingManifest? existing,
        ManifestConflictPolicy policy)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(fileName);

        if (existing is null)
        {
            return new ManifestConflictDecision(
                ManifestWriteAction.CreateNew, fileName, null, null,
                "No manifest exists at this location; a new one will be created.");
        }

        var parsed = _parser.TryParse(existing.Content);

        return policy switch
        {
            ManifestConflictPolicy.Replace => new ManifestConflictDecision(
                ManifestWriteAction.OverwriteWithoutMerge, fileName, parsed.Document, existing.ETag,
                "The existing manifest will be replaced, as configured."),

            ManifestConflictPolicy.Skip => new ManifestConflictDecision(
                ManifestWriteAction.Skip, fileName, parsed.Document, existing.ETag,
                "A manifest already exists and the conflict policy is to skip it."),

            ManifestConflictPolicy.Fail => new ManifestConflictDecision(
                ManifestWriteAction.Fail, fileName, parsed.Document, existing.ETag,
                "A manifest already exists and the conflict policy is to fail."),

            ManifestConflictPolicy.CreateTimestampedVersion => new ManifestConflictDecision(
                ManifestWriteAction.WriteTimestampedCopy, BuildTimestampedName(fileName), parsed.Document, null,
                "A timestamped copy will be written and the existing file left untouched."),

            ManifestConflictPolicy.UpdateSafely when parsed.Succeeded => new ManifestConflictDecision(
                ManifestWriteAction.MergeAndReplace, fileName, parsed.Document, existing.ETag,
                "The existing manifest was written by this application and will be updated in place."),

            // The safety valve: an unparseable file is somebody else's, so it is preserved.
            ManifestConflictPolicy.UpdateSafely => new ManifestConflictDecision(
                ManifestWriteAction.WriteTimestampedCopy, BuildTimestampedName(fileName), null, null,
                $"The existing file could not be recognised as a manifest written by this application "
                + $"({parsed.FailureReason}), so it will be left untouched and a timestamped copy written."),

            _ => new ManifestConflictDecision(
                ManifestWriteAction.WriteTimestampedCopy, BuildTimestampedName(fileName), null, null,
                "Unrecognised conflict policy; defaulting to writing a timestamped copy."),
        };
    }

    /// <summary>
    /// Builds a timestamped variant such as <c>_sharepoint-links-20260902-141530.txt</c>. The
    /// timestamp is UTC and sorts lexicographically.
    /// </summary>
    public string BuildTimestampedName(string fileName)
    {
        var stem = Path.GetFileNameWithoutExtension(fileName);
        var extension = Path.GetExtension(fileName);
        var stamp = _timeProvider.GetUtcNow().UtcDateTime.ToString("yyyyMMdd-HHmmss", CultureInfo.InvariantCulture);

        return $"{stem}-{stamp}{extension}";
    }
}
