using SharePointLinkManifestBuilder.Core.Models;

namespace SharePointLinkManifestBuilder.Core.Filtering;

/// <summary>
/// Applies the default exclusions and the user's filter rules to a discovered item, returning
/// the reason it would be skipped.
/// <para>
/// Pure and deterministic, so every rule in the product is unit-testable without a network.
/// </para>
/// </summary>
public sealed class FileFilterEvaluator
{
    private readonly FilterConfiguration _filters;
    private readonly HashSet<string> _generatedManifestNames;

    /// <summary>Creates an evaluator for one job's filter and manifest configuration.</summary>
    /// <param name="filters">The user's filter rules.</param>
    /// <param name="manifestConfiguration">
    /// Used to exclude the names this job will itself write, so a job never discovers and
    /// re-processes its own output.
    /// </param>
    public FileFilterEvaluator(FilterConfiguration filters, ManifestConfiguration manifestConfiguration)
    {
        ArgumentNullException.ThrowIfNull(filters);
        ArgumentNullException.ThrowIfNull(manifestConfiguration);

        _filters = filters;
        _generatedManifestNames = new HashSet<string>(
            manifestConfiguration.GeneratedFileNames(),
            StringComparer.OrdinalIgnoreCase);
    }

    /// <summary>
    /// Evaluates one item. Returns <see cref="SkipReason.None"/> when the item is eligible.
    /// </summary>
    /// <param name="file">The discovered item.</param>
    /// <param name="isHiddenOrSystem">Whether Graph flagged the item as hidden or a system artifact.</param>
    public SkipReason Evaluate(DiscoveredFile file, bool isHiddenOrSystem = false)
    {
        ArgumentNullException.ThrowIfNull(file);

        // Structural exclusions first: these are about what the item *is*, not what the user asked for.
        if (file.Kind == DriveItemKind.Folder)
        {
            return SkipReason.IsFolder;
        }

        if (file.Kind == DriveItemKind.Package && !_filters.IncludeSpecialItemTypes)
        {
            return SkipReason.PackageItem;
        }

        if (file.Kind == DriveItemKind.RemoteItem && !_filters.IncludeSpecialItemTypes)
        {
            return SkipReason.RemoteItem;
        }

        if (file.Kind == DriveItemKind.Unsupported)
        {
            return SkipReason.UnsupportedItemType;
        }

        if (IsTemporaryFile(file.Name) && !_filters.IncludeTemporaryFiles)
        {
            return SkipReason.TemporaryFile;
        }

        if (!_filters.IncludeGeneratedManifests && IsGeneratedManifest(file.Name))
        {
            return SkipReason.GeneratedManifest;
        }

        if (isHiddenOrSystem && !_filters.IncludeHiddenAndSystemItems)
        {
            return SkipReason.HiddenOrSystem;
        }

        // User-configured rules.
        return MatchesUserFilters(file) ? SkipReason.None : SkipReason.FilteredOut;
    }

    /// <summary>
    /// True when a name is an Office lock or temporary artifact. These are never useful in a
    /// manifest and often disappear before a link can be used.
    /// </summary>
    public static bool IsTemporaryFile(string? name)
    {
        if (string.IsNullOrEmpty(name))
        {
            return false;
        }

        return name.StartsWith("~$", StringComparison.Ordinal)
            || name.StartsWith(".~", StringComparison.Ordinal)
            || name.EndsWith(".tmp", StringComparison.OrdinalIgnoreCase)
            || name.EndsWith('~')
            || string.Equals(name, "thumbs.db", StringComparison.OrdinalIgnoreCase)
            || string.Equals(name, ".ds_store", StringComparison.OrdinalIgnoreCase);
    }

    /// <summary>True when the name is one this application generates.</summary>
    public bool IsGeneratedManifest(string? name) =>
        !string.IsNullOrEmpty(name)
        && (_generatedManifestNames.Contains(name) || ManifestDefaults.IsGeneratedManifestName(name));

    private bool MatchesUserFilters(DiscoveredFile file)
    {
        if (_filters.IncludeExtensions.Count > 0
            && !HasExtension(file.Name, _filters.IncludeExtensions))
        {
            return false;
        }

        if (_filters.ExcludeExtensions.Count > 0
            && HasExtension(file.Name, _filters.ExcludeExtensions))
        {
            return false;
        }

        if (_filters.IncludePatterns.Count > 0
            && !GlobMatcher.IsMatchAny(_filters.IncludePatterns, file.Name))
        {
            return false;
        }

        if (_filters.ExcludePatterns.Count > 0
            && GlobMatcher.IsMatchAny(_filters.ExcludePatterns, file.Name))
        {
            return false;
        }

        if (_filters.ModifiedAfterUtc is { } after
            && (file.LastModifiedUtc is null || file.LastModifiedUtc < after))
        {
            return false;
        }

        if (_filters.ModifiedBeforeUtc is { } before
            && (file.LastModifiedUtc is null || file.LastModifiedUtc > before))
        {
            return false;
        }

        if (_filters.MinimumSizeBytes is { } min && (file.Size is null || file.Size < min))
        {
            return false;
        }

        if (_filters.MaximumSizeBytes is { } max && (file.Size is null || file.Size > max))
        {
            return false;
        }

        return true;
    }

    /// <summary>
    /// Compares a name's extension against a list, tolerating entries written with or without
    /// a leading dot, because users type both.
    /// </summary>
    private static bool HasExtension(string name, IReadOnlyList<string> extensions)
    {
        var actual = Path.GetExtension(name);
        if (string.IsNullOrEmpty(actual))
        {
            return false;
        }

        foreach (var candidate in extensions)
        {
            if (string.IsNullOrWhiteSpace(candidate))
            {
                continue;
            }

            var normalized = candidate.StartsWith('.') ? candidate : "." + candidate;
            if (string.Equals(actual, normalized, StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }
        }

        return false;
    }
}
