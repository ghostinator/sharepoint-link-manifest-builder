using SharePointLinkManifestBuilder.Core.Models;

namespace SharePointLinkManifestBuilder.Core.Targets;

/// <summary>The outcome of reconciling a job's selected targets.</summary>
public sealed record TargetPlan
{
    /// <summary>The targets that will actually be enumerated.</summary>
    public IReadOnlyList<ProcessingTarget> EffectiveTargets { get; init; } = [];

    /// <summary>Every overlap detected among the selected targets.</summary>
    public IReadOnlyList<TargetOverlap> Overlaps { get; init; } = [];

    /// <summary>Targets removed by the chosen resolution, with the reason.</summary>
    public IReadOnlyList<(ProcessingTarget Target, string Reason)> Removed { get; init; } = [];

    /// <summary>Messages describing the reconciliation, shown before the job runs.</summary>
    public IReadOnlyList<string> Warnings { get; init; } = [];

    /// <summary>True when any overlap was found.</summary>
    public bool HasOverlaps => Overlaps.Count > 0;
}

/// <summary>
/// Detects overlapping selections and reduces them to a set of targets that processes each
/// file once.
/// <para>
/// Overlap is not merely "one path is a prefix of another". Whether a broader target actually
/// covers a narrower one depends on the broader target's recursion setting: a non-recursive
/// folder target processes only its direct children, so a target on a subfolder does not
/// overlap with it at all. Treating those as overlapping would silently drop files the user
/// asked for.
/// </para>
/// </summary>
public static class TargetPlanner
{
    /// <summary>Reconciles a set of targets under the chosen resolution strategy.</summary>
    /// <param name="targets">The targets the user selected.</param>
    /// <param name="resolution">How to reconcile detected overlaps.</param>
    public static TargetPlan Plan(
        IReadOnlyList<ProcessingTarget> targets,
        OverlapResolution resolution = OverlapResolution.KeepParent)
    {
        ArgumentNullException.ThrowIfNull(targets);

        var enabled = targets.Where(t => t.IsEnabled).ToList();
        var overlaps = DetectOverlaps(enabled);

        if (overlaps.Count == 0)
        {
            return new TargetPlan { EffectiveTargets = enabled };
        }

        var warnings = new List<string>();
        var removed = new List<(ProcessingTarget, string)>();
        var dropped = new HashSet<string>(StringComparer.Ordinal);

        switch (resolution)
        {
            case OverlapResolution.KeepParent:
                foreach (var overlap in overlaps)
                {
                    if (dropped.Add(overlap.Child.TargetId))
                    {
                        removed.Add((overlap.Child,
                            $"Already covered by the broader target '{overlap.Parent.DisplayPath}'."));
                    }
                }

                warnings.Add(
                    $"{overlaps.Count} overlapping selection(s) were reduced by keeping the broader target. "
                    + "Each file is processed once.");
                break;

            case OverlapResolution.KeepChild:
                foreach (var overlap in overlaps)
                {
                    if (dropped.Add(overlap.Parent.TargetId))
                    {
                        removed.Add((overlap.Parent,
                            $"Removed in favour of the narrower target '{overlap.Child.DisplayPath}'."));
                    }
                }

                warnings.Add(
                    "Broader targets were removed in favour of narrower ones. Files that were only inside "
                    + "the broader targets will not be processed.");
                break;

            case OverlapResolution.KeepBothDeduplicate:
                warnings.Add(
                    $"{overlaps.Count} overlapping selection(s) were kept. Files reachable through more than "
                    + "one target are processed once, matched on drive and item ID.");
                break;

            default:
                break;
        }

        // An exact duplicate is always collapsed: keeping both would be meaningless.
        foreach (var duplicate in overlaps.Where(o => o.IsExactDuplicate))
        {
            if (!dropped.Contains(duplicate.Parent.TargetId) && dropped.Add(duplicate.Child.TargetId))
            {
                removed.Add((duplicate.Child, "The same location was selected more than once."));
            }
        }

        return new TargetPlan
        {
            EffectiveTargets = enabled.Where(t => !dropped.Contains(t.TargetId)).ToArray(),
            Overlaps = overlaps,
            Removed = removed,
            Warnings = warnings,
        };
    }

    /// <summary>Finds every pair where one target's scope covers another's.</summary>
    public static IReadOnlyList<TargetOverlap> DetectOverlaps(IReadOnlyList<ProcessingTarget> targets)
    {
        ArgumentNullException.ThrowIfNull(targets);

        var overlaps = new List<TargetOverlap>();

        for (var i = 0; i < targets.Count; i++)
        {
            for (var j = 0; j < targets.Count; j++)
            {
                if (i == j)
                {
                    continue;
                }

                var parent = targets[i];
                var child = targets[j];

                if (!Covers(parent, child, out var explanation, out var isExactDuplicate))
                {
                    continue;
                }

                // An exact duplicate satisfies Covers in both directions; record it once.
                if (isExactDuplicate && j < i)
                {
                    continue;
                }

                overlaps.Add(new TargetOverlap
                {
                    Parent = parent,
                    Child = child,
                    Explanation = explanation,
                    IsExactDuplicate = isExactDuplicate,
                });
            }
        }

        return overlaps;
    }

    /// <summary>
    /// True when <paramref name="parent"/>'s scope wholly contains <paramref name="child"/>'s,
    /// taking recursion into account.
    /// </summary>
    /// <param name="parent">The candidate broader target.</param>
    /// <param name="child">The candidate narrower target.</param>
    /// <param name="explanation">Why they overlap, in plain language.</param>
    /// <param name="isExactDuplicate">True when both name the same location.</param>
    public static bool Covers(
        ProcessingTarget parent,
        ProcessingTarget child,
        out string explanation,
        out bool isExactDuplicate)
    {
        ArgumentNullException.ThrowIfNull(parent);
        ArgumentNullException.ThrowIfNull(child);

        explanation = string.Empty;
        isExactDuplicate = false;

        if (string.Equals(parent.TargetId, child.TargetId, StringComparison.Ordinal))
        {
            return false;
        }

        if (!string.Equals(parent.TenantId, child.TenantId, StringComparison.OrdinalIgnoreCase))
        {
            return false;
        }

        // A whole-site target is the root of every library in that site.
        if (parent.SourceType == TargetSourceType.SharePointSite)
        {
            if (!string.Equals(parent.SiteId, child.SiteId, StringComparison.OrdinalIgnoreCase))
            {
                return false;
            }

            if (child.SourceType == TargetSourceType.SharePointSite)
            {
                isExactDuplicate = true;
                explanation = $"The site '{parent.SiteName}' is selected more than once.";
                return true;
            }

            // A library root is always covered by its site, whatever the recursion setting,
            // because both begin at the same place.
            if (IsDriveRoot(child))
            {
                explanation =
                    $"The library '{child.DriveName}' is inside the selected site '{parent.SiteName}'.";
                return true;
            }

            // A folder inside a library is only reached when the site target recurses.
            if (parent.Recursive)
            {
                explanation =
                    $"The folder '{child.DisplayPath}' is inside the selected site '{parent.SiteName}', "
                    + "which is set to include subfolders.";
                return true;
            }

            return false;
        }

        // Everything else compares within a single drive.
        if (string.IsNullOrEmpty(parent.DriveId) || string.IsNullOrEmpty(child.DriveId))
        {
            return false;
        }

        if (!string.Equals(parent.DriveId, child.DriveId, StringComparison.OrdinalIgnoreCase))
        {
            return false;
        }

        var parentPath = NormalizePath(parent.StartingFolderRelativePath);
        var childPath = NormalizePath(child.StartingFolderRelativePath);

        if (string.Equals(parentPath, childPath, StringComparison.OrdinalIgnoreCase))
        {
            isExactDuplicate = true;
            explanation = $"'{parent.DisplayPath}' is selected more than once.";
            return true;
        }

        if (!IsUnder(childPath, parentPath))
        {
            return false;
        }

        // The decisive rule: a non-recursive parent never reaches into a subfolder, so the two
        // targets do not actually process the same files.
        if (!parent.Recursive)
        {
            return false;
        }

        explanation =
            $"'{child.DisplayPath}' is inside '{parent.DisplayPath}', which is set to include subfolders.";
        return true;
    }

    /// <summary>True when the target starts at the root of its drive.</summary>
    public static bool IsDriveRoot(ProcessingTarget target) =>
        string.IsNullOrEmpty(NormalizePath(target.StartingFolderRelativePath));

    /// <summary>True when <paramref name="candidate"/> lies strictly beneath <paramref name="ancestor"/>.</summary>
    public static bool IsUnder(string candidate, string ancestor)
    {
        candidate = NormalizePath(candidate);
        ancestor = NormalizePath(ancestor);

        if (ancestor.Length == 0)
        {
            return candidate.Length > 0;
        }

        // The trailing separator matters: "Reports" must not be treated as an ancestor of
        // "ReportsArchive".
        return candidate.StartsWith(ancestor + "/", StringComparison.OrdinalIgnoreCase);
    }

    /// <summary>Normalizes a relative path to forward slashes with no leading or trailing slash.</summary>
    public static string NormalizePath(string? path) =>
        string.IsNullOrWhiteSpace(path)
            ? string.Empty
            : path.Replace('\\', '/').Trim('/');
}
