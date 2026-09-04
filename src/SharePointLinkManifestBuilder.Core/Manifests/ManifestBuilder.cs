using SharePointLinkManifestBuilder.Core.Abstractions;
using SharePointLinkManifestBuilder.Core.Models;
using SharePointLinkManifestBuilder.Core.Targets;

namespace SharePointLinkManifestBuilder.Core.Manifests;

/// <summary>Assembles manifest documents from a run's per-file results.</summary>
public sealed class ManifestBuilder : IManifestBuilder
{
    /// <inheritdoc />
    public IReadOnlyList<(string FolderItemId, string FolderPath, ManifestDocument Document)>
        BuildPerFolderManifests(
            JobConfiguration configuration,
            ProcessingTarget target,
            IReadOnlyList<LinkResult> results,
            string applicationVersion)
    {
        ArgumentNullException.ThrowIfNull(configuration);
        ArgumentNullException.ThrowIfNull(target);
        ArgumentNullException.ThrowIfNull(results);

        var successes = results.Where(r => r.IsSuccess).ToArray();
        var output = new List<(string, string, ManifestDocument)>();

        // Group by the containing folder. A file whose parent folder is unknown is attributed
        // to the target's starting folder rather than being dropped.
        var groups = successes
            .GroupBy(r => (
                Id: r.File.ParentFolderItemId ?? target.StartingFolderItemId ?? string.Empty,
                Path: TargetPlanner.NormalizePath(r.File.ParentRelativePath)))
            .ToArray();

        foreach (var group in groups)
        {
            var entries = configuration.Manifest.AggregateDescendantsInPerFolderManifest
                ? successes.Where(r => IsAtOrUnder(r.File.ParentRelativePath, group.Key.Path)).ToArray()
                : group.ToArray();

            if (entries.Length == 0)
            {
                continue;
            }

            var document = new ManifestDocument
            {
                Header = BuildHeader(
                    configuration,
                    target,
                    results,
                    entries,
                    applicationVersion,
                    startingFolderOverride: FolderDisplayPath(target, group.Key.Path)),
                Entries = entries
                    .Select(r => ManifestEntry.FromResult(r) with
                    {
                        // Inside a per-folder manifest, paths read best relative to that folder.
                        RelativePath = configuration.Manifest.AggregateDescendantsInPerFolderManifest
                            ? MakeRelativeTo(r.File.RelativePath, group.Key.Path)
                            : r.File.Name,
                    })
                    .ToArray(),
            };

            output.Add((group.Key.Id, group.Key.Path, document));
        }

        return output;
    }

    /// <inheritdoc />
    public ManifestDocument BuildMasterManifest(
        JobConfiguration configuration,
        ProcessingTarget target,
        IReadOnlyList<LinkResult> results,
        string applicationVersion)
    {
        ArgumentNullException.ThrowIfNull(configuration);
        ArgumentNullException.ThrowIfNull(target);
        ArgumentNullException.ThrowIfNull(results);

        var successes = results.Where(r => r.IsSuccess).ToArray();

        return new ManifestDocument
        {
            Header = BuildHeader(configuration, target, results, successes, applicationVersion),
            Entries = successes.Select(ManifestEntry.FromResult).ToArray(),
        };
    }

    /// <inheritdoc />
    public ManifestDocument BuildCombinedMasterManifest(
        JobConfiguration configuration,
        IReadOnlyList<LinkResult> results,
        string applicationVersion)
    {
        ArgumentNullException.ThrowIfNull(configuration);
        ArgumentNullException.ThrowIfNull(results);

        var successes = results.Where(r => r.IsSuccess).ToArray();
        var targets = configuration.EnabledTargets;

        // A combined manifest spans several sources, so naming a single site or library would
        // be misleading. It reports the mix instead.
        var sourceType = targets.Select(t => t.SourceTypeLabel).Distinct(StringComparer.Ordinal).ToArray();

        return new ManifestDocument
        {
            Header = new ManifestHeader
            {
                ApplicationVersion = applicationVersion,
                JobId = configuration.JobId,
                TenantDisplayName = configuration.TenantDisplayName ?? configuration.TenantId,
                TenantId = configuration.TenantId,
                SourceType = sourceType.Length == 1
                    ? sourceType[0]
                    : $"Multiple ({string.Join(", ", sourceType)})",
                LibraryOrDriveName = targets.Count == 1
                    ? targets[0].DriveName
                    : $"{targets.Count} selected locations",
                StartingFolder = targets.Count == 1 ? FolderDisplayPath(targets[0], null) : "(multiple)",
                Recursive = targets.Any(t => t.Recursive),
                LinkPermission = configuration.Link.Permission.ToString(),
                LinkAudience = DescribeAudience(configuration.Link.Audience),
                SuccessfulFiles = successes.Length,
                ReusedLinks = results.Count(r =>
                    r.Status is LinkResultStatus.Reused or LinkResultStatus.Existing),
                SkippedFiles = results.Count(r => r.Status == LinkResultStatus.Skipped),
                FailedFiles = results.Count(r =>
                    r.Status is LinkResultStatus.Failed or LinkResultStatus.Unsupported
                        or LinkResultStatus.PolicyBlocked or LinkResultStatus.AccessDenied),
            },
            Entries = successes.Select(ManifestEntry.FromResult).ToArray(),
        };
    }

    /// <summary>Renders the audience as it appears in a manifest header.</summary>
    public static string DescribeAudience(LinkAudience audience) => audience switch
    {
        LinkAudience.Organization => "Organization",
        LinkAudience.SpecificPeople => "Specific People",
        LinkAudience.Anyone => "Anyone",
        _ => audience.ToString(),
    };

    private static ManifestHeader BuildHeader(
        JobConfiguration configuration,
        ProcessingTarget target,
        IReadOnlyList<LinkResult> allResults,
        IReadOnlyList<LinkResult> includedResults,
        string applicationVersion,
        string? startingFolderOverride = null) => new()
        {
            ApplicationVersion = applicationVersion,
            JobId = configuration.JobId,
            TenantDisplayName = configuration.TenantDisplayName ?? configuration.TenantId,
            TenantId = configuration.TenantId,
            SourceType = target.SourceTypeLabel,
            SiteName = target.SiteName,
            SiteUrl = target.SiteUrl,
            UserDisplayName = target.UserDisplayName,
            LibraryOrDriveName = target.DriveName,
            StartingFolder = startingFolderOverride ?? FolderDisplayPath(target, null),
            Recursive = target.Recursive,
            LinkPermission = configuration.Link.Permission.ToString(),
            LinkAudience = DescribeAudience(configuration.Link.Audience),
            SuccessfulFiles = includedResults.Count(r => r.IsSuccess),
            ReusedLinks = includedResults.Count(r =>
                r.Status is LinkResultStatus.Reused or LinkResultStatus.Existing),
            SkippedFiles = allResults.Count(r => r.Status == LinkResultStatus.Skipped),
            FailedFiles = allResults.Count(r =>
                r.Status is LinkResultStatus.Failed or LinkResultStatus.Unsupported
                    or LinkResultStatus.PolicyBlocked or LinkResultStatus.AccessDenied),
        };

    private static string FolderDisplayPath(ProcessingTarget target, string? relativePath)
    {
        var basePath = TargetPlanner.NormalizePath(target.StartingFolderRelativePath);
        var extra = TargetPlanner.NormalizePath(relativePath);

        var combined = (basePath, extra) switch
        {
            ("", "") => string.Empty,
            ("", _) => extra,
            (_, "") => basePath,
            _ => $"{basePath}/{extra}",
        };

        return combined.Length == 0 ? "/" : "/" + combined;
    }

    private static bool IsAtOrUnder(string candidate, string ancestor)
    {
        candidate = TargetPlanner.NormalizePath(candidate);
        ancestor = TargetPlanner.NormalizePath(ancestor);

        return string.Equals(candidate, ancestor, StringComparison.OrdinalIgnoreCase)
            || TargetPlanner.IsUnder(candidate, ancestor);
    }

    private static string MakeRelativeTo(string path, string ancestor)
    {
        path = TargetPlanner.NormalizePath(path);
        ancestor = TargetPlanner.NormalizePath(ancestor);

        if (ancestor.Length == 0)
        {
            return path;
        }

        return path.StartsWith(ancestor + "/", StringComparison.OrdinalIgnoreCase)
            ? path[(ancestor.Length + 1)..]
            : path;
    }
}
