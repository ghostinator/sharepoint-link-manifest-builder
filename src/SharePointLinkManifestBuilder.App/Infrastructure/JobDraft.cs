using System.Collections.ObjectModel;
using SharePointLinkManifestBuilder.Core.Models;

namespace SharePointLinkManifestBuilder.App.Infrastructure;

/// <summary>
/// The job the user is currently assembling.
/// <para>
/// Shared between the SharePoint browser, the OneDrive browser and the job page so that
/// "Add as processing target" works from anywhere and the job page always reflects it. The
/// draft is mutable by design; it is converted into an immutable
/// <see cref="JobConfiguration"/> the moment a run starts, so nothing can change underneath a
/// job in flight.
/// </para>
/// </summary>
public sealed class JobDraft
{
    /// <summary>Locations the job will process.</summary>
    public ObservableCollection<ProcessingTarget> Targets { get; } = [];

    /// <summary>What to request from Microsoft 365.</summary>
    public LinkConfiguration Link { get; set; } = new();

    /// <summary>How manifests are produced and written.</summary>
    public ManifestConfiguration Manifest { get; set; } = new();

    /// <summary>Which discovered files are eligible.</summary>
    public FilterConfiguration Filters { get; set; } = new();

    /// <summary>Execution tuning.</summary>
    public ExecutionOptions Execution { get; set; } = new();

    /// <summary>True to enumerate and validate without changing anything.</summary>
    public bool DryRun { get; set; } = true;

    /// <summary>How overlapping targets are reconciled.</summary>
    public OverlapResolution OverlapResolution { get; set; } = OverlapResolution.KeepParent;

    /// <summary>Optional job name.</summary>
    public string? Name { get; set; }

    /// <summary>Raised whenever the target list changes, so pages can refresh their counts.</summary>
    public event EventHandler? TargetsChanged;

    /// <summary>
    /// Adds a target unless an identical location is already present. Silently ignoring an
    /// exact duplicate is friendlier than an error dialog, and overlap between *different*
    /// locations is handled properly at preflight rather than here.
    /// </summary>
    /// <returns>True when the target was added.</returns>
    public bool AddTarget(ProcessingTarget target)
    {
        ArgumentNullException.ThrowIfNull(target);

        var duplicate = Targets.Any(existing =>
            string.Equals(existing.DriveId, target.DriveId, StringComparison.OrdinalIgnoreCase)
            && string.Equals(existing.SiteId, target.SiteId, StringComparison.OrdinalIgnoreCase)
            && string.Equals(
                existing.StartingFolderRelativePath.Trim('/'),
                target.StartingFolderRelativePath.Trim('/'),
                StringComparison.OrdinalIgnoreCase)
            && existing.SourceType == target.SourceType);

        if (duplicate)
        {
            return false;
        }

        Targets.Add(target);
        TargetsChanged?.Invoke(this, EventArgs.Empty);
        return true;
    }

    /// <summary>Removes a target.</summary>
    public void RemoveTarget(ProcessingTarget target)
    {
        if (Targets.Remove(target))
        {
            TargetsChanged?.Invoke(this, EventArgs.Empty);
        }
    }

    /// <summary>Replaces a target with an edited copy, preserving its position in the list.</summary>
    public void ReplaceTarget(ProcessingTarget original, ProcessingTarget replacement)
    {
        var index = Targets.IndexOf(original);

        if (index < 0)
        {
            return;
        }

        Targets[index] = replacement;
        TargetsChanged?.Invoke(this, EventArgs.Empty);
    }

    /// <summary>Removes every target.</summary>
    public void Clear()
    {
        if (Targets.Count == 0)
        {
            return;
        }

        Targets.Clear();
        TargetsChanged?.Invoke(this, EventArgs.Empty);
    }

    /// <summary>
    /// Freezes the draft into an immutable configuration for a run. A fresh job ID is minted
    /// each time, so re-running never reuses the identifier of a previous run's manifests.
    /// </summary>
    public JobConfiguration ToConfiguration(string tenantId, string? tenantDisplayName) => new()
    {
        JobId = Guid.NewGuid().ToString("n"),
        Name = Name,
        TenantId = tenantId,
        TenantDisplayName = tenantDisplayName,
        Targets = Targets.ToArray(),
        Link = Link,
        Manifest = Manifest,
        Filters = Filters,
        Execution = Execution,
        DryRun = DryRun,
        OverlapResolution = OverlapResolution,
    };

    /// <summary>Replaces the whole draft from a saved profile or a previous run.</summary>
    public void LoadFrom(JobConfiguration configuration)
    {
        ArgumentNullException.ThrowIfNull(configuration);

        Targets.Clear();

        foreach (var target in configuration.Targets)
        {
            Targets.Add(target);
        }

        Name = configuration.Name;
        Link = configuration.Link;
        Manifest = configuration.Manifest;
        Filters = configuration.Filters;
        Execution = configuration.Execution;
        DryRun = configuration.DryRun;
        OverlapResolution = configuration.OverlapResolution;

        TargetsChanged?.Invoke(this, EventArgs.Empty);
    }
}
