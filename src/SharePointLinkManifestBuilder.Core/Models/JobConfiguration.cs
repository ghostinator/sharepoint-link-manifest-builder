namespace SharePointLinkManifestBuilder.Core.Models;

/// <summary>Behaviour when an individual file fails.</summary>
public enum FailureBehavior
{
    /// <summary>Record the failure and keep going. The default.</summary>
    ContinueOnError = 0,

    /// <summary>Stop the job at the first failure, preserving results already produced.</summary>
    StopOnFirstError = 1,
}

/// <summary>Execution tuning that does not change what a job means, only how it runs.</summary>
public sealed record ExecutionOptions
{
    /// <summary>
    /// Maximum simultaneous link requests. Deliberately modest: Graph throttling is a
    /// tenant-wide shared resource and this tool can generate a lot of writes.
    /// </summary>
    public int MaxConcurrency { get; init; } = 4;

    /// <summary>Optional pause between requests, for tenants that throttle aggressively.</summary>
    public TimeSpan RequestDelay { get; init; } = TimeSpan.Zero;

    /// <summary>Maximum retry attempts for a retryable failure.</summary>
    public int MaxRetryAttempts { get; init; } = 5;

    /// <summary>What to do when an individual file fails.</summary>
    public FailureBehavior FailureBehavior { get; init; } = FailureBehavior.ContinueOnError;

    /// <summary>Validates the options, returning one message per problem.</summary>
    public IReadOnlyList<string> Validate()
    {
        var problems = new List<string>();

        if (MaxConcurrency is < 1 or > 32)
        {
            problems.Add("Maximum concurrency must be between 1 and 32.");
        }

        if (RequestDelay < TimeSpan.Zero || RequestDelay > TimeSpan.FromSeconds(60))
        {
            problems.Add("The request delay must be between 0 and 60 seconds.");
        }

        if (MaxRetryAttempts is < 0 or > 10)
        {
            problems.Add("The retry limit must be between 0 and 10.");
        }

        return problems;
    }
}

/// <summary>
/// A complete, immutable description of a job. Once
/// <c>ILinkJobRunner.RunAsync</c> is entered this cannot change, which removes a class of bugs
/// where the user edits settings while a run is in flight. Editing in the UI produces a new
/// instance for the next run.
/// </summary>
public sealed record JobConfiguration
{
    /// <summary>Identifier recorded in manifests, history and reports.</summary>
    public string JobId { get; init; } = Guid.NewGuid().ToString("n");

    /// <summary>Optional user-supplied name.</summary>
    public string? Name { get; init; }

    /// <summary>Tenant this job runs against. Compared against the signed-in tenant at preflight.</summary>
    public required string TenantId { get; init; }

    /// <summary>Tenant display name for manifest headers, when known.</summary>
    public string? TenantDisplayName { get; init; }

    /// <summary>Locations to process. Disabled targets are retained but not run.</summary>
    public IReadOnlyList<ProcessingTarget> Targets { get; init; } = [];

    /// <summary>What to ask Microsoft 365 for.</summary>
    public LinkConfiguration Link { get; init; } = new();

    /// <summary>How manifests are produced and written.</summary>
    public ManifestConfiguration Manifest { get; init; } = new();

    /// <summary>Which discovered files are eligible.</summary>
    public FilterConfiguration Filters { get; init; } = new();

    /// <summary>Execution tuning.</summary>
    public ExecutionOptions Execution { get; init; } = new();

    /// <summary>
    /// When true the job enumerates, filters and validates but creates no link, writes no
    /// manifest, and changes no permission or tenant configuration.
    /// </summary>
    public bool DryRun { get; init; }

    /// <summary>How overlapping targets are reconciled.</summary>
    public OverlapResolution OverlapResolution { get; init; } = OverlapResolution.KeepParent;

    /// <summary>When the configuration was created.</summary>
    public DateTimeOffset CreatedUtc { get; init; } = DateTimeOffset.UtcNow;

    /// <summary>Targets that will actually run.</summary>
    public IReadOnlyList<ProcessingTarget> EnabledTargets =>
        Targets.Where(t => t.IsEnabled).ToArray();

    /// <summary>
    /// Validates the whole configuration without contacting Microsoft Graph. Network-dependent
    /// checks belong to preflight.
    /// </summary>
    public IReadOnlyList<string> Validate()
    {
        var problems = new List<string>();

        if (EnabledTargets.Count == 0)
        {
            problems.Add("No enabled processing targets are selected.");
        }

        foreach (var target in EnabledTargets.Where(t => !string.Equals(t.TenantId, TenantId, StringComparison.OrdinalIgnoreCase)))
        {
            problems.Add($"Target '{target.DisplayPath}' belongs to a different tenant than this job.");
        }

        problems.AddRange(Link.Validate());
        problems.AddRange(Manifest.Validate());
        problems.AddRange(Filters.Validate());
        problems.AddRange(Execution.Validate());

        return problems;
    }
}
