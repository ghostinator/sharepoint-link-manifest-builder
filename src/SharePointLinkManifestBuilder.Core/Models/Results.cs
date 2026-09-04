namespace SharePointLinkManifestBuilder.Core.Models;

/// <summary>The outcome of requesting a sharing link for a single file.</summary>
public sealed record LinkResult
{
    /// <summary>The file the request was made for.</summary>
    public required DiscoveredFile File { get; init; }

    /// <summary>What actually happened. Never inferred; derived from the Graph response.</summary>
    public required LinkResultStatus Status { get; init; }

    /// <summary>The sharing URL, when one was produced.</summary>
    public string? SharingUrl { get; init; }

    /// <summary>The Graph permission ID of the sharing link, when returned.</summary>
    public string? PermissionId { get; init; }

    /// <summary>The link type Graph actually returned, which may differ from the request.</summary>
    public string? GrantedLinkType { get; init; }

    /// <summary>The link scope Graph actually returned, which may differ from the request.</summary>
    public string? GrantedScope { get; init; }

    /// <summary>Expiry Graph actually applied, which tenant policy may have adjusted.</summary>
    public DateTimeOffset? ExpirationUtc { get; init; }

    /// <summary>Per-recipient outcomes when the invite action was used.</summary>
    public IReadOnlyList<RecipientResult> RecipientResults { get; init; } = [];

    /// <summary>The failure, when the request did not succeed.</summary>
    public GraphError? Error { get; init; }

    /// <summary>When the result was produced.</summary>
    public DateTimeOffset TimestampUtc { get; init; } = DateTimeOffset.UtcNow;

    /// <summary>How many attempts were made, including the first.</summary>
    public int AttemptCount { get; init; } = 1;

    /// <summary>True when this file contributes an entry to a manifest.</summary>
    public bool IsSuccess =>
        Status is LinkResultStatus.Created or LinkResultStatus.Reused or LinkResultStatus.Existing;

    /// <summary>True when a retry could plausibly change the outcome.</summary>
    public bool IsRetryable =>
        Status is LinkResultStatus.Failed && (Error?.IsRetryable ?? false);

    /// <summary>The status word written into a plain-text manifest entry.</summary>
    public string ManifestStatus => Status switch
    {
        LinkResultStatus.Created => "Created",
        LinkResultStatus.Reused => "Reused",
        LinkResultStatus.Existing => "Existing",
        _ => Status.ToString(),
    };
}

/// <summary>
/// The outcome for one recipient of an invite action. Graph can return
/// <c>207 Multi-Status</c>, succeeding for some recipients and failing for others.
/// </summary>
public sealed record RecipientResult
{
    /// <summary>The recipient address as supplied.</summary>
    public required string Recipient { get; init; }

    /// <summary>True when Microsoft 365 granted this recipient access.</summary>
    public required bool Succeeded { get; init; }

    /// <summary>The failure for this recipient, when it did not succeed.</summary>
    public GraphError? Error { get; init; }
}

/// <summary>The phase a job is currently in.</summary>
public enum JobPhase
{
    /// <summary>Not started.</summary>
    NotStarted = 0,

    /// <summary>Checking the configuration without contacting Graph.</summary>
    ValidatingConfiguration,

    /// <summary>Resolving targets and expanding whole-site targets into libraries.</summary>
    ResolvingTargets,

    /// <summary>Enumerating candidate files.</summary>
    EnumeratingFiles,

    /// <summary>Awaiting the user's explicit Start after preview.</summary>
    AwaitingConfirmation,

    /// <summary>Requesting sharing links.</summary>
    CreatingLinks,

    /// <summary>Assembling manifest documents.</summary>
    BuildingManifests,

    /// <summary>Writing manifests back to SharePoint or OneDrive.</summary>
    UploadingManifests,

    /// <summary>Finished.</summary>
    Completed,

    /// <summary>Stopped by the user; results already produced are preserved.</summary>
    Cancelled,

    /// <summary>Stopped by an error that prevented continuing.</summary>
    Faulted,
}

/// <summary>A progress snapshot. Immutable so it can cross threads safely.</summary>
public sealed record JobProgress
{
    /// <summary>Current phase.</summary>
    public JobPhase Phase { get; init; } = JobPhase.NotStarted;

    /// <summary>Friendly name of the target being processed.</summary>
    public string? CurrentTargetName { get; init; }

    /// <summary>Index of the current target, one-based.</summary>
    public int CurrentTargetIndex { get; init; }

    /// <summary>Total targets to process.</summary>
    public int TotalTargets { get; init; }

    /// <summary>Name of the file being processed.</summary>
    public string? CurrentFileName { get; init; }

    /// <summary>Path of the file being processed, relative to its target.</summary>
    public string? CurrentRelativePath { get; init; }

    /// <summary>The Graph operation in flight, for example "createLink".</summary>
    public string? CurrentOperation { get; init; }

    /// <summary>Files processed so far.</summary>
    public int ProcessedCount { get; init; }

    /// <summary>Total candidate files, known after enumeration completes.</summary>
    public int TotalCount { get; init; }

    /// <summary>Links newly created.</summary>
    public int CreatedCount { get; init; }

    /// <summary>Existing equivalent links returned by Graph.</summary>
    public int ReusedCount { get; init; }

    /// <summary>Files skipped.</summary>
    public int SkippedCount { get; init; }

    /// <summary>Files that failed.</summary>
    public int FailedCount { get; init; }

    /// <summary>Retry attempts made.</summary>
    public int RetryCount { get; init; }

    /// <summary>True while the client is backing off because Graph asked it to.</summary>
    public bool IsThrottled { get; init; }

    /// <summary>When the client will resume after throttling.</summary>
    public DateTimeOffset? ThrottledUntilUtc { get; init; }

    /// <summary>True while the user has paused the run.</summary>
    public bool IsPaused { get; init; }

    /// <summary>Completion fraction from 0 to 1, or null while the total is unknown.</summary>
    public double? Fraction =>
        TotalCount > 0 ? Math.Clamp((double)ProcessedCount / TotalCount, 0, 1) : null;

    /// <summary>A screen-reader-friendly status sentence.</summary>
    public string AccessibleStatus =>
        Phase switch
        {
            JobPhase.EnumeratingFiles => $"Finding files. {ProcessedCount} found so far.",
            JobPhase.CreatingLinks =>
                $"Creating links. {ProcessedCount} of {TotalCount} processed, "
                + $"{CreatedCount} created, {ReusedCount} reused, {FailedCount} failed.",
            JobPhase.UploadingManifests => "Writing manifests.",
            JobPhase.Completed =>
                $"Finished. {CreatedCount} created, {ReusedCount} reused, "
                + $"{SkippedCount} skipped, {FailedCount} failed.",
            JobPhase.Cancelled => $"Cancelled. {ProcessedCount} files were processed before stopping.",
            _ => Phase.ToString(),
        };
}

/// <summary>Where a manifest ended up, for the results screen and job history.</summary>
public sealed record ManifestWriteResult
{
    /// <summary>Friendly destination path.</summary>
    public required string DisplayPath { get; init; }

    /// <summary>Absolute URL of the written manifest, when available.</summary>
    public string? WebUrl { get; init; }

    /// <summary>Format written.</summary>
    public required ManifestFormats Format { get; init; }

    /// <summary>True for a master manifest, false for a per-folder manifest.</summary>
    public bool IsMaster { get; init; }

    /// <summary>Entries written.</summary>
    public int EntryCount { get; init; }

    /// <summary>True when the write succeeded.</summary>
    public required bool Succeeded { get; init; }

    /// <summary>How an existing manifest was handled.</summary>
    public string? ConflictOutcome { get; init; }

    /// <summary>The failure, when the write did not succeed.</summary>
    public GraphError? Error { get; init; }
}

/// <summary>The final report for a run.</summary>
public sealed record JobSummary
{
    /// <summary>The job this summarizes.</summary>
    public required string JobId { get; init; }

    /// <summary>Application version that produced the run.</summary>
    public required string ApplicationVersion { get; init; }

    /// <summary>When the run started.</summary>
    public required DateTimeOffset StartedUtc { get; init; }

    /// <summary>When the run finished.</summary>
    public DateTimeOffset? CompletedUtc { get; init; }

    /// <summary>Terminal phase.</summary>
    public required JobPhase FinalPhase { get; init; }

    /// <summary>True when this was a dry run and nothing was modified.</summary>
    public bool WasDryRun { get; init; }

    /// <summary>True when the user stopped the run.</summary>
    public bool WasCancelled => FinalPhase == JobPhase.Cancelled;

    /// <summary>Every per-file outcome.</summary>
    public IReadOnlyList<LinkResult> Results { get; init; } = [];

    /// <summary>Every manifest write attempt.</summary>
    public IReadOnlyList<ManifestWriteResult> Manifests { get; init; } = [];

    /// <summary>Failures that were not tied to a single file.</summary>
    public IReadOnlyList<GraphError> JobLevelErrors { get; init; } = [];

    /// <summary>Warnings raised during preflight or execution.</summary>
    public IReadOnlyList<string> Warnings { get; init; } = [];

    /// <summary>Links newly created.</summary>
    public int CreatedCount => Results.Count(r => r.Status == LinkResultStatus.Created);

    /// <summary>Existing equivalent links returned by Graph.</summary>
    public int ReusedCount =>
        Results.Count(r => r.Status is LinkResultStatus.Reused or LinkResultStatus.Existing);

    /// <summary>Files skipped.</summary>
    public int SkippedCount => Results.Count(r => r.Status == LinkResultStatus.Skipped);

    /// <summary>Files that failed for any reason, including policy and access decisions.</summary>
    public int FailedCount => Results.Count(r =>
        r.Status is LinkResultStatus.Failed or LinkResultStatus.Unsupported
            or LinkResultStatus.PolicyBlocked or LinkResultStatus.AccessDenied);

    /// <summary>Files that produced a manifest entry.</summary>
    public int SuccessCount => Results.Count(r => r.IsSuccess);

    /// <summary>Failures a retry could plausibly fix.</summary>
    public IReadOnlyList<LinkResult> RetryableFailures =>
        Results.Where(r => r.IsRetryable).ToArray();

    /// <summary>How long the run took.</summary>
    public TimeSpan? Duration => CompletedUtc - StartedUtc;
}
