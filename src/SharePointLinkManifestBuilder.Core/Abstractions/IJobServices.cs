using SharePointLinkManifestBuilder.Core.Models;

namespace SharePointLinkManifestBuilder.Core.Abstractions;

/// <summary>Streamed, filtered, cancellable file discovery.</summary>
public interface IFileDiscoveryService
{
    /// <summary>
    /// Enumerates candidate files for one target, applying filters and honouring the target's
    /// recursion setting. Yields lazily; a large library is never materialized in full.
    /// </summary>
    /// <param name="target">The target to enumerate. Must be resolved to a concrete drive.</param>
    /// <param name="filters">Filters to apply.</param>
    /// <param name="manifestConfiguration">Used to exclude this application's own output.</param>
    /// <param name="cancellationToken">Cancellation token, honoured between pages and folders.</param>
    IAsyncEnumerable<DiscoveredFile> DiscoverAsync(
        ProcessingTarget target,
        FilterConfiguration filters,
        ManifestConfiguration manifestConfiguration,
        CancellationToken cancellationToken = default);
}

/// <summary>The result of preflight validation, run before any change is made.</summary>
public sealed record PreflightReport
{
    /// <summary>True when the job may proceed.</summary>
    public required bool CanProceed { get; init; }

    /// <summary>Conditions confirmed by an actual check.</summary>
    public IReadOnlyList<string> Validated { get; init; } = [];

    /// <summary>Conditions expected to hold but not directly confirmable.</summary>
    public IReadOnlyList<string> Expected { get; init; } = [];

    /// <summary>
    /// Conditions that cannot be known until execution. Tenant sharing policy in particular is
    /// only decided when a link is actually requested, and this is stated rather than glossed.
    /// </summary>
    public IReadOnlyList<string> UnknownUntilExecution { get; init; } = [];

    /// <summary>Problems that prevent the job from running.</summary>
    public IReadOnlyList<string> Blockers { get; init; } = [];

    /// <summary>Non-fatal concerns.</summary>
    public IReadOnlyList<string> Warnings { get; init; } = [];

    /// <summary>Targets after whole-site expansion and overlap resolution.</summary>
    public IReadOnlyList<ProcessingTarget> ResolvedTargets { get; init; } = [];

    /// <summary>Overlaps detected between selected targets.</summary>
    public IReadOnlyList<TargetOverlap> Overlaps { get; init; } = [];
}

/// <summary>The preview shown before a job runs.</summary>
public sealed record JobPreview
{
    /// <summary>Preflight findings.</summary>
    public required PreflightReport Preflight { get; init; }

    /// <summary>Files that would be processed.</summary>
    public IReadOnlyList<DiscoveredFile> Candidates { get; init; } = [];

    /// <summary>Items that would be skipped, with reasons.</summary>
    public IReadOnlyList<DiscoveredFile> Skipped { get; init; } = [];

    /// <summary>Where manifests would be written.</summary>
    public IReadOnlyList<string> ManifestDestinations { get; init; } = [];

    /// <summary>Manifests that already exist at those destinations.</summary>
    public IReadOnlyList<string> ExistingManifestConflicts { get; init; } = [];

    /// <summary>Candidate count.</summary>
    public int CandidateCount => Candidates.Count;

    /// <summary>Skipped count.</summary>
    public int SkippedCount => Skipped.Count;

    /// <summary>
    /// Estimated Graph write operations, or null when it cannot be calculated accurately.
    /// A guess is not offered in place of a real number.
    /// </summary>
    public int? EstimatedApiOperations { get; init; }

    /// <summary>True when this preview came from a dry run.</summary>
    public bool WasDryRun { get; init; }
}

/// <summary>Runs a link job through its phases.</summary>
public interface ILinkJobRunner
{
    /// <summary>
    /// Validates the configuration and resolves targets without changing anything. Safe to call
    /// repeatedly.
    /// </summary>
    Task<PreflightReport> PreflightAsync(
        JobConfiguration configuration,
        CancellationToken cancellationToken = default);

    /// <summary>
    /// Produces a preview by enumerating and filtering, without creating a link or writing a
    /// manifest. This is what dry-run mode returns.
    /// </summary>
    Task<JobPreview> PreviewAsync(
        JobConfiguration configuration,
        IProgress<JobProgress>? progress = null,
        CancellationToken cancellationToken = default);

    /// <summary>
    /// Executes the job. Requires a preview to have been produced and explicitly confirmed by
    /// the user; the runner never starts creating links on its own.
    /// </summary>
    /// <param name="configuration">The immutable job configuration.</param>
    /// <param name="preview">The confirmed preview.</param>
    /// <param name="progress">Progress sink.</param>
    /// <param name="pauseToken">Signals a user-requested pause.</param>
    /// <param name="cancellationToken">Cancellation token.</param>
    Task<JobSummary> RunAsync(
        JobConfiguration configuration,
        JobPreview preview,
        IProgress<JobProgress>? progress = null,
        IPauseToken? pauseToken = null,
        CancellationToken cancellationToken = default);

    /// <summary>
    /// Retries the retryable failures of a previous run, reusing the original configuration so
    /// the retry is exactly equivalent.
    /// </summary>
    Task<JobSummary> RetryFailuresAsync(
        JobConfiguration configuration,
        IReadOnlyList<LinkResult> failures,
        IProgress<JobProgress>? progress = null,
        CancellationToken cancellationToken = default);
}

/// <summary>A cooperative pause signal, separate from cancellation.</summary>
public interface IPauseToken
{
    /// <summary>True while the run is paused.</summary>
    bool IsPaused { get; }

    /// <summary>
    /// Completes when the run is not paused. Awaited at safe points, so pausing never
    /// interrupts an in-flight Graph write.
    /// </summary>
    Task WaitWhilePausedAsync(CancellationToken cancellationToken = default);
}
