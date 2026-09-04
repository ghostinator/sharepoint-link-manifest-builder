using System.Collections.ObjectModel;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using Microsoft.Extensions.Logging;
using SharePointLinkManifestBuilder.App.Infrastructure;
using SharePointLinkManifestBuilder.Core.Abstractions;
using SharePointLinkManifestBuilder.Core.Jobs;
using SharePointLinkManifestBuilder.Core.Models;

namespace SharePointLinkManifestBuilder.App.ViewModels;

/// <summary>
/// Assembles, previews and runs a link job.
/// <para>
/// The flow is deliberately gated: targets and settings, then a preview that must be produced
/// and looked at, and only then a Start button. Creating sharing links in bulk changes real
/// permissions on real content, so it should never be one click away from a half-configured
/// screen.
/// </para>
/// </summary>
public sealed partial class NewLinkJobViewModel : PageViewModelBase, IDisposable
{
    private readonly ILinkJobRunner _runner;
    private readonly ConnectionCoordinator _connection;
    private readonly JobDraft _draft;
    private readonly IJobHistoryStore _historyStore;
    private readonly ISettingsStore _settingsStore;
    private readonly IProductMetadataProvider _productMetadata;
    private readonly IClipboardService _clipboard;
    private readonly ISystemBrowser _browser;
    private readonly ILogger<NewLinkJobViewModel> _logger;

    private CancellationTokenSource? _runCancellation;
    private PauseToken? _pauseToken;
    private JobPreview? _preview;

    /// <summary>How many numbered steps the job page has.</summary>
    public const int StepCount = 6;

    /// <summary>
    /// Tab indices. The two browse tabs sit between Targets and Link so that choosing locations
    /// reads as part of step 1 rather than as leaving the job, and they are hidden until asked
    /// for. Indices are positional in the TabControl, so hidden tabs still occupy one -- which is
    /// what keeps these constants stable however the visibility changes.
    /// </summary>
    private const int TargetsTab = 0;
    private const int BrowseSharePointTab = 1;
    private const int BrowseOneDriveTab = 2;
    private const int FirstNumberedTabAfterBrowsing = 3;
    private const int LastTab = 7;

    /// <summary>The SharePoint browser, hosted inside this page rather than beside it.</summary>
    public SharePointBrowserViewModel SharePointBrowser { get; }

    /// <summary>The OneDrive browser, hosted inside this page rather than beside it.</summary>
    public OneDriveBrowserViewModel OneDriveBrowser { get; }

    /// <summary>The tab currently shown: 0 targets, 1 link, 2 manifest, 3 filters, 4 preview, 5 results.</summary>
    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(StepPosition))]
    [NotifyCanExecuteChangedFor(nameof(PreviousStepCommand))]
    [NotifyCanExecuteChangedFor(nameof(NextStepCommand))]
    private int _selectedTabIndex;

    /// <summary>A human-readable position, for the label between the Back and Next buttons.</summary>
    public string StepPosition => SelectedTabIndex switch
    {
        BrowseSharePointTab => "Choosing SharePoint locations",
        BrowseOneDriveTab => "Choosing OneDrive locations",
        TargetsTab => $"Step 1 of {StepCount}",
        _ => $"Step {SelectedTabIndex - FirstNumberedTabAfterBrowsing + 2} of {StepCount}",
    };

    /// <summary>True while the SharePoint browse tab should be shown.</summary>
    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(StepPosition))]
    [NotifyCanExecuteChangedFor(nameof(PreviousStepCommand))]
    [NotifyCanExecuteChangedFor(nameof(NextStepCommand))]
    private bool _isBrowsingSharePoint;

    /// <summary>True while the OneDrive browse tab should be shown.</summary>
    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(StepPosition))]
    [NotifyCanExecuteChangedFor(nameof(PreviousStepCommand))]
    [NotifyCanExecuteChangedFor(nameof(NextStepCommand))]
    private bool _isBrowsingOneDrive;

    /// <summary>True when either browse tab is showing, so the page can offer a way back.</summary>
    public bool IsBrowsing => IsBrowsingSharePoint || IsBrowsingOneDrive;

    /// <summary>Requested link permission.</summary>
    [ObservableProperty]
    private LinkPermission _linkPermission = LinkPermission.View;

    /// <summary>Requested link audience.</summary>
    [ObservableProperty]
    private LinkAudience _linkAudience = LinkAudience.Organization;

    /// <summary>Recipients for a specific-people link, one per line.</summary>
    [ObservableProperty]
    private string _recipientsText = string.Empty;

    /// <summary>True to send Microsoft's invitation email. Off by default.</summary>
    [ObservableProperty]
    private bool _sendInvitationEmail;

    /// <summary>Optional message included with an invitation.</summary>
    [ObservableProperty]
    private string _invitationMessage = string.Empty;

    /// <summary>True to request an expiry.</summary>
    [ObservableProperty]
    private bool _useExpiration;

    /// <summary>The requested expiry date.</summary>
    [ObservableProperty]
    private DateTimeOffset? _expirationDate = DateTimeOffset.UtcNow.AddDays(30);

    /// <summary>True to reuse an equivalent existing link rather than requesting a new one.</summary>
    [ObservableProperty]
    private bool _reuseExistingLinks = true;

    /// <summary>True to skip a file entirely when an equivalent link already exists.</summary>
    [ObservableProperty]
    private bool _skipWhenEquivalentLinkExists;

    /// <summary>True to retain inherited permissions on first share.</summary>
    [ObservableProperty]
    private bool _retainInheritedPermissions = true;

    /// <summary>True to write a manifest into each folder containing processed files.</summary>
    [ObservableProperty]
    private bool _writePerFolderManifest = true;

    /// <summary>True to write a manifest covering everything beneath each target.</summary>
    [ObservableProperty]
    private bool _writeMasterManifest;

    /// <summary>True to emit plain text. On by default.</summary>
    [ObservableProperty]
    private bool _formatPlainText = true;

    /// <summary>True to emit Markdown.</summary>
    [ObservableProperty]
    private bool _formatMarkdown;

    /// <summary>True to emit CSV.</summary>
    [ObservableProperty]
    private bool _formatCsv;

    /// <summary>True to emit JSON.</summary>
    [ObservableProperty]
    private bool _formatJson;

    /// <summary>True to list descendant files inside each per-folder manifest.</summary>
    [ObservableProperty]
    private bool _aggregateDescendants;

    /// <summary>How an existing manifest is handled.</summary>
    [ObservableProperty]
    private ManifestConflictPolicy _conflictPolicy = ManifestConflictPolicy.UpdateSafely;

    /// <summary>What happens to entries whose files were not seen on this run.</summary>
    [ObservableProperty]
    private MissingEntryPolicy _missingEntryPolicy = MissingEntryPolicy.Preserve;

    /// <summary>Comma-separated extensions to include.</summary>
    [ObservableProperty]
    private string _includeExtensions = string.Empty;

    /// <summary>Comma-separated extensions to exclude.</summary>
    [ObservableProperty]
    private string _excludeExtensions = string.Empty;

    /// <summary>Comma-separated glob patterns a name must match.</summary>
    [ObservableProperty]
    private string _includePatterns = string.Empty;

    /// <summary>Comma-separated glob patterns that exclude a name.</summary>
    [ObservableProperty]
    private string _excludePatterns = string.Empty;

    /// <summary>Only include files modified on or after this date.</summary>
    [ObservableProperty]
    private DateTimeOffset? _modifiedAfter;

    /// <summary>Only include files modified on or before this date.</summary>
    [ObservableProperty]
    private DateTimeOffset? _modifiedBefore;

    /// <summary>Minimum size in kilobytes.</summary>
    [ObservableProperty]
    private double? _minimumSizeKilobytes;

    /// <summary>Maximum size in kilobytes.</summary>
    [ObservableProperty]
    private double? _maximumSizeKilobytes;

    /// <summary>True to include hidden and system items.</summary>
    [ObservableProperty]
    private bool _includeHiddenItems;

    /// <summary>True to include Office lock and temporary files.</summary>
    [ObservableProperty]
    private bool _includeTemporaryFiles;

    /// <summary>Maximum simultaneous link requests.</summary>
    [ObservableProperty]
    private int _maxConcurrency = 4;

    /// <summary>Pause between requests, in milliseconds.</summary>
    [ObservableProperty]
    private int _requestDelayMilliseconds;

    /// <summary>Maximum retry attempts for a retryable failure.</summary>
    [ObservableProperty]
    private int _maxRetryAttempts = 5;

    /// <summary>True to stop the run at the first failure.</summary>
    [ObservableProperty]
    private bool _stopOnFirstError;

    /// <summary>True to enumerate and validate without changing anything. On by default.</summary>
    [ObservableProperty]
    private bool _dryRun = true;

    /// <summary>How overlapping targets are reconciled.</summary>
    [ObservableProperty]
    private OverlapResolution _overlapResolution = OverlapResolution.KeepParent;

    /// <summary>Live progress of the current run.</summary>
    [ObservableProperty]
    private JobProgress _progress = new();

    /// <summary>The summary of the last completed run.</summary>
    [ObservableProperty]
    private JobSummary? _summary;

    /// <summary>True while a run is in flight.</summary>
    [ObservableProperty]
    private bool _isRunning;

    /// <summary>True while the run is paused.</summary>
    [ObservableProperty]
    private bool _isPaused;

    /// <summary>Creates the page.</summary>
    public NewLinkJobViewModel(
        ILinkJobRunner runner,
        ConnectionCoordinator connection,
        JobDraft draft,
        IJobHistoryStore historyStore,
        ISettingsStore settingsStore,
        IProductMetadataProvider productMetadata,
        IClipboardService clipboard,
        ISystemBrowser browser,
        SharePointBrowserViewModel sharePointBrowser,
        OneDriveBrowserViewModel oneDriveBrowser,
        ILogger<NewLinkJobViewModel> logger)
        : base("New Link Job", "job")
    {
        SharePointBrowser = sharePointBrowser
            ?? throw new ArgumentNullException(nameof(sharePointBrowser));

        OneDriveBrowser = oneDriveBrowser ?? throw new ArgumentNullException(nameof(oneDriveBrowser));

        _runner = runner ?? throw new ArgumentNullException(nameof(runner));
        _connection = connection ?? throw new ArgumentNullException(nameof(connection));
        _draft = draft ?? throw new ArgumentNullException(nameof(draft));
        _historyStore = historyStore ?? throw new ArgumentNullException(nameof(historyStore));
        _settingsStore = settingsStore ?? throw new ArgumentNullException(nameof(settingsStore));
        _productMetadata = productMetadata ?? throw new ArgumentNullException(nameof(productMetadata));
        _clipboard = clipboard ?? throw new ArgumentNullException(nameof(clipboard));
        _browser = browser ?? throw new ArgumentNullException(nameof(browser));
        _logger = logger ?? throw new ArgumentNullException(nameof(logger));

        _draft.TargetsChanged += (_, _) => RefreshTargets();
        RefreshTargets();
    }

    /// <summary>The targets currently in the job.</summary>
    public ObservableCollection<ProcessingTarget> Targets { get; } = [];

    /// <summary>Files the preview says would be processed.</summary>
    public ObservableCollection<DiscoveredFile> PreviewCandidates { get; } = [];

    /// <summary>Items the preview says would be skipped, with reasons.</summary>
    public ObservableCollection<DiscoveredFile> PreviewSkipped { get; } = [];

    /// <summary>Per-file outcomes from the last run.</summary>
    public ObservableCollection<LinkResult> Results { get; } = [];

    /// <summary>Manifest write outcomes from the last run.</summary>
    public ObservableCollection<ManifestWriteResult> ManifestResults { get; } = [];

    /// <summary>Conditions preflight actually confirmed.</summary>
    public ObservableCollection<string> PreflightValidated { get; } = [];

    /// <summary>Conditions preflight expects but did not confirm.</summary>
    public ObservableCollection<string> PreflightExpected { get; } = [];

    /// <summary>Conditions that cannot be known until the job runs.</summary>
    public ObservableCollection<string> PreflightUnknown { get; } = [];

    /// <summary>Problems that prevent the job from running.</summary>
    public ObservableCollection<string> PreflightBlockers { get; } = [];

    /// <summary>Non-fatal concerns.</summary>
    public ObservableCollection<string> PreflightWarnings { get; } = [];

    /// <summary>Where manifests would be written.</summary>
    public ObservableCollection<string> ManifestDestinations { get; } = [];

    /// <summary>Available link permissions.</summary>
    public static IReadOnlyList<LinkPermission> LinkPermissions { get; } =
        [LinkPermission.View, LinkPermission.Edit];

    /// <summary>Available link audiences.</summary>
    public static IReadOnlyList<LinkAudience> LinkAudiences { get; } =
        [LinkAudience.Organization, LinkAudience.SpecificPeople, LinkAudience.Anyone];

    /// <summary>Available conflict policies.</summary>
    public static IReadOnlyList<ManifestConflictPolicy> ConflictPolicies { get; } =
        Enum.GetValues<ManifestConflictPolicy>();

    /// <summary>Available missing-entry policies.</summary>
    public static IReadOnlyList<MissingEntryPolicy> MissingEntryPolicies { get; } =
        Enum.GetValues<MissingEntryPolicy>();

    /// <summary>Available overlap resolutions.</summary>
    public static IReadOnlyList<OverlapResolution> OverlapResolutions { get; } =
        Enum.GetValues<OverlapResolution>();

    /// <summary>True when a preview has been produced and the job may be started.</summary>
    public bool CanStart => _preview is not null && _preview.Preflight.CanProceed && !IsRunning;

    /// <summary>
    /// Why the job cannot start, or what it will do differently if it does. Empty when the job is
    /// ready to run and will run for real.
    /// <para>
    /// A disabled button with no explanation is the worst of both worlds: the user can see that
    /// something is wrong but not what, and the reasons here are all actionable. Dry run is
    /// included even though it does not block anything, because "the job ran and created nothing"
    /// is otherwise indistinguishable from a failure.
    /// </para>
    /// </summary>
    public string StartBlockedReason
    {
        get
        {
            if (IsRunning)
            {
                return "A job is already running.";
            }

            if (ValidationProblems.Count > 0)
            {
                return ValidationProblems.Count == 1
                    ? "Fix the problem listed below before starting."
                    : $"Fix the {ValidationProblems.Count} problems listed below before starting.";
            }

            if (_preview is null)
            {
                return "Build a preview before the job can be started.";
            }

            if (!_preview.Preflight.CanProceed)
            {
                return PreflightBlockers.Count > 0
                    ? "The preview found problems that prevent this job from running. "
                      + "See 'Blocked' below."
                    : "The preview found a problem that prevents this job from running.";
            }

            return DryRun
                ? "Dry run is on. This will enumerate, filter and validate, but create no links "
                  + "and write no manifests. Clear the dry-run box to make real changes."
                : string.Empty;
        }
    }

    /// <summary>True when there is something to say about starting, blocking or otherwise.</summary>
    public bool HasStartBlockedReason => StartBlockedReason.Length > 0;

    /// <summary>True when the explanation is a blocker rather than a note about dry run.</summary>
    public bool IsStartBlocked => !CanStart;

    /// <summary>True when the audience needs a recipient list.</summary>
    public bool ShowRecipients => LinkAudience == LinkAudience.SpecificPeople;

    /// <summary>True when the highest-risk audience is selected.</summary>
    public bool ShowAnyoneWarning => LinkAudience == LinkAudience.Anyone;

    /// <summary>A one-line description of what will be requested.</summary>
    public string LinkSummary => BuildLinkConfiguration().Describe();

    /// <summary>A readable summary of the active filters.</summary>
    public string FilterSummary => BuildFilterConfiguration().Describe();

    /// <summary>Problems with the current settings, shown before a preview is attempted.</summary>
    public ObservableCollection<string> ValidationProblems { get; } = [];

    /// <summary>Failures from the last run that could be retried.</summary>
    public IReadOnlyList<LinkResult> RetryableFailures =>
        Summary?.RetryableFailures ?? [];

    /// <summary>True when there is anything worth retrying.</summary>
    public bool HasRetryableFailures => RetryableFailures.Count > 0;

    /// <inheritdoc />
    public override async Task OnNavigatedToAsync(CancellationToken cancellationToken = default)
    {
        RefreshTargets();

        var settings = await _settingsStore.LoadAsync(cancellationToken).ConfigureAwait(true);

        if (Targets.Count == 0 && Results.Count == 0)
        {
            MaxConcurrency = settings.DefaultExecution.MaxConcurrency;
            MaxRetryAttempts = settings.DefaultExecution.MaxRetryAttempts;
            RequestDelayMilliseconds = (int)settings.DefaultExecution.RequestDelay.TotalMilliseconds;
        }
    }

    /// <summary>Removes a target from the job.</summary>
    [RelayCommand]
    private void RemoveTarget(ProcessingTarget? target)
    {
        if (target is not null)
        {
            _draft.RemoveTarget(target);
        }
    }

    /// <summary>Enables or disables a target without removing it.</summary>
    [RelayCommand]
    private void ToggleTargetEnabled(ProcessingTarget? target)
    {
        if (target is not null)
        {
            _draft.ReplaceTarget(target, target with { IsEnabled = !target.IsEnabled });
        }
    }

    /// <summary>Toggles whether a target descends into subfolders.</summary>
    [RelayCommand]
    private void ToggleTargetRecursive(ProcessingTarget? target)
    {
        if (target is not null)
        {
            _draft.ReplaceTarget(target, target with { Recursive = !target.Recursive });
        }
    }

    /// <summary>Removes every target.</summary>
    [RelayCommand]
    private void ClearTargets() => _draft.Clear();

    /// <summary>Enumerates and validates without changing anything.</summary>
    [RelayCommand]
    private async Task BuildPreviewAsync(CancellationToken cancellationToken)
    {
        ClearMessages();
        ValidationProblems.Clear();

        if (_connection.Tenant is null)
        {
            ErrorMessage = "Connect to Microsoft 365 before previewing a job.";
            return;
        }

        SyncDraftFromInputs();

        var configuration = _draft.ToConfiguration(
            _connection.Tenant.TenantId,
            _connection.TenantDisplayName ?? _connection.Tenant.TenantDisplayName);

        foreach (var problem in configuration.Validate())
        {
            ValidationProblems.Add(problem);
        }

        if (ValidationProblems.Count > 0)
        {
            ErrorMessage = "Fix the problems listed before previewing.";
            return;
        }

        IsBusy = true;
        SelectedTabIndex = 4;

        try
        {
            var progress = new Progress<JobProgress>(p => Progress = p);

            _preview = await _runner.PreviewAsync(configuration, progress, cancellationToken)
                .ConfigureAwait(true);

            ApplyPreview(_preview);

            StatusMessage = _preview.Preflight.CanProceed
                ? $"{_preview.CandidateCount} file(s) would be processed, {_preview.SkippedCount} skipped."
                : "The job cannot run yet. Review the blockers below.";
        }
        catch (OperationCanceledException)
        {
            StatusMessage = "Preview cancelled.";
        }
#pragma warning disable CA1031 // Keep the UI responsive whatever goes wrong during preview.
        catch (Exception ex)
        {
            _logger.LogError(ex, "Building the job preview failed.");
            ErrorMessage = "The preview could not be completed. See Diagnostics for details.";
        }
#pragma warning restore CA1031
        finally
        {
            IsBusy = false;
            NotifyStartStateChanged();
        }
    }

    /// <summary>Runs the job. Requires a preview the user has seen.</summary>
    [RelayCommand]
    private async Task StartAsync()
    {
        if (_preview is null || _connection.Tenant is null)
        {
            ErrorMessage = "Build a preview first.";
            return;
        }

        SyncDraftFromInputs();

        var configuration = _draft.ToConfiguration(
            _connection.Tenant.TenantId,
            _connection.TenantDisplayName ?? _connection.Tenant.TenantDisplayName);

        _runCancellation?.Dispose();
        _runCancellation = new CancellationTokenSource();

        _pauseToken?.Dispose();
        _pauseToken = new PauseToken();
        _pauseToken.PauseStateChanged += (_, paused) => IsPaused = paused;

        IsRunning = true;
        ClearMessages();
        Results.Clear();
        ManifestResults.Clear();
        SelectedTabIndex = 5;

        try
        {
            var progress = new Progress<JobProgress>(p => Progress = p);

            var summary = await _runner
                .RunAsync(configuration, _preview, progress, _pauseToken, _runCancellation.Token)
                .ConfigureAwait(true);

            ApplySummary(summary);
            await RecordHistoryAsync(configuration, summary).ConfigureAwait(true);

            StatusMessage = summary.WasDryRun
                ? "Dry run finished. Nothing was created or modified."
                : summary.WasCancelled
                    ? $"Cancelled. {summary.SuccessCount} link(s) were produced before stopping and have "
                      + "been written to manifests."
                    : $"Finished. {summary.CreatedCount} created, {summary.ReusedCount} reused, "
                      + $"{summary.SkippedCount} skipped, {summary.FailedCount} failed.";
        }
#pragma warning disable CA1031 // A run failure must surface as a message, not an unhandled crash.
        catch (Exception ex)
        {
            _logger.LogError(ex, "The link job failed.");
            ErrorMessage = "The job stopped because of an unexpected error. See Diagnostics for details.";
        }
#pragma warning restore CA1031
        finally
        {
            IsRunning = false;
            IsPaused = false;
            NotifyStartStateChanged();
        }
    }

    /// <summary>Pauses the run at the next safe point.</summary>
    [RelayCommand]
    private void Pause() => _pauseToken?.Pause();

    /// <summary>Resumes a paused run.</summary>
    [RelayCommand]
    private void Resume() => _pauseToken?.Resume();

    /// <summary>Cancels the run, preserving results already produced.</summary>
    [RelayCommand]
    private void Cancel()
    {
        _runCancellation?.Cancel();
        StatusMessage = "Cancelling. Results already produced will be kept.";
    }

    /// <summary>Retries the retryable failures from the last run.</summary>
    [RelayCommand]
    private async Task RetryFailuresAsync(CancellationToken cancellationToken)
    {
        if (Summary is null || _connection.Tenant is null || !HasRetryableFailures)
        {
            return;
        }

        var configuration = _draft.ToConfiguration(
            _connection.Tenant.TenantId,
            _connection.TenantDisplayName ?? _connection.Tenant.TenantDisplayName);

        IsRunning = true;

        try
        {
            var progress = new Progress<JobProgress>(p => Progress = p);

            var summary = await _runner
                .RetryFailuresAsync(configuration, Summary.RetryableFailures, progress, cancellationToken)
                .ConfigureAwait(true);

            ApplySummary(summary);
            StatusMessage = $"Retry finished. {summary.SuccessCount} of {summary.Results.Count} succeeded.";
        }
        finally
        {
            IsRunning = false;
        }
    }

    /// <summary>Copies the last run's manifest locations.</summary>
    [RelayCommand]
    private async Task CopyManifestLocationsAsync()
    {
        if (ManifestResults.Count == 0)
        {
            return;
        }

        var text = string.Join(Environment.NewLine, ManifestResults.Select(m => m.WebUrl ?? m.DisplayPath));
        await _clipboard.SetTextAsync(text).ConfigureAwait(true);
        StatusMessage = "Manifest locations copied.";
    }

    /// <summary>Opens a written manifest in the browser.</summary>
    [RelayCommand]
    private async Task OpenManifestAsync(ManifestWriteResult? manifest)
    {
        if (manifest?.WebUrl is { Length: > 0 } url && Uri.TryCreate(url, UriKind.Absolute, out var uri))
        {
            await _browser.OpenAsync(uri).ConfigureAwait(true);
        }
    }

    /// <summary>Releases the run's cancellation source and pause token.</summary>
    public void Dispose()
    {
        _runCancellation?.Dispose();
        _pauseToken?.Dispose();
    }

    private void RefreshTargets()
    {
        Targets.Clear();

        foreach (var target in _draft.Targets)
        {
            Targets.Add(target);
        }

        NotifyStartStateChanged();
    }

    private void ApplyPreview(JobPreview preview)
    {
        PreviewCandidates.Clear();
        PreviewSkipped.Clear();
        ManifestDestinations.Clear();
        PreflightValidated.Clear();
        PreflightExpected.Clear();
        PreflightUnknown.Clear();
        PreflightBlockers.Clear();
        PreflightWarnings.Clear();

        // The grid is capped: a preview of 200,000 rows would freeze the UI and tell the user
        // nothing they cannot learn from the count plus a sample.
        foreach (var candidate in preview.Candidates.Take(2000))
        {
            PreviewCandidates.Add(candidate);
        }

        foreach (var skipped in preview.Skipped.Take(2000))
        {
            PreviewSkipped.Add(skipped);
        }

        foreach (var destination in preview.ManifestDestinations)
        {
            ManifestDestinations.Add(destination);
        }

        foreach (var item in preview.Preflight.Validated)
        {
            PreflightValidated.Add(item);
        }

        foreach (var item in preview.Preflight.Expected)
        {
            PreflightExpected.Add(item);
        }

        foreach (var item in preview.Preflight.UnknownUntilExecution)
        {
            PreflightUnknown.Add(item);
        }

        foreach (var item in preview.Preflight.Blockers)
        {
            PreflightBlockers.Add(item);
        }

        foreach (var item in preview.Preflight.Warnings)
        {
            PreflightWarnings.Add(item);
        }
    }

    private void ApplySummary(JobSummary summary)
    {
        Summary = summary;

        Results.Clear();
        ManifestResults.Clear();

        foreach (var result in summary.Results.Take(5000))
        {
            Results.Add(result);
        }

        foreach (var manifest in summary.Manifests)
        {
            ManifestResults.Add(manifest);
        }

        OnPropertyChanged(nameof(RetryableFailures));
        OnPropertyChanged(nameof(HasRetryableFailures));
    }

    private async Task RecordHistoryAsync(JobConfiguration configuration, JobSummary summary)
    {
        var settings = await _settingsStore.LoadAsync().ConfigureAwait(true);

        var entry = new JobHistoryEntry
        {
            JobId = summary.JobId,
            Name = configuration.Name,
            ApplicationVersion = _productMetadata.Version,
            StartedUtc = summary.StartedUtc,
            CompletedUtc = summary.CompletedUtc,

            // A privacy-conscious identifier, never the full user principal name.
            AccountIdentifier = _connection.Account?.PrivacyIdentifier ?? "unknown",
            TenantDisplayName = configuration.TenantDisplayName,
            TenantId = configuration.TenantId,
            TargetDescriptions = configuration.Targets.Select(t => t.DisplayPath).ToArray(),
            LinkSettingsSummary = configuration.Link.Describe(),
            ManifestSettingsSummary = string.Join(", ", configuration.Manifest.GeneratedFileNames()),
            FilterSummary = configuration.Filters.Describe(),
            CreatedCount = summary.CreatedCount,
            ReusedCount = summary.ReusedCount,
            SkippedCount = summary.SkippedCount,
            FailedCount = summary.FailedCount,
            ManifestLocations = summary.Manifests.Select(m => m.WebUrl ?? m.DisplayPath).ToArray(),
            SanitizedErrors = summary.Results
                .Where(r => r.Error is not null)
                .Select(r => $"{r.Error!.Kind}: {r.Error.Message}")
                .Distinct(StringComparer.Ordinal)
                .Take(100)
                .ToArray(),
            WasDryRun = summary.WasDryRun,
            WasCancelled = summary.WasCancelled,
            FinalPhase = summary.FinalPhase,
            Configuration = configuration,
        };

        await _historyStore.AppendAsync(entry, settings.JobHistoryRetentionCount).ConfigureAwait(true);
    }

    private void SyncDraftFromInputs()
    {
        _draft.Link = BuildLinkConfiguration();
        _draft.Manifest = BuildManifestConfiguration();
        _draft.Filters = BuildFilterConfiguration();
        _draft.Execution = BuildExecutionOptions();
        _draft.DryRun = DryRun;
        _draft.OverlapResolution = OverlapResolution;
    }

    private LinkConfiguration BuildLinkConfiguration() => new()
    {
        Permission = LinkPermission,
        Audience = LinkAudience,
        Recipients = ParseRecipients(RecipientsText),
        SendInvitationEmail = SendInvitationEmail,
        InvitationMessage = string.IsNullOrWhiteSpace(InvitationMessage) ? null : InvitationMessage,
        ExpirationUtc = UseExpiration ? ExpirationDate : null,
        RetainInheritedPermissions = RetainInheritedPermissions,
        ReuseExistingLinks = ReuseExistingLinks,
        SkipWhenEquivalentLinkExists = SkipWhenEquivalentLinkExists,
    };

    private ManifestConfiguration BuildManifestConfiguration()
    {
        var formats = ManifestFormats.None;

        if (FormatPlainText)
        {
            formats |= ManifestFormats.PlainText;
        }

        if (FormatMarkdown)
        {
            formats |= ManifestFormats.Markdown;
        }

        if (FormatCsv)
        {
            formats |= ManifestFormats.Csv;
        }

        if (FormatJson)
        {
            formats |= ManifestFormats.Json;
        }

        return new ManifestConfiguration
        {
            WritePerFolderManifest = WritePerFolderManifest,
            WriteMasterManifest = WriteMasterManifest,
            Formats = formats,
            AggregateDescendantsInPerFolderManifest = AggregateDescendants,
            ConflictPolicy = ConflictPolicy,
            MissingEntryPolicy = MissingEntryPolicy,
        };
    }

    private FilterConfiguration BuildFilterConfiguration() => new()
    {
        IncludeExtensions = SplitList(IncludeExtensions),
        ExcludeExtensions = SplitList(ExcludeExtensions),
        IncludePatterns = SplitList(IncludePatterns),
        ExcludePatterns = SplitList(ExcludePatterns),
        ModifiedAfterUtc = ModifiedAfter,
        ModifiedBeforeUtc = ModifiedBefore,
        MinimumSizeBytes = MinimumSizeKilobytes is { } min ? (long)(min * 1024) : null,
        MaximumSizeBytes = MaximumSizeKilobytes is { } max ? (long)(max * 1024) : null,
        IncludeHiddenAndSystemItems = IncludeHiddenItems,
        IncludeTemporaryFiles = IncludeTemporaryFiles,
    };

    private ExecutionOptions BuildExecutionOptions() => new()
    {
        MaxConcurrency = MaxConcurrency,
        RequestDelay = TimeSpan.FromMilliseconds(RequestDelayMilliseconds),
        MaxRetryAttempts = MaxRetryAttempts,
        FailureBehavior = StopOnFirstError
            ? FailureBehavior.StopOnFirstError
            : FailureBehavior.ContinueOnError,
    };

    /// <summary>
    /// Splits a user-typed list on commas, semicolons and newlines, so a pasted column from a
    /// spreadsheet works as well as a typed comma-separated list.
    /// </summary>
    internal static IReadOnlyList<string> SplitList(string? value) =>
        string.IsNullOrWhiteSpace(value)
            ? []
            : value.Split([',', ';', '\n', '\r'], StringSplitOptions.RemoveEmptyEntries
                    | StringSplitOptions.TrimEntries)
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToArray();

    /// <summary>Parses the recipient text box into a distinct list of addresses.</summary>
    internal static IReadOnlyList<string> ParseRecipients(string? value) => SplitList(value);

    partial void OnLinkAudienceChanged(LinkAudience value)
    {
        OnPropertyChanged(nameof(ShowRecipients));
        OnPropertyChanged(nameof(ShowAnyoneWarning));
        OnPropertyChanged(nameof(LinkSummary));

        // An invitation email is only meaningful for a specific-people link, so the option is
        // reset rather than left silently set on a configuration that cannot honour it.
        if (value != LinkAudience.SpecificPeople)
        {
            SendInvitationEmail = false;
        }
    }

    partial void OnLinkPermissionChanged(LinkPermission value) => OnPropertyChanged(nameof(LinkSummary));

    partial void OnUseExpirationChanged(bool value) => OnPropertyChanged(nameof(LinkSummary));

    partial void OnIsRunningChanged(bool value) => NotifyStartStateChanged();

    partial void OnDryRunChanged(bool value) => NotifyStartStateChanged();

    /// <summary>
    /// Raises change notification for everything the Start button and its explanation depend on.
    /// Kept together so the button and the sentence beside it can never disagree.
    /// </summary>
    private void NotifyStartStateChanged()
    {
        OnPropertyChanged(nameof(CanStart));
        OnPropertyChanged(nameof(StartBlockedReason));
        OnPropertyChanged(nameof(HasStartBlockedReason));
        OnPropertyChanged(nameof(IsStartBlocked));
        StartCommand.NotifyCanExecuteChanged();
    }

    /// <summary>Whether a tab index is currently showing.</summary>
    private bool IsTabVisible(int index) => index switch
    {
        BrowseSharePointTab => IsBrowsingSharePoint,
        BrowseOneDriveTab => IsBrowsingOneDrive,
        _ => true,
    };

    /// <summary>
    /// The next showing tab in a direction, or null when there is none. Back and Next have to
    /// skip hidden tabs: a hidden tab still occupies an index, so stepping by one would land on
    /// a tab that renders nothing.
    /// </summary>
    private int? AdjacentVisibleTab(int direction)
    {
        for (var index = SelectedTabIndex + direction; index >= 0 && index <= LastTab; index += direction)
        {
            if (IsTabVisible(index))
            {
                return index;
            }
        }

        return null;
    }

    /// <summary>Moves to the previous step that is showing.</summary>
    [RelayCommand(CanExecute = nameof(CanGoToPreviousStep))]
    private void PreviousStep()
    {
        if (AdjacentVisibleTab(-1) is { } index)
        {
            SelectedTabIndex = index;
        }
    }

    private bool CanGoToPreviousStep() => AdjacentVisibleTab(-1) is not null;

    /// <summary>Moves to the next step that is showing.</summary>
    [RelayCommand(CanExecute = nameof(CanGoToNextStep))]
    private void NextStep()
    {
        if (AdjacentVisibleTab(1) is { } index)
        {
            SelectedTabIndex = index;
        }
    }

    private bool CanGoToNextStep() => AdjacentVisibleTab(1) is not null;

    /// <summary>Opens the SharePoint browser as a step of this job.</summary>
    [RelayCommand]
    private async Task BrowseSharePointAsync(CancellationToken cancellationToken)
    {
        IsBrowsingSharePoint = true;
        SelectedTabIndex = BrowseSharePointTab;
        OnPropertyChanged(nameof(IsBrowsing));

        // The browser used to be a page, so it loaded when navigated to. Hosting it here means
        // this call is what replaces that.
        await SharePointBrowser.OnNavigatedToAsync(cancellationToken).ConfigureAwait(true);
    }

    /// <summary>Opens the OneDrive browser as a step of this job.</summary>
    [RelayCommand]
    private async Task BrowseOneDriveAsync(CancellationToken cancellationToken)
    {
        IsBrowsingOneDrive = true;
        SelectedTabIndex = BrowseOneDriveTab;
        OnPropertyChanged(nameof(IsBrowsing));

        await OneDriveBrowser.OnNavigatedToAsync(cancellationToken).ConfigureAwait(true);
    }

    /// <summary>
    /// Returns to the targets step and puts the browse tabs away.
    /// <para>
    /// Both are closed rather than only the one being left, so the tab strip returns to the six
    /// numbered steps and does not accumulate browse tabs from earlier in the session. The
    /// selections themselves live in the shared draft and are unaffected.
    /// </para>
    /// </summary>
    [RelayCommand]
    private void ReturnToTargets()
    {
        IsBrowsingSharePoint = false;
        IsBrowsingOneDrive = false;
        OnPropertyChanged(nameof(IsBrowsing));
        SelectedTabIndex = TargetsTab;
        RefreshTargets();
    }
}
