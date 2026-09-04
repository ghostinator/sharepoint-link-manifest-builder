using Microsoft.Extensions.Logging;
using SharePointLinkManifestBuilder.Core.Abstractions;
using SharePointLinkManifestBuilder.Core.Manifests;
using SharePointLinkManifestBuilder.Core.Models;
using SharePointLinkManifestBuilder.Core.Targets;

namespace SharePointLinkManifestBuilder.Core.Jobs;

/// <summary>
/// Runs a link job through its phases: validate, resolve, enumerate, preview, execute, build
/// manifests, upload, summarize.
/// <para>
/// The runner never starts creating links on its own. <see cref="RunAsync"/> requires a preview
/// the caller has already shown to the user and had confirmed, which is what keeps a bulk
/// permission change from being one mis-click away.
/// </para>
/// </summary>
public sealed class LinkJobRunner : ILinkJobRunner
{
    private readonly IFileDiscoveryService _discovery;
    private readonly ISharingLinkService _sharing;
    private readonly IManifestStorageService _manifestStorage;
    private readonly IManifestBuilder _manifestBuilder;
    private readonly IManifestMerger _manifestMerger;
    private readonly ManifestConflictResolver _conflictResolver;
    private readonly ISiteService _siteService;
    private readonly IDriveService _driveService;
    private readonly IAuthenticationService _authentication;
    private readonly IProductMetadataProvider _productMetadata;
    private readonly IReadOnlyList<IManifestFormatter> _formatters;
    private readonly ILogger<LinkJobRunner> _logger;
    private readonly TimeProvider _timeProvider;

    /// <summary>Creates the runner.</summary>
    public LinkJobRunner(
        IFileDiscoveryService discovery,
        ISharingLinkService sharing,
        IManifestStorageService manifestStorage,
        IManifestBuilder manifestBuilder,
        IManifestMerger manifestMerger,
        ManifestConflictResolver conflictResolver,
        ISiteService siteService,
        IDriveService driveService,
        IAuthenticationService authentication,
        IProductMetadataProvider productMetadata,
        IEnumerable<IManifestFormatter> formatters,
        ILogger<LinkJobRunner> logger,
        TimeProvider? timeProvider = null)
    {
        _discovery = discovery ?? throw new ArgumentNullException(nameof(discovery));
        _sharing = sharing ?? throw new ArgumentNullException(nameof(sharing));
        _manifestStorage = manifestStorage ?? throw new ArgumentNullException(nameof(manifestStorage));
        _manifestBuilder = manifestBuilder ?? throw new ArgumentNullException(nameof(manifestBuilder));
        _manifestMerger = manifestMerger ?? throw new ArgumentNullException(nameof(manifestMerger));
        _conflictResolver = conflictResolver ?? throw new ArgumentNullException(nameof(conflictResolver));
        _siteService = siteService ?? throw new ArgumentNullException(nameof(siteService));
        _driveService = driveService ?? throw new ArgumentNullException(nameof(driveService));
        _authentication = authentication ?? throw new ArgumentNullException(nameof(authentication));
        _productMetadata = productMetadata ?? throw new ArgumentNullException(nameof(productMetadata));
        _formatters = formatters?.ToArray() ?? throw new ArgumentNullException(nameof(formatters));
        _logger = logger ?? throw new ArgumentNullException(nameof(logger));
        _timeProvider = timeProvider ?? TimeProvider.System;
    }

    /// <inheritdoc />
    public async Task<PreflightReport> PreflightAsync(
        JobConfiguration configuration,
        CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(configuration);

        var validated = new List<string>();
        var expected = new List<string>();
        var unknown = new List<string>();
        var blockers = new List<string>(configuration.Validate());
        var warnings = new List<string>();

        // 1. Identity and tenant consistency.
        var account = _authentication.CurrentAccount;

        if (account is null)
        {
            blockers.Add("You are not signed in to Microsoft 365.");
        }
        else if (!string.Equals(account.TenantId, configuration.TenantId, StringComparison.OrdinalIgnoreCase))
        {
            blockers.Add(
                "The signed-in account belongs to a different tenant than this job. "
                + "Switch tenant, or rebuild the job against the current tenant.");
        }
        else
        {
            validated.Add($"Signed in as {account.DisplayName} in the expected tenant.");
        }

        // 2. Resolve targets, expanding whole-site targets into one target per library.
        var resolved = new List<ProcessingTarget>();

        foreach (var target in configuration.EnabledTargets)
        {
            cancellationToken.ThrowIfCancellationRequested();

            if (target.SourceType == TargetSourceType.SharePointSite)
            {
                var expansion = await ExpandSiteTargetAsync(target, cancellationToken).ConfigureAwait(false);

                if (expansion.Count == 0)
                {
                    warnings.Add($"No accessible document libraries were found in '{target.SiteName}'.");
                    continue;
                }

                resolved.AddRange(expansion);
                validated.Add($"'{target.SiteName}' expanded to {expansion.Count} document librar(y/ies).");
                continue;
            }

            if (!target.IsResolved)
            {
                blockers.Add($"Target '{target.DisplayPath}' could not be resolved to a drive.");
                continue;
            }

            resolved.Add(target);
        }

        // 3. Overlap reconciliation.
        var plan = TargetPlanner.Plan(resolved, configuration.OverlapResolution);
        warnings.AddRange(plan.Warnings);

        foreach (var (removed, reason) in plan.Removed)
        {
            warnings.Add($"'{removed.DisplayPath}' will not be processed separately: {reason}");
        }

        // 4. Reachability of each surviving target.
        foreach (var target in plan.EffectiveTargets)
        {
            cancellationToken.ThrowIfCancellationRequested();

            var folder = string.IsNullOrEmpty(target.StartingFolderItemId)
                ? await _driveService
                    .GetFolderByPathAsync(target.DriveId!, target.StartingFolderRelativePath, cancellationToken)
                    .ConfigureAwait(false)
                : OperationResult<SharePointFolder>.Success(new SharePointFolder
                {
                    DriveId = target.DriveId!,
                    ItemId = target.StartingFolderItemId!,
                    Name = target.StartingFolderName ?? "folder",
                });

            if (folder.Succeeded)
            {
                validated.Add($"'{target.DisplayPath}' is reachable.");
            }
            else
            {
                blockers.Add($"'{target.DisplayPath}' could not be opened: {folder.Error!.Message}");
            }
        }

        // 5. Things that genuinely cannot be known before execution. Saying so plainly is more
        //    useful than a green tick that turns out to have meant nothing.
        if (!configuration.DryRun)
        {
            unknown.Add(
                "Whether your organization's sharing policy permits the requested link type and audience. "
                + "Microsoft 365 decides this when each link is actually requested.");

            unknown.Add(
                "Whether every individual file allows sharing. Item-level permissions and sensitivity "
                + "labels are evaluated per file at the moment the link is created.");

            if (configuration.Link.Audience == LinkAudience.Anyone)
            {
                warnings.Add(
                    "'Anyone with the link' produces links usable without signing in, including outside your "
                    + "organization. Many tenants disable this; those requests will be reported as policy-blocked.");
            }

            if (configuration.Link.Recipients.Count > 0)
            {
                expected.Add(
                    $"{configuration.Link.Recipients.Count} recipient(s) will be granted access"
                    + (configuration.Link.SendInvitationEmail
                        ? " and Microsoft 365 will email them."
                        : " without any email being sent."));
            }
        }
        else
        {
            validated.Add("Dry run: no link will be created and no manifest will be written.");
        }

        expected.Add(
            configuration.Manifest.WritePerFolderManifest && configuration.Manifest.WriteMasterManifest
                ? "Per-folder and master manifests will be written."
                : configuration.Manifest.WriteMasterManifest
                    ? "A master manifest will be written."
                    : "A per-folder manifest will be written in each folder containing processed files.");

        return new PreflightReport
        {
            CanProceed = blockers.Count == 0,
            Validated = validated,
            Expected = expected,
            UnknownUntilExecution = unknown,
            Blockers = blockers,
            Warnings = warnings,
            ResolvedTargets = plan.EffectiveTargets,
            Overlaps = plan.Overlaps,
        };
    }

    /// <inheritdoc />
    public async Task<JobPreview> PreviewAsync(
        JobConfiguration configuration,
        IProgress<JobProgress>? progress = null,
        CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(configuration);

        var preflight = await PreflightAsync(configuration, cancellationToken).ConfigureAwait(false);

        var candidates = new List<DiscoveredFile>();
        var skipped = new List<DiscoveredFile>();

        // Deduplication across overlapping targets keys on (driveId, itemId), never on path,
        // so a file reachable by two routes is still processed exactly once.
        var seen = new HashSet<string>(StringComparer.Ordinal);

        var targetIndex = 0;

        foreach (var target in preflight.ResolvedTargets)
        {
            cancellationToken.ThrowIfCancellationRequested();
            targetIndex++;

            progress?.Report(new JobProgress
            {
                Phase = JobPhase.EnumeratingFiles,
                CurrentTargetName = target.DisplayPath,
                CurrentTargetIndex = targetIndex,
                TotalTargets = preflight.ResolvedTargets.Count,
                ProcessedCount = candidates.Count,
            });

            await foreach (var file in _discovery
                .DiscoverAsync(target, configuration.Filters, configuration.Manifest, cancellationToken)
                .ConfigureAwait(false))
            {
                if (file.SkipReason != SkipReason.None)
                {
                    skipped.Add(file);
                    continue;
                }

                if (!seen.Add(file.IdentityKey))
                {
                    skipped.Add(file with { SkipReason = SkipReason.DuplicateAcrossTargets });
                    continue;
                }

                candidates.Add(file);

                if (candidates.Count % 100 == 0)
                {
                    progress?.Report(new JobProgress
                    {
                        Phase = JobPhase.EnumeratingFiles,
                        CurrentTargetName = target.DisplayPath,
                        CurrentTargetIndex = targetIndex,
                        TotalTargets = preflight.ResolvedTargets.Count,
                        ProcessedCount = candidates.Count,
                        CurrentFileName = file.Name,
                        CurrentRelativePath = file.RelativePath,
                    });
                }
            }
        }

        var destinations = DescribeManifestDestinations(configuration, preflight.ResolvedTargets, candidates);

        return new JobPreview
        {
            Preflight = preflight,
            Candidates = candidates,
            Skipped = skipped,
            ManifestDestinations = destinations,
            WasDryRun = configuration.DryRun,

            // One createLink per file, plus one invite per file when recipients are named, plus
            // one write per manifest. Reported only because it can be computed exactly; a guess
            // would be worse than no number.
            EstimatedApiOperations = configuration.DryRun
                ? 0
                : candidates.Count * (configuration.Link.RequiresInviteAction ? 2 : 1) + destinations.Count,
        };
    }

    /// <summary>
    /// How long manifest writing may continue after the job was cancelled. Short on purpose: it
    /// exists to record links that already exist, not to keep working against the user's wish.
    /// </summary>
    private static readonly TimeSpan ManifestWriteAfterCancelBudget = TimeSpan.FromSeconds(30);

    /// <inheritdoc />
    public async Task<JobSummary> RunAsync(
        JobConfiguration configuration,
        JobPreview preview,
        IProgress<JobProgress>? progress = null,
        IPauseToken? pauseToken = null,
        CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(configuration);
        ArgumentNullException.ThrowIfNull(preview);

        var startedUtc = _timeProvider.GetUtcNow();
        var results = new List<LinkResult>();
        var manifests = new List<ManifestWriteResult>();
        var jobErrors = new List<GraphError>();
        var warnings = new List<string>(preview.Preflight.Warnings);

        if (configuration.DryRun)
        {
            // Belt and braces. A dry run should never reach here, but if it does it must not
            // change anything.
            _logger.LogInformation("Dry run requested; no links will be created and no manifests written.");

            return new JobSummary
            {
                JobId = configuration.JobId,
                ApplicationVersion = _productMetadata.Version,
                StartedUtc = startedUtc,
                CompletedUtc = _timeProvider.GetUtcNow(),
                FinalPhase = JobPhase.Completed,
                WasDryRun = true,
                Warnings = warnings,
            };
        }

        var finalPhase = JobPhase.Completed;

        try
        {
            results.AddRange(await CreateLinksAsync(
                configuration, preview.Candidates, progress, pauseToken, cancellationToken)
                .ConfigureAwait(false));
        }
        catch (OperationCanceledException)
        {
            // Reached only when cancellation happens before CreateLinksAsync can return, such
            // as before the first file starts. Ordinary mid-run cancellation returns partial
            // results instead of throwing, and is detected below.
            finalPhase = JobPhase.Cancelled;
        }

        // Cancellation is read from the token, not inferred from an exception, because
        // CreateLinksAsync now returns its partial results rather than throwing them away.
        if (cancellationToken.IsCancellationRequested)
        {
            finalPhase = JobPhase.Cancelled;

            _logger.LogInformation(
                "Job cancelled after processing {Count} file(s). Those results are preserved and "
                + "reported, because the links they describe exist in the tenant.",
                results.Count);
        }

        // Manifests are written even after cancellation, because the successes already produced
        // are real and discarding them would waste the tenant writes that created them, leaving
        // links that exist but that nothing records.
        //
        // That needs its own token. Passing the cancelled one made this block throw on its first
        // await, so the promise in the comment above was never kept: a cancelled run created
        // links and then wrote nothing describing them. The window is deliberately short, and
        // cancelling again during it does stop the writes.
        if (results.Any(r => r.IsSuccess))
        {
            using var manifestCancellation = cancellationToken.IsCancellationRequested
                ? new CancellationTokenSource(ManifestWriteAfterCancelBudget)
                : null;

            var manifestToken = manifestCancellation?.Token ?? cancellationToken;

            if (manifestCancellation is not null)
            {
                _logger.LogInformation(
                    "Writing manifests for the {Count} link(s) already created, despite the "
                    + "cancellation, so they are not left unrecorded.",
                    results.Count(r => r.IsSuccess));
            }

            try
            {
                manifests.AddRange(await WriteManifestsAsync(
                    configuration, preview.Preflight.ResolvedTargets, results, progress, manifestToken)
                    .ConfigureAwait(false));
            }
            catch (OperationCanceledException)
            {
                finalPhase = JobPhase.Cancelled;
            }
#pragma warning disable CA1031 // A manifest failure must not discard the link results.
            catch (Exception ex)
            {
                _logger.LogError(ex, "Manifest writing failed. Link results are preserved.");
                jobErrors.Add(new GraphError
                {
                    Kind = GraphErrorKind.Unknown,
                    Message = "The links were created, but writing the manifests failed.",
                    SuggestedAction = "Review the errors below and re-run with manifest writing only.",
                });
            }
#pragma warning restore CA1031
        }

        var completedUtc = _timeProvider.GetUtcNow();

        progress?.Report(new JobProgress
        {
            Phase = finalPhase,
            ProcessedCount = results.Count,
            TotalCount = preview.Candidates.Count,
            CreatedCount = results.Count(r => r.Status == LinkResultStatus.Created),
            ReusedCount = results.Count(r => r.Status is LinkResultStatus.Reused or LinkResultStatus.Existing),
            FailedCount = results.Count(r => !r.IsSuccess),
        });

        return new JobSummary
        {
            JobId = configuration.JobId,
            ApplicationVersion = _productMetadata.Version,
            StartedUtc = startedUtc,
            CompletedUtc = completedUtc,
            FinalPhase = finalPhase,
            WasDryRun = false,
            Results = results,
            Manifests = manifests,
            JobLevelErrors = jobErrors,
            Warnings = warnings,
        };
    }

    /// <inheritdoc />
    public async Task<JobSummary> RetryFailuresAsync(
        JobConfiguration configuration,
        IReadOnlyList<LinkResult> failures,
        IProgress<JobProgress>? progress = null,
        CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(configuration);
        ArgumentNullException.ThrowIfNull(failures);

        // Reusing the original configuration is what makes a retry genuinely equivalent to the
        // original attempt rather than a subtly different job.
        var retryCandidates = failures.Where(f => f.IsRetryable).Select(f => f.File).ToArray();

        var preview = new JobPreview
        {
            Preflight = await PreflightAsync(configuration, cancellationToken).ConfigureAwait(false),
            Candidates = retryCandidates,
        };

        return await RunAsync(configuration, preview, progress, pauseToken: null, cancellationToken)
            .ConfigureAwait(false);
    }

    private async Task<IReadOnlyList<LinkResult>> CreateLinksAsync(
        JobConfiguration configuration,
        IReadOnlyList<DiscoveredFile> candidates,
        IProgress<JobProgress>? progress,
        IPauseToken? pauseToken,
        CancellationToken cancellationToken)
    {
        var results = new List<LinkResult>(candidates.Count);
        // Disposed rather than left to the finalizer: one is created per job run, and every
        // task that takes it is awaited by the Task.WhenAll below, so the using scope cannot
        // close while anything still holds it.
        using var gate = new SemaphoreSlim(configuration.Execution.MaxConcurrency);
        var resultsLock = new Lock();
        var stopRequested = false;

        var counters = new Counters();

        var tasks = candidates.Select(async file =>
        {
            if (Volatile.Read(ref stopRequested))
            {
                return;
            }

            await gate.WaitAsync(cancellationToken).ConfigureAwait(false);

            try
            {
                if (pauseToken is not null)
                {
                    await pauseToken.WaitWhilePausedAsync(cancellationToken).ConfigureAwait(false);
                }

                if (Volatile.Read(ref stopRequested))
                {
                    return;
                }

                if (configuration.Execution.RequestDelay > TimeSpan.Zero)
                {
                    await Task.Delay(configuration.Execution.RequestDelay, _timeProvider, cancellationToken)
                        .ConfigureAwait(false);
                }

                var result = await _sharing
                    .CreateOrGetLinkAsync(file, configuration.Link, cancellationToken)
                    .ConfigureAwait(false);

                lock (resultsLock)
                {
                    results.Add(result);
                    counters.Record(result);

                    progress?.Report(new JobProgress
                    {
                        Phase = JobPhase.CreatingLinks,
                        CurrentFileName = file.Name,
                        CurrentRelativePath = file.RelativePath,
                        CurrentOperation = "createLink",
                        ProcessedCount = results.Count,
                        TotalCount = candidates.Count,
                        CreatedCount = counters.Created,
                        ReusedCount = counters.Reused,
                        SkippedCount = counters.Skipped,
                        FailedCount = counters.Failed,
                    });
                }

                if (!result.IsSuccess
                    && configuration.Execution.FailureBehavior == FailureBehavior.StopOnFirstError)
                {
                    _logger.LogWarning(
                        "Stopping on first error, as configured. Results already produced are preserved.");
                    Volatile.Write(ref stopRequested, true);
                }
            }
            finally
            {
                gate.Release();
            }
        });

        try
        {
            await Task.WhenAll(tasks).ConfigureAwait(false);
        }
        catch (OperationCanceledException)
        {
            // Swallowed on purpose, and this is the whole point of the method's contract.
            //
            // Cancelling makes every in-flight task throw, so WhenAll throws, so this method
            // used to unwind without ever reaching the return below. The caller's
            // `results.AddRange(await CreateLinksAsync(...))` therefore never ran and the list
            // it was adding to stayed empty -- a cancelled job reported zero created, zero
            // reused, zero failed, however many links it had actually made in the tenant.
            //
            // The results are not lost when the exception is thrown; they are already in this
            // list, added under resultsLock as each file finished. Returning them is all that
            // was ever needed. The caller distinguishes a cancelled run by inspecting the
            // token, not by catching this.
            _logger.LogInformation(
                "Cancelled with {Count} of {Total} file(s) already processed. "
                + "Those results are real and are preserved.",
                results.Count,
                candidates.Count);
        }

        return results;
    }

    private async Task<IReadOnlyList<ManifestWriteResult>> WriteManifestsAsync(
        JobConfiguration configuration,
        IReadOnlyList<ProcessingTarget> targets,
        IReadOnlyList<LinkResult> results,
        IProgress<JobProgress>? progress,
        CancellationToken cancellationToken)
    {
        var written = new List<ManifestWriteResult>();

        progress?.Report(new JobProgress { Phase = JobPhase.BuildingManifests });

        foreach (var target in targets)
        {
            cancellationToken.ThrowIfCancellationRequested();

            var targetResults = results
                .Where(r => string.Equals(r.File.TargetId, target.TargetId, StringComparison.Ordinal))
                .ToArray();

            if (targetResults.Length == 0)
            {
                continue;
            }

            if (configuration.Manifest.WritePerFolderManifest)
            {
                var perFolder = _manifestBuilder.BuildPerFolderManifests(
                    configuration, target, targetResults, _productMetadata.Version);

                foreach (var (folderItemId, _, document) in perFolder)
                {
                    written.AddRange(await WriteDocumentAsync(
                        configuration, target.DriveId!, folderItemId,
                        configuration.Manifest.PerFolderFileName, document, isMaster: false,
                        progress, cancellationToken).ConfigureAwait(false));
                }
            }

            if (configuration.Manifest.WriteMasterManifest && configuration.Manifest.MasterManifestPerTarget)
            {
                var master = _manifestBuilder.BuildMasterManifest(
                    configuration, target, targetResults, _productMetadata.Version);

                var destinationFolderId = target.StartingFolderItemId
                    ?? (await _driveService.GetRootFolderAsync(target.DriveId!, cancellationToken)
                        .ConfigureAwait(false)).Value?.ItemId;

                if (destinationFolderId is not null)
                {
                    written.AddRange(await WriteDocumentAsync(
                        configuration, target.DriveId!, destinationFolderId,
                        configuration.Manifest.MasterFileName, master, isMaster: true,
                        progress, cancellationToken).ConfigureAwait(false));
                }
            }
        }

        if (configuration.Manifest.WriteMasterManifest
            && !configuration.Manifest.MasterManifestPerTarget
            && configuration.Manifest.CombinedMasterDestination is { } combined)
        {
            var document = _manifestBuilder.BuildCombinedMasterManifest(
                configuration, results, _productMetadata.Version);

            written.AddRange(await WriteDocumentAsync(
                configuration, combined.DriveId, combined.FolderItemId,
                configuration.Manifest.MasterFileName, document, isMaster: true,
                progress, cancellationToken).ConfigureAwait(false));
        }

        return written;
    }

    private async Task<IReadOnlyList<ManifestWriteResult>> WriteDocumentAsync(
        JobConfiguration configuration,
        string driveId,
        string parentItemId,
        string baseName,
        ManifestDocument document,
        bool isMaster,
        IProgress<JobProgress>? progress,
        CancellationToken cancellationToken)
    {
        var written = new List<ManifestWriteResult>();

        foreach (var format in ManifestDefaults.Split(configuration.Manifest.Formats))
        {
            cancellationToken.ThrowIfCancellationRequested();

            var formatter = _formatters.FirstOrDefault(f => f.Format == format);

            if (formatter is null)
            {
                _logger.LogWarning("No formatter is registered for {Format}; skipping.", format);
                continue;
            }

            var fileName = baseName + formatter.FileExtension;

            progress?.Report(new JobProgress
            {
                Phase = JobPhase.UploadingManifests,
                CurrentFileName = fileName,
                CurrentOperation = "upload manifest",
            });

            var existing = await _manifestStorage
                .ReadManifestAsync(driveId, parentItemId, fileName, cancellationToken)
                .ConfigureAwait(false);

            var decision = _conflictResolver.Resolve(
                fileName,
                existing.Succeeded ? existing.Value : null,
                configuration.Manifest.ConflictPolicy);

            if (decision.Action is ManifestWriteAction.Skip or ManifestWriteAction.Fail)
            {
                written.Add(new ManifestWriteResult
                {
                    DisplayPath = fileName,
                    Format = format,
                    IsMaster = isMaster,
                    EntryCount = document.Entries.Count,
                    Succeeded = decision.Action == ManifestWriteAction.Skip,
                    ConflictOutcome = decision.Explanation,
                    Error = decision.Action == ManifestWriteAction.Fail
                        ? new GraphError
                        {
                            Kind = GraphErrorKind.ManifestConflict,
                            Message = decision.Explanation,
                        }
                        : null,
                });

                continue;
            }

            // Merging happens only for a document this application recognises as its own, and
            // only for the plain-text format, which is the one with a parser.
            var toWrite = decision.Action == ManifestWriteAction.MergeAndReplace
                && decision.ExistingDocument is { } previous
                    ? _manifestMerger.Merge(previous, document, configuration.Manifest.MissingEntryPolicy)
                    : document;

            var content = formatter.Render(toWrite);

            var result = await _manifestStorage.WriteManifestAsync(
                driveId,
                parentItemId,
                decision.FileName,
                content,
                format,
                isMaster,
                toWrite.Entries.Count,
                decision.IfMatchETag,
                cancellationToken).ConfigureAwait(false);

            written.Add(result.Succeeded
                ? result.Value! with { ConflictOutcome = decision.Explanation }
                : new ManifestWriteResult
                {
                    DisplayPath = decision.FileName,
                    Format = format,
                    IsMaster = isMaster,
                    EntryCount = toWrite.Entries.Count,
                    Succeeded = false,
                    ConflictOutcome = decision.Explanation,
                    Error = result.Error,
                });
        }

        return written;
    }

    private async Task<IReadOnlyList<ProcessingTarget>> ExpandSiteTargetAsync(
        ProcessingTarget siteTarget,
        CancellationToken cancellationToken)
    {
        if (string.IsNullOrEmpty(siteTarget.SiteId))
        {
            return [];
        }

        var drives = await _siteService.GetSiteDrivesAsync(siteTarget.SiteId, cancellationToken)
            .ConfigureAwait(false);

        if (!drives.Succeeded)
        {
            _logger.LogWarning(
                "Could not list document libraries for '{Site}': {Reason}",
                siteTarget.SiteName,
                drives.Error!.Message);

            return [];
        }

        return drives.Value!
            .Select(drive => siteTarget with
            {
                TargetId = $"{siteTarget.TargetId}:{drive.DriveId}",
                SourceType = TargetSourceType.DocumentLibrary,
                DriveId = drive.DriveId,
                DriveName = drive.Name,
                StartingFolderItemId = null,
                StartingFolderRelativePath = string.Empty,
                WebUrl = drive.WebUrl,
            })
            .ToArray();
    }

    private static List<string> DescribeManifestDestinations(
        JobConfiguration configuration,
        IReadOnlyList<ProcessingTarget> targets,
        IReadOnlyList<DiscoveredFile> candidates)
    {
        var destinations = new List<string>();
        var extensions = ManifestDefaults.ExtensionsFor(configuration.Manifest.Formats).ToArray();

        if (configuration.Manifest.WritePerFolderManifest)
        {
            var folders = candidates
                .Select(c => c.ParentRelativePath)
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .Count();

            foreach (var extension in extensions)
            {
                destinations.Add(
                    $"{folders} x {configuration.Manifest.PerFolderFileName}{extension} "
                    + "(one in each folder containing processed files)");
            }
        }

        if (configuration.Manifest.WriteMasterManifest)
        {
            if (configuration.Manifest.MasterManifestPerTarget)
            {
                foreach (var target in targets)
                {
                    foreach (var extension in extensions)
                    {
                        destinations.Add(
                            $"{target.DisplayPath}/{configuration.Manifest.MasterFileName}{extension}");
                    }
                }
            }
            else if (configuration.Manifest.CombinedMasterDestination is { } combined)
            {
                foreach (var extension in extensions)
                {
                    destinations.Add(
                        $"{combined.DisplayPath}/{configuration.Manifest.MasterFileName}{extension} (combined)");
                }
            }
        }

        return destinations;
    }

    /// <summary>Running tallies, guarded by the caller's lock.</summary>
    private sealed class Counters
    {
        public int Created { get; private set; }

        public int Reused { get; private set; }

        public int Skipped { get; private set; }

        public int Failed { get; private set; }

        public void Record(LinkResult result)
        {
            switch (result.Status)
            {
                case LinkResultStatus.Created:
                    Created++;
                    break;

                case LinkResultStatus.Reused:
                case LinkResultStatus.Existing:
                    Reused++;
                    break;

                case LinkResultStatus.Skipped:
                    Skipped++;
                    break;

                default:
                    Failed++;
                    break;
            }
        }
    }
}
