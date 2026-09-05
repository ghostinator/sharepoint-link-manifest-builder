using Microsoft.Extensions.Logging.Abstractions;
using NSubstitute;
using SharePointLinkManifestBuilder.Core.Abstractions;
using SharePointLinkManifestBuilder.Core.Jobs;
using SharePointLinkManifestBuilder.Core.Manifests;
using SharePointLinkManifestBuilder.Core.Models;

namespace SharePointLinkManifestBuilder.Core.Tests.Jobs;

/// <summary>
/// Cancelling a running job must still report what it did. The links already created exist in
/// the tenant whether or not the run finished, so a summary of zeros is not a smaller truth —
/// it is a false one, and it hides work the user has to reconcile by hand.
/// </summary>
public sealed class LinkJobCancellationTests
{
    private static DiscoveredFile File(int i) => new()
    {
        DriveId = "drive-1",
        ItemId = $"item-{i}",
        Name = $"file-{i}.docx",
        RelativePath = $"Folder/file-{i}.docx",
    };

    private static JobConfiguration Configuration() => new()
    {
        JobId = Guid.NewGuid().ToString("n"),
        TenantId = "11111111-1111-1111-1111-111111111111",
        Targets = [],
        Link = new LinkConfiguration { Permission = LinkPermission.View, Audience = LinkAudience.Organization },
        Manifest = new ManifestConfiguration { WritePerFolderManifest = false, WriteMasterManifest = false },
        Execution = new ExecutionOptions { MaxConcurrency = 1, RequestDelay = TimeSpan.Zero },
        DryRun = false,
    };

    private static JobPreview Preview(IReadOnlyList<DiscoveredFile> candidates) => new()
    {
        Preflight = new PreflightReport { CanProceed = true },
        Candidates = candidates,
    };

    /// <summary>
    /// Builds a runner whose sharing service succeeds, and cancels once <paramref name="cancelAfter"/>
    /// files have been processed — the shape of a user pressing Cancel mid-run.
    /// </summary>
    private static (LinkJobRunner Runner, CancellationTokenSource Cancellation) Build(int cancelAfter)
    {
        var cancellation = new CancellationTokenSource();
        var processed = 0;

        var sharing = Substitute.For<ISharingLinkService>();
        sharing.CreateOrGetLinkAsync(Arg.Any<DiscoveredFile>(), Arg.Any<LinkConfiguration>(), Arg.Any<CancellationToken>())
            .Returns(call =>
            {
                var file = call.Arg<DiscoveredFile>();

                if (Interlocked.Increment(ref processed) >= cancelAfter)
                {
                    cancellation.Cancel();
                }

                return Task.FromResult(new LinkResult
                {
                    File = file,
                    Status = LinkResultStatus.Created,
                    SharingUrl = $"https://example.sharepoint.test/:w:/s/x/{file.ItemId}",
                });
            });

        var productMetadata = Substitute.For<IProductMetadataProvider>();
        productMetadata.Version.Returns("0.0.0-test");

        var runner = new LinkJobRunner(
            Substitute.For<IFileDiscoveryService>(),
            sharing,
            Substitute.For<IManifestStorageService>(),
            Substitute.For<IManifestBuilder>(),
            Substitute.For<IManifestMerger>(),
            new ManifestConflictResolver(Substitute.For<IManifestParser>()),
            Substitute.For<ISiteService>(),
            Substitute.For<IDriveService>(),
            Substitute.For<IAuthenticationService>(),
            productMetadata,
            [],
            NullLogger<LinkJobRunner>.Instance);

        return (runner, cancellation);
    }

    /// <summary>
    /// The reported bug. Cancelling made every counter read zero even though links had been
    /// created, because the partial results were thrown away with the cancellation exception.
    /// </summary>
    [Fact]
    public async Task Cancelling_StillReportsTheLinksAlreadyCreated()
    {
        var (runner, cancellation) = Build(cancelAfter: 3);
        var candidates = Enumerable.Range(1, 50).Select(File).ToArray();

        var summary = await runner.RunAsync(
            Configuration(), Preview(candidates), progress: null, pauseToken: null, cancellation.Token);

        Assert.Equal(JobPhase.Cancelled, summary.FinalPhase);
        Assert.NotEmpty(summary.Results);
        Assert.All(summary.Results, r => Assert.Equal(LinkResultStatus.Created, r.Status));
    }

    /// <summary>
    /// It must report fewer than everything, too. A cancellation that silently completed the job
    /// would also be wrong, and would pass the assertion above.
    /// </summary>
    [Fact]
    public async Task Cancelling_ReportsAPartialRunNotAFullOne()
    {
        var (runner, cancellation) = Build(cancelAfter: 3);
        var candidates = Enumerable.Range(1, 50).Select(File).ToArray();

        var summary = await runner.RunAsync(
            Configuration(), Preview(candidates), progress: null, pauseToken: null, cancellation.Token);

        Assert.InRange(summary.Results.Count, 1, candidates.Length - 1);
    }

    /// <summary>
    /// Every preserved result must describe a real file, since these are what the user reconciles
    /// against the tenant afterwards.
    /// </summary>
    [Fact]
    public async Task PreservedResults_CarryTheirFileAndLink()
    {
        var (runner, cancellation) = Build(cancelAfter: 2);
        var candidates = Enumerable.Range(1, 20).Select(File).ToArray();

        var summary = await runner.RunAsync(
            Configuration(), Preview(candidates), progress: null, pauseToken: null, cancellation.Token);

        Assert.All(summary.Results, r =>
        {
            Assert.False(string.IsNullOrWhiteSpace(r.File.ItemId));
            Assert.False(string.IsNullOrWhiteSpace(r.SharingUrl));
        });
    }

    /// <summary>A run that is never cancelled must still complete normally and process everything.</summary>
    [Fact]
    public async Task WithoutCancellation_TheJobCompletesAndProcessesEveryCandidate()
    {
        var (runner, _) = Build(cancelAfter: int.MaxValue);
        var candidates = Enumerable.Range(1, 10).Select(File).ToArray();

        var summary = await runner.RunAsync(
            Configuration(), Preview(candidates), progress: null, pauseToken: null, CancellationToken.None);

        Assert.Equal(JobPhase.Completed, summary.FinalPhase);
        Assert.Equal(candidates.Length, summary.Results.Count);
    }

    /// <summary>A dry run changes nothing and says so, cancelled or not.</summary>
    [Fact]
    public async Task ADryRun_ReportsItselfAsSuchAndCreatesNothing()
    {
        var (runner, _) = Build(cancelAfter: int.MaxValue);
        var candidates = Enumerable.Range(1, 5).Select(File).ToArray();

        var summary = await runner.RunAsync(
            Configuration() with { DryRun = true },
            Preview(candidates), progress: null, pauseToken: null, CancellationToken.None);

        Assert.True(summary.WasDryRun);
        Assert.Empty(summary.Results);
    }
}
