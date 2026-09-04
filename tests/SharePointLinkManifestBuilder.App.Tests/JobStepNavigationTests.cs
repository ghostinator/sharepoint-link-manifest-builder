using Microsoft.Extensions.DependencyInjection;
using SharePointLinkManifestBuilder.App.Composition;
using SharePointLinkManifestBuilder.App.ViewModels;
using SharePointLinkManifestBuilder.Core.Settings;

namespace SharePointLinkManifestBuilder.App.Tests;

/// <summary>
/// The resource browsers live inside the Targets step rather than as tabs of their own, so the
/// numbered strip stays exactly six wide and choosing locations never moves the selected step.
/// </summary>
public sealed class JobStepNavigationTests : IDisposable
{
    private const int TargetsTab = 0;
    private const int LinkTab = 1;
    private const int ResultsTab = 5;

    private readonly string _stateDirectory;
    private readonly ServiceProvider _services;

    /// <summary>Builds a provider rooted at a temporary directory, never the real user profile.</summary>
    public JobStepNavigationTests()
    {
        _stateDirectory = Path.Combine(
            Path.GetTempPath(), "splmb-tests", Guid.NewGuid().ToString("n"));

        _services = ServiceRegistration.Build(new ApplicationPaths(_stateDirectory));
    }

    /// <inheritdoc />
    public void Dispose()
    {
        _services.Dispose();

        if (Directory.Exists(_stateDirectory))
        {
            Directory.Delete(_stateDirectory, recursive: true);
        }
    }

    private NewLinkJobViewModel Job() => _services.GetRequiredService<NewLinkJobViewModel>();

    /// <summary>Both browse tabs stay out of the way until asked for.</summary>
    [Fact]
    public void BrowseTabs_AreHiddenInitially()
    {
        var job = Job();

        Assert.False(job.IsBrowsingSharePoint);
        Assert.False(job.IsBrowsingOneDrive);
        Assert.False(job.IsBrowsing);
        Assert.Equal(TargetsTab, job.SelectedTabIndex);
    }

    /// <summary>The numbered strip is contiguous, so Next from Targets reaches Link.</summary>
    [Fact]
    public void Next_FromTargets_ReachesLink()
    {
        var job = Job();

        job.NextStepCommand.Execute(null);

        Assert.Equal(LinkTab, job.SelectedTabIndex);
    }

    /// <summary>Browsing must not shift the numbered steps.</summary>
    [Fact]
    public async Task Next_WhileBrowsing_StillReachesLink()
    {
        var job = Job();
        await job.BrowseSharePointCommand.ExecuteAsync(null);

        job.NextStepCommand.Execute(null);

        Assert.Equal(LinkTab, job.SelectedTabIndex);
    }

    /// <summary>Browsing opens inside the Targets step and leaves it selected.</summary>
    [Fact]
    public async Task BrowseSharePoint_StaysOnTheTargetsStep()
    {
        var job = Job();

        await job.BrowseSharePointCommand.ExecuteAsync(null);

        Assert.True(job.IsBrowsingSharePoint);
        Assert.True(job.IsBrowsing);
        Assert.Equal(TargetsTab, job.SelectedTabIndex);
        Assert.False(job.IsBrowsingOneDrive);
    }

    /// <summary>Only one browser shows at a time, so switching replaces rather than stacks.</summary>
    [Fact]
    public async Task SwitchingBrowsers_ShowsOnlyTheNewOne()
    {
        var job = Job();

        await job.BrowseSharePointCommand.ExecuteAsync(null);
        await job.BrowseOneDriveCommand.ExecuteAsync(null);

        Assert.True(job.IsBrowsingOneDrive);
        Assert.False(job.IsBrowsingSharePoint);
        Assert.Equal(TargetsTab, job.SelectedTabIndex);
    }

    /// <summary>
    /// Stepping away from Targets puts the browsers away, so coming back shows the target list
    /// the user just built rather than the browser they left open.
    /// </summary>
    [Fact]
    public async Task LeavingTheTargetsStep_ClosesTheBrowsers()
    {
        var job = Job();
        await job.BrowseSharePointCommand.ExecuteAsync(null);

        job.NextStepCommand.Execute(null);

        Assert.False(job.IsBrowsing);

        job.PreviousStepCommand.Execute(null);

        Assert.Equal(TargetsTab, job.SelectedTabIndex);
        Assert.False(job.IsBrowsing);
    }

    /// <summary>
    /// Returning puts both browse tabs away, not only the one being left, so the tab strip does
    /// not accumulate them across a session.
    /// </summary>
    [Fact]
    public async Task ReturnToTargets_ClosesBothBrowseTabs()
    {
        var job = Job();
        await job.BrowseSharePointCommand.ExecuteAsync(null);
        await job.BrowseOneDriveCommand.ExecuteAsync(null);

        job.ReturnToTargetsCommand.Execute(null);

        Assert.False(job.IsBrowsingSharePoint);
        Assert.False(job.IsBrowsingOneDrive);
        Assert.False(job.IsBrowsing);
        Assert.Equal(TargetsTab, job.SelectedTabIndex);
    }

    /// <summary>Navigation stops at the ends rather than wrapping, which would misread as a loop.</summary>
    [Fact]
    public void Navigation_StopsAtBothEnds()
    {
        var job = Job();

        Assert.False(job.PreviousStepCommand.CanExecute(null));

        job.SelectedTabIndex = ResultsTab;

        Assert.False(job.NextStepCommand.CanExecute(null));
        Assert.True(job.PreviousStepCommand.CanExecute(null));
    }

    /// <summary>
    /// The six numbered steps map straight onto their indices, and browsing annotates step 1
    /// rather than becoming a step of its own.
    /// </summary>
    [Fact]
    public async Task StepPosition_TracksTheSixNumberedSteps()
    {
        var job = Job();

        Assert.Equal("Step 1 of 6", job.StepPosition);

        job.SelectedTabIndex = LinkTab;
        Assert.Equal("Step 2 of 6", job.StepPosition);

        job.SelectedTabIndex = ResultsTab;
        Assert.Equal("Step 6 of 6", job.StepPosition);

        await job.BrowseSharePointCommand.ExecuteAsync(null);
        Assert.Equal("Step 1 of 6 — choosing locations", job.StepPosition);
    }

    /// <summary>The job page owns the browsers, so both are reachable from it.</summary>
    [Fact]
    public void JobPage_HostsBothBrowsers()
    {
        var job = Job();

        Assert.NotNull(job.SharePointBrowser);
        Assert.NotNull(job.OneDriveBrowser);
        Assert.Same(_services.GetRequiredService<SharePointBrowserViewModel>(), job.SharePointBrowser);
    }
}
