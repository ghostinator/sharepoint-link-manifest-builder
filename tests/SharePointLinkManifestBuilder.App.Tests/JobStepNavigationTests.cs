using Microsoft.Extensions.DependencyInjection;
using SharePointLinkManifestBuilder.App.Composition;
using SharePointLinkManifestBuilder.App.ViewModels;
using SharePointLinkManifestBuilder.Core.Settings;

namespace SharePointLinkManifestBuilder.App.Tests;

/// <summary>
/// The resource browsers are steps of the job rather than separate destinations. Back and Next
/// therefore have to walk the tabs that are actually showing: a hidden tab still occupies an
/// index, so stepping by one would land on a tab that renders nothing.
/// </summary>
public sealed class JobStepNavigationTests : IDisposable
{
    private const int TargetsTab = 0;
    private const int BrowseSharePointTab = 1;
    private const int BrowseOneDriveTab = 2;
    private const int LinkTab = 3;
    private const int ResultsTab = 7;

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

    /// <summary>
    /// The key property: with the browse tabs hidden, Next from Targets must reach Link, not the
    /// hidden SharePoint tab immediately after it.
    /// </summary>
    [Fact]
    public void Next_FromTargets_SkipsTheHiddenBrowseTabs()
    {
        var job = Job();

        job.NextStepCommand.Execute(null);

        Assert.Equal(LinkTab, job.SelectedTabIndex);
    }

    /// <summary>And Back from Link must return to Targets rather than a hidden tab.</summary>
    [Fact]
    public void Previous_FromLink_SkipsTheHiddenBrowseTabs()
    {
        var job = Job();
        job.SelectedTabIndex = LinkTab;

        job.PreviousStepCommand.Execute(null);

        Assert.Equal(TargetsTab, job.SelectedTabIndex);
    }

    /// <summary>Opening a browser shows its tab and selects it.</summary>
    [Fact]
    public async Task BrowseSharePoint_ShowsAndSelectsThatTab()
    {
        var job = Job();

        await job.BrowseSharePointCommand.ExecuteAsync(null);

        Assert.True(job.IsBrowsingSharePoint);
        Assert.True(job.IsBrowsing);
        Assert.Equal(BrowseSharePointTab, job.SelectedTabIndex);
        Assert.False(job.IsBrowsingOneDrive);
    }

    /// <summary>The OneDrive browser behaves the same way.</summary>
    [Fact]
    public async Task BrowseOneDrive_ShowsAndSelectsThatTab()
    {
        var job = Job();

        await job.BrowseOneDriveCommand.ExecuteAsync(null);

        Assert.True(job.IsBrowsingOneDrive);
        Assert.Equal(BrowseOneDriveTab, job.SelectedTabIndex);
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

    /// <summary>While a browse tab is open, Next must step through it rather than over it.</summary>
    [Fact]
    public async Task Next_FromTargets_EntersAVisibleBrowseTab()
    {
        var job = Job();
        await job.BrowseSharePointCommand.ExecuteAsync(null);
        job.SelectedTabIndex = TargetsTab;

        job.NextStepCommand.Execute(null);

        Assert.Equal(BrowseSharePointTab, job.SelectedTabIndex);
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
    /// The position label counts the six numbered steps and names the browse tabs instead of
    /// numbering them, so "Step 3 of 6" always means the same tab.
    /// </summary>
    [Fact]
    public async Task StepPosition_NumbersOnlyTheNumberedSteps()
    {
        var job = Job();

        Assert.Equal("Step 1 of 6", job.StepPosition);

        job.SelectedTabIndex = LinkTab;
        Assert.Equal("Step 2 of 6", job.StepPosition);

        job.SelectedTabIndex = ResultsTab;
        Assert.Equal("Step 6 of 6", job.StepPosition);

        await job.BrowseSharePointCommand.ExecuteAsync(null);
        Assert.Equal("Choosing SharePoint locations", job.StepPosition);
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
