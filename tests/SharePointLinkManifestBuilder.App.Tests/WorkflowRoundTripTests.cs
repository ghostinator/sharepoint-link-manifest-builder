using Microsoft.Extensions.DependencyInjection;
using SharePointLinkManifestBuilder.App.Composition;
using SharePointLinkManifestBuilder.App.ViewModels;
using SharePointLinkManifestBuilder.Core.Settings;

namespace SharePointLinkManifestBuilder.App.Tests;

/// <summary>
/// The job draft is shared state, and the actions that move things into and out of it used to
/// live on pages other than the job. A button that names a destination has to reach it, or the
/// user is left on a page whose work is done with no indication anything happened.
/// </summary>
public sealed class WorkflowRoundTripTests : IDisposable
{
    private readonly string _stateDirectory;
    private readonly ServiceProvider _services;

    /// <summary>Builds a provider rooted at a temporary directory, never the real user profile.</summary>
    public WorkflowRoundTripTests()
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

    /// <summary>Every page can ask the shell to navigate, which is what makes the round trips work.</summary>
    [Fact]
    public void EveryPage_CanRequestNavigation()
    {
        var shell = _services.GetRequiredService<MainWindowViewModel>();

        Assert.All(shell.Pages, page =>
            Assert.NotNull(page.GetType().GetEvent(nameof(PageViewModelBase.NavigationRequested))));
    }

    /// <summary>
    /// "Load into New Link Job" has to arrive at the job page. With nothing selected it must do
    /// nothing at all rather than navigating to an unchanged draft.
    /// </summary>
    [Fact]
    public void LoadingAProfile_WithNothingSelected_DoesNotNavigate()
    {
        var profiles = _services.GetRequiredService<SavedProfilesViewModel>();
        var navigatedTo = new List<string>();
        profiles.NavigationRequested += (_, key) => navigatedTo.Add(key);

        profiles.LoadProfileCommand.Execute(null);

        Assert.Empty(navigatedTo);
    }

    /// <summary>The same guard on the history page's equivalent action.</summary>
    [Fact]
    public void LoadingAHistoryEntry_WithNothingSelected_DoesNotNavigate()
    {
        var history = _services.GetRequiredService<JobHistoryViewModel>();
        var navigatedTo = new List<string>();
        history.NavigationRequested += (_, key) => navigatedTo.Add(key);

        history.RerunConfigurationCommand.Execute(null);

        Assert.Empty(navigatedTo);
        Assert.NotNull(history.ErrorMessage);
    }

    /// <summary>
    /// A profile can be saved from the job page, which is where the configuration is built.
    /// Without a name it must refuse rather than saving something unfindable.
    /// </summary>
    [Fact]
    public async Task SavingAProfile_WithoutAName_IsRefused()
    {
        var job = _services.GetRequiredService<NewLinkJobViewModel>();
        job.ProfileName = "   ";

        await job.SaveAsProfileCommand.ExecuteAsync(null);

        Assert.NotNull(job.ErrorMessage);
        Assert.Contains("name", job.ErrorMessage!, StringComparison.OrdinalIgnoreCase);
    }

    /// <summary>
    /// Saving needs a tenant, because a profile records which organization it was built for.
    /// Disconnected, it must say so rather than failing later.
    /// </summary>
    [Fact]
    public async Task SavingAProfile_WhileDisconnected_SaysSo()
    {
        var job = _services.GetRequiredService<NewLinkJobViewModel>();
        job.ProfileName = "Nightly marketing links";

        await job.SaveAsProfileCommand.ExecuteAsync(null);

        Assert.NotNull(job.ErrorMessage);
        Assert.Contains("Connect", job.ErrorMessage!, StringComparison.OrdinalIgnoreCase);
    }

    /// <summary>
    /// The shell resolves the key these pages request. A typo here would be a silent no-op,
    /// exactly the failure mode this whole change is about.
    /// </summary>
    [Fact]
    public async Task TheJobPageKey_ResolvesInTheShell()
    {
        var shell = _services.GetRequiredService<MainWindowViewModel>();

        await shell.NavigateToAsync("job");

        Assert.Equal("job", shell.CurrentPage!.NavigationKey);
    }
}
