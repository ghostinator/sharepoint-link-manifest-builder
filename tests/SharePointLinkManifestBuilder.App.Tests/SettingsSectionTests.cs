using Microsoft.Extensions.DependencyInjection;
using SharePointLinkManifestBuilder.App.Composition;
using SharePointLinkManifestBuilder.App.ViewModels;
using SharePointLinkManifestBuilder.Core.Settings;

namespace SharePointLinkManifestBuilder.App.Tests;

/// <summary>
/// Settings hosts the permissions review and diagnostics as sections rather than leaving them as
/// sidebar destinations. Both used to load when navigated to, so hosting them means Settings has
/// to load them instead.
/// </summary>
public sealed class SettingsSectionTests : IDisposable
{
    private readonly string _stateDirectory;
    private readonly ServiceProvider _services;

    /// <summary>Builds a provider rooted at a temporary directory, never the real user profile.</summary>
    public SettingsSectionTests()
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

    /// <summary>Settings owns both sections, and they are the instances the container holds.</summary>
    [Fact]
    public void Settings_HostsPermissionsAndDiagnostics()
    {
        var settings = _services.GetRequiredService<SettingsViewModel>();

        Assert.Same(_services.GetRequiredService<PermissionsViewModel>(), settings.Permissions);
        Assert.Same(_services.GetRequiredService<DiagnosticsViewModel>(), settings.Diagnostics);
    }

    /// <summary>
    /// Navigating to Settings must load both sections. A tab strip does not tell the view model
    /// which tab is showing, so loading only the visible one would leave the others blank until
    /// something else happened to populate them.
    /// </summary>
    [Fact]
    public async Task NavigatingToSettings_LoadsBothSectionsWithoutError()
    {
        var shell = _services.GetRequiredService<MainWindowViewModel>();

        await shell.NavigateToAsync("settings");

        Assert.Equal("settings", shell.CurrentPage!.NavigationKey);
        Assert.False(shell.CurrentPage.HasError);
    }

    /// <summary>The sidebar is shorter by exactly those two entries.</summary>
    [Fact]
    public void TheSidebar_NoLongerListsThemAsDestinations()
    {
        var shell = _services.GetRequiredService<MainWindowViewModel>();

        Assert.Equal(8, shell.Pages.Count);
        Assert.DoesNotContain(shell.Pages, p => p.NavigationKey == "permissions");
        Assert.DoesNotContain(shell.Pages, p => p.NavigationKey == "diagnostics");
        Assert.Contains(shell.Pages, p => p.NavigationKey == "settings");
    }

    /// <summary>
    /// This build ships a placeholder update endpoint, so the check is unavailable rather than
    /// offered and then apologised to. A button that can only report its own uselessness is
    /// worse than one that is visibly unavailable with the reason beside it.
    /// </summary>
    [Fact]
    public void UpdateCheck_IsUnavailableAndSaysWhy()
    {
        var about = _services.GetRequiredService<AboutViewModel>();

        Assert.False(about.IsUpdateCheckConfigured);
        Assert.False(about.CheckForUpdatesCommand.CanExecute(null));
        Assert.NotEmpty(about.UpdateCheckUnavailableReason);
    }
}
