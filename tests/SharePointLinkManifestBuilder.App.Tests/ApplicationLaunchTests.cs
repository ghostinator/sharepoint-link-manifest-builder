using Avalonia;
using Avalonia.Controls;
using Avalonia.Headless;
using Avalonia.Headless.XUnit;
using Microsoft.Extensions.DependencyInjection;
using SharePointLinkManifestBuilder.App.Composition;
using SharePointLinkManifestBuilder.App.ViewModels;
using SharePointLinkManifestBuilder.App.Views;
using SharePointLinkManifestBuilder.Core.Abstractions;
using SharePointLinkManifestBuilder.Core.Settings;

[assembly: AvaloniaTestApplication(typeof(SharePointLinkManifestBuilder.App.Tests.TestAppBuilder))]

namespace SharePointLinkManifestBuilder.App.Tests;

/// <summary>Builds the real application under Avalonia's headless platform.</summary>
public static class TestAppBuilder
{
    /// <summary>Configures the application for headless testing.</summary>
    public static AppBuilder BuildAvaloniaApp() =>
        AppBuilder.Configure<SharePointLinkManifestBuilder.App.App>()
            .UseHeadless(new AvaloniaHeadlessPlatformOptions { UseHeadlessDrawing = true });
}

/// <summary>
/// Proves the application actually starts: the composition root resolves, the shell window is
/// created, and every page renders without throwing.
/// <para>
/// This is the difference between "it compiles" and "it runs". A missing dependency
/// registration, a XAML binding to a property that does not exist, or a view model that throws
/// in its constructor all compile perfectly and fail only at launch.
/// </para>
/// </summary>
public sealed class ApplicationLaunchTests : IDisposable
{
    private readonly string _stateDirectory;
    private readonly ServiceProvider _services;

    /// <summary>Builds a provider rooted at a temporary directory, never the real user profile.</summary>
    public ApplicationLaunchTests()
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

    /// <summary>
    /// The container is built with ValidateOnBuild, so this fails if any registered service has
    /// a dependency that was never registered.
    /// </summary>
    [Fact]
    public void CompositionRoot_ResolvesEveryViewModel()
    {
        Assert.NotNull(_services.GetRequiredService<MainWindowViewModel>());
        Assert.NotNull(_services.GetRequiredService<HomeViewModel>());
        Assert.NotNull(_services.GetRequiredService<NewLinkJobViewModel>());
        Assert.NotNull(_services.GetRequiredService<SharePointBrowserViewModel>());
        Assert.NotNull(_services.GetRequiredService<OneDriveBrowserViewModel>());
        Assert.NotNull(_services.GetRequiredService<SavedProfilesViewModel>());
        Assert.NotNull(_services.GetRequiredService<JobHistoryViewModel>());
        Assert.NotNull(_services.GetRequiredService<TenantSetupViewModel>());
        Assert.NotNull(_services.GetRequiredService<PermissionsViewModel>());
        Assert.NotNull(_services.GetRequiredService<SettingsViewModel>());
        Assert.NotNull(_services.GetRequiredService<DiagnosticsViewModel>());
        Assert.NotNull(_services.GetRequiredService<HelpViewModel>());
        Assert.NotNull(_services.GetRequiredService<AboutViewModel>());
    }

    /// <summary>Every Graph and Core service the application depends on must resolve.</summary>
    [Fact]
    public void CompositionRoot_ResolvesEveryService()
    {
        Assert.NotNull(_services.GetRequiredService<IAuthenticationService>());
        Assert.NotNull(_services.GetRequiredService<IGraphApiClient>());
        Assert.NotNull(_services.GetRequiredService<ISiteService>());
        Assert.NotNull(_services.GetRequiredService<IDriveService>());
        Assert.NotNull(_services.GetRequiredService<IUserDirectoryService>());
        Assert.NotNull(_services.GetRequiredService<ISharingLinkService>());
        Assert.NotNull(_services.GetRequiredService<IManifestStorageService>());
        Assert.NotNull(_services.GetRequiredService<ILinkJobRunner>());
        Assert.NotNull(_services.GetRequiredService<IFileDiscoveryService>());
        Assert.NotNull(_services.GetRequiredService<IAppRegistrationService>());
        Assert.NotNull(_services.GetRequiredService<IConsentService>());
        Assert.NotNull(_services.GetRequiredService<IDiagnosticsService>());
        Assert.NotNull(_services.GetRequiredService<ISecureTokenStorage>());
    }

    /// <summary>All four manifest formatters must be registered, or a format silently produces nothing.</summary>
    [Fact]
    public void CompositionRoot_RegistersEveryManifestFormatter()
    {
        var formatters = _services.GetServices<IManifestFormatter>().ToArray();

        Assert.Equal(4, formatters.Length);
        Assert.Contains(formatters, f => f.FileExtension == ".txt");
        Assert.Contains(formatters, f => f.FileExtension == ".md");
        Assert.Contains(formatters, f => f.FileExtension == ".csv");
        Assert.Contains(formatters, f => f.FileExtension == ".json");
    }

    /// <summary>The shell window must construct and bind without throwing.</summary>
    [AvaloniaFact]
    public void MainWindow_IsCreatedAndBound()
    {
        var viewModel = _services.GetRequiredService<MainWindowViewModel>();
        var window = new MainWindow { DataContext = viewModel };

        window.Show();

        Assert.NotNull(window.DataContext);
        Assert.Equal(12, viewModel.Pages.Count);
        Assert.Contains(viewModel.Pages, p => p.NavigationKey == "setup");
        Assert.False(string.IsNullOrWhiteSpace(window.Title));
    }

    /// <summary>
    /// Renders every page in turn. A XAML binding to a property that does not exist, or a
    /// converter that throws, surfaces here rather than in front of a user.
    /// </summary>
    [AvaloniaFact]
    public void EveryPage_RendersWithoutThrowing()
    {
        var viewModel = _services.GetRequiredService<MainWindowViewModel>();
        var window = new MainWindow { DataContext = viewModel };

        window.Show();

        foreach (var page in viewModel.Pages)
        {
            viewModel.CurrentPage = page;

            // Forces a layout pass, which is what actually evaluates the bindings.
            window.Measure(new Size(1280, 840));
            window.Arrange(new Rect(0, 0, 1280, 840));

            Assert.Equal(page, viewModel.CurrentPage);
        }
    }

    /// <summary>Navigation by key must reach each page and run its load hook.</summary>
    [AvaloniaFact]
    public async Task NavigateTo_ReachesEveryPageByKey()
    {
        var viewModel = _services.GetRequiredService<MainWindowViewModel>();
        var window = new MainWindow { DataContext = viewModel };

        window.Show();

        foreach (var key in new[]
        {
            "home", "job", "sharepoint", "onedrive", "profiles", "history",
            "setup", "permissions", "settings", "diagnostics", "help", "about",
        })
        {
            await viewModel.NavigateToAsync(key);

            Assert.NotNull(viewModel.CurrentPage);
            Assert.Equal(key, viewModel.CurrentPage!.NavigationKey);

            // A page that fails to load records an error rather than throwing, so assert the
            // pages load cleanly in a default, unconfigured state.
            Assert.False(
                viewModel.CurrentPage.HasError,
                $"Page '{key}' reported an error on load: {viewModel.CurrentPage.ErrorMessage}");
        }
    }

    /// <summary>
    /// With no tenant configured, startup must land on the setup wizard rather than an empty
    /// Home page the user cannot do anything with.
    /// </summary>
    [AvaloniaFact]
    public async Task Initialize_WithNoTenantConfigured_OpensTheSetupWizard()
    {
        var viewModel = _services.GetRequiredService<MainWindowViewModel>();
        var window = new MainWindow { DataContext = viewModel };

        window.Show();

        await viewModel.InitializeAsync();

        Assert.True(viewModel.IsInitialized);
        Assert.Equal("setup", viewModel.CurrentPage?.NavigationKey);
        Assert.True(viewModel.ShowSetupPrompt);
        Assert.Equal("Not connected", viewModel.ConnectionBadge);
    }

    /// <summary>
    /// The wizard must be usable end to end without a tenant, and must state plainly that
    /// automatic setup is unavailable rather than offering a path that cannot work.
    /// </summary>
    [AvaloniaFact]
    public void SetupWizard_ReportsAutomaticSetupUnavailableWithoutABootstrapClientId()
    {
        var wizard = _services.GetRequiredService<TenantSetupViewModel>();

        Assert.False(wizard.IsAutomaticSetupAvailable);
        Assert.Contains("bootstrap client ID", wizard.AutomaticSetupUnavailableReason, StringComparison.OrdinalIgnoreCase);

        // The existing-registration path stays fully available.
        Assert.Equal(SetupMethod.ExistingRegistration, wizard.Method);
        Assert.NotEmpty(wizard.RequestedPermissions);
        Assert.NotEmpty(wizard.PlannedChanges);
    }

    /// <summary>The wizard's permission list must react to the least-privilege options.</summary>
    [AvaloniaFact]
    public void SetupWizard_PermissionListRespondsToOptions()
    {
        var wizard = _services.GetRequiredService<TenantSetupViewModel>();

        wizard.ReadOnlyMode = true;
        Assert.Contains(wizard.RequestedPermissions, p => p.Scope == "Files.Read.All");
        Assert.DoesNotContain(wizard.RequestedPermissions, p => p.Scope == "Files.ReadWrite.All");

        wizard.ReadOnlyMode = false;
        Assert.Contains(wizard.RequestedPermissions, p => p.Scope == "Files.ReadWrite.All");

        wizard.IncludeUserPicker = true;
        Assert.Contains(wizard.RequestedPermissions, p => p.Scope == "User.ReadBasic.All");

        wizard.IncludeBroadSharePointWrite = true;
        Assert.Contains(wizard.RequestedPermissions, p => p.Scope == "Sites.ReadWrite.All");
    }
}
