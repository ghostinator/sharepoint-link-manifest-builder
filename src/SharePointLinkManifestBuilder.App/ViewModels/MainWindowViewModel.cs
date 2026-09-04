using System.Collections.ObjectModel;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using Microsoft.Extensions.Logging;
using SharePointLinkManifestBuilder.App.Infrastructure;
using SharePointLinkManifestBuilder.Core.Abstractions;

namespace SharePointLinkManifestBuilder.App.ViewModels;

/// <summary>The application shell: navigation, connection banner and the active page.</summary>
public sealed partial class MainWindowViewModel : ViewModelBase
{
    private readonly ConnectionCoordinator _connection;
    private readonly IProductMetadataProvider _productMetadata;
    private readonly ILogger<MainWindowViewModel> _logger;

    /// <summary>The page currently shown in the content area.</summary>
    [ObservableProperty]
    private PageViewModelBase? _currentPage;

    /// <summary>True once startup initialization has finished.</summary>
    [ObservableProperty]
    private bool _isInitialized;

    /// <summary>Creates the shell.</summary>
    public MainWindowViewModel(
        HomeViewModel home,
        NewLinkJobViewModel job,
        SavedProfilesViewModel profiles,
        JobHistoryViewModel history,
        TenantSetupViewModel setup,
        PermissionsViewModel permissions,
        SettingsViewModel settings,
        DiagnosticsViewModel diagnostics,
        HelpViewModel help,
        AboutViewModel about,
        ConnectionCoordinator connection,
        IProductMetadataProvider productMetadata,
        ILogger<MainWindowViewModel> logger)
    {
        _connection = connection ?? throw new ArgumentNullException(nameof(connection));
        _productMetadata = productMetadata ?? throw new ArgumentNullException(nameof(productMetadata));
        _logger = logger ?? throw new ArgumentNullException(nameof(logger));

        // The two resource browsers are deliberately absent: they are steps inside the job page
        // now, not destinations of their own. Choosing what to process is part of building a
        // job, and having them in the sidebar made it a detour with no way back.
        Pages =
        [
            home, job, profiles, history,
            setup, permissions, settings, diagnostics, help, about,
        ];

        _currentPage = home;
        _connection.ConnectionChanged += (_, _) => RaiseConnectionProperties();
    }

    /// <summary>Every navigable page, in navigation order.</summary>
    public IReadOnlyList<PageViewModelBase> Pages { get; }

    /// <summary>Application title, including version.</summary>
    public string WindowTitle =>
        $"{_productMetadata.Metadata.ProductName} {_productMetadata.Version}";

    /// <summary>A short connection descriptor shown in the shell header.</summary>
    public string ConnectionBadge => _connection.State switch
    {
        ConnectionState.Connected =>
            $"{_connection.TenantDisplayName ?? _connection.Tenant?.TenantId} - "
            + $"{_connection.Account?.UserPrincipalName}",
        ConnectionState.ConnectedWithMissingPermissions => "Connected - permissions missing",
        ConnectionState.PendingAdministratorConsent => "Waiting for administrator approval",
        ConnectionState.ConfiguredSignedOut => "Signed out",
        _ => "Not connected",
    };

    /// <summary>True when the header should draw attention to the connection state.</summary>
    public bool ConnectionNeedsAttention => _connection.State != ConnectionState.Connected;

    /// <summary>A banner shown when setup has not been completed.</summary>
    public string? SetupPrompt => _connection.State == ConnectionState.NotConfigured
        ? "This application is not connected to Microsoft 365 yet. Open Tenant Setup to connect."
        : null;

    /// <summary>True when the setup banner applies.</summary>
    public bool ShowSetupPrompt => SetupPrompt is not null;

    /// <summary>
    /// Runs startup initialization and routes the user to setup when no tenant is configured.
    /// </summary>
    public async Task InitializeAsync(CancellationToken cancellationToken = default)
    {
        IsBusy = true;

        try
        {
            await _connection.InitializeAsync(cancellationToken).ConfigureAwait(true);

            // First run lands on the wizard rather than an empty Home page the user cannot use.
            if (_connection.State == ConnectionState.NotConfigured)
            {
                await NavigateToAsync("setup").ConfigureAwait(true);
            }
            else
            {
                await NavigateToAsync("home").ConfigureAwait(true);
            }
        }
#pragma warning disable CA1031 // Startup must not fail closed with an unhandled exception.
        catch (Exception ex)
        {
            _logger.LogError(ex, "Application initialization failed.");
            ErrorMessage = "The application could not finish starting up. See Diagnostics for details.";
        }
#pragma warning restore CA1031
        finally
        {
            IsBusy = false;
            IsInitialized = true;
            RaiseConnectionProperties();
        }
    }

    /// <summary>Navigates to a page by its stable key.</summary>
    [RelayCommand]
    public async Task NavigateToAsync(string? navigationKey)
    {
        var page = Pages.FirstOrDefault(p =>
            string.Equals(p.NavigationKey, navigationKey, StringComparison.OrdinalIgnoreCase));

        if (page is null)
        {
            return;
        }

        CurrentPage = page;

        try
        {
            await page.OnNavigatedToAsync().ConfigureAwait(true);
        }
#pragma warning disable CA1031 // A page that fails to load must not take the shell down with it.
        catch (Exception ex)
        {
            _logger.LogError(ex, "Loading the {Page} page failed.", page.NavigationKey);
            page.ErrorMessage = "This page could not be loaded. See Diagnostics for details.";
        }
#pragma warning restore CA1031
    }

    /// <summary>Navigates by page instance, used by the navigation list binding.</summary>
    public async Task NavigateToPageAsync(PageViewModelBase? page)
    {
        if (page is not null)
        {
            await NavigateToAsync(page.NavigationKey).ConfigureAwait(true);
        }
    }

    private void RaiseConnectionProperties()
    {
        OnPropertyChanged(nameof(ConnectionBadge));
        OnPropertyChanged(nameof(ConnectionNeedsAttention));
        OnPropertyChanged(nameof(SetupPrompt));
        OnPropertyChanged(nameof(ShowSetupPrompt));
    }
}
