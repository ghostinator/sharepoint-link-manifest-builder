using System.Collections.ObjectModel;
using System.Globalization;
using System.Runtime.InteropServices;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using Microsoft.Extensions.Logging;
using SharePointLinkManifestBuilder.App.Infrastructure;
using SharePointLinkManifestBuilder.Core.Abstractions;
using SharePointLinkManifestBuilder.Core.Models;
using SharePointLinkManifestBuilder.Core.Settings;

namespace SharePointLinkManifestBuilder.App.ViewModels;

/// <summary>The landing page: connection status, warnings and shortcuts.</summary>
public sealed partial class HomeViewModel : PageViewModelBase
{
    private readonly ConnectionCoordinator _connection;
    private readonly IJobHistoryStore _historyStore;
    private readonly ISettingsStore _settingsStore;
    private readonly IProductMetadataProvider _productMetadata;
    private readonly ISecureTokenStorage _tokenStorage;

    /// <summary>Creates the page.</summary>
    public HomeViewModel(
        ConnectionCoordinator connection,
        IJobHistoryStore historyStore,
        ISettingsStore settingsStore,
        IProductMetadataProvider productMetadata,
        ISecureTokenStorage tokenStorage)
        : base("Home", "home")
    {
        _connection = connection ?? throw new ArgumentNullException(nameof(connection));
        _historyStore = historyStore ?? throw new ArgumentNullException(nameof(historyStore));
        _settingsStore = settingsStore ?? throw new ArgumentNullException(nameof(settingsStore));
        _productMetadata = productMetadata ?? throw new ArgumentNullException(nameof(productMetadata));
        _tokenStorage = tokenStorage ?? throw new ArgumentNullException(nameof(tokenStorage));

        _connection.ConnectionChanged += (_, _) => RaiseAll();
    }

    /// <summary>The most recent jobs.</summary>
    public ObservableCollection<JobHistoryEntry> RecentJobs { get; } = [];

    /// <summary>Locations used recently.</summary>
    public ObservableCollection<RecentLocation> RecentLocations { get; } = [];

    /// <summary>The current connection state.</summary>
    public ConnectionState State => _connection.State;

    /// <summary>A one-line summary of the connection.</summary>
    public string ConnectionSummary => _connection.State switch
    {
        ConnectionState.NotConfigured =>
            "Not connected. Run Tenant Setup to connect to Microsoft 365.",
        ConnectionState.ConfiguredSignedOut =>
            "Configured, but you are signed out. Sign in to continue.",
        ConnectionState.PendingAdministratorConsent =>
            "Waiting for a Microsoft Entra administrator to approve the requested permissions.",
        ConnectionState.ConnectedWithMissingPermissions =>
            "Connected, but some required permissions have not been granted.",
        ConnectionState.Connected =>
            $"Connected to {_connection.TenantDisplayName ?? _connection.Tenant?.TenantId} "
            + $"as {_connection.Account?.UserPrincipalName}.",
        _ => "Unknown state.",
    };

    /// <summary>True when the connection is usable.</summary>
    public bool IsConnected => _connection.State == ConnectionState.Connected;

    /// <summary>True when something needs the user's attention.</summary>
    public bool NeedsAttention =>
        _connection.State is ConnectionState.NotConfigured
            or ConnectionState.ConfiguredSignedOut
            or ConnectionState.PendingAdministratorConsent
            or ConnectionState.ConnectedWithMissingPermissions;

    /// <summary>Tenant display name or ID.</summary>
    public string TenantDisplay =>
        _connection.TenantDisplayName ?? _connection.Tenant?.TenantId ?? "(none)";

    /// <summary>Signed-in account.</summary>
    public string AccountDisplay => _connection.Account?.UserPrincipalName ?? "(not signed in)";

    /// <summary>Registration source.</summary>
    public string RegistrationDisplay => _connection.Tenant?.Source.ToString() ?? "None";

    /// <summary>Consent state.</summary>
    public string ConsentDisplay => _connection.Tenant?.ConsentState.ToString() ?? "Unknown";

    /// <summary>Scopes required but not granted.</summary>
    public IReadOnlyList<string> MissingScopes => _connection.Tenant?.MissingScopes ?? [];

    /// <summary>True when a permission warning should be shown.</summary>
    public bool HasMissingScopes => MissingScopes.Count > 0;

    /// <summary>Warning shown when tokens are memory-only.</summary>
    public string? SecureStorageWarning =>
        _tokenStorage.Status.Availability == SecureStorageAvailability.UnavailableUsingMemoryOnly
            ? _tokenStorage.Status.Detail
            : null;

    /// <summary>True when the secure-storage warning applies.</summary>
    public bool HasSecureStorageWarning => SecureStorageWarning is not null;

    /// <summary>Application version.</summary>
    public string Version => _productMetadata.Version;

    /// <inheritdoc />
    public override async Task OnNavigatedToAsync(CancellationToken cancellationToken = default)
    {
        RaiseAll();

        RecentJobs.Clear();

        foreach (var job in (await _historyStore.ListAsync(cancellationToken).ConfigureAwait(true)).Take(5))
        {
            RecentJobs.Add(job);
        }

        var settings = await _settingsStore.LoadAsync(cancellationToken).ConfigureAwait(true);

        RecentLocations.Clear();

        foreach (var location in settings.RecentLocations.Take(5))
        {
            RecentLocations.Add(location);
        }
    }

    /// <summary>Signs in through the system browser.</summary>
    [RelayCommand]
    private async Task SignInAsync(CancellationToken cancellationToken)
    {
        IsBusy = true;
        ClearMessages();

        try
        {
            var result = await _connection.SignInAsync(cancellationToken).ConfigureAwait(true);

            if (!result.Succeeded)
            {
                ErrorMessage = result.Error?.Message;
            }
        }
        finally
        {
            IsBusy = false;
            RaiseAll();
        }
    }

    private void RaiseAll()
    {
        OnPropertyChanged(nameof(State));
        OnPropertyChanged(nameof(ConnectionSummary));
        OnPropertyChanged(nameof(IsConnected));
        OnPropertyChanged(nameof(NeedsAttention));
        OnPropertyChanged(nameof(TenantDisplay));
        OnPropertyChanged(nameof(AccountDisplay));
        OnPropertyChanged(nameof(RegistrationDisplay));
        OnPropertyChanged(nameof(ConsentDisplay));
        OnPropertyChanged(nameof(MissingScopes));
        OnPropertyChanged(nameof(HasMissingScopes));
        OnPropertyChanged(nameof(SecureStorageWarning));
        OnPropertyChanged(nameof(HasSecureStorageWarning));
    }
}

/// <summary>Application settings, including the privacy and local-data controls.</summary>
public sealed partial class SettingsViewModel : PageViewModelBase
{
    private readonly ISettingsStore _settingsStore;
    private readonly ISecureTokenStorage _tokenStorage;
    private readonly ConnectionCoordinator _connection;
    private readonly IDiagnosticsService _diagnostics;
    private readonly ApplicationPaths _paths;
    private readonly FolderLauncher _folders;

    /// <summary>Theme preference.</summary>
    [ObservableProperty]
    private ThemePreference _theme = ThemePreference.System;

    /// <summary>Honour the operating system's reduced-motion setting.</summary>
    [ObservableProperty]
    private bool _reduceMotion;

    /// <summary>Telemetry opt-in. Off by default and not implemented as a live pipeline.</summary>
    [ObservableProperty]
    private bool _telemetryEnabled;

    /// <summary>Default maximum concurrency for new jobs.</summary>
    [ObservableProperty]
    private int _defaultMaxConcurrency = 4;

    /// <summary>Default request delay for new jobs, in milliseconds.</summary>
    [ObservableProperty]
    private int _defaultRequestDelayMilliseconds;

    /// <summary>Default retry limit for new jobs.</summary>
    [ObservableProperty]
    private int _defaultMaxRetryAttempts = 5;

    /// <summary>How many history entries to keep. Zero means unlimited.</summary>
    [ObservableProperty]
    private int _historyRetentionCount = 100;

    /// <summary>Minimum level written to the log file.</summary>
    [ObservableProperty]
    private string _logLevel = "Information";

    /// <summary>Creates the page.</summary>
    public SettingsViewModel(
        ISettingsStore settingsStore,
        ISecureTokenStorage tokenStorage,
        ConnectionCoordinator connection,
        IDiagnosticsService diagnostics,
        ApplicationPaths paths,
        FolderLauncher folders)
        : base("Settings", "settings")
    {
        _settingsStore = settingsStore ?? throw new ArgumentNullException(nameof(settingsStore));
        _tokenStorage = tokenStorage ?? throw new ArgumentNullException(nameof(tokenStorage));
        _connection = connection ?? throw new ArgumentNullException(nameof(connection));
        _diagnostics = diagnostics ?? throw new ArgumentNullException(nameof(diagnostics));
        _paths = paths ?? throw new ArgumentNullException(nameof(paths));
        _folders = folders ?? throw new ArgumentNullException(nameof(folders));
    }

    /// <summary>Available log levels.</summary>
    public static IReadOnlyList<string> LogLevels { get; } =
        ["Trace", "Debug", "Information", "Warning", "Error"];

    /// <summary>Available themes.</summary>
    public static IReadOnlyList<ThemePreference> Themes { get; } = Enum.GetValues<ThemePreference>();

    /// <summary>Where local data is stored, shown so the answer is discoverable in the UI.</summary>
    public IReadOnlyList<StorageLocationInfo> StorageLocations => _paths.Describe();

    /// <summary>How tokens are protected on this machine.</summary>
    public string SecureStorageStatus =>
        $"{_tokenStorage.Status.Availability} ({_tokenStorage.Status.Mechanism})";

    /// <summary>Detail shown when secure storage is unavailable.</summary>
    public string? SecureStorageDetail => _tokenStorage.Status.Detail;

    /// <summary>The privacy statement shown in Settings and on first run.</summary>
    public static IReadOnlyList<string> PrivacyPoints =>
    [
        "This application sends requests only to Microsoft identity and Microsoft Graph endpoints. It "
        + "contacts no other service.",

        "Telemetry is disabled by default and no telemetry pipeline is implemented. No filename, URL, "
        + "tenant identifier, identity or sharing link is ever transmitted anywhere.",

        "Sign-in details are stored by the operating system's secure store, never in the settings file.",

        "Job history, profiles and the tenant audit trail stay on this machine and can be cleared at any "
        + "time from this page.",

        "Diagnostic bundles are built from an explicit allow-list and never contain tokens, secrets or "
        + "sharing links.",
    ];

    /// <inheritdoc />
    public override async Task OnNavigatedToAsync(CancellationToken cancellationToken = default)
    {
        var settings = await _settingsStore.LoadAsync(cancellationToken).ConfigureAwait(true);

        Theme = settings.Theme;
        ReduceMotion = settings.ReduceMotion;
        TelemetryEnabled = settings.TelemetryEnabled;
        DefaultMaxConcurrency = settings.DefaultExecution.MaxConcurrency;
        DefaultRequestDelayMilliseconds = (int)settings.DefaultExecution.RequestDelay.TotalMilliseconds;
        DefaultMaxRetryAttempts = settings.DefaultExecution.MaxRetryAttempts;
        HistoryRetentionCount = settings.JobHistoryRetentionCount;
        LogLevel = settings.LogLevel;

        OnPropertyChanged(nameof(SecureStorageStatus));
        OnPropertyChanged(nameof(SecureStorageDetail));
    }

    /// <summary>Saves the settings.</summary>
    [RelayCommand]
    private async Task SaveAsync(CancellationToken cancellationToken)
    {
        var existing = await _settingsStore.LoadAsync(cancellationToken).ConfigureAwait(true);

        await _settingsStore.SaveAsync(
            existing with
            {
                Theme = Theme,
                ReduceMotion = ReduceMotion,
                TelemetryEnabled = TelemetryEnabled,
                JobHistoryRetentionCount = HistoryRetentionCount,
                LogLevel = LogLevel,
                DefaultExecution = existing.DefaultExecution with
                {
                    MaxConcurrency = DefaultMaxConcurrency,
                    RequestDelay = TimeSpan.FromMilliseconds(DefaultRequestDelayMilliseconds),
                    MaxRetryAttempts = DefaultMaxRetryAttempts,
                },
            },
            cancellationToken).ConfigureAwait(true);

        StatusMessage = "Settings saved. The log level takes effect next time the application starts.";
    }

    /// <summary>Signs out and forgets the cached account.</summary>
    [RelayCommand]
    private async Task ForgetAccountAsync(CancellationToken cancellationToken)
    {
        await _connection.SignOutAsync(cancellationToken).ConfigureAwait(true);
        StatusMessage = "Signed out and removed the cached account.";
    }

    /// <summary>Clears the token cache entirely.</summary>
    [RelayCommand]
    private async Task ClearTokenCacheAsync(CancellationToken cancellationToken)
    {
        // Sign out first so each account is removed individually, which is what Microsoft
        // recommends; the cache wipe then just removes any residue.
        await _connection.SignOutAsync(cancellationToken).ConfigureAwait(true);
        await _tokenStorage.ClearAsync(cancellationToken).ConfigureAwait(true);

        StatusMessage = "Token cache cleared. You will need to sign in again.";
    }

    /// <summary>Clears cached data without touching sign-in details.</summary>
    [RelayCommand]
    private async Task ClearCachedDataAsync(CancellationToken cancellationToken)
    {
        await _diagnostics.ClearCacheAsync(cancellationToken).ConfigureAwait(true);
        StatusMessage = "Cached data cleared. Sign-in details were not affected.";
    }

    /// <summary>Removes the local tenant configuration. Nothing in the tenant is changed.</summary>
    [RelayCommand]
    private async Task RemoveTenantConfigurationAsync(CancellationToken cancellationToken)
    {
        await _connection.RemoveLocalConfigurationAsync(cancellationToken).ConfigureAwait(true);
        StatusMessage = "Local Microsoft 365 configuration removed. Your tenant was not changed.";
    }

    /// <summary>Opens the local data folder.</summary>
    [RelayCommand]
    private void OpenDataFolder() => _folders.OpenFolder(_paths.RootDirectory);
}

/// <summary>Diagnostics: environment, connectivity, and the sanitized bundle export.</summary>
public sealed partial class DiagnosticsViewModel : PageViewModelBase
{
    private readonly IDiagnosticsService _diagnostics;
    private readonly ConnectionCoordinator _connection;
    private readonly ISecureTokenStorage _tokenStorage;
    private readonly IJobHistoryStore _historyStore;
    private readonly IProductMetadataProvider _productMetadata;
    private readonly ApplicationPaths _paths;
    private readonly FolderLauncher _folders;

    /// <summary>Result of the last connectivity test.</summary>
    [ObservableProperty]
    private string? _connectivityResult;

    /// <summary>Include file and folder names in the bundle. Off by default.</summary>
    [ObservableProperty]
    private bool _includeFileNames;

    /// <summary>Include email addresses in the bundle. Off by default.</summary>
    [ObservableProperty]
    private bool _includeEmailAddresses;

    /// <summary>Include full tenant-specific URLs in the bundle. Off by default.</summary>
    [ObservableProperty]
    private bool _includeFullUrls;

    /// <summary>Where the last bundle was written.</summary>
    [ObservableProperty]
    private string? _lastBundlePath;

    /// <summary>Creates the page.</summary>
    public DiagnosticsViewModel(
        IDiagnosticsService diagnostics,
        ConnectionCoordinator connection,
        ISecureTokenStorage tokenStorage,
        IJobHistoryStore historyStore,
        IProductMetadataProvider productMetadata,
        ApplicationPaths paths,
        FolderLauncher folders)
        : base("Diagnostics", "diagnostics")
    {
        _diagnostics = diagnostics ?? throw new ArgumentNullException(nameof(diagnostics));
        _connection = connection ?? throw new ArgumentNullException(nameof(connection));
        _tokenStorage = tokenStorage ?? throw new ArgumentNullException(nameof(tokenStorage));
        _historyStore = historyStore ?? throw new ArgumentNullException(nameof(historyStore));
        _productMetadata = productMetadata ?? throw new ArgumentNullException(nameof(productMetadata));
        _paths = paths ?? throw new ArgumentNullException(nameof(paths));
        _folders = folders ?? throw new ArgumentNullException(nameof(folders));
    }

    /// <summary>Sanitized errors from recent jobs.</summary>
    public ObservableCollection<string> RecentErrors { get; } = [];

    /// <summary>Application version.</summary>
    public string Version => _productMetadata.Version;

    /// <summary>Operating system description.</summary>
    public static string Platform => RuntimeInformation.OSDescription;

    /// <summary>Process architecture.</summary>
    public static string Architecture => RuntimeInformation.OSArchitecture.ToString();

    /// <summary>.NET runtime version.</summary>
    public static string RuntimeVersion => RuntimeInformation.FrameworkDescription;

    /// <summary>Connection state.</summary>
    public string ConnectionStateDisplay => _connection.State.ToString();

    /// <summary>Registration source.</summary>
    public string RegistrationDisplay => _connection.Tenant?.Source.ToString() ?? "None";

    /// <summary>Consent state.</summary>
    public string ConsentDisplay => _connection.Tenant?.ConsentState.ToString() ?? "Unknown";

    /// <summary>Secure-storage availability.</summary>
    public string SecureStorageDisplay =>
        $"{_tokenStorage.Status.Availability} ({_tokenStorage.Status.Mechanism})";

    /// <summary>Summary of the most recent job.</summary>
    public string LastJobSummary { get; private set; } = "No jobs have been run.";

    /// <summary>Categories always included in a bundle.</summary>
    public static IReadOnlyList<string> AlwaysIncluded => DiagnosticBundleMetadata.AlwaysIncluded;

    /// <summary>Categories never included, under any option.</summary>
    public static IReadOnlyList<string> NeverIncluded => DiagnosticBundleMetadata.NeverIncluded;

    /// <inheritdoc />
    public override async Task OnNavigatedToAsync(CancellationToken cancellationToken = default)
    {
        var history = await _historyStore.ListAsync(cancellationToken).ConfigureAwait(true);
        var latest = history.Count > 0 ? history[0] : null;

        LastJobSummary = latest is null
            ? "No jobs have been run."
            : $"{latest.StartedUtc.UtcDateTime.ToString("u", CultureInfo.InvariantCulture)} - "
              + $"{latest.FinalPhase}, {latest.CreatedCount} created, {latest.ReusedCount} reused, "
              + $"{latest.SkippedCount} skipped, {latest.FailedCount} failed";

        RecentErrors.Clear();

        foreach (var error in history.SelectMany(h => h.SanitizedErrors).Distinct(StringComparer.Ordinal).Take(50))
        {
            RecentErrors.Add(error);
        }

        OnPropertyChanged(nameof(LastJobSummary));
        OnPropertyChanged(nameof(ConnectionStateDisplay));
        OnPropertyChanged(nameof(RegistrationDisplay));
        OnPropertyChanged(nameof(ConsentDisplay));
        OnPropertyChanged(nameof(SecureStorageDisplay));
    }

    /// <summary>Runs a live Graph connectivity test.</summary>
    [RelayCommand]
    private async Task TestConnectivityAsync(CancellationToken cancellationToken)
    {
        IsBusy = true;

        try
        {
            var result = await _diagnostics.TestGraphConnectivityAsync(cancellationToken).ConfigureAwait(true);

            ConnectivityResult = result.Succeeded
                ? $"Microsoft Graph responded in {result.Value.TotalMilliseconds:0} ms."
                : $"Not reachable: {result.Error!.Message}";
        }
        finally
        {
            IsBusy = false;
        }
    }

    /// <summary>Writes a sanitized diagnostic bundle.</summary>
    [RelayCommand]
    private async Task ExportBundleAsync(CancellationToken cancellationToken)
    {
        IsBusy = true;
        ClearMessages();

        try
        {
            var metadata = new DiagnosticBundleMetadata
            {
                ApplicationVersion = Version,
                Platform = Platform,
                RuntimeVersion = RuntimeVersion,
                IncludeFileNames = IncludeFileNames,
                IncludeEmailAddresses = IncludeEmailAddresses,
                IncludeFullUrls = IncludeFullUrls,
            };

            var fileName = $"diagnostics-{DateTimeOffset.UtcNow:yyyyMMdd-HHmmss}.zip";
            var destination = Path.Combine(_paths.ExportsDirectory, fileName);

            var result = await _diagnostics.ExportBundleAsync(metadata, destination, cancellationToken)
                .ConfigureAwait(true);

            if (result.Succeeded)
            {
                LastBundlePath = result.Value;
                StatusMessage = $"Diagnostic bundle written to {result.Value}";
            }
            else
            {
                ErrorMessage = result.Error!.Message;
            }
        }
        finally
        {
            IsBusy = false;
        }
    }

    /// <summary>Opens the log folder.</summary>
    [RelayCommand]
    private void OpenLogFolder() => _folders.OpenFolder(_paths.LogDirectory);

    /// <summary>Opens the folder holding exported bundles.</summary>
    [RelayCommand]
    private void OpenExportsFolder() => _folders.OpenFolder(_paths.ExportsDirectory);

    /// <summary>Clears cached data.</summary>
    [RelayCommand]
    private async Task ClearCacheAsync(CancellationToken cancellationToken)
    {
        await _diagnostics.ClearCacheAsync(cancellationToken).ConfigureAwait(true);
        StatusMessage = "Cached data cleared.";
    }

    /// <summary>Clears the job history.</summary>
    [RelayCommand]
    private async Task ClearHistoryAsync(CancellationToken cancellationToken)
    {
        await _historyStore.ClearAsync(cancellationToken).ConfigureAwait(true);
        RecentErrors.Clear();
        LastJobSummary = "No jobs have been run.";
        OnPropertyChanged(nameof(LastJobSummary));
        StatusMessage = "Job history cleared.";
    }
}

/// <summary>Job history, with rerun and retry.</summary>
public sealed partial class JobHistoryViewModel : PageViewModelBase
{
    private readonly IJobHistoryStore _historyStore;
    private readonly JobDraft _draft;
    private readonly IClipboardService _clipboard;

    /// <summary>The history entry currently selected.</summary>
    [ObservableProperty]
    private JobHistoryEntry? _selectedEntry;

    /// <summary>Creates the page.</summary>
    public JobHistoryViewModel(IJobHistoryStore historyStore, JobDraft draft, IClipboardService clipboard)
        : base("Job History", "history")
    {
        _historyStore = historyStore ?? throw new ArgumentNullException(nameof(historyStore));
        _draft = draft ?? throw new ArgumentNullException(nameof(draft));
        _clipboard = clipboard ?? throw new ArgumentNullException(nameof(clipboard));
    }

    /// <summary>Past runs, newest first.</summary>
    public ObservableCollection<JobHistoryEntry> Entries { get; } = [];

    /// <inheritdoc />
    public override async Task OnNavigatedToAsync(CancellationToken cancellationToken = default) =>
        await ReloadAsync(cancellationToken).ConfigureAwait(true);

    /// <summary>Reloads the history.</summary>
    [RelayCommand]
    private async Task ReloadAsync(CancellationToken cancellationToken)
    {
        Entries.Clear();

        foreach (var entry in await _historyStore.ListAsync(cancellationToken).ConfigureAwait(true))
        {
            Entries.Add(entry);
        }
    }

    /// <summary>Loads a past run's configuration into the job page.</summary>
    [RelayCommand]
    private void RerunConfiguration()
    {
        if (SelectedEntry?.Configuration is not { } configuration)
        {
            ErrorMessage = "This entry does not include a reusable configuration.";
            return;
        }

        _draft.LoadFrom(configuration);
        StatusMessage = "Configuration loaded into New Link Job. Review it, then preview before running.";
    }

    /// <summary>Copies a sanitized report for the selected run.</summary>
    [RelayCommand]
    private async Task ExportReportAsync()
    {
        if (SelectedEntry is not { } entry)
        {
            return;
        }

        var lines = new List<string>
        {
            "SharePoint Link Manifest Builder - job report",
            $"Job ID          : {entry.JobId}",
            $"Application     : {entry.ApplicationVersion}",
            $"Started         : {entry.StartedUtc:O}",
            $"Completed       : {entry.CompletedUtc:O}",
            $"Final phase     : {entry.FinalPhase}",
            $"Dry run         : {entry.WasDryRun}",
            $"Cancelled       : {entry.WasCancelled}",
            $"Account         : {entry.AccountIdentifier}",
            $"Tenant          : {entry.TenantDisplayName ?? entry.TenantId}",
            $"Link settings   : {entry.LinkSettingsSummary}",
            $"Manifests       : {entry.ManifestSettingsSummary}",
            $"Filters         : {entry.FilterSummary}",
            $"Created         : {entry.CreatedCount}",
            $"Reused          : {entry.ReusedCount}",
            $"Skipped         : {entry.SkippedCount}",
            $"Failed          : {entry.FailedCount}",
            string.Empty,
            "Targets:",
        };

        lines.AddRange(entry.TargetDescriptions.Select(t => $"  - {t}"));
        lines.Add(string.Empty);
        lines.Add("Manifest locations:");
        lines.AddRange(entry.ManifestLocations.Select(m => $"  - {m}"));
        lines.Add(string.Empty);
        lines.Add("Sanitized errors:");
        lines.AddRange(entry.SanitizedErrors.Select(e => $"  - {e}"));

        await _clipboard.SetTextAsync(string.Join(Environment.NewLine, lines)).ConfigureAwait(true);
        StatusMessage = "Report copied to the clipboard.";
    }

    /// <summary>Deletes the selected history entry.</summary>
    [RelayCommand]
    private async Task DeleteEntryAsync(CancellationToken cancellationToken)
    {
        if (SelectedEntry is not { } entry)
        {
            return;
        }

        await _historyStore.DeleteAsync(entry.JobId, cancellationToken).ConfigureAwait(true);
        await ReloadAsync(cancellationToken).ConfigureAwait(true);
        StatusMessage = "History entry deleted.";
    }

    /// <summary>Clears the whole history.</summary>
    [RelayCommand]
    private async Task ClearAllAsync(CancellationToken cancellationToken)
    {
        await _historyStore.ClearAsync(cancellationToken).ConfigureAwait(true);
        Entries.Clear();
        StatusMessage = "Job history cleared.";
    }
}

/// <summary>Saved job profiles.</summary>
public sealed partial class SavedProfilesViewModel : PageViewModelBase
{
    private readonly IProfileStore _profileStore;
    private readonly JobDraft _draft;
    private readonly ConnectionCoordinator _connection;

    /// <summary>The profile currently selected.</summary>
    [ObservableProperty]
    private SavedProfile? _selectedProfile;

    /// <summary>Name for a new profile.</summary>
    [ObservableProperty]
    private string _newProfileName = string.Empty;

    /// <summary>Description for a new profile.</summary>
    [ObservableProperty]
    private string _newProfileDescription = string.Empty;

    /// <summary>Creates the page.</summary>
    public SavedProfilesViewModel(
        IProfileStore profileStore,
        JobDraft draft,
        ConnectionCoordinator connection)
        : base("Saved Profiles", "profiles")
    {
        _profileStore = profileStore ?? throw new ArgumentNullException(nameof(profileStore));
        _draft = draft ?? throw new ArgumentNullException(nameof(draft));
        _connection = connection ?? throw new ArgumentNullException(nameof(connection));
    }

    /// <summary>The saved profiles.</summary>
    public ObservableCollection<SavedProfile> Profiles { get; } = [];

    /// <inheritdoc />
    public override async Task OnNavigatedToAsync(CancellationToken cancellationToken = default) =>
        await ReloadAsync(cancellationToken).ConfigureAwait(true);

    /// <summary>Reloads the profile list.</summary>
    [RelayCommand]
    private async Task ReloadAsync(CancellationToken cancellationToken)
    {
        Profiles.Clear();

        foreach (var profile in await _profileStore.ListAsync(cancellationToken).ConfigureAwait(true))
        {
            Profiles.Add(profile);
        }
    }

    /// <summary>Saves the current job draft as a profile.</summary>
    [RelayCommand]
    private async Task SaveCurrentAsync(CancellationToken cancellationToken)
    {
        if (string.IsNullOrWhiteSpace(NewProfileName))
        {
            ErrorMessage = "Give the profile a name.";
            return;
        }

        if (_connection.Tenant is null)
        {
            ErrorMessage = "Connect to Microsoft 365 first.";
            return;
        }

        var profile = new SavedProfile
        {
            Name = NewProfileName.Trim(),
            Description = string.IsNullOrWhiteSpace(NewProfileDescription) ? null : NewProfileDescription.Trim(),
            Configuration = _draft.ToConfiguration(
                _connection.Tenant.TenantId,
                _connection.TenantDisplayName ?? _connection.Tenant.TenantDisplayName),
        };

        await _profileStore.SaveAsync(profile, cancellationToken).ConfigureAwait(true);
        await ReloadAsync(cancellationToken).ConfigureAwait(true);

        NewProfileName = string.Empty;
        NewProfileDescription = string.Empty;
        StatusMessage = "Profile saved.";
    }

    /// <summary>Loads the selected profile into the job page.</summary>
    [RelayCommand]
    private void LoadProfile()
    {
        if (SelectedProfile is null)
        {
            return;
        }

        _draft.LoadFrom(SelectedProfile.Configuration);
        StatusMessage = $"Loaded '{SelectedProfile.Name}'. Review it, then preview before running.";
    }

    /// <summary>Deletes the selected profile.</summary>
    [RelayCommand]
    private async Task DeleteProfileAsync(CancellationToken cancellationToken)
    {
        if (SelectedProfile is null)
        {
            return;
        }

        await _profileStore.DeleteAsync(SelectedProfile.ProfileId, cancellationToken).ConfigureAwait(true);
        await ReloadAsync(cancellationToken).ConfigureAwait(true);
        StatusMessage = "Profile deleted.";
    }
}

/// <summary>Help: concepts, link types, and troubleshooting pointers.</summary>
public sealed class HelpViewModel : PageViewModelBase
{
    private readonly ISystemBrowser _browser;
    private readonly IProductMetadataProvider _productMetadata;

    /// <summary>Creates the page.</summary>
    public HelpViewModel(ISystemBrowser browser, IProductMetadataProvider productMetadata)
        : base("Help", "help")
    {
        _browser = browser ?? throw new ArgumentNullException(nameof(browser));
        _productMetadata = productMetadata ?? throw new ArgumentNullException(nameof(productMetadata));
    }

    /// <summary>Explanations of each link permission and audience.</summary>
    public static IReadOnlyList<HelpTopic> LinkTypes =>
    [
        new HelpTopic("View link",
            "Recipients can open and read the file but not change it."),

        new HelpTopic("Edit link",
            "Recipients can open and change the file. Use this only when collaboration is intended."),

        new HelpTopic("People in the organization",
            "Anyone signed in to your Microsoft 365 organization can use the link. This is usually the "
            + "right choice for a Copilot manifest, because it keeps access inside the tenant."),

        new HelpTopic("Specific people",
            "Only the people you name can use the link. Microsoft Graph creates the link and, separately, "
            + "grants each named recipient access. No email is sent unless you explicitly ask for one."),

        new HelpTopic("Anyone with the link",
            "Usable without signing in, including outside your organization. Many organizations disable "
            + "this. If yours does, the request is refused and reported as blocked rather than silently "
            + "downgraded."),
    ];

    /// <summary>How recursion and manifest modes interact.</summary>
    public static IReadOnlyList<HelpTopic> Concepts =>
    [
        new HelpTopic("Recursive vs non-recursive",
            "A non-recursive target processes only the files directly inside the folder you selected. A "
            + "recursive target also processes every subfolder beneath it. This is set per target, so one "
            + "job can mix both."),

        new HelpTopic("Per-folder manifest",
            "One manifest is written into each folder that contained processed files, listing the files "
            + "in that folder."),

        new HelpTopic("Master manifest",
            "One manifest lists every successful file beneath the target. For a folder target it is "
            + "written into the starting folder; for a library it goes at the library root. A site "
            + "spanning several libraries needs you to choose a destination, because guessing one could "
            + "put data somewhere unrelated."),

        new HelpTopic("Dry run",
            "Enumerates, filters and validates without creating a single link or writing a single file. "
            + "Always the safest first step."),

        new HelpTopic("Reused links",
            "Microsoft Graph returns an existing equivalent link rather than creating a duplicate. This "
            + "application records that outcome as 'Reused' and never reports it as 'Created'."),

        new HelpTopic("Retry",
            "Failures caused by throttling or a transient network problem can be retried using the same "
            + "configuration. Failures caused by policy or permissions cannot, and are not offered."),
    ];

    /// <summary>Product metadata, including any unset placeholders.</summary>
    public ProductMetadata Product => _productMetadata.Metadata;

    /// <summary>Opens a documentation URL in the system browser.</summary>
    public async Task OpenUrlAsync(string url)
    {
        if (Uri.TryCreate(url, UriKind.Absolute, out var uri))
        {
            await _browser.OpenAsync(uri).ConfigureAwait(true);
        }
    }
}

/// <summary>About: version, publisher metadata, licences and update check.</summary>
public sealed partial class AboutViewModel : PageViewModelBase
{
    private readonly IProductMetadataProvider _productMetadata;
    private readonly ISystemBrowser _browser;

    /// <summary>Outcome of the last manual update check.</summary>
    [ObservableProperty]
    private string? _updateCheckResult;

    /// <summary>Creates the page.</summary>
    public AboutViewModel(IProductMetadataProvider productMetadata, ISystemBrowser browser)
        : base("About", "about")
    {
        _productMetadata = productMetadata ?? throw new ArgumentNullException(nameof(productMetadata));
        _browser = browser ?? throw new ArgumentNullException(nameof(browser));
    }

    /// <summary>Product and publisher metadata.</summary>
    public ProductMetadata Product => _productMetadata.Metadata;

    /// <summary>Application version.</summary>
    public string Version => _productMetadata.Version;

    /// <summary>Runtime description.</summary>
    public static string Runtime => RuntimeInformation.FrameworkDescription;

    /// <summary>Platform description.</summary>
    public static string Platform => RuntimeInformation.OSDescription;

    /// <summary>True when publisher metadata still contains placeholders.</summary>
    public bool HasPlaceholders => Product.HasPlaceholders;

    /// <summary>Shown when metadata has not been set by a publisher.</summary>
    public static string PlaceholderNotice =>
        "This build has not had its publisher details set. Values shown as PLACEHOLDER must be replaced "
        + "before the application is distributed.";

    /// <summary>The third-party components this application depends on.</summary>
    public static IReadOnlyList<ThirdPartyComponent> ThirdPartyComponents =>
    [
        new ThirdPartyComponent("Avalonia UI", "MIT", "Cross-platform user interface"),
        new ThirdPartyComponent("CommunityToolkit.Mvvm", "MIT", "MVVM source generators"),
        new ThirdPartyComponent("Microsoft.Identity.Client (MSAL)", "MIT", "Microsoft identity authentication"),
        new ThirdPartyComponent("Microsoft.Identity.Client.Extensions.Msal", "MIT", "OS-native token cache"),
        new ThirdPartyComponent("Microsoft.Extensions.*", "MIT", "Dependency injection, configuration and logging"),
        new ThirdPartyComponent(".NET runtime", "MIT", "Application runtime"),
    ];

    /// <summary>
    /// Performs a manual update check. Nothing is downloaded or installed automatically; the
    /// user is only ever pointed at the publisher's release page.
    /// </summary>
    [RelayCommand]
    private async Task CheckForUpdatesAsync()
    {
        if (Product.UpdateCheckUrl.Contains("PLACEHOLDER", StringComparison.OrdinalIgnoreCase))
        {
            UpdateCheckResult =
                "No update endpoint is configured in this build, so there is nothing to check against.";

            return;
        }

        UpdateCheckResult = "Opening the releases page in your browser. "
            + "This application never downloads or installs an update by itself.";

        if (Uri.TryCreate(Product.UpdateCheckUrl, UriKind.Absolute, out var uri))
        {
            await _browser.OpenAsync(uri).ConfigureAwait(true);
        }
    }

    /// <summary>Opens a metadata URL in the system browser.</summary>
    [RelayCommand]
    private async Task OpenUrlAsync(string? url)
    {
        if (!string.IsNullOrWhiteSpace(url)
            && !url.Contains("PLACEHOLDER", StringComparison.OrdinalIgnoreCase)
            && Uri.TryCreate(url, UriKind.Absolute, out var uri))
        {
            await _browser.OpenAsync(uri).ConfigureAwait(true);
        }
        else
        {
            StatusMessage = "That link has not been configured in this build.";
        }
    }
}
