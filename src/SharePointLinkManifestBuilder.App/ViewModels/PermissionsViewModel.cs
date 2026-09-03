using System.Collections.ObjectModel;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using Microsoft.Extensions.Logging;
using SharePointLinkManifestBuilder.App.Infrastructure;
using SharePointLinkManifestBuilder.Core.Abstractions;
using SharePointLinkManifestBuilder.Core.Models;
using SharePointLinkManifestBuilder.Core.Security;

namespace SharePointLinkManifestBuilder.App.ViewModels;

/// <summary>
/// Settings &gt; Microsoft 365 Connection &gt; Permissions.
/// <para>
/// Shows exactly what has been requested, what has actually been granted, and what is missing.
/// It never displays a token or any secret material, and it never claims a permission is
/// granted on the strength of configuration alone: the granted list comes from a real token.
/// </para>
/// </summary>
public sealed partial class PermissionsViewModel : PageViewModelBase
{
    private readonly ConnectionCoordinator _connection;
    private readonly IConsentService _consentService;
    private readonly IAppRegistrationService _registrationService;
    private readonly IRegistrationAuditStore _auditStore;
    private readonly ISystemBrowser _browser;
    private readonly IClipboardService _clipboard;
    private readonly IDiagnosticsService _diagnostics;
    private readonly ILogger<PermissionsViewModel> _logger;

    /// <summary>Result of the last verification.</summary>
    [ObservableProperty]
    private RegistrationVerification? _verification;

    /// <summary>Round-trip time of the last Graph connectivity test.</summary>
    [ObservableProperty]
    private string? _graphConnectivity;

    /// <summary>The display name the user must type before a registration can be deleted.</summary>
    [ObservableProperty]
    private string _deleteConfirmationText = string.Empty;

    /// <summary>True when the guarded deletion panel is expanded.</summary>
    [ObservableProperty]
    private bool _showDangerZone;

    /// <summary>Creates the page.</summary>
    public PermissionsViewModel(
        ConnectionCoordinator connection,
        IConsentService consentService,
        IAppRegistrationService registrationService,
        IRegistrationAuditStore auditStore,
        ISystemBrowser browser,
        IClipboardService clipboard,
        IDiagnosticsService diagnostics,
        ILogger<PermissionsViewModel> logger)
        : base("Permissions", "permissions")
    {
        _connection = connection ?? throw new ArgumentNullException(nameof(connection));
        _consentService = consentService ?? throw new ArgumentNullException(nameof(consentService));
        _registrationService = registrationService ?? throw new ArgumentNullException(nameof(registrationService));
        _auditStore = auditStore ?? throw new ArgumentNullException(nameof(auditStore));
        _browser = browser ?? throw new ArgumentNullException(nameof(browser));
        _clipboard = clipboard ?? throw new ArgumentNullException(nameof(clipboard));
        _diagnostics = diagnostics ?? throw new ArgumentNullException(nameof(diagnostics));
        _logger = logger ?? throw new ArgumentNullException(nameof(logger));
    }

    /// <summary>Per-permission grant state.</summary>
    public ObservableCollection<PermissionGrantState> Permissions { get; } = [];

    /// <summary>The local audit trail of tenant modifications.</summary>
    public ObservableCollection<RegistrationAuditEntry> AuditEntries { get; } = [];

    /// <summary>Things verification could not check, and why.</summary>
    public ObservableCollection<string> NotVerified { get; } = [];

    /// <summary>The tenant configuration in force.</summary>
    public TenantConfiguration? Tenant => _connection.Tenant;

    /// <summary>Tenant display name, or the ID when the name cannot be read.</summary>
    public string TenantDisplay =>
        _connection.TenantDisplayName ?? _connection.Tenant?.TenantId ?? "(not connected)";

    /// <summary>The signed-in account.</summary>
    public string AccountDisplay =>
        _connection.Account?.UserPrincipalName ?? "(not signed in)";

    /// <summary>Client ID in use.</summary>
    public string ClientIdDisplay => _connection.Tenant?.ClientId ?? "(none)";

    /// <summary>How the registration came to exist.</summary>
    public string RegistrationSourceDisplay =>
        _connection.Tenant?.Source.ToString() ?? "None";

    /// <summary>Consent state as last verified.</summary>
    public string ConsentStateDisplay =>
        _connection.Tenant?.ConsentState.ToString() ?? "Unknown";

    /// <summary>Consent type, where it could be determined.</summary>
    public string ConsentTypeDisplay =>
        _connection.Tenant?.ConsentType.ToString() ?? "Unknown";

    /// <summary>When verification last ran.</summary>
    public string LastVerifiedDisplay =>
        _connection.Tenant?.LastVerifiedUtc is { } when
            ? when.UtcDateTime.ToString("u", System.Globalization.CultureInfo.InvariantCulture)
            : "Never";

    /// <summary>Scopes required but not granted.</summary>
    public IReadOnlyList<string> MissingScopes => Verification?.MissingScopes ?? [];

    /// <summary>True when something is missing.</summary>
    public bool HasMissingScopes => MissingScopes.Count > 0;

    /// <summary>True when deletion may be offered at all.</summary>
    public bool CanOfferDeletion =>
        _connection.Tenant?.Source == RegistrationSource.AutomaticSetup;

    /// <summary>Why deletion is not offered, when it is not.</summary>
    public static string DeletionNotOfferedReason =>
        "This registration was supplied by your organization rather than created here, so this application "
        + "will not delete it. Remove the local configuration instead, which changes nothing in your tenant.";

    /// <summary>True when the typed name matches, unlocking deletion.</summary>
    public bool IsDeleteConfirmed =>
        !string.IsNullOrWhiteSpace(_connection.Tenant?.ApplicationDisplayName)
        && string.Equals(
            DeleteConfirmationText.Trim(),
            _connection.Tenant!.ApplicationDisplayName,
            StringComparison.Ordinal);

    /// <inheritdoc />
    public override async Task OnNavigatedToAsync(CancellationToken cancellationToken = default)
    {
        RaiseAll();
        await LoadAuditAsync(cancellationToken).ConfigureAwait(true);

        if (_connection.Tenant is not null && Verification is null)
        {
            await CheckAgainAsync(cancellationToken).ConfigureAwait(true);
        }
    }

    /// <summary>Re-verifies consent and connectivity.</summary>
    [RelayCommand]
    private async Task CheckAgainAsync(CancellationToken cancellationToken)
    {
        var tenant = _connection.Tenant;

        if (tenant is null)
        {
            ErrorMessage = "No Microsoft 365 tenant is configured.";
            return;
        }

        IsBusy = true;
        ClearMessages();

        try
        {
            var required = tenant.RequiredScopes
                .Select(s => GraphScopes.Find(s) ?? new PermissionRequirement
                {
                    Scope = s,
                    Purpose = "Requested by the stored configuration.",
                    DataAccessImpact = "Not described by this build.",
                })
                .ToArray();

            var verification = await _consentService
                .VerifyConsentAsync(tenant, required, cancellationToken)
                .ConfigureAwait(true);

            Verification = verification;

            Permissions.Clear();

            foreach (var state in verification.PermissionStates)
            {
                Permissions.Add(state);
            }

            NotVerified.Clear();

            foreach (var item in verification.NotVerified)
            {
                NotVerified.Add(item);
            }

            var connectivity = await _diagnostics.TestGraphConnectivityAsync(cancellationToken)
                .ConfigureAwait(true);

            GraphConnectivity = connectivity.Succeeded
                ? $"Reachable ({connectivity.Value.TotalMilliseconds:0} ms)"
                : $"Not reachable: {connectivity.Error!.Message}";

            StatusMessage = verification.IsUsable
                ? "Everything required has been granted."
                : "Some required permissions are missing.";
        }
        finally
        {
            IsBusy = false;
            RaiseAll();
        }
    }

    /// <summary>Signs in again, prompting for any missing consent.</summary>
    [RelayCommand]
    private async Task ReauthorizeAsync(CancellationToken cancellationToken)
    {
        IsBusy = true;
        ClearMessages();

        try
        {
            var result = await _connection.SignInAsync(cancellationToken).ConfigureAwait(true);

            StatusMessage = result.Succeeded
                ? "Signed in again."
                : result.Error?.Message;

            if (result.Succeeded)
            {
                await CheckAgainAsync(cancellationToken).ConfigureAwait(true);
            }
        }
        finally
        {
            IsBusy = false;
        }
    }

    /// <summary>Starts Microsoft's consent experience for the missing permissions.</summary>
    [RelayCommand]
    private async Task RequestMissingConsentAsync(CancellationToken cancellationToken)
    {
        var tenant = _connection.Tenant;

        if (tenant is null)
        {
            return;
        }

        IsBusy = true;
        ClearMessages();

        try
        {
            var required = tenant.RequiredScopes
                .Select(s => GraphScopes.Find(s))
                .Where(p => p is not null)
                .Select(p => p!)
                .ToArray();

            var outcome = await _consentService
                .RequestAdminConsentAsync(tenant, required, cancellationToken: cancellationToken)
                .ConfigureAwait(true);

            StatusMessage = outcome.Approved
                ? "Microsoft reported consent. Verifying it now."
                : outcome.Error?.Message ?? "Consent was not completed.";

            if (outcome.Approved)
            {
                await CheckAgainAsync(cancellationToken).ConfigureAwait(true);
            }
        }
        finally
        {
            IsBusy = false;
        }
    }

    /// <summary>Copies the client ID.</summary>
    [RelayCommand]
    private async Task CopyClientIdAsync()
    {
        if (_connection.Tenant?.ClientId is { Length: > 0 } clientId)
        {
            await _clipboard.SetTextAsync(clientId).ConfigureAwait(true);
            StatusMessage = "Client ID copied.";
        }
    }

    /// <summary>Copies a sanitized configuration summary.</summary>
    [RelayCommand]
    private async Task ExportSanitizedSummaryAsync()
    {
        var tenant = _connection.Tenant;

        if (tenant is null)
        {
            return;
        }

        var lines = new List<string>
        {
            "SharePoint Link Manifest Builder - permissions summary",
            $"Generated       : {DateTimeOffset.UtcNow:O}",
            $"Tenant          : {TenantDisplay}",
            $"Tenant ID       : {SensitiveDataRedactor.MaskIdentifier(tenant.TenantId)}",
            $"Client ID       : {SensitiveDataRedactor.MaskIdentifier(tenant.ClientId)}",
            $"Registration    : {tenant.Source}",
            $"Consent state   : {tenant.ConsentState}",
            $"Last verified   : {LastVerifiedDisplay}",
            $"Graph           : {GraphConnectivity ?? "not tested"}",
            string.Empty,
            "Permissions:",
        };

        lines.AddRange(Permissions.Select(p =>
            $"  [{(p.IsGranted ? "granted" : "MISSING")}] {p.Requirement.Scope} ({p.Requirement.Type})"));

        lines.Add(string.Empty);
        lines.Add("No token, secret or certificate is included in this summary.");

        await _clipboard.SetTextAsync(string.Join(Environment.NewLine, lines)).ConfigureAwait(true);
        StatusMessage = "Sanitized summary copied to the clipboard.";
    }

    /// <summary>Opens the app registration in the Entra admin center.</summary>
    [RelayCommand]
    private async Task OpenAppRegistrationAsync()
    {
        if (_connection.Tenant?.ClientId is { Length: > 0 } clientId)
        {
            await _browser.OpenAsync(new Uri(
                "https://entra.microsoft.com/#view/Microsoft_AAD_RegisteredApps/ApplicationMenuBlade/~/Overview/appId/"
                + Uri.EscapeDataString(clientId))).ConfigureAwait(true);
        }
    }

    /// <summary>Opens the enterprise application in the Entra admin center.</summary>
    [RelayCommand]
    private async Task OpenEnterpriseApplicationAsync()
    {
        if (_connection.Tenant?.ClientId is { Length: > 0 } clientId)
        {
            await _browser.OpenAsync(new Uri(
                "https://entra.microsoft.com/#view/Microsoft_AAD_IAM/ManagedAppMenuBlade/~/Overview/objectId//appId/"
                + Uri.EscapeDataString(clientId))).ConfigureAwait(true);
        }
    }

    /// <summary>Repairs the registration's public-client, redirect URI and permission settings.</summary>
    [RelayCommand]
    private async Task RepairRegistrationAsync(CancellationToken cancellationToken)
    {
        var tenant = _connection.Tenant;

        if (tenant?.ApplicationObjectId is not { Length: > 0 } objectId)
        {
            ErrorMessage = "Repair needs the registration's object ID, which is only recorded for "
                + "registrations this application created.";
            return;
        }

        IsBusy = true;
        ClearMessages();

        try
        {
            var desired = new AppRegistrationConfiguration
            {
                DisplayName = tenant.ApplicationDisplayName ?? "SharePoint Link Manifest Builder",
                RequestedPermissions = tenant.RequiredScopes
                    .Select(s => GraphScopes.Find(s))
                    .Where(p => p is not null)
                    .Select(p => p!)
                    .ToArray(),
            };

            var result = await _registrationService
                .RepairRegistrationAsync(objectId, desired, cancellationToken)
                .ConfigureAwait(true);

            await _auditStore.AppendAsync(new RegistrationAuditEntry
            {
                Action = RegistrationAuditAction.ApplicationUpdated,
                TenantId = tenant.TenantId,
                ApplicationDisplayName = desired.DisplayName,
                ClientId = tenant.ClientId,
                PerformedBy = _connection.Account?.PrivacyIdentifier ?? "unknown",
                Changes = result.ChangesApplied,
                Succeeded = result.Succeeded,
                FailureReason = result.Error?.Message,
            }, cancellationToken).ConfigureAwait(true);

            StatusMessage = result.Succeeded
                ? "Registration repaired."
                : result.Error?.Message;

            await LoadAuditAsync(cancellationToken).ConfigureAwait(true);
        }
        finally
        {
            IsBusy = false;
        }
    }

    /// <summary>Removes the local configuration. Nothing in the tenant is changed.</summary>
    [RelayCommand]
    private async Task RemoveLocalConfigurationAsync(CancellationToken cancellationToken)
    {
        IsBusy = true;

        try
        {
            var tenantId = _connection.Tenant?.TenantId;

            await _connection.RemoveLocalConfigurationAsync(cancellationToken).ConfigureAwait(true);

            if (tenantId is not null)
            {
                await _auditStore.AppendAsync(new RegistrationAuditEntry
                {
                    Action = RegistrationAuditAction.LocalConfigurationRemoved,
                    TenantId = tenantId,
                    PerformedBy = "local user",
                    Changes = ["Removed local settings and the token cache on this machine only."],
                    Succeeded = true,
                }, cancellationToken).ConfigureAwait(true);
            }

            StatusMessage = "Local configuration removed. Nothing in your Microsoft 365 tenant was changed.";
            Permissions.Clear();
            Verification = null;
        }
        finally
        {
            IsBusy = false;
            RaiseAll();
        }
    }

    /// <summary>
    /// Deletes the application registration from the tenant. Guarded: only offered for a
    /// registration this application created, and only once the display name has been typed.
    /// </summary>
    [RelayCommand]
    private async Task DeleteRegistrationAsync(CancellationToken cancellationToken)
    {
        var tenant = _connection.Tenant;

        if (tenant?.ApplicationObjectId is not { Length: > 0 } objectId || !CanOfferDeletion)
        {
            ErrorMessage = DeletionNotOfferedReason;
            return;
        }

        if (!IsDeleteConfirmed)
        {
            ErrorMessage = "Type the application's display name exactly to confirm deletion.";
            return;
        }

        IsBusy = true;
        ClearMessages();

        try
        {
            var result = await _registrationService
                .DeleteRegistrationAsync(objectId, DeleteConfirmationText.Trim(), cancellationToken)
                .ConfigureAwait(true);

            await _auditStore.AppendAsync(new RegistrationAuditEntry
            {
                Action = RegistrationAuditAction.ApplicationDeleted,
                TenantId = tenant.TenantId,
                ApplicationDisplayName = tenant.ApplicationDisplayName,
                ClientId = tenant.ClientId,
                PerformedBy = _connection.Account?.PrivacyIdentifier ?? "unknown",
                Changes = ["Deleted the application registration from the tenant."],
                Succeeded = result.Succeeded,
                FailureReason = result.Error?.Message,
            }, cancellationToken).ConfigureAwait(true);

            if (!result.Succeeded)
            {
                ErrorMessage = result.Error!.Message;
                return;
            }

            await _connection.RemoveLocalConfigurationAsync(cancellationToken).ConfigureAwait(true);

            StatusMessage = "The registration was deleted and the local configuration removed.";
            DeleteConfirmationText = string.Empty;
        }
        finally
        {
            IsBusy = false;
            RaiseAll();
        }
    }

    private async Task LoadAuditAsync(CancellationToken cancellationToken)
    {
        AuditEntries.Clear();

        foreach (var entry in (await _auditStore.ListAsync(cancellationToken).ConfigureAwait(true)).Take(50))
        {
            AuditEntries.Add(entry);
        }
    }

    private void RaiseAll()
    {
        OnPropertyChanged(nameof(Tenant));
        OnPropertyChanged(nameof(TenantDisplay));
        OnPropertyChanged(nameof(AccountDisplay));
        OnPropertyChanged(nameof(ClientIdDisplay));
        OnPropertyChanged(nameof(RegistrationSourceDisplay));
        OnPropertyChanged(nameof(ConsentStateDisplay));
        OnPropertyChanged(nameof(ConsentTypeDisplay));
        OnPropertyChanged(nameof(LastVerifiedDisplay));
        OnPropertyChanged(nameof(MissingScopes));
        OnPropertyChanged(nameof(HasMissingScopes));
        OnPropertyChanged(nameof(CanOfferDeletion));
        OnPropertyChanged(nameof(IsDeleteConfirmed));
    }

    partial void OnDeleteConfirmationTextChanged(string value) =>
        OnPropertyChanged(nameof(IsDeleteConfirmed));
}
