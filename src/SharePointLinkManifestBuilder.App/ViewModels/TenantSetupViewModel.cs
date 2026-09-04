using System.Collections.ObjectModel;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using Microsoft.Extensions.Logging;
using SharePointLinkManifestBuilder.App.Infrastructure;
using SharePointLinkManifestBuilder.Core.Abstractions;
using SharePointLinkManifestBuilder.Core.Models;
using SharePointLinkManifestBuilder.Core.Settings;

namespace SharePointLinkManifestBuilder.App.ViewModels;

/// <summary>The pages of the first-run setup wizard.</summary>
public enum SetupWizardPage
{
    /// <summary>What the application does and what it accesses.</summary>
    Welcome = 0,

    /// <summary>Automatic setup, an existing registration, or manual instructions.</summary>
    ChooseMethod = 1,

    /// <summary>Sign in with a Microsoft work or school account.</summary>
    SignIn = 2,

    /// <summary>Review every permission before anything is requested.</summary>
    PermissionReview = 3,

    /// <summary>Create and configure the tenant-specific registration.</summary>
    Provisioning = 4,

    /// <summary>Grant administrator consent in Microsoft's own experience.</summary>
    Consent = 5,

    /// <summary>Verify the result rather than assuming it.</summary>
    Verification = 6,

    /// <summary>Summary and next steps.</summary>
    Completion = 7,
}

/// <summary>Which onboarding path the user chose.</summary>
public enum SetupMethod
{
    /// <summary>Let the application create the registration.</summary>
    Automatic = 0,

    /// <summary>Supply a client ID the organization already controls.</summary>
    ExistingRegistration = 1,

    /// <summary>Show the steps for an administrator to follow by hand.</summary>
    ManualInstructions = 2,
}

/// <summary>
/// The graphical tenant setup wizard.
/// <para>
/// Three rules shape this whole flow: nothing about the tenant changes without being shown
/// first; consent always happens in Microsoft's own browser experience; and success is verified
/// by acquiring a real token rather than inferred from a redirect that merely looked right.
/// </para>
/// </summary>
public sealed partial class TenantSetupViewModel : PageViewModelBase
{
    private readonly IAuthenticationService _authentication;
    private readonly IAppRegistrationService _registrationService;
    private readonly IConsentService _consentService;
    private readonly IBootstrapConfigurationProvider _bootstrap;
    private readonly ITenantConfigurationStore _tenantStore;
    private readonly IRegistrationAuditStore _auditStore;
    private readonly ConnectionCoordinator _connection;
    private readonly ISystemBrowser _browser;
    private readonly IClipboardService _clipboard;
    private readonly IProductMetadataProvider _productMetadata;
    private readonly ApplicationPaths _paths;
    private readonly ILogger<TenantSetupViewModel> _logger;

    /// <summary>
    /// How long to wait for a step that depends on the user returning from the system browser.
    /// <para>
    /// Both interactive sign-in and administrator consent finish only when Microsoft Entra
    /// redirects back to the loopback listener. When Entra instead shows an error *in the
    /// browser* — a wrong client ID being the common case — no redirect is ever sent, and
    /// without a bound the wait never ends. It is generous on purpose: multi-factor prompts,
    /// password changes and Conditional Access can all legitimately take minutes.
    /// </para>
    /// </summary>
    private static readonly TimeSpan BrowserRoundTripTimeout = TimeSpan.FromMinutes(5);

    /// <summary>
    /// How many times to try signing in to a freshly created registration. With the backoff
    /// below this spans roughly half a minute, which comfortably covers Entra replication.
    /// </summary>
    private const int NewRegistrationSignInAttempts = 5;

    /// <summary>The page currently shown.</summary>
    [ObservableProperty]
    private SetupWizardPage _currentPage = SetupWizardPage.Welcome;

    /// <summary>The onboarding path chosen on page 2.</summary>
    [ObservableProperty]
    private SetupMethod _method = SetupMethod.ExistingRegistration;

    /// <summary>Tenant ID typed or discovered during sign-in.</summary>
    [ObservableProperty]
    private string _tenantId = string.Empty;

    /// <summary>Client ID supplied by the user, for the existing-registration path.</summary>
    [ObservableProperty]
    private string _existingClientId = string.Empty;

    /// <summary>
    /// True when the registration accepts any work or school organization. This makes the
    /// Directory (tenant) ID optional, because the organization is whichever one the user signs
    /// in to rather than one fixed at setup time.
    /// </summary>
    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(TenantIdIsRequired))]
    [NotifyPropertyChangedFor(nameof(TenantIdHint))]
    private bool _useAnyOrganization;

    /// <summary>Display name proposed for a registration this application creates.</summary>
    [ObservableProperty]
    private string _proposedApplicationName = "SharePoint Link Manifest Builder";

    /// <summary>Bootstrap client ID typed into the Advanced field.</summary>
    [ObservableProperty]
    private string _bootstrapClientIdOverride = string.Empty;

    /// <summary>
    /// Whether the bootstrap registration should request <c>Application.ReadWrite.All</c> instead
    /// of <c>AppRegistration.Create</c>.
    /// <para>
    /// Off by default, and deliberately phrased as a downgrade rather than an option.
    /// <c>AppRegistration.Create</c> is the least-privileged permission Microsoft documents for
    /// <c>POST /applications</c> and is all this application needs, but it is new enough that
    /// some tenants' portal permission picker does not list it. Where that happens the only
    /// consentable alternative is far broader, so the user chooses it knowingly.
    /// </para>
    /// </summary>
    [ObservableProperty]
    private bool _useBroadBootstrapPermission;

    /// <summary>The account signed in during the wizard.</summary>
    [ObservableProperty]
    private UserAccount? _signedInAccount;

    /// <summary>True to request the optional people-picker permission.</summary>
    [ObservableProperty]
    private bool _includeUserPicker;

    /// <summary>True to request the broad SharePoint write permission.</summary>
    [ObservableProperty]
    private bool _includeBroadSharePointWrite;

    /// <summary>True to request no write capability at all.</summary>
    [ObservableProperty]
    private bool _readOnlyMode;

    /// <summary>The result of the last verification.</summary>
    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(VerificationVerdict))]
    [NotifyPropertyChangedFor(nameof(VerificationPassed))]
    [NotifyPropertyChangedFor(nameof(VerificationFailed))]
    private RegistrationVerification? _verification;

    /// <summary>The consent URL, offered for copying when the user cannot consent themselves.</summary>
    [ObservableProperty]
    private string? _consentUrl;

    /// <summary>True once the configuration has been saved as pending administrator approval.</summary>
    [ObservableProperty]
    private bool _isPendingAdministratorApproval;

    /// <summary>The client ID of the registration in use once setup completes.</summary>
    [ObservableProperty]
    private string? _resultingClientId;

    /// <summary>
    /// True when the Directory (tenant) ID must be supplied. A multi-organization registration
    /// does not need one up front: the organization is resolved from the token at sign-in.
    /// </summary>
    public bool TenantIdIsRequired => !UseAnyOrganization;

    /// <summary>Guidance shown beside the Directory (tenant) ID field.</summary>
    public string TenantIdHint => UseAnyOrganization
        ? "Optional. Leave this blank and the organization will be taken from the account you "
          + "sign in with. Supply one only to pre-select a specific organization."
        : "Required. The Directory (tenant) ID GUID shown on the Microsoft Entra overview page.";

    /// <summary>Creates the wizard.</summary>
    public TenantSetupViewModel(
        IAuthenticationService authentication,
        IAppRegistrationService registrationService,
        IConsentService consentService,
        IBootstrapConfigurationProvider bootstrap,
        ITenantConfigurationStore tenantStore,
        IRegistrationAuditStore auditStore,
        ConnectionCoordinator connection,
        ISystemBrowser browser,
        IClipboardService clipboard,
        IProductMetadataProvider productMetadata,
        ApplicationPaths paths,
        ILogger<TenantSetupViewModel> logger)
        : base("Tenant Setup", "setup")
    {
        _authentication = authentication ?? throw new ArgumentNullException(nameof(authentication));
        _registrationService = registrationService ?? throw new ArgumentNullException(nameof(registrationService));
        _consentService = consentService ?? throw new ArgumentNullException(nameof(consentService));
        _bootstrap = bootstrap ?? throw new ArgumentNullException(nameof(bootstrap));
        _tenantStore = tenantStore ?? throw new ArgumentNullException(nameof(tenantStore));
        _auditStore = auditStore ?? throw new ArgumentNullException(nameof(auditStore));
        _connection = connection ?? throw new ArgumentNullException(nameof(connection));
        _browser = browser ?? throw new ArgumentNullException(nameof(browser));
        _clipboard = clipboard ?? throw new ArgumentNullException(nameof(clipboard));
        _productMetadata = productMetadata ?? throw new ArgumentNullException(nameof(productMetadata));
        _paths = paths ?? throw new ArgumentNullException(nameof(paths));
        _logger = logger ?? throw new ArgumentNullException(nameof(logger));

        RefreshRequestedPermissions();
    }

    /// <summary>The permissions that will be requested, with their justifications.</summary>
    public ObservableCollection<PermissionRequirement> RequestedPermissions { get; } = [];

    /// <summary>The exact tenant changes that will be made, shown before anything happens.</summary>
    public ObservableCollection<string> PlannedChanges { get; } = [];

    /// <summary>Warnings raised during verification.</summary>
    public ObservableCollection<string> VerificationWarnings { get; } = [];

    /// <summary>Things verification could not check, and why.</summary>
    public ObservableCollection<string> VerificationNotChecked { get; } = [];

    /// <summary>
    /// The one-line verdict: whether this configuration can actually be used. Everything else on
    /// the page is detail behind this answer.
    /// </summary>
    public string VerificationVerdict => Verification switch
    {
        null => "Not verified yet.",
        { IsUsable: true } => "Ready to use. A token was issued with every required permission.",
        { CanAcquireToken: false } => "Not ready. No token could be obtained for this configuration.",
        _ => "Not ready. A token was issued, but it is missing at least one required permission.",
    };

    /// <summary>True once verification has run and succeeded.</summary>
    public bool VerificationPassed => Verification?.IsUsable == true;

    /// <summary>True once verification has run and did not succeed.</summary>
    public bool VerificationFailed => Verification is not null && !Verification.IsUsable;

    /// <summary>Where this application stores local data, shown on the welcome page.</summary>
    public IReadOnlyList<StorageLocationInfo> StorageLocations => _paths.Describe();

    /// <summary>Product metadata, including any unset placeholders.</summary>
    public ProductMetadata Product => _productMetadata.Metadata;

    /// <summary>True when a publisher has configured a bootstrap identity.</summary>
    /// <summary>
    /// True when a bootstrap client ID is configured, so automatic setup is ready to run.
    /// <para>
    /// This is deliberately <em>not</em> what gates selecting the automatic method. This build
    /// ships no bootstrap client ID, and the Advanced field that supplies one lives on the
    /// automatic path — so gating selection on this made the only way to enable automatic setup
    /// reachable only after it was already enabled. Selecting the method is always allowed; the
    /// run is guarded instead, in CreateRegistrationAsync.
    /// </para>
    /// </summary>
    public bool IsAutomaticSetupAvailable => _bootstrap.Current.IsConfigured;

    /// <summary>Why automatic setup is unavailable, when it is.</summary>
    public string AutomaticSetupUnavailableReason => _bootstrap.Current.UnavailableReason;

    /// <summary>True when the wizard is on its last page.</summary>
    public bool IsOnLastPage => CurrentPage == SetupWizardPage.Completion;

    /// <summary>True when Back is meaningful.</summary>
    public bool CanGoBack => CurrentPage > SetupWizardPage.Welcome && !IsBusy;

    /// <summary>The heading for the current page.</summary>
    public string PageTitle => CurrentPage switch
    {
        SetupWizardPage.Welcome => "Welcome",
        SetupWizardPage.ChooseMethod => "Choose a setup method",
        SetupWizardPage.SignIn => "Sign in to Microsoft 365",
        SetupWizardPage.PermissionReview => "Review permissions",
        SetupWizardPage.Provisioning => "Create the application registration",
        SetupWizardPage.Consent => "Grant administrator consent",
        SetupWizardPage.Verification => "Verify the result",
        SetupWizardPage.Completion => "Setup complete",
        _ => "Setup",
    };

    /// <summary>Step indicator, for example "Step 3 of 8".</summary>
    public string StepIndicator => $"Step {(int)CurrentPage + 1} of 8";

    /// <summary>What the welcome page tells the user before anything happens.</summary>
    public static IReadOnlyList<string> WelcomePoints =>
    [
        "This application builds manifests of SharePoint and OneDrive file links, so an AI system such as "
        + "Microsoft Copilot can be given an explicit list of files instead of having to discover them.",

        "It reads site, library, folder and file information through Microsoft Graph, and creates sharing "
        + "links only for the locations you select.",

        "It acts as you. It can never see or share anything your own account cannot already open.",

        "Sharing links remain governed by your organization's Microsoft 365 policy. If your organization "
        + "blocks a kind of link, the request will be refused and reported as blocked.",

        "No password, client secret or certificate is ever collected. Sign-in happens in your normal web "
        + "browser, not inside this application.",

        "Telemetry is disabled and no usage data is sent anywhere.",
    ];

    /// <inheritdoc />
    public override Task OnNavigatedToAsync(CancellationToken cancellationToken = default)
    {
        if (_connection.Tenant is { } tenant)
        {
            TenantId = tenant.TenantId;
            ExistingClientId = tenant.ClientId;
        }

        SignedInAccount = _authentication.CurrentAccount;
        RefreshRequestedPermissions();

        return Task.CompletedTask;
    }

    /// <summary>Moves to the next page, running that page's work where appropriate.</summary>
    [RelayCommand]
    private async Task NextAsync(CancellationToken cancellationToken)
    {
        ClearMessages();

        switch (CurrentPage)
        {
            case SetupWizardPage.Welcome:
                CurrentPage = SetupWizardPage.ChooseMethod;
                break;

            case SetupWizardPage.ChooseMethod:
                if (Method == SetupMethod.ManualInstructions)
                {
                    StatusMessage = "Follow the steps shown, then return and choose "
                        + "'Use an existing app registration'.";
                    return;
                }

                CurrentPage = SetupWizardPage.SignIn;
                break;

            case SetupWizardPage.SignIn:
                if (SignedInAccount is null)
                {
                    ErrorMessage = "Sign in before continuing.";
                    return;
                }

                RefreshRequestedPermissions();
                CurrentPage = SetupWizardPage.PermissionReview;
                break;

            case SetupWizardPage.PermissionReview:
                CurrentPage = Method == SetupMethod.Automatic
                    ? SetupWizardPage.Provisioning
                    : SetupWizardPage.Consent;
                break;

            case SetupWizardPage.Provisioning:
                CurrentPage = SetupWizardPage.Consent;
                break;

            case SetupWizardPage.Consent:
                CurrentPage = SetupWizardPage.Verification;
                await VerifyAsync(cancellationToken).ConfigureAwait(true);
                break;

            case SetupWizardPage.Verification:
                CurrentPage = SetupWizardPage.Completion;
                break;

            default:
                break;
        }

        RaisePageChanged();
    }

    /// <summary>Moves back a page.</summary>
    [RelayCommand]
    private void Back()
    {
        if (CurrentPage > SetupWizardPage.Welcome)
        {
            CurrentPage--;
            RaisePageChanged();
        }
    }

    /// <summary>
    /// Signs in through the system browser. For the existing-registration path this uses the
    /// supplied client ID; for automatic setup it uses the publisher's bootstrap identity.
    /// </summary>
    [RelayCommand(IncludeCancelCommand = true)]
    private async Task SignInAsync(CancellationToken cancellationToken)
    {
        ClearMessages();

        var clientId = Method == SetupMethod.Automatic
            ? ResolveBootstrapClientId()
            : ExistingClientId.Trim();

        if (string.IsNullOrWhiteSpace(clientId) || !Guid.TryParse(clientId, out _))
        {
            ErrorMessage = Method == SetupMethod.Automatic
                ? AutomaticSetupUnavailableReason
                : "Enter the Application (client) ID from your app registration. It is a GUID.";

            return;
        }

        if (TenantIdIsRequired && !Guid.TryParse(TenantId.Trim(), out _))
        {
            ErrorMessage = "Enter your Directory (tenant) ID. It is a GUID, shown on the Entra overview page.";
            return;
        }

        if (!TenantIdIsRequired
            && TenantId.Trim() is { Length: > 0 } typedTenant
            && !Guid.TryParse(typedTenant, out _))
        {
            ErrorMessage = "The Directory (tenant) ID is optional for a multi-organization "
                + "registration, but if you supply one it must be a GUID.";

            return;
        }

        IsBusy = true;

        try
        {
            var configuration = new TenantConfiguration
            {
                TenantId = TenantId.Trim(),
                Audience = UseAnyOrganization ? TenantAudience.AnyOrganization : TenantAudience.SingleTenant,
                ClientId = clientId,
                RequiredScopes = RequestedPermissions.Select(p => p.Scope).ToArray(),
                Source = Method == SetupMethod.Automatic
                    ? RegistrationSource.AutomaticSetup
                    : RegistrationSource.ExistingRegistration,
            };

            await _connection.ApplyTenantAsync(configuration, cancellationToken).ConfigureAwait(true);

            // The narrow tier is the default and the right answer. The broad tier exists because
            // AppRegistration.Create is recent enough that some tenants' portal permission picker
            // does not offer it yet, leaving Application.ReadWrite.All as the only consentable
            // alternative. That is a real escalation -- it can modify and delete every
            // registration in the directory, not just create one -- so it is an explicit choice
            // the user makes and can see, never a silent fallback.
            var bootstrapTier = UseBroadBootstrapPermission
                ? GraphScopes.BootstrapManageTier
                : GraphScopes.BootstrapCreateOnlyTier;

            var scopes = Method == SetupMethod.Automatic
                ? bootstrapTier.Select(p => p.Scope).ToArray()
                : configuration.RequiredScopes.ToArray();

            if (Method == SetupMethod.Automatic && UseBroadBootstrapPermission)
            {
                _logger.LogWarning(
                    "Automatic setup is using the broad bootstrap permission tier at the user's "
                    + "explicit request, because AppRegistration.Create was unavailable.");
            }

            using var timeout = new CancellationTokenSource(BrowserRoundTripTimeout);
            using var linked = CancellationTokenSource
                .CreateLinkedTokenSource(cancellationToken, timeout.Token);

            // Through the coordinator, not the authentication service directly. The coordinator
            // is what records the granted scopes, saves the tenant and moves the application to
            // Connected; bypassing it produced a wizard that reached "Setup complete, consent
            // granted" while every other page still showed "Not connected".
            var result = await _connection.SignInAsync(scopes, linked.Token).ConfigureAwait(true);

            if (!result.Succeeded)
            {
                // A timeout and a deliberate cancellation both arrive as a cancelled result, but
                // they mean different things to the user and only one of them is their doing.
                ErrorMessage = timeout.IsCancellationRequested && !cancellationToken.IsCancellationRequested
                    ? "Sign-in did not complete within five minutes, so it was stopped. If the "
                      + "browser showed an error instead of returning here, the Application "
                      + "(client) ID or Directory (tenant) ID is usually wrong. Correct them and "
                      + "sign in again."
                    : Describe(result.Error);

                return;
            }

            SignedInAccount = result.Account;

            // For a multi-organization registration the organization is discovered, not typed.
            // Recording it here is what lets the consent step name an explicit directory.
            if (UseAnyOrganization
                && result.Account?.TenantId is { Length: > 0 } discovered
                && !string.Equals(TenantId.Trim(), discovered, StringComparison.OrdinalIgnoreCase))
            {
                TenantId = discovered;
            }

            StatusMessage = $"Signed in as {result.Account?.DisplayName} in tenant {result.Account?.TenantId}.";
        }
        finally
        {
            IsBusy = false;
        }
    }

    /// <summary>
    /// Creates the tenant-specific registration. Only reached after the user has seen the exact
    /// list of changes on this page.
    /// </summary>
    [RelayCommand]
    private async Task CreateRegistrationAsync(CancellationToken cancellationToken)
    {
        ClearMessages();

        if (!IsAutomaticSetupAvailable && string.IsNullOrWhiteSpace(BootstrapClientIdOverride))
        {
            ErrorMessage = AutomaticSetupUnavailableReason;
            return;
        }

        IsBusy = true;

        try
        {
            var configuration = new AppRegistrationConfiguration
            {
                DisplayName = ProposedApplicationName.Trim(),
                Audience = UseAnyOrganization ? TenantAudience.AnyOrganization : TenantAudience.SingleTenant,
                RequestedPermissions = RequestedPermissions.ToArray(),
            };

            var result = await _registrationService
                .CreateRegistrationAsync(configuration, TenantId.Trim(), cancellationToken)
                .ConfigureAwait(true);

            // Every material tenant change is audited locally, whether or not it succeeded.
            await _auditStore.AppendAsync(new RegistrationAuditEntry
            {
                Action = RegistrationAuditAction.ApplicationCreated,
                TenantId = TenantId.Trim(),
                TenantDisplayName = _connection.TenantDisplayName,
                ApplicationDisplayName = configuration.DisplayName,
                ClientId = result.ClientId,
                PerformedBy = SignedInAccount?.PrivacyIdentifier ?? "unknown",
                Changes = result.Succeeded ? result.ChangesApplied : configuration.DescribePlannedChanges(),
                Succeeded = result.Succeeded,
                FailureReason = result.Error?.Message,
            }, cancellationToken).ConfigureAwait(true);

            if (!result.Succeeded)
            {
                ErrorMessage = Describe(result.Error);
                return;
            }

            ResultingClientId = result.ClientId;

            var tenantConfiguration = result.Configuration! with
            {
                TenantDisplayName = _connection.TenantDisplayName,
                RequiredScopes = RequestedPermissions.Select(p => p.Scope).ToArray(),
            };

            await _connection.SaveTenantAsync(tenantConfiguration, cancellationToken).ConfigureAwait(true);

            // Sign in to the registration that was just created, with the operating scopes.
            //
            // Without this the wizard configured the application for a client ID nobody had ever
            // authenticated against. Consent was then requested for an app the user had no grant
            // on, the connection never reached Connected because no sign-in with the operating
            // scopes had happened, and the only way to actually finish was to leave the wizard
            // and use the Home page's sign-in button. Doing it here also creates the user's
            // grant, which is what makes consent verification able to see anything at all.
            StatusMessage = "Registration created. Signing in to it now.";

            var signIn = await SignInToNewRegistrationAsync(
                tenantConfiguration.RequiredScopes, cancellationToken).ConfigureAwait(true);

            if (!signIn.Succeeded)
            {
                ErrorMessage = Describe(signIn.Error)
                    + " The registration was created successfully; only signing in to it failed. "
                    + "Use 'Sign in' to try again.";

                return;
            }

            SignedInAccount = signIn.Account;
            RefreshRequestedPermissions();

            StatusMessage = "The application registration was created and signed in to. "
                + "If any permission still needs an administrator, the consent step is next.";
        }
        finally
        {
            IsBusy = false;
        }
    }

    /// <summary>Opens Microsoft's official administrator consent experience.</summary>
    [RelayCommand(IncludeCancelCommand = true)]
    private async Task RequestConsentAsync(CancellationToken cancellationToken)
    {
        ClearMessages();

        var tenant = _connection.Tenant;

        if (tenant is null)
        {
            ErrorMessage = "Sign in first so the tenant and application are known.";
            return;
        }

        IsBusy = true;

        try
        {
            using var timeout = new CancellationTokenSource(BrowserRoundTripTimeout);
            using var linked = CancellationTokenSource
                .CreateLinkedTokenSource(cancellationToken, timeout.Token);

            var outcome = await _consentService
                .RequestAdminConsentAsync(
                    tenant,
                    RequestedPermissions.ToArray(),
                    cancellationToken: linked.Token)
                .ConfigureAwait(true);

            await _auditStore.AppendAsync(new RegistrationAuditEntry
            {
                Action = RegistrationAuditAction.ConsentRequested,
                TenantId = tenant.TenantId,
                TenantDisplayName = tenant.TenantDisplayName,
                ApplicationDisplayName = tenant.ApplicationDisplayName,
                ClientId = tenant.ClientId,
                PerformedBy = SignedInAccount?.PrivacyIdentifier ?? "unknown",
                Changes = [$"Requested consent for: {string.Join(", ", RequestedPermissions.Select(p => p.Scope))}"],
                Succeeded = outcome.Approved,
                FailureReason = outcome.Error?.Message,
            }, cancellationToken).ConfigureAwait(true);

            if (outcome.WasCancelled)
            {
                StatusMessage = "Consent was cancelled. Nothing was changed.";
                return;
            }

            if (!outcome.Approved)
            {
                // Before reporting a failure, check whether consent was needed at all. The
                // consent endpoint can fail for reasons that have nothing to do with whether the
                // permissions are granted -- a Conditional Access device requirement the browser
                // cannot satisfy, for instance -- while a token carrying every required scope has
                // already been issued. Reporting "consent failed" then sends the user to fix
                // something that is not broken.
                var alreadyGranted = await _consentService
                    .VerifyConsentAsync(tenant, RequestedPermissions.ToArray(),
                        allowInteractive: false, cancellationToken)
                    .ConfigureAwait(true);

                if (alreadyGranted.IsUsable)
                {
                    Verification = alreadyGranted;

                    StatusMessage = "The administrator consent step did not complete, but it is not "
                        + "needed: a token has already been issued with every permission this "
                        + "application requires.";

                    _logger.LogInformation(
                        "Consent request failed but verification shows every required permission is "
                        + "already granted, so the failure is not reported as a problem.");

                    return;
                }

                ErrorMessage = Describe(outcome.Error);
                await SaveAsPendingAsync(cancellationToken).ConfigureAwait(true);
                return;
            }

            // Approval is not proof. The next page verifies it against a real token.
            StatusMessage = "Microsoft reported consent. The next step verifies it independently.";
        }
        finally
        {
            IsBusy = false;
        }
    }

    /// <summary>Copies the consent URL so an administrator elsewhere can complete the step.</summary>
    [RelayCommand]
    private async Task CopyConsentLinkAsync()
    {
        var tenant = _connection.Tenant;

        if (tenant is null)
        {
            return;
        }

        // A shareable link uses the plain loopback redirect, since the administrator will not
        // have this application's listener running on their machine.
        var url = _consentService.BuildAdminConsentUrl(
            tenant,
            RequestedPermissions.ToArray(),
            AuthorityDefaults.LoopbackRedirectUri,
            LoopbackState());

        ConsentUrl = url.ToString();
        await _clipboard.SetTextAsync(ConsentUrl).ConfigureAwait(true);

        StatusMessage = "Consent link copied. Send it to an authorized Microsoft Entra administrator.";
    }

    /// <summary>Saves the configuration as pending administrator approval.</summary>
    [RelayCommand]
    private async Task SaveAsPendingAsync(CancellationToken cancellationToken)
    {
        var tenant = _connection.Tenant;

        if (tenant is null)
        {
            return;
        }

        await _connection
            .SaveTenantAsync(
                tenant with { ConsentState = ConsentState.PendingAdministratorApproval },
                cancellationToken)
            .ConfigureAwait(true);

        IsPendingAdministratorApproval = true;

        StatusMessage = "Saved as waiting for administrator approval. Use 'Check again' once an "
            + "administrator has approved the request.";
    }

    /// <summary>Verifies consent by acquiring a real token and comparing granted scopes.</summary>
    [RelayCommand]
    private async Task VerifyAsync(CancellationToken cancellationToken)
    {
        ClearMessages();

        var tenant = _connection.Tenant;

        if (tenant is null)
        {
            ErrorMessage = "There is no configuration to verify yet.";
            return;
        }

        IsBusy = true;
        VerificationWarnings.Clear();
        VerificationNotChecked.Clear();

        try
        {
            var verification = await _consentService
                .VerifyConsentAsync(tenant, RequestedPermissions.ToArray(), allowInteractive: true,
                    cancellationToken)
                .ConfigureAwait(true);

            Verification = verification;

            foreach (var warning in verification.Warnings)
            {
                VerificationWarnings.Add(warning);
            }

            foreach (var item in verification.NotVerified)
            {
                VerificationNotChecked.Add(item);
            }

            await _connection.SaveTenantAsync(
                tenant with
                {
                    ConsentState = verification.ConsentState,
                    ConsentType = verification.ConsentType,
                    GrantedScopes = verification.PermissionStates
                        .Where(p => p.IsGranted)
                        .Select(p => p.Requirement.Scope)
                        .ToArray(),
                    LastVerifiedUtc = DateTimeOffset.UtcNow,
                },
                cancellationToken).ConfigureAwait(true);

            IsPendingAdministratorApproval =
                verification.ConsentState == ConsentState.PendingAdministratorApproval;

            StatusMessage = verification.IsUsable
                ? "Verified: a token was issued with every required permission."
                : verification.ConsentState == ConsentState.PendingAdministratorApproval
                    ? "Still waiting for an administrator to approve the request."
                    : $"Verification incomplete. Missing: {string.Join(", ", verification.MissingScopes)}";
        }
        finally
        {
            IsBusy = false;
            OnPropertyChanged(nameof(Verification));
        }
    }

    /// <summary>Opens the app registration in the Entra admin center.</summary>
    [RelayCommand]
    private async Task OpenAppRegistrationAsync()
    {
        var clientId = _connection.Tenant?.ClientId;

        if (string.IsNullOrWhiteSpace(clientId))
        {
            return;
        }

        await _browser.OpenAsync(new Uri(
            "https://entra.microsoft.com/#view/Microsoft_AAD_RegisteredApps/ApplicationMenuBlade/~/Overview/appId/"
            + Uri.EscapeDataString(clientId))).ConfigureAwait(true);
    }

    /// <summary>Opens the enterprise application in the Entra admin center.</summary>
    [RelayCommand]
    private async Task OpenEnterpriseApplicationAsync()
    {
        var clientId = _connection.Tenant?.ClientId;

        if (string.IsNullOrWhiteSpace(clientId))
        {
            return;
        }

        await _browser.OpenAsync(new Uri(
            "https://entra.microsoft.com/#view/Microsoft_AAD_IAM/ManagedAppMenuBlade/~/Overview/objectId//appId/"
            + Uri.EscapeDataString(clientId))).ConfigureAwait(true);
    }

    /// <summary>Copies a sanitized summary of the configuration for a support ticket.</summary>
    [RelayCommand]
    private async Task ExportSanitizedSummaryAsync()
    {
        var tenant = _connection.Tenant;

        if (tenant is null)
        {
            return;
        }

        var summary = string.Join(Environment.NewLine,
        [
            "SharePoint Link Manifest Builder - configuration summary",
            $"Generated       : {DateTimeOffset.UtcNow:O}",
            $"Application     : {tenant.ApplicationDisplayName ?? "(not recorded)"}",

            // Masked: not secret, but identifying, and a summary is meant to be shareable.
            $"Client ID       : {Core.Security.SensitiveDataRedactor.MaskIdentifier(tenant.ClientId)}",
            $"Tenant ID       : {Core.Security.SensitiveDataRedactor.MaskIdentifier(tenant.TenantId)}",
            $"Registration    : {tenant.Source}",
            $"Consent state   : {tenant.ConsentState}",
            $"Consent type    : {tenant.ConsentType}",
            $"Required scopes : {string.Join(", ", tenant.RequiredScopes)}",
            $"Granted scopes  : {string.Join(", ", tenant.GrantedScopes)}",
            $"Missing scopes  : {string.Join(", ", tenant.MissingScopes)}",
            $"Last verified   : {tenant.LastVerifiedUtc:O}",
            "No token, secret or certificate is included in this summary.",
        ]);

        await _clipboard.SetTextAsync(summary).ConfigureAwait(true);
        StatusMessage = "Sanitized configuration summary copied to the clipboard.";
    }

    /// <summary>Rebuilds the requested permission list from the current options.</summary>
    [RelayCommand]
    private void RefreshRequestedPermissions()
    {
        RequestedPermissions.Clear();

        var permissions = GraphScopes.BuildOperatingSet(
            readOnly: ReadOnlyMode,
            includeUserOneDrivePicker: IncludeUserPicker,
            includeBroadSharePointWrite: IncludeBroadSharePointWrite);

        foreach (var permission in permissions)
        {
            RequestedPermissions.Add(permission);
        }

        PlannedChanges.Clear();

        var configuration = new AppRegistrationConfiguration
        {
            DisplayName = ProposedApplicationName,
            Audience = UseAnyOrganization ? TenantAudience.AnyOrganization : TenantAudience.SingleTenant,
            RequestedPermissions = permissions,
        };

        foreach (var change in configuration.DescribePlannedChanges())
        {
            PlannedChanges.Add(change);
        }
    }

    /// <summary>
    /// Signs in to a registration that was created moments ago, retrying while Microsoft Entra
    /// still reports it as unknown.
    /// <para>
    /// A new application is not usable the instant <c>POST /applications</c> returns 201. Entra
    /// replicates it first, and until that finishes a sign-in fails with AADSTS700016 -- "this
    /// app registration does not exist" -- which is true only for another second or two. Failing
    /// the wizard on it would report a registration that was created perfectly well as broken.
    /// </para>
    /// </summary>
    private async Task<AuthenticationResultInfo> SignInToNewRegistrationAsync(
        IReadOnlyList<string> scopes,
        CancellationToken cancellationToken)
    {
        var delay = TimeSpan.FromSeconds(2);
        AuthenticationResultInfo result;

        for (var attempt = 1; ; attempt++)
        {
            result = await _connection.SignInAsync(scopes, cancellationToken).ConfigureAwait(true);

            var stillReplicating = result.Error?.Kind == GraphErrorKind.ApplicationNotFoundInTenant;

            if (result.Succeeded || !stillReplicating || attempt >= NewRegistrationSignInAttempts)
            {
                return result;
            }

            _logger.LogInformation(
                "The new registration is not visible to Microsoft Entra yet. "
                + "Waiting {Delay} before attempt {NextAttempt}.",
                delay,
                attempt + 1);

            StatusMessage = "Waiting for Microsoft to finish provisioning the new registration.";

            await Task.Delay(delay, cancellationToken).ConfigureAwait(true);
            delay += TimeSpan.FromSeconds(3);
        }
    }

    private string? ResolveBootstrapClientId()
    {
        if (!string.IsNullOrWhiteSpace(BootstrapClientIdOverride))
        {
            _bootstrap.SetClientId(BootstrapClientIdOverride.Trim());
        }

        return _bootstrap.Current.ClientId;
    }

    private static string LoopbackState() =>
        Convert.ToHexString(System.Security.Cryptography.RandomNumberGenerator.GetBytes(16));

    private static string Describe(GraphError? error)
    {
        if (error is null)
        {
            return "The operation did not succeed.";
        }

        var text = error.SuggestedAction is { Length: > 0 } action
            ? $"{error.Message} {action}"
            : error.Message;

        // The Microsoft error code is included deliberately. It is the difference between a
        // failure the user can look up or hand to an administrator and one they can only guess
        // at, and it identifies no person, tenant, or resource.
        return error.GraphErrorCode is { Length: > 0 } code
            ? $"{text} (Microsoft error code: {code})"
            : text;
    }

    private void RaisePageChanged()
    {
        OnPropertyChanged(nameof(PageTitle));
        OnPropertyChanged(nameof(StepIndicator));
        OnPropertyChanged(nameof(CanGoBack));
        OnPropertyChanged(nameof(IsOnLastPage));
    }

    partial void OnCurrentPageChanged(SetupWizardPage value) => RaisePageChanged();

    partial void OnReadOnlyModeChanged(bool value) => RefreshRequestedPermissions();

    partial void OnIncludeUserPickerChanged(bool value) => RefreshRequestedPermissions();

    partial void OnIncludeBroadSharePointWriteChanged(bool value) => RefreshRequestedPermissions();

    partial void OnProposedApplicationNameChanged(string value) => RefreshRequestedPermissions();
}
