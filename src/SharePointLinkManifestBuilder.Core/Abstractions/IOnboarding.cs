using SharePointLinkManifestBuilder.Core.Models;

namespace SharePointLinkManifestBuilder.Core.Abstractions;

/// <summary>Whether the signed-in user appears able to create Entra objects.</summary>
public enum RegistrationCapability
{
    /// <summary>Not determined.</summary>
    Unknown = 0,

    /// <summary>The user is expected to be able to create an app registration.</summary>
    Likely = 1,

    /// <summary>The user is unlikely to be able to; the wizard explains and offers Path B.</summary>
    Unlikely = 2,

    /// <summary>A creation attempt was made and refused by the directory.</summary>
    Denied = 3,
}

/// <summary>The outcome of provisioning a tenant-specific registration.</summary>
public sealed record RegistrationProvisioningResult
{
    /// <summary>True when the registration now exists.</summary>
    public required bool Succeeded { get; init; }

    /// <summary>The resulting configuration, when provisioning succeeded.</summary>
    public TenantConfiguration? Configuration { get; init; }

    /// <summary>Object ID of the created application.</summary>
    public string? ApplicationObjectId { get; init; }

    /// <summary>Client ID of the created application.</summary>
    public string? ClientId { get; init; }

    /// <summary>The failure, when provisioning did not succeed.</summary>
    public GraphError? Error { get; init; }

    /// <summary>The exact changes made, mirrored into the local audit history.</summary>
    public IReadOnlyList<string> ChangesApplied { get; init; } = [];
}

/// <summary>The state of a registration as observed in the tenant.</summary>
public sealed record RegistrationVerification
{
    /// <summary>True when an application object with this client ID was found.</summary>
    public bool ApplicationFound { get; init; }

    /// <summary>True when a service principal for the client ID was found.</summary>
    public bool ServicePrincipalFound { get; init; }

    /// <summary>True when the app is configured as a public client.</summary>
    public bool IsPublicClient { get; init; }

    /// <summary>True when the expected loopback redirect URI is registered.</summary>
    public bool RedirectUriConfigured { get; init; }

    /// <summary>True when a bearer token could be obtained for the configured scopes.</summary>
    public bool CanAcquireToken { get; init; }

    /// <summary>True when a basic Graph call succeeded.</summary>
    public bool GraphConnectivityOk { get; init; }

    /// <summary>Per-scope grant state.</summary>
    public IReadOnlyList<PermissionGrantState> PermissionStates { get; init; } = [];

    /// <summary>Consent state as determined by verification.</summary>
    public ConsentState ConsentState { get; init; } = ConsentState.Unknown;

    /// <summary>Consent type, where determinable.</summary>
    public ConsentType ConsentType { get; init; } = ConsentType.Unknown;

    /// <summary>Non-fatal issues found during verification.</summary>
    public IReadOnlyList<string> Warnings { get; init; } = [];

    /// <summary>
    /// Parts of verification that could not be performed, and why. Stated explicitly rather
    /// than reported as a pass.
    /// </summary>
    public IReadOnlyList<string> NotVerified { get; init; } = [];

    /// <summary>Scopes required but not granted.</summary>
    public IReadOnlyList<string> MissingScopes =>
        PermissionStates.Where(p => !p.IsGranted).Select(p => p.Requirement.Scope).ToArray();

    /// <summary>True when the application can actually be used.</summary>
    public bool IsUsable => CanAcquireToken && GraphConnectivityOk && MissingScopes.Count == 0;
}

/// <summary>
/// Creates and inspects the tenant-specific application registration during onboarding.
/// <para>
/// Every method that changes the tenant requires the caller to have shown the planned changes
/// to the user first. Nothing here happens silently.
/// </para>
/// </summary>
public interface IAppRegistrationService
{
    /// <summary>
    /// Estimates whether the signed-in user can create a registration, so the wizard can warn
    /// early rather than failing late. An estimate, never a guarantee.
    /// </summary>
    Task<RegistrationCapability> EstimateCapabilityAsync(CancellationToken cancellationToken = default);

    /// <summary>
    /// Creates the registration by POSTing a complete application object, so no follow-up
    /// PATCH is needed and the bootstrap identity can remain create-only.
    /// </summary>
    /// <param name="configuration">The registration to create, as reviewed by the user.</param>
    /// <param name="tenantId">The tenant to create it in.</param>
    /// <param name="cancellationToken">Cancellation token.</param>
    Task<RegistrationProvisioningResult> CreateRegistrationAsync(
        AppRegistrationConfiguration configuration,
        string tenantId,
        CancellationToken cancellationToken = default);

    /// <summary>Reads a registration by client ID. Requires an elevated read permission.</summary>
    Task<OperationResult<RegistrationVerification>> InspectRegistrationAsync(
        string clientId,
        CancellationToken cancellationToken = default);

    /// <summary>
    /// Creates the service principal explicitly. Not needed on the happy path, because
    /// consenting through Microsoft's endpoint provisions it. Requires the elevated tier.
    /// </summary>
    Task<OperationResult<string>> EnsureServicePrincipalAsync(
        string clientId,
        CancellationToken cancellationToken = default);

    /// <summary>Repairs a registration's public-client, redirect URI and permission settings.</summary>
    Task<RegistrationProvisioningResult> RepairRegistrationAsync(
        string applicationObjectId,
        AppRegistrationConfiguration desired,
        CancellationToken cancellationToken = default);

    /// <summary>
    /// Deletes a registration. Destructive and heavily guarded: the caller must already have
    /// confirmed the display name with the user and verified the registration was created by
    /// this application.
    /// </summary>
    /// <param name="applicationObjectId">Object ID of the registration to delete.</param>
    /// <param name="confirmedDisplayName">The name the user typed, re-checked by the implementation.</param>
    /// <param name="cancellationToken">Cancellation token.</param>
    Task<OperationResult<bool>> DeleteRegistrationAsync(
        string applicationObjectId,
        string confirmedDisplayName,
        CancellationToken cancellationToken = default);
}

/// <summary>
/// The outcome of an administrator consent attempt.
/// <para>
/// An approval here means only that the browser returned a well-formed success for the expected
/// tenant and state. It is never sufficient evidence that consent was actually granted; callers
/// must confirm with <see cref="IConsentService.VerifyConsentAsync"/>.
/// </para>
/// </summary>
public sealed record ConsentOutcome
{
    /// <summary>True when the browser returned an approval for the expected tenant and state.</summary>
    public required bool Approved { get; init; }

    /// <summary>The tenant that returned the response, compared against the expected tenant.</summary>
    public string? ReturnedTenantId { get; init; }

    /// <summary>True when the user closed or cancelled the flow.</summary>
    public bool WasCancelled { get; init; }

    /// <summary>The failure, when consent did not complete.</summary>
    public GraphError? Error { get; init; }
}

/// <summary>
/// Builds and runs Microsoft's official consent experience, and verifies the result.
/// This application never renders a consent screen of its own and never collects credentials.
/// </summary>
public interface IConsentService
{
    /// <summary>
    /// Builds the official Microsoft administrator-consent URL. Exposed so the wizard can offer
    /// "Copy Consent Link" for an administrator who is not at this machine.
    /// </summary>
    /// <param name="tenantConfiguration">Tenant and client to request consent for.</param>
    /// <param name="permissions">Permissions to request.</param>
    /// <param name="redirectUri">A registered redirect URI.</param>
    /// <param name="state">Random state, validated on return.</param>
    /// <param name="targetTenantId">
    /// The organization to request consent in. Optional: a single-organization configuration
    /// uses its own tenant, and a multi-organization one uses the signed-in organization. The
    /// URL always names one explicit organization, never <c>/organizations</c>.
    /// </param>
    Uri BuildAdminConsentUrl(
        TenantConfiguration tenantConfiguration,
        IReadOnlyList<PermissionRequirement> permissions,
        string redirectUri,
        string state,
        string? targetTenantId = null);

    /// <summary>
    /// Opens the official consent experience in the system browser and awaits the redirect,
    /// validating both state and the returned tenant.
    /// </summary>
    /// <param name="tenantConfiguration">Tenant and client to request consent for.</param>
    /// <param name="permissions">Permissions to request.</param>
    /// <param name="targetTenantId">
    /// The organization to request consent in. Optional; see <see cref="BuildAdminConsentUrl"/>.
    /// </param>
    /// <param name="cancellationToken">Cancellation token.</param>
    Task<ConsentOutcome> RequestAdminConsentAsync(
        TenantConfiguration tenantConfiguration,
        IReadOnlyList<PermissionRequirement> permissions,
        string? targetTenantId = null,
        CancellationToken cancellationToken = default);

    /// <summary>
    /// Verifies consent by acquiring a real token and comparing the scopes Microsoft Entra
    /// actually issued against those required. See docs/adr/0006.
    /// </summary>
    Task<RegistrationVerification> VerifyConsentAsync(
        TenantConfiguration tenantConfiguration,
        IReadOnlyList<PermissionRequirement> requiredPermissions,
        CancellationToken cancellationToken = default);
}

/// <summary>Supplies the bootstrap identity used only during automatic setup.</summary>
public interface IBootstrapConfigurationProvider
{
    /// <summary>The bootstrap configuration, which may be unconfigured.</summary>
    BootstrapConfiguration Current { get; }

    /// <summary>Overrides the bootstrap client ID for this session, from the wizard's Advanced field.</summary>
    void SetClientId(string? clientId);
}
