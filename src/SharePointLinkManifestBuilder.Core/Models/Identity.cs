namespace SharePointLinkManifestBuilder.Core.Models;

/// <summary>How the tenant-specific application registration came to exist.</summary>
public enum RegistrationSource
{
    /// <summary>No registration is configured yet.</summary>
    None = 0,

    /// <summary>The user supplied a client ID their organization already controls.</summary>
    ExistingRegistration = 1,

    /// <summary>This application created the registration through the setup wizard.</summary>
    AutomaticSetup = 2,
}

/// <summary>The consent state of a configured tenant, as last verified.</summary>
public enum ConsentState
{
    /// <summary>Consent has never been checked.</summary>
    Unknown = 0,

    /// <summary>Every required scope was issued in a real token.</summary>
    Granted = 1,

    /// <summary>Some required scopes were issued; others were not.</summary>
    PartiallyGranted = 2,

    /// <summary>An authorized administrator has not yet approved the request.</summary>
    PendingAdministratorApproval = 3,

    /// <summary>Consent was explicitly denied.</summary>
    Denied = 4,

    /// <summary>The last verification attempt failed for a reason other than denial.</summary>
    VerificationFailed = 5,
}

/// <summary>How consent was obtained, where this can be determined.</summary>
public enum ConsentType
{
    /// <summary>Not determined.</summary>
    Unknown = 0,

    /// <summary>An individual user consented for themselves.</summary>
    User = 1,

    /// <summary>An administrator consented on behalf of the organization.</summary>
    Administrator = 2,
}

/// <summary>
/// The signed-in Microsoft 365 account. Contains no token material; MSAL owns tokens.
/// </summary>
public sealed record UserAccount
{
    /// <summary>Entra object ID of the signed-in user.</summary>
    public required string UserId { get; init; }

    /// <summary>Display name as returned by Microsoft Graph.</summary>
    public required string DisplayName { get; init; }

    /// <summary>User principal name as returned by Microsoft Graph.</summary>
    public required string UserPrincipalName { get; init; }

    /// <summary>The tenant this account is signed into.</summary>
    public required string TenantId { get; init; }

    /// <summary>Tenant display name, when the tenant permits reading it.</summary>
    public string? TenantDisplayName { get; init; }

    /// <summary>MSAL's home account identifier, used to target silent token acquisition.</summary>
    public string? HomeAccountId { get; init; }

    /// <summary>A privacy-conscious identifier for job history: the UPN domain plus a hash.</summary>
    public string PrivacyIdentifier =>
        UserPrincipalName.Contains('@', StringComparison.Ordinal)
            ? $"user@{UserPrincipalName.Split('@')[^1]}"
            : "user";
}

/// <summary>
/// Which Microsoft Entra organizations an application registration will accept sign-ins from.
/// </summary>
public enum TenantAudience
{
    /// <summary>
    /// One named organization. The authority is tenant-specific, so a token issued by any
    /// other tenant is structurally impossible.
    /// </summary>
    SingleTenant = 0,

    /// <summary>
    /// Any work or school organization, using the <c>/organizations</c> authority.
    /// <para>
    /// Deliberately <em>not</em> <c>/common</c>. <c>/common</c> additionally accepts personal
    /// Microsoft accounts, which have no Entra tenant and no SharePoint or OneDrive for
    /// Business, so they can only ever fail later and more confusingly. The tenant a token was
    /// issued by is resolved at sign-in and pinned for the session; see
    /// docs/adr/0011-multi-tenant-authority.md.
    /// </para>
    /// </summary>
    AnyOrganization = 1,
}

/// <summary>
/// Non-secret configuration identifying the tenant and the application registration this
/// installation uses. Persisted locally. Never contains a secret, token, or certificate.
/// </summary>
public sealed record TenantConfiguration
{
    /// <summary>
    /// Entra tenant (directory) ID of the organization the application registration lives in.
    /// Required for <see cref="TenantAudience.SingleTenant"/>, where it also fixes the
    /// authority. Optional for <see cref="TenantAudience.AnyOrganization"/>, where it is only
    /// used to build administrator-consent URLs for the registration's home tenant.
    /// </summary>
    public string TenantId { get; init; } = string.Empty;

    /// <summary>Which organizations this registration accepts sign-ins from.</summary>
    public TenantAudience Audience { get; init; } = TenantAudience.SingleTenant;

    /// <summary>Tenant display name for the UI and manifest headers, when readable.</summary>
    public string? TenantDisplayName { get; init; }

    /// <summary>The application (client) ID used for normal operation.</summary>
    public required string ClientId { get; init; }

    /// <summary>Display name of the application registration, for the UI and audit entries.</summary>
    public string? ApplicationDisplayName { get; init; }

    /// <summary>Object ID of the application registration, when known.</summary>
    public string? ApplicationObjectId { get; init; }

    /// <summary>Object ID of the service principal, when known.</summary>
    public string? ServicePrincipalObjectId { get; init; }

    /// <summary>Identity provider instance, allowing sovereign clouds to be configured.</summary>
    public string Instance { get; init; } = AuthorityDefaults.PublicCloudInstance;

    /// <summary>Microsoft Graph base endpoint, allowing sovereign clouds to be configured.</summary>
    public string GraphEndpoint { get; init; } = AuthorityDefaults.PublicCloudGraphEndpoint;

    /// <summary>How this registration came to exist. Governs whether deletion is offered.</summary>
    public RegistrationSource Source { get; init; } = RegistrationSource.None;

    /// <summary>Consent state as of <see cref="LastVerifiedUtc"/>.</summary>
    public ConsentState ConsentState { get; init; } = ConsentState.Unknown;

    /// <summary>How consent was obtained, where determinable.</summary>
    public ConsentType ConsentType { get; init; } = ConsentType.Unknown;

    /// <summary>Scopes this installation requires.</summary>
    public IReadOnlyList<string> RequiredScopes { get; init; } = [];

    /// <summary>Scopes confirmed issued by Microsoft Entra at the last verification.</summary>
    public IReadOnlyList<string> GrantedScopes { get; init; } = [];

    /// <summary>When this configuration was created.</summary>
    public DateTimeOffset CreatedUtc { get; init; } = DateTimeOffset.UtcNow;

    /// <summary>When consent and connectivity were last verified.</summary>
    public DateTimeOffset? LastVerifiedUtc { get; init; }

    /// <summary>
    /// The authority URL. Tenant-specific for a single-tenant registration; the
    /// <c>/organizations</c> authority for a multi-tenant one. Never <c>/common</c>, which
    /// would additionally admit personal Microsoft accounts.
    /// </summary>
    public string Authority => Audience == TenantAudience.AnyOrganization
        ? $"{Instance.TrimEnd('/')}/{AuthorityDefaults.OrganizationsSegment}"
        : $"{Instance.TrimEnd('/')}/{TenantId}";

    /// <summary>True when sign-in may come from an organization other than <see cref="TenantId"/>.</summary>
    public bool IsMultiTenant => Audience == TenantAudience.AnyOrganization;

    /// <summary>Required scopes that were not present in the last verified token.</summary>
    public IReadOnlyList<string> MissingScopes =>
        RequiredScopes.Where(s => !GrantedScopes.Contains(s, StringComparer.OrdinalIgnoreCase)).ToArray();

    /// <summary>
    /// True when this configuration has everything sign-in needs. A multi-tenant registration
    /// needs only a client ID: the tenant is whichever organization the user signs in to, and
    /// requiring one up front would defeat the point.
    /// </summary>
    public bool IsUsable => Guid.TryParse(ClientId, out _)
        && (IsMultiTenant || Guid.TryParse(TenantId, out _));
}

/// <summary>Default endpoints for the Microsoft public cloud. Sovereign clouds override these.</summary>
public static class AuthorityDefaults
{
    /// <summary>Public cloud identity instance.</summary>
    public const string PublicCloudInstance = "https://login.microsoftonline.com";

    /// <summary>Public cloud Microsoft Graph v1.0 endpoint.</summary>
    public const string PublicCloudGraphEndpoint = "https://graph.microsoft.com/v1.0";

    /// <summary>The well-known Microsoft Graph resource application ID.</summary>
    public const string MicrosoftGraphResourceAppId = "00000003-0000-0000-c000-000000000000";

    /// <summary>The loopback redirect URI registered for the public client.</summary>
    public const string LoopbackRedirectUri = "http://localhost";

    /// <summary>
    /// Authority segment accepting any work or school organization. Chosen over
    /// <c>common</c>, which also admits personal Microsoft accounts.
    /// </summary>
    public const string OrganizationsSegment = "organizations";

    /// <summary>Entra value for a registration scoped to one organization.</summary>
    public const string SignInAudienceSingleTenant = "AzureADMyOrg";

    /// <summary>Entra value for a registration accepting any work or school organization.</summary>
    public const string SignInAudienceMultiTenant = "AzureADMultipleOrgs";
}

/// <summary>
/// Configuration for the publisher-owned bootstrap application used only during automatic
/// setup. This repository ships no client ID; automatic setup stays disabled until one is
/// supplied through configuration (see docs/ENTRA-SETUP.md section 6).
/// </summary>
public sealed record BootstrapConfiguration
{
    /// <summary>Publisher-owned multitenant public client ID, or null when not configured.</summary>
    public string? ClientId { get; init; }

    /// <summary>Identity instance for the bootstrap flow.</summary>
    public string Instance { get; init; } = AuthorityDefaults.PublicCloudInstance;

    /// <summary>True when automatic setup can be offered at all.</summary>
    public bool IsConfigured => !string.IsNullOrWhiteSpace(ClientId) && Guid.TryParse(ClientId, out _);

    /// <summary>
    /// Explains to the user why automatic setup is unavailable, when it is.
    /// </summary>
    public string UnavailableReason =>
        ClientId is null or ""
            ? "Automatic tenant setup is unavailable because this build has no bootstrap client ID configured. "
              + "Use 'Existing app registration' instead, or ask the publisher to configure one."
            : "The configured bootstrap client ID is not a valid GUID.";
}

/// <summary>
/// The application registration this product intends to create, rendered for user review
/// before anything is sent to Microsoft Entra.
/// </summary>
public sealed record AppRegistrationConfiguration
{
    /// <summary>Display name proposed for the registration. User-editable.</summary>
    public required string DisplayName { get; init; }

    /// <summary>
    /// Which organizations the created registration will accept sign-ins from. Defaults to the
    /// narrower single-tenant choice; the user opts in to multi-tenant explicitly.
    /// </summary>
    public TenantAudience Audience { get; init; } = TenantAudience.SingleTenant;

    /// <summary>Supported account types, in the value Microsoft Entra expects.</summary>
    public string SignInAudience => Audience == TenantAudience.AnyOrganization
        ? AuthorityDefaults.SignInAudienceMultiTenant
        : AuthorityDefaults.SignInAudienceSingleTenant;

    /// <summary>Marks the app as a public client, so no secret is ever required.</summary>
    public bool IsFallbackPublicClient { get; init; } = true;

    /// <summary>Native/loopback redirect URIs for the desktop platform.</summary>
    public IReadOnlyList<string> RedirectUris { get; init; } = [AuthorityDefaults.LoopbackRedirectUri];

    /// <summary>Delegated Microsoft Graph scopes to request.</summary>
    public IReadOnlyList<PermissionRequirement> RequestedPermissions { get; init; } = [];

    /// <summary>
    /// A human-readable summary of every tenant change this will make, shown before execution.
    /// </summary>
    public IReadOnlyList<string> DescribePlannedChanges() =>
    [
        $"Create an application registration named \"{DisplayName}\".",
        $"Set supported account types to {SignInAudience} ("
            + (Audience == TenantAudience.AnyOrganization
                ? "any work or school organization; each organization must still consent separately"
                : "this organization only")
            + ").",
        "Enable public client behaviour (no client secret will be created).",
        $"Register redirect URI(s): {string.Join(", ", RedirectUris)}.",
        $"Request {RequestedPermissions.Count} delegated Microsoft Graph permission(s): "
            + string.Join(", ", RequestedPermissions.Select(p => p.Scope)) + ".",
        "No client secret, certificate, or password credential will be created.",
    ];
}
