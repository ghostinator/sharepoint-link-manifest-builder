using Microsoft.Extensions.Logging;
using SharePointLinkManifestBuilder.Core.Abstractions;
using SharePointLinkManifestBuilder.Core.Models;

namespace SharePointLinkManifestBuilder.Graph.Onboarding;

/// <summary>
/// Runs Microsoft's official administrator-consent experience and verifies the outcome.
/// <para>
/// This application only ever <em>builds a URL</em> and opens the system browser. It renders no
/// consent-like screen of its own and accepts no credential, because a convincing in-app
/// imitation of the Microsoft consent page is exactly the shape of a consent-phishing attack.
/// </para>
/// <para>
/// Success is never inferred from the redirect. A redirect can carry an error, name the wrong
/// tenant, or reflect only a partial grant, so the result is confirmed by acquiring a real
/// token and comparing the scopes Entra actually issued. See
/// docs/adr/0006-consent-verification-by-token-acquisition.md.
/// </para>
/// </summary>
public sealed class ConsentService : IConsentService
{
    private readonly IAuthenticationService _authentication;
    private readonly ISystemBrowser _browser;
    private readonly IGraphApiClient _graphClient;
    private readonly ILogger<ConsentService> _logger;

    /// <summary>Creates the service.</summary>
    public ConsentService(
        IAuthenticationService authentication,
        ISystemBrowser browser,
        IGraphApiClient graphClient,
        ILogger<ConsentService> logger)
    {
        _authentication = authentication ?? throw new ArgumentNullException(nameof(authentication));
        _browser = browser ?? throw new ArgumentNullException(nameof(browser));
        _graphClient = graphClient ?? throw new ArgumentNullException(nameof(graphClient));
        _logger = logger ?? throw new ArgumentNullException(nameof(logger));
    }

    /// <summary>
    /// Determines which directory consent must be requested in.
    /// <para>
    /// A single-tenant configuration always names its own tenant. A multi-tenant one has no
    /// statically configured tenant, so the directory to consent in is the one the user is
    /// currently signed in to. Returns null when that cannot be determined, which callers must
    /// treat as "sign in first" rather than falling back to <c>/organizations</c>: an explicit
    /// tenant in the consent URL is precisely what stops an administrator who is signed into
    /// several directories from consenting in the wrong one.
    /// </para>
    /// </summary>
    internal static string? ResolveConsentTenantId(
        TenantConfiguration tenantConfiguration,
        string? explicitTenantId,
        UserAccount? signedInAccount)
    {
        if (!string.IsNullOrWhiteSpace(explicitTenantId))
        {
            return explicitTenantId.Trim();
        }

        if (!tenantConfiguration.IsMultiTenant)
        {
            return string.IsNullOrWhiteSpace(tenantConfiguration.TenantId)
                ? null
                : tenantConfiguration.TenantId;
        }

        return string.IsNullOrWhiteSpace(signedInAccount?.TenantId) ? null : signedInAccount.TenantId;
    }

    /// <inheritdoc />
    public Uri BuildAdminConsentUrl(
        TenantConfiguration tenantConfiguration,
        IReadOnlyList<PermissionRequirement> permissions,
        string redirectUri,
        string state,
        string? targetTenantId = null)
    {
        ArgumentNullException.ThrowIfNull(tenantConfiguration);
        ArgumentNullException.ThrowIfNull(permissions);
        ArgumentException.ThrowIfNullOrWhiteSpace(redirectUri);
        ArgumentException.ThrowIfNullOrWhiteSpace(state);

        var tenantId = ResolveConsentTenantId(
            tenantConfiguration, targetTenantId, _authentication.CurrentAccount)
            ?? throw new InvalidOperationException(
                "The directory to request consent in could not be determined. Sign in first so the "
                + "consent request names an explicit organization.");

        // The tenant-specific /v2.0/adminconsent endpoint. Naming the tenant explicitly rather
        // than using /common or /organizations is what keeps an administrator from accidentally
        // consenting in a different directory than the one being configured.
        var scope = GraphScopes.ToQualifiedScopeParameter(permissions, tenantConfiguration.GraphEndpoint);

        var query = string.Join('&',
        [
            $"client_id={Uri.EscapeDataString(tenantConfiguration.ClientId)}",
            $"scope={Uri.EscapeDataString(scope)}",
            $"redirect_uri={Uri.EscapeDataString(redirectUri)}",
            $"state={Uri.EscapeDataString(state)}",
        ]);

        var instance = tenantConfiguration.Instance.TrimEnd('/');
        return new Uri($"{instance}/{tenantId}/v2.0/adminconsent?{query}");
    }

    /// <inheritdoc />
    public async Task<ConsentOutcome> RequestAdminConsentAsync(
        TenantConfiguration tenantConfiguration,
        IReadOnlyList<PermissionRequirement> permissions,
        string? targetTenantId = null,
        CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(tenantConfiguration);
        ArgumentNullException.ThrowIfNull(permissions);

        var expectedTenantId = ResolveConsentTenantId(
            tenantConfiguration, targetTenantId, _authentication.CurrentAccount);

        if (expectedTenantId is null)
        {
            return new ConsentOutcome
            {
                Approved = false,
                Error = new GraphError
                {
                    Kind = GraphErrorKind.TenantMismatch,
                    Message = "The organization to request consent in is not known yet.",
                    SuggestedAction = "Sign in first. Consent is always requested in one named "
                        + "organization so it cannot be granted in the wrong directory.",
                },
            };
        }

        var state = LoopbackRedirectListener.GenerateState();

        using var listener = new LoopbackRedirectListener(_logger);
        var url = BuildAdminConsentUrl(
            tenantConfiguration, permissions, listener.RedirectUri, state, expectedTenantId);

        _logger.LogInformation(
            "Opening Microsoft's administrator consent experience in the system browser for tenant {TenantId}.",
            expectedTenantId);

        await _browser.OpenAsync(url, cancellationToken).ConfigureAwait(false);

        LoopbackRedirectResult redirect;

        try
        {
            redirect = await listener.WaitForRedirectAsync(state, cancellationToken).ConfigureAwait(false);
        }
        catch (OperationCanceledException)
        {
            return new ConsentOutcome
            {
                Approved = false,
                WasCancelled = true,
                Error = GraphError.Canceled(),
            };
        }

        if (redirect.StateMismatch)
        {
            return new ConsentOutcome
            {
                Approved = false,
                Error = new GraphError
                {
                    Kind = GraphErrorKind.AuthenticationFailed,
                    Message = "The consent response could not be verified and was rejected. Nothing was changed.",
                    SuggestedAction = "Start the consent step again from this application.",
                },
            };
        }

        // Guard against an administrator consenting in the wrong directory, which is easy to do
        // when signed into several.
        if (!string.IsNullOrEmpty(redirect.TenantId)
            && !string.Equals(redirect.TenantId, expectedTenantId, StringComparison.OrdinalIgnoreCase))
        {
            return new ConsentOutcome
            {
                Approved = false,
                ReturnedTenantId = redirect.TenantId,
                Error = new GraphError
                {
                    Kind = GraphErrorKind.TenantMismatch,
                    Message = "Consent was granted in a different Microsoft 365 organization than the one being "
                        + "configured, so it was not accepted.",
                    SuggestedAction = "Sign in as an administrator of the organization shown in the wizard.",
                },
            };
        }

        if (redirect.Error is not null)
        {
            return new ConsentOutcome
            {
                Approved = false,
                ReturnedTenantId = redirect.TenantId,
                Error = MapConsentError(redirect.Error, redirect.ErrorDescription),
            };
        }

        return new ConsentOutcome
        {
            Approved = redirect.AdminConsentGranted,
            ReturnedTenantId = redirect.TenantId,
        };
    }

    /// <inheritdoc />
    /// <summary>
    /// True when the failure means "this user has no cached grant", which an interactive sign-in
    /// can resolve, rather than "consent was refused", which it cannot.
    /// </summary>
    private static bool NeedsInteraction(GraphError? error) =>
        error?.Kind is GraphErrorKind.ConsentRequired
            or GraphErrorKind.AdminConsentRequired
            or GraphErrorKind.TokenExpired
            or GraphErrorKind.ConditionalAccessInterrupted;

    public async Task<RegistrationVerification> VerifyConsentAsync(
        TenantConfiguration tenantConfiguration,
        IReadOnlyList<PermissionRequirement> requiredPermissions,
        bool allowInteractive = false,
        CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(tenantConfiguration);
        ArgumentNullException.ThrowIfNull(requiredPermissions);

        var notVerified = new List<string>();
        var warnings = new List<string>();

        await _authentication.ConfigureAsync(tenantConfiguration, cancellationToken).ConfigureAwait(false);

        var scopes = requiredPermissions.Select(p => p.Scope).ToArray();

        // The heart of verification: attempt a real token acquisition and read back the scopes
        // Entra issued. This needs no directory permission and tests the thing that actually
        // matters, which is whether the application can obtain a usable token.
        //
        // Silent first, because when a grant is already cached this answers without disturbing
        // the user. But silent-only was wrong: a consent that an administrator has genuinely
        // granted is invisible to this process until some interactive sign-in records it for
        // this user, so AADSTS65001 came back and was reported as "not consented" no matter how
        // many times the user pressed Check again. Asking silently and concluding from silence
        // conflates "nobody has consented" with "this user has no cached grant yet".
        var token = await _authentication
            .AcquireTokenAsync(scopes, allowInteractive: false, cancellationToken)
            .ConfigureAwait(false);

        if (!token.Succeeded && allowInteractive && NeedsInteraction(token.Error))
        {
            _logger.LogInformation(
                "Silent verification needs interaction; retrying with an interactive sign-in.");

            token = await _authentication
                .AcquireTokenAsync(scopes, allowInteractive: true, cancellationToken)
                .ConfigureAwait(false);
        }

        if (!token.Succeeded)
        {
            var consentState = token.Error?.Kind switch
            {
                GraphErrorKind.AdminConsentRequired or GraphErrorKind.ConsentRequired =>
                    ConsentState.PendingAdministratorApproval,
                GraphErrorKind.ConsentDenied => ConsentState.Denied,
                _ => ConsentState.VerificationFailed,
            };

            return new RegistrationVerification
            {
                CanAcquireToken = false,
                GraphConnectivityOk = false,
                ConsentState = consentState,
                PermissionStates = requiredPermissions
                    .Select(p => new PermissionGrantState
                    {
                        Requirement = p,
                        IsGranted = false,
                        GrantSource = "Not granted at last check",
                    })
                    .ToArray(),
                Warnings = [token.Error?.Message ?? "A token could not be acquired."],
                NotVerified =
                [
                    "Graph connectivity was not tested, because no token could be obtained.",
                    "The application registration was not inspected, which needs an elevated read permission.",
                ],
            };
        }

        var granted = new HashSet<string>(token.GrantedScopes, StringComparer.OrdinalIgnoreCase);

        var states = requiredPermissions
            .Select(p => new PermissionGrantState
            {
                Requirement = p,
                IsGranted = granted.Contains(p.Scope)
                    || GraphScopes.Reserved.Contains(p.Scope, StringComparer.OrdinalIgnoreCase),
                GrantSource = "Confirmed by a token issued by Microsoft Entra",
            })
            .ToArray();

        var missing = states.Where(s => !s.IsGranted).ToArray();

        // A basic Graph call proves the token is not merely issued but actually accepted.
        var connectivity = await _graphClient
            .GetAsync<object>("/me?$select=id", cancellationToken)
            .ConfigureAwait(false);

        if (!connectivity.Succeeded)
        {
            warnings.Add(
                "A token was issued, but a test call to Microsoft Graph did not succeed: "
                + connectivity.Error?.Message);
        }

        // A token alone cannot distinguish tenant-wide admin consent from this user's own
        // consent. Saying so is more useful than presenting a guess as fact.
        notVerified.Add(
            "Whether consent was granted tenant-wide by an administrator or only by the signed-in user. "
            + "Determining this requires a directory read permission that this application does not request.");

        notVerified.Add(
            "The application registration and service principal objects were not inspected, which would need "
            + "an elevated directory read permission.");

        return new RegistrationVerification
        {
            CanAcquireToken = true,
            GraphConnectivityOk = connectivity.Succeeded,
            PermissionStates = states,
            ConsentState = missing.Length == 0
                ? ConsentState.Granted
                : ConsentState.PartiallyGranted,
            ConsentType = ConsentType.Unknown,
            Warnings = warnings,
            NotVerified = notVerified,
        };
    }

    /// <summary>Maps an OAuth error code from the consent redirect onto the normalized set.</summary>
    internal static GraphError MapConsentError(string error, string? description) => error switch
    {
        "access_denied" => new GraphError
        {
            Kind = GraphErrorKind.ConsentDenied,
            Message = "Consent was declined in the Microsoft consent experience.",
            GraphErrorCode = error,
        },

        "invalid_client" => new GraphError
        {
            Kind = GraphErrorKind.AppRegistrationNotFound,
            Message = "Microsoft Entra did not recognise the application. The registration may not exist "
                + "in this tenant yet.",
            GraphErrorCode = error,
            SuggestedAction = "Run Verification, or use Repair Registration.",
        },

        "unauthorized_client" or "insufficient_privileges" => new GraphError
        {
            Kind = GraphErrorKind.UnauthorizedAdministratorRole,
            Message = "The account used does not have permission to grant consent for the organization.",
            GraphErrorCode = error,
            SuggestedAction = "An authorized Microsoft Entra administrator must approve the request. "
                + "Use 'Copy Consent Link' to send it to one.",
        },

        "consent_required" or "interaction_required" => new GraphError
        {
            Kind = GraphErrorKind.AdminConsentRequired,
            Message = "Further approval is required before this application can be used.",
            GraphErrorCode = error,
        },

        _ => new GraphError
        {
            Kind = GraphErrorKind.ConsentRequired,

            // The description comes from Microsoft and is safe to show, but it is length-capped
            // so a long service message cannot swamp the dialog.
            Message = string.IsNullOrWhiteSpace(description)
                ? $"Consent did not complete ({error})."
                : Truncate(description, 400),
            GraphErrorCode = error,
        },
    };

    private static string Truncate(string value, int max) =>
        value.Length <= max ? value : value[..max] + "…";
}
