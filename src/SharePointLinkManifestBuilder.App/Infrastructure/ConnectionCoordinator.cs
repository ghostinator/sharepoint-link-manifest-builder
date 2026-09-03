using Microsoft.Extensions.Logging;
using SharePointLinkManifestBuilder.Core.Abstractions;
using SharePointLinkManifestBuilder.Core.Models;
using SharePointLinkManifestBuilder.Graph.Dto;
using SharePointLinkManifestBuilder.Graph.Http;
using SharePointLinkManifestBuilder.Graph.Services;

namespace SharePointLinkManifestBuilder.App.Infrastructure;

/// <summary>The application's Microsoft 365 connection state, as the UI understands it.</summary>
public enum ConnectionState
{
    /// <summary>No tenant configuration exists; the setup wizard is required.</summary>
    NotConfigured = 0,

    /// <summary>A tenant is configured but nobody is signed in.</summary>
    ConfiguredSignedOut = 1,

    /// <summary>Signed in and usable.</summary>
    Connected = 2,

    /// <summary>Signed in, but one or more required permissions are missing.</summary>
    ConnectedWithMissingPermissions = 3,

    /// <summary>Configuration exists but is waiting for an administrator to approve consent.</summary>
    PendingAdministratorConsent = 4,
}

/// <summary>
/// Owns the application's connection to Microsoft 365 and keeps every dependent piece in step.
/// <para>
/// Tenant configuration, the MSAL application, the Graph transport's endpoint and scopes, and
/// the signed-in account all have to change together. Scattering that across view models is how
/// an application ends up authenticating against one tenant while querying another, so it is
/// coordinated in one place.
/// </para>
/// </summary>
public sealed class ConnectionCoordinator
{
    private readonly IAuthenticationService _authentication;
    private readonly ITenantConfigurationStore _tenantStore;
    private readonly ISecureTokenStorage _tokenStorage;
    private readonly GraphClientContext _graphContext;
    private readonly IGraphApiClient _graphClient;
    private readonly ILogger<ConnectionCoordinator> _logger;

    /// <summary>Creates the coordinator.</summary>
    public ConnectionCoordinator(
        IAuthenticationService authentication,
        ITenantConfigurationStore tenantStore,
        ISecureTokenStorage tokenStorage,
        GraphClientContext graphContext,
        IGraphApiClient graphClient,
        ILogger<ConnectionCoordinator> logger)
    {
        _authentication = authentication ?? throw new ArgumentNullException(nameof(authentication));
        _tenantStore = tenantStore ?? throw new ArgumentNullException(nameof(tenantStore));
        _tokenStorage = tokenStorage ?? throw new ArgumentNullException(nameof(tokenStorage));
        _graphContext = graphContext ?? throw new ArgumentNullException(nameof(graphContext));
        _graphClient = graphClient ?? throw new ArgumentNullException(nameof(graphClient));
        _logger = logger ?? throw new ArgumentNullException(nameof(logger));
    }

    /// <summary>The tenant configuration in force, or null when none is configured.</summary>
    public TenantConfiguration? Tenant { get; private set; }

    /// <summary>The signed-in account, or null.</summary>
    public UserAccount? Account => _authentication.CurrentAccount;

    /// <summary>The current connection state.</summary>
    public ConnectionState State { get; private set; } = ConnectionState.NotConfigured;

    /// <summary>Tenant display name resolved from Graph, when the tenant permits reading it.</summary>
    public string? TenantDisplayName { get; private set; }

    /// <summary>Raised whenever any part of the connection state changes.</summary>
    public event EventHandler? ConnectionChanged;

    /// <summary>
    /// Loads stored configuration and attempts a silent sign-in. Called once at startup, and
    /// again after the setup wizard completes.
    /// </summary>
    public async Task InitializeAsync(CancellationToken cancellationToken = default)
    {
        await _tokenStorage.ProbeAsync(cancellationToken).ConfigureAwait(false);

        Tenant = await _tenantStore.LoadAsync(cancellationToken).ConfigureAwait(false);

        if (Tenant is null || !Tenant.IsUsable)
        {
            State = ConnectionState.NotConfigured;
            RaiseChanged();
            return;
        }

        await ApplyTenantAsync(Tenant, cancellationToken).ConfigureAwait(false);

        // Silent only. Startup must never pop a browser window at the user unprompted.
        var token = await _authentication
            .AcquireTokenAsync(Tenant.RequiredScopes, allowInteractive: false, cancellationToken)
            .ConfigureAwait(false);

        if (token.Succeeded)
        {
            await OnSignedInAsync(token, cancellationToken).ConfigureAwait(false);
        }
        else
        {
            State = Tenant.ConsentState == ConsentState.PendingAdministratorApproval
                ? ConnectionState.PendingAdministratorConsent
                : ConnectionState.ConfiguredSignedOut;

            RaiseChanged();
        }
    }

    /// <summary>Points the whole application at a tenant configuration.</summary>
    public async Task ApplyTenantAsync(
        TenantConfiguration configuration,
        CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(configuration);

        Tenant = configuration;

        var scopes = configuration.RequiredScopes.Count > 0
            ? configuration.RequiredScopes
            : GraphScopes.StandardTier.Select(p => p.Scope).ToArray();

        // Endpoint and scopes are updated before authentication, so the very first token is
        // requested for the right resource.
        _graphContext.Update(configuration.GraphEndpoint, scopes);

        await _authentication.ConfigureAsync(configuration, cancellationToken).ConfigureAwait(false);

        RaiseChanged();
    }

    /// <summary>Signs in interactively through the system browser.</summary>
    public async Task<AuthenticationResultInfo> SignInAsync(CancellationToken cancellationToken = default)
    {
        if (Tenant is null)
        {
            return new AuthenticationResultInfo
            {
                Succeeded = false,
                Error = new GraphError
                {
                    Kind = GraphErrorKind.AuthenticationFailed,
                    Message = "No Microsoft 365 tenant is configured yet.",
                    SuggestedAction = "Run the setup wizard from Tenant Setup.",
                },
            };
        }

        var result = await _authentication
            .SignInAsync(Tenant.RequiredScopes, cancellationToken: cancellationToken)
            .ConfigureAwait(false);

        if (result.Succeeded)
        {
            await OnSignedInAsync(result, cancellationToken).ConfigureAwait(false);
        }

        return result;
    }

    /// <summary>Signs out and clears the resolved tenant name.</summary>
    public async Task SignOutAsync(CancellationToken cancellationToken = default)
    {
        await _authentication.SignOutAsync(cancellationToken).ConfigureAwait(false);

        State = Tenant is null ? ConnectionState.NotConfigured : ConnectionState.ConfiguredSignedOut;
        TenantDisplayName = null;

        RaiseChanged();
    }

    /// <summary>Removes the local configuration. Nothing in the tenant is changed.</summary>
    public async Task RemoveLocalConfigurationAsync(CancellationToken cancellationToken = default)
    {
        await _authentication.SignOutAsync(cancellationToken).ConfigureAwait(false);
        await _tokenStorage.ClearAsync(cancellationToken).ConfigureAwait(false);
        await _tenantStore.RemoveAsync(cancellationToken).ConfigureAwait(false);

        Tenant = null;
        TenantDisplayName = null;
        State = ConnectionState.NotConfigured;

        _logger.LogInformation(
            "Local Microsoft 365 configuration removed. The application registration in the tenant was "
            + "not touched.");

        RaiseChanged();
    }

    /// <summary>Persists an updated tenant configuration and re-applies it.</summary>
    public async Task SaveTenantAsync(
        TenantConfiguration configuration,
        CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(configuration);

        await _tenantStore.SaveAsync(configuration, cancellationToken).ConfigureAwait(false);
        await ApplyTenantAsync(configuration, cancellationToken).ConfigureAwait(false);
    }

    private async Task OnSignedInAsync(
        AuthenticationResultInfo result,
        CancellationToken cancellationToken)
    {
        if (Tenant is null)
        {
            return;
        }

        var granted = result.GrantedScopes;

        var missing = Tenant.RequiredScopes
            .Where(s => !granted.Contains(s, StringComparer.OrdinalIgnoreCase))
            .Where(s => !GraphScopes.Reserved.Contains(s, StringComparer.OrdinalIgnoreCase))
            .ToArray();

        Tenant = Tenant with
        {
            GrantedScopes = granted,
            LastVerifiedUtc = DateTimeOffset.UtcNow,
            ConsentState = missing.Length == 0 ? ConsentState.Granted : ConsentState.PartiallyGranted,
        };

        await _tenantStore.SaveAsync(Tenant, cancellationToken).ConfigureAwait(false);

        State = missing.Length == 0
            ? ConnectionState.Connected
            : ConnectionState.ConnectedWithMissingPermissions;

        await ResolveTenantNameAsync(cancellationToken).ConfigureAwait(false);

        RaiseChanged();
    }

    /// <summary>
    /// Resolves a friendly tenant name. Restricted tenants deny this, which is not an error:
    /// the UI falls back to the tenant ID rather than inventing a name.
    /// </summary>
    private async Task ResolveTenantNameAsync(CancellationToken cancellationToken)
    {
        var response = await _graphClient
            .GetAsync<GraphOrganizationCollection>(GraphPaths.Organization(), cancellationToken)
            .ConfigureAwait(false);

        var organizations = response.Value?.Value;
        var organization = organizations is { Count: > 0 } ? organizations[0] : null;

        if (response.Succeeded && organization?.DisplayName is { Length: > 0 } name)
        {
            TenantDisplayName = name;

            if (Tenant is not null)
            {
                Tenant = Tenant with { TenantDisplayName = name };
            }

            return;
        }

        _logger.LogInformation(
            "The organization name could not be read; the tenant ID will be shown instead.");

        TenantDisplayName = null;
    }

    private void RaiseChanged() => ConnectionChanged?.Invoke(this, EventArgs.Empty);

    /// <summary>The collection envelope returned by <c>GET /organization</c>.</summary>
    private sealed record GraphOrganizationCollection
    {
        public IReadOnlyList<GraphOrganizationDto>? Value { get; init; }
    }
}
