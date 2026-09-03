using Microsoft.Extensions.Logging;
using Microsoft.Identity.Client;
using SharePointLinkManifestBuilder.Core.Abstractions;
using SharePointLinkManifestBuilder.Core.Models;
using SharePointLinkManifestBuilder.Graph.Http;

namespace SharePointLinkManifestBuilder.Graph.Identity;

/// <summary>
/// Delegated authentication through MSAL.
/// <para>
/// Uses Authorization Code Flow with PKCE against the <em>system browser</em>. That choice is
/// what makes MFA, Conditional Access and device compliance work, and it means this application
/// never sees a credential. There is deliberately no embedded web view: an in-app browser can
/// observe what the user types, which is precisely the consent-phishing shape this product must
/// not have. See docs/adr/0004-public-client-pkce-no-secret.md.
/// </para>
/// </summary>
public sealed class MsalAuthenticationService : IAuthenticationService, IDisposable
{
    private readonly SecureTokenStorage _tokenStorage;
    private readonly ILogger<MsalAuthenticationService> _logger;
    private readonly SemaphoreSlim _gate = new(1, 1);

    private IPublicClientApplication? _application;
    private TenantConfiguration? _configuration;
    private UserAccount? _currentAccount;

    /// <summary>Creates the service.</summary>
    public MsalAuthenticationService(
        SecureTokenStorage tokenStorage,
        ILogger<MsalAuthenticationService> logger)
    {
        _tokenStorage = tokenStorage ?? throw new ArgumentNullException(nameof(tokenStorage));
        _logger = logger ?? throw new ArgumentNullException(nameof(logger));
    }

    /// <inheritdoc />
    public UserAccount? CurrentAccount => _currentAccount;

    /// <inheritdoc />
    public event EventHandler<UserAccount?>? AccountChanged;

    /// <inheritdoc />
    public async Task ConfigureAsync(
        TenantConfiguration configuration,
        CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(configuration);

        await _gate.WaitAsync(cancellationToken).ConfigureAwait(false);

        try
        {
            await _tokenStorage.ProbeAsync(cancellationToken).ConfigureAwait(false);

            _configuration = configuration;

            // The authority is always tenant-specific, never /common or /organizations. A
            // tenant-specific authority is what prevents a token from another tenant being
            // accepted here, which is the cross-tenant confusion threat in the threat model.
            _application = PublicClientApplicationBuilder
                .Create(configuration.ClientId)
                .WithAuthority($"{configuration.Instance.TrimEnd('/')}/{configuration.TenantId}", validateAuthority: true)
                .WithRedirectUri(AuthorityDefaults.LoopbackRedirectUri)
                .WithClientName("SharePointLinkManifestBuilder")
                .WithClientVersion(typeof(MsalAuthenticationService).Assembly.GetName().Version?.ToString() ?? "0.0.0")
                .Build();

            _tokenStorage.RegisterCache(_application);

            _logger.LogInformation(
                "Authentication configured for tenant {TenantId} using a public client with no secret.",
                configuration.TenantId);
        }
        finally
        {
            _gate.Release();
        }
    }

    /// <inheritdoc />
    public async Task<AuthenticationResultInfo> SignInAsync(
        IEnumerable<string> scopes,
        string? loginHint = null,
        CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(scopes);

        if (_application is null || _configuration is null)
        {
            return NotConfigured();
        }

        var requested = Normalize(scopes);

        try
        {
            var builder = _application
                .AcquireTokenInteractive(requested)

                // Force the system browser. MSAL would otherwise be free to choose an embedded
                // control on some platforms.
                .WithUseEmbeddedWebView(false)
                .WithPrompt(Prompt.SelectAccount);

            if (!string.IsNullOrWhiteSpace(loginHint))
            {
                builder = builder.WithLoginHint(loginHint);
            }

            var result = await builder.ExecuteAsync(cancellationToken).ConfigureAwait(false);

            return await AcceptResultAsync(result).ConfigureAwait(false);
        }
        catch (MsalException ex)
        {
            return Failed(MapMsalException(ex, "sign in"));
        }
        catch (OperationCanceledException)
        {
            return Failed(GraphError.Canceled());
        }
    }

    /// <inheritdoc />
    public async Task<AuthenticationResultInfo> AcquireTokenAsync(
        IEnumerable<string> scopes,
        bool allowInteractive = false,
        CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(scopes);

        if (_application is null || _configuration is null)
        {
            return NotConfigured();
        }

        var requested = Normalize(scopes);
        var accounts = await _application.GetAccountsAsync().ConfigureAwait(false);
        var account = SelectAccount(accounts);

        if (account is not null)
        {
            try
            {
                var silent = await _application
                    .AcquireTokenSilent(requested, account)
                    .ExecuteAsync(cancellationToken)
                    .ConfigureAwait(false);

                return await AcceptResultAsync(silent).ConfigureAwait(false);
            }
            catch (MsalUiRequiredException ex)
            {
                // The expected path when consent is missing, the token expired beyond refresh,
                // or Conditional Access demands interaction. It is a normal state, not a bug.
                if (!allowInteractive)
                {
                    return Failed(MapMsalException(ex, "acquire a token without interaction"));
                }

                _logger.LogInformation(
                    "Silent token acquisition needs interaction ({Classification}); prompting the user.",
                    ex.Classification);
            }
            catch (MsalException ex)
            {
                return Failed(MapMsalException(ex, "acquire a token"));
            }
        }
        else if (!allowInteractive)
        {
            return Failed(new GraphError
            {
                Kind = GraphErrorKind.AuthenticationFailed,
                Message = "No signed-in account is available.",
                SuggestedAction = "Sign in from the Microsoft 365 connection settings.",
            });
        }

        return await SignInAsync(requested, account?.Username, cancellationToken).ConfigureAwait(false);
    }

    /// <inheritdoc />
    public async Task<string?> GetAccessTokenAsync(
        IEnumerable<string> scopes,
        CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(scopes);

        if (_application is null)
        {
            return null;
        }

        var requested = Normalize(scopes);
        var accounts = await _application.GetAccountsAsync().ConfigureAwait(false);
        var account = SelectAccount(accounts);

        if (account is null)
        {
            return null;
        }

        try
        {
            var result = await _application
                .AcquireTokenSilent(requested, account)
                .ExecuteAsync(cancellationToken)
                .ConfigureAwait(false);

            // The token is handed straight to the transport and is never logged, cached by this
            // class, or exposed on any public surface.
            return result.AccessToken;
        }
        catch (MsalUiRequiredException)
        {
            // Returning null lets the transport produce a clean "sign in again" error rather
            // than surfacing an MSAL exception to the UI.
            return null;
        }
        catch (MsalException ex)
        {
            _logger.LogWarning(
                "Silent token acquisition failed with {ErrorCode}; the caller will be asked to sign in again.",
                ex.ErrorCode);

            return null;
        }
    }

    /// <inheritdoc />
    public async Task<IReadOnlyList<UserAccount>> GetCachedAccountsAsync(
        CancellationToken cancellationToken = default)
    {
        if (_application is null || _configuration is null)
        {
            return [];
        }

        var accounts = await _application.GetAccountsAsync().ConfigureAwait(false);

        return accounts.Select(a => new UserAccount
        {
            UserId = a.HomeAccountId?.ObjectId ?? a.Username,
            DisplayName = a.Username,
            UserPrincipalName = a.Username,
            TenantId = a.HomeAccountId?.TenantId ?? _configuration.TenantId,
            HomeAccountId = a.HomeAccountId?.Identifier,
        }).ToArray();
    }

    /// <inheritdoc />
    public async Task SignOutAsync(CancellationToken cancellationToken = default)
    {
        if (_application is null)
        {
            return;
        }

        foreach (var account in await _application.GetAccountsAsync().ConfigureAwait(false))
        {
            await _application.RemoveAsync(account).ConfigureAwait(false);
        }

        _currentAccount = null;
        AccountChanged?.Invoke(this, null);
        _logger.LogInformation("Signed out and removed all cached accounts.");
    }

    /// <inheritdoc />
    public async Task ForgetAccountAsync(string homeAccountId, CancellationToken cancellationToken = default)
    {
        if (_application is null)
        {
            return;
        }

        var accounts = await _application.GetAccountsAsync().ConfigureAwait(false);
        var match = accounts.FirstOrDefault(a =>
            string.Equals(a.HomeAccountId?.Identifier, homeAccountId, StringComparison.Ordinal));

        if (match is null)
        {
            return;
        }

        await _application.RemoveAsync(match).ConfigureAwait(false);

        if (string.Equals(_currentAccount?.HomeAccountId, homeAccountId, StringComparison.Ordinal))
        {
            _currentAccount = null;
            AccountChanged?.Invoke(this, null);
        }
    }

    /// <summary>Releases the synchronization primitive.</summary>
    public void Dispose() => _gate.Dispose();

    private Task<AuthenticationResultInfo> AcceptResultAsync(AuthenticationResult result)
    {
        // Cross-tenant confusion guard: a token from a tenant other than the configured one is
        // refused outright rather than used.
        var tenantId = result.Account?.HomeAccountId?.TenantId ?? result.TenantId;

        if (_configuration is not null
            && !string.IsNullOrEmpty(tenantId)
            && !string.Equals(tenantId, _configuration.TenantId, StringComparison.OrdinalIgnoreCase))
        {
            _logger.LogWarning("Rejecting a token issued for a different tenant than the configured one.");

            return Task.FromResult(Failed(new GraphError
            {
                Kind = GraphErrorKind.TenantMismatch,
                Message = "You signed in to a different Microsoft 365 organization than this application is "
                    + "configured for.",
                SuggestedAction = "Sign in with an account in the configured organization, or change tenant "
                    + "in the Microsoft 365 connection settings.",
            }));
        }

        var account = new UserAccount
        {
            UserId = result.Account?.HomeAccountId?.ObjectId ?? result.UniqueId ?? "unknown",
            DisplayName = result.Account?.Username ?? "Signed-in user",
            UserPrincipalName = result.Account?.Username ?? string.Empty,
            TenantId = tenantId ?? _configuration?.TenantId ?? string.Empty,
            HomeAccountId = result.Account?.HomeAccountId?.Identifier,
        };

        _currentAccount = account;
        AccountChanged?.Invoke(this, account);

        return Task.FromResult(new AuthenticationResultInfo
        {
            Succeeded = true,
            Account = account,

            // These are the scopes Microsoft Entra actually issued, which is the basis of
            // consent verification. The access token itself is never surfaced here.
            GrantedScopes = result.Scopes?.ToArray() ?? [],
            ExpiresOnUtc = result.ExpiresOn,
        });
    }

    private IAccount? SelectAccount(IEnumerable<IAccount> accounts)
    {
        var list = accounts.ToArray();

        if (_currentAccount?.HomeAccountId is { } current)
        {
            var match = list.FirstOrDefault(a =>
                string.Equals(a.HomeAccountId?.Identifier, current, StringComparison.Ordinal));

            if (match is not null)
            {
                return match;
            }
        }

        // Prefer an account in the configured tenant; otherwise there is nothing usable.
        return list.FirstOrDefault(a =>
            _configuration is null
            || string.Equals(a.HomeAccountId?.TenantId, _configuration.TenantId, StringComparison.OrdinalIgnoreCase))
            ?? list.FirstOrDefault();
    }

    /// <summary>
    /// Maps an MSAL failure onto the normalized error set. The distinctions matter: "an
    /// administrator must approve this" and "your sign-in expired" call for completely
    /// different user actions.
    /// </summary>
    internal static GraphError MapMsalException(MsalException exception, string operation) => exception switch
    {
        MsalUiRequiredException ui when ui.Classification == UiRequiredExceptionClassification.ConsentRequired =>
            new GraphError
            {
                Kind = GraphErrorKind.ConsentRequired,
                Message = "This application needs consent for the permissions it requires.",
                GraphErrorCode = ui.ErrorCode,
                SuggestedAction = "Open Settings, then Microsoft 365 Connection, then Permissions, and grant consent.",
            },

        MsalUiRequiredException ui when ui.ErrorCode == "invalid_grant"
            && ui.Message.Contains("AADSTS65001", StringComparison.OrdinalIgnoreCase) =>
            new GraphError
            {
                Kind = GraphErrorKind.AdminConsentRequired,
                Message = "An authorized Microsoft Entra administrator must approve the requested permissions.",
                GraphErrorCode = ui.ErrorCode,
                SuggestedAction = "Use 'Request Missing Consent' to send an administrator the consent link.",
            },

        MsalUiRequiredException ui when ui.Classification
            is UiRequiredExceptionClassification.BasicAction
            or UiRequiredExceptionClassification.AdditionalAction =>
            new GraphError
            {
                Kind = GraphErrorKind.ConditionalAccessInterrupted,
                Message = "Your organization requires an additional step, such as multi-factor authentication, "
                    + "before this application can continue.",
                GraphErrorCode = ui.ErrorCode,
                SuggestedAction = "Sign in again and complete the prompts your organization requires.",
            },

        MsalUiRequiredException ui => new GraphError
        {
            Kind = GraphErrorKind.TokenExpired,
            Message = "Your sign-in has expired and could not be renewed without interaction.",
            GraphErrorCode = ui.ErrorCode,
            SuggestedAction = "Sign in again.",
        },

        MsalServiceException service when service.ErrorCode == "access_denied" => new GraphError
        {
            Kind = GraphErrorKind.ConsentDenied,
            Message = "The request was declined in the Microsoft sign-in experience.",
            GraphErrorCode = service.ErrorCode,
        },

        MsalServiceException service when service.StatusCode == 429 => new GraphError
        {
            Kind = GraphErrorKind.Throttled,
            Message = "Microsoft Entra is asking this application to slow down.",
            IsRetryable = true,
        },

        MsalClientException client when client.ErrorCode == "authentication_canceled" => GraphError.Canceled(),

        MsalClientException client => new GraphError
        {
            Kind = GraphErrorKind.AuthenticationFailed,
            Message = $"Sign-in could not be completed on this device while trying to {operation}.",
            GraphErrorCode = client.ErrorCode,
        },

        _ => new GraphError
        {
            Kind = GraphErrorKind.AuthenticationFailed,
            Message = $"Microsoft Entra refused the request to {operation}.",
            GraphErrorCode = exception.ErrorCode,
        },
    };

    private static AuthenticationResultInfo NotConfigured() => Failed(new GraphError
    {
        Kind = GraphErrorKind.AuthenticationFailed,
        Message = "No Microsoft 365 tenant is configured yet.",
        SuggestedAction = "Run the setup wizard from Settings, then Microsoft 365 Connection.",
    });

    private static AuthenticationResultInfo Failed(GraphError error) =>
        new() { Succeeded = false, Error = error };

    /// <summary>
    /// Removes the reserved OIDC scopes. MSAL adds <c>openid</c>, <c>profile</c> and
    /// <c>offline_access</c> itself and rejects them if they are passed in explicitly.
    /// </summary>
    internal static string[] Normalize(IEnumerable<string> scopes) =>
        scopes
            .Where(s => !string.IsNullOrWhiteSpace(s))
            .Where(s => !GraphScopes.Reserved.Contains(s, StringComparer.OrdinalIgnoreCase))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToArray();
}
