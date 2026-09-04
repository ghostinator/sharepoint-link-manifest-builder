using System.Text.RegularExpressions;
using Microsoft.Extensions.Logging;
using Microsoft.Identity.Client;
using SharePointLinkManifestBuilder.Core.Abstractions;
using SharePointLinkManifestBuilder.Core.Models;
using SharePointLinkManifestBuilder.Core.Security;
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
public sealed partial class MsalAuthenticationService : IAuthenticationService, IDisposable
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

            // A single-tenant configuration keeps the tenant-specific authority, which makes a
            // token from another tenant structurally impossible. A multi-tenant configuration
            // uses /organizations and moves that protection to an explicit check at sign-in:
            // the issuing tenant is resolved from the token and pinned for the session. It is
            // never /common, which would also admit personal Microsoft accounts that have no
            // SharePoint or OneDrive for Business at all. See
            // docs/adr/0011-multi-tenant-authority.md.
            _application = PublicClientApplicationBuilder
                .Create(configuration.ClientId)
                .WithAuthority(configuration.Authority, validateAuthority: true)
                .WithRedirectUri(AuthorityDefaults.LoopbackRedirectUri)
                .WithClientName("SharePointLinkManifestBuilder")
                .WithClientVersion(typeof(MsalAuthenticationService).Assembly.GetName().Version?.ToString() ?? "0.0.0")
                .Build();

            _tokenStorage.RegisterCache(_application);

            _logger.LogInformation(
                "Authentication configured for {Audience} using authority {Authority} "
                + "and a public client with no secret.",
                configuration.IsMultiTenant ? "any work or school organization" : $"tenant {configuration.TenantId}",
                configuration.Authority);
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
            return Failed(Classify(ex, "sign in"));
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
                    return Failed(Classify(ex, "acquire a token without interaction"));
                }

                _logger.LogInformation(
                    "Silent token acquisition needs interaction ({Classification}); prompting the user.",
                    ex.Classification);
            }
            catch (MsalException ex)
            {
                return Failed(Classify(ex, "acquire a token"));
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
    public async Task<AuthenticationResultInfo> SwitchToAccountAsync(
        string homeAccountId,
        IEnumerable<string> scopes,
        CancellationToken cancellationToken = default)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(homeAccountId);
        ArgumentNullException.ThrowIfNull(scopes);

        if (_application is null || _configuration is null)
        {
            return NotConfigured();
        }

        var requested = Normalize(scopes);
        var accounts = await _application.GetAccountsAsync().ConfigureAwait(false);
        var target = accounts.FirstOrDefault(a =>
            string.Equals(a.HomeAccountId?.Identifier, homeAccountId, StringComparison.Ordinal));

        if (target is null)
        {
            return Failed(new GraphError
            {
                Kind = GraphErrorKind.AuthenticationFailed,
                Message = "That account is no longer cached on this device.",
                SuggestedAction = "Add the account again.",
            });
        }

        try
        {
            // The switch is silent whenever the refresh token is still good, which is what makes
            // it one click. Each organization keeps its own tokens and its own consent state, so
            // switching never carries authorization from one organization into another.
            var silent = await _application
                .AcquireTokenSilent(requested, target)
                .ExecuteAsync(cancellationToken)
                .ConfigureAwait(false);

            return await AcceptResultAsync(silent).ConfigureAwait(false);
        }
        catch (MsalUiRequiredException ex)
        {
            _logger.LogInformation(
                "Switching accounts needs interaction ({Classification}); prompting the user.",
                ex.Classification);

            // Falling through to interactive with the account as the hint keeps the switch a
            // single user gesture even when consent for this organization is still missing.
            return await SignInAsync(requested, target.Username, cancellationToken).ConfigureAwait(false);
        }
        catch (MsalException ex)
        {
            return Failed(Classify(ex, "switch accounts"));
        }
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
        var tenantId = result.Account?.HomeAccountId?.TenantId ?? result.TenantId;

        // Cross-tenant confusion guard. For a single-tenant configuration a token from any
        // other tenant is refused outright. For a multi-tenant configuration there is no single
        // expected tenant by definition, so the issuing tenant is instead resolved here and
        // recorded on the account: every subsequent Graph call runs against the token for that
        // account, and site and drive identifiers are tenant-scoped, so the tenant in effect is
        // always the one the user chose. The account switcher makes that choice explicit.
        if (_configuration is { IsMultiTenant: false } singleTenant
            && !string.IsNullOrEmpty(tenantId)
            && !string.Equals(tenantId, singleTenant.TenantId, StringComparison.OrdinalIgnoreCase))
        {
            _logger.LogWarning("Rejecting a token issued for a different tenant than the configured one.");

            return Task.FromResult(Failed(new GraphError
            {
                Kind = GraphErrorKind.TenantMismatch,
                Message = "You signed in to a different Microsoft 365 organization than this application is "
                    + "configured for.",
                SuggestedAction = "Sign in with an account in the configured organization, or turn on "
                    + "multi-organization mode in the Microsoft 365 connection settings.",
            }));
        }

        if (_configuration is { IsMultiTenant: true })
        {
            _logger.LogInformation(
                "Signed in to organization {TenantId}. All subsequent requests use this organization.",
                tenantId ?? "(unknown)");
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

        // A multi-tenant configuration has no single "correct" tenant, so any cached account is
        // a candidate and the user's explicit choice (handled above) is what decides. A
        // single-tenant configuration prefers an account in the configured tenant; anything else
        // would be rejected by the guard in AcceptResultAsync anyway.
        if (_configuration is null or { IsMultiTenant: true })
        {
            return list.FirstOrDefault();
        }

        return list.FirstOrDefault(a =>
            string.Equals(a.HomeAccountId?.TenantId, _configuration.TenantId, StringComparison.OrdinalIgnoreCase))
            ?? list.FirstOrDefault();
    }

    /// <summary>
    /// Matches the <c>AADSTSnnnnn</c> code Microsoft Entra embeds in its error descriptions.
    /// That code is the only reliable discriminator between failures that share an OAuth
    /// error code — <c>invalid_client</c> alone cannot distinguish "not a public client" from
    /// several unrelated problems.
    /// </summary>
    // The upper bound is deliberately loose. Entra codes are not fixed width -- AADSTS50011
    // is five digits and AADSTS7000218 is seven -- and a bound that is too tight does not
    // fail to match, it matches a truncated prefix and misclassifies the error.
    [GeneratedRegex(@"AADSTS\d{4,8}", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant)]
    private static partial Regex EntraErrorCodeRegex();

    /// <summary>Extracts the <c>AADSTSnnnnn</c> code from an MSAL message, when present.</summary>
    internal static string? ExtractEntraErrorCode(string? message)
    {
        if (string.IsNullOrEmpty(message))
        {
            return null;
        }

        var match = EntraErrorCodeRegex().Match(message);

        return match.Success ? match.Value.ToUpperInvariant() : null;
    }

    /// <summary>
    /// Logs the full MSAL diagnostic, then normalizes it for the UI.
    /// <para>
    /// Both halves are necessary and they are not the same thing. <see cref="GraphError"/> is
    /// deliberately sanitized and user-facing, so on its own it discards exactly what is needed
    /// to diagnose a failure: the OAuth error code, the <c>AADSTS</c> number, and the
    /// correlation ID that Microsoft Support asks for. Normalizing without logging first turns
    /// every unclassified failure into an undiagnosable one.
    /// </para>
    /// </summary>
    private GraphError Classify(MsalException exception, string operation)
    {
        var entraCode = ExtractEntraErrorCode(exception.Message);
        var correlationId = exception is MsalServiceException { CorrelationId: { Length: > 0 } id } ? id : null;
        var statusCode = exception is MsalServiceException service ? service.StatusCode : 0;

        // The message is redacted defensively. MSAL puts the Entra error description here, not
        // token material, but this class must never be the reason a token reaches a log file.
        _logger.LogError(
            "Could not {Operation}. MSAL error code {ErrorCode}, Entra code {EntraCode}, "
            + "HTTP status {StatusCode}, correlation ID {CorrelationId}. Detail: {Detail}",
            operation,
            exception.ErrorCode ?? "(none)",
            entraCode ?? "(none)",
            statusCode,
            correlationId ?? "(none)",
            SensitiveDataRedactor.Redact(exception.Message));

        var error = MapMsalException(exception, operation);

        // Surface the Entra code in the UI when there is one. It is the difference between
        // "sign-in failed" and a specific, searchable cause the user can act on.
        return entraCode is null
            ? error
            : error with
            {
                GraphErrorCode = error.GraphErrorCode is { Length: > 0 } existing
                ? $"{existing} ({entraCode})"
                : entraCode
            };
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

        // Registration and authority misconfiguration, keyed on the AADSTS code because the
        // OAuth error code alone is ambiguous. Every failure below surfaces *after* the browser
        // has displayed "authentication complete": MSAL renders that page the moment the
        // redirect arrives, whether or not the redirect carries an error. These arms must stay
        // below the MsalUiRequiredException arms, because that type derives from
        // MsalServiceException and would otherwise be captured here.
        MsalServiceException service when ExtractEntraErrorCode(service.Message) == "AADSTS7000218" =>
            new GraphError
            {
                Kind = GraphErrorKind.PublicClientNotConfigured,
                Message = "Microsoft Entra asked this application for a client secret, which means "
                    + "the app registration is not marked as a public client.",
                GraphErrorCode = service.ErrorCode,
                SuggestedAction = "In the app registration, open Authentication, set 'Allow public "
                    + "client flows' to Yes, and register the redirect URI http://localhost under "
                    + "'Mobile and desktop applications' rather than 'Web'. This application has no "
                    + "client secret by design and must never be given one.",
            },

        MsalServiceException service when ExtractEntraErrorCode(service.Message)
            is "AADSTS50011" or "AADSTS900971" =>
            new GraphError
            {
                Kind = GraphErrorKind.RedirectUriMismatch,
                Message = "The redirect address this application used is not registered on the app "
                    + "registration.",
                GraphErrorCode = service.ErrorCode,
                SuggestedAction = "In the app registration, open Authentication and add the redirect "
                    + "URI http://localhost under 'Mobile and desktop applications'. No port is "
                    + "needed: loopback redirects match on any port.",
            },

        MsalServiceException service when ExtractEntraErrorCode(service.Message) == "AADSTS50020" =>
            new GraphError
            {
                Kind = GraphErrorKind.AccountFromUnsupportedTenant,
                Message = "The account you signed in with belongs to a different organization than "
                    + "the one this app registration accepts.",
                GraphErrorCode = service.ErrorCode,
                SuggestedAction = "Either sign in with an account in the configured organization, or "
                    + "set the registration to accept any work or school organization and enable "
                    + "multi-organization mode in the Microsoft 365 connection settings.",
            },

        MsalServiceException service when ExtractEntraErrorCode(service.Message)
            is "AADSTS700016" or "AADSTS90002" =>
            new GraphError
            {
                Kind = GraphErrorKind.ApplicationNotFoundInTenant,
                Message = "This app registration does not exist in the organization that was signed "
                    + "in to.",
                GraphErrorCode = service.ErrorCode,
                SuggestedAction = "Check the client ID, and confirm the organization matches the one "
                    + "the registration was created in. A single-organization registration cannot be "
                    + "used from another organization.",
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
