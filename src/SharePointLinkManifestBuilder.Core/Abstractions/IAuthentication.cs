using SharePointLinkManifestBuilder.Core.Models;

namespace SharePointLinkManifestBuilder.Core.Abstractions;

/// <summary>Whether OS-native secure storage is usable on this machine.</summary>
public enum SecureStorageAvailability
{
    /// <summary>Not yet probed.</summary>
    Unknown = 0,

    /// <summary>Available and verified by a real read/write round-trip.</summary>
    Available = 1,

    /// <summary>
    /// Unavailable. Tokens are kept in memory only and sign-in is required each launch.
    /// Tokens are never written in plaintext as a fallback.
    /// </summary>
    UnavailableUsingMemoryOnly = 2,
}

/// <summary>The outcome of probing secure storage, including why it is unavailable.</summary>
/// <param name="Availability">Whether secure storage can be used.</param>
/// <param name="Mechanism">The mechanism in use, for example "macOS Keychain".</param>
/// <param name="Detail">A user-facing explanation when storage is unavailable.</param>
public readonly record struct SecureStorageStatus(
    SecureStorageAvailability Availability,
    string Mechanism,
    string? Detail);

/// <summary>
/// Provides the MSAL token cache backed by OS-native secure storage, with an explicit
/// memory-only fallback. See docs/adr/0008-token-storage.md.
/// </summary>
public interface ISecureTokenStorage
{
    /// <summary>The status determined by the last probe.</summary>
    SecureStorageStatus Status { get; }

    /// <summary>
    /// Probes the platform store with a real round-trip. Never throws for an unavailable
    /// store: unavailability is a supported state, reported to the user.
    /// </summary>
    Task<SecureStorageStatus> ProbeAsync(CancellationToken cancellationToken = default);

    /// <summary>Removes all cached tokens from the store.</summary>
    Task ClearAsync(CancellationToken cancellationToken = default);
}

/// <summary>The result of a sign-in or token acquisition attempt.</summary>
public sealed record AuthenticationResultInfo
{
    /// <summary>True when a usable token was obtained.</summary>
    public required bool Succeeded { get; init; }

    /// <summary>The authenticated account, when sign-in succeeded.</summary>
    public UserAccount? Account { get; init; }

    /// <summary>
    /// The scopes Microsoft Entra actually issued. This is the basis of consent verification
    /// (see docs/adr/0006). The access token itself is never exposed here.
    /// </summary>
    public IReadOnlyList<string> GrantedScopes { get; init; } = [];

    /// <summary>When the token expires.</summary>
    public DateTimeOffset? ExpiresOnUtc { get; init; }

    /// <summary>The failure, when the attempt did not succeed.</summary>
    public GraphError? Error { get; init; }
}

/// <summary>
/// Delegated authentication against Microsoft Entra. Implemented with MSAL using
/// Authorization Code Flow with PKCE through the system browser. Never handles a password,
/// never embeds a secret, and never exposes a raw token to callers other than the Graph
/// transport.
/// </summary>
public interface IAuthenticationService
{
    /// <summary>The currently signed-in account, or null.</summary>
    UserAccount? CurrentAccount { get; }

    /// <summary>Raised when the signed-in account changes, including on sign-out.</summary>
    event EventHandler<UserAccount?>? AccountChanged;

    /// <summary>Configures the tenant and client this service authenticates against.</summary>
    Task ConfigureAsync(TenantConfiguration configuration, CancellationToken cancellationToken = default);

    /// <summary>
    /// Signs in interactively through the system browser. Supports MFA and Conditional Access
    /// because authentication happens in the real browser.
    /// </summary>
    /// <param name="scopes">Scopes to request.</param>
    /// <param name="loginHint">Optional account hint.</param>
    /// <param name="cancellationToken">Cancellation token.</param>
    Task<AuthenticationResultInfo> SignInAsync(
        IEnumerable<string> scopes,
        string? loginHint = null,
        CancellationToken cancellationToken = default);

    /// <summary>
    /// Acquires a token silently, falling back to interactive only when
    /// <paramref name="allowInteractive"/> is true. Used by the Graph transport per request.
    /// </summary>
    Task<AuthenticationResultInfo> AcquireTokenAsync(
        IEnumerable<string> scopes,
        bool allowInteractive = false,
        CancellationToken cancellationToken = default);

    /// <summary>
    /// Returns the bearer token for the given scopes, or null when none can be obtained
    /// without interaction. Only the Graph transport calls this.
    /// </summary>
    Task<string?> GetAccessTokenAsync(
        IEnumerable<string> scopes,
        CancellationToken cancellationToken = default);

    /// <summary>Accounts MSAL has cached, for the account switcher.</summary>
    Task<IReadOnlyList<UserAccount>> GetCachedAccountsAsync(CancellationToken cancellationToken = default);

    /// <summary>
    /// Switches to an already-cached account, silently when its refresh token is still valid and
    /// falling back to an interactive prompt when the target organization still needs consent.
    /// This is what makes moving between organizations a single gesture.
    /// </summary>
    /// <param name="homeAccountId">MSAL home account identifier of the account to activate.</param>
    /// <param name="scopes">Scopes to request.</param>
    /// <param name="cancellationToken">Cancellation token.</param>
    Task<AuthenticationResultInfo> SwitchToAccountAsync(
        string homeAccountId,
        IEnumerable<string> scopes,
        CancellationToken cancellationToken = default);

    /// <summary>Signs out and removes the account from the cache.</summary>
    Task SignOutAsync(CancellationToken cancellationToken = default);

    /// <summary>Removes a specific cached account.</summary>
    Task ForgetAccountAsync(string homeAccountId, CancellationToken cancellationToken = default);
}
