using Microsoft.Extensions.Logging;
using Microsoft.Identity.Client;
using Microsoft.Identity.Client.Extensions.Msal;
using SharePointLinkManifestBuilder.Core.Abstractions;

namespace SharePointLinkManifestBuilder.Graph.Identity;

/// <summary>
/// Places the MSAL token cache in OS-native secure storage, with an explicit memory-only
/// fallback when that is unavailable.
/// <para>
/// The store is probed with a real persistence check rather than assumed to work. A headless
/// Linux session with no keyring is the common failure, and the correct response is to tell the
/// user and keep tokens in memory — never to quietly write them to a plaintext file.
/// See docs/adr/0008-token-storage.md.
/// </para>
/// </summary>
public sealed class SecureTokenStorage : ISecureTokenStorage, IDisposable
{
    /// <summary>Cache file name. On Windows this file is DPAPI-protected.</summary>
    public const string CacheFileName = "msal.cache";

    /// <summary>Keychain service name used on macOS.</summary>
    public const string MacKeyChainServiceName = "SharePointLinkManifestBuilder";

    /// <summary>Keychain account name used on macOS.</summary>
    public const string MacKeyChainAccountName = "MsalTokenCache";

    /// <summary>Schema name used for the Linux Secret Service keyring.</summary>
    public const string LinuxKeyRingSchema = "com.example.placeholder.sharepointlinkmanifestbuilder";

    /// <summary>Collection used for the Linux Secret Service keyring.</summary>
    public const string LinuxKeyRingCollection = MsalCacheHelper.LinuxKeyRingDefaultCollection;

    private readonly string _cacheDirectory;
    private readonly ILogger<SecureTokenStorage> _logger;
    private readonly SemaphoreSlim _initializationGate = new(1, 1);

    private MsalCacheHelper? _cacheHelper;
    private SecureStorageStatus _status =
        new(SecureStorageAvailability.Unknown, DescribeMechanism(), null);

    /// <summary>Creates the storage.</summary>
    /// <param name="cacheDirectory">Directory holding the cache file, on platforms that use one.</param>
    /// <param name="logger">Logger. Never receives token material.</param>
    public SecureTokenStorage(string cacheDirectory, ILogger<SecureTokenStorage> logger)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(cacheDirectory);

        _cacheDirectory = cacheDirectory;
        _logger = logger ?? throw new ArgumentNullException(nameof(logger));
    }

    /// <inheritdoc />
    public SecureStorageStatus Status => _status;

    /// <inheritdoc />
    public async Task<SecureStorageStatus> ProbeAsync(CancellationToken cancellationToken = default)
    {
        await _initializationGate.WaitAsync(cancellationToken).ConfigureAwait(false);

        try
        {
            if (_status.Availability != SecureStorageAvailability.Unknown)
            {
                return _status;
            }

            try
            {
                Directory.CreateDirectory(_cacheDirectory);

                var properties = new StorageCreationPropertiesBuilder(CacheFileName, _cacheDirectory)
                    .WithMacKeyChain(MacKeyChainServiceName, MacKeyChainAccountName)
                    .WithLinuxKeyring(
                        LinuxKeyRingSchema,
                        LinuxKeyRingCollection,
                        "MSAL token cache for SharePoint Link Manifest Builder",
                        new KeyValuePair<string, string>("Product", "SharePointLinkManifestBuilder"),
                        new KeyValuePair<string, string>("Component", "TokenCache"))
                    .Build();

                var helper = await MsalCacheHelper.CreateAsync(properties).ConfigureAwait(false);

                // A real round-trip through the platform store. Constructing the helper alone
                // proves nothing; this is what actually fails on a keyring-less Linux session.
                helper.VerifyPersistence();

                _cacheHelper = helper;
                _status = new SecureStorageStatus(
                    SecureStorageAvailability.Available, DescribeMechanism(), null);

                _logger.LogInformation(
                    "Secure token storage is available using {Mechanism}.", _status.Mechanism);
            }
#pragma warning disable CA1031 // Any storage failure must degrade safely rather than crash.
            catch (Exception ex)
            {
                _cacheHelper = null;
                _status = new SecureStorageStatus(
                    SecureStorageAvailability.UnavailableUsingMemoryOnly,
                    DescribeMechanism(),
                    "Secure storage could not be initialised, so sign-in details are kept in memory only "
                    + "and you will need to sign in again each time the application starts. "
                    + "Tokens are never written to disk unprotected.");

                _logger.LogWarning(
                    ex,
                    "Secure token storage is unavailable; falling back to a memory-only cache. "
                    + "No token will be written to disk.");
            }
#pragma warning restore CA1031

            return _status;
        }
        finally
        {
            _initializationGate.Release();
        }
    }

    /// <summary>
    /// Attaches the cache to an MSAL application. When secure storage is unavailable this does
    /// nothing, leaving MSAL's default in-memory cache in place.
    /// </summary>
    public void RegisterCache(IPublicClientApplication application)
    {
        ArgumentNullException.ThrowIfNull(application);

        if (_cacheHelper is null)
        {
            _logger.LogInformation(
                "Using an in-memory token cache; sign-in will be required again next launch.");
            return;
        }

        _cacheHelper.RegisterCache(application.UserTokenCache);
    }

    /// <inheritdoc />
    /// <remarks>
    /// This is the explicit, user-initiated "forget everything on this machine" action, not a
    /// sign-out. Ordinary sign-out goes through
    /// <see cref="IAuthenticationService.SignOutAsync"/>, which removes each account
    /// individually as Microsoft recommends. This method exists for the case where the user
    /// wants no trace of the cache left behind at all.
    /// </remarks>
    public Task ClearAsync(CancellationToken cancellationToken = default)
    {
#pragma warning disable CS0618 // The obsoletion warns against using Clear() as a logout
        // mechanism. This is not logout: it is the deliberate "Clear Token Cache" command, and
        // per-account removal has already run via SignOutAsync before this point.
        _cacheHelper?.Clear();
#pragma warning restore CS0618

        _logger.LogInformation("The token cache has been cleared at the user's request.");
        return Task.CompletedTask;
    }

    /// <summary>Releases the synchronization primitive.</summary>
    public void Dispose() => _initializationGate.Dispose();

    /// <summary>Describes the platform mechanism, for display on the Diagnostics page.</summary>
    public static string DescribeMechanism()
    {
        if (OperatingSystem.IsWindows())
        {
            return "Windows DPAPI-protected file";
        }

        if (OperatingSystem.IsMacOS())
        {
            return "macOS Keychain";
        }

        if (OperatingSystem.IsLinux())
        {
            return "Linux Secret Service (libsecret) keyring";
        }

        return "Unknown platform";
    }
}
