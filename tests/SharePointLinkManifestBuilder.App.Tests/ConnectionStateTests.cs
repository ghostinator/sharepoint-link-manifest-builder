using Microsoft.Extensions.DependencyInjection;
using SharePointLinkManifestBuilder.App.Composition;
using SharePointLinkManifestBuilder.App.Infrastructure;
using SharePointLinkManifestBuilder.Core.Abstractions;
using SharePointLinkManifestBuilder.Core.Models;
using SharePointLinkManifestBuilder.Core.Settings;

namespace SharePointLinkManifestBuilder.App.Tests;

/// <summary>
/// A successful sign-in has to reach the rest of the application. The wizard used to call the
/// authentication service directly, which left the coordinator -- and therefore every page that
/// reads it -- believing nothing had happened: the wizard showed "Setup complete, consent
/// granted" while the header still said "Not connected".
/// </summary>
public sealed class ConnectionStateTests : IDisposable
{
    private readonly string _stateDirectory;
    private readonly ServiceProvider _services;
    private readonly StubAuthenticationService _auth = new();

    /// <summary>Builds the real container with only authentication substituted.</summary>
    public ConnectionStateTests()
    {
        _stateDirectory = Path.Combine(
            Path.GetTempPath(), "splmb-tests", Guid.NewGuid().ToString("n"));

        _services = ServiceRegistration.Build(
            new ApplicationPaths(_stateDirectory),
            services => services.AddSingleton<IAuthenticationService>(_auth));
    }

    /// <inheritdoc />
    public void Dispose()
    {
        _services.Dispose();

        if (Directory.Exists(_stateDirectory))
        {
            Directory.Delete(_stateDirectory, recursive: true);
        }
    }

    private static TenantConfiguration Tenant => new()
    {
        TenantId = "11111111-1111-1111-1111-111111111111",
        ClientId = "22222222-2222-2222-2222-222222222222",
        RequiredScopes = ["User.Read", "Sites.Read.All", "Files.ReadWrite.All"],
    };

    /// <summary>Signing in with the tenant's own scopes must reach Connected.</summary>
    [Fact]
    public async Task SignIn_WithTenantScopes_ReachesConnected()
    {
        var connection = _services.GetRequiredService<ConnectionCoordinator>();
        await connection.ApplyTenantAsync(Tenant);

        Assert.NotEqual(ConnectionState.Connected, connection.State);

        var result = await connection.SignInAsync();

        Assert.True(result.Succeeded);
        Assert.Equal(ConnectionState.Connected, connection.State);
        Assert.Equal(ConsentState.Granted, connection.Tenant!.ConsentState);
    }

    /// <summary>
    /// The scoped overload the wizard uses must behave identically when the scopes happen to be
    /// the tenant's own, so routing through the coordinator costs the wizard nothing.
    /// </summary>
    [Fact]
    public async Task SignIn_WithExplicitTenantScopes_AlsoReachesConnected()
    {
        var connection = _services.GetRequiredService<ConnectionCoordinator>();
        await connection.ApplyTenantAsync(Tenant);

        await connection.SignInAsync(Tenant.RequiredScopes);

        Assert.Equal(ConnectionState.Connected, connection.State);
    }

    /// <summary>
    /// A bootstrap sign-in asks for a deliberately different, smaller scope set. It must not be
    /// judged against the operating scopes, or automatic setup would record a healthy tenant as
    /// only partially consented purely because of a setup-time sign-in.
    /// </summary>
    [Fact]
    public async Task SignIn_WithBootstrapScopes_DoesNotDowngradeConsent()
    {
        var connection = _services.GetRequiredService<ConnectionCoordinator>();
        await connection.ApplyTenantAsync(Tenant with { ConsentState = ConsentState.Granted });

        _auth.GrantedScopes = ["User.Read", "AppRegistration.Create"];
        await connection.SignInAsync(["User.Read", "AppRegistration.Create"]);

        Assert.NotEqual(ConsentState.PartiallyGranted, connection.Tenant!.ConsentState);
    }

    /// <summary>A sign-in that fails must not report the application as connected.</summary>
    [Fact]
    public async Task SignIn_WhenItFails_DoesNotReachConnected()
    {
        var connection = _services.GetRequiredService<ConnectionCoordinator>();
        await connection.ApplyTenantAsync(Tenant);

        _auth.ShouldSucceed = false;
        var result = await connection.SignInAsync();

        Assert.False(result.Succeeded);
        Assert.NotEqual(ConnectionState.Connected, connection.State);
    }

    /// <summary>Signing in before a tenant is configured is refused, not attempted.</summary>
    [Fact]
    public async Task SignIn_WithoutATenant_IsRefused()
    {
        var connection = _services.GetRequiredService<ConnectionCoordinator>();

        var result = await connection.SignInAsync();

        Assert.False(result.Succeeded);
        Assert.Equal(0, _auth.SignInCalls);
    }

    /// <summary>Authentication stand-in: records calls and reports whatever scopes it is told to.</summary>
    private sealed class StubAuthenticationService : IAuthenticationService
    {
        public bool ShouldSucceed { get; set; } = true;

        public int SignInCalls { get; private set; }

        public IReadOnlyList<string> GrantedScopes { get; set; } =
            ["User.Read", "Sites.Read.All", "Files.ReadWrite.All"];

        public UserAccount? CurrentAccount { get; private set; }

        public event EventHandler<UserAccount?>? AccountChanged;

        public Task ConfigureAsync(TenantConfiguration configuration, CancellationToken ct = default) =>
            Task.CompletedTask;

        public Task<AuthenticationResultInfo> SignInAsync(
            IEnumerable<string> scopes, string? loginHint = null, CancellationToken ct = default)
        {
            SignInCalls++;
            return Task.FromResult(Result(scopes));
        }

        public Task<AuthenticationResultInfo> AcquireTokenAsync(
            IEnumerable<string> scopes, bool allowInteractive = false, CancellationToken ct = default) =>
            Task.FromResult(Result(scopes));

        public Task<string?> GetAccessTokenAsync(IEnumerable<string> scopes, CancellationToken ct = default) =>
            Task.FromResult<string?>(ShouldSucceed ? "STUB-TOKEN-NOT-A-CREDENTIAL" : null);

        public Task<IReadOnlyList<UserAccount>> GetCachedAccountsAsync(CancellationToken ct = default) =>
            Task.FromResult<IReadOnlyList<UserAccount>>([]);

        public Task<AuthenticationResultInfo> SwitchToAccountAsync(
            string homeAccountId, IEnumerable<string> scopes, CancellationToken ct = default) =>
            Task.FromResult(Result(scopes));

        public Task SignOutAsync(CancellationToken ct = default)
        {
            CurrentAccount = null;
            AccountChanged?.Invoke(this, null);
            return Task.CompletedTask;
        }

        public Task ForgetAccountAsync(string homeAccountId, CancellationToken ct = default) =>
            Task.CompletedTask;

        private AuthenticationResultInfo Result(IEnumerable<string> scopes)
        {
            if (!ShouldSucceed)
            {
                return new AuthenticationResultInfo
                {
                    Succeeded = false,
                    Error = new GraphError
                    {
                        Kind = GraphErrorKind.AuthenticationFailed,
                        Message = "Stubbed failure.",
                    },
                };
            }

            CurrentAccount = new UserAccount
            {
                UserId = "user-1",
                DisplayName = "Test User",
                UserPrincipalName = "test.user@example.test",
                TenantId = "11111111-1111-1111-1111-111111111111",
            };

            AccountChanged?.Invoke(this, CurrentAccount);

            return new AuthenticationResultInfo
            {
                Succeeded = true,
                Account = CurrentAccount,
                GrantedScopes = GrantedScopes,
                ExpiresOnUtc = DateTimeOffset.UtcNow.AddHours(1),
            };
        }
    }
}

/// <summary>
/// Automatic setup signs in twice with different identities: the bootstrap client to create the
/// registration, then the new registration itself. These pin that the second sign-in is what
/// leaves the application connected, which is the step that was missing.
/// </summary>
public sealed class AutomaticSetupConnectionTests : IDisposable
{
    private readonly string _stateDirectory;
    private readonly ServiceProvider _services;
    private readonly ScriptedAuthenticationService _auth = new();

    /// <summary>Builds the real container with only authentication substituted.</summary>
    public AutomaticSetupConnectionTests()
    {
        _stateDirectory = Path.Combine(
            Path.GetTempPath(), "splmb-tests", Guid.NewGuid().ToString("n"));

        _services = ServiceRegistration.Build(
            new ApplicationPaths(_stateDirectory),
            services => services.AddSingleton<IAuthenticationService>(_auth));
    }

    /// <inheritdoc />
    public void Dispose()
    {
        _services.Dispose();

        if (Directory.Exists(_stateDirectory))
        {
            Directory.Delete(_stateDirectory, recursive: true);
        }
    }

    private static TenantConfiguration Bootstrap => new()
    {
        TenantId = "11111111-1111-1111-1111-111111111111",
        ClientId = "bbbbbbbb-bbbb-bbbb-bbbb-bbbbbbbbbbbb",
        RequiredScopes = ["User.Read", "Sites.Read.All", "Files.ReadWrite.All"],
    };

    /// <summary>
    /// The bootstrap sign-in asks only for setup scopes, so it must not be mistaken for the
    /// application being ready to use.
    /// </summary>
    [Fact]
    public async Task BootstrapSignIn_AloneDoesNotReportConnected()
    {
        var connection = _services.GetRequiredService<ConnectionCoordinator>();
        await connection.ApplyTenantAsync(Bootstrap);

        _auth.GrantedScopes = ["User.Read", "AppRegistration.Create"];
        await connection.SignInAsync(["User.Read", "AppRegistration.Create"]);

        Assert.NotEqual(ConnectionState.Connected, connection.State);
    }

    /// <summary>
    /// Signing in to the new registration with the operating scopes is what finishes setup. Any
    /// flow that omits it leaves the user to discover the Home page's sign-in button.
    /// </summary>
    [Fact]
    public async Task SigningInToTheNewRegistration_CompletesTheConnection()
    {
        var connection = _services.GetRequiredService<ConnectionCoordinator>();
        await connection.ApplyTenantAsync(Bootstrap);

        _auth.GrantedScopes = ["User.Read", "AppRegistration.Create"];
        await connection.SignInAsync(["User.Read", "AppRegistration.Create"]);

        // The registration now exists, so the wizard reconfigures and signs in to it.
        await connection.SaveTenantAsync(Bootstrap with
        {
            ClientId = "cccccccc-cccc-cccc-cccc-cccccccccccc",
        });

        _auth.GrantedScopes = ["User.Read", "Sites.Read.All", "Files.ReadWrite.All"];
        await connection.SignInAsync(Bootstrap.RequiredScopes);

        Assert.Equal(ConnectionState.Connected, connection.State);
        Assert.Equal(ConsentState.Granted, connection.Tenant!.ConsentState);
    }

    /// <summary>
    /// Entra replicates a new registration before it can be used, so the first sign-in attempts
    /// can legitimately fail with "this registration does not exist". Retrying must recover
    /// rather than reporting a registration that was created correctly as broken.
    /// </summary>
    [Fact]
    public async Task SignIn_RecoversOnceReplicationCatchesUp()
    {
        var connection = _services.GetRequiredService<ConnectionCoordinator>();
        await connection.ApplyTenantAsync(Bootstrap);

        _auth.FailuresBeforeSuccess = 2;
        _auth.FailureKind = GraphErrorKind.ApplicationNotFoundInTenant;

        var first = await connection.SignInAsync();
        Assert.False(first.Succeeded);
        Assert.Equal(GraphErrorKind.ApplicationNotFoundInTenant, first.Error!.Kind);

        var second = await connection.SignInAsync();
        Assert.False(second.Succeeded);

        var third = await connection.SignInAsync();
        Assert.True(third.Succeeded);
        Assert.Equal(ConnectionState.Connected, connection.State);
    }

    /// <summary>Authentication stand-in that can fail a fixed number of times first.</summary>
    private sealed class ScriptedAuthenticationService : IAuthenticationService
    {
        public int FailuresBeforeSuccess { get; set; }

        public GraphErrorKind FailureKind { get; set; } = GraphErrorKind.AuthenticationFailed;

        public IReadOnlyList<string> GrantedScopes { get; set; } =
            ["User.Read", "Sites.Read.All", "Files.ReadWrite.All"];

        public UserAccount? CurrentAccount { get; private set; }

        public event EventHandler<UserAccount?>? AccountChanged;

        public Task ConfigureAsync(TenantConfiguration configuration, CancellationToken ct = default) =>
            Task.CompletedTask;

        public Task<AuthenticationResultInfo> SignInAsync(
            IEnumerable<string> scopes, string? loginHint = null, CancellationToken ct = default) =>
            Task.FromResult(Next());

        public Task<AuthenticationResultInfo> AcquireTokenAsync(
            IEnumerable<string> scopes, bool allowInteractive = false, CancellationToken ct = default) =>
            Task.FromResult(Next());

        public Task<string?> GetAccessTokenAsync(IEnumerable<string> scopes, CancellationToken ct = default) =>
            Task.FromResult<string?>("STUB-TOKEN-NOT-A-CREDENTIAL");

        public Task<IReadOnlyList<UserAccount>> GetCachedAccountsAsync(CancellationToken ct = default) =>
            Task.FromResult<IReadOnlyList<UserAccount>>([]);

        public Task<AuthenticationResultInfo> SwitchToAccountAsync(
            string homeAccountId, IEnumerable<string> scopes, CancellationToken ct = default) =>
            Task.FromResult(Next());

        public Task SignOutAsync(CancellationToken ct = default) => Task.CompletedTask;

        public Task ForgetAccountAsync(string homeAccountId, CancellationToken ct = default) =>
            Task.CompletedTask;

        private AuthenticationResultInfo Next()
        {
            if (FailuresBeforeSuccess > 0)
            {
                FailuresBeforeSuccess--;

                return new AuthenticationResultInfo
                {
                    Succeeded = false,
                    Error = new GraphError { Kind = FailureKind, Message = "Scripted failure." },
                };
            }

            CurrentAccount = new UserAccount
            {
                UserId = "user-1",
                DisplayName = "Test User",
                UserPrincipalName = "test.user@example.test",
                TenantId = "11111111-1111-1111-1111-111111111111",
            };

            AccountChanged?.Invoke(this, CurrentAccount);

            return new AuthenticationResultInfo
            {
                Succeeded = true,
                Account = CurrentAccount,
                GrantedScopes = GrantedScopes,
                ExpiresOnUtc = DateTimeOffset.UtcNow.AddHours(1),
            };
        }
    }
}
