using Microsoft.Extensions.DependencyInjection;
using SharePointLinkManifestBuilder.App.Composition;
using SharePointLinkManifestBuilder.App.Infrastructure;
using SharePointLinkManifestBuilder.App.ViewModels;
using SharePointLinkManifestBuilder.Core.Abstractions;
using SharePointLinkManifestBuilder.Core.Models;
using SharePointLinkManifestBuilder.Core.Settings;

namespace SharePointLinkManifestBuilder.App.Tests;

/// <summary>
/// Opening several people's OneDrives at once has to stay legible. Each drive is a full folder
/// tree, so leaving them all expanded buries the one just asked for, and opening the same drive
/// twice produces two subtrees whose checkboxes disagree about the same files.
/// </summary>
public sealed class OneDriveMultiUserTests : IDisposable
{
    private const string MyDriveId = "drive-mine";

    private readonly string _stateDirectory;
    private readonly ServiceProvider _services;
    private readonly StubDriveService _drives = new();

    /// <summary>Builds the real container with only the drive service substituted.</summary>
    public OneDriveMultiUserTests()
    {
        _stateDirectory = Path.Combine(
            Path.GetTempPath(), "splmb-tests", Guid.NewGuid().ToString("n"));

        _services = ServiceRegistration.Build(
            new ApplicationPaths(_stateDirectory),
            services =>
            {
                services.AddSingleton<IDriveService>(_drives);
                services.AddSingleton<IAuthenticationService>(new StubAuthenticationService());
            });
    }

    /// <summary>
    /// Brings the connection to Connected. The browser refuses to load a drive while
    /// disconnected, which is correct behaviour and would otherwise make every test here assert
    /// against an empty tree.
    /// </summary>
    private async Task ConnectAsync()
    {
        var connection = _services.GetRequiredService<ConnectionCoordinator>();

        await connection.ApplyTenantAsync(new TenantConfiguration
        {
            TenantId = "11111111-1111-1111-1111-111111111111",
            ClientId = "22222222-2222-2222-2222-222222222222",
            RequiredScopes = ["User.Read", "Sites.Read.All", "Files.ReadWrite.All"],
        });

        await connection.SignInAsync();
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

    private OneDriveBrowserViewModel Browser() =>
        _services.GetRequiredService<OneDriveBrowserViewModel>();

    private static OneDriveUser User(string id, string name) =>
        new() { UserId = id, DisplayName = name, UserPrincipalName = $"{id}@example.test" };

    private async Task<OneDriveBrowserViewModel> WithMyDriveOpenAsync()
    {
        await ConnectAsync();
        var browser = Browser();
        await browser.LoadMyDriveCommand.ExecuteAsync(null);
        return browser;
    }

    /// <summary>Opening another drive collapses the rest and expands the new one.</summary>
    [Fact]
    public async Task OpeningAnotherDrive_CollapsesTheOthersAndExpandsTheNewOne()
    {
        var browser = await WithMyDriveOpenAsync();
        var mine = Assert.Single(browser.Roots);
        Assert.True(mine.IsExpanded);

        browser.SelectedUser = User("alice", "Alice");
        await browser.OpenUserDriveCommand.ExecuteAsync(null);

        Assert.Equal(2, browser.Roots.Count);
        Assert.False(mine.IsExpanded);

        var opened = browser.Roots[1];
        Assert.True(opened.IsExpanded);
        Assert.Same(opened, browser.SelectedNode);
    }

    /// <summary>The same user's drive must not open twice.</summary>
    [Fact]
    public async Task OpeningTheSameUserTwice_DoesNotAddASecondCopy()
    {
        var browser = await WithMyDriveOpenAsync();
        browser.SelectedUser = User("alice", "Alice");

        await browser.OpenUserDriveCommand.ExecuteAsync(null);
        await browser.OpenUserDriveCommand.ExecuteAsync(null);

        Assert.Equal(2, browser.Roots.Count);
        Assert.Contains("already open", browser.StatusMessage, StringComparison.OrdinalIgnoreCase);
    }

    /// <summary>Re-opening reveals the existing drive rather than doing nothing visible.</summary>
    [Fact]
    public async Task ReopeningAnAlreadyOpenDrive_RevealsIt()
    {
        var browser = await WithMyDriveOpenAsync();
        browser.SelectedUser = User("alice", "Alice");
        await browser.OpenUserDriveCommand.ExecuteAsync(null);

        var alice = browser.Roots[1];
        await browser.LoadMyDriveCommand.ExecuteAsync(null);

        Assert.False(alice.IsExpanded);
        Assert.True(browser.Roots[0].IsExpanded);

        await browser.OpenUserDriveCommand.ExecuteAsync(null);

        Assert.True(alice.IsExpanded);
        Assert.Same(alice, browser.SelectedNode);
    }

    /// <summary>Identity is the drive, so two directory entries for one drive open once.</summary>
    [Fact]
    public async Task TwoUsersResolvingToTheSameDrive_OpenOnlyOnce()
    {
        await ConnectAsync();
        var browser = Browser();

        browser.SelectedUser = User("alice", "Alice");
        await browser.OpenUserDriveCommand.ExecuteAsync(null);

        _drives.UserDriveIdOverride = "drive-alice";
        browser.SelectedUser = User("alice-alias", "Alice (alias)");
        await browser.OpenUserDriveCommand.ExecuteAsync(null);

        Assert.Single(browser.Roots);
    }

    /// <summary>One drive can be closed without touching the others.</summary>
    [Fact]
    public async Task CloseDrive_RemovesOnlyThatDrive()
    {
        var browser = await WithMyDriveOpenAsync();
        browser.SelectedUser = User("alice", "Alice");
        await browser.OpenUserDriveCommand.ExecuteAsync(null);

        browser.CloseDriveCommand.Execute(browser.Roots[1]);

        var remaining = Assert.Single(browser.Roots);
        Assert.Equal(MyDriveId, remaining.DriveId);
        Assert.False(browser.HasOtherUserDrives);
    }

    /// <summary>Closing every other drive keeps your own.</summary>
    [Fact]
    public async Task CloseOtherUserDrives_KeepsYourOwn()
    {
        var browser = await WithMyDriveOpenAsync();

        foreach (var (id, name) in new[] { ("alice", "Alice"), ("bob", "Bob") })
        {
            _drives.UserDriveIdOverride = $"drive-{id}";
            browser.SelectedUser = User(id, name);
            await browser.OpenUserDriveCommand.ExecuteAsync(null);
        }

        Assert.Equal(3, browser.Roots.Count);
        Assert.True(browser.HasOtherUserDrives);

        browser.CloseOtherUserDrivesCommand.Execute(null);

        var remaining = Assert.Single(browser.Roots);
        Assert.Equal(MyDriveId, remaining.DriveId);
        Assert.False(browser.HasOtherUserDrives);
    }

    /// <summary>With nothing to close, the command stays unavailable.</summary>
    [Fact]
    public async Task CloseOtherUserDrives_IsUnavailableWithNoneOpen()
    {
        var browser = await WithMyDriveOpenAsync();

        Assert.False(browser.HasOtherUserDrives);
        Assert.False(browser.CloseOtherUserDrivesCommand.CanExecute(null));
    }

    /// <summary>Only a drive row offers Close; folders inside one do not.</summary>
    [Fact]
    public async Task OnlyDriveRootsAreClosable()
    {
        var browser = await WithMyDriveOpenAsync();

        Assert.True(Assert.Single(browser.Roots).IsRoot);
    }

    /// <summary>Authentication stand-in, so the coordinator can reach Connected.</summary>
    private sealed class StubAuthenticationService : IAuthenticationService
    {
        public UserAccount? CurrentAccount { get; private set; }

        public event EventHandler<UserAccount?>? AccountChanged;

        public Task ConfigureAsync(TenantConfiguration configuration, CancellationToken ct = default) =>
            Task.CompletedTask;

        public Task<AuthenticationResultInfo> SignInAsync(
            IEnumerable<string> scopes, string? loginHint = null, CancellationToken ct = default) =>
            Task.FromResult(Result());

        public Task<AuthenticationResultInfo> AcquireTokenAsync(
            IEnumerable<string> scopes, bool allowInteractive = false, CancellationToken ct = default) =>
            Task.FromResult(Result());

        public Task<string?> GetAccessTokenAsync(IEnumerable<string> scopes, CancellationToken ct = default) =>
            Task.FromResult<string?>("STUB-TOKEN-NOT-A-CREDENTIAL");

        public Task<IReadOnlyList<UserAccount>> GetCachedAccountsAsync(CancellationToken ct = default) =>
            Task.FromResult<IReadOnlyList<UserAccount>>([]);

        public Task<AuthenticationResultInfo> SwitchToAccountAsync(
            string homeAccountId, IEnumerable<string> scopes, CancellationToken ct = default) =>
            Task.FromResult(Result());

        public Task SignOutAsync(CancellationToken ct = default) => Task.CompletedTask;

        public Task ForgetAccountAsync(string homeAccountId, CancellationToken ct = default) =>
            Task.CompletedTask;

        private AuthenticationResultInfo Result()
        {
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
                GrantedScopes = ["User.Read", "Sites.Read.All", "Files.ReadWrite.All"],
                ExpiresOnUtc = DateTimeOffset.UtcNow.AddHours(1),
            };
        }
    }

    /// <summary>Drive service stand-in returning distinct drives per user.</summary>
    private sealed class StubDriveService : IDriveService
    {
        public string UserDriveIdOverride { get; set; } = "drive-alice";

        public Task<OperationResult<DriveResource>> GetMyDriveAsync(CancellationToken ct = default) =>
            Task.FromResult(OperationResult<DriveResource>.Success(new DriveResource
            {
                DriveId = MyDriveId,
                Name = "OneDrive",
                DriveType = "business",
            }));

        public Task<OperationResult<DriveResource>> GetUserDriveAsync(
            string userId, CancellationToken ct = default) =>
            Task.FromResult(OperationResult<DriveResource>.Success(new DriveResource
            {
                DriveId = UserDriveIdOverride,
                Name = "OneDrive",
                DriveType = "business",
            }));

        public Task<OperationResult<DriveResource>> GetDriveAsync(
            string driveId, CancellationToken ct = default) =>
            GetMyDriveAsync(ct);

        public Task<OperationResult<SharePointFolder>> GetRootFolderAsync(
            string driveId, CancellationToken ct = default) =>
            Task.FromResult(OperationResult<SharePointFolder>.Success(new SharePointFolder
            {
                DriveId = driveId,
                ItemId = "root",
                Name = "root",
            }));

        public Task<OperationResult<SharePointFolder>> GetFolderByPathAsync(
            string driveId, string relativePath, CancellationToken ct = default) =>
            GetRootFolderAsync(driveId, ct);

        public Task<OperationResult<DiscoveredFile>> GetItemAsync(
            string driveId, string itemId, CancellationToken ct = default) =>
            throw new NotSupportedException();

        public async IAsyncEnumerable<DiscoveredFile> GetChildrenAsync(
            string driveId,
            string folderItemId,
            [System.Runtime.CompilerServices.EnumeratorCancellation] CancellationToken ct = default)
        {
            await Task.CompletedTask.ConfigureAwait(false);
            yield break;
        }

        public Task<OperationResult<IReadOnlyList<SharePointFolder>>> GetSubfoldersAsync(
            string driveId, string folderItemId, CancellationToken ct = default) =>
            Task.FromResult(OperationResult<IReadOnlyList<SharePointFolder>>.Success([]));

        public Task<OperationResult<DiscoveredFile>> ResolveSharingUrlAsync(
            string sharingUrl, CancellationToken ct = default) =>
            throw new NotSupportedException();
    }
}
