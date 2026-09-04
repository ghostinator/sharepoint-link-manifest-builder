using System.Collections.ObjectModel;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using Microsoft.Extensions.Logging;
using SharePointLinkManifestBuilder.App.Infrastructure;
using SharePointLinkManifestBuilder.Core.Abstractions;
using SharePointLinkManifestBuilder.Core.Models;

namespace SharePointLinkManifestBuilder.App.ViewModels;

/// <summary>
/// The OneDrive selector, covering both the signed-in user's own OneDrive and another user's
/// OneDrive where access is permitted.
/// <para>
/// Finding a user in the picker says nothing about whether their OneDrive can be opened. The
/// drive may be unprovisioned, access-denied or blocked by policy, and this page reports the
/// actual Graph outcome rather than implying that administrator consent grants access to
/// everyone's files.
/// </para>
/// </summary>
public sealed partial class OneDriveBrowserViewModel : PageViewModelBase
{
    private readonly IDriveService _driveService;
    private readonly IUserDirectoryService _userDirectory;
    private readonly ConnectionCoordinator _connection;
    private readonly JobDraft _draft;
    private readonly ISystemBrowser _browser;
    private readonly IClipboardService _clipboard;
    private readonly ILogger<OneDriveBrowserViewModel> _logger;

    /// <summary>The user-search query for the people picker.</summary>
    [ObservableProperty]
    private string _userSearchQuery = string.Empty;

    /// <summary>The user chosen in the people picker.</summary>
    [ObservableProperty]
    private OneDriveUser? _selectedUser;

    /// <summary>The node whose details are shown.</summary>
    [ObservableProperty]
    private ResourceNodeViewModel? _selectedNode;

    /// <summary>True to include subfolders when adding a target.</summary>
    [ObservableProperty]
    private bool _includeSubfolders = true;

    /// <summary>True to reveal Graph identifiers.</summary>
    [ObservableProperty]
    private bool _showAdvancedDetails;

    /// <summary>True while a user-OneDrive lookup is running.</summary>
    [ObservableProperty]
    private bool _isSearchingUsers;

    /// <summary>Explains why a selected user's OneDrive could not be opened.</summary>
    [ObservableProperty]
    private string? _userDriveUnavailableReason;

    /// <summary>Creates the page.</summary>
    public OneDriveBrowserViewModel(
        IDriveService driveService,
        IUserDirectoryService userDirectory,
        ConnectionCoordinator connection,
        JobDraft draft,
        ISystemBrowser browser,
        IClipboardService clipboard,
        ILogger<OneDriveBrowserViewModel> logger)
        : base("OneDrive", "onedrive")
    {
        _driveService = driveService ?? throw new ArgumentNullException(nameof(driveService));
        _userDirectory = userDirectory ?? throw new ArgumentNullException(nameof(userDirectory));
        _connection = connection ?? throw new ArgumentNullException(nameof(connection));
        _draft = draft ?? throw new ArgumentNullException(nameof(draft));
        _browser = browser ?? throw new ArgumentNullException(nameof(browser));
        _clipboard = clipboard ?? throw new ArgumentNullException(nameof(clipboard));
        _logger = logger ?? throw new ArgumentNullException(nameof(logger));
    }

    /// <summary>Root nodes: My OneDrive, and any user drives that were opened.</summary>
    public ObservableCollection<ResourceNodeViewModel> Roots { get; } = [];

    /// <summary>Users matching the current search.</summary>
    public ObservableCollection<OneDriveUser> UserResults { get; } = [];

    /// <summary>Number of targets currently in the draft job.</summary>
    public int TargetCount => _draft.Targets.Count;

    /// <summary>The caveat shown above the people picker.</summary>
    public static string UserDriveNotice =>
        "Finding a user here does not mean their OneDrive can be opened. Access still depends on your own "
        + "permissions and your organization's policy, and a user who has never opened OneDrive does not "
        + "have one yet. This application never creates a OneDrive for someone.";

    /// <inheritdoc />
    public override async Task OnNavigatedToAsync(CancellationToken cancellationToken = default)
    {
        if (Roots.Count == 0)
        {
            await LoadMyDriveAsync(cancellationToken).ConfigureAwait(true);
        }
    }

    /// <summary>Opens the signed-in user's own OneDrive.</summary>
    [RelayCommand]
    private async Task LoadMyDriveAsync(CancellationToken cancellationToken)
    {
        if (_connection.State is ConnectionState.NotConfigured or ConnectionState.ConfiguredSignedOut)
        {
            StatusMessage = "Connect to Microsoft 365 to browse OneDrive.";
            return;
        }

        IsBusy = true;
        ClearMessages();

        try
        {
            var drive = await _driveService.GetMyDriveAsync(cancellationToken).ConfigureAwait(true);

            if (!drive.Succeeded)
            {
                ErrorMessage = drive.Error!.Message;
                return;
            }

            var existing = FindRootByDriveId(drive.Value!.DriveId);

            if (existing is not null)
            {
                RevealOnly(existing);
                StatusMessage = "Your OneDrive is already open.";
                return;
            }

            var node = CreateDriveNode(drive.Value!, TargetSourceType.MyOneDrive, "My OneDrive");
            Roots.Insert(0, node);
            RevealOnly(node);

            StatusMessage = "Opened your OneDrive.";
        }
        finally
        {
            IsBusy = false;
        }
    }

    /// <summary>Searches the directory for a user whose OneDrive should be browsed.</summary>
    [RelayCommand]
    private async Task SearchUsersAsync(CancellationToken cancellationToken)
    {
        if (string.IsNullOrWhiteSpace(UserSearchQuery) || UserSearchQuery.Trim().Length < 2)
        {
            StatusMessage = "Type at least two characters to search for a user.";
            return;
        }

        IsSearchingUsers = true;
        ClearMessages();
        UserDriveUnavailableReason = null;

        try
        {
            var result = await _userDirectory.SearchUsersAsync(UserSearchQuery, cancellationToken)
                .ConfigureAwait(true);

            UserResults.Clear();

            if (!result.Succeeded)
            {
                // The most common cause is the optional people-picker scope not being granted,
                // which is a configuration state rather than a fault.
                ErrorMessage = result.Error!.Kind == GraphErrorKind.SharePointAccessDenied
                    ? "Searching for users needs the optional User.ReadBasic.All permission, which has not "
                      + "been granted. You can still open your own OneDrive."
                    : result.Error.Message;

                return;
            }

            foreach (var user in result.Value!)
            {
                UserResults.Add(user);
            }

            StatusMessage = UserResults.Count == 0
                ? "No users matched that search."
                : $"{UserResults.Count} user(s) found.";
        }
        finally
        {
            IsSearchingUsers = false;
        }
    }

    /// <summary>Opens the selected user's OneDrive, reporting honestly when it cannot be opened.</summary>
    [RelayCommand]
    private async Task OpenUserDriveAsync(CancellationToken cancellationToken)
    {
        if (SelectedUser is null)
        {
            StatusMessage = "Select a user first.";
            return;
        }

        IsBusy = true;
        ClearMessages();
        UserDriveUnavailableReason = null;

        try
        {
            var drive = await _driveService.GetUserDriveAsync(SelectedUser.UserId, cancellationToken)
                .ConfigureAwait(true);

            if (!drive.Succeeded)
            {
                // The normalized error already distinguishes unprovisioned from access-denied,
                // so the user is told which of the two it actually is.
                UserDriveUnavailableReason = drive.Error!.SuggestedAction is { Length: > 0 } action
                    ? $"{drive.Error.Message} {action}"
                    : drive.Error.Message;

                return;
            }

            // Identity is the drive, not the person who was searched for: two directory entries
            // can resolve to the same drive, and the same person can be selected twice. Opening
            // a second copy would give the user two subtrees whose checkboxes disagree about
            // the same files.
            var alreadyOpen = FindRootByDriveId(drive.Value!.DriveId);

            if (alreadyOpen is not null)
            {
                RevealOnly(alreadyOpen);

                StatusMessage = $"{SelectedUser.DisplayName}'s OneDrive is already open. "
                    + "Showing it rather than opening a second copy.";

                return;
            }

            var node = CreateDriveNode(
                drive.Value! with
                {
                    OwnerDisplayName = SelectedUser.DisplayName,
                    OwnerUserId = SelectedUser.UserId,
                },
                TargetSourceType.UserOneDrive,
                $"{SelectedUser.DisplayName}'s OneDrive");

            Roots.Add(node);
            RevealOnly(node);
            OnPropertyChanged(nameof(HasOtherUserDrives));
            CloseOtherUserDrivesCommand.NotifyCanExecuteChanged();

            StatusMessage = $"Opened {SelectedUser.DisplayName}'s OneDrive.";
        }
        finally
        {
            IsBusy = false;
        }
    }

    private ResourceNodeViewModel? FindRootByDriveId(string driveId) =>
        Roots.FirstOrDefault(r =>
            string.Equals(r.DriveId, driveId, StringComparison.OrdinalIgnoreCase));

    /// <summary>
    /// Expands one drive and collapses every other, so the drive just opened is the one on
    /// screen.
    /// <para>
    /// Several expanded drives at once is not a useful view: each is a full folder tree, and the
    /// one the user just asked for ends up below however much of the previous ones happens to be
    /// showing. Collapsing does not discard anything -- loaded children stay loaded, so
    /// re-expanding costs no further Graph calls, and any targets already added are untouched.
    /// </para>
    /// </summary>
    private void RevealOnly(ResourceNodeViewModel node)
    {
        foreach (var root in Roots)
        {
            root.IsExpanded = ReferenceEquals(root, node);
        }

        SelectedNode = node;
    }

    /// <summary>True when at least one other user's OneDrive is open.</summary>
    public bool HasOtherUserDrives =>
        Roots.Any(r => r.SourceType == TargetSourceType.UserOneDrive);

    /// <summary>Closes one open drive, leaving any targets already added from it alone.</summary>
    [RelayCommand]
    private void CloseDrive(ResourceNodeViewModel? node)
    {
        if (node is null || !Roots.Remove(node))
        {
            return;
        }

        OnPropertyChanged(nameof(HasOtherUserDrives));
        CloseDriveCommand.NotifyCanExecuteChanged();
        CloseOtherUserDrivesCommand.NotifyCanExecuteChanged();

        // Deliberately not removing targets added from this drive. They were an explicit choice
        // and are listed on the Targets step, where they can be removed individually; dropping
        // them here would undo that choice silently as a side effect of tidying the tree.
        StatusMessage = $"Closed {node.DisplayName}. Any targets already added from it are unchanged.";
    }

    /// <summary>Closes every other user's OneDrive, keeping your own.</summary>
    [RelayCommand(CanExecute = nameof(HasOtherUserDrives))]
    private void CloseOtherUserDrives()
    {
        var closed = Roots
            .Where(r => r.SourceType == TargetSourceType.UserOneDrive)
            .ToArray();

        foreach (var root in closed)
        {
            Roots.Remove(root);
        }

        OnPropertyChanged(nameof(HasOtherUserDrives));
        CloseOtherUserDrivesCommand.NotifyCanExecuteChanged();

        StatusMessage = closed.Length == 1
            ? "Closed 1 other user's OneDrive. Any targets already added are unchanged."
            : $"Closed {closed.Length} other users' OneDrives. Any targets already added are unchanged.";
    }

    /// <summary>Adds every checked node as a processing target.</summary>
    [RelayCommand]
    private void AddSelectedAsTargets()
    {
        ClearMessages();

        var tenantId = _connection.Tenant?.TenantId;

        if (tenantId is null)
        {
            ErrorMessage = "Connect to Microsoft 365 before adding targets.";
            return;
        }

        var selected = Roots
            .SelectMany(r => r.SelfAndDescendants())
            .Where(n => n.IsChecked == true && n.CanBeTarget)
            .ToArray();

        if (selected.Length == 0)
        {
            StatusMessage = "Select a OneDrive or folder first.";
            return;
        }

        var added = selected.Count(node => _draft.AddTarget(BuildTarget(node, tenantId)));

        OnPropertyChanged(nameof(TargetCount));

        StatusMessage = added == selected.Length
            ? $"Added {added} target(s) to the job."
            : $"Added {added} target(s). {selected.Length - added} were already in the job.";
    }

    /// <summary>Adds the node shown in the details panel as a target.</summary>
    [RelayCommand]
    private void AddNodeAsTarget(ResourceNodeViewModel? node)
    {
        var target = node ?? SelectedNode;
        var tenantId = _connection.Tenant?.TenantId;

        if (target is null || tenantId is null || !target.CanBeTarget)
        {
            return;
        }

        StatusMessage = _draft.AddTarget(BuildTarget(target, tenantId))
            ? $"Added '{target.DisplayName}' to the job."
            : $"'{target.DisplayName}' is already in the job.";

        OnPropertyChanged(nameof(TargetCount));
    }

    /// <summary>Clears every checkbox.</summary>
    [RelayCommand]
    private void ClearSelection()
    {
        foreach (var node in Roots.SelectMany(r => r.SelfAndDescendants()))
        {
            node.SetCheckedFromParent(false);
        }

        StatusMessage = "Selection cleared.";
    }

    /// <summary>Retries a node whose children failed to load.</summary>
    [RelayCommand]
    private static async Task RetryNodeAsync(ResourceNodeViewModel? node)
    {
        if (node is null)
        {
            return;
        }

        node.Reset();
        await node.EnsureChildrenLoadedAsync().ConfigureAwait(true);
    }

    /// <summary>Opens the selected node in the system browser.</summary>
    [RelayCommand]
    private async Task OpenInBrowserAsync(ResourceNodeViewModel? node)
    {
        var url = (node ?? SelectedNode)?.WebUrl;

        if (!string.IsNullOrWhiteSpace(url) && Uri.TryCreate(url, UriKind.Absolute, out var uri))
        {
            await _browser.OpenAsync(uri).ConfigureAwait(true);
        }
    }

    /// <summary>Copies the selected node's URL.</summary>
    [RelayCommand]
    private async Task CopyWebUrlAsync(ResourceNodeViewModel? node)
    {
        var url = (node ?? SelectedNode)?.WebUrl;

        if (!string.IsNullOrWhiteSpace(url))
        {
            await _clipboard.SetTextAsync(url).ConfigureAwait(true);
            StatusMessage = "URL copied.";
        }
    }

    private ResourceNodeViewModel CreateDriveNode(
        DriveResource drive,
        TargetSourceType sourceType,
        string displayName) =>
        new(ResourceKind.Drive, displayName, LoadFolderChildrenAsync)
        {
            Description = drive.WebUrl,
            WebUrl = drive.WebUrl,
            DriveId = drive.DriveId,
            DriveName = drive.Name,
            OwnerUserId = drive.OwnerUserId,
            OwnerDisplayName = drive.OwnerDisplayName,
            SourceType = sourceType,
        };

    private async Task<IReadOnlyList<ResourceNodeViewModel>> LoadFolderChildrenAsync(
        ResourceNodeViewModel node,
        CancellationToken cancellationToken)
    {
        var driveId = node.DriveId
            ?? throw new InvalidOperationException("This node has no drive to enumerate.");

        var folderItemId = node.ItemId;

        if (folderItemId is null)
        {
            var root = await _driveService.GetRootFolderAsync(driveId, cancellationToken)
                .ConfigureAwait(true);

            if (!root.Succeeded)
            {
                throw new InvalidOperationException(root.Error!.Message);
            }

            folderItemId = root.Value!.ItemId;
        }

        var folders = await _driveService.GetSubfoldersAsync(driveId, folderItemId, cancellationToken)
            .ConfigureAwait(true);

        if (!folders.Succeeded)
        {
            throw new InvalidOperationException(folders.Error!.Message);
        }

        return folders.Value!
            .Select(folder => new ResourceNodeViewModel(
                ResourceKind.Folder, folder.Name, LoadFolderChildrenAsync)
            {
                Description = folder.ChildCount is { } count ? $"{count} item(s)" : null,
                WebUrl = folder.WebUrl,
                DriveId = driveId,
                DriveName = node.DriveName,
                ItemId = folder.ItemId,
                OwnerUserId = node.OwnerUserId,
                OwnerDisplayName = node.OwnerDisplayName,
                RelativePath = string.IsNullOrEmpty(node.RelativePath)
                    ? folder.Name
                    : $"{node.RelativePath}/{folder.Name}",
                SourceType = node.SourceType,
            })
            .ToArray();
    }

    private ProcessingTarget BuildTarget(ResourceNodeViewModel node, string tenantId) => new()
    {
        SourceType = node.SourceType,
        TenantId = tenantId,
        UserId = node.OwnerUserId,
        UserDisplayName = node.OwnerDisplayName,
        DriveId = node.DriveId,
        DriveName = node.DriveName,
        StartingFolderName = node.Kind == ResourceKind.Folder ? node.DisplayName : null,
        StartingFolderItemId = node.Kind == ResourceKind.Folder ? node.ItemId : null,
        StartingFolderRelativePath = node.Kind == ResourceKind.Folder ? node.RelativePath : string.Empty,
        WebUrl = node.WebUrl,
        Recursive = IncludeSubfolders,
    };
}
