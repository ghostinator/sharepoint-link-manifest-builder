using System.Collections.ObjectModel;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using Microsoft.Extensions.Logging;
using SharePointLinkManifestBuilder.App.Infrastructure;
using SharePointLinkManifestBuilder.Core.Abstractions;
using SharePointLinkManifestBuilder.Core.Models;
using SharePointLinkManifestBuilder.Core.Urls;

namespace SharePointLinkManifestBuilder.App.ViewModels;

/// <summary>
/// The graphical SharePoint selector: search or paste a URL, then expand
/// site -&gt; library -&gt; folder -&gt; subfolder and add any level as a processing target.
/// <para>
/// Every level loads lazily. Nothing enumerates a whole tenant, or even a whole library, merely
/// to draw the screen.
/// </para>
/// </summary>
public sealed partial class SharePointBrowserViewModel : PageViewModelBase
{
    private readonly ISiteService _siteService;
    private readonly IDriveService _driveService;
    private readonly ConnectionCoordinator _connection;
    private readonly JobDraft _draft;
    private readonly ISystemBrowser _browser;
    private readonly IClipboardService _clipboard;
    private readonly ISettingsStore _settingsStore;
    private readonly ILogger<SharePointBrowserViewModel> _logger;

    /// <summary>Free-text search query.</summary>
    [ObservableProperty]
    private string _searchQuery = string.Empty;

    /// <summary>A SharePoint or OneDrive URL the user pasted.</summary>
    [ObservableProperty]
    private string _pastedUrl = string.Empty;

    /// <summary>The node whose details are shown in the side panel.</summary>
    [ObservableProperty]
    private ResourceNodeViewModel? _selectedNode;

    /// <summary>True to include subfolders when a selected node is added as a target.</summary>
    [ObservableProperty]
    private bool _includeSubfolders = true;

    /// <summary>True to reveal Graph identifiers in the details panel.</summary>
    [ObservableProperty]
    private bool _showAdvancedDetails;

    /// <summary>Creates the page.</summary>
    public SharePointBrowserViewModel(
        ISiteService siteService,
        IDriveService driveService,
        ConnectionCoordinator connection,
        JobDraft draft,
        ISystemBrowser browser,
        IClipboardService clipboard,
        ISettingsStore settingsStore,
        ILogger<SharePointBrowserViewModel> logger)
        : base("SharePoint Sites", "sharepoint")
    {
        _siteService = siteService ?? throw new ArgumentNullException(nameof(siteService));
        _driveService = driveService ?? throw new ArgumentNullException(nameof(driveService));
        _connection = connection ?? throw new ArgumentNullException(nameof(connection));
        _draft = draft ?? throw new ArgumentNullException(nameof(draft));
        _browser = browser ?? throw new ArgumentNullException(nameof(browser));
        _clipboard = clipboard ?? throw new ArgumentNullException(nameof(clipboard));
        _settingsStore = settingsStore ?? throw new ArgumentNullException(nameof(settingsStore));
        _logger = logger ?? throw new ArgumentNullException(nameof(logger));
    }

    /// <summary>Root nodes of the tree.</summary>
    public ObservableCollection<ResourceNodeViewModel> Roots { get; } = [];

    /// <summary>Locations the user pinned locally.</summary>
    public ObservableCollection<RecentLocation> PinnedLocations { get; } = [];

    /// <summary>Locations used recently.</summary>
    public ObservableCollection<RecentLocation> RecentLocations { get; } = [];

    /// <summary>Number of targets currently in the draft job.</summary>
    public int TargetCount => _draft.Targets.Count;

    /// <summary>
    /// The honesty notice shown above the results. Site search reflects the search index for
    /// the signed-in user, so claiming it lists every site in the tenant would be false.
    /// </summary>
    public static string SearchScopeNotice =>
        "Search shows sites Microsoft 365 returns for your account. It is not necessarily every site in "
        + "the organization. If a site is missing, paste its URL below to open it directly.";

    /// <inheritdoc />
    public override async Task OnNavigatedToAsync(CancellationToken cancellationToken = default)
    {
        var settings = await _settingsStore.LoadAsync(cancellationToken).ConfigureAwait(true);

        PinnedLocations.Clear();
        RecentLocations.Clear();

        var tenantId = _connection.Tenant?.TenantId;

        // Locations are scoped to the tenant they came from, so switching tenant never shows a
        // previous tenant's sites.
        foreach (var pinned in settings.PinnedLocations.Where(l =>
            l.SourceType is TargetSourceType.SharePointSite or TargetSourceType.DocumentLibrary
                or TargetSourceType.SharePointFolder
            && string.Equals(l.TenantId, tenantId, StringComparison.OrdinalIgnoreCase)))
        {
            PinnedLocations.Add(pinned);
        }

        foreach (var recent in settings.RecentLocations
            .Where(l => string.Equals(l.TenantId, tenantId, StringComparison.OrdinalIgnoreCase))
            .Take(10))
        {
            RecentLocations.Add(recent);
        }

        if (Roots.Count == 0)
        {
            await LoadFollowedSitesAsync(cancellationToken).ConfigureAwait(true);
        }
    }

    /// <summary>Searches for sites matching the query.</summary>
    [RelayCommand]
    private async Task SearchAsync(CancellationToken cancellationToken)
    {
        if (string.IsNullOrWhiteSpace(SearchQuery))
        {
            await LoadFollowedSitesAsync(cancellationToken).ConfigureAwait(true);
            return;
        }

        IsBusy = true;
        ClearMessages();

        try
        {
            var result = await _siteService.SearchSitesAsync(SearchQuery, cancellationToken)
                .ConfigureAwait(true);

            if (!result.Succeeded)
            {
                ErrorMessage = result.Error!.Message;
                return;
            }

            Roots.Clear();

            foreach (var site in result.Value!)
            {
                Roots.Add(CreateSiteNode(site));
            }

            StatusMessage = Roots.Count == 0
                ? "No sites matched that search. Try a different term, or paste the site URL."
                : $"{Roots.Count} site(s) found.";
        }
        finally
        {
            IsBusy = false;
        }
    }

    /// <summary>Resolves a pasted SharePoint URL and adds it to the tree.</summary>
    [RelayCommand]
    private async Task OpenPastedUrlAsync(CancellationToken cancellationToken)
    {
        if (string.IsNullOrWhiteSpace(PastedUrl))
        {
            return;
        }

        IsBusy = true;
        ClearMessages();

        try
        {
            var parsed = SharePointUrlParser.Parse(PastedUrl);

            if (!parsed.IsValid)
            {
                ErrorMessage = parsed.FailureReason;
                return;
            }

            var site = await _siteService.ResolveSiteByUrlAsync(PastedUrl, cancellationToken)
                .ConfigureAwait(true);

            if (!site.Succeeded)
            {
                ErrorMessage = site.Error!.Message;
                return;
            }

            var node = CreateSiteNode(site.Value!);
            Roots.Insert(0, node);
            node.IsExpanded = true;

            StatusMessage = parsed.SiteRelativeItemPath is { Length: > 0 } path
                ? $"Opened '{site.Value!.DisplayName}'. Expand to '{path}' to select that folder."
                : $"Opened '{site.Value!.DisplayName}'.";

            PastedUrl = string.Empty;
        }
        finally
        {
            IsBusy = false;
        }
    }

    /// <summary>Reloads the initial site list.</summary>
    [RelayCommand]
    private async Task RefreshAsync(CancellationToken cancellationToken)
    {
        Roots.Clear();
        await LoadFollowedSitesAsync(cancellationToken).ConfigureAwait(true);
    }

    /// <summary>Retries loading a node whose children failed to load.</summary>
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

    /// <summary>Adds every checked node in the tree as a processing target.</summary>
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
            StatusMessage = "Select a site, library or folder first.";
            return;
        }

        var added = 0;

        foreach (var node in selected)
        {
            if (_draft.AddTarget(BuildTarget(node, tenantId)))
            {
                added++;
            }
        }

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

    /// <summary>Clears every checkbox in the tree.</summary>
    [RelayCommand]
    private void ClearSelection()
    {
        foreach (var node in Roots.SelectMany(r => r.SelfAndDescendants()))
        {
            node.SetCheckedFromParent(false);
        }

        StatusMessage = "Selection cleared.";
    }

    /// <summary>Checks every currently visible root node.</summary>
    [RelayCommand]
    private void SelectAllVisible()
    {
        foreach (var root in Roots)
        {
            root.IsChecked = true;
        }
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

    /// <summary>Opens a pinned or recent location in the tree.</summary>
    [RelayCommand]
    private async Task OpenLocationAsync(RecentLocation? location)
    {
        if (location?.WebUrl is null)
        {
            return;
        }

        PastedUrl = location.WebUrl;
        await OpenPastedUrlAsync(CancellationToken.None).ConfigureAwait(true);
    }

    private async Task LoadFollowedSitesAsync(CancellationToken cancellationToken)
    {
        if (_connection.State is ConnectionState.NotConfigured or ConnectionState.ConfiguredSignedOut)
        {
            StatusMessage = "Connect to Microsoft 365 to browse SharePoint.";
            return;
        }

        IsBusy = true;
        ClearMessages();

        try
        {
            var followed = await _siteService.GetFollowedSitesAsync(cancellationToken).ConfigureAwait(true);

            if (followed.Succeeded && followed.Value!.Count > 0)
            {
                foreach (var site in followed.Value!)
                {
                    Roots.Add(CreateSiteNode(site));
                }

                StatusMessage = $"Showing {Roots.Count} site(s) you follow. Search to find others.";
                return;
            }

            var root = await _siteService.GetRootSiteAsync(cancellationToken).ConfigureAwait(true);

            if (root.Succeeded)
            {
                Roots.Add(CreateSiteNode(root.Value!));
                StatusMessage = "Showing the organization's root site. Search or paste a URL to find others.";
            }
            else
            {
                ErrorMessage = root.Error!.Message;
            }
        }
        finally
        {
            IsBusy = false;
        }
    }

    private ResourceNodeViewModel CreateSiteNode(SharePointSite site) =>
        new(ResourceKind.Site, site.DisplayName, LoadSiteDrivesAsync)
        {
            Description = site.WebUrl,
            WebUrl = site.WebUrl,
            SiteId = site.SiteId,
            SiteName = site.DisplayName,
            SourceType = TargetSourceType.SharePointSite,
        };

    private async Task<IReadOnlyList<ResourceNodeViewModel>> LoadSiteDrivesAsync(
        ResourceNodeViewModel node,
        CancellationToken cancellationToken)
    {
        var result = await _siteService.GetSiteDrivesAsync(node.SiteId!, cancellationToken)
            .ConfigureAwait(true);

        if (!result.Succeeded)
        {
            throw new InvalidOperationException(result.Error!.Message);
        }

        return result.Value!
            .Select(drive => new ResourceNodeViewModel(
                ResourceKind.Drive, drive.Name, LoadFolderChildrenAsync)
            {
                Description = drive.WebUrl,
                WebUrl = drive.WebUrl,
                SiteId = node.SiteId,
                SiteName = node.SiteName,
                DriveId = drive.DriveId,
                DriveName = drive.Name,
                SourceType = TargetSourceType.DocumentLibrary,
            })
            .ToArray();
    }

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
                SiteId = node.SiteId,
                SiteName = node.SiteName,
                DriveId = driveId,
                DriveName = node.DriveName,
                ItemId = folder.ItemId,
                RelativePath = CombinePath(node.RelativePath, folder.Name),
                SourceType = TargetSourceType.SharePointFolder,
            })
            .ToArray();
    }

    private ProcessingTarget BuildTarget(ResourceNodeViewModel node, string tenantId) => new()
    {
        SourceType = node.SourceType,
        TenantId = tenantId,
        SiteId = node.SiteId,
        SiteName = node.SiteName,
        SiteUrl = node.Kind == ResourceKind.Site ? node.WebUrl : null,
        DriveId = node.DriveId,
        DriveName = node.DriveName,
        StartingFolderName = node.Kind == ResourceKind.Folder ? node.DisplayName : null,
        StartingFolderItemId = node.Kind == ResourceKind.Folder ? node.ItemId : null,
        StartingFolderRelativePath = node.Kind == ResourceKind.Folder ? node.RelativePath : string.Empty,
        WebUrl = node.WebUrl,
        Recursive = IncludeSubfolders,
    };

    private static string CombinePath(string parent, string name) =>
        string.IsNullOrEmpty(parent) ? name : $"{parent}/{name}";
}
