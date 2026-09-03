using System.Collections.ObjectModel;
using CommunityToolkit.Mvvm.ComponentModel;
using SharePointLinkManifestBuilder.Core.Models;

namespace SharePointLinkManifestBuilder.App.ViewModels;

/// <summary>
/// One node in the unified resource tree: a site, a library or OneDrive, a folder, or a file.
/// <para>
/// Children load lazily on first expansion. A tenant can hold millions of items, so the tree
/// fetches a folder's contents only when a user actually opens it. A placeholder child makes
/// the expander arrow appear before anything has been fetched.
/// </para>
/// </summary>
public sealed partial class ResourceNodeViewModel : ObservableObject
{
    private readonly Func<ResourceNodeViewModel, CancellationToken, Task<IReadOnlyList<ResourceNodeViewModel>>>?
        _loadChildren;

    private readonly Action<ResourceNodeViewModel>? _onSelectionChanged;
    private bool _isUpdatingFromParent;

    /// <summary>True while this node's children are being fetched.</summary>
    [ObservableProperty]
    private bool _isLoading;

    /// <summary>True when the node is expanded in the tree.</summary>
    [ObservableProperty]
    private bool _isExpanded;

    /// <summary>
    /// Tri-state selection: true, false, or null for partially selected. Null is what lets a
    /// parent honestly show that some but not all of its children are selected.
    /// </summary>
    [ObservableProperty]
    private bool? _isChecked = false;

    /// <summary>An error from the last attempt to load this node's children.</summary>
    [ObservableProperty]
    private string? _loadError;

    /// <summary>True when loading failed and a retry is worth offering.</summary>
    [ObservableProperty]
    private bool _canRetry;

    /// <summary>Creates a node.</summary>
    /// <param name="kind">What the node represents.</param>
    /// <param name="displayName">The friendly name shown to the user.</param>
    /// <param name="loadChildren">Callback that fetches children, or null for a leaf.</param>
    /// <param name="onSelectionChanged">Notified when this node's selection changes.</param>
    public ResourceNodeViewModel(
        ResourceKind kind,
        string displayName,
        Func<ResourceNodeViewModel, CancellationToken, Task<IReadOnlyList<ResourceNodeViewModel>>>? loadChildren = null,
        Action<ResourceNodeViewModel>? onSelectionChanged = null)
    {
        Kind = kind;
        DisplayName = displayName;
        _loadChildren = loadChildren;
        _onSelectionChanged = onSelectionChanged;

        if (loadChildren is not null)
        {
            // A placeholder gives the node an expander arrow before its real children exist.
            Children.Add(CreatePlaceholder());
        }
    }

    /// <summary>What this node represents.</summary>
    public ResourceKind Kind { get; }

    /// <summary>The friendly name. Always shown in preference to any identifier.</summary>
    public string DisplayName { get; }

    /// <summary>A secondary line, such as a URL or item count.</summary>
    public string? Description { get; init; }

    /// <summary>Absolute URL, for "Open in browser" and "Copy web URL".</summary>
    public string? WebUrl { get; init; }

    /// <summary>Graph site ID, when applicable. Shown only under Advanced details.</summary>
    public string? SiteId { get; init; }

    /// <summary>Graph drive ID, when applicable. Shown only under Advanced details.</summary>
    public string? DriveId { get; init; }

    /// <summary>Graph item ID, when applicable. Shown only under Advanced details.</summary>
    public string? ItemId { get; init; }

    /// <summary>Path relative to the drive root.</summary>
    public string RelativePath { get; init; } = string.Empty;

    /// <summary>Owning user ID, for a User OneDrive node.</summary>
    public string? OwnerUserId { get; init; }

    /// <summary>Owning user display name, for a User OneDrive node.</summary>
    public string? OwnerDisplayName { get; init; }

    /// <summary>Owning site display name.</summary>
    public string? SiteName { get; init; }

    /// <summary>Owning drive display name.</summary>
    public string? DriveName { get; init; }

    /// <summary>The source family this node belongs to, used when building a target.</summary>
    public TargetSourceType SourceType { get; init; } = TargetSourceType.DocumentLibrary;

    /// <summary>The parent node, or null at the root.</summary>
    public ResourceNodeViewModel? Parent { get; private set; }

    /// <summary>Child nodes, possibly holding a single placeholder until first expansion.</summary>
    public ObservableCollection<ResourceNodeViewModel> Children { get; } = [];

    /// <summary>True when this node is only a loading placeholder.</summary>
    public bool IsPlaceholder { get; private init; }

    /// <summary>True when this node can become a processing target.</summary>
    public bool CanBeTarget =>
        Kind is ResourceKind.Site or ResourceKind.Drive or ResourceKind.Folder;

    /// <summary>True when this node has real, loaded children.</summary>
    public bool HasLoadedChildren => Children.Count > 0 && !Children[0].IsPlaceholder;

    /// <summary>
    /// A screen-reader label combining the kind and name, so a tree item is not announced as a
    /// bare string with no indication of what it is.
    /// </summary>
    public string AccessibleName => Kind switch
    {
        ResourceKind.Site => $"SharePoint site {DisplayName}",
        ResourceKind.Drive => $"Document library {DisplayName}",
        ResourceKind.Folder => $"Folder {DisplayName}",
        ResourceKind.File => $"File {DisplayName}",
        ResourceKind.User => $"User {DisplayName}",
        _ => DisplayName,
    };

    /// <summary>An icon glyph for the node kind. Never the sole indicator of state.</summary>
    public string Glyph => Kind switch
    {
        ResourceKind.Site => "\U0001F310",
        ResourceKind.Drive => "\U0001F4DA",
        ResourceKind.Folder => "\U0001F4C1",
        ResourceKind.File => "\U0001F4C4",
        ResourceKind.User => "\U0001F464",
        _ => "\U0001F4C2",
    };

    /// <summary>Loads children if they have not been loaded yet.</summary>
    public async Task EnsureChildrenLoadedAsync(CancellationToken cancellationToken = default)
    {
        if (_loadChildren is null || HasLoadedChildren || IsLoading)
        {
            return;
        }

        IsLoading = true;
        LoadError = null;
        CanRetry = false;

        try
        {
            var children = await _loadChildren(this, cancellationToken).ConfigureAwait(true);

            Children.Clear();

            foreach (var child in children)
            {
                child.Parent = this;
                Children.Add(child);
            }

            // A newly loaded subtree inherits an already-selected parent's state, so expanding
            // after selecting does not silently deselect anything.
            if (IsChecked == true)
            {
                foreach (var child in Children)
                {
                    child.SetCheckedFromParent(true);
                }
            }
        }
        catch (OperationCanceledException)
        {
            throw;
        }
#pragma warning disable CA1031 // One unreadable folder must not break the whole tree.
        catch (Exception ex)
        {
            Children.Clear();
            LoadError = ex.Message;
            CanRetry = true;
        }
#pragma warning restore CA1031
        finally
        {
            IsLoading = false;
        }
    }

    /// <summary>Discards loaded children so the next expansion refetches them.</summary>
    public void Reset()
    {
        Children.Clear();

        if (_loadChildren is not null)
        {
            Children.Add(CreatePlaceholder());
        }

        LoadError = null;
        CanRetry = false;
    }

    /// <summary>Applies a checked state pushed down from a parent, without recursing upward.</summary>
    public void SetCheckedFromParent(bool value)
    {
        _isUpdatingFromParent = true;

        try
        {
            IsChecked = value;

            foreach (var child in Children.Where(c => !c.IsPlaceholder))
            {
                child.SetCheckedFromParent(value);
            }
        }
        finally
        {
            _isUpdatingFromParent = false;
        }
    }

    /// <summary>Enumerates this node and every loaded descendant.</summary>
    public IEnumerable<ResourceNodeViewModel> SelfAndDescendants()
    {
        yield return this;

        foreach (var child in Children.Where(c => !c.IsPlaceholder))
        {
            foreach (var descendant in child.SelfAndDescendants())
            {
                yield return descendant;
            }
        }
    }

    partial void OnIsExpandedChanged(bool value)
    {
        if (value)
        {
            // Fire and forget is acceptable here: EnsureChildrenLoadedAsync captures its own
            // failures onto LoadError rather than throwing into the void.
            _ = EnsureChildrenLoadedAsync();
        }
    }

    partial void OnIsCheckedChanged(bool? value)
    {
        if (_isUpdatingFromParent)
        {
            return;
        }

        // A user click resolves the indeterminate state to a definite one, then cascades.
        if (value is { } definite)
        {
            foreach (var child in Children.Where(c => !c.IsPlaceholder))
            {
                child.SetCheckedFromParent(definite);
            }
        }

        Parent?.RecalculateFromChildren();
        _onSelectionChanged?.Invoke(this);
    }

    /// <summary>
    /// Recomputes this node's tri-state from its children and propagates upward, which is what
    /// produces the partial-selection indicator on ancestors.
    /// </summary>
    private void RecalculateFromChildren()
    {
        var loaded = Children.Where(c => !c.IsPlaceholder).ToArray();

        if (loaded.Length == 0)
        {
            return;
        }

        var all = loaded.All(c => c.IsChecked == true);
        var none = loaded.All(c => c.IsChecked == false);

        _isUpdatingFromParent = true;

        try
        {
            IsChecked = all ? true : none ? false : null;
        }
        finally
        {
            _isUpdatingFromParent = false;
        }

        Parent?.RecalculateFromChildren();
    }

    private static ResourceNodeViewModel CreatePlaceholder() =>
        new(ResourceKind.Category, "Loading…") { IsPlaceholder = true };
}
