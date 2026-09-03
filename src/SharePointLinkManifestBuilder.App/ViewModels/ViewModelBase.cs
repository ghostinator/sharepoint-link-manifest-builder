using CommunityToolkit.Mvvm.ComponentModel;

namespace SharePointLinkManifestBuilder.App.ViewModels;

/// <summary>Base for every view model.</summary>
public abstract partial class ViewModelBase : ObservableObject
{
    /// <summary>True while a long-running operation is in flight.</summary>
    [ObservableProperty]
    private bool _isBusy;

    /// <summary>A short status line shown near the top of the page.</summary>
    [ObservableProperty]
    private string? _statusMessage;

    /// <summary>
    /// A user-facing error. Kept separate from <see cref="StatusMessage"/> so the view can
    /// present it with an alert role, which screen readers announce.
    /// </summary>
    [ObservableProperty]
    private string? _errorMessage;

    /// <summary>True when an error is being shown.</summary>
    public bool HasError => !string.IsNullOrEmpty(ErrorMessage);

    /// <summary>Clears any error and status text.</summary>
    public void ClearMessages()
    {
        ErrorMessage = null;
        StatusMessage = null;
    }

    partial void OnErrorMessageChanged(string? value) => OnPropertyChanged(nameof(HasError));
}

/// <summary>Base for a navigable page.</summary>
public abstract partial class PageViewModelBase : ViewModelBase
{
    /// <summary>Creates the page.</summary>
    /// <param name="title">Title shown in the navigation list and page header.</param>
    /// <param name="navigationKey">Stable key used for navigation and testing.</param>
    protected PageViewModelBase(string title, string navigationKey)
    {
        Title = title;
        NavigationKey = navigationKey;
    }

    /// <summary>The page title.</summary>
    public string Title { get; }

    /// <summary>A stable identifier for this page.</summary>
    public string NavigationKey { get; }

    /// <summary>
    /// Called each time the page becomes visible. Loading here rather than in the constructor
    /// keeps startup fast and means a page reflects current state whenever it is revisited.
    /// </summary>
    public virtual Task OnNavigatedToAsync(CancellationToken cancellationToken = default) =>
        Task.CompletedTask;
}
