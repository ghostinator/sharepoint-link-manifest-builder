using Avalonia.Controls;
using Avalonia.Markup.Xaml;

namespace SharePointLinkManifestBuilder.App.Views;

/// <summary>Code-behind for <see cref="SharePointBrowserView"/>. Layout and behaviour live in XAML and the view model.</summary>
public partial class SharePointBrowserView : UserControl
{
    /// <summary>Creates the view.</summary>
    public SharePointBrowserView() => InitializeComponent();

    private void InitializeComponent() => AvaloniaXamlLoader.Load(this);
}
