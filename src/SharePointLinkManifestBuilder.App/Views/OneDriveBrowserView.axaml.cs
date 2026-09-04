using Avalonia.Controls;
using Avalonia.Markup.Xaml;

namespace SharePointLinkManifestBuilder.App.Views;

/// <summary>Code-behind for <see cref="OneDriveBrowserView"/>. Layout and behaviour live in XAML and the view model.</summary>
public partial class OneDriveBrowserView : UserControl
{
    /// <summary>Creates the view.</summary>
    public OneDriveBrowserView() => InitializeComponent();

    private void InitializeComponent() => AvaloniaXamlLoader.Load(this);
}
