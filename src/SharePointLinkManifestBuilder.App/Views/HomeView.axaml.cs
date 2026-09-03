using Avalonia.Controls;
using Avalonia.Markup.Xaml;

namespace SharePointLinkManifestBuilder.App.Views;

/// <summary>Code-behind for <see cref="HomeView"/>. Layout and behaviour live in XAML and the view model.</summary>
public partial class HomeView : UserControl
{
    /// <summary>Creates the view.</summary>
    public HomeView() => InitializeComponent();

    private void InitializeComponent() => AvaloniaXamlLoader.Load(this);
}
