using Avalonia.Controls;
using Avalonia.Markup.Xaml;

namespace SharePointLinkManifestBuilder.App.Views;

/// <summary>Code-behind for <see cref="AboutView"/>. Layout and behaviour live in XAML and the view model.</summary>
public partial class AboutView : UserControl
{
    /// <summary>Creates the view.</summary>
    public AboutView() => InitializeComponent();

    private void InitializeComponent() => AvaloniaXamlLoader.Load(this);
}
