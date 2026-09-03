using Avalonia.Controls;
using Avalonia.Markup.Xaml;

namespace SharePointLinkManifestBuilder.App.Views;

/// <summary>Code-behind for <see cref="PermissionsView"/>. Layout and behaviour live in XAML and the view model.</summary>
public partial class PermissionsView : UserControl
{
    /// <summary>Creates the view.</summary>
    public PermissionsView() => InitializeComponent();

    private void InitializeComponent() => AvaloniaXamlLoader.Load(this);
}
