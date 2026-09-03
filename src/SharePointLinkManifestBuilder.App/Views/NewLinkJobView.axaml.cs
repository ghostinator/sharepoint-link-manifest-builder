using Avalonia.Controls;
using Avalonia.Markup.Xaml;

namespace SharePointLinkManifestBuilder.App.Views;

/// <summary>Code-behind for <see cref="NewLinkJobView"/>. Layout and behaviour live in XAML and the view model.</summary>
public partial class NewLinkJobView : UserControl
{
    /// <summary>Creates the view.</summary>
    public NewLinkJobView() => InitializeComponent();

    private void InitializeComponent() => AvaloniaXamlLoader.Load(this);
}
