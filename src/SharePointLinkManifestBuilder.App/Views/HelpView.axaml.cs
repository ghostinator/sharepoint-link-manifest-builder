using Avalonia.Controls;
using Avalonia.Markup.Xaml;

namespace SharePointLinkManifestBuilder.App.Views;

/// <summary>Code-behind for <see cref="HelpView"/>. Layout and behaviour live in XAML and the view model.</summary>
public partial class HelpView : UserControl
{
    /// <summary>Creates the view.</summary>
    public HelpView() => InitializeComponent();

    private void InitializeComponent() => AvaloniaXamlLoader.Load(this);
}
