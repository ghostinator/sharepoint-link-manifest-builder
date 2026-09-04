using Avalonia.Controls;
using Avalonia.Markup.Xaml;

namespace SharePointLinkManifestBuilder.App.Views;

/// <summary>Code-behind for <see cref="SettingsView"/>. Layout and behaviour live in XAML and the view model.</summary>
public partial class SettingsView : UserControl
{
    /// <summary>Creates the view.</summary>
    public SettingsView() => InitializeComponent();

    private void InitializeComponent() => AvaloniaXamlLoader.Load(this);
}
