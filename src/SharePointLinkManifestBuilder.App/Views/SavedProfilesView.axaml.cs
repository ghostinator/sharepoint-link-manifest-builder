using Avalonia.Controls;
using Avalonia.Markup.Xaml;

namespace SharePointLinkManifestBuilder.App.Views;

/// <summary>Code-behind for <see cref="SavedProfilesView"/>. Layout and behaviour live in XAML and the view model.</summary>
public partial class SavedProfilesView : UserControl
{
    /// <summary>Creates the view.</summary>
    public SavedProfilesView() => InitializeComponent();

    private void InitializeComponent() => AvaloniaXamlLoader.Load(this);
}
