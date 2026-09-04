using Avalonia.Controls;
using Avalonia.Markup.Xaml;

namespace SharePointLinkManifestBuilder.App.Views;

/// <summary>Code-behind for <see cref="TenantSetupView"/>. Layout and behaviour live in XAML and the view model.</summary>
public partial class TenantSetupView : UserControl
{
    /// <summary>Creates the view.</summary>
    public TenantSetupView() => InitializeComponent();

    private void InitializeComponent() => AvaloniaXamlLoader.Load(this);
}
