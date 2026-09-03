using Avalonia.Controls;
using Avalonia.Markup.Xaml;

namespace SharePointLinkManifestBuilder.App.Views;

/// <summary>Code-behind for <see cref="DiagnosticsView"/>. Layout and behaviour live in XAML and the view model.</summary>
public partial class DiagnosticsView : UserControl
{
    /// <summary>Creates the view.</summary>
    public DiagnosticsView() => InitializeComponent();

    private void InitializeComponent() => AvaloniaXamlLoader.Load(this);
}
