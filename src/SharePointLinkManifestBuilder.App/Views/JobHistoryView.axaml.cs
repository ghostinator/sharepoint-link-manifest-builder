using Avalonia.Controls;
using Avalonia.Markup.Xaml;

namespace SharePointLinkManifestBuilder.App.Views;

/// <summary>Code-behind for <see cref="JobHistoryView"/>. Layout and behaviour live in XAML and the view model.</summary>
public partial class JobHistoryView : UserControl
{
    /// <summary>Creates the view.</summary>
    public JobHistoryView() => InitializeComponent();

    private void InitializeComponent() => AvaloniaXamlLoader.Load(this);
}
