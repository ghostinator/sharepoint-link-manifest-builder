using Avalonia;
using Avalonia.Controls.ApplicationLifetimes;
using Avalonia.Markup.Xaml;
using Microsoft.Extensions.DependencyInjection;
using SharePointLinkManifestBuilder.App.Composition;
using SharePointLinkManifestBuilder.App.ViewModels;
using SharePointLinkManifestBuilder.App.Views;

namespace SharePointLinkManifestBuilder.App;

/// <summary>The Avalonia application.</summary>
public partial class App : Application
{
    private ServiceProvider? _services;

    /// <summary>The resolved service provider, exposed for the headless test host.</summary>
    public ServiceProvider? Services => _services;

    /// <inheritdoc />
    public override void Initialize() => AvaloniaXamlLoader.Load(this);

    /// <inheritdoc />
    public override void OnFrameworkInitializationCompleted()
    {
        if (ApplicationLifetime is IClassicDesktopStyleApplicationLifetime desktop)
        {
            _services = ServiceRegistration.Build();

            var viewModel = _services.GetRequiredService<MainWindowViewModel>();
            var window = new MainWindow { DataContext = viewModel };

            desktop.MainWindow = window;

            // Startup work runs after the window exists, so the user sees the shell immediately
            // rather than a blank screen while the token cache is probed.
            window.Opened += async (_, _) => await viewModel.InitializeAsync().ConfigureAwait(true);

            desktop.ShutdownRequested += (_, _) => _services?.Dispose();
        }

        base.OnFrameworkInitializationCompleted();
    }
}
