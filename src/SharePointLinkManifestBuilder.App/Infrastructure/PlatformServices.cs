using System.Diagnostics;
using System.Runtime.InteropServices;
using Avalonia;
using Avalonia.Controls.ApplicationLifetimes;
using Avalonia.Input.Platform;
using Microsoft.Extensions.Logging;
using SharePointLinkManifestBuilder.Core.Abstractions;
using SharePointLinkManifestBuilder.Core.Models;

namespace SharePointLinkManifestBuilder.App.Infrastructure;

/// <summary>
/// Opens URLs in the operating system's default browser.
/// <para>
/// Only absolute https URLs are opened. Handing an arbitrary string to the shell would let a
/// crafted value launch a local program, so the scheme is checked before anything is executed.
/// </para>
/// </summary>
public sealed class SystemBrowser : ISystemBrowser
{
    private readonly ILogger<SystemBrowser> _logger;

    /// <summary>Creates the service.</summary>
    public SystemBrowser(ILogger<SystemBrowser> logger) =>
        _logger = logger ?? throw new ArgumentNullException(nameof(logger));

    /// <inheritdoc />
    public Task OpenAsync(Uri url, CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(url);

        if (!url.IsAbsoluteUri || !string.Equals(url.Scheme, Uri.UriSchemeHttps, StringComparison.OrdinalIgnoreCase))
        {
            throw new ArgumentException(
                $"Only absolute https URLs can be opened; '{url.Scheme}' was supplied.", nameof(url));
        }

        try
        {
            // UseShellExecute lets each platform pick the user's default browser rather than
            // this application guessing at an executable name.
            using var process = Process.Start(new ProcessStartInfo
            {
                FileName = url.ToString(),
                UseShellExecute = true,
            });
        }
#pragma warning disable CA1031 // Failing to open a browser must not take the application down.
        catch (Exception ex)
        {
            _logger.LogError(ex, "Could not open the system browser. The URL can be copied manually instead.");
        }
#pragma warning restore CA1031

        return Task.CompletedTask;
    }
}

/// <summary>Opens a local folder in the platform file manager.</summary>
public sealed class FolderLauncher
{
    private readonly ILogger<FolderLauncher> _logger;

    /// <summary>Creates the service.</summary>
    public FolderLauncher(ILogger<FolderLauncher> logger) =>
        _logger = logger ?? throw new ArgumentNullException(nameof(logger));

    /// <summary>Reveals a directory in the platform file manager.</summary>
    public void OpenFolder(string path)
    {
        if (string.IsNullOrWhiteSpace(path) || !Directory.Exists(path))
        {
            _logger.LogWarning("Cannot open a folder that does not exist.");
            return;
        }

        try
        {
            var (fileName, arguments) = RuntimeInformation.IsOSPlatform(OSPlatform.Windows)
                ? ("explorer.exe", $"\"{path}\"")
                : RuntimeInformation.IsOSPlatform(OSPlatform.OSX)
                    ? ("open", $"\"{path}\"")
                    : ("xdg-open", $"\"{path}\"");

            using var process = Process.Start(new ProcessStartInfo
            {
                FileName = fileName,
                Arguments = arguments,
                UseShellExecute = false,
            });
        }
#pragma warning disable CA1031 // A missing file manager must not crash the application.
        catch (Exception ex)
        {
            _logger.LogError(ex, "Could not open the folder in the platform file manager.");
        }
#pragma warning restore CA1031
    }
}

/// <summary>Clipboard access through Avalonia's top-level clipboard.</summary>
public sealed class ClipboardService : IClipboardService
{
    private readonly ILogger<ClipboardService> _logger;

    /// <summary>Creates the service.</summary>
    public ClipboardService(ILogger<ClipboardService> logger) =>
        _logger = logger ?? throw new ArgumentNullException(nameof(logger));

    /// <inheritdoc />
    public async Task SetTextAsync(string text, CancellationToken cancellationToken = default)
    {
        var clipboard = ResolveClipboard();

        if (clipboard is null)
        {
            _logger.LogWarning("No clipboard is available on this platform.");
            return;
        }

        await clipboard.SetTextAsync(text ?? string.Empty).ConfigureAwait(false);
    }

    private static IClipboard? ResolveClipboard() =>
        Application.Current?.ApplicationLifetime is IClassicDesktopStyleApplicationLifetime desktop
            ? desktop.MainWindow?.Clipboard
            : null;
}

/// <summary>
/// Supplies product and publisher metadata.
/// <para>
/// Values come from configuration and default to obvious placeholders. They are displayed
/// verbatim, so a publisher who has not set them sees PLACEHOLDER rather than a plausible-
/// looking but wrong value.
/// </para>
/// </summary>
public sealed class ProductMetadataProvider : IProductMetadataProvider
{
    /// <summary>Creates the provider.</summary>
    /// <param name="metadata">Metadata bound from configuration.</param>
    public ProductMetadataProvider(ProductMetadata? metadata = null)
    {
        Metadata = metadata ?? new ProductMetadata();

        var informational = typeof(ProductMetadataProvider).Assembly
            .GetCustomAttributes(typeof(System.Reflection.AssemblyInformationalVersionAttribute), false)
            .OfType<System.Reflection.AssemblyInformationalVersionAttribute>()
            .FirstOrDefault()?.InformationalVersion;

        // The informational version carries a build hash suffix such as "0.1.0+abc1234";
        // the UI wants the semantic part only.
        Version = informational?.Split('+')[0]
            ?? typeof(ProductMetadataProvider).Assembly.GetName().Version?.ToString(3)
            ?? "0.0.0";
    }

    /// <inheritdoc />
    public ProductMetadata Metadata { get; }

    /// <inheritdoc />
    public string Version { get; }
}
