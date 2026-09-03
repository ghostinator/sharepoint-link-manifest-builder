using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Logging;
using SharePointLinkManifestBuilder.App.Infrastructure;
using SharePointLinkManifestBuilder.App.ViewModels;
using SharePointLinkManifestBuilder.Core.Abstractions;
using SharePointLinkManifestBuilder.Core.Jobs;
using SharePointLinkManifestBuilder.Core.Manifests;
using SharePointLinkManifestBuilder.Core.Models;
using SharePointLinkManifestBuilder.Core.Resilience;
using SharePointLinkManifestBuilder.Core.Settings;
using SharePointLinkManifestBuilder.Graph.Http;
using SharePointLinkManifestBuilder.Graph.Identity;
using SharePointLinkManifestBuilder.Graph.Onboarding;
using SharePointLinkManifestBuilder.Graph.Services;

namespace SharePointLinkManifestBuilder.App.Composition;

/// <summary>
/// The composition root. Every dependency is registered here and nowhere else, so the whole
/// object graph can be read in one place and nothing resolves a service locator at runtime.
/// </summary>
public static class ServiceRegistration
{
    /// <summary>Environment variable a publisher can use to supply the bootstrap client ID.</summary>
    public const string BootstrapClientIdVariable = "SPLMB_BOOTSTRAP_CLIENT_ID";

    /// <summary>Builds the application's service provider.</summary>
    /// <param name="paths">Local state locations; tests override this to a temporary directory.</param>
    public static ServiceProvider Build(ApplicationPaths? paths = null)
    {
        var services = new ServiceCollection();
        var applicationPaths = paths ?? new ApplicationPaths();
        applicationPaths.EnsureCreated();

        var configuration = BuildConfiguration(applicationPaths);

        services.AddSingleton<IConfiguration>(configuration);
        services.AddSingleton(applicationPaths);

        AddLogging(services, applicationPaths, configuration);
        AddCoreServices(services);
        AddGraphServices(services, configuration);
        AddViewModels(services);

        return services.BuildServiceProvider(new ServiceProviderOptions
        {
            ValidateOnBuild = true,
            ValidateScopes = true,
        });
    }

    /// <summary>
    /// Configuration sources, in increasing order of precedence: bundled defaults, a local
    /// override file, then environment variables. Environment variables win so a publisher or
    /// an administrator can set the bootstrap client ID without editing a file.
    /// </summary>
    public static IConfigurationRoot BuildConfiguration(ApplicationPaths paths)
    {
        ArgumentNullException.ThrowIfNull(paths);

        return new ConfigurationBuilder()
            .SetBasePath(AppContext.BaseDirectory)
            .AddJsonFile("appsettings.json", optional: true, reloadOnChange: false)
            .AddJsonFile(Path.Combine(paths.RootDirectory, "appsettings.Local.json"),
                optional: true, reloadOnChange: false)
            .AddEnvironmentVariables()
            .Build();
    }

    private static void AddLogging(
        IServiceCollection services,
        ApplicationPaths paths,
        IConfiguration configuration)
    {
        var configuredLevel = configuration["Logging:LogLevel:Default"];

        var minimumLevel = Enum.TryParse<LogLevel>(configuredLevel, ignoreCase: true, out var parsed)
            ? parsed
            : LogLevel.Information;

        services.AddLogging(builder =>
        {
            builder.SetMinimumLevel(minimumLevel);
            builder.AddProvider(new FileLoggerProvider(paths.LogDirectory, minimumLevel));

            // Redaction wraps every provider registered above it. Adding it last is what
            // guarantees nothing bypasses the filter, including third-party libraries.
            builder.AddRedaction();
        });
    }

    private static void AddCoreServices(IServiceCollection services)
    {
        services.AddSingleton(TimeProvider.System);

        services.AddSingleton<ISettingsStore, JsonSettingsStore>();
        services.AddSingleton<ITenantConfigurationStore, JsonTenantConfigurationStore>();
        services.AddSingleton<IProfileStore, JsonProfileStore>();
        services.AddSingleton<IJobHistoryStore, JsonJobHistoryStore>();
        services.AddSingleton<IRegistrationAuditStore, JsonRegistrationAuditStore>();

        // One instance per format, resolved as a set by the job runner.
        services.AddSingleton<IManifestFormatter, PlainTextManifestFormatter>();
        services.AddSingleton<IManifestFormatter, MarkdownManifestFormatter>();
        services.AddSingleton<IManifestFormatter, CsvManifestFormatter>();
        services.AddSingleton<IManifestFormatter, JsonManifestFormatter>();

        services.AddSingleton<IManifestParser, PlainTextManifestParser>();
        services.AddSingleton<IManifestMerger, ManifestMerger>();
        services.AddSingleton<IManifestBuilder, ManifestBuilder>();
        services.AddSingleton<ManifestConflictResolver>();

        services.AddSingleton<IFileDiscoveryService, FileDiscoveryService>();
        services.AddSingleton<ILinkJobRunner, LinkJobRunner>();

        // The in-progress job is shared: the browsers add targets to it and the job page reads it.
        services.AddSingleton<JobDraft>();

        services.AddSingleton<ISystemBrowser, SystemBrowser>();
        services.AddSingleton<IClipboardService, ClipboardService>();
        services.AddSingleton<FolderLauncher>();
        services.AddSingleton<IDiagnosticsService, DiagnosticsService>();
    }

    private static void AddGraphServices(IServiceCollection services, IConfiguration configuration)
    {
        services.AddSingleton<GraphClientContext>();

        services.AddSingleton(sp => new SecureTokenStorage(
            sp.GetRequiredService<ApplicationPaths>().TokenCacheDirectory,
            sp.GetRequiredService<ILogger<SecureTokenStorage>>()));

        services.AddSingleton<ISecureTokenStorage>(sp => sp.GetRequiredService<SecureTokenStorage>());

        services.AddSingleton<MsalAuthenticationService>();
        services.AddSingleton<IAuthenticationService>(sp => sp.GetRequiredService<MsalAuthenticationService>());

        // Retry defaults are conservative: this tool can generate a great many writes, and
        // Graph throttling is a shared tenant resource rather than a per-application budget.
        services.AddSingleton(new GraphRetryPolicy(
            maxAttempts: 5,
            baseDelay: TimeSpan.FromSeconds(1),
            maxDelay: TimeSpan.FromSeconds(60)));

        services.AddHttpClient<IGraphApiClient, GraphApiClient>(client =>
        {
            client.Timeout = TimeSpan.FromMinutes(2);
            client.DefaultRequestHeaders.Add("Accept", "application/json");
        });

        services.AddSingleton<ISiteService, SiteService>();
        services.AddSingleton<IDriveService, DriveService>();
        services.AddSingleton<IUserDirectoryService, UserDirectoryService>();
        services.AddSingleton<ISharingLinkService, SharingLinkService>();
        services.AddSingleton<IManifestStorageService, ManifestStorageService>();

        services.AddSingleton<IAppRegistrationService, AppRegistrationService>();
        services.AddSingleton<IConsentService, ConsentService>();

        // No client ID ships in this repository. Automatic setup stays unavailable until a
        // publisher supplies one through configuration or the wizard's Advanced field.
        var bootstrapClientId = configuration["Bootstrap:ClientId"]
            ?? Environment.GetEnvironmentVariable(BootstrapClientIdVariable);

        services.AddSingleton<IBootstrapConfigurationProvider>(
            new BootstrapConfigurationProvider(bootstrapClientId, configuration["Bootstrap:Instance"]));

        var metadata = new ProductMetadata();
        configuration.GetSection("Product").Bind(metadata);
        services.AddSingleton<IProductMetadataProvider>(new ProductMetadataProvider(metadata));

        services.AddSingleton<ConnectionCoordinator>();
    }

    private static void AddViewModels(IServiceCollection services)
    {
        services.AddSingleton<MainWindowViewModel>();
        services.AddSingleton<HomeViewModel>();
        services.AddSingleton<NewLinkJobViewModel>();
        services.AddSingleton<SharePointBrowserViewModel>();
        services.AddSingleton<OneDriveBrowserViewModel>();
        services.AddSingleton<SavedProfilesViewModel>();
        services.AddSingleton<JobHistoryViewModel>();
        services.AddSingleton<TenantSetupViewModel>();
        services.AddSingleton<PermissionsViewModel>();
        services.AddSingleton<SettingsViewModel>();
        services.AddSingleton<DiagnosticsViewModel>();
        services.AddSingleton<HelpViewModel>();
        services.AddSingleton<AboutViewModel>();
    }
}
