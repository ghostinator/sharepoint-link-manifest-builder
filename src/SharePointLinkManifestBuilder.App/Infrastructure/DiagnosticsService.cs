using System.Diagnostics;
using System.Globalization;
using System.IO.Compression;
using System.Runtime.InteropServices;
using System.Text;
using Microsoft.Extensions.Logging;
using SharePointLinkManifestBuilder.Core.Abstractions;
using SharePointLinkManifestBuilder.Core.Models;
using SharePointLinkManifestBuilder.Core.Security;
using SharePointLinkManifestBuilder.Core.Settings;

namespace SharePointLinkManifestBuilder.App.Infrastructure;

/// <summary>
/// Produces sanitized diagnostics for support.
/// <para>
/// The bundle is built from an explicit allow-list. It contains only what
/// <see cref="DiagnosticBundleMetadata"/> declares, and the user is shown those categories
/// before anything is written. Building it by exclusion instead — copying everything and then
/// stripping what looks sensitive — is how tenant data ends up in support tickets.
/// </para>
/// </summary>
public sealed class DiagnosticsService : IDiagnosticsService
{
    private readonly IGraphApiClient _graphClient;
    private readonly ISecureTokenStorage _tokenStorage;
    private readonly ITenantConfigurationStore _tenantStore;
    private readonly IJobHistoryStore _historyStore;
    private readonly ISettingsStore _settingsStore;
    private readonly IProductMetadataProvider _productMetadata;
    private readonly ApplicationPaths _paths;
    private readonly ILogger<DiagnosticsService> _logger;

    /// <summary>Creates the service.</summary>
    public DiagnosticsService(
        IGraphApiClient graphClient,
        ISecureTokenStorage tokenStorage,
        ITenantConfigurationStore tenantStore,
        IJobHistoryStore historyStore,
        ISettingsStore settingsStore,
        IProductMetadataProvider productMetadata,
        ApplicationPaths paths,
        ILogger<DiagnosticsService> logger)
    {
        _graphClient = graphClient ?? throw new ArgumentNullException(nameof(graphClient));
        _tokenStorage = tokenStorage ?? throw new ArgumentNullException(nameof(tokenStorage));
        _tenantStore = tenantStore ?? throw new ArgumentNullException(nameof(tenantStore));
        _historyStore = historyStore ?? throw new ArgumentNullException(nameof(historyStore));
        _settingsStore = settingsStore ?? throw new ArgumentNullException(nameof(settingsStore));
        _productMetadata = productMetadata ?? throw new ArgumentNullException(nameof(productMetadata));
        _paths = paths ?? throw new ArgumentNullException(nameof(paths));
        _logger = logger ?? throw new ArgumentNullException(nameof(logger));
    }

    /// <inheritdoc />
    public string LogDirectory => _paths.LogDirectory;

    /// <inheritdoc />
    public async Task<OperationResult<TimeSpan>> TestGraphConnectivityAsync(
        CancellationToken cancellationToken = default)
    {
        var stopwatch = Stopwatch.StartNew();

        var response = await _graphClient
            .GetAsync<object>("/me?$select=id", cancellationToken)
            .ConfigureAwait(false);

        stopwatch.Stop();

        return response.Succeeded
            ? OperationResult<TimeSpan>.Success(stopwatch.Elapsed)
            : OperationResult<TimeSpan>.Failure(response.Error ?? new GraphError
            {
                Kind = GraphErrorKind.NetworkOutage,
                Message = "The connectivity test did not succeed.",
            });
    }

    /// <inheritdoc />
    public async Task<OperationResult<string>> ExportBundleAsync(
        DiagnosticBundleMetadata metadata,
        string destinationPath,
        CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(metadata);
        ArgumentException.ThrowIfNullOrWhiteSpace(destinationPath);

        try
        {
            Directory.CreateDirectory(Path.GetDirectoryName(destinationPath)!);

            if (File.Exists(destinationPath))
            {
                // The caller confirms overwriting; silently replacing an export a user may be
                // about to send to support would be unhelpful.
                File.Delete(destinationPath);
            }

            using var archive = ZipFile.Open(destinationPath, ZipArchiveMode.Create);

            await WriteEntryAsync(archive, "summary.txt",
                await BuildSummaryAsync(metadata, cancellationToken).ConfigureAwait(false))
                .ConfigureAwait(false);

            await WriteEntryAsync(archive, "included-categories.txt", BuildCategoryManifest(metadata))
                .ConfigureAwait(false);

            await WriteEntryAsync(archive, "recent-errors.txt",
                await BuildRecentErrorsAsync(metadata, cancellationToken).ConfigureAwait(false))
                .ConfigureAwait(false);

            _logger.LogInformation("Diagnostic bundle written. Tokens and secrets are never included.");

            return OperationResult<string>.Success(destinationPath);
        }
#pragma warning disable CA1031 // Surface any export failure as a result rather than an exception.
        catch (Exception ex)
        {
            _logger.LogError(ex, "Could not write the diagnostic bundle.");

            return OperationResult<string>.Failure(new GraphError
            {
                Kind = GraphErrorKind.Unknown,
                Message = "The diagnostic bundle could not be written to that location.",
                SuggestedAction = "Choose a different folder and try again.",
            });
        }
#pragma warning restore CA1031
    }

    /// <inheritdoc />
    public Task ClearCacheAsync(CancellationToken cancellationToken = default)
    {
        // Deliberately does not touch the token cache, which has its own explicit command, so
        // "clear cached data" cannot sign a user out as a side effect.
        if (Directory.Exists(_paths.ExportsDirectory))
        {
            foreach (var file in Directory.EnumerateFiles(_paths.ExportsDirectory))
            {
                File.Delete(file);
            }
        }

        _logger.LogInformation("Local cached data cleared. Sign-in details were not affected.");
        return Task.CompletedTask;
    }

    private async Task<string> BuildSummaryAsync(
        DiagnosticBundleMetadata metadata,
        CancellationToken cancellationToken)
    {
        var tenant = await _tenantStore.LoadAsync(cancellationToken).ConfigureAwait(false);
        var settings = await _settingsStore.LoadAsync(cancellationToken).ConfigureAwait(false);
        var history = await _historyStore.ListAsync(cancellationToken).ConfigureAwait(false);
        var latest = history.Count > 0 ? history[0] : null;

        var builder = new StringBuilder();
        builder.AppendLine("SharePoint Link Manifest Builder - diagnostic summary");
        builder.AppendLine(CultureInfo.InvariantCulture, $"Generated: {metadata.GeneratedUtc:O}");
        builder.AppendLine();

        builder.AppendLine("Environment");
        builder.AppendLine(CultureInfo.InvariantCulture, $"  Application version : {metadata.ApplicationVersion}");
        builder.AppendLine(CultureInfo.InvariantCulture, $"  Platform            : {metadata.Platform}");
        builder.AppendLine(CultureInfo.InvariantCulture, $"  Runtime             : {metadata.RuntimeVersion}");
        builder.AppendLine(CultureInfo.InvariantCulture, $"  Architecture        : {RuntimeInformation.OSArchitecture}");
        builder.AppendLine();

        builder.AppendLine("Microsoft 365 connection");
        builder.AppendLine(CultureInfo.InvariantCulture, $"  Configured          : {(tenant is not null ? "yes" : "no")}");

        if (tenant is not null)
        {
            // Tenant and client IDs are not secret, but they identify the customer, so they are
            // masked unless the user opted in to including identifying values.
            builder.AppendLine(CultureInfo.InvariantCulture,
                $"  Tenant ID           : {(metadata.IncludeFullUrls ? tenant.TenantId : SensitiveDataRedactor.MaskIdentifier(tenant.TenantId))}");
            builder.AppendLine(CultureInfo.InvariantCulture,
                $"  Client ID           : {(metadata.IncludeFullUrls ? tenant.ClientId : SensitiveDataRedactor.MaskIdentifier(tenant.ClientId))}");
            builder.AppendLine(CultureInfo.InvariantCulture, $"  Registration source : {tenant.Source}");
            builder.AppendLine(CultureInfo.InvariantCulture, $"  Consent state       : {tenant.ConsentState}");
            builder.AppendLine(CultureInfo.InvariantCulture, $"  Consent type        : {tenant.ConsentType}");
            builder.AppendLine(CultureInfo.InvariantCulture, $"  Last verified       : {tenant.LastVerifiedUtc:O}");
            builder.AppendLine(CultureInfo.InvariantCulture, $"  Required scopes     : {string.Join(", ", tenant.RequiredScopes)}");
            builder.AppendLine(CultureInfo.InvariantCulture, $"  Granted scopes      : {string.Join(", ", tenant.GrantedScopes)}");
            builder.AppendLine(CultureInfo.InvariantCulture, $"  Missing scopes      : {string.Join(", ", tenant.MissingScopes)}");
        }

        builder.AppendLine();
        builder.AppendLine("Secure storage");
        builder.AppendLine(CultureInfo.InvariantCulture, $"  Availability        : {_tokenStorage.Status.Availability}");
        builder.AppendLine(CultureInfo.InvariantCulture, $"  Mechanism           : {_tokenStorage.Status.Mechanism}");
        builder.AppendLine();

        builder.AppendLine("Settings");
        builder.AppendLine(CultureInfo.InvariantCulture, $"  Theme               : {settings.Theme}");
        builder.AppendLine(CultureInfo.InvariantCulture, $"  Telemetry enabled   : {settings.TelemetryEnabled}");
        builder.AppendLine(CultureInfo.InvariantCulture, $"  Log level           : {settings.LogLevel}");
        builder.AppendLine(CultureInfo.InvariantCulture, $"  Max concurrency     : {settings.DefaultExecution.MaxConcurrency}");
        builder.AppendLine();

        builder.AppendLine("Most recent job");

        if (latest is null)
        {
            builder.AppendLine("  (no jobs have been run)");
        }
        else
        {
            builder.AppendLine(CultureInfo.InvariantCulture, $"  Started             : {latest.StartedUtc:O}");
            builder.AppendLine(CultureInfo.InvariantCulture, $"  Final phase         : {latest.FinalPhase}");
            builder.AppendLine(CultureInfo.InvariantCulture, $"  Dry run             : {latest.WasDryRun}");
            builder.AppendLine(CultureInfo.InvariantCulture, $"  Created / reused    : {latest.CreatedCount} / {latest.ReusedCount}");
            builder.AppendLine(CultureInfo.InvariantCulture, $"  Skipped / failed    : {latest.SkippedCount} / {latest.FailedCount}");
            builder.AppendLine(CultureInfo.InvariantCulture, $"  Targets             : {latest.TargetDescriptions.Count}");
        }

        return builder.ToString();
    }

    private async Task<string> BuildRecentErrorsAsync(
        DiagnosticBundleMetadata metadata,
        CancellationToken cancellationToken)
    {
        var history = await _historyStore.ListAsync(cancellationToken).ConfigureAwait(false);

        var builder = new StringBuilder();
        builder.AppendLine("Sanitized recent errors");
        builder.AppendLine("Every line below has been passed through the redaction filter.");
        builder.AppendLine();

        foreach (var entry in history.Take(10))
        {
            builder.AppendLine(CultureInfo.InvariantCulture, $"Job started {entry.StartedUtc:O} ({entry.FinalPhase})");

            foreach (var error in entry.SanitizedErrors.Take(25))
            {
                builder.AppendLine(CultureInfo.InvariantCulture, $"  - {SensitiveDataRedactor.Redact(error)}");
            }

            if (entry.SanitizedErrors.Count == 0)
            {
                builder.AppendLine("  (no errors recorded)");
            }

            builder.AppendLine();
        }

        if (!metadata.IncludeFileNames)
        {
            builder.AppendLine(
                "File and folder names were excluded from this bundle because that option was not enabled.");
        }

        return builder.ToString();
    }

    private static string BuildCategoryManifest(DiagnosticBundleMetadata metadata)
    {
        var builder = new StringBuilder();

        builder.AppendLine("Categories included in this bundle");
        builder.AppendLine();
        builder.AppendLine("Always included:");

        foreach (var category in DiagnosticBundleMetadata.AlwaysIncluded)
        {
            builder.AppendLine(CultureInfo.InvariantCulture, $"  - {category}");
        }

        builder.AppendLine();
        builder.AppendLine("Included because you approved them:");

        var optional = metadata.OptionalIncluded();

        if (optional.Count == 0)
        {
            builder.AppendLine("  (none)");
        }
        else
        {
            foreach (var category in optional)
            {
                builder.AppendLine(CultureInfo.InvariantCulture, $"  - {category}");
            }
        }

        builder.AppendLine();
        builder.AppendLine("Never included, under any option:");

        foreach (var category in DiagnosticBundleMetadata.NeverIncluded)
        {
            builder.AppendLine(CultureInfo.InvariantCulture, $"  - {category}");
        }

        return builder.ToString();
    }

    private static async Task WriteEntryAsync(ZipArchive archive, string name, string content)
    {
        var entry = archive.CreateEntry(name, CompressionLevel.Optimal);
        await using var stream = entry.Open();
        await using var writer = new StreamWriter(stream, new UTF8Encoding(false));
        await writer.WriteAsync(content).ConfigureAwait(false);
    }
}
