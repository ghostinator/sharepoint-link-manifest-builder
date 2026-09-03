using System.Text.Json;
using System.Text.Json.Serialization;
using Microsoft.Extensions.Logging;
using SharePointLinkManifestBuilder.Core.Abstractions;
using SharePointLinkManifestBuilder.Core.Models;
using SharePointLinkManifestBuilder.Core.Security;

namespace SharePointLinkManifestBuilder.Core.Settings;

/// <summary>Shared JSON conventions and safe file handling for the local stores.</summary>
public static class LocalStoreJson
{
    /// <summary>Serializer options used for every local file this application writes.</summary>
    public static readonly JsonSerializerOptions Options = new()
    {
        WriteIndented = true,
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        PropertyNameCaseInsensitive = true,
        DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull,
        Converters = { new JsonStringEnumConverter() },
    };

    /// <summary>
    /// Writes a file atomically: serialize to a temporary file in the same directory, then
    /// replace. A crash mid-write therefore leaves the previous file intact rather than a
    /// truncated one, which matters most for job history a user may need afterwards.
    /// </summary>
    public static async Task WriteAtomicAsync<T>(
        string path,
        T value,
        CancellationToken cancellationToken = default)
    {
        var directory = Path.GetDirectoryName(path);

        if (!string.IsNullOrEmpty(directory))
        {
            Directory.CreateDirectory(directory);
        }

        var temporaryPath = path + ".tmp";

        await using (var stream = File.Create(temporaryPath))
        {
            await JsonSerializer.SerializeAsync(stream, value, Options, cancellationToken).ConfigureAwait(false);
        }

        File.Move(temporaryPath, path, overwrite: true);
    }

    /// <summary>
    /// Reads and deserializes a file, returning the fallback when it is missing or corrupt.
    /// Local state is a convenience, so a damaged file resets rather than blocking startup.
    /// </summary>
    public static async Task<T?> ReadOrDefaultAsync<T>(
        string path,
        ILogger logger,
        CancellationToken cancellationToken = default)
    {
        if (!File.Exists(path))
        {
            return default;
        }

        try
        {
            await using var stream = File.OpenRead(path);
            return await JsonSerializer.DeserializeAsync<T>(stream, Options, cancellationToken)
                .ConfigureAwait(false);
        }
        catch (Exception ex) when (ex is JsonException or IOException or UnauthorizedAccessException)
        {
            logger.LogWarning(
                ex,
                "A local state file could not be read and will be treated as absent: {FileName}",
                Path.GetFileName(path));

            return default;
        }
    }
}

/// <summary>Stores non-secret application settings as JSON.</summary>
public sealed class JsonSettingsStore : ISettingsStore
{
    private readonly ApplicationPaths _paths;
    private readonly ILogger<JsonSettingsStore> _logger;

    /// <summary>Creates the store.</summary>
    public JsonSettingsStore(ApplicationPaths paths, ILogger<JsonSettingsStore> logger)
    {
        _paths = paths ?? throw new ArgumentNullException(nameof(paths));
        _logger = logger ?? throw new ArgumentNullException(nameof(logger));
    }

    /// <inheritdoc />
    public string StorageDirectory => _paths.RootDirectory;

    /// <inheritdoc />
    public async Task<ApplicationSettings> LoadAsync(CancellationToken cancellationToken = default) =>
        await LocalStoreJson
            .ReadOrDefaultAsync<ApplicationSettings>(_paths.SettingsFile, _logger, cancellationToken)
            .ConfigureAwait(false)
        ?? new ApplicationSettings();

    /// <inheritdoc />
    public Task SaveAsync(ApplicationSettings settings, CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(settings);
        return LocalStoreJson.WriteAtomicAsync(_paths.SettingsFile, settings, cancellationToken);
    }
}

/// <summary>
/// Stores the tenant configuration as JSON. This file holds a tenant ID and a client ID, which
/// are identifying but not secret. No token, secret or certificate is ever written here.
/// </summary>
public sealed class JsonTenantConfigurationStore : ITenantConfigurationStore
{
    private readonly ApplicationPaths _paths;
    private readonly ILogger<JsonTenantConfigurationStore> _logger;

    /// <summary>Creates the store.</summary>
    public JsonTenantConfigurationStore(ApplicationPaths paths, ILogger<JsonTenantConfigurationStore> logger)
    {
        _paths = paths ?? throw new ArgumentNullException(nameof(paths));
        _logger = logger ?? throw new ArgumentNullException(nameof(logger));
    }

    /// <inheritdoc />
    public Task<TenantConfiguration?> LoadAsync(CancellationToken cancellationToken = default) =>
        LocalStoreJson.ReadOrDefaultAsync<TenantConfiguration>(
            _paths.TenantConfigurationFile, _logger, cancellationToken);

    /// <inheritdoc />
    public Task SaveAsync(TenantConfiguration configuration, CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(configuration);
        return LocalStoreJson.WriteAtomicAsync(
            _paths.TenantConfigurationFile, configuration, cancellationToken);
    }

    /// <inheritdoc />
    public Task RemoveAsync(CancellationToken cancellationToken = default)
    {
        if (File.Exists(_paths.TenantConfigurationFile))
        {
            File.Delete(_paths.TenantConfigurationFile);
            _logger.LogInformation(
                "Local tenant configuration removed. Nothing in the Microsoft 365 tenant was changed.");
        }

        return Task.CompletedTask;
    }
}

/// <summary>Stores saved job profiles, one file per profile.</summary>
public sealed class JsonProfileStore : IProfileStore
{
    private readonly ApplicationPaths _paths;
    private readonly ILogger<JsonProfileStore> _logger;

    /// <summary>Creates the store.</summary>
    public JsonProfileStore(ApplicationPaths paths, ILogger<JsonProfileStore> logger)
    {
        _paths = paths ?? throw new ArgumentNullException(nameof(paths));
        _logger = logger ?? throw new ArgumentNullException(nameof(logger));
    }

    /// <inheritdoc />
    public async Task<IReadOnlyList<SavedProfile>> ListAsync(CancellationToken cancellationToken = default)
    {
        if (!Directory.Exists(_paths.ProfilesDirectory))
        {
            return [];
        }

        var profiles = new List<SavedProfile>();

        foreach (var file in Directory.EnumerateFiles(_paths.ProfilesDirectory, "*.json"))
        {
            var profile = await LocalStoreJson
                .ReadOrDefaultAsync<SavedProfile>(file, _logger, cancellationToken)
                .ConfigureAwait(false);

            if (profile is not null)
            {
                profiles.Add(profile);
            }
        }

        return profiles.OrderBy(p => p.Name, StringComparer.CurrentCultureIgnoreCase).ToArray();
    }

    /// <inheritdoc />
    public Task SaveAsync(SavedProfile profile, CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(profile);
        return LocalStoreJson.WriteAtomicAsync(PathFor(profile.ProfileId), profile, cancellationToken);
    }

    /// <inheritdoc />
    public Task DeleteAsync(string profileId, CancellationToken cancellationToken = default)
    {
        var path = PathFor(profileId);

        if (File.Exists(path))
        {
            File.Delete(path);
        }

        return Task.CompletedTask;
    }

    private string PathFor(string profileId) =>
        Path.Combine(_paths.ProfilesDirectory, SafePathBuilder.MakeSafeFileName(profileId, "profile") + ".json");
}

/// <summary>Stores job history as a single JSON list, newest first.</summary>
public sealed class JsonJobHistoryStore : IJobHistoryStore, IDisposable
{
    private readonly ApplicationPaths _paths;
    private readonly ILogger<JsonJobHistoryStore> _logger;
    private readonly SemaphoreSlim _gate = new(1, 1);

    /// <summary>Creates the store.</summary>
    public JsonJobHistoryStore(ApplicationPaths paths, ILogger<JsonJobHistoryStore> logger)
    {
        _paths = paths ?? throw new ArgumentNullException(nameof(paths));
        _logger = logger ?? throw new ArgumentNullException(nameof(logger));
    }

    private string HistoryFile => Path.Combine(_paths.HistoryDirectory, "job-history.json");

    /// <inheritdoc />
    public async Task<IReadOnlyList<JobHistoryEntry>> ListAsync(CancellationToken cancellationToken = default) =>
        await LocalStoreJson
            .ReadOrDefaultAsync<List<JobHistoryEntry>>(HistoryFile, _logger, cancellationToken)
            .ConfigureAwait(false)
        ?? [];

    /// <inheritdoc />
    public async Task AppendAsync(
        JobHistoryEntry entry,
        int retentionCount,
        CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(entry);

        await _gate.WaitAsync(cancellationToken).ConfigureAwait(false);

        try
        {
            var existing = (await ListAsync(cancellationToken).ConfigureAwait(false)).ToList();
            existing.Insert(0, entry);

            // Retention is applied on write rather than on read, so the file cannot grow
            // without bound between sessions. Zero means keep everything.
            if (retentionCount > 0 && existing.Count > retentionCount)
            {
                existing = existing.Take(retentionCount).ToList();
            }

            await LocalStoreJson.WriteAtomicAsync(HistoryFile, existing, cancellationToken)
                .ConfigureAwait(false);
        }
        finally
        {
            _gate.Release();
        }
    }

    /// <inheritdoc />
    public async Task DeleteAsync(string jobId, CancellationToken cancellationToken = default)
    {
        await _gate.WaitAsync(cancellationToken).ConfigureAwait(false);

        try
        {
            var existing = (await ListAsync(cancellationToken).ConfigureAwait(false))
                .Where(e => !string.Equals(e.JobId, jobId, StringComparison.Ordinal))
                .ToList();

            await LocalStoreJson.WriteAtomicAsync(HistoryFile, existing, cancellationToken)
                .ConfigureAwait(false);
        }
        finally
        {
            _gate.Release();
        }
    }

    /// <inheritdoc />
    public async Task ClearAsync(CancellationToken cancellationToken = default)
    {
        await _gate.WaitAsync(cancellationToken).ConfigureAwait(false);

        try
        {
            if (File.Exists(HistoryFile))
            {
                File.Delete(HistoryFile);
            }

            _logger.LogInformation("Job history cleared.");
        }
        finally
        {
            _gate.Release();
        }
    }

    /// <summary>Releases the synchronization primitive.</summary>
    public void Dispose() => _gate.Dispose();
}

/// <summary>
/// Stores the local audit trail of tenant modifications. Append-only by design: an audit
/// record this application can quietly rewrite is not much of an audit record.
/// </summary>
public sealed class JsonRegistrationAuditStore : IRegistrationAuditStore, IDisposable
{
    private readonly ApplicationPaths _paths;
    private readonly ILogger<JsonRegistrationAuditStore> _logger;
    private readonly SemaphoreSlim _gate = new(1, 1);

    /// <summary>Creates the store.</summary>
    public JsonRegistrationAuditStore(ApplicationPaths paths, ILogger<JsonRegistrationAuditStore> logger)
    {
        _paths = paths ?? throw new ArgumentNullException(nameof(paths));
        _logger = logger ?? throw new ArgumentNullException(nameof(logger));
    }

    private string AuditFile => Path.Combine(_paths.AuditDirectory, "registration-audit.json");

    /// <inheritdoc />
    public async Task<IReadOnlyList<RegistrationAuditEntry>> ListAsync(
        CancellationToken cancellationToken = default) =>
        await LocalStoreJson
            .ReadOrDefaultAsync<List<RegistrationAuditEntry>>(AuditFile, _logger, cancellationToken)
            .ConfigureAwait(false)
        ?? [];

    /// <inheritdoc />
    public async Task AppendAsync(RegistrationAuditEntry entry, CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(entry);

        await _gate.WaitAsync(cancellationToken).ConfigureAwait(false);

        try
        {
            var existing = (await ListAsync(cancellationToken).ConfigureAwait(false)).ToList();
            existing.Insert(0, entry);

            await LocalStoreJson.WriteAtomicAsync(AuditFile, existing, cancellationToken)
                .ConfigureAwait(false);

            _logger.LogInformation(
                "Recorded a tenant modification in the local audit history: {Action} ({Outcome}).",
                entry.Action,
                entry.Succeeded ? "succeeded" : "failed");
        }
        finally
        {
            _gate.Release();
        }
    }

    /// <summary>Releases the synchronization primitive.</summary>
    public void Dispose() => _gate.Dispose();
}
