using SharePointLinkManifestBuilder.Core.Models;

namespace SharePointLinkManifestBuilder.Core.Abstractions;

/// <summary>Opens URLs in the operating system's default browser.</summary>
public interface ISystemBrowser
{
    /// <summary>Opens an absolute https URL. Refuses any other scheme.</summary>
    Task OpenAsync(Uri url, CancellationToken cancellationToken = default);
}

/// <summary>Clipboard access, abstracted so view models stay testable.</summary>
public interface IClipboardService
{
    /// <summary>Copies text to the clipboard.</summary>
    Task SetTextAsync(string text, CancellationToken cancellationToken = default);
}

/// <summary>Persists non-secret application settings.</summary>
public interface ISettingsStore
{
    /// <summary>Loads settings, returning defaults when none are stored.</summary>
    Task<ApplicationSettings> LoadAsync(CancellationToken cancellationToken = default);

    /// <summary>Saves settings. Never writes a token or secret.</summary>
    Task SaveAsync(ApplicationSettings settings, CancellationToken cancellationToken = default);

    /// <summary>The directory holding local state, shown to the user on the privacy page.</summary>
    string StorageDirectory { get; }
}

/// <summary>Persists the tenant configuration. Contains no secret.</summary>
public interface ITenantConfigurationStore
{
    /// <summary>Loads the stored configuration, or null when none exists.</summary>
    Task<TenantConfiguration?> LoadAsync(CancellationToken cancellationToken = default);

    /// <summary>Saves the configuration.</summary>
    Task SaveAsync(TenantConfiguration configuration, CancellationToken cancellationToken = default);

    /// <summary>Removes the local configuration. Does not change anything in the tenant.</summary>
    Task RemoveAsync(CancellationToken cancellationToken = default);
}

/// <summary>Persists saved job profiles.</summary>
public interface IProfileStore
{
    /// <summary>Lists all profiles.</summary>
    Task<IReadOnlyList<SavedProfile>> ListAsync(CancellationToken cancellationToken = default);

    /// <summary>Creates or replaces a profile.</summary>
    Task SaveAsync(SavedProfile profile, CancellationToken cancellationToken = default);

    /// <summary>Deletes a profile.</summary>
    Task DeleteAsync(string profileId, CancellationToken cancellationToken = default);
}

/// <summary>Persists job history. Never stores a credential or a complete token.</summary>
public interface IJobHistoryStore
{
    /// <summary>Lists history entries, most recent first.</summary>
    Task<IReadOnlyList<JobHistoryEntry>> ListAsync(CancellationToken cancellationToken = default);

    /// <summary>Appends an entry, trimming to the configured retention count.</summary>
    Task AppendAsync(JobHistoryEntry entry, int retentionCount, CancellationToken cancellationToken = default);

    /// <summary>Deletes one entry.</summary>
    Task DeleteAsync(string jobId, CancellationToken cancellationToken = default);

    /// <summary>Deletes every entry.</summary>
    Task ClearAsync(CancellationToken cancellationToken = default);
}

/// <summary>Persists the local audit trail of tenant modifications.</summary>
public interface IRegistrationAuditStore
{
    /// <summary>Lists audit entries, most recent first.</summary>
    Task<IReadOnlyList<RegistrationAuditEntry>> ListAsync(CancellationToken cancellationToken = default);

    /// <summary>Appends an audit entry.</summary>
    Task AppendAsync(RegistrationAuditEntry entry, CancellationToken cancellationToken = default);
}

/// <summary>Produces sanitized diagnostics for support.</summary>
public interface IDiagnosticsService
{
    /// <summary>Runs a live Graph connectivity check.</summary>
    Task<OperationResult<TimeSpan>> TestGraphConnectivityAsync(CancellationToken cancellationToken = default);

    /// <summary>
    /// Writes a diagnostic bundle containing only the categories declared in
    /// <paramref name="metadata"/>, which the user has already been shown.
    /// </summary>
    Task<OperationResult<string>> ExportBundleAsync(
        DiagnosticBundleMetadata metadata,
        string destinationPath,
        CancellationToken cancellationToken = default);

    /// <summary>The directory holding log files, for "Open Log Folder".</summary>
    string LogDirectory { get; }

    /// <summary>Deletes cached data. Does not touch the token cache, which has its own action.</summary>
    Task ClearCacheAsync(CancellationToken cancellationToken = default);
}

/// <summary>Supplies product and publisher metadata.</summary>
public interface IProductMetadataProvider
{
    /// <summary>Product and publisher metadata, which may contain placeholders.</summary>
    ProductMetadata Metadata { get; }

    /// <summary>The application's semantic version.</summary>
    string Version { get; }
}
