namespace SharePointLinkManifestBuilder.Core.Models;

/// <summary>
/// Publisher and product metadata. A fork that redistributes this application replaces these;
/// anything still reading PLACEHOLDER is reported rather than shown as though it were real. The
/// About and Help screens display them verbatim so a placeholder is visible rather than hidden.
/// </summary>
public sealed record ProductMetadata
{
    /// <summary>Product display name.</summary>
    public string ProductName { get; init; } = "SharePoint Link Manifest Builder";

    /// <summary>Publisher name.</summary>
    public string Publisher { get; init; } = "Brandon Cook";

    /// <summary>Product homepage.</summary>
    public string HomepageUrl { get; init; } = "https://github.com/ghostinator/sharepoint-link-manifest-builder";

    /// <summary>Support landing page.</summary>
    public string SupportUrl { get; init; } = "https://github.com/ghostinator/sharepoint-link-manifest-builder/issues";

    /// <summary>Privacy policy, shown on the consent review page.</summary>
    public string PrivacyPolicyUrl { get; init; } = "https://github.com/ghostinator/sharepoint-link-manifest-builder/blob/main/docs/PRIVACY.md";

    /// <summary>Terms of use.</summary>
    public string TermsUrl { get; init; } = "https://github.com/ghostinator/sharepoint-link-manifest-builder/blob/main/LICENSE";

    /// <summary>Source repository.</summary>
    public string SourceCodeUrl { get; init; } = "https://github.com/ghostinator/sharepoint-link-manifest-builder";

    /// <summary>Issue tracker.</summary>
    public string IssueTrackerUrl { get; init; } = "https://github.com/ghostinator/sharepoint-link-manifest-builder/issues";

    /// <summary>Endpoint consulted by a manual update check. No automatic check is performed.</summary>
    public string UpdateCheckUrl { get; init; } = "https://example.invalid/PLACEHOLDER-UPDATES";

    /// <summary>Contact address for security and support correspondence.</summary>
    public string ContactAddress { get; init; } = "github@ghostinator.co";

    /// <summary>
    /// True when the publisher identity is still unset, so the UI can say so plainly.
    /// <para>
    /// <see cref="UpdateCheckUrl"/> is deliberately excluded. It is unset until a release exists
    /// to check against, which is a different and expected condition, and the About page already
    /// reports it specifically. Including it here would keep a "publisher details are not set"
    /// warning on screen after they had been, which is the opposite of what the warning is for.
    /// </para>
    /// </summary>
    public bool HasPlaceholders =>
        new[]
        {
            Publisher, HomepageUrl, SupportUrl, PrivacyPolicyUrl, TermsUrl,
            SourceCodeUrl, IssueTrackerUrl, ContactAddress,
        }.Any(v => v.Contains("PLACEHOLDER", StringComparison.OrdinalIgnoreCase));
}

/// <summary>Application theme preference.</summary>
public enum ThemePreference
{
    /// <summary>Follow the operating system.</summary>
    System = 0,

    /// <summary>Always light.</summary>
    Light = 1,

    /// <summary>Always dark.</summary>
    Dark = 2,
}

/// <summary>Non-secret user preferences, persisted locally as JSON.</summary>
public sealed record ApplicationSettings
{
    /// <summary>Theme preference.</summary>
    public ThemePreference Theme { get; init; } = ThemePreference.System;

    /// <summary>Honour the operating system's reduced-motion setting.</summary>
    public bool ReduceMotion { get; init; }

    /// <summary>
    /// Telemetry opt-in. Disabled by default and no telemetry pipeline is implemented; the
    /// setting exists so the architecture is explicit about consent rather than implicit.
    /// </summary>
    public bool TelemetryEnabled { get; init; }

    /// <summary>Default execution options for new jobs.</summary>
    public ExecutionOptions DefaultExecution { get; init; } = new();

    /// <summary>Default link settings for new jobs.</summary>
    public LinkConfiguration DefaultLink { get; init; } = new();

    /// <summary>Default manifest settings for new jobs.</summary>
    public ManifestConfiguration DefaultManifest { get; init; } = new();

    /// <summary>Recently used locations, most recent first.</summary>
    public IReadOnlyList<RecentLocation> RecentLocations { get; init; } = [];

    /// <summary>Locations the user pinned.</summary>
    public IReadOnlyList<RecentLocation> PinnedLocations { get; init; } = [];

    /// <summary>How many job history entries to retain. Zero means unlimited.</summary>
    public int JobHistoryRetentionCount { get; init; } = 100;

    /// <summary>Minimum log level written to the local log file.</summary>
    public string LogLevel { get; init; } = "Information";

    /// <summary>True once the first-run privacy notice has been acknowledged.</summary>
    public bool PrivacyNoticeAcknowledged { get; init; }
}

/// <summary>A remembered or pinned location, stored without any token or secret.</summary>
public sealed record RecentLocation
{
    /// <summary>Friendly label shown in the UI.</summary>
    public required string DisplayName { get; init; }

    /// <summary>Which source family this location belongs to.</summary>
    public required TargetSourceType SourceType { get; init; }

    /// <summary>Tenant this location belongs to, so locations do not leak across tenants.</summary>
    public required string TenantId { get; init; }

    /// <summary>Graph site ID, for SharePoint locations.</summary>
    public string? SiteId { get; init; }

    /// <summary>Graph drive ID.</summary>
    public string? DriveId { get; init; }

    /// <summary>Graph item ID of the folder, when the location is not a drive root.</summary>
    public string? FolderItemId { get; init; }

    /// <summary>Absolute URL, for "Open in browser".</summary>
    public string? WebUrl { get; init; }

    /// <summary>When the location was last used.</summary>
    public DateTimeOffset LastUsedUtc { get; init; } = DateTimeOffset.UtcNow;
}

/// <summary>A reusable job configuration the user saved.</summary>
public sealed record SavedProfile
{
    /// <summary>Stable identifier.</summary>
    public string ProfileId { get; init; } = Guid.NewGuid().ToString("n");

    /// <summary>User-supplied name.</summary>
    public required string Name { get; init; }

    /// <summary>Optional description.</summary>
    public string? Description { get; init; }

    /// <summary>The saved configuration. Contains no token or secret.</summary>
    public required JobConfiguration Configuration { get; init; }

    /// <summary>When the profile was created.</summary>
    public DateTimeOffset CreatedUtc { get; init; } = DateTimeOffset.UtcNow;

    /// <summary>When the profile was last run.</summary>
    public DateTimeOffset? LastUsedUtc { get; init; }
}

/// <summary>A completed run, retained locally. Contains no credential or token.</summary>
public sealed record JobHistoryEntry
{
    /// <summary>The job identifier.</summary>
    public required string JobId { get; init; }

    /// <summary>Optional job name.</summary>
    public string? Name { get; init; }

    /// <summary>Application version that produced the run.</summary>
    public required string ApplicationVersion { get; init; }

    /// <summary>When the run started.</summary>
    public required DateTimeOffset StartedUtc { get; init; }

    /// <summary>When the run finished.</summary>
    public DateTimeOffset? CompletedUtc { get; init; }

    /// <summary>A privacy-conscious account identifier, never the full user principal name.</summary>
    public required string AccountIdentifier { get; init; }

    /// <summary>Tenant display name.</summary>
    public string? TenantDisplayName { get; init; }

    /// <summary>Tenant ID.</summary>
    public required string TenantId { get; init; }

    /// <summary>Friendly descriptions of the targets processed.</summary>
    public IReadOnlyList<string> TargetDescriptions { get; init; } = [];

    /// <summary>Summary of the link settings used.</summary>
    public required string LinkSettingsSummary { get; init; }

    /// <summary>Summary of the manifest settings used.</summary>
    public required string ManifestSettingsSummary { get; init; }

    /// <summary>Summary of the filters used.</summary>
    public required string FilterSummary { get; init; }

    /// <summary>Links newly created.</summary>
    public int CreatedCount { get; init; }

    /// <summary>Existing links returned rather than created.</summary>
    public int ReusedCount { get; init; }

    /// <summary>Files skipped.</summary>
    public int SkippedCount { get; init; }

    /// <summary>Files that failed.</summary>
    public int FailedCount { get; init; }

    /// <summary>Where manifests were written.</summary>
    public IReadOnlyList<string> ManifestLocations { get; init; } = [];

    /// <summary>Sanitized failure descriptions.</summary>
    public IReadOnlyList<string> SanitizedErrors { get; init; } = [];

    /// <summary>True when the run was a dry run.</summary>
    public bool WasDryRun { get; init; }

    /// <summary>True when the user stopped the run.</summary>
    public bool WasCancelled { get; init; }

    /// <summary>Terminal phase.</summary>
    public required JobPhase FinalPhase { get; init; }

    /// <summary>The configuration, retained so the run can be repeated or its failures retried.</summary>
    public JobConfiguration? Configuration { get; init; }
}

/// <summary>The kind of tenant modification recorded in the local audit history.</summary>
public enum RegistrationAuditAction
{
    /// <summary>An application registration was created.</summary>
    ApplicationCreated = 0,

    /// <summary>An application registration was modified.</summary>
    ApplicationUpdated,

    /// <summary>A service principal was created.</summary>
    ServicePrincipalCreated,

    /// <summary>An administrator consent flow was started.</summary>
    ConsentRequested,

    /// <summary>A consent verification was performed.</summary>
    ConsentVerified,

    /// <summary>An application registration was deleted.</summary>
    ApplicationDeleted,

    /// <summary>Local configuration was removed; nothing in the tenant changed.</summary>
    LocalConfigurationRemoved,
}

/// <summary>
/// A local, sanitized record of a tenant modification. Written for every material change so a
/// user can see exactly what this application did, and when.
/// </summary>
public sealed record RegistrationAuditEntry
{
    /// <summary>Stable identifier.</summary>
    public string EntryId { get; init; } = Guid.NewGuid().ToString("n");

    /// <summary>When the action occurred.</summary>
    public DateTimeOffset TimestampUtc { get; init; } = DateTimeOffset.UtcNow;

    /// <summary>What was done.</summary>
    public required RegistrationAuditAction Action { get; init; }

    /// <summary>Tenant affected.</summary>
    public required string TenantId { get; init; }

    /// <summary>Tenant display name, when known.</summary>
    public string? TenantDisplayName { get; init; }

    /// <summary>Application display name affected.</summary>
    public string? ApplicationDisplayName { get; init; }

    /// <summary>Client ID affected. Not a secret, but identifying, so exports redact it.</summary>
    public string? ClientId { get; init; }

    /// <summary>A privacy-conscious identifier of who performed the action.</summary>
    public required string PerformedBy { get; init; }

    /// <summary>The exact changes as they were presented to the user beforehand.</summary>
    public IReadOnlyList<string> Changes { get; init; } = [];

    /// <summary>True when the action succeeded.</summary>
    public required bool Succeeded { get; init; }

    /// <summary>Sanitized failure description, when the action failed.</summary>
    public string? FailureReason { get; init; }
}

/// <summary>
/// Declares exactly what a diagnostic bundle will contain. Shown to the user before export so
/// nothing leaves the machine without informed consent.
/// </summary>
public sealed record DiagnosticBundleMetadata
{
    /// <summary>When the bundle was produced.</summary>
    public DateTimeOffset GeneratedUtc { get; init; } = DateTimeOffset.UtcNow;

    /// <summary>Application version.</summary>
    public required string ApplicationVersion { get; init; }

    /// <summary>Operating system description.</summary>
    public required string Platform { get; init; }

    /// <summary>.NET runtime version.</summary>
    public required string RuntimeVersion { get; init; }

    /// <summary>Include file names. Off by default; requires explicit approval.</summary>
    public bool IncludeFileNames { get; init; }

    /// <summary>Include email addresses. Off by default; requires explicit approval.</summary>
    public bool IncludeEmailAddresses { get; init; }

    /// <summary>Include full tenant-specific URLs. Off by default; requires explicit approval.</summary>
    public bool IncludeFullUrls { get; init; }

    /// <summary>
    /// The categories that will always be included. Presented verbatim in the export dialog.
    /// </summary>
    public static IReadOnlyList<string> AlwaysIncluded =>
    [
        "Application version, platform and .NET runtime version",
        "Tenant connection status (connected or not) and consent status",
        "Registration status and service principal status",
        "Secure storage availability",
        "Sanitized recent errors: HTTP status, Graph error code and correlation ID",
        "Counts from the most recent job (created, reused, skipped, failed)",
        "Application settings, excluding any identifier",
    ];

    /// <summary>The categories that are never included, under any option.</summary>
    public static IReadOnlyList<string> NeverIncluded =>
    [
        "Access tokens, refresh tokens and authorization codes",
        "Authorization headers",
        "Client secrets, certificates and passwords",
        "Sharing links produced by any job",
        "File contents",
    ];

    /// <summary>The categories included only when the user explicitly approves them.</summary>
    public IReadOnlyList<string> OptionalIncluded()
    {
        var included = new List<string>();
        if (IncludeFileNames)
        {
            included.Add("File and folder names");
        }

        if (IncludeEmailAddresses)
        {
            included.Add("Email addresses and user principal names");
        }

        if (IncludeFullUrls)
        {
            included.Add("Full tenant-specific SharePoint and OneDrive URLs");
        }

        return included;
    }
}
