namespace SharePointLinkManifestBuilder.Core.Settings;

/// <summary>
/// Resolves where this application keeps local state on each platform.
/// <para>
/// The locations are exposed to the user on the privacy page. Nothing here ever holds a token:
/// credentials live in OS-native secure storage, which is a separate mechanism entirely
/// (see docs/adr/0008-token-storage.md).
/// </para>
/// </summary>
public sealed class ApplicationPaths
{
    /// <summary>The directory name used under the platform's application-data location.</summary>
    public const string ProductFolderName = "SharePointLinkManifestBuilder";

    /// <summary>Creates paths rooted at the platform application-data directory.</summary>
    public ApplicationPaths()
        : this(ResolveDefaultRoot())
    {
    }

    /// <summary>Creates paths rooted at an explicit directory. Used by tests.</summary>
    /// <param name="rootDirectory">The directory to hold all local state.</param>
    public ApplicationPaths(string rootDirectory)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(rootDirectory);
        RootDirectory = rootDirectory;
    }

    /// <summary>The directory holding all local state for this application.</summary>
    public string RootDirectory { get; }

    /// <summary>Non-secret application settings.</summary>
    public string SettingsFile => Path.Combine(RootDirectory, "settings.json");

    /// <summary>Tenant configuration. Contains a tenant ID and client ID, never a secret.</summary>
    public string TenantConfigurationFile => Path.Combine(RootDirectory, "tenant.json");

    /// <summary>Saved job profiles.</summary>
    public string ProfilesDirectory => Path.Combine(RootDirectory, "profiles");

    /// <summary>Local job history.</summary>
    public string HistoryDirectory => Path.Combine(RootDirectory, "history");

    /// <summary>Local audit trail of tenant modifications.</summary>
    public string AuditDirectory => Path.Combine(RootDirectory, "audit");

    /// <summary>Log files.</summary>
    public string LogDirectory => Path.Combine(RootDirectory, "logs");

    /// <summary>Directory holding the OS-protected MSAL cache, where the platform uses a file.</summary>
    public string TokenCacheDirectory => Path.Combine(RootDirectory, "cache");

    /// <summary>Scratch directory for exports before the user chooses a destination.</summary>
    public string ExportsDirectory => Path.Combine(RootDirectory, "exports");

    /// <summary>Creates every directory this application writes to.</summary>
    public void EnsureCreated()
    {
        foreach (var directory in new[]
        {
            RootDirectory, ProfilesDirectory, HistoryDirectory,
            AuditDirectory, LogDirectory, TokenCacheDirectory, ExportsDirectory,
        })
        {
            Directory.CreateDirectory(directory);
        }
    }

    /// <summary>
    /// A user-facing description of every local location, shown on the privacy page so the
    /// answer to "what does this store on my machine, and where" is discoverable in the UI.
    /// </summary>
    public IReadOnlyList<(string Description, string Path)> Describe() =>
    [
        ("Application settings (no identifiers)", SettingsFile),
        ("Microsoft 365 connection settings (tenant ID and client ID; no secrets)", TenantConfigurationFile),
        ("Saved job profiles", ProfilesDirectory),
        ("Job history", HistoryDirectory),
        ("Record of changes made to your tenant", AuditDirectory),
        ("Log files", LogDirectory),
        ("Sign-in cache, protected by the operating system", TokenCacheDirectory),
    ];

    private static string ResolveDefaultRoot()
    {
        // SpecialFolder.ApplicationData maps to %APPDATA% on Windows, ~/.config on Linux and
        // ~/.config on macOS under .NET. Create is requested so a fresh profile works.
        var baseDirectory = Environment.GetFolderPath(
            Environment.SpecialFolder.ApplicationData,
            Environment.SpecialFolderOption.Create);

        if (string.IsNullOrWhiteSpace(baseDirectory))
        {
            // A last resort for constrained environments where the platform reports no
            // application-data location at all.
            baseDirectory = Path.Combine(Path.GetTempPath(), ProductFolderName + "-fallback");
        }

        return Path.Combine(baseDirectory, ProductFolderName);
    }
}
