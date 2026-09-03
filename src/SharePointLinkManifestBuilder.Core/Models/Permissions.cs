namespace SharePointLinkManifestBuilder.Core.Models;

/// <summary>Whether a permission acts as the signed-in user or as the application itself.</summary>
public enum PermissionType
{
    /// <summary>The application acts as the signed-in user, bounded by that user's own access.</summary>
    Delegated = 0,

    /// <summary>
    /// The application acts as itself with no user. This product does not use application
    /// permissions for routine operation; the value exists so such permissions can be
    /// displayed and classified accurately if a tenant has granted them.
    /// </summary>
    Application = 1,
}

/// <summary>How risky a permission is, used to drive UI emphasis.</summary>
public enum PermissionBreadth
{
    /// <summary>Narrow, low-risk, usually consentable by a user.</summary>
    Minimal = 0,

    /// <summary>Tenant-wide read or scoped write. Normally requires administrator consent.</summary>
    Elevated = 1,

    /// <summary>Broad write across the tenant. Flagged prominently and always optional.</summary>
    Broad = 2,
}

/// <summary>
/// One Microsoft Graph permission this application may request, together with the
/// justification and least-privilege guidance shown to an administrator before consent.
/// </summary>
public sealed record PermissionRequirement
{
    /// <summary>The scope string sent to Microsoft Entra, for example <c>Sites.Read.All</c>.</summary>
    public required string Scope { get; init; }

    /// <summary>Delegated or application. This product requests delegated permissions only.</summary>
    public PermissionType Type { get; init; } = PermissionType.Delegated;

    /// <summary>Why this application needs the permission, in plain language.</summary>
    public required string Purpose { get; init; }

    /// <summary>What an administrator is practically allowing, in terms of data access.</summary>
    public required string DataAccessImpact { get; init; }

    /// <summary>True when tenant-wide administrator consent is normally required.</summary>
    public bool AdminConsentExpected { get; init; }

    /// <summary>Risk emphasis for the UI.</summary>
    public PermissionBreadth Breadth { get; init; } = PermissionBreadth.Minimal;

    /// <summary>A narrower permission that may suffice, when one exists.</summary>
    public string? LeastPrivilegeAlternative { get; init; }

    /// <summary>True when the application functions without this permission.</summary>
    public bool IsOptional { get; init; }

    /// <summary>The resource the scope belongs to. Microsoft Graph unless stated otherwise.</summary>
    public string ResourceAppId { get; init; } = AuthorityDefaults.MicrosoftGraphResourceAppId;

    /// <summary>True when the permission grants tenant-wide write and warrants a warning.</summary>
    public bool RequiresWarning => Breadth == PermissionBreadth.Broad;
}

/// <summary>Whether a required permission was actually issued.</summary>
public sealed record PermissionGrantState
{
    /// <summary>The permission that was checked.</summary>
    public required PermissionRequirement Requirement { get; init; }

    /// <summary>True when Microsoft Entra issued this scope in a real token.</summary>
    public required bool IsGranted { get; init; }

    /// <summary>How the grant was determined.</summary>
    public string GrantSource { get; init; } = "Token acquisition";
}

/// <summary>
/// The catalog of Microsoft Graph permissions this application can request.
/// <para>
/// Permission requirements were researched against the Microsoft Graph v1.0 permission tables
/// on Microsoft Learn rather than assumed. See docs/GRAPH-OPERATIONS.md for the per-operation
/// mapping and docs/adr/0009-least-privilege-scope-tiers.md for the tiering rationale.
/// </para>
/// </summary>
public static class GraphScopes
{
    /// <summary>Scopes MSAL always includes; they are never listed as consent requirements.</summary>
    public static readonly IReadOnlyList<string> Reserved = ["openid", "profile", "offline_access"];

    /// <summary>Read the signed-in user's own profile.</summary>
    public static readonly PermissionRequirement UserRead = new()
    {
        Scope = "User.Read",
        Purpose = "Identify the signed-in user and the tenant they are connected to.",
        DataAccessImpact = "Reads only the signed-in user's own basic profile (name, user principal name, ID).",
        AdminConsentExpected = false,
        Breadth = PermissionBreadth.Minimal,
    };

    /// <summary>Discover and read SharePoint sites and their document libraries.</summary>
    public static readonly PermissionRequirement SitesReadAll = new()
    {
        Scope = "Sites.Read.All",
        Purpose = "Find SharePoint sites, read site metadata, and list document libraries.",
        DataAccessImpact =
            "Allows reading SharePoint site and library information that the signed-in user can already "
            + "access. It does not grant access to sites the user cannot already open.",
        AdminConsentExpected = true,
        Breadth = PermissionBreadth.Elevated,
        LeastPrivilegeAlternative =
            "Sites.Selected restricts the application to individually assigned sites, but requires each "
            + "site to be assigned explicitly and behaves differently from normal delegated access.",
    };

    /// <summary>Read files across drives the user can access. Sufficient for dry runs.</summary>
    public static readonly PermissionRequirement FilesReadAll = new()
    {
        Scope = "Files.Read.All",
        Purpose = "Enumerate folders and files in the selected locations without changing anything.",
        DataAccessImpact =
            "Allows reading file and folder metadata in SharePoint and OneDrive locations the signed-in "
            + "user can already access. No content is modified.",
        AdminConsentExpected = true,
        Breadth = PermissionBreadth.Elevated,
    };

    /// <summary>Create sharing links and write manifests.</summary>
    public static readonly PermissionRequirement FilesReadWriteAll = new()
    {
        Scope = "Files.ReadWrite.All",
        Purpose = "Create sharing links for selected files and write manifest files back to SharePoint or OneDrive.",
        DataAccessImpact =
            "Allows creating sharing links and writing files in SharePoint and OneDrive locations the "
            + "signed-in user can already access. Sharing remains governed by tenant policy.",
        AdminConsentExpected = true,
        Breadth = PermissionBreadth.Elevated,
        LeastPrivilegeAlternative =
            "Files.ReadWrite covers only the signed-in user's own OneDrive. Choose it if you will never "
            + "target SharePoint libraries or another user's OneDrive.",
    };

    /// <summary>Read basic profiles of other users, for the User OneDrive picker.</summary>
    public static readonly PermissionRequirement UserReadBasicAll = new()
    {
        Scope = "User.ReadBasic.All",
        Purpose = "Search for a user by name so their OneDrive can be selected.",
        DataAccessImpact =
            "Allows reading basic profile information (display name, user principal name) for users in the "
            + "directory. It does not grant access to their files.",
        AdminConsentExpected = true,
        Breadth = PermissionBreadth.Elevated,
        IsOptional = true,
    };

    /// <summary>Broad SharePoint write. Optional, flagged, and never requested by default.</summary>
    public static readonly PermissionRequirement SitesReadWriteAll = new()
    {
        Scope = "Sites.ReadWrite.All",
        Purpose = "Write manifests to SharePoint libraries that reject writes made under Files.ReadWrite.All.",
        DataAccessImpact =
            "Allows reading and writing across SharePoint sites the signed-in user can access. This is a "
            + "broad write permission; enable it only if a library actually rejects the standard write.",
        AdminConsentExpected = true,
        Breadth = PermissionBreadth.Broad,
        IsOptional = true,
        LeastPrivilegeAlternative = "Files.ReadWrite.All is sufficient for most tenants and is preferred.",
    };

    /// <summary>Bootstrap: create an application registration. Least privileged for POST /applications.</summary>
    public static readonly PermissionRequirement AppRegistrationCreate = new()
    {
        Scope = "AppRegistration.Create",
        Purpose = "Create the tenant-specific application registration during automatic setup.",
        DataAccessImpact =
            "Allows creating a new application registration in your tenant. It does not allow reading, "
            + "modifying, or deleting existing registrations, and grants no access to files or directory data.",
        AdminConsentExpected = true,
        Breadth = PermissionBreadth.Elevated,
    };

    /// <summary>Bootstrap, opt-in: read/write application objects. Only for repair and deletion.</summary>
    public static readonly PermissionRequirement ApplicationReadWriteAll = new()
    {
        Scope = "Application.ReadWrite.All",
        Purpose =
            "Repair or replace an existing registration, create a service principal explicitly, or delete a "
            + "registration this application created.",
        DataAccessImpact =
            "Allows reading and writing ALL application registrations and service principals in the tenant. "
            + "This is a broad directory permission. It is requested only when you choose a repair, replace, "
            + "or delete action, and never for routine setup.",
        AdminConsentExpected = true,
        Breadth = PermissionBreadth.Broad,
        IsOptional = true,
        LeastPrivilegeAlternative =
            "AppRegistration.Create is sufficient to create a registration. This wider permission is needed "
            + "only to modify or delete one.",
    };

    /// <summary>Read-only browsing and dry runs. No write capability at all.</summary>
    public static IReadOnlyList<PermissionRequirement> DiscoveryTier =>
        [UserRead, SitesReadAll, FilesReadAll];

    /// <summary>The default operating set: discovery plus link creation and manifest writing.</summary>
    public static IReadOnlyList<PermissionRequirement> StandardTier =>
        [UserRead, SitesReadAll, FilesReadWriteAll];

    /// <summary>Bootstrap default: create-only. Sufficient because the POST body is complete.</summary>
    public static IReadOnlyList<PermissionRequirement> BootstrapCreateOnlyTier =>
        [UserRead, AppRegistrationCreate];

    /// <summary>Bootstrap opt-in: needed only for repair, replace, or delete.</summary>
    public static IReadOnlyList<PermissionRequirement> BootstrapManageTier =>
        [UserRead, ApplicationReadWriteAll];

    /// <summary>Every permission this application knows how to explain.</summary>
    public static IReadOnlyList<PermissionRequirement> All =>
    [
        UserRead, SitesReadAll, FilesReadAll, FilesReadWriteAll, UserReadBasicAll,
        SitesReadWriteAll, AppRegistrationCreate, ApplicationReadWriteAll,
    ];

    /// <summary>
    /// Builds the operating scope set for the given feature choices.
    /// </summary>
    /// <param name="readOnly">True to request no write capability at all.</param>
    /// <param name="includeUserOneDrivePicker">True when the User OneDrive source is enabled.</param>
    /// <param name="includeBroadSharePointWrite">True to add the broad, flagged SharePoint write scope.</param>
    public static IReadOnlyList<PermissionRequirement> BuildOperatingSet(
        bool readOnly = false,
        bool includeUserOneDrivePicker = false,
        bool includeBroadSharePointWrite = false)
    {
        var set = new List<PermissionRequirement>(readOnly ? DiscoveryTier : StandardTier);
        if (includeUserOneDrivePicker)
        {
            set.Add(UserReadBasicAll);
        }

        if (includeBroadSharePointWrite)
        {
            set.Add(SitesReadWriteAll);
        }

        return set;
    }

    /// <summary>
    /// Formats a scope list for a Microsoft Entra authorization or consent request. Reserved
    /// OIDC scopes are excluded because MSAL adds them, and duplicates are removed.
    /// </summary>
    public static string ToScopeParameter(IEnumerable<PermissionRequirement> permissions) =>
        string.Join(' ', permissions
            .Select(p => p.Scope)
            .Where(s => !Reserved.Contains(s, StringComparer.OrdinalIgnoreCase))
            .Distinct(StringComparer.OrdinalIgnoreCase));

    /// <summary>
    /// Converts scope names into the fully-qualified form Microsoft Entra expects in an
    /// admin-consent request, for example <c>https://graph.microsoft.com/Sites.Read.All</c>.
    /// </summary>
    public static string ToQualifiedScopeParameter(
        IEnumerable<PermissionRequirement> permissions,
        string graphEndpoint)
    {
        var resource = graphEndpoint.Replace("/v1.0", string.Empty, StringComparison.OrdinalIgnoreCase).TrimEnd('/');
        return string.Join(' ', permissions
            .Select(p => p.Scope)
            .Where(s => !Reserved.Contains(s, StringComparer.OrdinalIgnoreCase))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .Select(s => $"{resource}/{s}"));
    }

    /// <summary>Finds a known permission by scope name, for display of an unfamiliar grant.</summary>
    public static PermissionRequirement? Find(string scope) =>
        All.FirstOrDefault(p => string.Equals(p.Scope, scope, StringComparison.OrdinalIgnoreCase));
}
