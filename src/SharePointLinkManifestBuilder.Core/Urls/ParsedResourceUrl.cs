namespace SharePointLinkManifestBuilder.Core.Urls;

/// <summary>What a pasted SharePoint or OneDrive URL refers to.</summary>
public enum ResourceUrlKind
{
    /// <summary>The URL could not be understood.</summary>
    Invalid = 0,

    /// <summary>The tenant root site, for example <c>https://contoso.sharepoint.com</c>.</summary>
    RootSite,

    /// <summary>A site, for example <c>/sites/Marketing</c> or <c>/teams/Engineering</c>.</summary>
    Site,

    /// <summary>A personal OneDrive site, for example <c>/personal/jane_contoso_com</c>.</summary>
    PersonalSite,

    /// <summary>A path inside a document library or OneDrive.</summary>
    DocumentPath,

    /// <summary>
    /// A sharing link such as <c>/:f:/s/Marketing/Ex4...</c>, which must be resolved through
    /// the Graph <c>/shares</c> endpoint rather than parsed.
    /// </summary>
    SharingLink,
}

/// <summary>
/// The structured result of parsing a pasted URL.
/// <para>
/// Parsing is entirely local. The URL is never fetched: resolution happens exclusively through
/// Microsoft Graph, so a hostile URL cannot induce an arbitrary outbound request.
/// </para>
/// </summary>
public sealed record ParsedResourceUrl
{
    /// <summary>What the URL refers to.</summary>
    public required ResourceUrlKind Kind { get; init; }

    /// <summary>The URL as supplied, trimmed.</summary>
    public required string OriginalUrl { get; init; }

    /// <summary>Host name, for example <c>contoso.sharepoint.com</c>.</summary>
    public string? Hostname { get; init; }

    /// <summary>
    /// Server-relative site path with a leading slash and no trailing slash, for example
    /// <c>/sites/Marketing</c>. Empty for the tenant root site.
    /// </summary>
    public string SitePath { get; init; } = string.Empty;

    /// <summary>
    /// Decoded server-relative path of the item, for example
    /// <c>/sites/Marketing/Shared Documents/Reports</c>. Null when the URL names only a site.
    /// </summary>
    public string? ServerRelativeItemPath { get; init; }

    /// <summary>
    /// Item path relative to the site, for example <c>Shared Documents/Reports</c>. This still
    /// includes the library segment, because only Graph can say which library that is.
    /// </summary>
    public string? SiteRelativeItemPath { get; init; }

    /// <summary>True when the URL points into a personal OneDrive site.</summary>
    public bool IsPersonalSite => Kind == ResourceUrlKind.PersonalSite
        || SitePath.StartsWith("/personal/", StringComparison.OrdinalIgnoreCase);

    /// <summary>
    /// The encoded user segment of a personal site, for example <c>jane_contoso_com</c>.
    /// Not a user principal name; Graph must resolve it.
    /// </summary>
    public string? PersonalSiteSegment { get; init; }

    /// <summary>Why parsing failed, when it did.</summary>
    public string? FailureReason { get; init; }

    /// <summary>True when the URL was understood.</summary>
    public bool IsValid => Kind != ResourceUrlKind.Invalid;

    /// <summary>
    /// The Graph path used to resolve the site, for example
    /// <c>/sites/contoso.sharepoint.com:/sites/Marketing</c>. Null when not applicable.
    /// </summary>
    public string? ToGraphSitePath() => Kind switch
    {
        ResourceUrlKind.RootSite => $"/sites/{Hostname}",
        ResourceUrlKind.Site or ResourceUrlKind.PersonalSite or ResourceUrlKind.DocumentPath
            when !string.IsNullOrEmpty(SitePath) => $"/sites/{Hostname}:{SitePath}",
        ResourceUrlKind.DocumentPath => $"/sites/{Hostname}",
        _ => null,
    };

    /// <summary>Creates a failed parse result.</summary>
    public static ParsedResourceUrl Invalid(string originalUrl, string reason) => new()
    {
        Kind = ResourceUrlKind.Invalid,
        OriginalUrl = originalUrl,
        FailureReason = reason,
    };
}
