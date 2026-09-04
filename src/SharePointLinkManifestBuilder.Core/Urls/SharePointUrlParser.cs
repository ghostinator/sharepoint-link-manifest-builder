using System.Diagnostics.CodeAnalysis;

namespace SharePointLinkManifestBuilder.Core.Urls;

/// <summary>
/// Parses the SharePoint and OneDrive URLs a user can realistically paste, including the
/// <c>_layouts</c> forms produced by the SharePoint "Copy link" and address-bar experiences.
/// <para>
/// Pure and side-effect free. The URL is validated and decomposed locally and is never
/// requested; resolution goes through Microsoft Graph.
/// </para>
/// </summary>
public static class SharePointUrlParser
{
    /// <summary>
    /// Host suffixes accepted as SharePoint or OneDrive, covering the public cloud and the
    /// sovereign clouds. Defence in depth: resolution goes through Graph regardless.
    /// </summary>
    public static readonly IReadOnlyList<string> AllowedHostSuffixes =
    [
        ".sharepoint.com",
        ".sharepoint.us",
        ".sharepoint-mil.us",
        ".sharepoint.cn",
        ".sharepoint.de",
    ];

    private static readonly string[] SitePrefixes = ["/sites/", "/teams/", "/personal/"];

    /// <summary>
    /// Segments that mark the start of SharePoint's own application pages. Everything from here
    /// on is chrome, not a document path.
    /// </summary>
    private static readonly string[] LayoutMarkers = ["/_layouts/", "/_api/", "/_vti_bin/", "/forms/"];

    /// <summary>Query parameters that carry the real server-relative path in <c>_layouts</c> URLs.</summary>
    private static readonly string[] PathQueryKeys = ["id", "RootFolder", "FolderCTID", "parent"];

    /// <summary>Parses a pasted URL. Never throws for malformed input.</summary>
    /// <param name="url">The URL as the user supplied it.</param>
    public static ParsedResourceUrl Parse(string? url)
    {
        var original = (url ?? string.Empty).Trim();

        if (string.IsNullOrEmpty(original))
        {
            return ParsedResourceUrl.Invalid(original, "No URL was supplied.");
        }

        if (!Uri.TryCreate(original, UriKind.Absolute, out var uri))
        {
            return ParsedResourceUrl.Invalid(original, "The value is not a complete URL. Include https:// at the start.");
        }

        if (!string.Equals(uri.Scheme, Uri.UriSchemeHttps, StringComparison.OrdinalIgnoreCase))
        {
            return ParsedResourceUrl.Invalid(
                original,
                $"Only https URLs are accepted; this URL uses '{uri.Scheme}'.");
        }

        if (!IsAllowedHost(uri.Host))
        {
            return ParsedResourceUrl.Invalid(
                original,
                $"'{uri.Host}' is not a recognised SharePoint or OneDrive host name.");
        }

        var host = uri.Host;
        var path = Uri.UnescapeDataString(uri.AbsolutePath).TrimEnd('/');

        // A sharing link (/:f:/s/..., /:w:/r/..., /:x:/g/...) is opaque and must go through
        // the Graph /shares endpoint. Detect it before any structural parsing.
        if (IsSharingLink(path))
        {
            return new ParsedResourceUrl
            {
                Kind = ResourceUrlKind.SharingLink,
                OriginalUrl = original,
                Hostname = host,
            };
        }

        // SharePoint's own pages carry the real folder path in a query parameter rather than
        // in the path, e.g. /sites/M/_layouts/15/onedrive.aspx?id=%2Fsites%2FM%2FShared%20Documents%2FX
        var layoutIndex = IndexOfLayoutMarker(path);
        if (layoutIndex >= 0)
        {
            var fromQuery = ExtractPathFromQuery(uri);
            if (fromQuery is not null)
            {
                return BuildFromServerRelativePath(original, host, fromQuery);
            }

            // No usable query parameter: fall back to the site portion that precedes /_layouts/.
            path = path[..layoutIndex];
        }

        if (path.Length == 0)
        {
            return new ParsedResourceUrl
            {
                Kind = ResourceUrlKind.RootSite,
                OriginalUrl = original,
                Hostname = host,
                SitePath = string.Empty,
            };
        }

        return BuildFromServerRelativePath(original, host, path);
    }

    /// <summary>True when the host is a recognised SharePoint or OneDrive host.</summary>
    public static bool IsAllowedHost([NotNullWhen(true)] string? host) =>
        !string.IsNullOrWhiteSpace(host)
        && AllowedHostSuffixes.Any(suffix => host.EndsWith(suffix, StringComparison.OrdinalIgnoreCase));

    /// <summary>
    /// True when the path is a SharePoint sharing link. These use a short marker segment such
    /// as <c>:f:</c> (folder), <c>:w:</c> (Word), <c>:x:</c> (Excel) or <c>:b:</c> (PDF).
    /// </summary>
    public static bool IsSharingLink(string path)
    {
        var segments = path.Split('/', StringSplitOptions.RemoveEmptyEntries);
        return segments.Length > 0
            && segments[0].Length >= 3
            && segments[0].StartsWith(':')
            && segments[0].EndsWith(':');
    }

    private static ParsedResourceUrl BuildFromServerRelativePath(string original, string host, string path)
    {
        if (!path.StartsWith('/'))
        {
            path = "/" + path;
        }

        path = path.TrimEnd('/');

        // A bare "/sites", "/teams" or "/personal" names no site. Trimming the trailing slash
        // makes it look like an ordinary root-relative path, so reject it explicitly.
        if (SitePrefixes.Any(p => string.Equals(path, p.TrimEnd('/'), StringComparison.OrdinalIgnoreCase)))
        {
            return ParsedResourceUrl.Invalid(
                original, $"The URL is missing a site name after '{path.TrimStart('/')}'.");
        }

        var prefix = SitePrefixes.FirstOrDefault(p =>
            path.StartsWith(p, StringComparison.OrdinalIgnoreCase));

        if (prefix is null)
        {
            // No /sites/, /teams/ or /personal/ prefix: the path lives directly under the root
            // site, e.g. https://contoso.sharepoint.com/Shared Documents/Report.docx
            return new ParsedResourceUrl
            {
                Kind = path.Length == 0 ? ResourceUrlKind.RootSite : ResourceUrlKind.DocumentPath,
                OriginalUrl = original,
                Hostname = host,
                SitePath = string.Empty,
                ServerRelativeItemPath = path.Length == 0 ? null : path,
                SiteRelativeItemPath = path.Length == 0 ? null : path.TrimStart('/'),
            };
        }

        // /sites/Marketing/Shared Documents/Reports
        //  ^prefix ^name    ^remainder
        var afterPrefix = path[prefix.Length..];
        var slash = afterPrefix.IndexOf('/', StringComparison.Ordinal);

        var siteName = slash < 0 ? afterPrefix : afterPrefix[..slash];
        var remainder = slash < 0 ? string.Empty : afterPrefix[(slash + 1)..].Trim('/');

        if (siteName.Length == 0)
        {
            return ParsedResourceUrl.Invalid(original, "The URL is missing a site name after '" + prefix.Trim('/') + "'.");
        }

        var sitePath = prefix + siteName;
        var isPersonal = prefix.Equals("/personal/", StringComparison.OrdinalIgnoreCase);

        if (remainder.Length == 0)
        {
            return new ParsedResourceUrl
            {
                Kind = isPersonal ? ResourceUrlKind.PersonalSite : ResourceUrlKind.Site,
                OriginalUrl = original,
                Hostname = host,
                SitePath = sitePath,
                PersonalSiteSegment = isPersonal ? siteName : null,
            };
        }

        return new ParsedResourceUrl
        {
            Kind = ResourceUrlKind.DocumentPath,
            OriginalUrl = original,
            Hostname = host,
            SitePath = sitePath,
            ServerRelativeItemPath = $"{sitePath}/{remainder}",
            SiteRelativeItemPath = remainder,
            PersonalSiteSegment = isPersonal ? siteName : null,
        };
    }

    private static int IndexOfLayoutMarker(string path)
    {
        var best = -1;
        foreach (var marker in LayoutMarkers)
        {
            var index = path.IndexOf(marker, StringComparison.OrdinalIgnoreCase);
            if (index >= 0 && (best < 0 || index < best))
            {
                best = index;
            }
        }

        return best;
    }

    private static string? ExtractPathFromQuery(Uri uri)
    {
        if (string.IsNullOrEmpty(uri.Query))
        {
            return null;
        }

        foreach (var pair in uri.Query.TrimStart('?').Split('&', StringSplitOptions.RemoveEmptyEntries))
        {
            var separator = pair.IndexOf('=', StringComparison.Ordinal);
            if (separator <= 0)
            {
                continue;
            }

            var key = Uri.UnescapeDataString(pair[..separator]);
            if (!PathQueryKeys.Contains(key, StringComparer.OrdinalIgnoreCase))
            {
                continue;
            }

            var value = Uri.UnescapeDataString(pair[(separator + 1)..]);
            if (value.StartsWith('/'))
            {
                return value.TrimEnd('/');
            }
        }

        return null;
    }
}
