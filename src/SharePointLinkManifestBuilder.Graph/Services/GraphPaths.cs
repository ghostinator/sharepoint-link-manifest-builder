namespace SharePointLinkManifestBuilder.Graph.Services;

/// <summary>
/// Builds Microsoft Graph request paths.
/// <para>
/// Centralized so escaping is applied consistently. SharePoint names routinely contain spaces,
/// ampersands, plus signs and non-ASCII characters; an unescaped name produces either a 400 or,
/// worse, a request against a different item than the caller intended.
/// </para>
/// </summary>
public static class GraphPaths
{
    /// <summary>
    /// The <c>$select</c> used when enumerating children: enough to classify, filter and
    /// identify an item without over-fetching tenant data.
    /// </summary>
    public const string DriveItemSelect =
        "id,name,size,webUrl,eTag,cTag,lastModifiedDateTime,createdDateTime,"
        + "folder,file,package,remoteItem,shared,parentReference";

    /// <summary>Page size requested when listing children. Graph may return fewer.</summary>
    public const int ChildrenPageSize = 200;

    /// <summary>Escapes one path segment, leaving the rest of the URL structure intact.</summary>
    public static string EscapeSegment(string segment) => Uri.EscapeDataString(segment);

    /// <summary>
    /// Escapes a multi-segment relative path, preserving the forward slashes that separate
    /// segments while escaping the characters inside each one.
    /// </summary>
    public static string EscapePath(string relativePath)
    {
        if (string.IsNullOrWhiteSpace(relativePath))
        {
            return string.Empty;
        }

        return string.Join(
            '/',
            relativePath
                .Replace('\\', '/')
                .Trim('/')
                .Split('/', StringSplitOptions.RemoveEmptyEntries)
                .Select(Uri.EscapeDataString));
    }

    /// <summary>Search for sites the signed-in user can see.</summary>
    public static string SearchSites(string query) =>
        $"/sites?search={Uri.EscapeDataString(query)}&$top=50";

    /// <summary>Resolve a site from a hostname and server-relative site path.</summary>
    public static string SiteByPath(string hostname, string sitePath) =>
        string.IsNullOrEmpty(sitePath)
            ? $"/sites/{EscapeSegment(hostname)}"
            : $"/sites/{EscapeSegment(hostname)}:{EscapePathPreservingLeadingSlash(sitePath)}";

    /// <summary>The tenant root site.</summary>
    public static string RootSite() => "/sites/root";

    /// <summary>A site by composite ID.</summary>
    public static string Site(string siteId) => $"/sites/{siteId}";

    /// <summary>Sites the signed-in user follows.</summary>
    public static string FollowedSites() => "/me/followedSites";

    /// <summary>Document libraries of a site.</summary>
    public static string SiteDrives(string siteId) =>
        $"/sites/{siteId}/drives?$select=id,name,driveType,webUrl";

    /// <summary>The signed-in user's OneDrive.</summary>
    public static string MyDrive() => "/me/drive?$select=id,name,driveType,webUrl,owner";

    /// <summary>Another user's OneDrive.</summary>
    public static string UserDrive(string userId) =>
        $"/users/{EscapeSegment(userId)}/drive?$select=id,name,driveType,webUrl,owner";

    /// <summary>A drive by ID.</summary>
    public static string Drive(string driveId) =>
        $"/drives/{EscapeSegment(driveId)}?$select=id,name,driveType,webUrl,owner";

    /// <summary>A drive's root folder.</summary>
    public static string DriveRoot(string driveId) =>
        $"/drives/{EscapeSegment(driveId)}/root?$select={DriveItemSelect}";

    /// <summary>Immediate children of a folder, with pagination and a bounded projection.</summary>
    public static string Children(string driveId, string itemId) =>
        $"/drives/{EscapeSegment(driveId)}/items/{EscapeSegment(itemId)}/children"
        + $"?$top={ChildrenPageSize}&$select={DriveItemSelect}";

    /// <summary>A single item by ID.</summary>
    public static string Item(string driveId, string itemId) =>
        $"/drives/{EscapeSegment(driveId)}/items/{EscapeSegment(itemId)}?$select={DriveItemSelect}";

    /// <summary>A folder addressed by its path relative to the drive root.</summary>
    public static string ItemByPath(string driveId, string relativePath) =>
        string.IsNullOrWhiteSpace(relativePath)
            ? DriveRoot(driveId)
            : $"/drives/{EscapeSegment(driveId)}/root:/{EscapePath(relativePath)}?$select={DriveItemSelect}";

    /// <summary>A child of a folder, addressed by name.</summary>
    public static string ChildByName(string driveId, string parentItemId, string name) =>
        $"/drives/{EscapeSegment(driveId)}/items/{EscapeSegment(parentItemId)}:/{EscapeSegment(name)}"
        + $"?$select={DriveItemSelect}";

    /// <summary>Content of a child addressed by name.</summary>
    public static string ChildContentByName(string driveId, string parentItemId, string name) =>
        $"/drives/{EscapeSegment(driveId)}/items/{EscapeSegment(parentItemId)}:/{EscapeSegment(name)}:/content";

    /// <summary>Create an upload session for a large manifest.</summary>
    public static string CreateUploadSession(string driveId, string parentItemId, string name) =>
        $"/drives/{EscapeSegment(driveId)}/items/{EscapeSegment(parentItemId)}:/{EscapeSegment(name)}:"
        + "/createUploadSession";

    /// <summary>Create or obtain a sharing link.</summary>
    public static string CreateLink(string driveId, string itemId) =>
        $"/drives/{EscapeSegment(driveId)}/items/{EscapeSegment(itemId)}/createLink";

    /// <summary>Grant named recipients access.</summary>
    public static string Invite(string driveId, string itemId) =>
        $"/drives/{EscapeSegment(driveId)}/items/{EscapeSegment(itemId)}/invite";

    /// <summary>List existing permissions on an item.</summary>
    public static string Permissions(string driveId, string itemId) =>
        $"/drives/{EscapeSegment(driveId)}/items/{EscapeSegment(itemId)}/permissions";

    /// <summary>Resolve a sharing URL through the shares endpoint.</summary>
    public static string SharedItem(string shareToken) =>
        $"/shares/{shareToken}/driveItem?$select={DriveItemSelect}";

    /// <summary>Search the directory for users.</summary>
    public static string SearchUsers(string query)
    {
        var escaped = Uri.EscapeDataString(query.Replace("'", "''", StringComparison.Ordinal));

        // startswith on the two display fields is used rather than $search, because $search on
        // /users requires the ConsistencyLevel=eventual header and is not universally enabled.
        return "/users?$select=id,displayName,userPrincipalName,jobTitle&$top=25&$filter="
            + $"startswith(displayName,'{escaped}') or startswith(userPrincipalName,'{escaped}')";
    }

    /// <summary>A user by ID or user principal name.</summary>
    public static string User(string userIdOrUpn) =>
        $"/users/{EscapeSegment(userIdOrUpn)}?$select=id,displayName,userPrincipalName,jobTitle";

    /// <summary>The signed-in user's profile.</summary>
    public static string Me() => "/me?$select=id,displayName,userPrincipalName";

    /// <summary>The tenant's organization profile.</summary>
    public static string Organization() => "/organization?$select=id,displayName,verifiedDomains";

    /// <summary>Create an application registration.</summary>
    public static string Applications() => "/applications";

    /// <summary>An application addressed by client ID rather than object ID.</summary>
    public static string ApplicationByAppId(string appId) =>
        $"/applications(appId='{Uri.EscapeDataString(appId)}')";

    /// <summary>An application addressed by object ID.</summary>
    public static string ApplicationByObjectId(string objectId) =>
        $"/applications/{EscapeSegment(objectId)}";

    /// <summary>Create a service principal.</summary>
    public static string ServicePrincipals() => "/servicePrincipals";

    /// <summary>A service principal addressed by client ID.</summary>
    public static string ServicePrincipalByAppId(string appId) =>
        $"/servicePrincipals(appId='{Uri.EscapeDataString(appId)}')";

    /// <summary>Delegated permission grants for a service principal.</summary>
    public static string OAuth2PermissionGrants(string clientServicePrincipalId) =>
        $"/oauth2PermissionGrants?$filter=clientId eq '{Uri.EscapeDataString(clientServicePrincipalId)}'";

    private static string EscapePathPreservingLeadingSlash(string path)
    {
        var escaped = EscapePath(path);
        return escaped.Length == 0 ? string.Empty : "/" + escaped;
    }
}
