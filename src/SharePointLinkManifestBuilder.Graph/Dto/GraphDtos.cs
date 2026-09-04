using System.Text.Json.Serialization;

namespace SharePointLinkManifestBuilder.Graph.Dto;

/// <summary>
/// Minimal data-transfer types for the Microsoft Graph responses this application consumes.
/// <para>
/// Only fields the application actually uses are modelled. Keeping these small is deliberate:
/// it limits the surface that must track Graph changes, and it means no unused tenant data is
/// ever deserialized into memory. See docs/adr/0003-raw-graph-http-over-sdk.md.
/// </para>
/// </summary>
internal static class GraphDtoDocumentation
{
    // Marker type for documentation only.
}

/// <summary>A Microsoft Graph <c>user</c>.</summary>
public sealed record GraphUserDto
{
    /// <summary>Entra object ID.</summary>
    public string? Id { get; init; }

    /// <summary>Display name.</summary>
    public string? DisplayName { get; init; }

    /// <summary>User principal name.</summary>
    public string? UserPrincipalName { get; init; }

    /// <summary>Job title, when the directory exposes it.</summary>
    public string? JobTitle { get; init; }

    /// <summary>Primary mail address, when present.</summary>
    public string? Mail { get; init; }
}

/// <summary>A Microsoft Graph <c>organization</c>.</summary>
public sealed record GraphOrganizationDto
{
    /// <summary>Tenant ID.</summary>
    public string? Id { get; init; }

    /// <summary>Tenant display name.</summary>
    public string? DisplayName { get; init; }

    /// <summary>Verified domains, used to show a recognisable tenant name.</summary>
    public IReadOnlyList<GraphVerifiedDomainDto>? VerifiedDomains { get; init; }
}

/// <summary>A verified domain on a tenant.</summary>
public sealed record GraphVerifiedDomainDto
{
    /// <summary>The domain name.</summary>
    public string? Name { get; init; }

    /// <summary>True when this is the tenant's default domain.</summary>
    public bool? IsDefault { get; init; }
}

/// <summary>A Microsoft Graph <c>site</c>.</summary>
public sealed record GraphSiteDto
{
    /// <summary>Composite site ID (<c>hostname,siteCollectionId,siteId</c>).</summary>
    public string? Id { get; init; }

    /// <summary>Site name (the URL segment).</summary>
    public string? Name { get; init; }

    /// <summary>Site title.</summary>
    public string? DisplayName { get; init; }

    /// <summary>Absolute site URL.</summary>
    public string? WebUrl { get; init; }

    /// <summary>Description, when set.</summary>
    public string? Description { get; init; }

    /// <summary>Creation timestamp.</summary>
    public DateTimeOffset? CreatedDateTime { get; init; }

    /// <summary>Root marker; present only on the tenant root site.</summary>
    public object? Root { get; init; }
}

/// <summary>A Microsoft Graph <c>drive</c>.</summary>
public sealed record GraphDriveDto
{
    /// <summary>Drive ID.</summary>
    public string? Id { get; init; }

    /// <summary>Friendly library or drive name.</summary>
    public string? Name { get; init; }

    /// <summary>One of <c>documentLibrary</c>, <c>business</c> or <c>personal</c>.</summary>
    public string? DriveType { get; init; }

    /// <summary>Absolute URL.</summary>
    public string? WebUrl { get; init; }

    /// <summary>Owner information, present for personal drives.</summary>
    public GraphIdentitySetDto? Owner { get; init; }
}

/// <summary>A Graph identity set.</summary>
public sealed record GraphIdentitySetDto
{
    /// <summary>The user identity, when present.</summary>
    public GraphIdentityDto? User { get; init; }

    /// <summary>The application identity, when present.</summary>
    public GraphIdentityDto? Application { get; init; }
}

/// <summary>A Graph identity.</summary>
public sealed record GraphIdentityDto
{
    /// <summary>Identity ID.</summary>
    public string? Id { get; init; }

    /// <summary>Identity display name.</summary>
    public string? DisplayName { get; init; }
}

/// <summary>A Microsoft Graph <c>driveItem</c>.</summary>
public sealed record GraphDriveItemDto
{
    /// <summary>Item ID.</summary>
    public string? Id { get; init; }

    /// <summary>Item name.</summary>
    public string? Name { get; init; }

    /// <summary>Size in bytes.</summary>
    public long? Size { get; init; }

    /// <summary>Absolute URL.</summary>
    public string? WebUrl { get; init; }

    /// <summary>Item ETag.</summary>
    public string? ETag { get; init; }

    /// <summary>Content tag, which changes only when content changes.</summary>
    public string? CTag { get; init; }

    /// <summary>Last modification timestamp.</summary>
    public DateTimeOffset? LastModifiedDateTime { get; init; }

    /// <summary>Creation timestamp.</summary>
    public DateTimeOffset? CreatedDateTime { get; init; }

    /// <summary>Present when the item is a folder.</summary>
    public GraphFolderFacetDto? Folder { get; init; }

    /// <summary>Present when the item is a file.</summary>
    public GraphFileFacetDto? File { get; init; }

    /// <summary>Present when the item is a package, such as a OneNote notebook.</summary>
    public GraphPackageFacetDto? Package { get; init; }

    /// <summary>Present when the item is a shortcut to something stored elsewhere.</summary>
    public GraphRemoteItemDto? RemoteItem { get; init; }

    /// <summary>Location of the item's parent.</summary>
    public GraphItemReferenceDto? ParentReference { get; init; }

    /// <summary>Sharing information, when Graph reports the item as shared.</summary>
    public GraphSharedFacetDto? Shared { get; init; }
}

/// <summary>Marks a drive item as a folder.</summary>
public sealed record GraphFolderFacetDto
{
    /// <summary>Number of direct children.</summary>
    public int? ChildCount { get; init; }
}

/// <summary>Marks a drive item as a file.</summary>
public sealed record GraphFileFacetDto
{
    /// <summary>MIME type.</summary>
    public string? MimeType { get; init; }
}

/// <summary>Marks a drive item as a package.</summary>
public sealed record GraphPackageFacetDto
{
    /// <summary>Package type, for example <c>oneNote</c>.</summary>
    public string? Type { get; init; }
}

/// <summary>Marks a drive item as a shortcut to a remote item.</summary>
public sealed record GraphRemoteItemDto
{
    /// <summary>Remote item ID.</summary>
    public string? Id { get; init; }

    /// <summary>Remote item name.</summary>
    public string? Name { get; init; }
}

/// <summary>Sharing information on a drive item.</summary>
public sealed record GraphSharedFacetDto
{
    /// <summary>Sharing scope, for example <c>users</c> or <c>anonymous</c>.</summary>
    public string? Scope { get; init; }
}

/// <summary>A reference to a drive item's parent.</summary>
public sealed record GraphItemReferenceDto
{
    /// <summary>Owning drive ID.</summary>
    public string? DriveId { get; init; }

    /// <summary>Parent item ID.</summary>
    public string? Id { get; init; }

    /// <summary>Parent path, in the form <c>/drive/root:/Folder/Sub</c>.</summary>
    public string? Path { get; init; }

    /// <summary>Owning site ID, for SharePoint items.</summary>
    public string? SiteId { get; init; }
}

/// <summary>A Microsoft Graph <c>permission</c>, returned by createLink and invite.</summary>
public sealed record GraphPermissionDto
{
    /// <summary>Permission ID.</summary>
    public string? Id { get; init; }

    /// <summary>Roles granted, for example <c>read</c> or <c>write</c>.</summary>
    public IReadOnlyList<string>? Roles { get; init; }

    /// <summary>The sharing link, when this permission represents one.</summary>
    public GraphSharingLinkDto? Link { get; init; }

    /// <summary>Expiry, when set.</summary>
    public DateTimeOffset? ExpirationDateTime { get; init; }

    /// <summary>True when the link is password protected.</summary>
    public bool? HasPassword { get; init; }

    /// <summary>Invitation details, present on permissions created by the invite action.</summary>
    public GraphSharingInvitationDto? Invitation { get; init; }

    /// <summary>Who the permission was granted to.</summary>
    public GraphIdentitySetDto? GrantedToV2 { get; init; }

    /// <summary>
    /// Per-entry error, present in a <c>207 Multi-Status</c> invite response where some
    /// recipients succeeded and others failed.
    /// </summary>
    public GraphInnerErrorDto? Error { get; init; }
}

/// <summary>The sharing-link portion of a permission.</summary>
public sealed record GraphSharingLinkDto
{
    /// <summary>Link type: <c>view</c>, <c>edit</c> or <c>embed</c>.</summary>
    public string? Type { get; init; }

    /// <summary>Link scope: <c>anonymous</c>, <c>organization</c> or <c>users</c>.</summary>
    public string? Scope { get; init; }

    /// <summary>The link URL.</summary>
    public string? WebUrl { get; init; }

    /// <summary>Embed HTML, for embed links only.</summary>
    public string? WebHtml { get; init; }

    /// <summary>True when the link prevents downloading.</summary>
    public bool? PreventsDownload { get; init; }
}

/// <summary>Invitation details on a permission.</summary>
public sealed record GraphSharingInvitationDto
{
    /// <summary>The invited email address.</summary>
    public string? Email { get; init; }

    /// <summary>True when the recipient must sign in.</summary>
    public bool? SignInRequired { get; init; }
}

/// <summary>An error nested inside a multi-status response entry.</summary>
public sealed record GraphInnerErrorDto
{
    /// <summary>Graph error code.</summary>
    public string? Code { get; init; }

    /// <summary>Human-readable message.</summary>
    public string? Message { get; init; }
}

/// <summary>The response to createUploadSession.</summary>
public sealed record GraphUploadSessionDto
{
    /// <summary>Pre-authorized upload URL. Carries its own credentials in the query string.</summary>
    public string? UploadUrl { get; init; }

    /// <summary>When the session expires.</summary>
    public DateTimeOffset? ExpirationDateTime { get; init; }

    /// <summary>Byte ranges the service still expects.</summary>
    public IReadOnlyList<string>? NextExpectedRanges { get; init; }
}

/// <summary>A Microsoft Entra <c>application</c>.</summary>
public sealed record GraphApplicationDto
{
    /// <summary>Object ID.</summary>
    public string? Id { get; init; }

    /// <summary>Application (client) ID.</summary>
    public string? AppId { get; init; }

    /// <summary>Display name.</summary>
    public string? DisplayName { get; init; }

    /// <summary>Supported account types.</summary>
    public string? SignInAudience { get; init; }

    /// <summary>True when the app is treated as a public client.</summary>
    public bool? IsFallbackPublicClient { get; init; }

    /// <summary>Public client configuration, including redirect URIs.</summary>
    public GraphPublicClientDto? PublicClient { get; init; }

    /// <summary>Requested resource access, that is, the configured API permissions.</summary>
    public IReadOnlyList<GraphRequiredResourceAccessDto>? RequiredResourceAccess { get; init; }

    /// <summary>
    /// Password credentials. Read only so the UI can warn if a secret exists; this application
    /// never creates one.
    /// </summary>
    public IReadOnlyList<GraphPasswordCredentialDto>? PasswordCredentials { get; init; }
}

/// <summary>Public client platform configuration.</summary>
public sealed record GraphPublicClientDto
{
    /// <summary>Registered redirect URIs.</summary>
    public IReadOnlyList<string>? RedirectUris { get; init; }
}

/// <summary>A resource and the permissions requested against it.</summary>
public sealed record GraphRequiredResourceAccessDto
{
    /// <summary>The resource application ID, for example Microsoft Graph.</summary>
    public string? ResourceAppId { get; init; }

    /// <summary>The individual permissions requested.</summary>
    public IReadOnlyList<GraphResourceAccessDto>? ResourceAccess { get; init; }
}

/// <summary>One requested permission.</summary>
public sealed record GraphResourceAccessDto
{
    /// <summary>The permission's GUID.</summary>
    public string? Id { get; init; }

    /// <summary><c>Scope</c> for delegated, <c>Role</c> for application permissions.</summary>
    public string? Type { get; init; }
}

/// <summary>A password credential on an application. This application never creates one.</summary>
public sealed record GraphPasswordCredentialDto
{
    /// <summary>Credential key ID.</summary>
    public string? KeyId { get; init; }

    /// <summary>Credential display name.</summary>
    public string? DisplayName { get; init; }

    /// <summary>Expiry.</summary>
    public DateTimeOffset? EndDateTime { get; init; }
}

/// <summary>A Microsoft Entra <c>servicePrincipal</c>.</summary>
public sealed record GraphServicePrincipalDto
{
    /// <summary>Object ID.</summary>
    public string? Id { get; init; }

    /// <summary>Application (client) ID.</summary>
    public string? AppId { get; init; }

    /// <summary>Display name.</summary>
    public string? DisplayName { get; init; }

    /// <summary>True when the service principal is enabled.</summary>
    public bool? AccountEnabled { get; init; }
}

/// <summary>An OAuth2 delegated permission grant.</summary>
public sealed record GraphOAuth2PermissionGrantDto
{
    /// <summary>Grant ID.</summary>
    public string? Id { get; init; }

    /// <summary>The client service principal the grant applies to.</summary>
    public string? ClientId { get; init; }

    /// <summary><c>AllPrincipals</c> for admin consent, <c>Principal</c> for user consent.</summary>
    public string? ConsentType { get; init; }

    /// <summary>Space-separated granted scopes.</summary>
    public string? Scope { get; init; }

    /// <summary>The resource service principal.</summary>
    public string? ResourceId { get; init; }
}

/// <summary>Request body for the createLink action.</summary>
public sealed record CreateLinkRequest
{
    /// <summary><c>view</c>, <c>edit</c> or <c>embed</c>.</summary>
    public required string Type { get; init; }

    /// <summary><c>anonymous</c>, <c>organization</c> or <c>users</c>.</summary>
    public string? Scope { get; init; }

    /// <summary>Requested expiry, in ISO 8601.</summary>
    public string? ExpirationDateTime { get; init; }

    /// <summary>Whether existing inherited permissions are retained on first share.</summary>
    public bool? RetainInheritedPermissions { get; init; }
}

/// <summary>
/// Request body for the invite action, which is the only v1.0 operation that accepts
/// recipients. The createLink action has no recipients parameter.
/// </summary>
public sealed record InviteRequest
{
    /// <summary>The people to grant access to.</summary>
    public required IReadOnlyList<DriveRecipientDto> Recipients { get; init; }

    /// <summary><c>read</c> or <c>write</c>.</summary>
    public required IReadOnlyList<string> Roles { get; init; }

    /// <summary>Whether recipients must sign in.</summary>
    public bool RequireSignIn { get; init; } = true;

    /// <summary>
    /// Whether Microsoft sends a notification email. Defaults to false everywhere in this
    /// application; no message is ever sent unless the user explicitly asks for one.
    /// </summary>
    public bool SendInvitation { get; init; }

    /// <summary>An optional message, included only when an invitation is sent.</summary>
    public string? Message { get; init; }

    /// <summary>Requested expiry, in ISO 8601.</summary>
    public string? ExpirationDateTime { get; init; }

    /// <summary>Whether existing inherited permissions are retained on first share.</summary>
    public bool? RetainInheritedPermissions { get; init; }
}

/// <summary>A recipient of a sharing invitation.</summary>
public sealed record DriveRecipientDto
{
    /// <summary>The recipient's email address.</summary>
    public string? Email { get; init; }

    /// <summary>The recipient's object ID, when known.</summary>
    [JsonPropertyName("objectId")]
    public string? ObjectId { get; init; }
}

/// <summary>Request body for creating an application registration.</summary>
public sealed record CreateApplicationRequest
{
    /// <summary>Display name.</summary>
    public required string DisplayName { get; init; }

    /// <summary>Supported account types. Always single-tenant here.</summary>
    public required string SignInAudience { get; init; }

    /// <summary>Marks the app as a public client so no secret is needed.</summary>
    public required bool IsFallbackPublicClient { get; init; }

    /// <summary>Desktop platform configuration.</summary>
    public required GraphPublicClientDto PublicClient { get; init; }

    /// <summary>
    /// Requested API permissions. Supplying these in the initial POST is what lets the
    /// bootstrap identity stay create-only: no follow-up PATCH is required, and PATCH would
    /// need Application.ReadWrite.All.
    /// </summary>
    public required IReadOnlyList<GraphRequiredResourceAccessDto> RequiredResourceAccess { get; init; }
}

/// <summary>Request body for creating a service principal.</summary>
public sealed record CreateServicePrincipalRequest
{
    /// <summary>The application (client) ID to create a service principal for.</summary>
    public required string AppId { get; init; }
}

/// <summary>Request body for creating an upload session.</summary>
public sealed record CreateUploadSessionRequest
{
    /// <summary>Behaviour on conflict and the target name.</summary>
    public required UploadSessionItemDto Item { get; init; }
}

/// <summary>Item settings for an upload session.</summary>
public sealed record UploadSessionItemDto
{
    /// <summary>Conflict behaviour: <c>replace</c>, <c>rename</c> or <c>fail</c>.</summary>
    [JsonPropertyName("@microsoft.graph.conflictBehavior")]
    public required string ConflictBehavior { get; init; }

    /// <summary>The file name.</summary>
    public required string Name { get; init; }
}
