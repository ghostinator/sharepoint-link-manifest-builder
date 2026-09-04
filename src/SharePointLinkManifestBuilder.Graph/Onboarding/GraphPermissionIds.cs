using SharePointLinkManifestBuilder.Core.Models;
using SharePointLinkManifestBuilder.Graph.Dto;

namespace SharePointLinkManifestBuilder.Graph.Onboarding;

/// <summary>
/// The fixed identifiers Microsoft Entra uses for Microsoft Graph delegated permissions.
/// <para>
/// The <c>requiredResourceAccess</c> property of an application registration takes permission
/// <em>GUIDs</em>, not names. These identifiers are assigned by Microsoft, are published as
/// part of the Microsoft Graph service principal, and are identical in every tenant worldwide,
/// which is why they can be treated as constants. The Azure CLI and the Terraform AzureAD
/// provider embed the same values for the same reason.
/// </para>
/// <para>
/// They are collected here so they can be audited in one place. To confirm any of them in a
/// tenant:
/// <c>GET /servicePrincipals(appId='00000003-0000-0000-c000-000000000000')?$select=oauth2PermissionScopes</c>
/// and match on the <c>value</c> field.
/// </para>
/// </summary>
public static class GraphPermissionIds
{
    /// <summary>Type marker used for delegated permissions in <c>requiredResourceAccess</c>.</summary>
    public const string DelegatedScopeType = "Scope";

    /// <summary>Type marker used for application permissions. This product does not request any.</summary>
    public const string ApplicationRoleType = "Role";

    /// <summary>Delegated <c>User.Read</c>.</summary>
    public const string UserRead = "e1fe6dd8-ba31-4d61-89e7-88639da4683d";

    /// <summary>Delegated <c>User.ReadBasic.All</c>.</summary>
    public const string UserReadBasicAll = "b340eb25-3456-403f-be2f-af7a0d370277";

    /// <summary>Delegated <c>Sites.Read.All</c>.</summary>
    public const string SitesReadAll = "205e70e5-aba6-4c52-a976-6d2d46c48043";

    /// <summary>Delegated <c>Sites.ReadWrite.All</c>.</summary>
    public const string SitesReadWriteAll = "89fe6a52-be36-487e-b7d8-d061c450a026";

    /// <summary>Delegated <c>Files.Read.All</c>.</summary>
    public const string FilesReadAll = "df85f4d6-205c-4ac5-a5ea-6bf408dba283";

    /// <summary>Delegated <c>Files.ReadWrite.All</c>.</summary>
    public const string FilesReadWriteAll = "863451e7-0667-486c-a5d6-d135439485f0";

    private static readonly Dictionary<string, string> ScopeIdsByName =
        new(StringComparer.OrdinalIgnoreCase)
        {
            ["User.Read"] = UserRead,
            ["User.ReadBasic.All"] = UserReadBasicAll,
            ["Sites.Read.All"] = SitesReadAll,
            ["Sites.ReadWrite.All"] = SitesReadWriteAll,
            ["Files.Read.All"] = FilesReadAll,
            ["Files.ReadWrite.All"] = FilesReadWriteAll,
        };

    /// <summary>Looks up the identifier for a delegated scope name, or null when unknown.</summary>
    public static string? TryGetScopeId(string scopeName) =>
        ScopeIdsByName.TryGetValue(scopeName, out var id) ? id : null;

    /// <summary>
    /// Builds the <c>requiredResourceAccess</c> block for an application registration.
    /// <para>
    /// A scope whose identifier is unknown to this build is omitted rather than guessed at, and
    /// reported through <paramref name="unmapped"/> so the wizard can tell the user plainly
    /// which permissions it could not configure automatically.
    /// </para>
    /// </summary>
    /// <param name="permissions">The delegated permissions to request.</param>
    /// <param name="unmapped">Scope names that had no known identifier.</param>
    public static IReadOnlyList<GraphRequiredResourceAccessDto> BuildRequiredResourceAccess(
        IEnumerable<PermissionRequirement> permissions,
        out IReadOnlyList<string> unmapped)
    {
        ArgumentNullException.ThrowIfNull(permissions);

        var missing = new List<string>();
        var access = new List<GraphResourceAccessDto>();

        foreach (var permission in permissions)
        {
            if (GraphScopes.Reserved.Contains(permission.Scope, StringComparer.OrdinalIgnoreCase))
            {
                // openid, profile and offline_access are implicit for a public client and are
                // not listed in requiredResourceAccess.
                continue;
            }

            var id = TryGetScopeId(permission.Scope);

            if (id is null)
            {
                missing.Add(permission.Scope);
                continue;
            }

            access.Add(new GraphResourceAccessDto { Id = id, Type = DelegatedScopeType });
        }

        unmapped = missing;

        if (access.Count == 0)
        {
            return [];
        }

        return
        [
            new GraphRequiredResourceAccessDto
            {
                ResourceAppId = AuthorityDefaults.MicrosoftGraphResourceAppId,
                ResourceAccess = access.DistinctBy(a => a.Id).ToArray(),
            },
        ];
    }
}
