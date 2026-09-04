namespace SharePointLinkManifestBuilder.Core.Models;

/// <summary>The permission a sharing link grants. Maps to the Graph <c>type</c> parameter.</summary>
[System.Diagnostics.CodeAnalysis.SuppressMessage(
    "Naming", "CA1711:Identifiers should not have incorrect suffix",
    Justification = "\"Permission\" is the exact term Microsoft Graph uses for this concept. " +
                    "Renaming it would make the code harder to match against the Graph documentation.")]
public enum LinkPermission
{
    /// <summary>Read-only. Graph <c>type: view</c>.</summary>
    View = 0,

    /// <summary>Read-write. Graph <c>type: edit</c>.</summary>
    Edit = 1,

    /// <summary>
    /// Embeddable link. Graph <c>type: embed</c>. Supported by OneDrive personal only, so it
    /// is offered only where applicable and never for SharePoint targets.
    /// </summary>
    Embed = 2,
}

/// <summary>Who can use a sharing link. Maps to the Graph <c>scope</c> parameter.</summary>
public enum LinkAudience
{
    /// <summary>Anyone signed into the organization. Graph <c>scope: organization</c>.</summary>
    Organization = 0,

    /// <summary>Only chosen people. Graph <c>scope: users</c>, plus an optional invite.</summary>
    SpecificPeople = 1,

    /// <summary>
    /// Anyone with the link, without signing in. Graph <c>scope: anonymous</c>. Frequently
    /// disabled by tenant policy; the request may be rejected and that is reported honestly.
    /// </summary>
    Anyone = 2,
}

/// <summary>What actually happened when a link was requested for one file.</summary>
public enum LinkResultStatus
{
    /// <summary>Graph returned 201: a new sharing link was created.</summary>
    Created = 0,

    /// <summary>Graph returned 200: an equivalent link already existed and was returned.</summary>
    Reused = 1,

    /// <summary>An equivalent link was found without requesting a new one.</summary>
    Existing = 2,

    /// <summary>No link was requested, by configuration or filtering.</summary>
    Skipped = 3,

    /// <summary>The request failed for a reason that is not a policy or access decision.</summary>
    Failed = 4,

    /// <summary>The requested link type is not supported for this item.</summary>
    Unsupported = 5,

    /// <summary>Tenant policy refused the requested link, for example anonymous sharing disabled.</summary>
    PolicyBlocked = 6,

    /// <summary>The signed-in user lacks permission to share this item.</summary>
    AccessDenied = 7,
}

/// <summary>
/// What the job will ask Microsoft 365 for. Every field is a <em>request</em>; the service
/// decides. Nothing here can override tenant policy.
/// </summary>
public sealed record LinkConfiguration
{
    /// <summary>Requested link permission.</summary>
    public LinkPermission Permission { get; init; } = LinkPermission.View;

    /// <summary>Requested link audience.</summary>
    public LinkAudience Audience { get; init; } = LinkAudience.Organization;

    /// <summary>
    /// Recipients for a specific-people link. Only meaningful when
    /// <see cref="Audience"/> is <see cref="LinkAudience.SpecificPeople"/>.
    /// </summary>
    public IReadOnlyList<string> Recipients { get; init; } = [];

    /// <summary>
    /// Whether to send Microsoft's invitation email. Default false: no message is ever sent
    /// unless the user explicitly asks for it.
    /// </summary>
    public bool SendInvitationEmail { get; init; }

    /// <summary>An optional message included with an invitation, when one is sent.</summary>
    public string? InvitationMessage { get; init; }

    /// <summary>Requested expiry. Support depends on tenant policy and licensing.</summary>
    public DateTimeOffset? ExpirationUtc { get; init; }

    /// <summary>
    /// Whether existing inherited permissions are retained when an item is shared for the first
    /// time. Maps to Graph <c>retainInheritedPermissions</c>; default true, matching Graph.
    /// </summary>
    public bool RetainInheritedPermissions { get; init; } = true;

    /// <summary>
    /// When true, an equivalent existing link is reused rather than a new one requested.
    /// Graph does this natively by returning 200 instead of 201.
    /// </summary>
    public bool ReuseExistingLinks { get; init; } = true;

    /// <summary>
    /// When true, files that already have an equivalent link are reported as Skipped rather
    /// than having a request sent at all. Reduces write operations against the tenant.
    /// </summary>
    public bool SkipWhenEquivalentLinkExists { get; init; }

    /// <summary>Graph <c>type</c> value for this configuration.</summary>
    public string GraphLinkType => Permission switch
    {
        LinkPermission.View => "view",
        LinkPermission.Edit => "edit",
        LinkPermission.Embed => "embed",
        _ => "view",
    };

    /// <summary>Graph <c>scope</c> value for this configuration.</summary>
    public string GraphScope => Audience switch
    {
        LinkAudience.Organization => "organization",
        LinkAudience.SpecificPeople => "users",
        LinkAudience.Anyone => "anonymous",
        _ => "organization",
    };

    /// <summary>Graph <c>roles</c> value used by the invite action.</summary>
    public IReadOnlyList<string> GraphRoles =>
        Permission == LinkPermission.Edit ? ["write"] : ["read"];

    /// <summary>
    /// True when the configuration needs the invite action in addition to createLink, because
    /// named recipients were supplied. The v1.0 createLink action has no recipients parameter.
    /// </summary>
    public bool RequiresInviteAction =>
        Audience == LinkAudience.SpecificPeople && Recipients.Count > 0;

    /// <summary>
    /// Validates combinations that Microsoft Graph or this application does not support,
    /// returning one message per problem. An empty result means the configuration is coherent.
    /// </summary>
    public IReadOnlyList<string> Validate()
    {
        var problems = new List<string>();

        if (Audience == LinkAudience.SpecificPeople && Recipients.Count == 0)
        {
            problems.Add(
                "A 'Specific people' link has no recipients. The link will be created with scope 'users', "
                + "but nobody will be granted access until people are added.");
        }

        if (Audience != LinkAudience.SpecificPeople && Recipients.Count > 0)
        {
            problems.Add("Recipients are only used when the audience is 'Specific people'.");
        }

        if (SendInvitationEmail && !RequiresInviteAction)
        {
            problems.Add(
                "An invitation email can only be sent when the audience is 'Specific people' and at least "
                + "one recipient is supplied.");
        }

        if (Permission == LinkPermission.Embed)
        {
            problems.Add(
                "Embed links are supported by OneDrive personal only and will fail for SharePoint and "
                + "OneDrive for Business items.");
        }

        if (ExpirationUtc.HasValue && ExpirationUtc.Value <= DateTimeOffset.UtcNow)
        {
            problems.Add("The expiration date is in the past.");
        }

        foreach (var invalid in Recipients.Where(r => !IsPlausibleRecipient(r)))
        {
            problems.Add($"Recipient '{invalid}' does not look like a valid email address.");
        }

        return problems;
    }

    /// <summary>
    /// A conservative syntactic check. Deliberately not a full RFC 5322 validator: Microsoft 365
    /// is the authority on whether a recipient is real, and over-strict local validation would
    /// reject addresses that work.
    /// </summary>
    public static bool IsPlausibleRecipient(string value)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return false;
        }

        var at = value.IndexOf('@', StringComparison.Ordinal);
        if (at <= 0 || at != value.LastIndexOf('@'))
        {
            return false;
        }

        var domain = value[(at + 1)..];
        return domain.Contains('.', StringComparison.Ordinal)
            && !domain.StartsWith('.')
            && !domain.EndsWith('.')
            && !value.Any(char.IsWhiteSpace);
    }

    /// <summary>A one-line summary of the request, shown in previews and manifest headers.</summary>
    public string Describe()
    {
        var audience = Audience switch
        {
            LinkAudience.Organization => "people in the organization",
            LinkAudience.SpecificPeople => Recipients.Count > 0
                ? $"{Recipients.Count} specific recipient(s)"
                : "specific people (none named yet)",
            LinkAudience.Anyone => "anyone with the link",
            _ => "unknown audience",
        };

        var expiry = ExpirationUtc.HasValue
            ? $", expiring {ExpirationUtc.Value.UtcDateTime:yyyy-MM-dd HH:mm} UTC"
            : string.Empty;

        return $"{Permission} access for {audience}{expiry}";
    }
}
