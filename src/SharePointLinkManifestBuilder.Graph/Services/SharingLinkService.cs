using System.Globalization;
using Microsoft.Extensions.Logging;
using SharePointLinkManifestBuilder.Core.Abstractions;
using SharePointLinkManifestBuilder.Core.Models;
using SharePointLinkManifestBuilder.Graph.Dto;
using SharePointLinkManifestBuilder.Graph.Http;

namespace SharePointLinkManifestBuilder.Graph.Services;

/// <summary>
/// Creates sharing links and grants recipients access.
/// <para>
/// Two Microsoft Graph behaviours drive the design here, both verified against the v1.0
/// reference rather than assumed:
/// </para>
/// <list type="number">
///   <item>
///     <c>createLink</c> returns <c>201 Created</c> for a new link and <c>200 OK</c> when an
///     equivalent link already existed. Reporting a 200 as "Created" would be a lie, so the
///     status code is what determines the recorded outcome.
///   </item>
///   <item>
///     The v1.0 <c>createLink</c> action has no <c>recipients</c> parameter. Named recipients
///     require the separate <c>invite</c> action, which can return <c>207 Multi-Status</c> when
///     some recipients succeed and others fail.
///   </item>
/// </list>
/// </summary>
public sealed class SharingLinkService : ISharingLinkService
{
    private readonly IGraphApiClient _client;
    private readonly ILogger<SharingLinkService> _logger;

    /// <summary>Creates the service.</summary>
    public SharingLinkService(IGraphApiClient client, ILogger<SharingLinkService> logger)
    {
        _client = client ?? throw new ArgumentNullException(nameof(client));
        _logger = logger ?? throw new ArgumentNullException(nameof(logger));
    }

    /// <inheritdoc />
    public async Task<LinkResult> CreateOrGetLinkAsync(
        DiscoveredFile file,
        LinkConfiguration configuration,
        CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(file);
        ArgumentNullException.ThrowIfNull(configuration);

        // Optional pre-check that avoids a write entirely when an equivalent link already
        // exists. Off by default because it costs an extra read per file.
        if (configuration.SkipWhenEquivalentLinkExists)
        {
            var existing = await GetExistingLinksAsync(file, cancellationToken).ConfigureAwait(false);

            if (existing.Succeeded && existing.Value!.FirstOrDefault(l => l.Matches(configuration)) is { } match)
            {
                return new LinkResult
                {
                    File = file,
                    Status = LinkResultStatus.Existing,
                    SharingUrl = match.WebUrl,
                    PermissionId = match.PermissionId,
                    GrantedLinkType = match.LinkType,
                    GrantedScope = match.Scope,
                    ExpirationUtc = match.ExpirationUtc,
                };
            }
        }

        var request = new CreateLinkRequest
        {
            Type = configuration.GraphLinkType,
            Scope = configuration.GraphScope,
            ExpirationDateTime = FormatExpiration(configuration.ExpirationUtc),

            // Only sent when it differs from the Graph default, so the request stays minimal
            // and a tenant default is not overridden by accident.
            RetainInheritedPermissions = configuration.RetainInheritedPermissions ? null : false,
        };

        var response = await _client
            .PostAsync<GraphPermissionDto>(
                GraphPaths.CreateLink(file.DriveId, file.ItemId), request, cancellationToken)
            .ConfigureAwait(false);

        if (!response.Succeeded || response.Value is null)
        {
            return Failed(file, response.Error
                ?? GraphErrorMapper.Map(response.StatusCode, null, null, "create a sharing link"));
        }

        var permission = response.Value;

        // The status code is the only reliable signal of whether a link was created or reused.
        var status = response.IsCreatedResource ? LinkResultStatus.Created : LinkResultStatus.Reused;

        var result = new LinkResult
        {
            File = file,
            Status = status,
            SharingUrl = permission.Link?.WebUrl,
            PermissionId = permission.Id,
            GrantedLinkType = permission.Link?.Type,
            GrantedScope = permission.Link?.Scope,
            ExpirationUtc = permission.ExpirationDateTime,
        };

        // Graph can honour a request only partially: for example returning an organization link
        // where an anonymous one was requested because tenant policy forbids anonymous sharing.
        // Recording the granted values, and warning, keeps the manifest honest.
        if (permission.Link is { } link
            && !string.Equals(link.Scope, configuration.GraphScope, StringComparison.OrdinalIgnoreCase))
        {
            _logger.LogWarning(
                "Requested link scope '{Requested}' but Microsoft 365 returned '{Granted}'. "
                + "Tenant policy may have adjusted the request.",
                configuration.GraphScope,
                link.Scope);
        }

        if (configuration.RequiresInviteAction)
        {
            var recipients = await InviteRecipientsAsync(file, configuration, cancellationToken)
                .ConfigureAwait(false);

            result = result with { RecipientResults = recipients };
        }

        return result;
    }

    /// <inheritdoc />
    public async Task<IReadOnlyList<RecipientResult>> InviteRecipientsAsync(
        DiscoveredFile file,
        LinkConfiguration configuration,
        CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(file);
        ArgumentNullException.ThrowIfNull(configuration);

        if (configuration.Recipients.Count == 0)
        {
            return [];
        }

        var request = new InviteRequest
        {
            Recipients = configuration.Recipients.Select(r => new DriveRecipientDto { Email = r }).ToArray(),
            Roles = configuration.GraphRoles,
            RequireSignIn = configuration.Audience != LinkAudience.Anyone,

            // Never send a notification unless the user explicitly asked for one.
            SendInvitation = configuration.SendInvitationEmail,
            Message = configuration.SendInvitationEmail ? configuration.InvitationMessage : null,
            ExpirationDateTime = FormatExpiration(configuration.ExpirationUtc),
            RetainInheritedPermissions = configuration.RetainInheritedPermissions ? null : false,
        };

        var response = await _client
            .PostAsync<GraphInviteResponse>(
                GraphPaths.Invite(file.DriveId, file.ItemId), request, cancellationToken)
            .ConfigureAwait(false);

        if (!response.Succeeded || response.Value?.Value is null)
        {
            var error = response.Error
                ?? GraphErrorMapper.Map(response.StatusCode, null, null, "grant access to recipients");

            // The whole call failed, so every recipient failed. Reporting them individually
            // keeps the results grid consistent between whole and partial failures.
            return configuration.Recipients
                .Select(r => new RecipientResult { Recipient = r, Succeeded = false, Error = error })
                .ToArray();
        }

        var results = new List<RecipientResult>();
        var returned = response.Value.Value;

        for (var i = 0; i < configuration.Recipients.Count; i++)
        {
            var requested = configuration.Recipients[i];

            // Match the response entry by the invitation email where Graph echoes it; fall back
            // to positional matching, which is the documented ordering.
            var entry = returned.FirstOrDefault(p =>
                    string.Equals(p.Invitation?.Email, requested, StringComparison.OrdinalIgnoreCase))
                ?? (i < returned.Count ? returned[i] : null);

            if (entry is null)
            {
                results.Add(new RecipientResult
                {
                    Recipient = requested,
                    Succeeded = false,
                    Error = new GraphError
                    {
                        Kind = GraphErrorKind.RecipientRejected,
                        Message = "Microsoft 365 did not report an outcome for this recipient.",
                    },
                });

                continue;
            }

            // A 207 carries a per-entry error object for the recipients that failed.
            if (entry.Error is { } entryError)
            {
                results.Add(new RecipientResult
                {
                    Recipient = requested,
                    Succeeded = false,
                    Error = new GraphError
                    {
                        Kind = GraphErrorKind.RecipientRejected,
                        Message = entryError.Message ?? "Microsoft 365 rejected this recipient.",
                        GraphErrorCode = entryError.Code,
                        StatusCode = response.StatusCode,
                    },
                });

                continue;
            }

            results.Add(new RecipientResult { Recipient = requested, Succeeded = true });
        }

        if (response.IsPartialSuccess)
        {
            _logger.LogWarning(
                "Sharing invitation partially succeeded: {Succeeded} of {Total} recipient(s) were granted access.",
                results.Count(r => r.Succeeded),
                results.Count);
        }

        return results;
    }

    /// <inheritdoc />
    public async Task<OperationResult<IReadOnlyList<ExistingSharingLink>>> GetExistingLinksAsync(
        DiscoveredFile file,
        CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(file);

        try
        {
            var links = new List<ExistingSharingLink>();

            await foreach (var permission in _client
                .EnumeratePagedAsync<GraphPermissionDto>(
                    GraphPaths.Permissions(file.DriveId, file.ItemId), cancellationToken)
                .ConfigureAwait(false))
            {
                // Only permissions that actually represent a link are relevant; direct grants
                // have no URL and cannot be recorded in a manifest.
                if (permission.Id is null || permission.Link is null)
                {
                    continue;
                }

                links.Add(new ExistingSharingLink
                {
                    PermissionId = permission.Id,
                    LinkType = permission.Link.Type,
                    Scope = permission.Link.Scope,
                    WebUrl = permission.Link.WebUrl,
                    ExpirationUtc = permission.ExpirationDateTime,
                });
            }

            return OperationResult<IReadOnlyList<ExistingSharingLink>>.Success(links);
        }
        catch (GraphOperationException ex)
        {
            return OperationResult<IReadOnlyList<ExistingSharingLink>>.Failure(ex.Error);
        }
    }

    /// <summary>
    /// Maps a normalized error onto the specific result status, so the results grid can show
    /// "blocked by policy" rather than a generic failure.
    /// </summary>
    internal static LinkResult Failed(DiscoveredFile file, GraphError error) => new()
    {
        File = file,
        Status = error.Kind switch
        {
            GraphErrorKind.AnonymousSharingDisabled or GraphErrorKind.ExternalSharingDisabled
                or GraphErrorKind.PolicyBlocked => LinkResultStatus.PolicyBlocked,

            GraphErrorKind.SharePointAccessDenied or GraphErrorKind.OneDriveAccessDenied
                => LinkResultStatus.AccessDenied,

            GraphErrorKind.UnsupportedLinkType or GraphErrorKind.UnsupportedItemType
                => LinkResultStatus.Unsupported,

            _ => LinkResultStatus.Failed,
        },
        Error = error,
    };

    /// <summary>
    /// Formats an expiry for Graph. The documented shape is <c>yyyy-MM-ddTHH:mm:ssZ</c>.
    /// </summary>
    internal static string? FormatExpiration(DateTimeOffset? value) =>
        value?.UtcDateTime.ToString("yyyy-MM-dd'T'HH:mm:ss'Z'", CultureInfo.InvariantCulture);
}

/// <summary>The envelope returned by the invite action.</summary>
internal sealed record GraphInviteResponse
{
    /// <summary>One permission per recipient.</summary>
    public IReadOnlyList<GraphPermissionDto>? Value { get; init; }
}
