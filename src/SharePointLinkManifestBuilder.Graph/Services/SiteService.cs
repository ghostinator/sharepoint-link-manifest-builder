using Microsoft.Extensions.Logging;
using SharePointLinkManifestBuilder.Core.Abstractions;
using SharePointLinkManifestBuilder.Core.Models;
using SharePointLinkManifestBuilder.Core.Urls;
using SharePointLinkManifestBuilder.Graph.Dto;
using SharePointLinkManifestBuilder.Graph.Http;

namespace SharePointLinkManifestBuilder.Graph.Services;

/// <summary>SharePoint site discovery and metadata over Microsoft Graph.</summary>
public sealed class SiteService : ISiteService
{
    private readonly IGraphApiClient _client;
    private readonly ILogger<SiteService> _logger;

    /// <summary>Creates the service.</summary>
    public SiteService(IGraphApiClient client, ILogger<SiteService> logger)
    {
        _client = client ?? throw new ArgumentNullException(nameof(client));
        _logger = logger ?? throw new ArgumentNullException(nameof(logger));
    }

    /// <inheritdoc />
    public async Task<OperationResult<IReadOnlyList<SharePointSite>>> SearchSitesAsync(
        string query,
        CancellationToken cancellationToken = default)
    {
        if (string.IsNullOrWhiteSpace(query))
        {
            return OperationResult<IReadOnlyList<SharePointSite>>.Success([]);
        }

        try
        {
            var sites = new List<SharePointSite>();

            await foreach (var dto in _client
                .EnumeratePagedAsync<GraphSiteDto>(GraphPaths.SearchSites(query), cancellationToken)
                .ConfigureAwait(false))
            {
                if (Map(dto) is { } site)
                {
                    sites.Add(site);
                }
            }

            _logger.LogInformation(
                "Site search returned {Count} site(s). This reflects the search index for the signed-in "
                + "user and is not necessarily every site in the tenant.",
                sites.Count);

            return OperationResult<IReadOnlyList<SharePointSite>>.Success(sites);
        }
        catch (GraphOperationException ex)
        {
            return OperationResult<IReadOnlyList<SharePointSite>>.Failure(ex.Error);
        }
    }

    /// <inheritdoc />
    public async Task<OperationResult<SharePointSite>> ResolveSiteByUrlAsync(
        string siteUrl,
        CancellationToken cancellationToken = default)
    {
        var parsed = SharePointUrlParser.Parse(siteUrl);

        if (!parsed.IsValid)
        {
            return OperationResult<SharePointSite>.Failure(new GraphError
            {
                Kind = GraphErrorKind.InvalidUrl,
                Message = parsed.FailureReason ?? "That URL could not be understood.",
                SuggestedAction = "Paste the address of a SharePoint site, library or folder.",
            });
        }

        if (parsed.Kind == ResourceUrlKind.SharingLink)
        {
            return OperationResult<SharePointSite>.Failure(new GraphError
            {
                Kind = GraphErrorKind.InvalidUrl,
                Message = "That is a sharing link, which points at a single item rather than a site.",
                SuggestedAction = "Use the OneDrive or SharePoint browser to select a location, or paste a site URL.",
            });
        }

        var path = GraphPaths.SiteByPath(parsed.Hostname!, parsed.SitePath);
        var response = await _client.GetAsync<GraphSiteDto>(path, cancellationToken).ConfigureAwait(false);

        if (!response.Succeeded || response.Value is null)
        {
            return OperationResult<SharePointSite>.Failure(
                response.Error ?? GraphErrorMapper.Map(response.StatusCode, null, null, "resolve a SharePoint site"));
        }

        return Map(response.Value) is { } site
            ? OperationResult<SharePointSite>.Success(site)
            : OperationResult<SharePointSite>.Failure(new GraphError
            {
                Kind = GraphErrorKind.SiteNotFound,
                Message = "Microsoft Graph returned a site without an identifier.",
            });
    }

    /// <inheritdoc />
    public async Task<OperationResult<SharePointSite>> GetSiteAsync(
        string siteId,
        CancellationToken cancellationToken = default) =>
        await GetSiteFromPathAsync(GraphPaths.Site(siteId), cancellationToken).ConfigureAwait(false);

    /// <inheritdoc />
    public async Task<OperationResult<SharePointSite>> GetRootSiteAsync(
        CancellationToken cancellationToken = default) =>
        await GetSiteFromPathAsync(GraphPaths.RootSite(), cancellationToken).ConfigureAwait(false);

    /// <inheritdoc />
    public async Task<OperationResult<IReadOnlyList<SharePointSite>>> GetFollowedSitesAsync(
        CancellationToken cancellationToken = default)
    {
        try
        {
            var sites = new List<SharePointSite>();

            await foreach (var dto in _client
                .EnumeratePagedAsync<GraphSiteDto>(GraphPaths.FollowedSites(), cancellationToken)
                .ConfigureAwait(false))
            {
                if (Map(dto) is { } site)
                {
                    sites.Add(site);
                }
            }

            return OperationResult<IReadOnlyList<SharePointSite>>.Success(sites);
        }
        catch (GraphOperationException ex)
        {
            return OperationResult<IReadOnlyList<SharePointSite>>.Failure(ex.Error);
        }
    }

    /// <inheritdoc />
    public async Task<OperationResult<IReadOnlyList<DriveResource>>> GetSiteDrivesAsync(
        string siteId,
        CancellationToken cancellationToken = default)
    {
        try
        {
            var drives = new List<DriveResource>();

            await foreach (var dto in _client
                .EnumeratePagedAsync<GraphDriveDto>(GraphPaths.SiteDrives(siteId), cancellationToken)
                .ConfigureAwait(false))
            {
                if (string.IsNullOrEmpty(dto.Id))
                {
                    continue;
                }

                drives.Add(new DriveResource
                {
                    DriveId = dto.Id,
                    Name = dto.Name ?? "Documents",
                    DriveType = dto.DriveType,
                    WebUrl = dto.WebUrl,
                    SiteId = siteId,
                });
            }

            return OperationResult<IReadOnlyList<DriveResource>>.Success(drives);
        }
        catch (GraphOperationException ex)
        {
            return OperationResult<IReadOnlyList<DriveResource>>.Failure(ex.Error);
        }
    }

    private async Task<OperationResult<SharePointSite>> GetSiteFromPathAsync(
        string path,
        CancellationToken cancellationToken)
    {
        var response = await _client.GetAsync<GraphSiteDto>(path, cancellationToken).ConfigureAwait(false);

        if (!response.Succeeded || response.Value is null)
        {
            return OperationResult<SharePointSite>.Failure(
                response.Error ?? GraphErrorMapper.Map(response.StatusCode, null, null, "read a SharePoint site"));
        }

        return Map(response.Value) is { } site
            ? OperationResult<SharePointSite>.Success(site)
            : OperationResult<SharePointSite>.Failure(new GraphError
            {
                Kind = GraphErrorKind.SiteNotFound,
                Message = "Microsoft Graph returned a site without an identifier.",
            });
    }

    /// <summary>
    /// Maps a Graph site. Returns null when the payload has no ID, since a site that cannot be
    /// addressed is useless downstream and silently keeping it would produce confusing failures
    /// much later in a job.
    /// </summary>
    internal static SharePointSite? Map(GraphSiteDto dto)
    {
        if (string.IsNullOrEmpty(dto.Id))
        {
            return null;
        }

        string? hostname = null;
        if (!string.IsNullOrEmpty(dto.WebUrl)
            && Uri.TryCreate(dto.WebUrl, UriKind.Absolute, out var uri))
        {
            hostname = uri.Host;
        }

        return new SharePointSite
        {
            SiteId = dto.Id,

            // Some sites return only a name; falling back keeps the tree readable rather than
            // showing an opaque composite ID to the user.
            DisplayName = FirstNonEmpty(dto.DisplayName, dto.Name) ?? "Untitled site",
            WebUrl = dto.WebUrl,
            Hostname = hostname,
            Description = dto.Description,
            IsRootSite = dto.Root is not null,
            CreatedUtc = dto.CreatedDateTime,
        };
    }

    private static string? FirstNonEmpty(params string?[] values) =>
        values.FirstOrDefault(v => !string.IsNullOrWhiteSpace(v));
}
