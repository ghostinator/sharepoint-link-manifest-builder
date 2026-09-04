using SharePointLinkManifestBuilder.Core.Models;

namespace SharePointLinkManifestBuilder.Graph.Http;

/// <summary>
/// The endpoint and scopes the Graph transport should use right now.
/// <para>
/// Held as a single mutable object so that changing tenant, switching account, or granting an
/// additional scope takes effect immediately without rebuilding the dependency graph. Access
/// is synchronized because the transport reads it from many concurrent worker tasks.
/// </para>
/// </summary>
public sealed class GraphClientContext
{
    private readonly Lock _gate = new();
    private string _endpoint = AuthorityDefaults.PublicCloudGraphEndpoint;
    private IReadOnlyList<string> _scopes = [];

    /// <summary>The Graph v1.0 base endpoint, which sovereign clouds may override.</summary>
    public string Endpoint
    {
        get
        {
            lock (_gate)
            {
                return _endpoint;
            }
        }
    }

    /// <summary>The scopes tokens are requested for.</summary>
    public IReadOnlyList<string> Scopes
    {
        get
        {
            lock (_gate)
            {
                return _scopes;
            }
        }
    }

    /// <summary>Points the transport at a tenant's endpoint and scope set.</summary>
    public void Update(string endpoint, IReadOnlyList<string> scopes)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(endpoint);
        ArgumentNullException.ThrowIfNull(scopes);

        lock (_gate)
        {
            _endpoint = endpoint.TrimEnd('/');
            _scopes = scopes;
        }
    }
}
