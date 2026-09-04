using System.Net;
using System.Net.Http.Headers;
using System.Text;
using Microsoft.Extensions.Logging.Abstractions;
using SharePointLinkManifestBuilder.Core.Abstractions;
using SharePointLinkManifestBuilder.Core.Models;
using SharePointLinkManifestBuilder.Core.Resilience;
using SharePointLinkManifestBuilder.Graph.Http;

namespace SharePointLinkManifestBuilder.Graph.Tests;

/// <summary>One canned HTTP response in a scripted exchange.</summary>
public sealed record FakeResponse
{
    /// <summary>Status code to return.</summary>
    public HttpStatusCode Status { get; init; } = HttpStatusCode.OK;

    /// <summary>JSON body, or null for no content.</summary>
    public string? Json { get; init; }

    /// <summary>Raw body, used for content downloads.</summary>
    public byte[]? Bytes { get; init; }

    /// <summary>ETag header to return.</summary>
    public string? ETag { get; init; }

    /// <summary>Retry-After, as a delay in seconds.</summary>
    public int? RetryAfterSeconds { get; init; }

    /// <summary>Retry-After, as an HTTP date.</summary>
    public DateTimeOffset? RetryAfterDate { get; init; }

    /// <summary>Creates a 200 with a JSON body.</summary>
    public static FakeResponse Ok(string json) => new() { Json = json };

    /// <summary>Creates a 201 with a JSON body.</summary>
    public static FakeResponse Created(string json) =>
        new() { Status = HttpStatusCode.Created, Json = json };

    /// <summary>Creates a status-only response, optionally with a Graph error body.</summary>
    public static FakeResponse Error(HttpStatusCode status, string? code = null, string? message = null) =>
        new()
        {
            Status = status,
            // Plain concatenation: a raw interpolated literal cannot express the JSON's
            // trailing "}}" without escalating the "$" count, which reads far worse than this.
            Json = code is null
                ? null
                : "{\"error\":{\"code\":\"" + code + "\",\"message\":\"" + (message ?? code) + "\"}}",
        };

    /// <summary>Creates a 429 with a Retry-After delay.</summary>
    public static FakeResponse Throttled(int retryAfterSeconds) => new()
    {
        Status = HttpStatusCode.TooManyRequests,
        RetryAfterSeconds = retryAfterSeconds,
        Json = """{"error":{"code":"activityLimitReached","message":"Too many requests"}}""",
    };
}

/// <summary>A request the fake handler observed.</summary>
/// <param name="Method">HTTP method.</param>
/// <param name="Uri">Full request URI.</param>
/// <param name="Body">Request body, when there was one.</param>
/// <param name="IfMatch">The If-Match header, when present.</param>
/// <param name="Authorization">The Authorization scheme, never the token value.</param>
/// <param name="ClientRequestId">The correlation ID this application sent.</param>
public readonly record struct ObservedRequest(
    string Method,
    Uri Uri,
    string? Body,
    string? IfMatch,
    string? Authorization,
    string? ClientRequestId);

/// <summary>
/// A scripted <see cref="HttpMessageHandler"/> standing in for Microsoft Graph.
/// <para>
/// Mocking at the transport boundary is deliberate: the real
/// <see cref="GraphApiClient"/> runs, so its pagination loop, retry policy, ETag handling and
/// error mapping are genuinely exercised. Faking <c>IGraphApiClient</c> instead would test the
/// mock rather than the code.
/// </para>
/// </summary>
public sealed class FakeGraphHandler : HttpMessageHandler
{
    private readonly Queue<FakeResponse> _sequence = new();
    private readonly Dictionary<string, FakeResponse> _byPathFragment = new(StringComparer.OrdinalIgnoreCase);
    private readonly List<ObservedRequest> _requests = [];

    /// <summary>Every request the handler saw, in order.</summary>
    public IReadOnlyList<ObservedRequest> Requests => _requests;

    /// <summary>How many requests were made.</summary>
    public int RequestCount => _requests.Count;

    /// <summary>Queues a response to be returned by the next unmatched request.</summary>
    public FakeGraphHandler Enqueue(FakeResponse response)
    {
        _sequence.Enqueue(response);
        return this;
    }

    /// <summary>Queues several responses in order.</summary>
    public FakeGraphHandler Enqueue(params FakeResponse[] responses)
    {
        foreach (var response in responses)
        {
            _sequence.Enqueue(response);
        }

        return this;
    }

    /// <summary>
    /// Maps a URL fragment to a response, for tests where request order is not the point.
    /// A path match wins over the queue.
    /// </summary>
    public FakeGraphHandler Map(string pathFragment, FakeResponse response)
    {
        _byPathFragment[pathFragment] = response;
        return this;
    }

    /// <inheritdoc />
    protected override async Task<HttpResponseMessage> SendAsync(
        HttpRequestMessage request,
        CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();

        var body = request.Content is null
            ? null
            : await request.Content.ReadAsStringAsync(cancellationToken).ConfigureAwait(false);

        _requests.Add(new ObservedRequest(
            request.Method.Method,
            request.RequestUri!,
            body,
            request.Headers.TryGetValues("If-Match", out var ifMatch) ? ifMatch.FirstOrDefault() : null,
            request.Headers.Authorization?.Scheme,
            request.Headers.TryGetValues(GraphApiClient.ClientRequestIdHeader, out var correlation)
                ? correlation.FirstOrDefault()
                : null));

        var url = request.RequestUri!.ToString();

        var matched = _byPathFragment
            .FirstOrDefault(kvp => url.Contains(kvp.Key, StringComparison.OrdinalIgnoreCase));

        var canned = matched.Key is not null
            ? matched.Value
            : _sequence.Count > 0
                ? _sequence.Dequeue()
                : FakeResponse.Error(HttpStatusCode.NotFound, "itemNotFound", "No response was scripted.");

        var response = new HttpResponseMessage(canned.Status);

        if (canned.Bytes is not null)
        {
            response.Content = new ByteArrayContent(canned.Bytes);
        }
        else if (canned.Json is not null)
        {
            response.Content = new StringContent(canned.Json, Encoding.UTF8, "application/json");
        }
        else
        {
            response.Content = new StringContent(string.Empty);
        }

        if (canned.ETag is not null)
        {
            response.Headers.ETag = new EntityTagHeaderValue(canned.ETag);
        }

        if (canned.RetryAfterSeconds is { } seconds)
        {
            response.Headers.RetryAfter = new RetryConditionHeaderValue(TimeSpan.FromSeconds(seconds));
        }
        else if (canned.RetryAfterDate is { } date)
        {
            response.Headers.RetryAfter = new RetryConditionHeaderValue(date);
        }

        response.Headers.TryAddWithoutValidation(GraphApiClient.ServiceRequestIdHeader, "fake-request-id");

        return response;
    }
}

/// <summary>An authentication stand-in that returns a fixed, obviously fake token.</summary>
public sealed class FakeAuthenticationService : IAuthenticationService
{
    /// <summary>The token handed to the transport. Deliberately not JWT-shaped.</summary>
    public const string Token = "FAKE-TEST-TOKEN-NOT-A-REAL-CREDENTIAL";

    /// <summary>When false, token acquisition fails, simulating a signed-out state.</summary>
    public bool IsSignedIn { get; set; } = true;

    /// <summary>
    /// When set, a silent acquisition fails with this error and an interactive one succeeds.
    /// Models the real AADSTS65001 case: the tenant is consented, but this user has no cached
    /// grant, so only an interactive request can produce a token.
    /// </summary>
    public GraphErrorKind? SilentOnlyFailure { get; set; }

    /// <summary>How many times a token was requested silently.</summary>
    public int SilentAttempts { get; private set; }

    /// <summary>How many times a token was requested interactively.</summary>
    public int InteractiveAttempts { get; private set; }

    /// <summary>Scopes the fake reports as granted.</summary>
    public IReadOnlyList<string> GrantedScopes { get; set; } =
        ["User.Read", "Sites.Read.All", "Files.ReadWrite.All"];

    /// <inheritdoc />
    public UserAccount? CurrentAccount { get; set; } = new()
    {
        UserId = "user-1",
        DisplayName = "Test User",
        UserPrincipalName = "test.user@example.test",
        TenantId = "11111111-1111-1111-1111-111111111111",
    };

    /// <inheritdoc />
    public event EventHandler<UserAccount?>? AccountChanged;

    /// <inheritdoc />
    public Task ConfigureAsync(TenantConfiguration configuration, CancellationToken cancellationToken = default) =>
        Task.CompletedTask;

    /// <inheritdoc />
    public Task<AuthenticationResultInfo> SignInAsync(
        IEnumerable<string> scopes,
        string? loginHint = null,
        CancellationToken cancellationToken = default)
    {
        AccountChanged?.Invoke(this, CurrentAccount);

        return Task.FromResult(new AuthenticationResultInfo
        {
            Succeeded = IsSignedIn,
            Account = CurrentAccount,
            GrantedScopes = GrantedScopes,
        });
    }

    /// <summary>Accounts the fake reports as cached, for account-switcher tests.</summary>
    public List<UserAccount> CachedAccounts { get; } = [];

    /// <summary>The home account ID passed to the last switch, for assertions.</summary>
    public string? LastSwitchedTo { get; private set; }

    /// <inheritdoc />
    public Task<AuthenticationResultInfo> SwitchToAccountAsync(
        string homeAccountId,
        IEnumerable<string> scopes,
        CancellationToken cancellationToken = default)
    {
        LastSwitchedTo = homeAccountId;

        var match = CachedAccounts.FirstOrDefault(a =>
            string.Equals(a.HomeAccountId, homeAccountId, StringComparison.Ordinal));

        if (match is null)
        {
            return Task.FromResult(new AuthenticationResultInfo
            {
                Succeeded = false,
                Error = new GraphError
                {
                    Kind = GraphErrorKind.AuthenticationFailed,
                    Message = "That account is no longer cached on this device.",
                },
            });
        }

        CurrentAccount = match;
        AccountChanged?.Invoke(this, match);

        return Task.FromResult(new AuthenticationResultInfo
        {
            Succeeded = true,
            Account = match,
            GrantedScopes = GrantedScopes,
        });
    }

    /// <inheritdoc />
    public Task<AuthenticationResultInfo> AcquireTokenAsync(
        IEnumerable<string> scopes,
        bool allowInteractive = false,
        CancellationToken cancellationToken = default)
    {
        if (allowInteractive)
        {
            InteractiveAttempts++;
        }
        else
        {
            SilentAttempts++;
        }

        if (SilentOnlyFailure is { } kind && !allowInteractive)
        {
            return Task.FromResult(new AuthenticationResultInfo
            {
                Succeeded = false,
                Error = new GraphError { Kind = kind, Message = "Interaction required." },
            });
        }

        return Task.FromResult(new AuthenticationResultInfo
        {
            Succeeded = IsSignedIn,
            Account = CurrentAccount,
            GrantedScopes = GrantedScopes,
            Error = IsSignedIn
                ? null
                : new GraphError { Kind = GraphErrorKind.AuthenticationFailed, Message = "Not signed in." },
        });
    }

    /// <inheritdoc />
    public Task<string?> GetAccessTokenAsync(
        IEnumerable<string> scopes,
        CancellationToken cancellationToken = default) =>
        Task.FromResult(IsSignedIn ? Token : null);

    /// <inheritdoc />
    public Task<IReadOnlyList<UserAccount>> GetCachedAccountsAsync(CancellationToken cancellationToken = default) =>
        Task.FromResult<IReadOnlyList<UserAccount>>(
            CurrentAccount is null ? [] : [CurrentAccount]);

    /// <inheritdoc />
    public Task SignOutAsync(CancellationToken cancellationToken = default)
    {
        CurrentAccount = null;
        AccountChanged?.Invoke(this, null);
        return Task.CompletedTask;
    }

    /// <inheritdoc />
    public Task ForgetAccountAsync(string homeAccountId, CancellationToken cancellationToken = default) =>
        Task.CompletedTask;
}

/// <summary>Builds a real <see cref="GraphApiClient"/> wired to a fake transport.</summary>
public static class GraphTestHarness
{
    /// <summary>The synthetic Graph endpoint used by tests. Never a real host.</summary>
    public const string Endpoint = "https://graph.example.test/v1.0";

    /// <summary>Creates a client over the supplied handler.</summary>
    /// <param name="handler">The scripted transport.</param>
    /// <param name="authentication">Authentication stand-in, or a default signed-in fake.</param>
    /// <param name="maxAttempts">Retry attempts; 1 disables retry.</param>
    public static GraphApiClient CreateClient(
        FakeGraphHandler handler,
        FakeAuthenticationService? authentication = null,
        int maxAttempts = 3)
    {
        var context = new GraphClientContext();
        context.Update(Endpoint, ["User.Read", "Sites.Read.All", "Files.ReadWrite.All"]);

        // Zero base delay keeps tests fast while still exercising the real retry loop; the
        // policy's timing behaviour is covered separately by GraphRetryPolicyTests.
        var policy = new GraphRetryPolicy(
            maxAttempts,
            TimeSpan.Zero,
            TimeSpan.FromMilliseconds(1),
            () => 0.0);

        return new GraphApiClient(
            new HttpClient(handler),
            authentication ?? new FakeAuthenticationService(),
            context,
            policy,
            NullLogger<GraphApiClient>.Instance);
    }
}
