using System.Net;
using System.Text.Json;
using SharePointLinkManifestBuilder.Core.Models;
using SharePointLinkManifestBuilder.Graph.Http;

namespace SharePointLinkManifestBuilder.Graph.Tests;

/// <summary>
/// Exercises the real transport against a scripted Microsoft Graph.
/// <para>
/// These cover the behaviour a live tenant would otherwise be needed to observe: pagination
/// across pages, throttling followed by success, permanent failures, ETag conflicts,
/// <c>207 Multi-Status</c>, and cancellation part-way through.
/// </para>
/// </summary>
public sealed class GraphApiClientTests
{
    private sealed record Item(string Id, string Name);

    [Fact]
    public async Task GetAsync_SuccessfulResponse_ReturnsTheDeserializedPayload()
    {
        var handler = new FakeGraphHandler()
            .Enqueue(FakeResponse.Ok("""{"id":"abc","name":"Report.docx"}"""));

        var response = await GraphTestHarness.CreateClient(handler)
            .GetAsync<Item>("/drives/d1/items/i1");

        Assert.True(response.Succeeded);
        Assert.Equal("abc", response.Value!.Id);
        Assert.Equal(200, response.StatusCode);
    }

    /// <summary>Every request must carry a correlation ID so a failure can be traced with support.</summary>
    [Fact]
    public async Task EveryRequest_CarriesACorrelationIdAndBearerScheme()
    {
        var handler = new FakeGraphHandler().Enqueue(FakeResponse.Ok("""{"id":"abc","name":"x"}"""));

        await GraphTestHarness.CreateClient(handler).GetAsync<Item>("/me");

        var request = Assert.Single(handler.Requests);
        Assert.False(string.IsNullOrWhiteSpace(request.ClientRequestId));
        Assert.True(Guid.TryParse(request.ClientRequestId, out _));
        Assert.Equal("Bearer", request.Authorization);
    }

    /// <summary>Pagination must follow every nextLink, not just the first page.</summary>
    [Fact]
    public async Task EnumeratePagedAsync_FollowsNextLinkAcrossThreePages()
    {
        var handler = new FakeGraphHandler().Enqueue(
            FakeResponse.Ok($$"""
                {"value":[{"id":"1","name":"a"},{"id":"2","name":"b"}],
                 "@odata.nextLink":"{{GraphTestHarness.Endpoint}}/page2"}
                """),
            FakeResponse.Ok($$"""
                {"value":[{"id":"3","name":"c"}],
                 "@odata.nextLink":"{{GraphTestHarness.Endpoint}}/page3"}
                """),
            FakeResponse.Ok("""{"value":[{"id":"4","name":"d"}]}"""));

        var items = new List<Item>();

        await foreach (var item in GraphTestHarness.CreateClient(handler)
            .EnumeratePagedAsync<Item>("/drives/d1/items/root/children"))
        {
            items.Add(item);
        }

        Assert.Equal(4, items.Count);
        Assert.Equal(["1", "2", "3", "4"], items.Select(i => i.Id));
        Assert.Equal(3, handler.RequestCount);
    }

    [Fact]
    public async Task EnumeratePagedAsync_EmptyCollection_YieldsNothingAndMakesOneRequest()
    {
        var handler = new FakeGraphHandler().Enqueue(FakeResponse.Ok("""{"value":[]}"""));

        var items = new List<Item>();

        await foreach (var item in GraphTestHarness.CreateClient(handler)
            .EnumeratePagedAsync<Item>("/drives/d1/items/root/children"))
        {
            items.Add(item);
        }

        Assert.Empty(items);
        Assert.Equal(1, handler.RequestCount);
    }

    /// <summary>
    /// A failure part-way through pagination must throw, not end the sequence quietly. Silently
    /// stopping would be indistinguishable from "the folder is empty", and a job would report
    /// success having enumerated nothing.
    /// </summary>
    [Fact]
    public async Task EnumeratePagedAsync_FailureMidPagination_Throws()
    {
        var handler = new FakeGraphHandler().Enqueue(
            FakeResponse.Ok($$"""
                {"value":[{"id":"1","name":"a"}],
                 "@odata.nextLink":"{{GraphTestHarness.Endpoint}}/page2"}
                """),
            FakeResponse.Error(HttpStatusCode.Forbidden, "accessDenied", "Access denied"));

        var client = GraphTestHarness.CreateClient(handler);
        var items = new List<Item>();

        var exception = await Assert.ThrowsAsync<GraphOperationException>(async () =>
        {
            await foreach (var item in client.EnumeratePagedAsync<Item>("/drives/d1/items/root/children"))
            {
                items.Add(item);
            }
        });

        Assert.Single(items);
        Assert.Equal(GraphErrorKind.SharePointAccessDenied, exception.Error.Kind);
    }

    /// <summary>Throttling must be retried, and the retry must succeed transparently.</summary>
    [Fact]
    public async Task Throttled_ThenSuccess_RetriesAndSucceeds()
    {
        var handler = new FakeGraphHandler().Enqueue(
            FakeResponse.Throttled(1),
            FakeResponse.Ok("""{"id":"abc","name":"Report.docx"}"""));

        var response = await GraphTestHarness.CreateClient(handler).GetAsync<Item>("/me");

        Assert.True(response.Succeeded);
        Assert.Equal(2, handler.RequestCount);
    }

    [Fact]
    public async Task Throttled_BeyondTheRetryLimit_ReportsThrottling()
    {
        var handler = new FakeGraphHandler().Enqueue(
            FakeResponse.Throttled(1), FakeResponse.Throttled(1), FakeResponse.Throttled(1));

        var response = await GraphTestHarness.CreateClient(handler, maxAttempts: 3)
            .GetAsync<Item>("/me");

        Assert.False(response.Succeeded);
        Assert.Equal(GraphErrorKind.Throttled, response.Error!.Kind);
        Assert.True(response.Error.IsRetryable);
        Assert.Equal(3, handler.RequestCount);
    }

    /// <summary>Retry-After may arrive as an HTTP date rather than a delay; both must work.</summary>
    [Fact]
    public async Task RetryAfterAsHttpDate_IsHonoured()
    {
        var handler = new FakeGraphHandler().Enqueue(
            new FakeResponse
            {
                Status = HttpStatusCode.TooManyRequests,
                RetryAfterDate = DateTimeOffset.UtcNow.AddMilliseconds(1),
            },
            FakeResponse.Ok("""{"id":"abc","name":"x"}"""));

        var response = await GraphTestHarness.CreateClient(handler).GetAsync<Item>("/me");

        Assert.True(response.Succeeded);
        Assert.Equal(2, handler.RequestCount);
    }

    /// <summary>A permanent 403 must fail immediately rather than consuming the retry budget.</summary>
    [Fact]
    public async Task PermanentForbidden_IsNotRetried()
    {
        var handler = new FakeGraphHandler()
            .Enqueue(FakeResponse.Error(HttpStatusCode.Forbidden, "accessDenied", "Access denied"));

        var response = await GraphTestHarness.CreateClient(handler).GetAsync<Item>("/sites/s1");

        Assert.False(response.Succeeded);
        Assert.Equal(GraphErrorKind.SharePointAccessDenied, response.Error!.Kind);
        Assert.False(response.Error.IsRetryable);
        Assert.Equal(1, handler.RequestCount);
    }

    [Fact]
    public async Task NotFound_ForASite_MapsToSiteNotFound()
    {
        var handler = new FakeGraphHandler()
            .Enqueue(FakeResponse.Error(HttpStatusCode.NotFound, "itemNotFound", "Not found"));

        var response = await GraphTestHarness.CreateClient(handler).GetAsync<Item>("/sites/does-not-exist");

        Assert.Equal(GraphErrorKind.SiteNotFound, response.Error!.Kind);
    }

    [Fact]
    public async Task NotFound_ForADriveItem_MapsToDeletedDuringProcessing()
    {
        var handler = new FakeGraphHandler()
            .Enqueue(FakeResponse.Error(HttpStatusCode.NotFound, "itemNotFound", "Not found"));

        var response = await GraphTestHarness.CreateClient(handler)
            .GetAsync<Item>("/drives/d1/items/gone");

        Assert.Equal(GraphErrorKind.FileDeletedDuringProcessing, response.Error!.Kind);
    }

    [Fact]
    public async Task Unauthorized_MapsToAuthenticationFailed()
    {
        var handler = new FakeGraphHandler()
            .Enqueue(FakeResponse.Error(HttpStatusCode.Unauthorized, "InvalidAuthenticationToken"));

        var response = await GraphTestHarness.CreateClient(handler).GetAsync<Item>("/me");

        Assert.Equal(GraphErrorKind.AuthenticationFailed, response.Error!.Kind);
    }

    [Fact]
    public async Task ServiceUnavailable_IsRetriedThenReported()
    {
        var handler = new FakeGraphHandler().Enqueue(
            FakeResponse.Error(HttpStatusCode.ServiceUnavailable),
            FakeResponse.Error(HttpStatusCode.ServiceUnavailable),
            FakeResponse.Error(HttpStatusCode.ServiceUnavailable));

        var response = await GraphTestHarness.CreateClient(handler, maxAttempts: 3).GetAsync<Item>("/me");

        Assert.False(response.Succeeded);
        Assert.Equal(GraphErrorKind.ServiceUnavailable, response.Error!.Kind);
        Assert.Equal(3, handler.RequestCount);
    }

    /// <summary>Without a token the transport must fail cleanly rather than sending an unauthenticated request.</summary>
    [Fact]
    public async Task NoSignedInAccount_FailsWithoutMakingARequest()
    {
        var handler = new FakeGraphHandler();
        var authentication = new FakeAuthenticationService { IsSignedIn = false };

        var response = await GraphTestHarness.CreateClient(handler, authentication).GetAsync<Item>("/me");

        Assert.False(response.Succeeded);
        Assert.Equal(GraphErrorKind.AuthenticationFailed, response.Error!.Kind);
        Assert.Equal(0, handler.RequestCount);
    }

    /// <summary>createLink's 201 versus 200 distinction is surfaced by the transport.</summary>
    [Fact]
    public async Task CreatedVersusOk_IsDistinguishable()
    {
        var created = new FakeGraphHandler().Enqueue(FakeResponse.Created("""{"id":"p1"}"""));
        var existing = new FakeGraphHandler().Enqueue(FakeResponse.Ok("""{"id":"p1"}"""));

        var createdResponse = await GraphTestHarness.CreateClient(created)
            .PostAsync<JsonElement>("/drives/d1/items/i1/createLink", new { type = "view" });

        var existingResponse = await GraphTestHarness.CreateClient(existing)
            .PostAsync<JsonElement>("/drives/d1/items/i1/createLink", new { type = "view" });

        Assert.True(createdResponse.IsCreatedResource);
        Assert.False(createdResponse.IsExistingResource);

        Assert.True(existingResponse.IsExistingResource);
        Assert.False(existingResponse.IsCreatedResource);
    }

    /// <summary>A 207 must be treated as a success carrying per-entry outcomes, not as a failure.</summary>
    [Fact]
    public async Task MultiStatus_IsTreatedAsSuccessWithPartialFlag()
    {
        var handler = new FakeGraphHandler()
            .Enqueue(new FakeResponse { Status = HttpStatusCode.MultiStatus, Json = """{"value":[]}""" });

        var response = await GraphTestHarness.CreateClient(handler)
            .PostAsync<JsonElement>("/drives/d1/items/i1/invite", new { });

        Assert.True(response.Succeeded);
        Assert.True(response.IsPartialSuccess);
    }

    /// <summary>A conditional write must actually send If-Match.</summary>
    [Fact]
    public async Task PutContentAsync_WithETag_SendsIfMatch()
    {
        var handler = new FakeGraphHandler().Enqueue(FakeResponse.Ok("""{"id":"i1"}"""));

        await GraphTestHarness.CreateClient(handler).PutContentAsync<JsonElement>(
            "/drives/d1/items/root:/manifest.txt:/content",
            new byte[] { 1, 2, 3 },
            "text/plain",
            "\"etag-1\"");

        Assert.Equal("\"etag-1\"", Assert.Single(handler.Requests).IfMatch);
    }

    /// <summary>412 must not be retried: the remote copy changed and retrying could clobber it.</summary>
    [Fact]
    public async Task PreconditionFailed_IsNotRetriedAndMapsToETagConflict()
    {
        var handler = new FakeGraphHandler().Enqueue(
            FakeResponse.Error(HttpStatusCode.PreconditionFailed, "resourceModified"));

        var response = await GraphTestHarness.CreateClient(handler).PutContentAsync<JsonElement>(
            "/drives/d1/items/root:/manifest.txt:/content",
            new byte[] { 1 },
            "text/plain",
            "\"stale\"");

        Assert.False(response.Succeeded);
        Assert.Equal(GraphErrorKind.ETagConflict, response.Error!.Kind);
        Assert.Equal(1, handler.RequestCount);
    }

    /// <summary>An upload-session chunk is pre-authorized; attaching a bearer token would leak it wider.</summary>
    [Fact]
    public async Task UploadSessionChunk_DoesNotAttachABearerToken()
    {
        var handler = new FakeGraphHandler().Enqueue(FakeResponse.Ok("""{"id":"i1"}"""));

        await GraphTestHarness.CreateClient(handler).PutUploadSessionChunkAsync(
            new Uri("https://upload.example.test/session/abc"),
            new byte[] { 1, 2, 3 },
            new System.Net.Http.Headers.ContentRangeHeaderValue(0, 2, 3),
            3);

        Assert.Null(Assert.Single(handler.Requests).Authorization);
    }

    [Fact]
    public async Task GetContentAsync_ReturnsRawBytes()
    {
        var payload = System.Text.Encoding.UTF8.GetBytes("SharePoint Link Manifest\n");
        var handler = new FakeGraphHandler().Enqueue(new FakeResponse { Bytes = payload, ETag = "\"e1\"" });

        var response = await GraphTestHarness.CreateClient(handler)
            .GetContentAsync("/drives/d1/items/root:/manifest.txt:/content");

        Assert.True(response.Succeeded);
        Assert.Equal(payload, response.Value);
    }

    /// <summary>Cancellation part-way through pagination must stop promptly and propagate.</summary>
    [Fact]
    public async Task Cancellation_DuringPagination_StopsEnumeration()
    {
        var handler = new FakeGraphHandler().Enqueue(
            FakeResponse.Ok($$"""
                {"value":[{"id":"1","name":"a"}],
                 "@odata.nextLink":"{{GraphTestHarness.Endpoint}}/page2"}
                """),
            FakeResponse.Ok("""{"value":[{"id":"2","name":"b"}]}"""));

        using var cancellation = new CancellationTokenSource();
        var client = GraphTestHarness.CreateClient(handler);
        var items = new List<Item>();

        await Assert.ThrowsAnyAsync<OperationCanceledException>(async () =>
        {
            await foreach (var item in client.EnumeratePagedAsync<Item>("/x", cancellation.Token))
            {
                items.Add(item);
                await cancellation.CancelAsync();
            }
        });

        Assert.Single(items);
    }

    [Fact]
    public async Task Cancellation_BeforeARequest_MakesNoRequest()
    {
        var handler = new FakeGraphHandler().Enqueue(FakeResponse.Ok("{}"));

        using var cancellation = new CancellationTokenSource();
        await cancellation.CancelAsync();

        await Assert.ThrowsAnyAsync<OperationCanceledException>(() =>
            GraphTestHarness.CreateClient(handler).GetAsync<Item>("/me", cancellation.Token));

        Assert.Equal(0, handler.RequestCount);
    }

    /// <summary>A nextLink is absolute; a relative path is resolved against the configured endpoint.</summary>
    [Fact]
    public async Task AbsoluteAndRelativeUrls_AreBothHandled()
    {
        var handler = new FakeGraphHandler().Enqueue(
            FakeResponse.Ok("""{"id":"1","name":"a"}"""),
            FakeResponse.Ok("""{"id":"2","name":"b"}"""));

        var client = GraphTestHarness.CreateClient(handler);

        await client.GetAsync<Item>("/me");
        await client.GetAsync<Item>("https://graph.example.test/v1.0/other");

        Assert.Equal($"{GraphTestHarness.Endpoint}/me", handler.Requests[0].Uri.ToString());
        Assert.Equal($"{GraphTestHarness.Endpoint}/other", handler.Requests[1].Uri.ToString());
    }

    /// <summary>A malformed error body must not break error reporting.</summary>
    [Fact]
    public async Task MalformedErrorBody_StillProducesANormalizedError()
    {
        var handler = new FakeGraphHandler()
            .Enqueue(new FakeResponse { Status = HttpStatusCode.BadRequest, Json = "this is not json" });

        var response = await GraphTestHarness.CreateClient(handler).GetAsync<Item>("/me");

        Assert.False(response.Succeeded);
        Assert.NotNull(response.Error);
        Assert.Equal(400, response.Error!.StatusCode);
    }

    /// <summary>The correlation ID reaches the error so support can trace the exact request.</summary>
    [Fact]
    public async Task ErrorsCarryCorrelationAndServiceRequestIds()
    {
        var handler = new FakeGraphHandler()
            .Enqueue(FakeResponse.Error(HttpStatusCode.Forbidden, "accessDenied"));

        var response = await GraphTestHarness.CreateClient(handler).GetAsync<Item>("/sites/s1");

        Assert.False(string.IsNullOrWhiteSpace(response.Error!.ClientRequestId));
        Assert.Equal("fake-request-id", response.Error.ServiceRequestId);
    }
}
