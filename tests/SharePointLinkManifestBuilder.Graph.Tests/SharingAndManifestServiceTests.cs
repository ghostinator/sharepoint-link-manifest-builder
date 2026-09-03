using System.Net;
using Microsoft.Extensions.Logging.Abstractions;
using SharePointLinkManifestBuilder.Core.Models;
using SharePointLinkManifestBuilder.Graph.Services;

namespace SharePointLinkManifestBuilder.Graph.Tests;

/// <summary>
/// Covers the two Graph behaviours this product's correctness hinges on: createLink's
/// 201-versus-200 distinction, and the fact that named recipients require the separate invite
/// action, which can partially fail.
/// </summary>
public sealed class SharingLinkServiceTests
{
    private static DiscoveredFile File => new()
    {
        DriveId = "drive-1",
        ItemId = "item-1",
        Name = "Quarterly Report.docx",
        RelativePath = "Reports/Quarterly Report.docx",
    };

    private static SharingLinkService Create(FakeGraphHandler handler) =>
        new(GraphTestHarness.CreateClient(handler), NullLogger<SharingLinkService>.Instance);

    private const string OrganizationViewLink = """
        {"id":"perm-1","roles":["read"],
         "link":{"type":"view","scope":"organization",
                 "webUrl":"https://example.sharepoint.test/:w:/s/A/EXAMPLE"}}
        """;

    /// <summary>HTTP 201 means Graph made a new link.</summary>
    [Fact]
    public async Task Created_Http201_IsRecordedAsCreated()
    {
        var handler = new FakeGraphHandler().Enqueue(FakeResponse.Created(OrganizationViewLink));

        var result = await Create(handler).CreateOrGetLinkAsync(File, new LinkConfiguration());

        Assert.Equal(LinkResultStatus.Created, result.Status);
        Assert.Equal("https://example.sharepoint.test/:w:/s/A/EXAMPLE", result.SharingUrl);
        Assert.True(result.IsSuccess);
    }

    /// <summary>
    /// HTTP 200 means an equivalent link already existed. Reporting that as "Created" would
    /// tell an administrator this run made links it did not make.
    /// </summary>
    [Fact]
    public async Task Existing_Http200_IsRecordedAsReusedNotCreated()
    {
        var handler = new FakeGraphHandler().Enqueue(FakeResponse.Ok(OrganizationViewLink));

        var result = await Create(handler).CreateOrGetLinkAsync(File, new LinkConfiguration());

        Assert.Equal(LinkResultStatus.Reused, result.Status);
        Assert.NotEqual(LinkResultStatus.Created, result.Status);
        Assert.True(result.IsSuccess);
        Assert.Equal("Reused", result.ManifestStatus);
    }

    [Fact]
    public async Task RequestBody_CarriesTheRequestedTypeAndScope()
    {
        var handler = new FakeGraphHandler().Enqueue(FakeResponse.Created(OrganizationViewLink));

        await Create(handler).CreateOrGetLinkAsync(
            File,
            new LinkConfiguration { Permission = LinkPermission.Edit, Audience = LinkAudience.Anyone });

        var body = Assert.Single(handler.Requests).Body;

        Assert.Contains("\"Type\":\"edit\"", body, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("\"Scope\":\"anonymous\"", body, StringComparison.OrdinalIgnoreCase);
    }

    /// <summary>The default matches Graph's, so the request stays minimal.</summary>
    [Fact]
    public async Task RetainInheritedPermissions_IsOmittedWhenItMatchesTheGraphDefault()
    {
        var handler = new FakeGraphHandler().Enqueue(FakeResponse.Created(OrganizationViewLink));

        await Create(handler).CreateOrGetLinkAsync(
            File, new LinkConfiguration { RetainInheritedPermissions = true });

        Assert.DoesNotContain(
            "retainInheritedPermissions",
            Assert.Single(handler.Requests).Body,
            StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task RetainInheritedPermissions_IsSentWhenItDiffersFromTheDefault()
    {
        var handler = new FakeGraphHandler().Enqueue(FakeResponse.Created(OrganizationViewLink));

        await Create(handler).CreateOrGetLinkAsync(
            File, new LinkConfiguration { RetainInheritedPermissions = false });

        Assert.Contains(
            "RetainInheritedPermissions",
            Assert.Single(handler.Requests).Body,
            StringComparison.OrdinalIgnoreCase);
    }

    /// <summary>A policy refusal is its own outcome, not a generic failure.</summary>
    [Fact]
    public async Task AnonymousSharingDisabled_IsRecordedAsPolicyBlocked()
    {
        var handler = new FakeGraphHandler().Enqueue(FakeResponse.Error(
            HttpStatusCode.Forbidden,
            "notAllowed",
            "Anonymous sharing is disabled for this organization"));

        var result = await Create(handler).CreateOrGetLinkAsync(
            File, new LinkConfiguration { Audience = LinkAudience.Anyone });

        Assert.Equal(LinkResultStatus.PolicyBlocked, result.Status);
        Assert.False(result.IsSuccess);
        Assert.NotNull(result.Error!.SuggestedAction);
    }

    [Fact]
    public async Task AccessDenied_IsRecordedAsAccessDenied()
    {
        var handler = new FakeGraphHandler()
            .Enqueue(FakeResponse.Error(HttpStatusCode.Forbidden, "accessDenied", "Access denied"));

        var result = await Create(handler).CreateOrGetLinkAsync(File, new LinkConfiguration());

        Assert.Equal(LinkResultStatus.AccessDenied, result.Status);
    }

    [Fact]
    public async Task DeletedDuringProcessing_IsRecordedAsFailedWithTheRightKind()
    {
        var handler = new FakeGraphHandler()
            .Enqueue(FakeResponse.Error(HttpStatusCode.NotFound, "itemNotFound", "Item not found"));

        var result = await Create(handler).CreateOrGetLinkAsync(File, new LinkConfiguration());

        Assert.Equal(LinkResultStatus.Failed, result.Status);
        Assert.Equal(GraphErrorKind.FileDeletedDuringProcessing, result.Error!.Kind);
    }

    /// <summary>Skip-if-exists avoids the write entirely when a matching link is already present.</summary>
    [Fact]
    public async Task SkipWhenEquivalentExists_ReturnsExistingWithoutWriting()
    {
        var handler = new FakeGraphHandler().Enqueue(FakeResponse.Ok("""
            {"value":[{"id":"perm-9","link":{"type":"view","scope":"organization",
                       "webUrl":"https://example.sharepoint.test/:w:/s/A/OLD"}}]}
            """));

        var result = await Create(handler).CreateOrGetLinkAsync(
            File, new LinkConfiguration { SkipWhenEquivalentLinkExists = true });

        Assert.Equal(LinkResultStatus.Existing, result.Status);
        Assert.Equal("https://example.sharepoint.test/:w:/s/A/OLD", result.SharingUrl);

        // One read, and crucially no POST.
        Assert.Equal(1, handler.RequestCount);
        Assert.Equal("GET", handler.Requests[0].Method);
    }

    [Fact]
    public async Task SkipWhenEquivalentExists_WithNoMatch_StillCreatesTheLink()
    {
        var handler = new FakeGraphHandler().Enqueue(
            FakeResponse.Ok("""
                {"value":[{"id":"perm-9","link":{"type":"edit","scope":"anonymous","webUrl":"https://x.test/1"}}]}
                """),
            FakeResponse.Created(OrganizationViewLink));

        var result = await Create(handler).CreateOrGetLinkAsync(
            File, new LinkConfiguration { SkipWhenEquivalentLinkExists = true });

        Assert.Equal(LinkResultStatus.Created, result.Status);
        Assert.Equal(2, handler.RequestCount);
    }

    /// <summary>
    /// Specific-people sharing needs two calls, because the v1.0 createLink action has no
    /// recipients parameter.
    /// </summary>
    [Fact]
    public async Task SpecificPeople_UsesCreateLinkThenInvite()
    {
        var handler = new FakeGraphHandler().Enqueue(
            FakeResponse.Created("""
                {"id":"perm-1","link":{"type":"view","scope":"users","webUrl":"https://x.test/u"}}
                """),
            FakeResponse.Ok("""
                {"value":[{"id":"p1","invitation":{"email":"a@example.test","signInRequired":true},
                           "roles":["read"]}]}
                """));

        var result = await Create(handler).CreateOrGetLinkAsync(
            File,
            new LinkConfiguration
            {
                Audience = LinkAudience.SpecificPeople,
                Recipients = ["a@example.test"],
            });

        Assert.Equal(LinkResultStatus.Created, result.Status);
        Assert.Equal(2, handler.RequestCount);
        Assert.Contains("createLink", handler.Requests[0].Uri.ToString(), StringComparison.Ordinal);
        Assert.Contains("invite", handler.Requests[1].Uri.ToString(), StringComparison.Ordinal);

        var recipient = Assert.Single(result.RecipientResults);
        Assert.True(recipient.Succeeded);
        Assert.Equal("a@example.test", recipient.Recipient);
    }

    /// <summary>No email may be sent unless the user explicitly asked for one.</summary>
    [Fact]
    public async Task Invite_DoesNotSendAnEmailByDefault()
    {
        var handler = new FakeGraphHandler()
            .Enqueue(FakeResponse.Ok("""{"value":[{"id":"p1","roles":["read"]}]}"""));

        await Create(handler).InviteRecipientsAsync(
            File,
            new LinkConfiguration
            {
                Audience = LinkAudience.SpecificPeople,
                Recipients = ["a@example.test"],
            });

        Assert.Contains(
            "\"SendInvitation\":false",
            Assert.Single(handler.Requests).Body,
            StringComparison.OrdinalIgnoreCase);
    }

    /// <summary>
    /// Graph returns 207 when some recipients succeed and others fail. Each recipient's
    /// outcome must be reported individually rather than failing the whole file.
    /// </summary>
    [Fact]
    public async Task Invite_MultiStatus_ReportsEachRecipientIndividually()
    {
        var handler = new FakeGraphHandler().Enqueue(new FakeResponse
        {
            Status = HttpStatusCode.MultiStatus,
            Json = """
                {"value":[
                  {"id":"p1","invitation":{"email":"good@example.test"},"roles":["read"]},
                  {"id":"p2","invitation":{"email":"bad@example.test"},"roles":["read"],
                   "error":{"code":"notAllowed","message":"Recipient rejected"}}
                ]}
                """,
        });

        var results = await Create(handler).InviteRecipientsAsync(
            File,
            new LinkConfiguration
            {
                Audience = LinkAudience.SpecificPeople,
                Recipients = ["good@example.test", "bad@example.test"],
            });

        Assert.Equal(2, results.Count);

        var good = Assert.Single(results, r => r.Recipient == "good@example.test");
        Assert.True(good.Succeeded);

        var bad = Assert.Single(results, r => r.Recipient == "bad@example.test");
        Assert.False(bad.Succeeded);
        Assert.Equal(GraphErrorKind.RecipientRejected, bad.Error!.Kind);
    }

    /// <summary>A whole-call failure still reports one outcome per recipient, for a consistent grid.</summary>
    [Fact]
    public async Task Invite_WholeCallFailure_ReportsEveryRecipientAsFailed()
    {
        var handler = new FakeGraphHandler()
            .Enqueue(FakeResponse.Error(HttpStatusCode.Forbidden, "accessDenied"));

        var results = await Create(handler).InviteRecipientsAsync(
            File,
            new LinkConfiguration
            {
                Audience = LinkAudience.SpecificPeople,
                Recipients = ["a@example.test", "b@example.test"],
            });

        Assert.Equal(2, results.Count);
        Assert.All(results, r => Assert.False(r.Succeeded));
    }

    [Fact]
    public async Task Expiration_IsFormattedAsIso8601Utc()
    {
        var handler = new FakeGraphHandler().Enqueue(FakeResponse.Created(OrganizationViewLink));

        await Create(handler).CreateOrGetLinkAsync(
            File,
            new LinkConfiguration
            {
                ExpirationUtc = new DateTimeOffset(2026, 12, 31, 23, 59, 59, TimeSpan.Zero),
            });

        Assert.Contains(
            "2026-12-31T23:59:59Z",
            Assert.Single(handler.Requests).Body,
            StringComparison.Ordinal);
    }
}

/// <summary>Manifest reading and writing, including ETag concurrency and chunked upload.</summary>
public sealed class ManifestStorageServiceTests
{
    private static ManifestStorageService Create(FakeGraphHandler handler) =>
        new(GraphTestHarness.CreateClient(handler), NullLogger<ManifestStorageService>.Instance);

    /// <summary>A missing manifest is the normal first-run case, not an error.</summary>
    [Fact]
    public async Task ReadManifest_WhenAbsent_ReturnsSuccessWithNull()
    {
        var handler = new FakeGraphHandler()
            .Enqueue(FakeResponse.Error(HttpStatusCode.NotFound, "itemNotFound"));

        var result = await Create(handler).ReadManifestAsync("d1", "folder-1", "_sharepoint-links.txt");

        Assert.True(result.Succeeded);
        Assert.Null(result.Value);
    }

    [Fact]
    public async Task ReadManifest_WhenPresent_ReturnsContentAndETag()
    {
        var content = "SharePoint Link Manifest\nSchema Version: 1.0\n";

        var handler = new FakeGraphHandler().Enqueue(
            new FakeResponse
            {
                Json = """{"id":"m1","webUrl":"https://x.test/m","eTag":"\"etag-1\""}""",
            },
            new FakeResponse { Bytes = System.Text.Encoding.UTF8.GetBytes(content) });

        var result = await Create(handler).ReadManifestAsync("d1", "folder-1", "_sharepoint-links.txt");

        Assert.True(result.Succeeded);
        Assert.Equal("m1", result.Value!.ItemId);
        Assert.Equal(content, result.Value.Content);
        Assert.Equal("\"etag-1\"", result.Value.ETag);
    }

    /// <summary>A byte-order mark another tool wrote must not corrupt the parsed content.</summary>
    [Fact]
    public async Task ReadManifest_ToleratesAByteOrderMark()
    {
        var bytes = new byte[] { 0xEF, 0xBB, 0xBF }
            .Concat(System.Text.Encoding.UTF8.GetBytes("SharePoint Link Manifest\n"))
            .ToArray();

        var handler = new FakeGraphHandler().Enqueue(
            new FakeResponse { Json = """{"id":"m1"}""" },
            new FakeResponse { Bytes = bytes });

        var result = await Create(handler).ReadManifestAsync("d1", "f1", "_sharepoint-links.txt");

        Assert.StartsWith("SharePoint Link Manifest", result.Value!.Content, StringComparison.Ordinal);
    }

    [Fact]
    public async Task WriteManifest_SmallContent_UsesASinglePutWithTheETag()
    {
        var handler = new FakeGraphHandler().Enqueue(FakeResponse.Ok("""{"id":"m1","webUrl":"https://x.test/m"}"""));

        var result = await Create(handler).WriteManifestAsync(
            "d1", "folder-1", "_sharepoint-links.txt", "content",
            ManifestFormats.PlainText, isMaster: false, entryCount: 3, ifMatchETag: "\"etag-1\"");

        Assert.True(result.Succeeded);
        Assert.Equal(3, result.Value!.EntryCount);

        var request = Assert.Single(handler.Requests);
        Assert.Equal("PUT", request.Method);
        Assert.Equal("\"etag-1\"", request.IfMatch);
    }

    /// <summary>
    /// A remote change between read and write must surface as a typed conflict so the
    /// conflict policy can decide, rather than being retried into an overwrite.
    /// </summary>
    [Fact]
    public async Task WriteManifest_ETagConflict_ReportsAManifestConflict()
    {
        var handler = new FakeGraphHandler()
            .Enqueue(FakeResponse.Error(HttpStatusCode.PreconditionFailed, "resourceModified"));

        var result = await Create(handler).WriteManifestAsync(
            "d1", "folder-1", "_sharepoint-links.txt", "content",
            ManifestFormats.PlainText, isMaster: false, entryCount: 1, ifMatchETag: "\"stale\"");

        Assert.False(result.Succeeded);
        Assert.Equal(GraphErrorKind.ManifestConflict, result.Error!.Kind);
        Assert.True(result.Error.IsRetryable);
        Assert.Equal(1, handler.RequestCount);
    }

    /// <summary>Content over the simple-upload limit must go through an upload session.</summary>
    [Fact]
    public async Task WriteManifest_LargeContent_UsesAnUploadSessionInChunks()
    {
        var large = new string('x', ManifestStorageService.SimpleUploadLimitBytes + 5_000);

        var handler = new FakeGraphHandler().Enqueue(
            FakeResponse.Ok("""{"uploadUrl":"https://upload.example.test/session/abc"}"""),
            FakeResponse.Ok("{}"),
            FakeResponse.Ok("{}"));

        var result = await Create(handler).WriteManifestAsync(
            "d1", "folder-1", "_sharepoint-links-master.txt", large,
            ManifestFormats.PlainText, isMaster: true, entryCount: 50_000, ifMatchETag: null);

        Assert.True(result.Succeeded);

        // One createUploadSession plus one PUT per chunk.
        Assert.Contains("createUploadSession", handler.Requests[0].Uri.ToString(), StringComparison.Ordinal);
        Assert.True(handler.RequestCount >= 2);
        Assert.All(handler.Requests.Skip(1), r => Assert.Equal("PUT", r.Method));
    }

    [Theory]
    [InlineData("_sharepoint-links.txt", "text/plain; charset=utf-8")]
    [InlineData("_sharepoint-links.json", "application/json; charset=utf-8")]
    [InlineData("_sharepoint-links.csv", "text/csv; charset=utf-8")]
    [InlineData("_sharepoint-links.md", "text/markdown; charset=utf-8")]
    public void ContentType_MatchesTheManifestFormat(string fileName, string expected) =>
        Assert.Equal(expected, ManifestStorageService.ContentTypeFor(fileName));

    /// <summary>Manifests are written as UTF-8 without a byte-order mark.</summary>
    [Fact]
    public async Task WriteManifest_DoesNotEmitAByteOrderMark()
    {
        var handler = new FakeGraphHandler().Enqueue(FakeResponse.Ok("""{"id":"m1"}"""));

        await Create(handler).WriteManifestAsync(
            "d1", "f1", "_sharepoint-links.txt", "SharePoint Link Manifest",
            ManifestFormats.PlainText, false, 0, null);

        var body = Assert.Single(handler.Requests).Body!;
        Assert.StartsWith("SharePoint", body, StringComparison.Ordinal);
        Assert.DoesNotContain('\uFEFF', body);
    }
}
