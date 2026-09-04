using System.Net;
using Microsoft.Extensions.Logging.Abstractions;
using SharePointLinkManifestBuilder.Core.Models;
using SharePointLinkManifestBuilder.Graph.Services;

namespace SharePointLinkManifestBuilder.Graph.Tests;

/// <summary>Site discovery, resolution and library listing.</summary>
public sealed class SiteServiceTests
{
    private static SiteService Create(FakeGraphHandler handler) =>
        new(GraphTestHarness.CreateClient(handler), NullLogger<SiteService>.Instance);

    [Fact]
    public async Task SearchSites_ReturnsMappedSites()
    {
        var handler = new FakeGraphHandler().Enqueue(FakeResponse.Ok("""
            {"value":[
              {"id":"host,sc1,web1","displayName":"Marketing","name":"Marketing",
               "webUrl":"https://example.sharepoint.test/sites/Marketing","description":"Marketing team"},
              {"id":"host,sc2,web2","name":"Engineering",
               "webUrl":"https://example.sharepoint.test/sites/Engineering"}
            ]}
            """));

        var result = await Create(handler).SearchSitesAsync("mark");

        Assert.True(result.Succeeded);
        Assert.Equal(2, result.Value!.Count);
        Assert.Equal("Marketing", result.Value[0].DisplayName);
        Assert.Equal("example.sharepoint.test", result.Value[0].Hostname);

        // Falls back to the name when no display name is returned, rather than showing an ID.
        Assert.Equal("Engineering", result.Value[1].DisplayName);
    }

    [Fact]
    public async Task SearchSites_EmptyQuery_MakesNoRequest()
    {
        var handler = new FakeGraphHandler();

        var result = await Create(handler).SearchSitesAsync("   ");

        Assert.True(result.Succeeded);
        Assert.Empty(result.Value!);
        Assert.Equal(0, handler.RequestCount);
    }

    /// <summary>A site with no usable ID cannot be addressed, so it is dropped rather than shown.</summary>
    [Fact]
    public async Task SearchSites_SkipsSitesWithNoIdentifier()
    {
        var handler = new FakeGraphHandler()
            .Enqueue(FakeResponse.Ok("""{"value":[{"displayName":"Broken"},{"id":"h,s,w","displayName":"Good"}]}"""));

        var result = await Create(handler).SearchSitesAsync("x");

        Assert.Equal("Good", Assert.Single(result.Value!).DisplayName);
    }

    [Fact]
    public async Task SearchSites_Forbidden_ReportsTheError()
    {
        var handler = new FakeGraphHandler()
            .Enqueue(FakeResponse.Error(HttpStatusCode.Forbidden, "accessDenied", "Access denied"));

        var result = await Create(handler).SearchSitesAsync("x");

        Assert.False(result.Succeeded);
        Assert.Equal(GraphErrorKind.SharePointAccessDenied, result.Error!.Kind);
    }

    /// <summary>A pasted URL resolves through the hostname-and-path form, not by fetching it.</summary>
    [Fact]
    public async Task ResolveSiteByUrl_UsesTheHostnameAndPathForm()
    {
        var handler = new FakeGraphHandler()
            .Enqueue(FakeResponse.Ok("""{"id":"host,sc1,web1","displayName":"Marketing"}"""));

        var result = await Create(handler)
            .ResolveSiteByUrlAsync("https://example.sharepoint.com/sites/Marketing");

        Assert.True(result.Succeeded);

        var requested = Assert.Single(handler.Requests).Uri.ToString();
        Assert.Contains("/sites/example.sharepoint.com", requested, StringComparison.Ordinal);
        Assert.Contains("Marketing", requested, StringComparison.Ordinal);
    }

    [Fact]
    public async Task ResolveSiteByUrl_InvalidUrl_FailsWithoutARequest()
    {
        var handler = new FakeGraphHandler();

        var result = await Create(handler).ResolveSiteByUrlAsync("https://evil.example.com/sites/A");

        Assert.False(result.Succeeded);
        Assert.Equal(GraphErrorKind.InvalidUrl, result.Error!.Kind);
        Assert.Equal(0, handler.RequestCount);
    }

    /// <summary>A sharing link names one item, not a site, and must be rejected with an explanation.</summary>
    [Fact]
    public async Task ResolveSiteByUrl_SharingLink_IsRejectedWithGuidance()
    {
        var handler = new FakeGraphHandler();

        var result = await Create(handler)
            .ResolveSiteByUrlAsync("https://example.sharepoint.com/:f:/s/Marketing/EQ123");

        Assert.False(result.Succeeded);
        Assert.Contains("sharing link", result.Error!.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Equal(0, handler.RequestCount);
    }

    /// <summary>A site with several libraries must surface all of them.</summary>
    [Fact]
    public async Task GetSiteDrives_ReturnsEveryLibrary()
    {
        var handler = new FakeGraphHandler().Enqueue(FakeResponse.Ok("""
            {"value":[
              {"id":"d1","name":"Documents","driveType":"documentLibrary",
               "webUrl":"https://example.sharepoint.test/sites/A/Shared%20Documents"},
              {"id":"d2","name":"Policies","driveType":"documentLibrary"},
              {"id":"d3","name":"Archive","driveType":"documentLibrary"}
            ]}
            """));

        var result = await Create(handler).GetSiteDrivesAsync("host,sc1,web1");

        Assert.Equal(3, result.Value!.Count);
        Assert.Equal(["Documents", "Policies", "Archive"], result.Value.Select(d => d.Name));
        Assert.All(result.Value, d => Assert.Equal("host,sc1,web1", d.SiteId));
    }

    [Fact]
    public async Task GetSiteDrives_PaginatesAcrossPages()
    {
        var handler = new FakeGraphHandler().Enqueue(
            FakeResponse.Ok($$"""
                {"value":[{"id":"d1","name":"Documents"}],
                 "@odata.nextLink":"{{GraphTestHarness.Endpoint}}/next"}
                """),
            FakeResponse.Ok("""{"value":[{"id":"d2","name":"Policies"}]}"""));

        var result = await Create(handler).GetSiteDrivesAsync("host,sc1,web1");

        Assert.Equal(2, result.Value!.Count);
        Assert.Equal(2, handler.RequestCount);
    }
}

/// <summary>Drive resolution and folder enumeration for SharePoint and OneDrive.</summary>
public sealed class DriveServiceTests
{
    private static DriveService Create(FakeGraphHandler handler) =>
        new(GraphTestHarness.CreateClient(handler), NullLogger<DriveService>.Instance);

    [Fact]
    public async Task GetMyDrive_ReturnsTheSignedInUsersDrive()
    {
        var handler = new FakeGraphHandler().Enqueue(FakeResponse.Ok("""
            {"id":"d-me","name":"OneDrive","driveType":"business",
             "webUrl":"https://example-my.sharepoint.test/personal/test",
             "owner":{"user":{"id":"u1","displayName":"Test User"}}}
            """));

        var result = await Create(handler).GetMyDriveAsync();

        Assert.True(result.Succeeded);
        Assert.Equal("d-me", result.Value!.DriveId);
        Assert.Equal("Test User", result.Value.OwnerDisplayName);
        Assert.True(result.Value.IsPersonal);
    }

    /// <summary>
    /// A user who has never opened OneDrive does not have one. That is a distinct, explainable
    /// outcome, not a generic access error, and the application must never provision it.
    /// </summary>
    [Fact]
    public async Task GetUserDrive_NotProvisioned_IsReportedSpecifically()
    {
        var handler = new FakeGraphHandler()
            .Enqueue(FakeResponse.Error(HttpStatusCode.NotFound, "itemNotFound", "Drive not found"));

        var result = await Create(handler).GetUserDriveAsync("user-2");

        Assert.False(result.Succeeded);
        Assert.Equal(GraphErrorKind.UserDriveUnprovisioned, result.Error!.Kind);
        Assert.Contains("does not provision", result.Error.SuggestedAction, StringComparison.OrdinalIgnoreCase);
    }

    /// <summary>Administrator consent does not grant access to everyone's OneDrive; say so.</summary>
    [Fact]
    public async Task GetUserDrive_AccessDenied_ExplainsThatConsentIsNotEnough()
    {
        var handler = new FakeGraphHandler()
            .Enqueue(FakeResponse.Error(HttpStatusCode.Forbidden, "accessDenied", "Access denied"));

        var result = await Create(handler).GetUserDriveAsync("user-2");

        Assert.Equal(GraphErrorKind.OneDriveAccessDenied, result.Error!.Kind);
        Assert.Contains(
            "Administrator consent alone",
            result.Error.SuggestedAction,
            StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task GetChildren_ClassifiesFilesFoldersPackagesAndShortcuts()
    {
        var handler = new FakeGraphHandler().Enqueue(FakeResponse.Ok("""
            {"value":[
              {"id":"f1","name":"Report.docx","size":1024,"file":{"mimeType":"application/msword"},
               "parentReference":{"driveId":"d1","id":"root","path":"/drive/root:"}},
              {"id":"f2","name":"Reports","folder":{"childCount":3},
               "parentReference":{"driveId":"d1","id":"root","path":"/drive/root:"}},
              {"id":"f3","name":"Notebook","package":{"type":"oneNote"},"folder":{"childCount":2},
               "parentReference":{"driveId":"d1","id":"root","path":"/drive/root:"}},
              {"id":"f4","name":"Shortcut","remoteItem":{"id":"r1","name":"Elsewhere"},
               "parentReference":{"driveId":"d1","id":"root","path":"/drive/root:"}}
            ]}
            """));

        var items = new List<DiscoveredFile>();

        await foreach (var item in Create(handler).GetChildrenAsync("d1", "root"))
        {
            items.Add(item);
        }

        Assert.Equal(4, items.Count);
        Assert.Equal(DriveItemKind.File, items[0].Kind);
        Assert.Equal(DriveItemKind.Folder, items[1].Kind);

        // A OneNote notebook carries both a folder facet and a package facet. Classifying it as
        // a folder would make a recursive job try to descend into something that is not one.
        Assert.Equal(DriveItemKind.Package, items[2].Kind);
        Assert.Equal(DriveItemKind.RemoteItem, items[3].Kind);
    }

    [Fact]
    public async Task GetChildren_EmptyFolder_YieldsNothing()
    {
        var handler = new FakeGraphHandler().Enqueue(FakeResponse.Ok("""{"value":[]}"""));

        var items = new List<DiscoveredFile>();

        await foreach (var item in Create(handler).GetChildrenAsync("d1", "root"))
        {
            items.Add(item);
        }

        Assert.Empty(items);
    }

    /// <summary>The drive-relative path is derived from the parent reference, not guessed.</summary>
    [Theory]
    [InlineData("/drive/root:", "")]
    [InlineData("/drive/root:/Reports", "Reports")]
    [InlineData("/drive/root:/Reports/Q1", "Reports/Q1")]
    [InlineData("/drives/d1/root:/Reports%20Archive", "Reports Archive")]
    public async Task RelativePath_IsDerivedFromTheParentReference(string parentPath, string expected)
    {
        // Concatenated rather than interpolated: the JSON's trailing "}}]}" cannot follow an
        // interpolation hole in a raw literal without escalating the "$" count.
        var json = "{\"value\":[{\"id\":\"f1\",\"name\":\"Report.docx\",\"file\":{},"
            + "\"parentReference\":{\"driveId\":\"d1\",\"id\":\"p1\",\"path\":\""
            + parentPath + "\"}}]}";

        var handler = new FakeGraphHandler().Enqueue(FakeResponse.Ok(json));

        var items = new List<DiscoveredFile>();

        await foreach (var item in Create(handler).GetChildrenAsync("d1", "p1"))
        {
            items.Add(item);
        }

        Assert.Equal(expected, Assert.Single(items).ParentRelativePath);
    }

    [Fact]
    public async Task GetSubfolders_ReturnsOnlyFolders()
    {
        var handler = new FakeGraphHandler().Enqueue(FakeResponse.Ok("""
            {"value":[
              {"id":"f1","name":"Report.docx","file":{}},
              {"id":"f2","name":"Reports","folder":{"childCount":3}},
              {"id":"f3","name":"Archive","folder":{"childCount":0}}
            ]}
            """));

        var result = await Create(handler).GetSubfoldersAsync("d1", "root");

        Assert.Equal(2, result.Value!.Count);
        Assert.Equal(["Reports", "Archive"], result.Value.Select(f => f.Name));
        Assert.Equal(3, result.Value[0].ChildCount);
    }

    [Fact]
    public async Task GetFolderByPath_WhenTheTargetIsAFile_ExplainsTheMistake()
    {
        var handler = new FakeGraphHandler()
            .Enqueue(FakeResponse.Ok("""{"id":"f1","name":"Report.docx","file":{}}"""));

        var result = await Create(handler).GetFolderByPathAsync("d1", "Reports/Report.docx");

        Assert.False(result.Succeeded);
        Assert.Equal(GraphErrorKind.FolderNotFound, result.Error!.Kind);
        Assert.Contains("is a file, not a folder", result.Error.Message, StringComparison.Ordinal);
    }

    /// <summary>A pasted sharing URL is resolved through the shares endpoint with a u! token.</summary>
    [Fact]
    public async Task ResolveSharingUrl_UsesTheSharesEndpoint()
    {
        var handler = new FakeGraphHandler().Enqueue(FakeResponse.Ok("""
            {"id":"i1","name":"Shared.docx","file":{},
             "parentReference":{"driveId":"d9","id":"p1","path":"/drive/root:/Shared"}}
            """));

        var result = await Create(handler)
            .ResolveSharingUrlAsync("https://example.sharepoint.com/:w:/s/Marketing/EQ123");

        Assert.True(result.Succeeded);
        Assert.Equal("d9", result.Value!.DriveId);

        var requested = Assert.Single(handler.Requests).Uri.ToString();
        Assert.Contains("/shares/u!", requested, StringComparison.Ordinal);
    }

    [Fact]
    public async Task ResolveSharingUrl_WithNoDriveInTheResponse_Fails()
    {
        var handler = new FakeGraphHandler().Enqueue(FakeResponse.Ok("""{"id":"i1","name":"x","file":{}}"""));

        var result = await Create(handler)
            .ResolveSharingUrlAsync("https://example.sharepoint.com/:w:/s/M/EQ1");

        Assert.False(result.Succeeded);
        Assert.Equal(GraphErrorKind.InvalidUrl, result.Error!.Kind);
    }
}

/// <summary>The people picker backing the User OneDrive source.</summary>
public sealed class UserDirectoryServiceTests
{
    private static UserDirectoryService Create(FakeGraphHandler handler) =>
        new(GraphTestHarness.CreateClient(handler), NullLogger<UserDirectoryService>.Instance);

    [Fact]
    public async Task SearchUsers_ReturnsOnlyWhatGraphReturned()
    {
        var handler = new FakeGraphHandler().Enqueue(FakeResponse.Ok("""
            {"value":[
              {"id":"u1","displayName":"Ada Example","userPrincipalName":"ada@example.test","jobTitle":"Engineer"},
              {"id":"u2","displayName":"Grace Example"}
            ]}
            """));

        var result = await Create(handler).SearchUsersAsync("exam");

        Assert.Equal(2, result.Value!.Count);
        Assert.Equal("ada@example.test", result.Value[0].UserPrincipalName);

        // Nothing is invented for fields the directory withheld.
        Assert.Null(result.Value[1].UserPrincipalName);
        Assert.Null(result.Value[1].JobTitle);
    }

    /// <summary>A one-character query would return an arbitrary slice of the directory.</summary>
    [Theory]
    [InlineData("a")]
    [InlineData("")]
    [InlineData("   ")]
    public async Task SearchUsers_TooShort_MakesNoRequest(string query)
    {
        var handler = new FakeGraphHandler();

        var result = await Create(handler).SearchUsersAsync(query);

        Assert.True(result.Succeeded);
        Assert.Empty(result.Value!);
        Assert.Equal(0, handler.RequestCount);
    }

    /// <summary>A quote in a search term must not break the OData filter.</summary>
    [Fact]
    public async Task SearchUsers_EscapesQuotesInTheFilter()
    {
        var handler = new FakeGraphHandler().Enqueue(FakeResponse.Ok("""{"value":[]}"""));

        await Create(handler).SearchUsersAsync("O'Brien");

        var url = Uri.UnescapeDataString(Assert.Single(handler.Requests).Uri.ToString());
        Assert.Contains("O''Brien", url, StringComparison.Ordinal);
    }
}
