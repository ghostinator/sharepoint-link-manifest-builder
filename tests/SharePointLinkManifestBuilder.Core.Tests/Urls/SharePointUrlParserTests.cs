using SharePointLinkManifestBuilder.Core.Urls;

namespace SharePointLinkManifestBuilder.Core.Tests.Urls;

public class SharePointUrlParserTests
{
    [Fact]
    public void Parse_RootSiteUrl_ReturnsRootSite()
    {
        var result = SharePointUrlParser.Parse("https://example.sharepoint.com");

        Assert.Equal(ResourceUrlKind.RootSite, result.Kind);
        Assert.Equal("example.sharepoint.com", result.Hostname);
        Assert.Equal(string.Empty, result.SitePath);
        Assert.Equal("/sites/example.sharepoint.com", result.ToGraphSitePath());
    }

    [Theory]
    [InlineData("https://example.sharepoint.com/sites/Marketing", "/sites/Marketing")]
    [InlineData("https://example.sharepoint.com/sites/Marketing/", "/sites/Marketing")]
    [InlineData("https://example.sharepoint.com/teams/Engineering", "/teams/Engineering")]
    public void Parse_SiteUrl_ExtractsSitePath(string url, string expectedSitePath)
    {
        var result = SharePointUrlParser.Parse(url);

        Assert.Equal(ResourceUrlKind.Site, result.Kind);
        Assert.Equal(expectedSitePath, result.SitePath);
        Assert.Null(result.ServerRelativeItemPath);
    }

    [Fact]
    public void Parse_DocumentPath_SeparatesSiteFromItemPath()
    {
        var result = SharePointUrlParser.Parse(
            "https://example.sharepoint.com/sites/Marketing/Shared%20Documents/Reports/2026");

        Assert.Equal(ResourceUrlKind.DocumentPath, result.Kind);
        Assert.Equal("/sites/Marketing", result.SitePath);
        Assert.Equal("/sites/Marketing/Shared Documents/Reports/2026", result.ServerRelativeItemPath);
        Assert.Equal("Shared Documents/Reports/2026", result.SiteRelativeItemPath);
        Assert.Equal("/sites/example.sharepoint.com:/sites/Marketing", result.ToGraphSitePath());
    }

    [Fact]
    public void Parse_PersonalOneDriveUrl_IsRecognisedAsPersonalSite()
    {
        var result = SharePointUrlParser.Parse(
            "https://example-my.sharepoint.com/personal/jane_example_com/Documents/Reports");

        Assert.Equal(ResourceUrlKind.DocumentPath, result.Kind);
        Assert.True(result.IsPersonalSite);
        Assert.Equal("/personal/jane_example_com", result.SitePath);
        Assert.Equal("jane_example_com", result.PersonalSiteSegment);
        Assert.Equal("Documents/Reports", result.SiteRelativeItemPath);
    }

    [Fact]
    public void Parse_PersonalSiteRootUrl_ReturnsPersonalSiteKind()
    {
        var result = SharePointUrlParser.Parse(
            "https://example-my.sharepoint.com/personal/jane_example_com");

        Assert.Equal(ResourceUrlKind.PersonalSite, result.Kind);
        Assert.Equal("jane_example_com", result.PersonalSiteSegment);
    }

    /// <summary>
    /// The SharePoint UI's "Copy link" for a folder produces a _layouts URL whose real target
    /// is in the 'id' query parameter. Resolving such a URL to the site instead of the folder
    /// would silently widen the job's scope to the entire site.
    /// </summary>
    [Fact]
    public void Parse_LayoutsUrlWithIdParameter_UsesTheEncodedFolderPath()
    {
        var result = SharePointUrlParser.Parse(
            "https://example.sharepoint.com/sites/Marketing/_layouts/15/onedrive.aspx"
            + "?id=%2Fsites%2FMarketing%2FShared%20Documents%2FReports%2FQ1"
            + "&viewid=2f8e1234%2D0000%2D0000%2D0000%2D000000000000");

        Assert.Equal(ResourceUrlKind.DocumentPath, result.Kind);
        Assert.Equal("/sites/Marketing", result.SitePath);
        Assert.Equal("Shared Documents/Reports/Q1", result.SiteRelativeItemPath);
    }

    [Fact]
    public void Parse_LayoutsUrlWithRootFolderParameter_UsesTheEncodedFolderPath()
    {
        var result = SharePointUrlParser.Parse(
            "https://example.sharepoint.com/sites/Marketing/Shared%20Documents/Forms/AllItems.aspx"
            + "?RootFolder=%2Fsites%2FMarketing%2FShared%20Documents%2FArchive");

        Assert.Equal(ResourceUrlKind.DocumentPath, result.Kind);
        Assert.Equal("Shared Documents/Archive", result.SiteRelativeItemPath);
    }

    [Fact]
    public void Parse_LayoutsUrlWithoutPathParameter_FallsBackToTheSite()
    {
        var result = SharePointUrlParser.Parse(
            "https://example.sharepoint.com/sites/Marketing/_layouts/15/settings.aspx");

        Assert.Equal(ResourceUrlKind.Site, result.Kind);
        Assert.Equal("/sites/Marketing", result.SitePath);
    }

    [Theory]
    [InlineData("https://example.sharepoint.com/:f:/s/Marketing/EQ12345abcde")]
    [InlineData("https://example.sharepoint.com/:w:/r/sites/Marketing/_layouts/15/Doc.aspx")]
    [InlineData("https://example.sharepoint.com/:x:/g/personal/jane_example_com/EX999")]
    public void Parse_SharingLink_IsRoutedToTheSharesEndpoint(string url)
    {
        var result = SharePointUrlParser.Parse(url);

        Assert.Equal(ResourceUrlKind.SharingLink, result.Kind);
    }

    [Theory]
    [InlineData("", "No URL was supplied.")]
    [InlineData("not a url", "complete URL")]
    [InlineData("ftp://example.sharepoint.com/sites/A", "Only https")]
    [InlineData("http://example.sharepoint.com/sites/A", "Only https")]
    [InlineData("https://evil.example.com/sites/A", "not a recognised SharePoint")]
    public void Parse_InvalidInput_IsRejectedWithAReason(string url, string expectedFragment)
    {
        var result = SharePointUrlParser.Parse(url);

        Assert.Equal(ResourceUrlKind.Invalid, result.Kind);
        Assert.False(result.IsValid);
        Assert.Contains(expectedFragment, result.FailureReason, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Parse_NullInput_DoesNotThrow()
    {
        var result = SharePointUrlParser.Parse(null);

        Assert.Equal(ResourceUrlKind.Invalid, result.Kind);
    }

    [Theory]
    [InlineData("example.sharepoint.com", true)]
    [InlineData("example.sharepoint.us", true)]
    [InlineData("example.sharepoint-mil.us", true)]
    [InlineData("example.sharepoint.cn", true)]
    [InlineData("example.com", false)]
    [InlineData("sharepoint.com.evil.example", false)]
    [InlineData(null, false)]
    public void IsAllowedHost_AcceptsOnlySharePointHosts(string? host, bool expected) =>
        Assert.Equal(expected, SharePointUrlParser.IsAllowedHost(host));

    [Fact]
    public void Parse_RootRelativeDocumentPath_HasNoSitePath()
    {
        var result = SharePointUrlParser.Parse(
            "https://example.sharepoint.com/Shared%20Documents/Policy.docx");

        Assert.Equal(ResourceUrlKind.DocumentPath, result.Kind);
        Assert.Equal(string.Empty, result.SitePath);
        Assert.Equal("Shared Documents/Policy.docx", result.SiteRelativeItemPath);
    }

    [Fact]
    public void Parse_SiteUrlWithMissingName_IsRejected()
    {
        var result = SharePointUrlParser.Parse("https://example.sharepoint.com/sites/");

        Assert.Equal(ResourceUrlKind.Invalid, result.Kind);
    }
}

public class GraphShareTokenEncoderTests
{
    /// <summary>
    /// The documented encoding is "u!" + unpadded base64url of the UTF-8 URL, with '+' and '/'
    /// replaced by '-' and '_'.
    /// </summary>
    [Fact]
    public void Encode_ProducesUnpaddedBase64UrlWithPrefix()
    {
        const string url = "https://example.sharepoint.com/:f:/s/Marketing/EQ12345";

        var token = GraphShareTokenEncoder.Encode(url);

        Assert.StartsWith("u!", token, StringComparison.Ordinal);
        Assert.DoesNotContain('=', token);
        Assert.DoesNotContain('+', token);
        Assert.DoesNotContain('/', token);

        var payload = token[2..].Replace('_', '/').Replace('-', '+');
        var padded = payload.PadRight(payload.Length + ((4 - (payload.Length % 4)) % 4), '=');
        Assert.Equal(url, System.Text.Encoding.UTF8.GetString(Convert.FromBase64String(padded)));
    }

    [Fact]
    public void Encode_TrimsSurroundingWhitespace()
    {
        Assert.Equal(
            GraphShareTokenEncoder.Encode("https://example.sharepoint.com/a"),
            GraphShareTokenEncoder.Encode("  https://example.sharepoint.com/a  "));
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("   ")]
    public void Encode_RejectsEmptyInput(string? url) =>
        Assert.Throws<ArgumentException>(() => GraphShareTokenEncoder.Encode(url!));
}
