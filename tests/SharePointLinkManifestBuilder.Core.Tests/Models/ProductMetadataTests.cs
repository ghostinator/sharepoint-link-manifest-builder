using SharePointLinkManifestBuilder.Core.Models;

namespace SharePointLinkManifestBuilder.Core.Tests.Models;

/// <summary>
/// The publisher metadata shipped with this build. These pin the distinction between "nobody has
/// set this" and "this is set to something real", because the About page shows a warning based
/// on it and a warning that is wrong in either direction is worse than none.
/// </summary>
public sealed class ProductMetadataTests
{
    private static readonly ProductMetadata Defaults = new();

    /// <summary>Publisher identity is set, so the About warning must not appear.</summary>
    [Fact]
    public void PublisherIdentity_IsNoLongerAPlaceholder()
    {
        Assert.False(Defaults.HasPlaceholders);
        Assert.DoesNotContain("PLACEHOLDER", Defaults.Publisher, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("PLACEHOLDER", Defaults.ContactAddress, StringComparison.OrdinalIgnoreCase);
    }

    /// <summary>
    /// The update endpoint is deliberately unset until a release exists. It is excluded from
    /// HasPlaceholders, and the About page reports it separately.
    /// </summary>
    [Fact]
    public void UpdateEndpoint_IsStillUnsetAndDoesNotTriggerThePublisherWarning()
    {
        Assert.Contains("PLACEHOLDER", Defaults.UpdateCheckUrl, StringComparison.OrdinalIgnoreCase);
        Assert.False(Defaults.HasPlaceholders);
    }

    /// <summary>Every published URL must be absolute and https, or it cannot be opened.</summary>
    [Fact]
    public void PublishedUrls_AreAbsoluteAndSecure()
    {
        foreach (var url in new[]
        {
            Defaults.HomepageUrl, Defaults.SupportUrl, Defaults.PrivacyPolicyUrl,
            Defaults.TermsUrl, Defaults.SourceCodeUrl, Defaults.IssueTrackerUrl,
        })
        {
            Assert.True(Uri.TryCreate(url, UriKind.Absolute, out var uri), url);
            Assert.Equal(Uri.UriSchemeHttps, uri!.Scheme);
        }
    }

    /// <summary>
    /// The privacy and terms URLs must point at documents that exist in this repository rather
    /// than pages nobody has written, since the consent screen shows the privacy URL.
    /// </summary>
    [Fact]
    public void PrivacyAndTerms_PointAtDocumentsThatExist()
    {
        Assert.EndsWith("docs/PRIVACY.md", Defaults.PrivacyPolicyUrl, StringComparison.Ordinal);
        Assert.EndsWith("LICENSE", Defaults.TermsUrl, StringComparison.Ordinal);
    }

    /// <summary>A contact address has to be an address, not a sentence.</summary>
    [Fact]
    public void ContactAddress_LooksLikeAnAddress()
    {
        Assert.Contains("@", Defaults.ContactAddress, StringComparison.Ordinal);
        Assert.DoesNotContain("example.invalid", Defaults.ContactAddress, StringComparison.OrdinalIgnoreCase);
    }
}
