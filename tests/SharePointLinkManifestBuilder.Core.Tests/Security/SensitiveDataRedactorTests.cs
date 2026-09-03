using SharePointLinkManifestBuilder.Core.Security;

namespace SharePointLinkManifestBuilder.Core.Tests.Security;

public class SensitiveDataRedactorTests
{
    /// <summary>
    /// A structurally-valid but entirely synthetic JWT: the payload decodes to {"sub":"test"}
    /// and the signature is the literal text "signature-placeholder". It is not a credential
    /// and grants nothing.
    /// </summary>
    private const string FakeJwt =
        "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJzdWIiOiJ0ZXN0In0.c2lnbmF0dXJlLXBsYWNlaG9sZGVy"; // SCAN-ALLOW: synthetic

    [Fact]
    public void Redact_BearerToken_IsRemoved()
    {
        var result = SensitiveDataRedactor.Redact($"Authorization: Bearer {FakeJwt}");

        Assert.DoesNotContain(FakeJwt, result, StringComparison.Ordinal);
        Assert.Contains(SensitiveDataRedactor.Placeholder, result, StringComparison.Ordinal);
    }

    [Fact]
    public void Redact_BareJwt_IsRemoved()
    {
        var result = SensitiveDataRedactor.Redact($"the token was {FakeJwt} apparently");

        Assert.DoesNotContain(FakeJwt, result, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("code")]
    [InlineData("access_token")]
    [InlineData("refresh_token")]
    [InlineData("id_token")]
    [InlineData("client_secret")]
    [InlineData("code_verifier")]
    [InlineData("password")]
    public void Redact_SensitiveQueryParameters_AreRemoved(string parameter)
    {
        var result = SensitiveDataRedactor.Redact(
            $"https://login.example.test/callback?{parameter}=SUPERSECRETVALUE&state=abc");

        Assert.DoesNotContain("SUPERSECRETVALUE", result, StringComparison.Ordinal);
        Assert.Contains(SensitiveDataRedactor.Placeholder, result, StringComparison.Ordinal);
    }

    /// <summary>Non-sensitive parameters stay readable, or logs become useless for diagnosis.</summary>
    [Fact]
    public void Redact_NonSensitiveParameters_ArePreserved()
    {
        var result = SensitiveDataRedactor.Redact("https://graph.example.test/v1.0/sites?search=marketing");

        Assert.Contains("search=marketing", result, StringComparison.Ordinal);
    }

    [Fact]
    public void Redact_AuthorizationHeaderLine_IsRemoved()
    {
        var result = SensitiveDataRedactor.Redact("authorization=abc123def456");

        Assert.DoesNotContain("abc123def456", result, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    public void Redact_EmptyInput_DoesNotThrow(string? input) =>
        Assert.Equal(string.Empty, SensitiveDataRedactor.Redact(input));

    [Fact]
    public void RedactUrl_DropsTheEntireQueryString()
    {
        var result = SensitiveDataRedactor.RedactUrl(
            "https://graph.example.test/v1.0/me?$select=id&code=secret");

        Assert.Equal("https://graph.example.test/v1.0/me?[REDACTED]", result);
    }

    [Fact]
    public void RedactUrl_WithoutQuery_IsUnchanged() =>
        Assert.Equal(
            "https://graph.example.test/v1.0/me",
            SensitiveDataRedactor.RedactUrl("https://graph.example.test/v1.0/me"));

    [Fact]
    public void MaskIdentifier_KeepsFirstAndLastFourCharacters()
    {
        var masked = SensitiveDataRedactor.MaskIdentifier("12345678-1234-1234-1234-123456789abc");

        Assert.StartsWith("1234", masked, StringComparison.Ordinal);
        Assert.EndsWith("9abc", masked, StringComparison.Ordinal);
        Assert.DoesNotContain("5678-1234", masked, StringComparison.Ordinal);
    }

    [Fact]
    public void MaskIdentifier_ShortValue_IsFullyMasked() =>
        Assert.Equal("****", SensitiveDataRedactor.MaskIdentifier("abcd"));

    [Fact]
    public void MaskEmail_KeepsTheDomainOnly() =>
        Assert.Equal("j***@example.test", SensitiveDataRedactor.MaskEmail("jane.doe@example.test"));
}
