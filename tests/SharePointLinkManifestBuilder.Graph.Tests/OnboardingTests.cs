using System.Net;
using Microsoft.Extensions.Logging.Abstractions;
using SharePointLinkManifestBuilder.Core.Abstractions;
using SharePointLinkManifestBuilder.Core.Models;
using SharePointLinkManifestBuilder.Graph.Onboarding;

namespace SharePointLinkManifestBuilder.Graph.Tests;

/// <summary>Creating and inspecting the tenant-specific application registration.</summary>
public sealed class AppRegistrationServiceTests
{
    private static AppRegistrationService Create(FakeGraphHandler handler) =>
        new(GraphTestHarness.CreateClient(handler), NullLogger<AppRegistrationService>.Instance);

    private static AppRegistrationConfiguration Configuration => new()
    {
        DisplayName = "SharePoint Link Manifest Builder",
        RequestedPermissions = GraphScopes.StandardTier,
    };

    private const string CreatedApplication = """
        {"id":"app-object-1","appId":"22222222-2222-2222-2222-222222222222",
         "displayName":"SharePoint Link Manifest Builder","signInAudience":"AzureADMyOrg",
         "isFallbackPublicClient":true,"publicClient":{"redirectUris":["http://localhost"]}}
        """;

    [Fact]
    public async Task CreateRegistration_ReturnsTheNewClientId()
    {
        var handler = new FakeGraphHandler().Enqueue(FakeResponse.Created(CreatedApplication));

        var result = await Create(handler).CreateRegistrationAsync(
            Configuration, "11111111-1111-1111-1111-111111111111");

        Assert.True(result.Succeeded);
        Assert.Equal("22222222-2222-2222-2222-222222222222", result.ClientId);
        Assert.Equal("app-object-1", result.ApplicationObjectId);
        Assert.Equal(RegistrationSource.AutomaticSetup, result.Configuration!.Source);
    }

    /// <summary>
    /// The whole least-privilege story depends on this: everything goes in the initial POST, so
    /// no PATCH is needed and the bootstrap identity can stay create-only.
    /// </summary>
    [Fact]
    public async Task CreateRegistration_SendsACompleteObjectInASingleRequest()
    {
        var handler = new FakeGraphHandler().Enqueue(FakeResponse.Created(CreatedApplication));

        await Create(handler).CreateRegistrationAsync(
            Configuration, "11111111-1111-1111-1111-111111111111");

        var request = Assert.Single(handler.Requests);

        Assert.Equal("POST", request.Method);
        Assert.EndsWith("/applications", request.Uri.AbsolutePath, StringComparison.Ordinal);

        var body = request.Body!;
        Assert.Contains("AzureADMyOrg", body, StringComparison.Ordinal);
        Assert.Contains("IsFallbackPublicClient", body, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("localhost", body, StringComparison.Ordinal);
        Assert.Contains("RequiredResourceAccess", body, StringComparison.OrdinalIgnoreCase);
        Assert.Contains(AuthorityDefaults.MicrosoftGraphResourceAppId, body, StringComparison.Ordinal);
    }

    /// <summary>A desktop public client must never be given a secret.</summary>
    [Fact]
    public async Task CreateRegistration_NeverRequestsAPasswordCredential()
    {
        var handler = new FakeGraphHandler().Enqueue(FakeResponse.Created(CreatedApplication));

        await Create(handler).CreateRegistrationAsync(
            Configuration, "11111111-1111-1111-1111-111111111111");

        var body = Assert.Single(handler.Requests).Body!;

        Assert.DoesNotContain("passwordCredential", body, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("keyCredential", body, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("secret", body, StringComparison.OrdinalIgnoreCase);
    }

    /// <summary>The planned changes must state that no secret will be created.</summary>
    [Fact]
    public async Task CreateRegistration_ReportsThatNoSecretWillBeCreated()
    {
        var handler = new FakeGraphHandler().Enqueue(FakeResponse.Created(CreatedApplication));

        var result = await Create(handler).CreateRegistrationAsync(
            Configuration, "11111111-1111-1111-1111-111111111111");

        Assert.Contains(
            result.ChangesApplied,
            c => c.Contains("No client secret", StringComparison.OrdinalIgnoreCase));
    }

    /// <summary>A tenant that forbids self-service registration must get a usable explanation.</summary>
    [Fact]
    public async Task CreateRegistration_Forbidden_ExplainsAndPointsAtTheAlternative()
    {
        var handler = new FakeGraphHandler()
            .Enqueue(FakeResponse.Error(HttpStatusCode.Forbidden, "Authorization_RequestDenied"));

        var result = await Create(handler).CreateRegistrationAsync(
            Configuration, "11111111-1111-1111-1111-111111111111");

        Assert.False(result.Succeeded);
        Assert.Equal(GraphErrorKind.InsufficientPrivilegesToCreateApplication, result.Error!.Kind);
        Assert.Contains("Existing app registration", result.Error.SuggestedAction, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task CreateRegistration_BadRequest_MapsToRegistrationBlocked()
    {
        var handler = new FakeGraphHandler()
            .Enqueue(FakeResponse.Error(HttpStatusCode.BadRequest, "Request_BadRequest"));

        var result = await Create(handler).CreateRegistrationAsync(
            Configuration, "11111111-1111-1111-1111-111111111111");

        Assert.Equal(GraphErrorKind.RegistrationCreationBlocked, result.Error!.Kind);
    }

    /// <summary>
    /// Being unable to read the registration is the normal state, because the product does not
    /// request a directory read permission. It must not be presented as a missing registration.
    /// </summary>
    [Fact]
    public async Task InspectRegistration_WithoutReadPermission_ReportsNotVerifiedNotMissing()
    {
        var handler = new FakeGraphHandler()
            .Enqueue(FakeResponse.Error(HttpStatusCode.Forbidden, "Authorization_RequestDenied"));

        var result = await Create(handler).InspectRegistrationAsync("22222222-2222-2222-2222-222222222222");

        Assert.True(result.Succeeded);
        Assert.False(result.Value!.ApplicationFound);
        Assert.Contains(
            result.Value.NotVerified,
            n => n.Contains("does not mean the registration is missing", StringComparison.OrdinalIgnoreCase));
    }

    /// <summary>A secret on a public-client registration is worth flagging loudly.</summary>
    [Fact]
    public async Task InspectRegistration_WithAClientSecret_WarnsAboutIt()
    {
        var handler = new FakeGraphHandler()
            .Map("/applications", FakeResponse.Ok("""
                {"id":"app-1","appId":"22222222-2222-2222-2222-222222222222",
                 "isFallbackPublicClient":true,
                 "publicClient":{"redirectUris":["http://localhost"]},
                 "passwordCredentials":[{"keyId":"k1","displayName":"secret"}]}
                """))
            .Map("/servicePrincipals", FakeResponse.Ok("""{"id":"sp-1","appId":"22222222-2222-2222-2222-222222222222"}"""));

        var result = await Create(handler).InspectRegistrationAsync("22222222-2222-2222-2222-222222222222");

        Assert.True(result.Value!.ApplicationFound);
        Assert.True(result.Value.IsPublicClient);
        Assert.True(result.Value.RedirectUriConfigured);
        Assert.Contains(result.Value.Warnings, w => w.Contains("client secret", StringComparison.OrdinalIgnoreCase));
    }

    /// <summary>Deletion must refuse when the typed name does not match, even if the caller asked.</summary>
    [Fact]
    public async Task DeleteRegistration_NameMismatch_RefusesWithoutDeleting()
    {
        var handler = new FakeGraphHandler()
            .Enqueue(FakeResponse.Ok("""{"id":"app-1","displayName":"The Real Name"}"""));

        var result = await Create(handler).DeleteRegistrationAsync("app-1", "Something Else");

        Assert.False(result.Succeeded);
        Assert.Contains("does not match", result.Error!.Message, StringComparison.OrdinalIgnoreCase);

        // Only the read happened; no DELETE was issued.
        Assert.Equal(1, handler.RequestCount);
        Assert.Equal("GET", handler.Requests[0].Method);
    }

    [Fact]
    public async Task DeleteRegistration_NameMatches_IssuesTheDelete()
    {
        var handler = new FakeGraphHandler().Enqueue(
            FakeResponse.Ok("""{"id":"app-1","displayName":"The Real Name"}"""),
            new FakeResponse { Status = HttpStatusCode.NoContent });

        var result = await Create(handler).DeleteRegistrationAsync("app-1", "The Real Name");

        Assert.True(result.Succeeded);
        Assert.Equal("DELETE", handler.Requests[1].Method);
    }

    /// <summary>Consent normally creates the service principal, so this path is repair-only.</summary>
    [Fact]
    public async Task EnsureServicePrincipal_WhenItAlreadyExists_DoesNotCreateAnother()
    {
        var handler = new FakeGraphHandler().Enqueue(FakeResponse.Ok("""{"id":"sp-1","appId":"a"}"""));

        var result = await Create(handler).EnsureServicePrincipalAsync("22222222-2222-2222-2222-222222222222");

        Assert.True(result.Succeeded);
        Assert.Equal("sp-1", result.Value);
        Assert.Equal(1, handler.RequestCount);
    }

    [Fact]
    public async Task EnsureServicePrincipal_WhenAbsent_CreatesIt()
    {
        var handler = new FakeGraphHandler().Enqueue(
            FakeResponse.Error(HttpStatusCode.NotFound, "Request_ResourceNotFound"),
            FakeResponse.Created("""{"id":"sp-new","appId":"a"}"""));

        var result = await Create(handler).EnsureServicePrincipalAsync("22222222-2222-2222-2222-222222222222");

        Assert.True(result.Succeeded);
        Assert.Equal("sp-new", result.Value);
        Assert.Equal("POST", handler.Requests[1].Method);
    }

    /// <summary>
    /// Capability is reported as Unknown rather than probed. Speculatively creating a
    /// registration to find out would be exactly the silent tenant change this product forbids.
    /// </summary>
    [Fact]
    public async Task EstimateCapability_MakesNoRequestAndClaimsNothing()
    {
        var handler = new FakeGraphHandler();

        var capability = await Create(handler).EstimateCapabilityAsync();

        Assert.Equal(RegistrationCapability.Unknown, capability);
        Assert.Equal(0, handler.RequestCount);
    }
}

/// <summary>The Microsoft Graph permission identifier table used to configure a registration.</summary>
public sealed class GraphPermissionIdsTests
{
    [Theory]
    [InlineData("User.Read")]
    [InlineData("User.ReadBasic.All")]
    [InlineData("Sites.Read.All")]
    [InlineData("Sites.ReadWrite.All")]
    [InlineData("Files.Read.All")]
    [InlineData("Files.ReadWrite.All")]
    public void EveryOperatingScope_HasAKnownIdentifier(string scope)
    {
        var id = GraphPermissionIds.TryGetScopeId(scope);

        Assert.NotNull(id);
        Assert.True(Guid.TryParse(id, out _), $"'{scope}' maps to '{id}', which is not a GUID.");
    }

    [Fact]
    public void BuildRequiredResourceAccess_TargetsMicrosoftGraphWithDelegatedScopes()
    {
        var access = GraphPermissionIds.BuildRequiredResourceAccess(
            GraphScopes.StandardTier, out var unmapped);

        var resource = Assert.Single(access);

        Assert.Equal(AuthorityDefaults.MicrosoftGraphResourceAppId, resource.ResourceAppId);
        Assert.Equal(3, resource.ResourceAccess!.Count);
        Assert.All(resource.ResourceAccess, a => Assert.Equal("Scope", a.Type));
        Assert.Empty(unmapped);
    }

    /// <summary>MSAL supplies the OIDC scopes; they are not listed in requiredResourceAccess.</summary>
    [Fact]
    public void BuildRequiredResourceAccess_ExcludesReservedOidcScopes()
    {
        var permissions = new List<PermissionRequirement>(GraphScopes.StandardTier)
        {
            new() { Scope = "openid", Purpose = "x", DataAccessImpact = "x" },
            new() { Scope = "offline_access", Purpose = "x", DataAccessImpact = "x" },
        };

        var access = GraphPermissionIds.BuildRequiredResourceAccess(permissions, out var unmapped);

        Assert.Equal(3, Assert.Single(access).ResourceAccess!.Count);
        Assert.Empty(unmapped);
    }

    /// <summary>
    /// An unknown scope is reported rather than guessed at. A wrong GUID would silently
    /// configure the wrong permission, which is worse than configuring none.
    /// </summary>
    [Fact]
    public void BuildRequiredResourceAccess_UnknownScope_IsReportedNotGuessed()
    {
        var permissions = new[]
        {
            GraphScopes.UserRead,
            new PermissionRequirement
            {
                Scope = "Some.Future.Scope",
                Purpose = "x",
                DataAccessImpact = "x",
            },
        };

        var access = GraphPermissionIds.BuildRequiredResourceAccess(permissions, out var unmapped);

        Assert.Single(Assert.Single(access).ResourceAccess!);
        Assert.Equal("Some.Future.Scope", Assert.Single(unmapped));
    }
}

/// <summary>The official administrator-consent URL and the errors a redirect can carry.</summary>
public sealed class ConsentUrlTests
{
    private static TenantConfiguration Tenant => new()
    {
        TenantId = "11111111-1111-1111-1111-111111111111",
        ClientId = "22222222-2222-2222-2222-222222222222",
        GraphEndpoint = "https://graph.example.test/v1.0",
        Instance = "https://login.example.test",
    };

    private static ConsentService Create() =>
        new(new FakeAuthenticationService(),
            new NoOpBrowser(),
            GraphTestHarness.CreateClient(new FakeGraphHandler()),
            NullLogger<ConsentService>.Instance);

    /// <summary>
    /// Tenant-specific, never /common. A /common consent URL lets an administrator consent in
    /// whichever directory they happen to be signed into, which is not necessarily this one.
    /// </summary>
    [Fact]
    public void AdminConsentUrl_IsTenantSpecificAndWellFormed()
    {
        var url = Create().BuildAdminConsentUrl(
            Tenant, GraphScopes.StandardTier, "http://localhost:5000/", "state-123");

        Assert.StartsWith(
            "https://login.example.test/11111111-1111-1111-1111-111111111111/v2.0/adminconsent",
            url.ToString(),
            StringComparison.Ordinal);

        Assert.DoesNotContain("/common/", url.ToString(), StringComparison.Ordinal);

        var query = Uri.UnescapeDataString(url.Query);
        Assert.Contains("client_id=22222222-2222-2222-2222-222222222222", query, StringComparison.Ordinal);
        Assert.Contains("state=state-123", query, StringComparison.Ordinal);
        Assert.Contains("redirect_uri=http://localhost:5000/", query, StringComparison.Ordinal);
    }

    /// <summary>Admin consent needs fully-qualified scope URIs, not bare names.</summary>
    [Fact]
    public void AdminConsentUrl_UsesFullyQualifiedScopes()
    {
        var url = Create().BuildAdminConsentUrl(
            Tenant, GraphScopes.StandardTier, "http://localhost:5000/", "s");

        var query = Uri.UnescapeDataString(url.Query);

        Assert.Contains("https://graph.example.test/Sites.Read.All", query, StringComparison.Ordinal);
        Assert.Contains("https://graph.example.test/Files.ReadWrite.All", query, StringComparison.Ordinal);
        Assert.DoesNotContain("openid", query, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("access_denied", GraphErrorKind.ConsentDenied)]
    [InlineData("invalid_client", GraphErrorKind.AppRegistrationNotFound)]
    [InlineData("unauthorized_client", GraphErrorKind.UnauthorizedAdministratorRole)]
    [InlineData("insufficient_privileges", GraphErrorKind.UnauthorizedAdministratorRole)]
    [InlineData("consent_required", GraphErrorKind.AdminConsentRequired)]
    public void ConsentErrors_MapToTheRightKind(string error, GraphErrorKind expected) =>
        Assert.Equal(expected, ConsentService.MapConsentError(error, null).Kind);

    /// <summary>An unfamiliar error still produces something a user can act on.</summary>
    [Fact]
    public void UnknownConsentError_StillProducesAMessage()
    {
        var mapped = ConsentService.MapConsentError("something_new", "A long service description");

        Assert.Equal(GraphErrorKind.ConsentRequired, mapped.Kind);
        Assert.Equal("A long service description", mapped.Message);
    }

    /// <summary>State must be unguessable, or the forgery protection is decorative.</summary>
    [Fact]
    public void GeneratedState_IsLongAndUnique()
    {
        var values = Enumerable.Range(0, 50)
            .Select(_ => LoopbackRedirectListener.GenerateState())
            .ToArray();

        Assert.Equal(50, values.Distinct(StringComparer.Ordinal).Count());
        Assert.All(values, v => Assert.Equal(64, v.Length));
    }

    private sealed class NoOpBrowser : ISystemBrowser
    {
        public Task OpenAsync(Uri url, CancellationToken cancellationToken = default) => Task.CompletedTask;
    }
}

/// <summary>The bootstrap identity, which this repository deliberately does not supply.</summary>
public sealed class BootstrapConfigurationProviderTests
{
    /// <summary>
    /// Shipping a client ID in source control would be a supply-chain problem, so automatic
    /// setup must report itself unavailable rather than fabricate one.
    /// </summary>
    [Fact]
    public void WithNoClientId_AutomaticSetupIsUnavailableWithAnExplanation()
    {
        var provider = new BootstrapConfigurationProvider();

        Assert.False(provider.Current.IsConfigured);
        Assert.Contains("bootstrap client ID", provider.Current.UnavailableReason, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("Existing app registration", provider.Current.UnavailableReason, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void WithAValidClientId_AutomaticSetupBecomesAvailable()
    {
        var provider = new BootstrapConfigurationProvider("33333333-3333-3333-3333-333333333333");

        Assert.True(provider.Current.IsConfigured);
    }

    [Fact]
    public void WithANonGuidClientId_ReportsThatItIsInvalid()
    {
        var provider = new BootstrapConfigurationProvider("not-a-guid");

        Assert.False(provider.Current.IsConfigured);
        Assert.Contains("not a valid GUID", provider.Current.UnavailableReason, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void SetClientId_UpdatesTheConfiguration()
    {
        var provider = new BootstrapConfigurationProvider();

        provider.SetClientId("  44444444-4444-4444-4444-444444444444  ");

        Assert.True(provider.Current.IsConfigured);
        Assert.Equal("44444444-4444-4444-4444-444444444444", provider.Current.ClientId);
    }
}
