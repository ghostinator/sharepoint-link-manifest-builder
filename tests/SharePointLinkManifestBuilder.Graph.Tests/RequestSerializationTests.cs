using System.Net;
using System.Text.Json;
using SharePointLinkManifestBuilder.Graph.Dto;
using SharePointLinkManifestBuilder.Graph.Http;

namespace SharePointLinkManifestBuilder.Graph.Tests;

/// <summary>
/// Microsoft Graph names every property in camelCase. These pin the wire format, because the
/// mismatch that made automatic setup fail was invisible from inside the application: reads were
/// unaffected, and Graph is lenient enough about scalar properties that most writes were
/// accepted in PascalCase anyway. It is not lenient about complex ones, and a registration POST
/// carrying "PublicClient" was rejected with "Invalid property 'PublicClient'".
/// </summary>
public sealed class RequestSerializationTests
{
    private static JsonElement Serialize<T>(T value) =>
        JsonDocument.Parse(JsonSerializer.Serialize(value, GraphApiClient.SerializerOptions))
            .RootElement;

    /// <summary>The exact failure that was reported: a complex property Graph rejected.</summary>
    [Fact]
    public void CreateApplicationRequest_NamesEveryPropertyInCamelCase()
    {
        var json = Serialize(new CreateApplicationRequest
        {
            DisplayName = "Test",
            SignInAudience = "AzureADMyOrg",
            IsFallbackPublicClient = true,
            PublicClient = new GraphPublicClientDto { RedirectUris = ["http://localhost"] },
            RequiredResourceAccess = [],
        });

        Assert.True(json.TryGetProperty("publicClient", out var publicClient));
        Assert.True(json.TryGetProperty("displayName", out _));
        Assert.True(json.TryGetProperty("signInAudience", out _));
        Assert.True(json.TryGetProperty("isFallbackPublicClient", out _));
        Assert.True(json.TryGetProperty("requiredResourceAccess", out _));

        // The rejected spelling must be gone, not merely accompanied by the right one.
        Assert.False(json.TryGetProperty("PublicClient", out _));

        // Nested objects follow the same policy; Graph rejects them just as readily.
        Assert.True(publicClient.TryGetProperty("redirectUris", out _));
    }

    /// <summary>
    /// This one already worked against a live tenant, because Graph tolerated the scalars. It is
    /// pinned anyway: it was working by leniency rather than by being correct.
    /// </summary>
    [Fact]
    public void CreateLinkRequest_NamesEveryPropertyInCamelCase()
    {
        var json = Serialize(new CreateLinkRequest
        {
            Type = "view",
            Scope = "organization",
            ExpirationDateTime = "2026-01-01T00:00:00Z",
            RetainInheritedPermissions = false,
        });

        Assert.True(json.TryGetProperty("type", out _));
        Assert.True(json.TryGetProperty("scope", out _));
        Assert.True(json.TryGetProperty("expirationDateTime", out _));
        Assert.True(json.TryGetProperty("retainInheritedPermissions", out _));
        Assert.False(json.TryGetProperty("Type", out _));
    }

    /// <summary>
    /// An explicit [JsonPropertyName] must win over the policy. "@microsoft.graph.conflictBehavior"
    /// is not camelCase of anything, and camelCasing it would silently disable conflict handling.
    /// </summary>
    [Fact]
    public void ExplicitPropertyNames_SurviveTheNamingPolicy()
    {
        var json = Serialize(new DriveRecipientDto { Email = "someone@example.test" });

        Assert.True(json.TryGetProperty("email", out _));

        var item = Serialize(new UploadSessionItemDto
        {
            ConflictBehavior = "replace",
            Name = "manifest.txt",
        });

        Assert.True(item.TryGetProperty("@microsoft.graph.conflictBehavior", out _));
        Assert.True(item.TryGetProperty("name", out _));
    }

    /// <summary>
    /// The end-to-end proof: what the transport actually puts on the wire, not what a
    /// standalone serializer call produces.
    /// </summary>
    [Fact]
    public async Task PostAsync_PutsCamelCaseOnTheWire()
    {
        var handler = new FakeGraphHandler();
        handler.Enqueue(new FakeResponse { Status = HttpStatusCode.Created, Json = "{}" });

        var client = GraphTestHarness.CreateClient(handler);

        await client.PostAsync<GraphApplicationDto>(
            "/applications",
            new CreateApplicationRequest
            {
                DisplayName = "Test",
                SignInAudience = "AzureADMultipleOrgs",
                IsFallbackPublicClient = true,
                PublicClient = new GraphPublicClientDto { RedirectUris = ["http://localhost"] },
                RequiredResourceAccess = [],
            },
            TestContext.Current.CancellationToken);

        var body = Assert.Single(handler.Requests).Body;

        Assert.NotNull(body);
        Assert.Contains("\"publicClient\"", body, StringComparison.Ordinal);
        Assert.DoesNotContain("\"PublicClient\"", body, StringComparison.Ordinal);
    }
}
