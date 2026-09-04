using SharePointLinkManifestBuilder.Core.Models;

namespace SharePointLinkManifestBuilder.Core.Tests.Models;

/// <summary>
/// Covers the audience rules. These decide the authority URL, which is the single most
/// security-relevant string in the application.
/// </summary>
public sealed class TenantAudienceTests
{
    private const string Tenant = "11111111-1111-1111-1111-111111111111";
    private const string Client = "22222222-2222-2222-2222-222222222222";

    /// <summary>A single-tenant configuration must keep its tenant-specific authority.</summary>
    [Fact]
    public void Authority_SingleTenant_IsTenantSpecific()
    {
        var configuration = new TenantConfiguration { TenantId = Tenant, ClientId = Client };

        Assert.Equal($"https://login.microsoftonline.com/{Tenant}", configuration.Authority);
        Assert.False(configuration.IsMultiTenant);
    }

    /// <summary>
    /// A multi-tenant configuration must use /organizations. It must never use /common, which
    /// would additionally admit personal Microsoft accounts that have no SharePoint at all.
    /// </summary>
    [Fact]
    public void Authority_MultiTenant_UsesOrganizationsAndNeverCommon()
    {
        var configuration = new TenantConfiguration
        {
            TenantId = Tenant,
            ClientId = Client,
            Audience = TenantAudience.AnyOrganization,
        };

        Assert.Equal("https://login.microsoftonline.com/organizations", configuration.Authority);
        Assert.DoesNotContain("/common", configuration.Authority, StringComparison.OrdinalIgnoreCase);
        Assert.True(configuration.IsMultiTenant);
    }

    /// <summary>The tenant ID is irrelevant to a multi-tenant authority, even when supplied.</summary>
    [Fact]
    public void Authority_MultiTenant_IgnoresConfiguredTenantId()
    {
        var configuration = new TenantConfiguration
        {
            TenantId = Tenant,
            ClientId = Client,
            Audience = TenantAudience.AnyOrganization,
        };

        Assert.DoesNotContain(Tenant, configuration.Authority, StringComparison.OrdinalIgnoreCase);
    }

    /// <summary>A sovereign-cloud instance must still be honoured in multi-tenant mode.</summary>
    [Fact]
    public void Authority_MultiTenant_RespectsSovereignInstance()
    {
        var configuration = new TenantConfiguration
        {
            ClientId = Client,
            Instance = "https://login.microsoftonline.us/",
            Audience = TenantAudience.AnyOrganization,
        };

        Assert.Equal("https://login.microsoftonline.us/organizations", configuration.Authority);
    }

    /// <summary>A multi-tenant configuration is usable with no tenant ID at all.</summary>
    [Fact]
    public void IsUsable_MultiTenant_DoesNotRequireTenantId()
    {
        var configuration = new TenantConfiguration
        {
            ClientId = Client,
            Audience = TenantAudience.AnyOrganization,
        };

        Assert.True(configuration.IsUsable);
    }

    /// <summary>A single-tenant configuration is not usable without a tenant ID.</summary>
    [Fact]
    public void IsUsable_SingleTenant_RequiresTenantId()
    {
        var configuration = new TenantConfiguration { ClientId = Client };

        Assert.False(configuration.IsUsable);
    }

    /// <summary>A client ID is required whatever the audience.</summary>
    [Theory]
    [InlineData(TenantAudience.SingleTenant)]
    [InlineData(TenantAudience.AnyOrganization)]
    public void IsUsable_AlwaysRequiresAValidClientId(TenantAudience audience)
    {
        var configuration = new TenantConfiguration
        {
            TenantId = Tenant,
            ClientId = "not-a-guid",
            Audience = audience,
        };

        Assert.False(configuration.IsUsable);
    }

    /// <summary>The Entra sign-in audience value must match the chosen audience.</summary>
    [Theory]
    [InlineData(TenantAudience.SingleTenant, "AzureADMyOrg")]
    [InlineData(TenantAudience.AnyOrganization, "AzureADMultipleOrgs")]
    public void SignInAudience_MatchesTheChosenAudience(TenantAudience audience, string expected)
    {
        var registration = new AppRegistrationConfiguration
        {
            DisplayName = "Test",
            Audience = audience,
        };

        Assert.Equal(expected, registration.SignInAudience);
    }

    /// <summary>
    /// The planned-changes text is the user's only preview before a tenant is modified, so a
    /// multi-tenant registration must not describe itself as single-organization.
    /// </summary>
    [Fact]
    public void DescribePlannedChanges_MultiTenant_DoesNotClaimThisOrganizationOnly()
    {
        var registration = new AppRegistrationConfiguration
        {
            DisplayName = "Test",
            Audience = TenantAudience.AnyOrganization,
        };

        var changes = string.Join('\n', registration.DescribePlannedChanges());

        Assert.Contains("any work or school organization", changes, StringComparison.Ordinal);
        Assert.Contains("consent separately", changes, StringComparison.Ordinal);
        Assert.DoesNotContain("this organization only", changes, StringComparison.Ordinal);
    }

    /// <summary>A single-tenant registration must still say so plainly.</summary>
    [Fact]
    public void DescribePlannedChanges_SingleTenant_SaysThisOrganizationOnly()
    {
        var registration = new AppRegistrationConfiguration { DisplayName = "Test" };
        var changes = string.Join('\n', registration.DescribePlannedChanges());

        Assert.Contains("this organization only", changes, StringComparison.Ordinal);
    }

    /// <summary>No audience may ever cause a secret to be planned.</summary>
    [Theory]
    [InlineData(TenantAudience.SingleTenant)]
    [InlineData(TenantAudience.AnyOrganization)]
    public void DescribePlannedChanges_NeverPlansASecret(TenantAudience audience)
    {
        var registration = new AppRegistrationConfiguration
        {
            DisplayName = "Test",
            Audience = audience,
        };

        var changes = string.Join('\n', registration.DescribePlannedChanges());

        Assert.Contains("No client secret", changes, StringComparison.Ordinal);
        Assert.True(registration.IsFallbackPublicClient);
    }
}
