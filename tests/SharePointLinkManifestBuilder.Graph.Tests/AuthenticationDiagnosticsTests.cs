using Microsoft.Identity.Client;
using SharePointLinkManifestBuilder.Core.Models;
using SharePointLinkManifestBuilder.Graph.Identity;
using SharePointLinkManifestBuilder.Graph.Onboarding;

namespace SharePointLinkManifestBuilder.Graph.Tests;

/// <summary>
/// Covers classification of the failures that appear *after* the system browser has shown
/// "authentication complete". MSAL renders that page as soon as the redirect arrives, whether
/// or not the redirect carries an error, so this whole family looks like success to the user
/// and must be distinguished by its Entra error code.
/// </summary>
public sealed class AuthenticationDiagnosticsTests
{
    /// <summary>The Entra code must be recovered from a realistic MSAL message.</summary>
    [Theory]
    [InlineData("AADSTS7000218: The request body must contain the following parameter.", "AADSTS7000218")]
    [InlineData("AADSTS50011: The redirect URI specified does not match.", "AADSTS50011")]
    [InlineData("aadsts50020: User account does not exist in tenant.", "AADSTS50020")]
    [InlineData("Something failed. AADSTS900971 no reply address. Trace ID: x", "AADSTS900971")]
    public void ExtractEntraErrorCode_FindsTheCode(string message, string expected) =>
        Assert.Equal(expected, MsalAuthenticationService.ExtractEntraErrorCode(message));

    /// <summary>Absence of a code must be reported as absence, not an empty string.</summary>
    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("A network error occurred.")]
    [InlineData("AADSTS")]
    [InlineData("AADSTS12")]
    public void ExtractEntraErrorCode_ReturnsNullWhenAbsent(string? message) =>
        Assert.Null(MsalAuthenticationService.ExtractEntraErrorCode(message));

    /// <summary>
    /// A registration that is not a public client is the single most likely cause of a sign-in
    /// that fails only after the browser reports success, so it must not land in the generic
    /// fallback.
    /// </summary>
    [Fact]
    public void MapMsalException_NotAPublicClient_IsIdentifiedPrecisely()
    {
        var exception = new MsalServiceException(
            "invalid_client",
            "AADSTS7000218: The request body must contain the following parameter: "
            + "'client_assertion' or 'client_secret'.");

        var error = MsalAuthenticationService.MapMsalException(exception, "sign in");

        Assert.Equal(GraphErrorKind.PublicClientNotConfigured, error.Kind);
        Assert.Contains("public client", error.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("Allow public client flows", error.SuggestedAction!, StringComparison.Ordinal);

        // The remedy must never be "add a secret": this application has none by design.
        Assert.Contains("never be given one", error.SuggestedAction!, StringComparison.Ordinal);
    }

    /// <summary>A redirect problem must name the redirect URI to register.</summary>
    [Theory]
    [InlineData("AADSTS50011: The redirect URI does not match.")]
    [InlineData("AADSTS900971: No reply address provided.")]
    public void MapMsalException_RedirectProblems_AreIdentified(string message)
    {
        var error = MsalAuthenticationService.MapMsalException(
            new MsalServiceException("invalid_request", message), "sign in");

        Assert.Equal(GraphErrorKind.RedirectUriMismatch, error.Kind);
        Assert.Contains("http://localhost", error.SuggestedAction!, StringComparison.Ordinal);
    }

    /// <summary>
    /// An account from another organization must point at multi-organization mode, since that
    /// is the actual fix rather than a retry.
    /// </summary>
    [Fact]
    public void MapMsalException_AccountFromAnotherTenant_PointsAtMultiOrganizationMode()
    {
        var error = MsalAuthenticationService.MapMsalException(
            new MsalServiceException(
                "invalid_request",
                "AADSTS50020: User account from identity provider does not exist in tenant."),
            "sign in");

        Assert.Equal(GraphErrorKind.AccountFromUnsupportedTenant, error.Kind);
        Assert.Contains("any work or school organization", error.SuggestedAction!, StringComparison.Ordinal);
    }

    /// <summary>A registration missing from the signed-in directory must be called out.</summary>
    [Theory]
    [InlineData("AADSTS700016: Application not found in the directory.")]
    [InlineData("AADSTS90002: Tenant not found.")]
    public void MapMsalException_ApplicationNotInTenant_IsIdentified(string message)
    {
        var error = MsalAuthenticationService.MapMsalException(
            new MsalServiceException("unauthorized_client", message), "sign in");

        Assert.Equal(GraphErrorKind.ApplicationNotFoundInTenant, error.Kind);
    }

    /// <summary>
    /// Ordering guard. MsalUiRequiredException derives from MsalServiceException, so the new
    /// AADSTS arms could capture consent failures if they were placed above the UI-required
    /// arms. Admin consent must still classify as admin consent.
    /// </summary>
    [Fact]
    public void MapMsalException_AdminConsentStillClassifiesAsAdminConsent()
    {
        var error = MsalAuthenticationService.MapMsalException(
            new MsalUiRequiredException(
                "invalid_grant",
                "AADSTS65001: The user or administrator has not consented."),
            "sign in");

        Assert.Equal(GraphErrorKind.AdminConsentRequired, error.Kind);
    }

    /// <summary>An unclassified failure must still be reported, not swallowed.</summary>
    [Fact]
    public void MapMsalException_UnknownFailure_StillReportsAFailure()
    {
        var error = MsalAuthenticationService.MapMsalException(
            new MsalServiceException("something_new", "AADSTS999999: Undocumented."), "sign in");

        Assert.Equal(GraphErrorKind.AuthenticationFailed, error.Kind);
    }

    /// <summary>Cancellation must never be reported as an Entra refusal.</summary>
    [Fact]
    public void MapMsalException_UserCancelled_IsCancellation()
    {
        var error = MsalAuthenticationService.MapMsalException(
            new MsalClientException("authentication_canceled", "User closed the window."), "sign in");

        Assert.Equal(GraphErrorKind.Canceled, error.Kind);
    }

    /// <summary>A single-tenant configuration always consents in its own directory.</summary>
    [Fact]
    public void ResolveConsentTenantId_SingleTenant_UsesTheConfiguredTenant()
    {
        var configuration = new TenantConfiguration
        {
            TenantId = "11111111-1111-1111-1111-111111111111",
            ClientId = "22222222-2222-2222-2222-222222222222",
        };

        Assert.Equal(
            configuration.TenantId,
            ConsentService.ResolveConsentTenantId(configuration, null, signedInAccount: null));
    }

    /// <summary>A multi-tenant configuration consents in the signed-in directory.</summary>
    [Fact]
    public void ResolveConsentTenantId_MultiTenant_UsesTheSignedInTenant()
    {
        var configuration = new TenantConfiguration
        {
            TenantId = "11111111-1111-1111-1111-111111111111",
            ClientId = "22222222-2222-2222-2222-222222222222",
            Audience = TenantAudience.AnyOrganization,
        };

        var account = new UserAccount
        {
            UserId = "u",
            DisplayName = "d",
            UserPrincipalName = "u@example.test",
            TenantId = "33333333-3333-3333-3333-333333333333",
        };

        Assert.Equal(
            account.TenantId,
            ConsentService.ResolveConsentTenantId(configuration, null, account));
    }

    /// <summary>
    /// With no signed-in account a multi-tenant configuration must refuse to guess. Falling
    /// back to /organizations here would let an administrator who is signed into several
    /// directories consent in the wrong one.
    /// </summary>
    [Fact]
    public void ResolveConsentTenantId_MultiTenantWithNoAccount_IsUndetermined()
    {
        var configuration = new TenantConfiguration
        {
            ClientId = "22222222-2222-2222-2222-222222222222",
            Audience = TenantAudience.AnyOrganization,
        };

        Assert.Null(ConsentService.ResolveConsentTenantId(configuration, null, signedInAccount: null));
    }

    /// <summary>An explicit target always wins, for onboarding a directory before sign-in.</summary>
    [Fact]
    public void ResolveConsentTenantId_ExplicitTarget_TakesPrecedence()
    {
        var configuration = new TenantConfiguration
        {
            TenantId = "11111111-1111-1111-1111-111111111111",
            ClientId = "22222222-2222-2222-2222-222222222222",
            Audience = TenantAudience.AnyOrganization,
        };

        Assert.Equal(
            "44444444-4444-4444-4444-444444444444",
            ConsentService.ResolveConsentTenantId(
                configuration, " 44444444-4444-4444-4444-444444444444 ", signedInAccount: null));
    }
}
