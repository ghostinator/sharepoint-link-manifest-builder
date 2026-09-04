using Microsoft.Extensions.Logging.Abstractions;
using SharePointLinkManifestBuilder.Core.Abstractions;
using SharePointLinkManifestBuilder.Core.Models;
using SharePointLinkManifestBuilder.Graph.Onboarding;

namespace SharePointLinkManifestBuilder.Graph.Tests;

/// <summary>
/// Consent verification acquires a token and reads back the scopes Entra issued (ADR-0006).
/// These pin the distinction that made a correctly consented tenant report itself unconsented:
/// a silent acquisition failing with "interaction required" says nothing about whether an
/// administrator has consented, only that this user has no cached grant yet.
/// </summary>
public sealed class ConsentVerificationTests
{
    private static readonly TenantConfiguration Tenant = new()
    {
        TenantId = "11111111-1111-1111-1111-111111111111",
        ClientId = "22222222-2222-2222-2222-222222222222",
        RequiredScopes = ["User.Read", "Sites.Read.All", "Files.ReadWrite.All"],
    };

    private static readonly PermissionRequirement[] Required =
    [
        GraphScopes.UserRead, GraphScopes.SitesReadAll, GraphScopes.FilesReadWriteAll,
    ];

    private static (ConsentService Service, FakeAuthenticationService Auth) Create(
        FakeAuthenticationService auth)
    {
        var handler = new FakeGraphHandler();
        var client = GraphTestHarness.CreateClient(handler, auth);

        return (new ConsentService(auth, new NoOpBrowser(), client,
            NullLogger<ConsentService>.Instance), auth);
    }

    /// <summary>A cached grant needs no prompt, so silent alone must be enough.</summary>
    [Fact]
    public async Task WhenSilentSucceeds_NoInteractivePromptIsRaised()
    {
        var (service, auth) = Create(new FakeAuthenticationService());

        var verification = await service.VerifyConsentAsync(Tenant, Required, allowInteractive: true);

        Assert.True(verification.CanAcquireToken);
        Assert.Equal(ConsentState.Granted, verification.ConsentState);
        Assert.Equal(1, auth.SilentAttempts);
        Assert.Equal(0, auth.InteractiveAttempts);
    }

    /// <summary>
    /// The reported bug. Consent is granted in the directory, but this user has no cached grant,
    /// so the silent request fails. Verification must escalate rather than conclude from silence.
    /// </summary>
    [Fact]
    public async Task WhenSilentNeedsInteraction_ItRetriesInteractivelyAndSucceeds()
    {
        var (service, auth) = Create(new FakeAuthenticationService
        {
            SilentOnlyFailure = GraphErrorKind.AdminConsentRequired,
        });

        var verification = await service.VerifyConsentAsync(Tenant, Required, allowInteractive: true);

        Assert.Equal(1, auth.SilentAttempts);
        Assert.Equal(1, auth.InteractiveAttempts);
        Assert.True(verification.CanAcquireToken);
        Assert.Equal(ConsentState.Granted, verification.ConsentState);
    }

    /// <summary>
    /// Without opting in, behaviour is unchanged: no prompt appears at a moment the user did not
    /// ask for one, such as a background refresh.
    /// </summary>
    [Fact]
    public async Task WhenInteractiveIsNotAllowed_ItDoesNotPrompt()
    {
        var (service, auth) = Create(new FakeAuthenticationService
        {
            SilentOnlyFailure = GraphErrorKind.AdminConsentRequired,
        });

        var verification = await service.VerifyConsentAsync(Tenant, Required, allowInteractive: false);

        Assert.Equal(0, auth.InteractiveAttempts);
        Assert.False(verification.CanAcquireToken);
        Assert.Equal(ConsentState.PendingAdministratorApproval, verification.ConsentState);
    }

    /// <summary>
    /// A refusal is a decision, not a missing cache entry. Re-prompting after one would nag the
    /// user with the dialog they just dismissed.
    /// </summary>
    [Fact]
    public async Task WhenConsentWasDenied_ItDoesNotRetryInteractively()
    {
        var (service, auth) = Create(new FakeAuthenticationService
        {
            SilentOnlyFailure = GraphErrorKind.ConsentDenied,
        });

        var verification = await service.VerifyConsentAsync(Tenant, Required, allowInteractive: true);

        Assert.Equal(0, auth.InteractiveAttempts);
        Assert.Equal(ConsentState.Denied, verification.ConsentState);
    }

    private sealed class NoOpBrowser : ISystemBrowser
    {
        public Task OpenAsync(Uri url, CancellationToken cancellationToken = default) =>
            Task.CompletedTask;
    }
}
