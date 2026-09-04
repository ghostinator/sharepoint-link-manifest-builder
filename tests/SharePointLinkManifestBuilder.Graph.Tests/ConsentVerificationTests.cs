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

/// <summary>
/// The consent redirect can fail for reasons unrelated to whether the permissions are granted.
/// These pin the classification, because the wrong one sends the user to fix something that is
/// not broken -- or worse, that cannot be fixed from the device they are on.
/// </summary>
public sealed class ConsentErrorMappingTests
{
    /// <summary>
    /// Conditional Access demanding a managed device arrives as interaction_required, which
    /// otherwise reads as "an administrator still needs to approve this". It is not that, and the
    /// remedy is completely different.
    /// </summary>
    [Fact]
    public void DeviceAuthenticationRequired_IsNotReportedAsMissingAdministratorApproval()
    {
        var error = ConsentService.MapConsentError(
            "interaction_required",
            "AADSTS50097: Device authentication is required. Trace ID: 00000000-0000-0000-0000-000000000000");

        Assert.Equal(GraphErrorKind.ConditionalAccessInterrupted, error.Kind);
        Assert.Contains("managed", error.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("Entra admin center", error.SuggestedAction!, StringComparison.OrdinalIgnoreCase);
    }

    /// <summary>An ordinary interaction_required still means approval is outstanding.</summary>
    [Fact]
    public void PlainInteractionRequired_StillMeansApprovalIsOutstanding()
    {
        var error = ConsentService.MapConsentError("interaction_required", "Something else entirely.");

        Assert.Equal(GraphErrorKind.AdminConsentRequired, error.Kind);
    }

    /// <summary>A declined consent is a decision and must stay distinguishable from a failure.</summary>
    [Fact]
    public void AccessDenied_IsReportedAsDeclined()
    {
        var error = ConsentService.MapConsentError("access_denied", null);

        Assert.Equal(GraphErrorKind.ConsentDenied, error.Kind);
    }

    /// <summary>A null description must not throw the AADSTS check.</summary>
    [Fact]
    public void MissingDescription_IsHandled()
    {
        var error = ConsentService.MapConsentError("interaction_required", null);

        Assert.Equal(GraphErrorKind.AdminConsentRequired, error.Kind);
    }
}
