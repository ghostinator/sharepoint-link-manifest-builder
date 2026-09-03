using Microsoft.Extensions.DependencyInjection;
using SharePointLinkManifestBuilder.App.Composition;
using SharePointLinkManifestBuilder.App.ViewModels;
using SharePointLinkManifestBuilder.Core.Settings;

namespace SharePointLinkManifestBuilder.App.Tests;

/// <summary>
/// Covers the wizard's escape hatches. Both failures these pin were reachable in normal use and
/// left the user with no way forward: a browser round-trip that never returns disabled its own
/// button, and the automatic-setup method could not be selected in order to supply the very
/// value that would enable it.
/// </summary>
public sealed class SetupWizardRecoveryTests : IDisposable
{
    private readonly string _stateDirectory;
    private readonly ServiceProvider _services;

    /// <summary>Builds a provider rooted at a temporary directory, never the real user profile.</summary>
    public SetupWizardRecoveryTests()
    {
        _stateDirectory = Path.Combine(
            Path.GetTempPath(), "splmb-tests", Guid.NewGuid().ToString("n"));

        _services = ServiceRegistration.Build(new ApplicationPaths(_stateDirectory));
    }

    /// <inheritdoc />
    public void Dispose()
    {
        _services.Dispose();

        if (Directory.Exists(_stateDirectory))
        {
            Directory.Delete(_stateDirectory, recursive: true);
        }
    }

    private TenantSetupViewModel Wizard() => _services.GetRequiredService<TenantSetupViewModel>();

    /// <summary>
    /// An in-flight sign-in must be cancellable. Without this the generated AsyncRelayCommand
    /// reports CanExecute false for as long as it runs, so a sign-in that never completes
    /// disables the only button that could retry it.
    /// </summary>
    [Fact]
    public void SignIn_ExposesACancelCommand()
    {
        var wizard = Wizard();

        Assert.NotNull(wizard.SignInCancelCommand);
        Assert.NotNull(wizard.RequestConsentCancelCommand);
    }

    /// <summary>
    /// Cancelling is safe when nothing is running, since the button is bound for the whole page
    /// rather than being created on demand.
    /// </summary>
    [Fact]
    public void SignInCancel_IsHarmlessWhenNothingIsRunning()
    {
        var wizard = Wizard();

        wizard.SignInCancelCommand.Execute(null);
        wizard.RequestConsentCancelCommand.Execute(null);

        Assert.False(wizard.IsBusy);
    }

    /// <summary>
    /// This build ships no bootstrap client ID, so automatic setup is not *ready*. It must still
    /// be *selectable*, because the field that supplies a bootstrap client ID is only reachable
    /// once the automatic method is chosen.
    /// </summary>
    [Fact]
    public void AutomaticSetup_IsSelectableEvenWhenNoBootstrapClientIdIsConfigured()
    {
        var wizard = Wizard();

        Assert.False(wizard.IsAutomaticSetupAvailable);

        wizard.Method = SetupMethod.Automatic;

        Assert.Equal(SetupMethod.Automatic, wizard.Method);
    }

    /// <summary>
    /// Selectable must not mean runnable. Creating a registration without a bootstrap client ID
    /// has to refuse, and say why, rather than attempting a call it cannot authenticate.
    /// </summary>
    [Fact]
    public async Task CreateRegistration_WithoutABootstrapClientId_RefusesAndExplains()
    {
        var wizard = Wizard();
        wizard.Method = SetupMethod.Automatic;
        wizard.BootstrapClientIdOverride = string.Empty;

        await wizard.CreateRegistrationCommand.ExecuteAsync(null);

        Assert.False(wizard.IsBusy);
        Assert.NotNull(wizard.ErrorMessage);
        Assert.Contains("bootstrap", wizard.ErrorMessage!, StringComparison.OrdinalIgnoreCase);
    }

    /// <summary>
    /// The wizard must not advance past sign-in without an account, and must say so rather than
    /// failing later in a less obvious place.
    /// </summary>
    [Fact]
    public async Task Next_OnSignInPageWithoutAnAccount_RefusesWithAnExplanation()
    {
        var wizard = Wizard();
        wizard.CurrentPage = SetupWizardPage.SignIn;

        await wizard.NextCommand.ExecuteAsync(null);

        Assert.Equal(SetupWizardPage.SignIn, wizard.CurrentPage);
        Assert.NotNull(wizard.ErrorMessage);
    }

    /// <summary>
    /// A correctable mistake must stay correctable. Typing a bad client ID, being refused, then
    /// fixing it must leave sign-in runnable — this is the state the reporter got stuck in.
    /// </summary>
    [Fact]
    public async Task SignIn_AfterAValidationFailure_CanBeRetried()
    {
        var wizard = Wizard();
        wizard.CurrentPage = SetupWizardPage.SignIn;
        wizard.Method = SetupMethod.ExistingRegistration;
        wizard.TenantId = "11111111-1111-1111-1111-111111111111";
        wizard.ExistingClientId = "not-a-guid";

        await wizard.SignInCommand.ExecuteAsync(null);

        Assert.NotNull(wizard.ErrorMessage);
        Assert.False(wizard.IsBusy);

        // The correction, and the retry it must permit.
        wizard.ExistingClientId = "22222222-2222-2222-2222-222222222222";

        Assert.True(wizard.SignInCommand.CanExecute(null));
        Assert.True(wizard.CanGoBack);
    }
}
