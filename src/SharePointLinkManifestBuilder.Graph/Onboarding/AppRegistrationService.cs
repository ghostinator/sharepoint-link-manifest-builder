using Microsoft.Extensions.Logging;
using SharePointLinkManifestBuilder.Core.Abstractions;
using SharePointLinkManifestBuilder.Core.Models;
using SharePointLinkManifestBuilder.Graph.Dto;
using SharePointLinkManifestBuilder.Graph.Http;
using SharePointLinkManifestBuilder.Graph.Services;

namespace SharePointLinkManifestBuilder.Graph.Onboarding;

/// <summary>
/// Creates and inspects the tenant-specific application registration during onboarding.
/// <para>
/// The registration is created with a <em>complete</em> POST body — display name, sign-in
/// audience, public-client flag, redirect URIs and requiredResourceAccess all at once. This is
/// deliberate and is what keeps the bootstrap identity least-privileged: <c>POST /applications</c>
/// accepts <c>AppRegistration.Create</c>, whereas <c>PATCH /applications/{id}</c> would demand
/// <c>Application.ReadWrite.All</c>. Getting it right first time avoids ever needing the wider
/// permission. See docs/adr/0005-bootstrap-application-model.md.
/// </para>
/// <para>
/// Nothing here happens silently. Every method that changes the tenant is called only after the
/// caller has shown the user exactly what will change, and every change is written to the local
/// sanitized audit history.
/// </para>
/// </summary>
public sealed class AppRegistrationService : IAppRegistrationService
{
    private readonly IGraphApiClient _client;
    private readonly ILogger<AppRegistrationService> _logger;

    /// <summary>Creates the service.</summary>
    public AppRegistrationService(IGraphApiClient client, ILogger<AppRegistrationService> logger)
    {
        _client = client ?? throw new ArgumentNullException(nameof(client));
        _logger = logger ?? throw new ArgumentNullException(nameof(logger));
    }

    /// <inheritdoc />
    public Task<RegistrationCapability> EstimateCapabilityAsync(CancellationToken cancellationToken = default)
    {
        // Deliberately an estimate, not a probe. The only reliable way to learn whether a user
        // may create an application is to try, and a speculative create would be exactly the
        // kind of silent tenant modification this product forbids. Directory-role inspection
        // would need a permission the default bootstrap tier does not request.
        _logger.LogInformation(
            "Registration capability cannot be determined without attempting the operation; "
            + "the wizard will explain the possible outcomes instead.");

        return Task.FromResult(RegistrationCapability.Unknown);
    }

    /// <inheritdoc />
    public async Task<RegistrationProvisioningResult> CreateRegistrationAsync(
        AppRegistrationConfiguration configuration,
        string tenantId,
        CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(configuration);
        ArgumentException.ThrowIfNullOrWhiteSpace(tenantId);

        var requiredResourceAccess = GraphPermissionIds.BuildRequiredResourceAccess(
            configuration.RequestedPermissions, out var unmapped);

        var changes = new List<string>(configuration.DescribePlannedChanges());

        if (unmapped.Count > 0)
        {
            // Reported rather than guessed at. A wrong permission GUID would silently configure
            // the wrong permission, which is worse than configuring none.
            changes.Add(
                $"The following permission(s) could not be configured automatically and must be added by hand: "
                + string.Join(", ", unmapped));

            _logger.LogWarning(
                "No known identifier for {Count} requested permission(s); they will not be pre-configured.",
                unmapped.Count);
        }

        var request = new CreateApplicationRequest
        {
            DisplayName = configuration.DisplayName,
            SignInAudience = configuration.SignInAudience,
            IsFallbackPublicClient = configuration.IsFallbackPublicClient,
            PublicClient = new GraphPublicClientDto { RedirectUris = configuration.RedirectUris },
            RequiredResourceAccess = requiredResourceAccess,
        };

        _logger.LogInformation(
            "Creating an application registration named '{DisplayName}' in tenant {TenantId}. "
            + "No client secret will be created.",
            configuration.DisplayName,
            tenantId);

        var response = await _client
            .PostAsync<GraphApplicationDto>(GraphPaths.Applications(), request, cancellationToken)
            .ConfigureAwait(false);

        if (!response.Succeeded || response.Value?.AppId is null)
        {
            var error = MapRegistrationError(response.Error, response.StatusCode);

            _logger.LogWarning("Application registration was not created: {Message}", error.Message);

            return new RegistrationProvisioningResult
            {
                Succeeded = false,
                Error = error,
                ChangesApplied = [],
            };
        }

        var created = response.Value;

        _logger.LogInformation(
            "Application registration created. A service principal will be provisioned by Microsoft when "
            + "consent is granted.");

        return new RegistrationProvisioningResult
        {
            Succeeded = true,
            ApplicationObjectId = created.Id,
            ClientId = created.AppId,
            ChangesApplied = changes,
            Configuration = new TenantConfiguration
            {
                TenantId = tenantId,
                ClientId = created.AppId,
                ApplicationDisplayName = created.DisplayName ?? configuration.DisplayName,
                ApplicationObjectId = created.Id,
                Source = RegistrationSource.AutomaticSetup,
                ConsentState = ConsentState.Unknown,
                RequiredScopes = configuration.RequestedPermissions.Select(p => p.Scope).ToArray(),
            },
        };
    }

    /// <inheritdoc />
    public async Task<OperationResult<RegistrationVerification>> InspectRegistrationAsync(
        string clientId,
        CancellationToken cancellationToken = default)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(clientId);

        var notVerified = new List<string>();
        var warnings = new List<string>();

        var application = await _client
            .GetAsync<GraphApplicationDto>(GraphPaths.ApplicationByAppId(clientId), cancellationToken)
            .ConfigureAwait(false);

        if (!application.Succeeded)
        {
            // Almost always means the elevated read permission is not granted, which is the
            // normal state for this product. It is not an error the user needs to fix.
            return OperationResult<RegistrationVerification>.Success(new RegistrationVerification
            {
                ApplicationFound = false,
                NotVerified =
                [
                    "The application registration could not be read. This is expected unless an elevated "
                    + "directory read permission has been granted; it does not mean the registration is missing.",
                ],
            });
        }

        var app = application.Value;
        var redirectOk = app?.PublicClient?.RedirectUris?
            .Any(u => u.StartsWith("http://localhost", StringComparison.OrdinalIgnoreCase)) ?? false;

        if (app?.PasswordCredentials is { Count: > 0 })
        {
            // Worth flagging loudly: this application never creates one, so its presence means
            // somebody else added a secret to a registration used by a public client.
            warnings.Add(
                $"This registration has {app.PasswordCredentials.Count} client secret(s). "
                + "This application never creates or uses one; consider removing them.");
        }

        var servicePrincipal = await _client
            .GetAsync<GraphServicePrincipalDto>(GraphPaths.ServicePrincipalByAppId(clientId), cancellationToken)
            .ConfigureAwait(false);

        if (!servicePrincipal.Succeeded)
        {
            notVerified.Add("The service principal could not be read with the current permissions.");
        }

        return OperationResult<RegistrationVerification>.Success(new RegistrationVerification
        {
            ApplicationFound = true,
            ServicePrincipalFound = servicePrincipal.Succeeded && servicePrincipal.Value?.Id is not null,
            IsPublicClient = app?.IsFallbackPublicClient ?? false,
            RedirectUriConfigured = redirectOk,
            Warnings = warnings,
            NotVerified = notVerified,
        });
    }

    /// <inheritdoc />
    public async Task<OperationResult<string>> EnsureServicePrincipalAsync(
        string clientId,
        CancellationToken cancellationToken = default)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(clientId);

        var existing = await _client
            .GetAsync<GraphServicePrincipalDto>(GraphPaths.ServicePrincipalByAppId(clientId), cancellationToken)
            .ConfigureAwait(false);

        if (existing.Succeeded && existing.Value?.Id is { } id)
        {
            return OperationResult<string>.Success(id);
        }

        _logger.LogInformation(
            "Creating a service principal explicitly. This is normally unnecessary: consenting through "
            + "Microsoft's endpoint provisions one automatically.");

        var response = await _client
            .PostAsync<GraphServicePrincipalDto>(
                GraphPaths.ServicePrincipals(),
                new CreateServicePrincipalRequest { AppId = clientId },
                cancellationToken)
            .ConfigureAwait(false);

        if (!response.Succeeded || response.Value?.Id is null)
        {
            return OperationResult<string>.Failure(
                MapRegistrationError(response.Error, response.StatusCode));
        }

        return OperationResult<string>.Success(response.Value.Id);
    }

    /// <inheritdoc />
    public async Task<RegistrationProvisioningResult> RepairRegistrationAsync(
        string applicationObjectId,
        AppRegistrationConfiguration desired,
        CancellationToken cancellationToken = default)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(applicationObjectId);
        ArgumentNullException.ThrowIfNull(desired);

        var requiredResourceAccess = GraphPermissionIds.BuildRequiredResourceAccess(
            desired.RequestedPermissions, out var unmapped);

        var patch = new
        {
            isFallbackPublicClient = desired.IsFallbackPublicClient,
            publicClient = new { redirectUris = desired.RedirectUris },
            requiredResourceAccess,
        };

        var response = await _client
            .PatchAsync<GraphApplicationDto>(
                GraphPaths.ApplicationByObjectId(applicationObjectId), patch, cancellationToken)
            .ConfigureAwait(false);

        if (!response.Succeeded)
        {
            return new RegistrationProvisioningResult
            {
                Succeeded = false,
                Error = MapRegistrationError(response.Error, response.StatusCode),
            };
        }

        var changes = new List<string>
        {
            "Enabled public client behaviour.",
            $"Set redirect URI(s) to {string.Join(", ", desired.RedirectUris)}.",
            $"Set {desired.RequestedPermissions.Count} delegated Microsoft Graph permission(s).",
        };

        if (unmapped.Count > 0)
        {
            changes.Add($"Could not configure: {string.Join(", ", unmapped)}.");
        }

        return new RegistrationProvisioningResult
        {
            Succeeded = true,
            ApplicationObjectId = applicationObjectId,
            ChangesApplied = changes,
        };
    }

    /// <inheritdoc />
    public async Task<OperationResult<bool>> DeleteRegistrationAsync(
        string applicationObjectId,
        string confirmedDisplayName,
        CancellationToken cancellationToken = default)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(applicationObjectId);
        ArgumentException.ThrowIfNullOrWhiteSpace(confirmedDisplayName);

        // The confirmation is re-checked here, at the point of no return, rather than trusted
        // from the UI. A destructive tenant operation should not depend on a caller having
        // remembered to validate.
        var application = await _client
            .GetAsync<GraphApplicationDto>(
                GraphPaths.ApplicationByObjectId(applicationObjectId), cancellationToken)
            .ConfigureAwait(false);

        if (!application.Succeeded || application.Value?.DisplayName is null)
        {
            return OperationResult<bool>.Failure(new GraphError
            {
                Kind = GraphErrorKind.AppRegistrationNotFound,
                Message = "The registration could not be read, so it was not deleted.",
                SuggestedAction = "Deleting a registration requires an elevated permission "
                    + "(Application.ReadWrite.All).",
            });
        }

        if (!string.Equals(application.Value.DisplayName, confirmedDisplayName, StringComparison.Ordinal))
        {
            _logger.LogWarning("Refusing to delete a registration whose name does not match the confirmation.");

            return OperationResult<bool>.Failure(new GraphError
            {
                Kind = GraphErrorKind.Unknown,
                Message = "The name you typed does not match the registration's display name, so nothing "
                    + "was deleted.",
            });
        }

        _logger.LogWarning(
            "Deleting application registration '{DisplayName}' at the user's explicit request.",
            application.Value.DisplayName);

        var response = await _client
            .DeleteAsync(GraphPaths.ApplicationByObjectId(applicationObjectId), cancellationToken)
            .ConfigureAwait(false);

        return response.Succeeded
            ? OperationResult<bool>.Success(true)
            : OperationResult<bool>.Failure(MapRegistrationError(response.Error, response.StatusCode));
    }

    /// <summary>
    /// Refines a generic transport error into the registration-specific cases, because
    /// "you are not allowed to create app registrations in this tenant" needs a very different
    /// response from a generic access error.
    /// </summary>
    internal static GraphError MapRegistrationError(GraphError? error, int statusCode)
    {
        if (error is null)
        {
            return GraphErrorMapper.Map(statusCode, null, null, "change the application registration");
        }

        return statusCode switch
        {
            403 => error with
            {
                Kind = GraphErrorKind.InsufficientPrivilegesToCreateApplication,
                Message = "Your account is not permitted to create or change application registrations in "
                    + "this organization.",
                SuggestedAction = "Ask a Microsoft Entra administrator to run setup, or use "
                    + "'Existing app registration' with a client ID your administrators created.",
            },

            401 => error with
            {
                Kind = GraphErrorKind.AuthenticationFailed,
                Message = "Your sign-in is not valid for directory operations.",
                SuggestedAction = "Sign in again and retry.",
            },

            404 => error with
            {
                Kind = GraphErrorKind.AppRegistrationNotFound,
                Message = "The application registration could not be found in this organization.",
            },

            400 => error with
            {
                Kind = GraphErrorKind.RegistrationCreationBlocked,
                Message = "Microsoft Entra rejected the registration request.",
                SuggestedAction = "Your organization may restrict application registration. Use "
                    + "'Existing app registration' instead.",
            },

            _ => error,
        };
    }
}

/// <summary>
/// Supplies the bootstrap client ID from configuration.
/// <para>
/// This repository ships no client ID and never will: a private developer identifier committed
/// to source control is a supply-chain problem, not a convenience. Automatic setup reports
/// itself unavailable until a publisher supplies one, and the existing-registration path stays
/// fully functional in the meantime.
/// </para>
/// </summary>
public sealed class BootstrapConfigurationProvider : IBootstrapConfigurationProvider
{
    private readonly Lock _gate = new();
    private BootstrapConfiguration _current;

    /// <summary>Creates the provider.</summary>
    /// <param name="clientId">Client ID from configuration, or null when none is set.</param>
    /// <param name="instance">Identity instance, for sovereign clouds.</param>
    public BootstrapConfigurationProvider(string? clientId = null, string? instance = null) =>
        _current = new BootstrapConfiguration
        {
            ClientId = string.IsNullOrWhiteSpace(clientId) ? null : clientId.Trim(),
            Instance = string.IsNullOrWhiteSpace(instance) ? AuthorityDefaults.PublicCloudInstance : instance,
        };

    /// <inheritdoc />
    public BootstrapConfiguration Current
    {
        get
        {
            lock (_gate)
            {
                return _current;
            }
        }
    }

    /// <inheritdoc />
    public void SetClientId(string? clientId)
    {
        lock (_gate)
        {
            _current = _current with
            {
                ClientId = string.IsNullOrWhiteSpace(clientId) ? null : clientId.Trim(),
            };
        }
    }
}
