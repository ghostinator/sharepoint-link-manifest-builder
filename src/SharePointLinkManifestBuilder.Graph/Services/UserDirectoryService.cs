using Microsoft.Extensions.Logging;
using SharePointLinkManifestBuilder.Core.Abstractions;
using SharePointLinkManifestBuilder.Core.Models;
using SharePointLinkManifestBuilder.Graph.Dto;
using SharePointLinkManifestBuilder.Graph.Http;

namespace SharePointLinkManifestBuilder.Graph.Services;

/// <summary>
/// Directory lookup backing the User OneDrive people picker.
/// <para>
/// This service returns exactly what Microsoft Graph returns for the signed-in user and adds
/// nothing. It cannot reveal information the caller is not authorized to read, and finding a
/// user here does not imply their OneDrive is accessible.
/// </para>
/// </summary>
public sealed class UserDirectoryService : IUserDirectoryService
{
    private readonly IGraphApiClient _client;
    private readonly ILogger<UserDirectoryService> _logger;

    /// <summary>Creates the service.</summary>
    public UserDirectoryService(IGraphApiClient client, ILogger<UserDirectoryService> logger)
    {
        _client = client ?? throw new ArgumentNullException(nameof(client));
        _logger = logger ?? throw new ArgumentNullException(nameof(logger));
    }

    /// <inheritdoc />
    public async Task<OperationResult<IReadOnlyList<OneDriveUser>>> SearchUsersAsync(
        string query,
        CancellationToken cancellationToken = default)
    {
        if (string.IsNullOrWhiteSpace(query) || query.Trim().Length < 2)
        {
            // Refusing a one-character query is deliberate: it would return an arbitrary slice
            // of the directory and waste a request on a result nobody can use.
            return OperationResult<IReadOnlyList<OneDriveUser>>.Success([]);
        }

        try
        {
            var users = new List<OneDriveUser>();

            await foreach (var dto in _client
                .EnumeratePagedAsync<GraphUserDto>(GraphPaths.SearchUsers(query.Trim()), cancellationToken)
                .ConfigureAwait(false))
            {
                if (Map(dto) is { } user)
                {
                    users.Add(user);
                }
            }

            _logger.LogInformation("Directory search matched {Count} user(s).", users.Count);
            return OperationResult<IReadOnlyList<OneDriveUser>>.Success(users);
        }
        catch (GraphOperationException ex)
        {
            return OperationResult<IReadOnlyList<OneDriveUser>>.Failure(ex.Error);
        }
    }

    /// <inheritdoc />
    public async Task<OperationResult<OneDriveUser>> GetUserAsync(
        string userIdOrUpn,
        CancellationToken cancellationToken = default)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(userIdOrUpn);

        var response = await _client
            .GetAsync<GraphUserDto>(GraphPaths.User(userIdOrUpn), cancellationToken)
            .ConfigureAwait(false);

        if (!response.Succeeded || response.Value is null)
        {
            return OperationResult<OneDriveUser>.Failure(
                response.Error ?? GraphErrorMapper.Map(response.StatusCode, null, null, "read a user"));
        }

        return Map(response.Value) is { } user
            ? OperationResult<OneDriveUser>.Success(user)
            : OperationResult<OneDriveUser>.Failure(new GraphError
            {
                Kind = GraphErrorKind.Unknown,
                Message = "Microsoft Graph returned a user without an identifier.",
            });
    }

    /// <summary>
    /// Maps a directory user. Only fields Graph actually returned are surfaced; nothing is
    /// inferred or filled in, so the picker never displays a value the directory withheld.
    /// </summary>
    internal static OneDriveUser? Map(GraphUserDto dto)
    {
        if (string.IsNullOrEmpty(dto.Id))
        {
            return null;
        }

        return new OneDriveUser
        {
            UserId = dto.Id,
            DisplayName = dto.DisplayName ?? dto.UserPrincipalName ?? dto.Id,
            UserPrincipalName = dto.UserPrincipalName,
            JobTitle = dto.JobTitle,
        };
    }
}
