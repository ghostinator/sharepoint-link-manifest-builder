namespace SharePointLinkManifestBuilder.Core.Models;

/// <summary>
/// A normalized classification of every failure this application can encounter. View models
/// render this enum rather than raw exception text, so no sensitive payload reaches the UI.
/// </summary>
public enum GraphErrorKind
{
    /// <summary>Failure that does not match a known classification.</summary>
    Unknown = 0,

    // Identity and consent
    AuthenticationFailed,
    ConsentRequired,
    AdminConsentRequired,
    ConsentDenied,
    ConsentCanceled,
    UnauthorizedAdministratorRole,
    ConditionalAccessInterrupted,
    TenantMismatch,
    TokenExpired,

    // Registration/authority misconfiguration. These all surface *after* the browser has
    // rendered its "authentication complete" page, because MSAL shows that page as soon as the
    // redirect arrives, whether or not the redirect carries an error.
    /// <summary>The registration is not marked as a public client, so Entra demanded a secret.</summary>
    PublicClientNotConfigured,

    /// <summary>The loopback redirect URI is not registered, or not registered as native.</summary>
    RedirectUriMismatch,

    /// <summary>The account signed in belongs to an organization the registration does not accept.</summary>
    AccountFromUnsupportedTenant,

    /// <summary>The registration does not exist in the organization that was signed in to.</summary>
    ApplicationNotFoundInTenant,

    // Registration
    RegistrationCreationBlocked,
    InsufficientPrivilegesToCreateApplication,
    AppRegistrationNotFound,
    ServicePrincipalNotFound,
    PermissionMissing,

    // Resource access
    SharePointAccessDenied,
    OneDriveAccessDenied,
    UserDriveUnavailable,
    UserDriveUnprovisioned,
    SiteNotFound,
    LibraryNotFound,
    FolderNotFound,
    FileDeletedDuringProcessing,
    FileMovedDuringProcessing,
    InvalidUrl,

    // Sharing
    UnsupportedLinkType,
    AnonymousSharingDisabled,
    ExternalSharingDisabled,
    RecipientRejected,
    PolicyBlocked,

    // Manifest
    ManifestWriteDenied,
    ManifestConflict,
    ETagConflict,
    NameConflict,

    // Transport
    Throttled,
    NetworkOutage,
    ServiceUnavailable,
    UnsupportedItemType,
    Canceled,
}

/// <summary>
/// A sanitized, user-presentable description of a failure. Never contains tokens, authorization
/// headers, or raw response bodies.
/// </summary>
public sealed record GraphError
{
    /// <summary>Normalized classification driving UI copy and retry eligibility.</summary>
    public required GraphErrorKind Kind { get; init; }

    /// <summary>A plain-language explanation suitable for a non-developer.</summary>
    public required string Message { get; init; }

    /// <summary>HTTP status code, when the failure came from an HTTP response.</summary>
    public int? StatusCode { get; init; }

    /// <summary>The Microsoft Graph error code, such as <c>accessDenied</c>.</summary>
    public string? GraphErrorCode { get; init; }

    /// <summary>The correlation ID sent with the request, for support escalation.</summary>
    public string? ClientRequestId { get; init; }

    /// <summary>The request ID returned by the service, for support escalation.</summary>
    public string? ServiceRequestId { get; init; }

    /// <summary>True when retrying the same operation could plausibly succeed.</summary>
    public bool IsRetryable { get; init; }

    /// <summary>What the user can do about it, when there is a useful action.</summary>
    public string? SuggestedAction { get; init; }

    /// <summary>When the failure was observed.</summary>
    public DateTimeOffset TimestampUtc { get; init; } = DateTimeOffset.UtcNow;

    /// <summary>Creates a cancellation result, which is never treated as an error condition.</summary>
    public static GraphError Canceled() => new()
    {
        Kind = GraphErrorKind.Canceled,
        Message = "The operation was cancelled.",
        IsRetryable = true,
    };
}

/// <summary>
/// The result of an operation that can fail in an expected way. Using this instead of
/// exceptions for expected failures keeps per-file error handling explicit and prevents
/// failures from being silently swallowed.
/// </summary>
/// <typeparam name="T">The success value type.</typeparam>
[System.Diagnostics.CodeAnalysis.SuppressMessage(
    "Design", "CA1000:Do not declare static members on generic types",
    Justification = "Success/Failure factories on the generic result type are the conventional " +
                    "result-pattern shape and read far better at call sites than a non-generic helper.")]
public readonly record struct OperationResult<T>
{
    private OperationResult(bool succeeded, T? value, GraphError? error)
    {
        Succeeded = succeeded;
        Value = value;
        Error = error;
    }

    /// <summary>True when the operation produced a value.</summary>
    public bool Succeeded { get; }

    /// <summary>The value, present only when <see cref="Succeeded"/> is true.</summary>
    public T? Value { get; }

    /// <summary>The failure, present only when <see cref="Succeeded"/> is false.</summary>
    public GraphError? Error { get; }

    /// <summary>Creates a successful result.</summary>
    public static OperationResult<T> Success(T value) => new(true, value, null);

    /// <summary>Creates a failed result.</summary>
    public static OperationResult<T> Failure(GraphError error) => new(false, default, error);

    /// <summary>Projects the success value, preserving any failure unchanged.</summary>
    public OperationResult<TOut> Map<TOut>(Func<T, TOut> selector) =>
        Succeeded
            ? OperationResult<TOut>.Success(selector(Value!))
            : OperationResult<TOut>.Failure(Error!);
}
