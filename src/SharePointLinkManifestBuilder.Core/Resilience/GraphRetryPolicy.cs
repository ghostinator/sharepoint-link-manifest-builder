using System.Net;

namespace SharePointLinkManifestBuilder.Core.Resilience;

/// <summary>What the retry policy decided about a failed attempt.</summary>
/// <param name="ShouldRetry">True when the attempt should be repeated.</param>
/// <param name="Delay">How long to wait before repeating.</param>
/// <param name="Reason">Why the decision was made, for logging and the progress display.</param>
public readonly record struct RetryDecision(bool ShouldRetry, TimeSpan Delay, string Reason)
{
    /// <summary>A decision not to retry.</summary>
    public static RetryDecision Stop(string reason) => new(false, TimeSpan.Zero, reason);
}

/// <summary>
/// Decides whether and when to repeat a failed Graph request.
/// <para>
/// Deliberately pure: it takes an attempt number, a status code and a <c>Retry-After</c> value
/// and returns a decision. Jitter is injected, so tests are deterministic and the real backoff
/// behaviour is exercised rather than mocked.
/// </para>
/// </summary>
public sealed class GraphRetryPolicy
{
    private readonly Func<double> _jitterSource;

    /// <summary>Creates a policy.</summary>
    /// <param name="maxAttempts">Total attempts including the first. Zero disables retry.</param>
    /// <param name="baseDelay">The base unit for exponential backoff.</param>
    /// <param name="maxDelay">An upper bound on any single wait.</param>
    /// <param name="jitterSource">
    /// Returns a value in [0,1). Defaults to <see cref="Random.Shared"/>. Tests inject a fixed
    /// value to make backoff deterministic.
    /// </param>
    public GraphRetryPolicy(
        int maxAttempts = 5,
        TimeSpan? baseDelay = null,
        TimeSpan? maxDelay = null,
        Func<double>? jitterSource = null)
    {
        MaxAttempts = Math.Max(0, maxAttempts);
        BaseDelay = baseDelay ?? TimeSpan.FromSeconds(1);
        MaxDelay = maxDelay ?? TimeSpan.FromSeconds(60);
        _jitterSource = jitterSource ?? Random.Shared.NextDouble;
    }

    /// <summary>Total attempts including the first.</summary>
    public int MaxAttempts { get; }

    /// <summary>The base unit for exponential backoff.</summary>
    public TimeSpan BaseDelay { get; }

    /// <summary>Upper bound on any single wait.</summary>
    public TimeSpan MaxDelay { get; }

    /// <summary>
    /// Status codes worth repeating. 429 is throttling; 502/503/504 are transient gateway and
    /// availability failures; 408 is a request timeout.
    /// </summary>
    public static readonly IReadOnlySet<int> RetryableStatusCodes = new HashSet<int>
    {
        (int)HttpStatusCode.RequestTimeout,        // 408
        (int)HttpStatusCode.TooManyRequests,       // 429
        (int)HttpStatusCode.BadGateway,            // 502
        (int)HttpStatusCode.ServiceUnavailable,    // 503
        (int)HttpStatusCode.GatewayTimeout,        // 504
    };

    /// <summary>
    /// Status codes never repeated, because repeating them cannot change the outcome and would
    /// waste the tenant's throttling budget. 412 in particular is an ETag conflict, which must
    /// surface to the conflict policy rather than being retried blindly.
    /// </summary>
    public static readonly IReadOnlySet<int> NonRetryableStatusCodes = new HashSet<int>
    {
        (int)HttpStatusCode.BadRequest,            // 400
        (int)HttpStatusCode.Unauthorized,          // 401
        (int)HttpStatusCode.Forbidden,             // 403
        (int)HttpStatusCode.NotFound,              // 404
        (int)HttpStatusCode.MethodNotAllowed,      // 405
        (int)HttpStatusCode.Conflict,              // 409
        (int)HttpStatusCode.Gone,                  // 410
        (int)HttpStatusCode.PreconditionFailed,    // 412
        (int)HttpStatusCode.RequestEntityTooLarge, // 413
        (int)HttpStatusCode.UnsupportedMediaType,  // 415
        (int)HttpStatusCode.UnprocessableEntity,   // 422
        (int)HttpStatusCode.NotImplemented,        // 501
    };

    /// <summary>
    /// Decides whether to repeat an attempt that produced an HTTP status.
    /// </summary>
    /// <param name="attemptNumber">One-based number of the attempt that just failed.</param>
    /// <param name="statusCode">The status code received.</param>
    /// <param name="retryAfter">The <c>Retry-After</c> value, when the service supplied one.</param>
    public RetryDecision Evaluate(int attemptNumber, int statusCode, TimeSpan? retryAfter = null)
    {
        if (attemptNumber >= MaxAttempts)
        {
            return RetryDecision.Stop($"Retry limit of {MaxAttempts} attempt(s) reached.");
        }

        if (NonRetryableStatusCodes.Contains(statusCode))
        {
            return RetryDecision.Stop($"HTTP {statusCode} will not change on retry.");
        }

        if (!RetryableStatusCodes.Contains(statusCode) && statusCode < 500)
        {
            return RetryDecision.Stop($"HTTP {statusCode} is not a retryable condition.");
        }

        // Retry-After is authoritative. Honouring it is how a client stays a good tenant
        // citizen, and ignoring it in favour of a shorter local backoff makes throttling worse.
        if (retryAfter is { } wait && wait > TimeSpan.Zero)
        {
            var capped = wait > MaxDelay ? MaxDelay : wait;
            return new RetryDecision(
                true,
                capped,
                $"Service requested a {capped.TotalSeconds:0.#}s wait (Retry-After).");
        }

        return new RetryDecision(
            true,
            ComputeBackoff(attemptNumber),
            $"HTTP {statusCode}: retrying with exponential backoff.");
    }

    /// <summary>
    /// Decides whether to repeat an attempt that failed with a transport exception rather than
    /// an HTTP status, such as a DNS failure or a dropped connection.
    /// </summary>
    /// <param name="attemptNumber">One-based number of the attempt that just failed.</param>
    /// <param name="isTransient">True when the exception indicates a transient network fault.</param>
    public RetryDecision EvaluateTransport(int attemptNumber, bool isTransient)
    {
        if (!isTransient)
        {
            return RetryDecision.Stop("The failure is not a transient network condition.");
        }

        if (attemptNumber >= MaxAttempts)
        {
            return RetryDecision.Stop($"Retry limit of {MaxAttempts} attempt(s) reached.");
        }

        return new RetryDecision(
            true,
            ComputeBackoff(attemptNumber),
            "Transient network failure: retrying with exponential backoff.");
    }

    /// <summary>
    /// Exponential backoff with full jitter: a wait uniformly distributed in
    /// [0, base * 2^(attempt-1)], capped. Full jitter is used rather than a fixed exponential
    /// wait because it prevents many concurrent workers from retrying in lockstep and
    /// re-throttling the tenant.
    /// </summary>
    public TimeSpan ComputeBackoff(int attemptNumber)
    {
        var exponent = Math.Max(0, attemptNumber - 1);

        // Clamp the exponent before shifting so a large attempt number cannot overflow.
        var multiplier = Math.Pow(2, Math.Min(exponent, 16));
        var ceiling = BaseDelay.TotalMilliseconds * multiplier;
        ceiling = Math.Min(ceiling, MaxDelay.TotalMilliseconds);

        var jittered = ceiling * _jitterSource();
        return TimeSpan.FromMilliseconds(Math.Max(0, jittered));
    }
}
