using SharePointLinkManifestBuilder.Core.Resilience;

namespace SharePointLinkManifestBuilder.Core.Tests.Resilience;

public class GraphRetryPolicyTests
{
    /// <summary>Jitter is injected so backoff is deterministic; 1.0 yields the full ceiling.</summary>
    private static GraphRetryPolicy Policy(int maxAttempts = 5, double jitter = 1.0) =>
        new(maxAttempts, TimeSpan.FromSeconds(1), TimeSpan.FromSeconds(60), () => jitter);

    [Theory]
    [InlineData(429)]
    [InlineData(503)]
    [InlineData(504)]
    [InlineData(502)]
    [InlineData(408)]
    public void Evaluate_RetryableStatus_Retries(int status) =>
        Assert.True(Policy().Evaluate(1, status).ShouldRetry);

    [Theory]
    [InlineData(400)]
    [InlineData(401)]
    [InlineData(403)]
    [InlineData(404)]
    [InlineData(409)]
    [InlineData(413)]
    public void Evaluate_NonRetryableStatus_Stops(int status) =>
        Assert.False(Policy().Evaluate(1, status).ShouldRetry);

    /// <summary>
    /// 412 is an ETag conflict. Retrying it would either fail identically or, worse, succeed
    /// against a changed remote file. It must reach the manifest conflict policy instead.
    /// </summary>
    [Fact]
    public void Evaluate_PreconditionFailed_IsNeverRetried()
    {
        var decision = Policy().Evaluate(1, 412);

        Assert.False(decision.ShouldRetry);
        Assert.Contains("412", decision.Reason, StringComparison.Ordinal);
    }

    /// <summary>Retry-After is authoritative; ignoring it in favour of a shorter local backoff makes throttling worse.</summary>
    [Fact]
    public void Evaluate_RetryAfterHeader_TakesPrecedenceOverBackoff()
    {
        var decision = Policy().Evaluate(1, 429, TimeSpan.FromSeconds(17));

        Assert.True(decision.ShouldRetry);
        Assert.Equal(TimeSpan.FromSeconds(17), decision.Delay);
        Assert.Contains("Retry-After", decision.Reason, StringComparison.Ordinal);
    }

    [Fact]
    public void Evaluate_RetryAfterLongerThanMax_IsCapped()
    {
        var decision = Policy().Evaluate(1, 429, TimeSpan.FromMinutes(30));

        Assert.Equal(TimeSpan.FromSeconds(60), decision.Delay);
    }

    [Fact]
    public void Evaluate_AtAttemptLimit_Stops()
    {
        var decision = Policy(maxAttempts: 3).Evaluate(3, 429);

        Assert.False(decision.ShouldRetry);
        Assert.Contains("limit", decision.Reason, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Evaluate_ZeroMaxAttempts_DisablesRetry() =>
        Assert.False(Policy(maxAttempts: 0).Evaluate(1, 429).ShouldRetry);

    [Fact]
    public void ComputeBackoff_GrowsExponentially()
    {
        var policy = Policy();

        Assert.Equal(TimeSpan.FromSeconds(1), policy.ComputeBackoff(1));
        Assert.Equal(TimeSpan.FromSeconds(2), policy.ComputeBackoff(2));
        Assert.Equal(TimeSpan.FromSeconds(4), policy.ComputeBackoff(3));
        Assert.Equal(TimeSpan.FromSeconds(8), policy.ComputeBackoff(4));
    }

    [Fact]
    public void ComputeBackoff_IsCappedAtMaxDelay() =>
        Assert.Equal(TimeSpan.FromSeconds(60), Policy().ComputeBackoff(20));

    /// <summary>
    /// Full jitter means the wait is uniform in [0, ceiling], so many workers do not retry in
    /// lockstep and immediately re-throttle the tenant.
    /// </summary>
    [Fact]
    public void ComputeBackoff_AppliesFullJitter()
    {
        Assert.Equal(TimeSpan.Zero, Policy(jitter: 0.0).ComputeBackoff(5));
        Assert.Equal(TimeSpan.FromSeconds(8), Policy(jitter: 0.5).ComputeBackoff(5));
        Assert.Equal(TimeSpan.FromSeconds(16), Policy(jitter: 1.0).ComputeBackoff(5));
    }

    [Fact]
    public void ComputeBackoff_LargeAttemptNumber_DoesNotOverflow()
    {
        var delay = Policy().ComputeBackoff(int.MaxValue);

        Assert.True(delay >= TimeSpan.Zero);
        Assert.True(delay <= TimeSpan.FromSeconds(60));
    }

    [Fact]
    public void EvaluateTransport_TransientFailure_Retries() =>
        Assert.True(Policy().EvaluateTransport(1, isTransient: true).ShouldRetry);

    [Fact]
    public void EvaluateTransport_NonTransientFailure_Stops() =>
        Assert.False(Policy().EvaluateTransport(1, isTransient: false).ShouldRetry);

    [Fact]
    public void EvaluateTransport_AtAttemptLimit_Stops() =>
        Assert.False(Policy(maxAttempts: 2).EvaluateTransport(2, isTransient: true).ShouldRetry);
}
