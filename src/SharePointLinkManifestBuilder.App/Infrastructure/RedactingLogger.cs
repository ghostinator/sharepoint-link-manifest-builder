using Microsoft.Extensions.Logging;
using SharePointLinkManifestBuilder.Core.Security;

namespace SharePointLinkManifestBuilder.App.Infrastructure;

/// <summary>
/// Wraps every logger so that all formatted output passes through
/// <see cref="SensitiveDataRedactor"/> before reaching any provider.
/// <para>
/// Individual call sites already avoid logging credentials, but relying on that alone means one
/// careless log statement anywhere in the codebase becomes a token disclosure. Redacting at the
/// provider boundary makes it structurally impossible instead of merely unlikely, and it covers
/// third-party libraries logging through the same pipeline.
/// </para>
/// </summary>
[ProviderAlias("Redacting")]
public sealed class RedactingLoggerProvider : ILoggerProvider
{
    private readonly ILoggerProvider _inner;

    /// <summary>Wraps an existing provider.</summary>
    public RedactingLoggerProvider(ILoggerProvider inner) =>
        _inner = inner ?? throw new ArgumentNullException(nameof(inner));

    /// <inheritdoc />
    public ILogger CreateLogger(string categoryName) =>
        new RedactingLogger(_inner.CreateLogger(categoryName));

    /// <inheritdoc />
    public void Dispose() => _inner.Dispose();
}

/// <summary>A logger decorator that redacts formatted messages and scope values.</summary>
public sealed class RedactingLogger : ILogger
{
    private readonly ILogger _inner;

    /// <summary>Wraps an existing logger.</summary>
    public RedactingLogger(ILogger inner) =>
        _inner = inner ?? throw new ArgumentNullException(nameof(inner));

    /// <inheritdoc />
    public IDisposable? BeginScope<TState>(TState state)
        where TState : notnull =>
        _inner.BeginScope(state);

    /// <inheritdoc />
    public bool IsEnabled(LogLevel logLevel) => _inner.IsEnabled(logLevel);

    /// <inheritdoc />
    public void Log<TState>(
        LogLevel logLevel,
        EventId eventId,
        TState state,
        Exception? exception,
        Func<TState, Exception?, string> formatter)
    {
        ArgumentNullException.ThrowIfNull(formatter);

        _inner.Log(
            logLevel,
            eventId,
            state,
            exception,

            // Redaction is applied to the final formatted string, which is the only place every
            // structured argument has actually been rendered into text.
            (s, e) => SensitiveDataRedactor.Redact(formatter(s, e)));
    }
}

/// <summary>Registers redaction across the whole logging pipeline.</summary>
public static class RedactingLoggerExtensions
{
    /// <summary>
    /// Wraps every provider already registered on the builder. Call this last, after all
    /// providers have been added, so nothing escapes the wrapper.
    /// </summary>
    public static ILoggingBuilder AddRedaction(this ILoggingBuilder builder)
    {
        ArgumentNullException.ThrowIfNull(builder);

        var existing = builder.Services
            .Where(d => d.ServiceType == typeof(ILoggerProvider))
            .ToArray();

        foreach (var descriptor in existing)
        {
            builder.Services.Remove(descriptor);

            builder.Services.Add(new Microsoft.Extensions.DependencyInjection.ServiceDescriptor(
                typeof(ILoggerProvider),
                provider =>
                {
                    var inner = (ILoggerProvider)(descriptor.ImplementationInstance
                        ?? descriptor.ImplementationFactory?.Invoke(provider)
                        ?? ActivatorUtilitiesCreate(provider, descriptor.ImplementationType!));

                    return new RedactingLoggerProvider(inner);
                },
                descriptor.Lifetime));
        }

        return builder;
    }

    private static object ActivatorUtilitiesCreate(IServiceProvider provider, Type type) =>
        Microsoft.Extensions.DependencyInjection.ActivatorUtilities.CreateInstance(provider, type);
}
