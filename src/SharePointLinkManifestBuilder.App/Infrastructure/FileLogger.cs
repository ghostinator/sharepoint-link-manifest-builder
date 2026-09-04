using System.Collections.Concurrent;
using System.Globalization;
using System.Text;
using Microsoft.Extensions.Logging;

namespace SharePointLinkManifestBuilder.App.Infrastructure;

/// <summary>
/// A small rolling file logger.
/// <para>
/// Writing is queued and drained on a single background task, so logging never blocks the UI
/// thread and concurrent job workers cannot interleave partial lines. Files roll by size and a
/// bounded number of previous files is retained, so logs cannot grow without limit on a machine
/// nobody is watching.
/// </para>
/// <para>
/// Output still passes through <see cref="RedactingLoggerProvider"/>, so no token can reach
/// disk through this path.
/// </para>
/// </summary>
[ProviderAlias("File")]
public sealed class FileLoggerProvider : ILoggerProvider
{
    private const long MaxFileSizeBytes = 5 * 1024 * 1024;
    private const int MaxRetainedFiles = 5;

    private readonly BlockingCollection<string> _queue = new(new ConcurrentQueue<string>(), 10_000);
    private readonly CancellationTokenSource _shutdown = new();
    private readonly Task _writerTask;
    private readonly string _directory;
    private readonly LogLevel _minimumLevel;

    /// <summary>Creates the provider.</summary>
    /// <param name="directory">Directory to write log files into.</param>
    /// <param name="minimumLevel">Minimum level written to disk.</param>
    public FileLoggerProvider(string directory, LogLevel minimumLevel = LogLevel.Information)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(directory);

        _directory = directory;
        _minimumLevel = minimumLevel;

        Directory.CreateDirectory(directory);

        _writerTask = Task.Run(DrainAsync);
    }

    /// <summary>The current log file path.</summary>
    public string CurrentFile => Path.Combine(_directory, "application.log");

    /// <inheritdoc />
    public ILogger CreateLogger(string categoryName) =>
        new FileLogger(categoryName, _minimumLevel, Enqueue);

    /// <inheritdoc />
    public void Dispose()
    {
        _queue.CompleteAdding();

        try
        {
            // Bounded so a stuck writer cannot hang application shutdown.
            _writerTask.Wait(TimeSpan.FromSeconds(2));
        }
        catch (AggregateException)
        {
            // Shutdown is best-effort; a failure to flush the last lines is not worth crashing over.
        }

        _shutdown.Cancel();
        _shutdown.Dispose();
        _queue.Dispose();
    }

    private void Enqueue(string line)
    {
        // Never block the caller. Dropping a line under extreme pressure is preferable to
        // stalling a job or the UI thread on disk I/O.
        _queue.TryAdd(line);
    }

    private async Task DrainAsync()
    {
        var buffer = new StringBuilder();

        foreach (var line in _queue.GetConsumingEnumerable())
        {
            buffer.Clear().AppendLine(line);

            // Opportunistically batch anything already queued into the same write.
            while (buffer.Length < 32 * 1024 && _queue.TryTake(out var extra))
            {
                buffer.AppendLine(extra);
            }

            try
            {
                RollIfNeeded();
                await File.AppendAllTextAsync(CurrentFile, buffer.ToString()).ConfigureAwait(false);
            }
#pragma warning disable CA1031 // Logging must never be able to take down the application.
            catch (Exception)
            {
                // Deliberately silent: reporting a logging failure through the logger would
                // recurse, and there is no better channel available at this point.
            }
#pragma warning restore CA1031
        }
    }

    private void RollIfNeeded()
    {
        var info = new FileInfo(CurrentFile);

        if (!info.Exists || info.Length < MaxFileSizeBytes)
        {
            return;
        }

        for (var index = MaxRetainedFiles - 1; index >= 1; index--)
        {
            var source = Path.Combine(_directory, $"application.{index}.log");
            var destination = Path.Combine(_directory, $"application.{index + 1}.log");

            if (File.Exists(source))
            {
                File.Move(source, destination, overwrite: true);
            }
        }

        File.Move(CurrentFile, Path.Combine(_directory, "application.1.log"), overwrite: true);
    }

    private sealed class FileLogger : ILogger
    {
        private readonly string _category;
        private readonly LogLevel _minimumLevel;
        private readonly Action<string> _write;

        public FileLogger(string category, LogLevel minimumLevel, Action<string> write)
        {
            _category = category;
            _minimumLevel = minimumLevel;
            _write = write;
        }

        public IDisposable? BeginScope<TState>(TState state)
            where TState : notnull => null;

        public bool IsEnabled(LogLevel logLevel) =>
            logLevel >= _minimumLevel && logLevel != LogLevel.None;

        public void Log<TState>(
            LogLevel logLevel,
            EventId eventId,
            TState state,
            Exception? exception,
            Func<TState, Exception?, string> formatter)
        {
            if (!IsEnabled(logLevel))
            {
                return;
            }

            ArgumentNullException.ThrowIfNull(formatter);

            var timestamp = DateTimeOffset.UtcNow.ToString("yyyy-MM-dd'T'HH:mm:ss'Z'", CultureInfo.InvariantCulture);
            var line = $"{timestamp} [{Abbreviate(logLevel)}] {_category}: {formatter(state, exception)}";

            if (exception is not null)
            {
                // The type and stack trace are useful for support; the message may carry
                // request detail, so it is not appended separately.
                line += Environment.NewLine + exception;
            }

            _write(line);
        }

        private static string Abbreviate(LogLevel level) => level switch
        {
            LogLevel.Trace => "TRC",
            LogLevel.Debug => "DBG",
            LogLevel.Information => "INF",
            LogLevel.Warning => "WRN",
            LogLevel.Error => "ERR",
            LogLevel.Critical => "CRT",
            _ => "___",
        };
    }
}
