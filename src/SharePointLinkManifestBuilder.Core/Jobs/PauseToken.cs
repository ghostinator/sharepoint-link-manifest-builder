using SharePointLinkManifestBuilder.Core.Abstractions;

namespace SharePointLinkManifestBuilder.Core.Jobs;

/// <summary>
/// A cooperative pause signal, distinct from cancellation.
/// <para>
/// Pausing is only honoured at safe points between operations, so it can never interrupt an
/// in-flight Graph write and leave a half-created sharing link.
/// </para>
/// </summary>
public sealed class PauseToken : IPauseToken, IDisposable
{
    private readonly Lock _gate = new();
    private TaskCompletionSource _resumeSignal =
        new(TaskCreationOptions.RunContinuationsAsynchronously);

    private bool _isPaused;
    private bool _disposed;

    /// <summary>Creates a token in the running state.</summary>
    public PauseToken()
    {
        // Start unpaused: the signal is already complete, so waiting is free.
        _resumeSignal.SetResult();
    }

    /// <inheritdoc />
    public bool IsPaused
    {
        get
        {
            lock (_gate)
            {
                return _isPaused;
            }
        }
    }

    /// <summary>Raised when the paused state changes, so the UI can reflect it.</summary>
    public event EventHandler<bool>? PauseStateChanged;

    /// <summary>Pauses the run at the next safe point.</summary>
    public void Pause()
    {
        bool changed;

        lock (_gate)
        {
            ObjectDisposedException.ThrowIf(_disposed, this);

            changed = !_isPaused;
            if (changed)
            {
                _isPaused = true;
                _resumeSignal = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
            }
        }

        if (changed)
        {
            PauseStateChanged?.Invoke(this, true);
        }
    }

    /// <summary>Resumes a paused run.</summary>
    public void Resume()
    {
        TaskCompletionSource? toRelease = null;

        lock (_gate)
        {
            ObjectDisposedException.ThrowIf(_disposed, this);

            if (_isPaused)
            {
                _isPaused = false;
                toRelease = _resumeSignal;
            }
        }

        if (toRelease is not null)
        {
            toRelease.TrySetResult();
            PauseStateChanged?.Invoke(this, false);
        }
    }

    /// <inheritdoc />
    public Task WaitWhilePausedAsync(CancellationToken cancellationToken = default)
    {
        Task signal;

        lock (_gate)
        {
            if (!_isPaused)
            {
                return Task.CompletedTask;
            }

            signal = _resumeSignal.Task;
        }

        return signal.WaitAsync(cancellationToken);
    }

    /// <summary>Releases any waiter so a disposed token cannot deadlock a run.</summary>
    public void Dispose()
    {
        lock (_gate)
        {
            if (_disposed)
            {
                return;
            }

            _disposed = true;
            _isPaused = false;
            _resumeSignal.TrySetResult();
        }
    }
}
