using System.Net;
using System.Text;
using Microsoft.Extensions.Logging;

namespace SharePointLinkManifestBuilder.Graph.Onboarding;

/// <summary>What a redirect back from Microsoft's consent experience carried.</summary>
/// <param name="Succeeded">True when the redirect arrived, was well-formed and matched the expected state.</param>
/// <param name="TenantId">The tenant the response reports, when present.</param>
/// <param name="AdminConsentGranted">True when the response reports admin consent as granted.</param>
/// <param name="Error">The OAuth error code, when the response carried one.</param>
/// <param name="ErrorDescription">The OAuth error description, when present.</param>
/// <param name="StateMismatch">True when the returned state did not match the one sent.</param>
public readonly record struct LoopbackRedirectResult(
    bool Succeeded,
    string? TenantId,
    bool AdminConsentGranted,
    string? Error,
    string? ErrorDescription,
    bool StateMismatch);

/// <summary>
/// A single-use loopback listener that receives the redirect from Microsoft's official
/// administrator-consent experience.
/// <para>
/// Binds only <c>127.0.0.1</c> on an ephemeral port, serves exactly one request, and validates
/// the <c>state</c> parameter it generated. State validation is what stops a forged or replayed
/// redirect from being accepted as a genuine consent result.
/// </para>
/// </summary>
public sealed class LoopbackRedirectListener : IDisposable
{
    private readonly HttpListener _listener = new();
    private readonly ILogger _logger;
    private bool _disposed;

    /// <summary>Starts listening on a free loopback port.</summary>
    /// <param name="logger">Logger. Never receives the query string, which can carry a code.</param>
    public LoopbackRedirectListener(ILogger logger)
    {
        _logger = logger ?? throw new ArgumentNullException(nameof(logger));

        Port = GetFreePort();
        RedirectUri = $"http://localhost:{Port}/";

        _listener.Prefixes.Add(RedirectUri);
        _listener.Start();

        _logger.LogInformation("Listening for the consent redirect on a loopback port.");
    }

    /// <summary>The ephemeral port chosen for this attempt.</summary>
    public int Port { get; }

    /// <summary>The redirect URI to send to Microsoft Entra.</summary>
    public string RedirectUri { get; }

    /// <summary>
    /// Waits for the single redirect, validating <paramref name="expectedState"/>.
    /// </summary>
    /// <param name="expectedState">The random state generated for this attempt.</param>
    /// <param name="cancellationToken">Cancellation token; cancelling stops the listener.</param>
    public async Task<LoopbackRedirectResult> WaitForRedirectAsync(
        string expectedState,
        CancellationToken cancellationToken = default)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(expectedState);

        using var registration = cancellationToken.Register(() =>
        {
            // Aborting is what unblocks GetContextAsync; there is no cancellable overload.
            if (_listener.IsListening)
            {
                _listener.Abort();
            }
        });

        HttpListenerContext context;

        try
        {
            context = await _listener.GetContextAsync().ConfigureAwait(false);
        }
        catch (Exception ex) when (ex is HttpListenerException or ObjectDisposedException)
        {
            cancellationToken.ThrowIfCancellationRequested();

            return new LoopbackRedirectResult(
                false, null, false, "listener_failed",
                "The local listener stopped before the response arrived.", false);
        }

        var query = context.Request.QueryString;

        var state = query["state"];
        var tenantId = query["tenant"];
        var adminConsent = query["admin_consent"];
        var error = query["error"];
        var errorDescription = query["error_description"];

        var stateMatches = string.Equals(state, expectedState, StringComparison.Ordinal);

        if (!stateMatches)
        {
            _logger.LogWarning(
                "Rejecting a consent redirect whose state did not match the one this application generated.");
        }

        var granted = stateMatches
            && error is null
            && string.Equals(adminConsent, "True", StringComparison.OrdinalIgnoreCase);

        await RespondAsync(context, stateMatches, granted, error).ConfigureAwait(false);

        return new LoopbackRedirectResult(
            Succeeded: stateMatches && error is null,
            TenantId: tenantId,
            AdminConsentGranted: granted,
            Error: error,
            ErrorDescription: errorDescription,
            StateMismatch: !stateMatches);
    }

    /// <summary>Stops the listener.</summary>
    public void Dispose()
    {
        if (_disposed)
        {
            return;
        }

        _disposed = true;

        if (_listener.IsListening)
        {
            _listener.Stop();
        }

        _listener.Close();
    }

    /// <summary>
    /// Generates a cryptographically random state value. Predictable state would defeat the
    /// forgery protection entirely, so this never uses <see cref="Random"/>.
    /// </summary>
    public static string GenerateState() =>
        Convert.ToHexString(System.Security.Cryptography.RandomNumberGenerator.GetBytes(32));

    private static async Task RespondAsync(
        HttpListenerContext context,
        bool stateMatches,
        bool granted,
        string? error)
    {
        var (title, detail) = (stateMatches, granted, error) switch
        {
            (false, _, _) => (
                "Request could not be verified",
                "The response did not match the request this application sent. Nothing has been changed. "
                + "Close this tab and start the consent step again from the application."),

            (_, _, not null) => (
                "Consent was not granted",
                "Microsoft reported that the request was declined or could not be completed. "
                + "You can close this tab and return to the application."),

            (true, true, null) => (
                "Consent recorded",
                "You can close this tab and return to the application, which will now verify the result."),

            _ => (
                "Response received",
                "You can close this tab and return to the application, which will now verify the result."),
        };

        // A deliberately static page. No value from the query string is echoed back, so the
        // response cannot be used to reflect attacker-supplied content.
        // $$ raw string: interpolation is {{expr}}, so the CSS braces below stay literal.
        var html = $$"""
            <!doctype html>
            <html lang="en">
            <head>
              <meta charset="utf-8">
              <meta name="viewport" content="width=device-width, initial-scale=1">
              <title>SharePoint Link Manifest Builder</title>
              <style>
                body { font-family: system-ui, -apple-system, "Segoe UI", sans-serif;
                        margin: 0; display: grid; place-items: center; min-height: 100vh;
                        background: #f5f5f7; color: #1c1c1e; }
                main { max-width: 30rem; padding: 2rem; background: #fff; border-radius: 12px;
                        box-shadow: 0 1px 4px rgba(0,0,0,.12); }
                h1 { font-size: 1.25rem; margin: 0 0 .75rem; }
                p { margin: 0; line-height: 1.5; }
              </style>
            </head>
            <body><main><h1>{{WebUtility.HtmlEncode(title)}}</h1>
            <p>{{WebUtility.HtmlEncode(detail)}}</p></main></body>
            </html>
            """;

        var bytes = Encoding.UTF8.GetBytes(html);

        context.Response.StatusCode = 200;
        context.Response.ContentType = "text/html; charset=utf-8";
        context.Response.ContentLength64 = bytes.Length;

        await context.Response.OutputStream.WriteAsync(bytes).ConfigureAwait(false);
        context.Response.Close();
    }

    /// <summary>
    /// Asks the operating system for a free port by binding to port 0. This is inherently a
    /// small race, but it is the standard approach for a native-client loopback redirect and
    /// the window is a few milliseconds.
    /// </summary>
    private static int GetFreePort()
    {
        using var socket = new System.Net.Sockets.TcpListener(IPAddress.Loopback, 0);
        socket.Start();
        var port = ((IPEndPoint)socket.LocalEndpoint).Port;
        socket.Stop();
        return port;
    }
}
