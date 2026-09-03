using System.Text.RegularExpressions;

namespace SharePointLinkManifestBuilder.Core.Security;

/// <summary>
/// Removes credential material from any text before it is logged, exported, or shown.
/// <para>
/// Applied by the logging decorator to every message and scope value at every level,
/// including Trace, so there is no log path that bypasses it.
/// </para>
/// </summary>
public static partial class SensitiveDataRedactor
{
    /// <summary>The text substituted for redacted material.</summary>
    public const string Placeholder = "[REDACTED]";

    /// <summary>Query and form parameters whose values are credentials.</summary>
    private static readonly string[] SensitiveParameterNames =
    [
        "code", "access_token", "refresh_token", "id_token", "client_secret",
        "assertion", "client_assertion", "code_verifier", "password", "pwd",
        "session_state", "sig", "signature", "key",
    ];

    [GeneratedRegex(
        @"\bBearer\s+[A-Za-z0-9\-._~+/]+=*",
        RegexOptions.IgnoreCase | RegexOptions.CultureInvariant)]
    private static partial Regex BearerTokenPattern();

    [GeneratedRegex(
        @"\beyJ[A-Za-z0-9_-]{5,}\.[A-Za-z0-9_-]{5,}\.[A-Za-z0-9_-]*",
        RegexOptions.CultureInvariant)]
    private static partial Regex JwtPattern();

    [GeneratedRegex(
        @"(?i)\b(authorization|x-ms-client-request-id-secret|cookie|set-cookie)\s*[:=]\s*\S+",
        RegexOptions.CultureInvariant)]
    private static partial Regex SensitiveHeaderPattern();

    /// <summary>
    /// Redacts credential material from arbitrary text. Safe to call on any string, including
    /// null or empty.
    /// </summary>
    public static string Redact(string? text)
    {
        if (string.IsNullOrEmpty(text))
        {
            return text ?? string.Empty;
        }

        var result = BearerTokenPattern().Replace(text, "Bearer " + Placeholder);
        result = JwtPattern().Replace(result, Placeholder);
        result = SensitiveHeaderPattern().Replace(result, match =>
        {
            var separator = match.Value.IndexOfAny([':', '=']);
            return separator < 0 ? Placeholder : match.Value[..(separator + 1)] + " " + Placeholder;
        });

        return RedactQueryParameters(result);
    }

    /// <summary>
    /// Replaces the values of sensitive query parameters while leaving the rest of a URL
    /// readable, so logs stay useful for diagnosis.
    /// </summary>
    public static string RedactQueryParameters(string text)
    {
        if (string.IsNullOrEmpty(text) || !text.Contains('=', StringComparison.Ordinal))
        {
            return text;
        }

        foreach (var name in SensitiveParameterNames)
        {
            text = Regex.Replace(
                text,
                $@"(?i)([?&;]|\b){Regex.Escape(name)}=([^&\s""']*)",
                $"$1{name}={Placeholder}",
                RegexOptions.CultureInvariant,
                TimeSpan.FromSeconds(1));
        }

        return text;
    }

    /// <summary>
    /// Reduces a URL to scheme, host and path, dropping the query entirely. Used when logging
    /// a request so no parameter can leak, sensitive or otherwise.
    /// </summary>
    public static string RedactUrl(string? url)
    {
        if (string.IsNullOrWhiteSpace(url))
        {
            return string.Empty;
        }

        if (!Uri.TryCreate(url, UriKind.Absolute, out var uri))
        {
            return RedactQueryParameters(url);
        }

        return string.IsNullOrEmpty(uri.Query)
            ? $"{uri.Scheme}://{uri.Host}{uri.AbsolutePath}"
            : $"{uri.Scheme}://{uri.Host}{uri.AbsolutePath}?{Placeholder}";
    }

    /// <summary>
    /// Masks a GUID-shaped identifier for display, keeping the first and last four characters
    /// so a user can still recognise it. Client IDs are not secret, but they are identifying,
    /// so sanitized exports mask them.
    /// </summary>
    public static string MaskIdentifier(string? value)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return string.Empty;
        }

        return value.Length <= 8
            ? new string('*', value.Length)
            : $"{value[..4]}{new string('*', Math.Min(value.Length - 8, 24))}{value[^4..]}";
    }

    /// <summary>
    /// Masks the local part of an email address, keeping the domain, so a diagnostic bundle can
    /// show that an address was present without disclosing who it belongs to.
    /// </summary>
    public static string MaskEmail(string? value)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return string.Empty;
        }

        var at = value.IndexOf('@', StringComparison.Ordinal);
        return at <= 0 ? Placeholder : $"{value[0]}***@{value[(at + 1)..]}";
    }
}
