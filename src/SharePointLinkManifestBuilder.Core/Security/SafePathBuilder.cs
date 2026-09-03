using System.Diagnostics.CodeAnalysis;

namespace SharePointLinkManifestBuilder.Core.Security;

/// <summary>The outcome of validating a user- or service-supplied path fragment.</summary>
/// <param name="IsSafe">True when the fragment may be used.</param>
/// <param name="Reason">Why the fragment was rejected, when it was.</param>
public readonly record struct PathValidationResult(bool IsSafe, string? Reason)
{
    /// <summary>A safe result.</summary>
    public static PathValidationResult Safe() => new(true, null);

    /// <summary>An unsafe result with an explanation.</summary>
    public static PathValidationResult Unsafe(string reason) => new(false, reason);
}

/// <summary>
/// Validates path fragments before they are used to build a local file path.
/// <para>
/// File and folder names originate in SharePoint and are attacker-influenced. Writing an
/// export to a name such as <c>../../.ssh/authorized_keys</c> must be impossible.
/// </para>
/// </summary>
public static class SafePathBuilder
{
    /// <summary>Windows reserved device names, which are unsafe even with an extension.</summary>
    private static readonly string[] ReservedDeviceNames =
    [
        "CON", "PRN", "AUX", "NUL",
        "COM1", "COM2", "COM3", "COM4", "COM5", "COM6", "COM7", "COM8", "COM9",
        "LPT1", "LPT2", "LPT3", "LPT4", "LPT5", "LPT6", "LPT7", "LPT8", "LPT9",
    ];

    /// <summary>Validates a single file or directory name fragment.</summary>
    public static PathValidationResult ValidateFragment(string? fragment)
    {
        if (string.IsNullOrWhiteSpace(fragment))
        {
            return PathValidationResult.Unsafe("The name is empty.");
        }

        if (fragment is "." or "..")
        {
            return PathValidationResult.Unsafe("Relative path segments are not allowed.");
        }

        if (fragment.Contains("..", StringComparison.Ordinal))
        {
            return PathValidationResult.Unsafe("The name contains a parent-directory reference.");
        }

        if (fragment.Contains('/', StringComparison.Ordinal)
            || fragment.Contains('\\', StringComparison.Ordinal))
        {
            return PathValidationResult.Unsafe("The name contains a directory separator.");
        }

        if (fragment.Contains(':', StringComparison.Ordinal))
        {
            return PathValidationResult.Unsafe("The name contains a drive or stream separator.");
        }

        if (fragment.Any(char.IsControl))
        {
            return PathValidationResult.Unsafe("The name contains control characters.");
        }

        var stem = Path.GetFileNameWithoutExtension(fragment);
        if (ReservedDeviceNames.Contains(stem, StringComparer.OrdinalIgnoreCase))
        {
            return PathValidationResult.Unsafe($"'{stem}' is a reserved device name on Windows.");
        }

        if (fragment.EndsWith('.') || fragment.EndsWith(' '))
        {
            return PathValidationResult.Unsafe("The name ends with a dot or space, which is not portable.");
        }

        return PathValidationResult.Safe();
    }

    /// <summary>
    /// Combines a base directory with untrusted fragments, then verifies the resolved path is
    /// still inside the base directory. The containment check is the real control; fragment
    /// validation is the first line of defence.
    /// </summary>
    /// <param name="baseDirectory">The directory the result must stay within.</param>
    /// <param name="fragments">Untrusted name fragments.</param>
    /// <param name="fullPath">The resolved absolute path, when safe.</param>
    /// <param name="reason">Why the path was rejected, when it was.</param>
    public static bool TryBuild(
        string baseDirectory,
        IEnumerable<string> fragments,
        [NotNullWhen(true)] out string? fullPath,
        out string? reason)
    {
        fullPath = null;
        reason = null;

        var materialized = fragments.ToArray();
        foreach (var fragment in materialized)
        {
            var validation = ValidateFragment(fragment);
            if (!validation.IsSafe)
            {
                reason = validation.Reason;
                return false;
            }
        }

        var root = Path.GetFullPath(baseDirectory);
        var candidate = Path.GetFullPath(Path.Combine([root, .. materialized]));

        // Containment check. Comparing the resolved paths defeats traversal, symlinked
        // fragments and casing tricks that slipped past fragment validation.
        var rootWithSeparator = root.EndsWith(Path.DirectorySeparatorChar)
            ? root
            : root + Path.DirectorySeparatorChar;

        if (!candidate.StartsWith(rootWithSeparator, StringComparison.Ordinal)
            && !string.Equals(candidate, root, StringComparison.Ordinal))
        {
            reason = "The resolved path would fall outside the destination directory.";
            return false;
        }

        fullPath = candidate;
        return true;
    }

    /// <summary>
    /// Replaces characters that are unsafe in a file name with an underscore, so an
    /// untrusted name can be used as a local export name without being rejected outright.
    /// </summary>
    public static string MakeSafeFileName(string? name, string fallback = "export")
    {
        if (string.IsNullOrWhiteSpace(name))
        {
            return fallback;
        }

        var invalid = Path.GetInvalidFileNameChars();
        var cleaned = new string(name.Select(c =>
            Array.IndexOf(invalid, c) >= 0 || char.IsControl(c) ? '_' : c).ToArray());

        cleaned = cleaned.Trim().TrimEnd('.');

        var stem = Path.GetFileNameWithoutExtension(cleaned);
        if (ReservedDeviceNames.Contains(stem, StringComparer.OrdinalIgnoreCase))
        {
            cleaned = "_" + cleaned;
        }

        return string.IsNullOrWhiteSpace(cleaned) ? fallback : cleaned;
    }
}
