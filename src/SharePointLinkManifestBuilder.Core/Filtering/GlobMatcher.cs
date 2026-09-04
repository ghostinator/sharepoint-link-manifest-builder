namespace SharePointLinkManifestBuilder.Core.Filtering;

/// <summary>
/// Matches file names against simple glob patterns supporting <c>*</c> (any run of characters)
/// and <c>?</c> (exactly one character).
/// <para>
/// Implemented as an iterative two-pointer scan with backtracking rather than by translating to
/// a regular expression. A user-supplied pattern becoming a regular expression is a
/// catastrophic-backtracking risk; this runs in bounded time on any input.
/// </para>
/// </summary>
public static class GlobMatcher
{
    /// <summary>Matches a name against a pattern, case-insensitively.</summary>
    /// <param name="pattern">A glob pattern. Null or empty matches nothing.</param>
    /// <param name="name">The name to test.</param>
    public static bool IsMatch(string? pattern, string? name)
    {
        if (string.IsNullOrEmpty(pattern))
        {
            return false;
        }

        name ??= string.Empty;

        var patternIndex = 0;
        var nameIndex = 0;
        var starIndex = -1;
        var nameIndexAtStar = 0;

        while (nameIndex < name.Length)
        {
            if (patternIndex < pattern.Length
                && (pattern[patternIndex] == '?' || EqualsIgnoreCase(pattern[patternIndex], name[nameIndex])))
            {
                patternIndex++;
                nameIndex++;
            }
            else if (patternIndex < pattern.Length && pattern[patternIndex] == '*')
            {
                // Record the star position so we can backtrack and let it consume one more char.
                starIndex = patternIndex;
                nameIndexAtStar = nameIndex;
                patternIndex++;
            }
            else if (starIndex >= 0)
            {
                patternIndex = starIndex + 1;
                nameIndexAtStar++;
                nameIndex = nameIndexAtStar;
            }
            else
            {
                return false;
            }
        }

        while (patternIndex < pattern.Length && pattern[patternIndex] == '*')
        {
            patternIndex++;
        }

        return patternIndex == pattern.Length;
    }

    /// <summary>True when the name matches at least one pattern.</summary>
    public static bool IsMatchAny(IEnumerable<string> patterns, string? name) =>
        patterns.Any(pattern => IsMatch(pattern, name));

    private static bool EqualsIgnoreCase(char a, char b) =>
        a == b || char.ToUpperInvariant(a) == char.ToUpperInvariant(b);
}
