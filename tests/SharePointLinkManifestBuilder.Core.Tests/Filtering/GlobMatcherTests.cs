using SharePointLinkManifestBuilder.Core.Filtering;

namespace SharePointLinkManifestBuilder.Core.Tests.Filtering;

public class GlobMatcherTests
{
    [Theory]
    [InlineData("*.docx", "report.docx", true)]
    [InlineData("*.docx", "report.DOCX", true)]
    [InlineData("*.docx", "report.pdf", false)]
    [InlineData("report*", "report-2026.docx", true)]
    [InlineData("*report*", "quarterly-report-final.docx", true)]
    [InlineData("report?.docx", "report1.docx", true)]
    [InlineData("report?.docx", "report12.docx", false)]
    [InlineData("*", "anything", true)]
    [InlineData("**", "anything", true)]
    [InlineData("exact.txt", "exact.txt", true)]
    [InlineData("exact.txt", "exact.txts", false)]
    [InlineData("a*b*c", "aXXbYYc", true)]
    [InlineData("a*b*c", "aXXbYY", false)]
    public void IsMatch_HandlesWildcards(string pattern, string name, bool expected) =>
        Assert.Equal(expected, GlobMatcher.IsMatch(pattern, name));

    [Fact]
    public void IsMatch_EmptyPattern_MatchesNothing() =>
        Assert.False(GlobMatcher.IsMatch(string.Empty, "anything"));

    [Fact]
    public void IsMatch_StarMatchesEmptyString() =>
        Assert.True(GlobMatcher.IsMatch("a*", "a"));

    /// <summary>
    /// A pathological pattern must not hang. This is why matching is an iterative scan rather
    /// than a translated regular expression, which could backtrack catastrophically.
    /// </summary>
    [Fact]
    public void IsMatch_PathologicalPattern_CompletesQuickly()
    {
        var pattern = new string('*', 40) + "z";
        var name = new string('a', 2000);

        var stopwatch = System.Diagnostics.Stopwatch.StartNew();
        var matched = GlobMatcher.IsMatch(pattern, name);
        stopwatch.Stop();

        Assert.False(matched);
        Assert.True(
            stopwatch.ElapsedMilliseconds < 1000,
            $"Matching took {stopwatch.ElapsedMilliseconds} ms, which suggests runaway backtracking.");
    }

    [Fact]
    public void IsMatchAny_ReturnsTrueWhenAnyPatternMatches() =>
        Assert.True(GlobMatcher.IsMatchAny(["*.pdf", "*.docx"], "notes.docx"));

    [Fact]
    public void IsMatchAny_ReturnsFalseWhenNoPatternMatches() =>
        Assert.False(GlobMatcher.IsMatchAny(["*.pdf", "*.xlsx"], "notes.docx"));
}
