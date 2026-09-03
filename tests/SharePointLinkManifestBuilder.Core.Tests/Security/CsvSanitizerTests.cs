using SharePointLinkManifestBuilder.Core.Security;

namespace SharePointLinkManifestBuilder.Core.Tests.Security;

public class CsvSanitizerTests
{
    /// <summary>
    /// A SharePoint file can genuinely be named this. Without neutralization it becomes a live
    /// formula the moment the exported CSV is opened in a spreadsheet.
    /// </summary>
    [Theory]
    [InlineData("=cmd|'/c calc'!A1")]
    [InlineData("+1+1")]
    [InlineData("-2+3")]
    [InlineData("@SUM(A1:A9)")]
    [InlineData("\tleading tab")]
    [InlineData("\rleading carriage return")]
    public void SanitizeField_FormulaTriggers_ArePrefixedWithAnApostrophe(string value)
    {
        var sanitized = CsvSanitizer.SanitizeField(value);

        Assert.StartsWith("'", sanitized.TrimStart('"'), StringComparison.Ordinal);
    }

    [Fact]
    public void SanitizeField_OrdinaryValue_IsUnchanged() =>
        Assert.Equal("Quarterly Report.docx", CsvSanitizer.SanitizeField("Quarterly Report.docx"));

    [Fact]
    public void SanitizeField_ValueWithComma_IsQuoted() =>
        Assert.Equal("\"Smith, John.docx\"", CsvSanitizer.SanitizeField("Smith, John.docx"));

    [Fact]
    public void SanitizeField_EmbeddedQuote_IsDoubled() =>
        Assert.Equal("\"He said \"\"hi\"\".docx\"", CsvSanitizer.SanitizeField("He said \"hi\".docx"));

    [Fact]
    public void SanitizeField_EmbeddedNewline_IsQuoted() =>
        Assert.StartsWith("\"", CsvSanitizer.SanitizeField("line1\nline2"), StringComparison.Ordinal);

    [Fact]
    public void SanitizeField_Null_BecomesEmpty() =>
        Assert.Equal(string.Empty, CsvSanitizer.SanitizeField(null));

    [Fact]
    public void BuildDocument_WritesHeaderAndCrLfDelimitedRows()
    {
        var csv = CsvSanitizer.BuildDocument(["A", "B"], [["1", "2"], ["3", "4"]]);

        Assert.Equal("A,B\r\n1,2\r\n3,4\r\n", csv);
    }

    [Fact]
    public void FormatTimestamp_IsInvariantIso8601Utc() =>
        Assert.Equal(
            "2026-02-01T09:30:00Z",
            CsvSanitizer.FormatTimestamp(new DateTimeOffset(2026, 2, 1, 9, 30, 0, TimeSpan.Zero)));
}

public class MarkdownEscaperTests
{
    [Fact]
    public void Escape_SpecialCharacters_AreBackslashEscaped() =>
        Assert.Equal("\\*not bold\\*", MarkdownEscaper.Escape("*not bold*"));

    [Fact]
    public void Escape_LineBreaks_BecomeSpaces()
    {
        var escaped = MarkdownEscaper.Escape("a\r\nb\nc");

        Assert.DoesNotContain('\n', escaped);
        Assert.DoesNotContain('\r', escaped);
    }

    [Fact]
    public void EscapeTableCell_Pipe_IsEscapedSoTheTableSurvives() =>
        Assert.Equal("a\\|b", MarkdownEscaper.EscapeTableCell("a|b"));

    /// <summary>
    /// SharePoint URLs routinely contain parentheses and encoded spaces. Angle brackets keep
    /// those from terminating the link target early.
    /// </summary>
    [Fact]
    public void Link_WrapsTheUrlInAngleBrackets() =>
        Assert.Equal(
            "[Report \\(final\\)](<https://example.test/a(b)c>)",
            MarkdownEscaper.Link("Report (final)", "https://example.test/a(b)c"));

    [Fact]
    public void Link_WithoutUrl_ReturnsEscapedLabelOnly() =>
        Assert.Equal("Report", MarkdownEscaper.Link("Report", null));

    [Fact]
    public void Escape_Null_ReturnsEmpty() =>
        Assert.Equal(string.Empty, MarkdownEscaper.Escape(null));
}

public class SafePathBuilderTests
{
    [Theory]
    [InlineData("report.docx", true)]
    [InlineData("Quarterly Report 2026.xlsx", true)]
    [InlineData("..", false)]
    [InlineData(".", false)]
    [InlineData("../etc/passwd", false)]
    [InlineData("..\\windows\\system32", false)]
    [InlineData("sub/dir", false)]
    [InlineData("C:file", false)]
    [InlineData("", false)]
    [InlineData("CON", false)]
    [InlineData("con.txt", false)]
    [InlineData("LPT1.log", false)]
    [InlineData("trailing.", false)]
    [InlineData("trailing ", false)]
    public void ValidateFragment_RejectsUnsafeNames(string fragment, bool expectedSafe) =>
        Assert.Equal(expectedSafe, SafePathBuilder.ValidateFragment(fragment).IsSafe);

    [Fact]
    public void ValidateFragment_ControlCharacters_AreRejected() =>
        Assert.False(SafePathBuilder.ValidateFragment("badname.txt").IsSafe);

    [Fact]
    public void TryBuild_SafeFragments_ProduceAContainedPath()
    {
        var root = Path.Combine(Path.GetTempPath(), "splmb-tests");

        var ok = SafePathBuilder.TryBuild(root, ["reports", "q1.txt"], out var full, out var reason);

        Assert.True(ok, reason);
        Assert.NotNull(full);
        Assert.StartsWith(Path.GetFullPath(root), full, StringComparison.Ordinal);
    }

    [Fact]
    public void TryBuild_TraversalFragment_IsRejected()
    {
        var root = Path.Combine(Path.GetTempPath(), "splmb-tests");

        var ok = SafePathBuilder.TryBuild(root, ["..", "escaped.txt"], out var full, out var reason);

        Assert.False(ok);
        Assert.Null(full);
        Assert.NotNull(reason);
    }

    [Fact]
    public void MakeSafeFileName_ReplacesInvalidCharacters()
    {
        var safe = SafePathBuilder.MakeSafeFileName("in/valid:name?.txt");

        Assert.DoesNotContain('/', safe);
        Assert.DoesNotContain('?', safe);
    }

    [Fact]
    public void MakeSafeFileName_ReservedName_IsPrefixed() =>
        Assert.Equal("_NUL.txt", SafePathBuilder.MakeSafeFileName("NUL.txt"));

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("   ")]
    public void MakeSafeFileName_EmptyInput_UsesFallback(string? input) =>
        Assert.Equal("export", SafePathBuilder.MakeSafeFileName(input));
}
