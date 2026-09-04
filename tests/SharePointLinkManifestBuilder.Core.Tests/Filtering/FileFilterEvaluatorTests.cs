using SharePointLinkManifestBuilder.Core.Filtering;
using SharePointLinkManifestBuilder.Core.Models;

namespace SharePointLinkManifestBuilder.Core.Tests.Filtering;

public class FileFilterEvaluatorTests
{
    private static FileFilterEvaluator Evaluator(FilterConfiguration? filters = null) =>
        new(filters ?? new FilterConfiguration(), new ManifestConfiguration());

    [Fact]
    public void Evaluate_OrdinaryFile_IsEligibleByDefault() =>
        Assert.Equal(SkipReason.None, Evaluator().Evaluate(TestData.File("report.docx")));

    [Fact]
    public void Evaluate_Folder_IsSkipped() =>
        Assert.Equal(
            SkipReason.IsFolder,
            Evaluator().Evaluate(TestData.File("Reports", kind: DriveItemKind.Folder)));

    [Theory]
    [InlineData("~$report.docx")]
    [InlineData(".~lock.report.docx")]
    [InlineData("scratch.tmp")]
    [InlineData("Thumbs.db")]
    [InlineData(".DS_Store")]
    public void Evaluate_TemporaryFiles_AreSkippedByDefault(string name) =>
        Assert.Equal(SkipReason.TemporaryFile, Evaluator().Evaluate(TestData.File(name)));

    [Fact]
    public void Evaluate_TemporaryFile_IsIncludedWhenExplicitlyRequested()
    {
        var evaluator = Evaluator(new FilterConfiguration { IncludeTemporaryFiles = true });

        Assert.Equal(SkipReason.None, evaluator.Evaluate(TestData.File("~$report.docx")));
    }

    /// <summary>
    /// A job must never discover the manifests it is about to write, or a second run would
    /// create links to its own output.
    /// </summary>
    [Theory]
    [InlineData("_sharepoint-links.txt")]
    [InlineData("_sharepoint-links-master.txt")]
    [InlineData("_sharepoint-links.csv")]
    [InlineData("_sharepoint-links-20260902-141530.txt")]
    public void Evaluate_GeneratedManifests_AreSkippedByDefault(string name) =>
        Assert.Equal(SkipReason.GeneratedManifest, Evaluator().Evaluate(TestData.File(name)));

    [Fact]
    public void Evaluate_FileNamedLikeAManifestButWithAnotherExtension_IsNotSkipped() =>
        Assert.Equal(SkipReason.None, Evaluator().Evaluate(TestData.File("_sharepoint-links.docx")));

    [Fact]
    public void Evaluate_HiddenItem_IsSkippedByDefault() =>
        Assert.Equal(
            SkipReason.HiddenOrSystem,
            Evaluator().Evaluate(TestData.File("hidden.docx"), isHiddenOrSystem: true));

    [Fact]
    public void Evaluate_Package_IsSkippedByDefault() =>
        Assert.Equal(
            SkipReason.PackageItem,
            Evaluator().Evaluate(TestData.File("Notebook", kind: DriveItemKind.Package)));

    [Fact]
    public void Evaluate_RemoteItem_IsSkippedByDefault() =>
        Assert.Equal(
            SkipReason.RemoteItem,
            Evaluator().Evaluate(TestData.File("Shortcut", kind: DriveItemKind.RemoteItem)));

    [Fact]
    public void Evaluate_UnsupportedKind_IsAlwaysSkipped()
    {
        var evaluator = Evaluator(new FilterConfiguration { IncludeSpecialItemTypes = true });

        Assert.Equal(
            SkipReason.UnsupportedItemType,
            evaluator.Evaluate(TestData.File("mystery", kind: DriveItemKind.Unsupported)));
    }

    [Theory]
    [InlineData("docx", "report.docx", SkipReason.None)]
    [InlineData(".docx", "report.docx", SkipReason.None)]
    [InlineData("pdf", "report.docx", SkipReason.FilteredOut)]
    public void Evaluate_IncludeExtensions_ToleratesLeadingDotOrNot(
        string extension, string fileName, SkipReason expected)
    {
        var evaluator = Evaluator(new FilterConfiguration { IncludeExtensions = [extension] });

        Assert.Equal(expected, evaluator.Evaluate(TestData.File(fileName)));
    }

    [Fact]
    public void Evaluate_ExcludeExtensions_RemovesMatchingFiles()
    {
        var evaluator = Evaluator(new FilterConfiguration { ExcludeExtensions = ["exe", "zip"] });

        Assert.Equal(SkipReason.FilteredOut, evaluator.Evaluate(TestData.File("setup.exe")));
        Assert.Equal(SkipReason.None, evaluator.Evaluate(TestData.File("notes.txt")));
    }

    [Fact]
    public void Evaluate_IncludeAndExcludePatterns_AreBothApplied()
    {
        var evaluator = Evaluator(new FilterConfiguration
        {
            IncludePatterns = ["report*"],
            ExcludePatterns = ["*draft*"],
        });

        Assert.Equal(SkipReason.None, evaluator.Evaluate(TestData.File("report-final.docx")));
        Assert.Equal(SkipReason.FilteredOut, evaluator.Evaluate(TestData.File("report-draft.docx")));
        Assert.Equal(SkipReason.FilteredOut, evaluator.Evaluate(TestData.File("summary.docx")));
    }

    [Fact]
    public void Evaluate_DateRange_IsInclusive()
    {
        var evaluator = Evaluator(new FilterConfiguration
        {
            ModifiedAfterUtc = new DateTimeOffset(2026, 1, 1, 0, 0, 0, TimeSpan.Zero),
            ModifiedBeforeUtc = new DateTimeOffset(2026, 12, 31, 23, 59, 59, TimeSpan.Zero),
        });

        Assert.Equal(SkipReason.None, evaluator.Evaluate(
            TestData.File("in.docx", modified: new DateTimeOffset(2026, 6, 1, 0, 0, 0, TimeSpan.Zero))));

        Assert.Equal(SkipReason.FilteredOut, evaluator.Evaluate(
            TestData.File("old.docx", modified: new DateTimeOffset(2025, 6, 1, 0, 0, 0, TimeSpan.Zero))));
    }

    [Fact]
    public void Evaluate_SizeRange_IsInclusive()
    {
        var evaluator = Evaluator(new FilterConfiguration
        {
            MinimumSizeBytes = 100,
            MaximumSizeBytes = 1000,
        });

        Assert.Equal(SkipReason.None, evaluator.Evaluate(TestData.File("mid.docx", size: 500)));
        Assert.Equal(SkipReason.None, evaluator.Evaluate(TestData.File("min.docx", size: 100)));
        Assert.Equal(SkipReason.None, evaluator.Evaluate(TestData.File("max.docx", size: 1000)));
        Assert.Equal(SkipReason.FilteredOut, evaluator.Evaluate(TestData.File("tiny.docx", size: 99)));
        Assert.Equal(SkipReason.FilteredOut, evaluator.Evaluate(TestData.File("huge.docx", size: 1001)));
    }

    [Fact]
    public void Evaluate_SizeFilterWithUnknownSize_ExcludesTheFile()
    {
        var evaluator = Evaluator(new FilterConfiguration { MinimumSizeBytes = 1 });

        Assert.Equal(SkipReason.FilteredOut, evaluator.Evaluate(TestData.File("unknown.docx", size: null)));
    }

    [Fact]
    public void Validate_ContradictoryDateRange_IsReported()
    {
        var problems = new FilterConfiguration
        {
            ModifiedAfterUtc = new DateTimeOffset(2026, 6, 1, 0, 0, 0, TimeSpan.Zero),
            ModifiedBeforeUtc = new DateTimeOffset(2026, 1, 1, 0, 0, 0, TimeSpan.Zero),
        }.Validate();

        Assert.Contains(problems, p => p.Contains("nothing can match", StringComparison.Ordinal));
    }

    [Fact]
    public void Validate_ContradictorySizeRange_IsReported()
    {
        var problems = new FilterConfiguration
        {
            MinimumSizeBytes = 1000,
            MaximumSizeBytes = 10,
        }.Validate();

        Assert.Contains(problems, p => p.Contains("nothing can match", StringComparison.Ordinal));
    }

    [Fact]
    public void Describe_EmptyConfiguration_MentionsDefaultExclusions() =>
        Assert.Contains("default exclusions", new FilterConfiguration().Describe(), StringComparison.Ordinal);
}
