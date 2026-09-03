using SharePointLinkManifestBuilder.Core.Abstractions;
using SharePointLinkManifestBuilder.Core.Manifests;
using SharePointLinkManifestBuilder.Core.Models;

namespace SharePointLinkManifestBuilder.Core.Tests.Manifests;

public class ManifestMergerTests
{
    private static ManifestHeader Header => new()
    {
        ApplicationVersion = "0.1.0",
        JobId = "job-0002",
        TenantDisplayName = "Example Organization",
        TenantId = TestData.TenantId,
        SourceType = "Document Library",
        LinkPermission = "View",
        LinkAudience = "Organization",
    };

    private static ManifestEntry Entry(
        string itemId,
        string name,
        string status = "Created",
        string path = "",
        string driveId = TestData.DriveA) => new()
        {
            FileName = name,
            RelativePath = string.IsNullOrEmpty(path) ? name : path,
            DriveId = driveId,
            ItemId = itemId,
            Status = status,
            SharingUrl = $"https://example.sharepoint.test/:w:/s/A/{itemId}",
            GeneratedUtc = new DateTimeOffset(2026, 2, 1, 9, 30, 0, TimeSpan.Zero),
        };

    private static ManifestDocument Document(params ManifestEntry[] entries) =>
        new() { Header = Header, Entries = entries };

    [Fact]
    public void Merge_NewFile_IsAppended()
    {
        var merged = new ManifestMerger().Merge(
            Document(Entry("item-1", "a.docx")),
            Document(Entry("item-1", "a.docx"), Entry("item-2", "b.docx")),
            MissingEntryPolicy.Preserve);

        Assert.Equal(2, merged.Entries.Count);
        Assert.Contains(merged.Entries, e => e.ItemId == "item-2");
    }

    /// <summary>
    /// The central guarantee of ADR-0007: a renamed file keeps its identity, so it updates in
    /// place rather than orphaning the old entry and adding a duplicate.
    /// </summary>
    [Fact]
    public void Merge_RenamedFile_UpdatesInPlaceRatherThanDuplicating()
    {
        var merged = new ManifestMerger().Merge(
            Document(Entry("item-1", "old-name.docx")),
            Document(Entry("item-1", "new-name.docx")),
            MissingEntryPolicy.Preserve);

        var entry = Assert.Single(merged.Entries);
        Assert.Equal("new-name.docx", entry.FileName);
        Assert.Equal("item-1", entry.ItemId);
    }

    [Fact]
    public void Merge_MovedFile_UpdatesItsPathInPlace()
    {
        var merged = new ManifestMerger().Merge(
            Document(Entry("item-1", "a.docx", path: "Reports/a.docx")),
            Document(Entry("item-1", "a.docx", path: "Archive/a.docx")),
            MissingEntryPolicy.Preserve);

        Assert.Equal("Archive/a.docx", Assert.Single(merged.Entries).RelativePath);
    }

    /// <summary>Same name, different item: two genuinely different files, so two entries.</summary>
    [Fact]
    public void Merge_SameNameDifferentItemId_ProducesTwoEntries()
    {
        var merged = new ManifestMerger().Merge(
            Document(Entry("item-1", "report.docx", path: "Q1/report.docx")),
            Document(Entry("item-2", "report.docx", path: "Q2/report.docx")),
            MissingEntryPolicy.Preserve);

        Assert.Equal(2, merged.Entries.Count);
    }

    /// <summary>Same item id in a different drive is a different file.</summary>
    [Fact]
    public void Merge_SameItemIdInAnotherDrive_IsTreatedAsDistinct()
    {
        var merged = new ManifestMerger().Merge(
            Document(Entry("item-1", "a.docx", driveId: TestData.DriveA)),
            Document(Entry("item-1", "a.docx", driveId: TestData.DriveB)),
            MissingEntryPolicy.Preserve);

        Assert.Equal(2, merged.Entries.Count);
    }

    [Fact]
    public void Merge_MissingEntryPreservePolicy_KeepsItUnchanged()
    {
        var merged = new ManifestMerger().Merge(
            Document(Entry("item-1", "gone.docx")),
            Document(Entry("item-2", "new.docx")),
            MissingEntryPolicy.Preserve);

        var preserved = Assert.Single(merged.Entries, e => e.ItemId == "item-1");
        Assert.False(preserved.IsMissing);
    }

    [Fact]
    public void Merge_MissingEntryMarkPolicy_FlagsIt()
    {
        var merged = new ManifestMerger().Merge(
            Document(Entry("item-1", "gone.docx")),
            Document(Entry("item-2", "new.docx")),
            MissingEntryPolicy.Mark);

        Assert.True(Assert.Single(merged.Entries, e => e.ItemId == "item-1").IsMissing);
    }

    [Fact]
    public void Merge_MissingEntryRemovePolicy_DropsIt()
    {
        var merged = new ManifestMerger().Merge(
            Document(Entry("item-1", "gone.docx")),
            Document(Entry("item-2", "new.docx")),
            MissingEntryPolicy.Remove);

        Assert.Equal("item-2", Assert.Single(merged.Entries).ItemId);
    }

    /// <summary>A file marked missing that reappears must lose the flag.</summary>
    [Fact]
    public void Merge_PreviouslyMissingFileSeenAgain_IsUnmarked()
    {
        var existing = Document(Entry("item-1", "a.docx") with { IsMissing = true });

        var merged = new ManifestMerger().Merge(
            existing,
            Document(Entry("item-1", "a.docx")),
            MissingEntryPolicy.Mark);

        Assert.False(Assert.Single(merged.Entries).IsMissing);
    }

    /// <summary>Stable ordering keeps version-history diffs readable across runs.</summary>
    [Fact]
    public void Merge_PreservesExistingOrderThenAppendsNewEntries()
    {
        var merged = new ManifestMerger().Merge(
            Document(Entry("item-3", "c.docx"), Entry("item-1", "a.docx")),
            Document(Entry("item-1", "a.docx"), Entry("item-2", "b.docx"), Entry("item-3", "c.docx")),
            MissingEntryPolicy.Preserve);

        Assert.Equal(["item-3", "item-1", "item-2"], merged.Entries.Select(e => e.ItemId));
    }

    [Fact]
    public void Merge_RecomputesCountsForTheMergedDocument()
    {
        var merged = new ManifestMerger().Merge(
            Document(Entry("item-1", "a.docx", status: "Reused")),
            Document(Entry("item-2", "b.docx", status: "Created")),
            MissingEntryPolicy.Preserve);

        Assert.Equal(2, merged.Header.SuccessfulFiles);
        Assert.Equal(1, merged.Header.ReusedLinks);
    }
}

public class ManifestConflictResolverTests
{
    private sealed class FixedClock(DateTimeOffset now) : TimeProvider
    {
        public override DateTimeOffset GetUtcNow() => now;
    }

    private static ManifestConflictResolver Resolver() =>
        new(new PlainTextManifestParser(),
            new FixedClock(new DateTimeOffset(2026, 9, 2, 14, 15, 30, TimeSpan.Zero)));

    private static string ValidManifest => new PlainTextManifestFormatter().Render(new ManifestDocument
    {
        Header = new ManifestHeader
        {
            ApplicationVersion = "0.1.0",
            JobId = "job-0001",
            TenantDisplayName = "Example Organization",
            TenantId = TestData.TenantId,
            SourceType = "Document Library",
            LinkPermission = "View",
            LinkAudience = "Organization",
        },
        Entries = [],
    });

    private static ExistingManifest Existing(string content) => new()
    {
        ItemId = "manifest-item",
        Content = content,
        ETag = "\"etag-1\"",
    };

    [Fact]
    public void Resolve_NoExistingFile_CreatesNew()
    {
        var decision = Resolver().Resolve("_sharepoint-links.txt", null, ManifestConflictPolicy.UpdateSafely);

        Assert.Equal(ManifestWriteAction.CreateNew, decision.Action);
        Assert.Equal("_sharepoint-links.txt", decision.FileName);
        Assert.Null(decision.IfMatchETag);
    }

    [Fact]
    public void Resolve_OurOwnManifest_IsMergedAndWrittenConditionally()
    {
        var decision = Resolver().Resolve(
            "_sharepoint-links.txt", Existing(ValidManifest), ManifestConflictPolicy.UpdateSafely);

        Assert.Equal(ManifestWriteAction.MergeAndReplace, decision.Action);
        Assert.Equal("\"etag-1\"", decision.IfMatchETag);
        Assert.NotNull(decision.ExistingDocument);
    }

    /// <summary>
    /// The core safety property: a file this application did not write is never overwritten by
    /// the default policy. Someone's own document at that path survives.
    /// </summary>
    [Fact]
    public void Resolve_ForeignFile_IsPreservedAndACopyIsTimestamped()
    {
        var decision = Resolver().Resolve(
            "_sharepoint-links.txt",
            Existing("This is somebody's own notes file."),
            ManifestConflictPolicy.UpdateSafely);

        Assert.Equal(ManifestWriteAction.WriteTimestampedCopy, decision.Action);
        Assert.Equal("_sharepoint-links-20260902-141530.txt", decision.FileName);
        Assert.Null(decision.IfMatchETag);
        Assert.Contains("left untouched", decision.Explanation, StringComparison.Ordinal);
    }

    [Fact]
    public void Resolve_ReplacePolicy_OverwritesButStillSendsTheETag()
    {
        var decision = Resolver().Resolve(
            "_sharepoint-links.txt", Existing(ValidManifest), ManifestConflictPolicy.Replace);

        Assert.Equal(ManifestWriteAction.OverwriteWithoutMerge, decision.Action);
        Assert.Equal("\"etag-1\"", decision.IfMatchETag);
    }

    [Fact]
    public void Resolve_SkipPolicy_WritesNothing() =>
        Assert.Equal(
            ManifestWriteAction.Skip,
            Resolver().Resolve("_sharepoint-links.txt", Existing(ValidManifest), ManifestConflictPolicy.Skip).Action);

    [Fact]
    public void Resolve_FailPolicy_ReportsAFailure() =>
        Assert.Equal(
            ManifestWriteAction.Fail,
            Resolver().Resolve("_sharepoint-links.txt", Existing(ValidManifest), ManifestConflictPolicy.Fail).Action);

    [Fact]
    public void Resolve_TimestampedPolicy_LeavesTheOriginalAlone()
    {
        var decision = Resolver().Resolve(
            "_sharepoint-links.txt", Existing(ValidManifest), ManifestConflictPolicy.CreateTimestampedVersion);

        Assert.Equal(ManifestWriteAction.WriteTimestampedCopy, decision.Action);
        Assert.Equal("_sharepoint-links-20260902-141530.txt", decision.FileName);
    }

    [Fact]
    public void BuildTimestampedName_PreservesTheExtension() =>
        Assert.Equal(
            "_sharepoint-links-master-20260902-141530.csv",
            Resolver().BuildTimestampedName("_sharepoint-links-master.csv"));
}

public class ManifestDefaultsTests
{
    [Theory]
    [InlineData("_sharepoint-links.txt", true)]
    [InlineData("_sharepoint-links.md", true)]
    [InlineData("_sharepoint-links.csv", true)]
    [InlineData("_sharepoint-links.json", true)]
    [InlineData("_sharepoint-links-master.txt", true)]
    [InlineData("_sharepoint-links-20260902-141530.txt", true)]
    [InlineData("report.docx", false)]
    [InlineData("_sharepoint-links.docx", false)]
    [InlineData("my-links.txt", false)]
    public void IsGeneratedManifestName_RecognisesOurOwnOutput(string name, bool expected) =>
        Assert.Equal(expected, ManifestDefaults.IsGeneratedManifestName(name));

    [Fact]
    public void GeneratedFileNames_CoverEverySelectedFormat()
    {
        var names = new ManifestConfiguration
        {
            Formats = ManifestFormats.PlainText | ManifestFormats.Json,
        }.GeneratedFileNames();

        Assert.Contains("_sharepoint-links.txt", names);
        Assert.Contains("_sharepoint-links.json", names);
        Assert.Contains("_sharepoint-links-master.txt", names);
        Assert.DoesNotContain("_sharepoint-links.csv", names);
    }

    [Fact]
    public void Validate_NoManifestSelected_IsReported()
    {
        var problems = new ManifestConfiguration
        {
            WritePerFolderManifest = false,
            WriteMasterManifest = false,
        }.Validate();

        Assert.Contains(problems, p => p.Contains("No manifest will be written", StringComparison.Ordinal));
    }

    /// <summary>A combined master manifest must never guess a destination.</summary>
    [Fact]
    public void Validate_CombinedMasterWithoutDestination_IsReported()
    {
        var problems = new ManifestConfiguration
        {
            WriteMasterManifest = true,
            MasterManifestPerTarget = false,
        }.Validate();

        Assert.Contains(problems, p => p.Contains("writable destination", StringComparison.Ordinal));
    }

    [Fact]
    public void Validate_InvalidFileName_IsReported()
    {
        var problems = new ManifestConfiguration { PerFolderFileName = "bad:name" }.Validate();

        Assert.Contains(problems, p => p.Contains("not a valid manifest file name", StringComparison.Ordinal));
    }
}
