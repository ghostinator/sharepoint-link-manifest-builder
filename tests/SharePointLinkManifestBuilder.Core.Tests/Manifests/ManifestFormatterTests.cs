using System.Text.Json;
using SharePointLinkManifestBuilder.Core.Manifests;
using SharePointLinkManifestBuilder.Core.Models;

namespace SharePointLinkManifestBuilder.Core.Tests.Manifests;

public class PlainTextManifestFormatterTests
{
    private static ManifestDocument Document(int entryCount = 2) => new()
    {
        Header = new ManifestHeader
        {
            ApplicationVersion = "0.1.0",
            JobId = "job-0001",
            GeneratedUtc = new DateTimeOffset(2026, 2, 1, 9, 30, 0, TimeSpan.Zero),
            TenantDisplayName = "Example Organization",
            TenantId = TestData.TenantId,
            SourceType = "Document Library",
            SiteName = "Example Site",
            SiteUrl = "https://example.sharepoint.test/sites/A",
            LibraryOrDriveName = "Documents",
            StartingFolder = "/Reports",
            Recursive = true,
            LinkPermission = "View",
            LinkAudience = "Organization",
            SuccessfulFiles = entryCount,
            ReusedLinks = 1,
            SkippedFiles = 3,
            FailedFiles = 0,
        },
        Entries = Enumerable.Range(1, entryCount).Select(i => new ManifestEntry
        {
            FileName = $"report-{i}.docx",
            RelativePath = $"Reports/report-{i}.docx",
            WebUrl = $"https://example.sharepoint.test/sites/A/Shared%20Documents/report-{i}.docx",
            SharingUrl = $"https://example.sharepoint.test/:w:/s/A/EXAMPLE{i}",
            DriveId = TestData.DriveA,
            ItemId = $"item-{i}",
            Status = i == 1 ? "Created" : "Reused",
            GeneratedUtc = new DateTimeOffset(2026, 2, 1, 9, 30, 0, TimeSpan.Zero),
        }).ToArray(),
    };

    [Fact]
    public void Render_BeginsWithTheDocumentHeader()
    {
        var text = new PlainTextManifestFormatter().Render(Document());

        Assert.StartsWith("SharePoint Link Manifest\n", text, StringComparison.Ordinal);
    }

    [Fact]
    public void Render_EmitsEveryRequiredHeaderField()
    {
        var text = new PlainTextManifestFormatter().Render(Document());

        foreach (var field in new[]
        {
            "Schema Version:", "Application Version:", "Job ID:", "Generated:", "Tenant:",
            "Tenant ID:", "Source Type:", "Site:", "Site URL:", "Document Library or Drive:",
            "Starting Folder:", "Recursive:", "Link Permission:", "Link Audience:",
            "Successful Files:", "Reused Links:", "Skipped Files:", "Failed Files:",
        })
        {
            Assert.Contains(field, text, StringComparison.Ordinal);
        }
    }

    [Fact]
    public void Render_EmitsEveryRequiredEntryField()
    {
        var text = new PlainTextManifestFormatter().Render(Document(1));

        foreach (var field in new[]
        {
            "File:", "Relative Path:", "Web URL:", "Sharing Link:",
            "Drive ID:", "Item ID:", "Status:", "Generated:",
        })
        {
            Assert.Contains(field, text, StringComparison.Ordinal);
        }

        Assert.Contains("\n---\n", text, StringComparison.Ordinal);
    }

    /// <summary>Fields marked "when applicable" are omitted, not written blank.</summary>
    [Fact]
    public void Render_OmitsInapplicableOptionalFields()
    {
        var document = Document(1);
        var oneDrive = document with
        {
            Header = document.Header with { SiteName = null, SiteUrl = null, SourceType = "My OneDrive" },
        };

        var text = new PlainTextManifestFormatter().Render(oneDrive);

        Assert.DoesNotContain("Site:", text, StringComparison.Ordinal);
        Assert.DoesNotContain("Site URL:", text, StringComparison.Ordinal);
        Assert.Contains("Source Type: My OneDrive", text, StringComparison.Ordinal);
    }

    [Fact]
    public void Render_RecursiveFlag_IsYesOrNo()
    {
        var document = Document(1);

        Assert.Contains("Recursive: Yes", new PlainTextManifestFormatter().Render(document), StringComparison.Ordinal);

        var nonRecursive = document with { Header = document.Header with { Recursive = false } };
        Assert.Contains("Recursive: No", new PlainTextManifestFormatter().Render(nonRecursive), StringComparison.Ordinal);
    }

    [Fact]
    public void Render_TimestampsAreIso8601Utc()
    {
        var text = new PlainTextManifestFormatter().Render(Document(1));

        Assert.Contains("Generated: 2026-02-01T09:30:00Z", text, StringComparison.Ordinal);
    }

    /// <summary>
    /// A newline smuggled into a value would otherwise create a bogus field and corrupt every
    /// entry that follows it.
    /// </summary>
    [Fact]
    public void Render_CollapsesLineBreaksInsideValues()
    {
        var document = Document(1);
        var hostile = document with
        {
            Entries = [document.Entries[0] with { FileName = "evil\nStatus: Created\nFile: fake.docx" }],
        };

        var text = new PlainTextManifestFormatter().Render(hostile);
        var statusLines = text.Split('\n').Count(l => l.StartsWith("Status:", StringComparison.Ordinal));

        Assert.Equal(1, statusLines);
    }

    [Fact]
    public void Render_ProducesNoByteOrderMark()
    {
        var bytes = System.Text.Encoding.UTF8.GetBytes(new PlainTextManifestFormatter().Render(Document(1)));

        Assert.False(bytes.Length >= 3 && bytes[0] == 0xEF && bytes[1] == 0xBB && bytes[2] == 0xBF);
    }
}

public class PlainTextManifestParserTests
{
    private static string Sample => new PlainTextManifestFormatter().Render(new ManifestDocument
    {
        Header = new ManifestHeader
        {
            ApplicationVersion = "0.1.0",
            JobId = "job-0001",
            GeneratedUtc = new DateTimeOffset(2026, 2, 1, 9, 30, 0, TimeSpan.Zero),
            TenantDisplayName = "Example Organization",
            TenantId = TestData.TenantId,
            SourceType = "Document Library",
            SiteName = "Example Site",
            LibraryOrDriveName = "Documents",
            StartingFolder = "/Reports",
            Recursive = true,
            LinkPermission = "View",
            LinkAudience = "Organization",
            SuccessfulFiles = 2,
            ReusedLinks = 1,
            SkippedFiles = 3,
            FailedFiles = 0,
        },
        Entries =
        [
            new ManifestEntry
            {
                FileName = "a.docx", RelativePath = "Reports/a.docx",
                WebUrl = "https://example.sharepoint.test/a", SharingUrl = "https://example.sharepoint.test/:w:/s/A/1",
                DriveId = TestData.DriveA, ItemId = "item-1", Status = "Created",
                GeneratedUtc = new DateTimeOffset(2026, 2, 1, 9, 30, 0, TimeSpan.Zero),
            },
            new ManifestEntry
            {
                FileName = "b.docx", RelativePath = "Reports/b.docx",
                WebUrl = "https://example.sharepoint.test/b", SharingUrl = "https://example.sharepoint.test/:w:/s/A/2",
                DriveId = TestData.DriveA, ItemId = "item-2", Status = "Reused",
                GeneratedUtc = new DateTimeOffset(2026, 2, 1, 9, 30, 0, TimeSpan.Zero),
            },
        ],
    });

    /// <summary>Update mode depends on this: what the formatter writes, the parser must read back.</summary>
    [Fact]
    public void TryParse_RoundTripsAFormattedManifest()
    {
        var result = new PlainTextManifestParser().TryParse(Sample);

        Assert.True(result.Succeeded, result.FailureReason);
        var document = result.Document!;

        Assert.Equal("0.1.0", document.Header.ApplicationVersion);
        Assert.Equal("job-0001", document.Header.JobId);
        Assert.Equal("Example Organization", document.Header.TenantDisplayName);
        Assert.Equal(TestData.TenantId, document.Header.TenantId);
        Assert.Equal("Document Library", document.Header.SourceType);
        Assert.Equal("Example Site", document.Header.SiteName);
        Assert.Equal("/Reports", document.Header.StartingFolder);
        Assert.True(document.Header.Recursive);
        Assert.Equal("View", document.Header.LinkPermission);
        Assert.Equal(3, document.Header.SkippedFiles);

        Assert.Equal(2, document.Entries.Count);
        Assert.Equal("a.docx", document.Entries[0].FileName);
        Assert.Equal("item-1", document.Entries[0].ItemId);
        Assert.Equal(TestData.DriveA, document.Entries[0].DriveId);
        Assert.Equal("Created", document.Entries[0].Status);
        Assert.Equal("Reused", document.Entries[1].Status);
    }

    /// <summary>
    /// A parse failure is not an error: it means the destination holds a file this application
    /// did not write, which the conflict policy uses to avoid destroying it.
    /// </summary>
    [Theory]
    [InlineData("Some unrelated document a user happened to name the same thing.")]
    [InlineData("")]
    [InlineData("   ")]
    public void TryParse_ForeignContent_FailsGracefully(string content)
    {
        var result = new PlainTextManifestParser().TryParse(content);

        Assert.False(result.Succeeded);
        Assert.NotNull(result.FailureReason);
        Assert.Null(result.Document);
    }

    [Fact]
    public void TryParse_MissingSchemaVersion_Fails()
    {
        var result = new PlainTextManifestParser().TryParse("SharePoint Link Manifest\nTenant: Example\n\n");

        Assert.False(result.Succeeded);
        Assert.Contains("schema version", result.FailureReason, StringComparison.OrdinalIgnoreCase);
    }

    /// <summary>
    /// A manifest from a future major version must not be silently mis-parsed and then
    /// overwritten with a downgraded document.
    /// </summary>
    [Fact]
    public void TryParse_NewerMajorSchemaVersion_IsRefused()
    {
        var future = Sample.Replace("Schema Version: 1.0", "Schema Version: 99.0", StringComparison.Ordinal);

        var result = new PlainTextManifestParser().TryParse(future);

        Assert.False(result.Succeeded);
        Assert.Contains("newer than this application supports", result.FailureReason, StringComparison.Ordinal);
    }

    /// <summary>Identity is (driveId, itemId); an entry lacking either cannot be matched later.</summary>
    [Fact]
    public void TryParse_EntryMissingIdentity_IsDropped()
    {
        var text = Sample.Replace("Item ID: item-2\n", string.Empty, StringComparison.Ordinal);

        var result = new PlainTextManifestParser().TryParse(text);

        Assert.True(result.Succeeded);
        Assert.Single(result.Document!.Entries);
    }

    [Fact]
    public void TryParse_UnknownFields_AreIgnored()
    {
        var text = Sample.Replace(
            "Link Permission: View",
            "Some Future Field: value\nLink Permission: View",
            StringComparison.Ordinal);

        var result = new PlainTextManifestParser().TryParse(text);

        Assert.True(result.Succeeded);
        Assert.Equal(2, result.Document!.Entries.Count);
    }

    [Fact]
    public void TryParse_MissingMarkedEntry_IsRecognised()
    {
        var text = Sample.Replace("Status: Created", "Status: Created (not found on last run)", StringComparison.Ordinal);

        var result = new PlainTextManifestParser().TryParse(text);

        Assert.True(result.Succeeded);
        Assert.True(result.Document!.Entries[0].IsMissing);
        Assert.Equal("Created", result.Document.Entries[0].Status);
    }

    [Fact]
    public void TryParse_TrailingEntryWithoutSeparator_IsStillRead()
    {
        var text = Sample.TrimEnd('\n');
        text = text[..text.LastIndexOf("---", StringComparison.Ordinal)];

        var result = new PlainTextManifestParser().TryParse(text);

        Assert.True(result.Succeeded);
        Assert.Equal(2, result.Document!.Entries.Count);
    }

    [Theory]
    [InlineData("1.0", true)]
    [InlineData("0.9", true)]
    [InlineData("2.0", false)]
    [InlineData("not-a-version", false)]
    [InlineData(null, false)]
    public void IsSupportedSchemaVersion_AcceptsOnlyKnownMajorVersions(string? version, bool expected) =>
        Assert.Equal(expected, PlainTextManifestParser.IsSupportedSchemaVersion(version));
}

public class CsvAndJsonAndMarkdownFormatterTests
{
    private static ManifestDocument HostileDocument => new()
    {
        Header = new ManifestHeader
        {
            ApplicationVersion = "0.1.0",
            JobId = "job-0001",
            TenantDisplayName = "Example Organization",
            TenantId = TestData.TenantId,
            SourceType = "Document Library",
            LibraryOrDriveName = "Documents",
            LinkPermission = "View",
            LinkAudience = "Organization",
            SuccessfulFiles = 1,
        },
        Entries =
        [
            new ManifestEntry
            {
                FileName = "=cmd|'/c calc'!A1.docx",
                RelativePath = "Reports/=cmd.docx",
                WebUrl = "https://example.sharepoint.test/a",
                SharingUrl = "https://example.sharepoint.test/:w:/s/A/1",
                DriveId = TestData.DriveA,
                ItemId = "item-1",
                Status = "Created",
                GeneratedUtc = new DateTimeOffset(2026, 2, 1, 9, 30, 0, TimeSpan.Zero),
            },
        ],
    };

    [Fact]
    public void Csv_EmitsTheDocumentedColumnOrder()
    {
        var csv = new CsvManifestFormatter().Render(HostileDocument);
        var header = csv.Split("\r\n")[0];

        Assert.Equal(string.Join(',', CsvManifestFormatter.Columns), header);
    }

    /// <summary>The formula-injection defence must survive the whole formatter, not just the sanitizer.</summary>
    [Fact]
    public void Csv_NeutralizesFormulaInjectionInFileNames()
    {
        var csv = new CsvManifestFormatter().Render(HostileDocument);
        var dataRow = csv.Split("\r\n")[1];

        // The apostrophe prefix is the defence. RFC 4180 quoting is not triggered here because
        // the value contains no comma, quote or newline.
        Assert.StartsWith("'=cmd", dataRow, StringComparison.Ordinal);
        Assert.DoesNotContain("\r\n=cmd", csv, StringComparison.Ordinal);
    }

    [Fact]
    public void Json_IsValidAndCarriesSchemaAndVersion()
    {
        var json = new JsonManifestFormatter().Render(HostileDocument);

        using var document = JsonDocument.Parse(json);
        var root = document.RootElement;

        Assert.Equal(JsonManifestFormatter.SchemaId, root.GetProperty("$schema").GetString());
        Assert.Equal(ManifestDefaults.SchemaVersion, root.GetProperty("schemaVersion").GetString());
        Assert.Equal(TestData.TenantId, root.GetProperty("tenant").GetProperty("tenantId").GetString());
        Assert.Equal(1, root.GetProperty("files").GetArrayLength());
        Assert.Equal(
            "=cmd|'/c calc'!A1.docx",
            root.GetProperty("files")[0].GetProperty("fileName").GetString());
    }

    [Fact]
    public void Json_UsesCamelCaseConsistently()
    {
        var json = new JsonManifestFormatter().Render(HostileDocument);

        Assert.Contains("\"applicationVersion\"", json, StringComparison.Ordinal);
        Assert.Contains("\"generatedUtc\"", json, StringComparison.Ordinal);
        Assert.DoesNotContain("\"ApplicationVersion\"", json, StringComparison.Ordinal);
    }

    [Fact]
    public void Markdown_EscapesUntrustedFileNames()
    {
        var markdown = new MarkdownManifestFormatter().Render(HostileDocument);

        // The raw pipe would otherwise add a phantom column to the table.
        Assert.DoesNotContain("| =cmd|'/c calc'!A1.docx |", markdown, StringComparison.Ordinal);
        Assert.Contains("\\|", markdown, StringComparison.Ordinal);
    }

    [Fact]
    public void Markdown_WithNoEntries_SaysSoExplicitly()
    {
        var empty = HostileDocument with { Entries = [] };

        Assert.Contains(
            "No files were successfully processed",
            new MarkdownManifestFormatter().Render(empty),
            StringComparison.Ordinal);
    }

    [Fact]
    public void Formatters_ReportTheirOwnFormatAndExtension()
    {
        Assert.Equal(".txt", new PlainTextManifestFormatter().FileExtension);
        Assert.Equal(".csv", new CsvManifestFormatter().FileExtension);
        Assert.Equal(".json", new JsonManifestFormatter().FileExtension);
        Assert.Equal(".md", new MarkdownManifestFormatter().FileExtension);

        Assert.Equal(ManifestFormats.PlainText, new PlainTextManifestFormatter().Format);
        Assert.Equal(ManifestFormats.Csv, new CsvManifestFormatter().Format);
        Assert.Equal(ManifestFormats.Json, new JsonManifestFormatter().Format);
        Assert.Equal(ManifestFormats.Markdown, new MarkdownManifestFormatter().Format);
    }
}
