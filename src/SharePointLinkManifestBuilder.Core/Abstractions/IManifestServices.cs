using SharePointLinkManifestBuilder.Core.Models;

namespace SharePointLinkManifestBuilder.Core.Abstractions;

/// <summary>Renders a manifest document to text in one format.</summary>
public interface IManifestFormatter
{
    /// <summary>The format this formatter produces.</summary>
    ManifestFormats Format { get; }

    /// <summary>The file extension for this format, including the dot.</summary>
    string FileExtension { get; }

    /// <summary>
    /// Renders the document to text. Implementations must neutralize untrusted content: file names and
    /// paths originate in SharePoint and are attacker-influenced.
    /// </summary>
    string Render(ManifestDocument document);
}

/// <summary>Parses a previously written manifest so it can be updated in place.</summary>
public interface IManifestParser
{
    /// <summary>The format this parser reads.</summary>
    ManifestFormats Format { get; }

    /// <summary>
    /// Attempts to parse manifest content.
    /// <para>
    /// Failure is a normal outcome meaning "this application did not write this file", which
    /// the conflict policy uses to avoid overwriting a foreign file. Parsing is defensive:
    /// unrecognized lines are ignored and no field is ever interpreted as an instruction.
    /// </para>
    /// </summary>
    ManifestParseResult TryParse(string content);
}

/// <summary>Merges a newly produced manifest with one that already exists.</summary>
public interface IManifestMerger
{
    /// <summary>
    /// Merges by identity: entries are matched on (driveId, itemId), never on name or path, so
    /// renames and moves update in place instead of creating duplicates.
    /// </summary>
    /// <param name="existing">The manifest already in the destination.</param>
    /// <param name="incoming">The manifest produced by this run.</param>
    /// <param name="missingEntryPolicy">What to do with entries not seen on this run.</param>
    ManifestDocument Merge(
        ManifestDocument existing,
        ManifestDocument incoming,
        MissingEntryPolicy missingEntryPolicy);
}

/// <summary>Assembles manifest documents from a run's results.</summary>
public interface IManifestBuilder
{
    /// <summary>Builds one manifest per folder that contained a successful file.</summary>
    IReadOnlyList<(string FolderItemId, string FolderPath, ManifestDocument Document)> BuildPerFolderManifests(
        JobConfiguration configuration,
        ProcessingTarget target,
        IReadOnlyList<LinkResult> results,
        string applicationVersion);

    /// <summary>Builds a single manifest covering every successful file for a target.</summary>
    ManifestDocument BuildMasterManifest(
        JobConfiguration configuration,
        ProcessingTarget target,
        IReadOnlyList<LinkResult> results,
        string applicationVersion);

    /// <summary>Builds one combined manifest spanning every target in a job.</summary>
    ManifestDocument BuildCombinedMasterManifest(
        JobConfiguration configuration,
        IReadOnlyList<LinkResult> results,
        string applicationVersion);
}
