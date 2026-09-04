using SharePointLinkManifestBuilder.Core.Abstractions;
using SharePointLinkManifestBuilder.Core.Models;

namespace SharePointLinkManifestBuilder.Core.Manifests;

/// <summary>
/// Merges a freshly produced manifest into one that already exists at the destination.
/// <para>
/// Entries are matched on (driveId, itemId) and never on file name or path. That is what makes
/// a rename or a move update an existing entry in place instead of orphaning it and adding a
/// duplicate. See docs/adr/0007-file-identity-drive-item.md.
/// </para>
/// </summary>
public sealed class ManifestMerger : IManifestMerger
{
    /// <inheritdoc />
    public ManifestDocument Merge(
        ManifestDocument existing,
        ManifestDocument incoming,
        MissingEntryPolicy missingEntryPolicy)
    {
        ArgumentNullException.ThrowIfNull(existing);
        ArgumentNullException.ThrowIfNull(incoming);

        var incomingByIdentity = incoming.Entries
            .GroupBy(e => e.IdentityKey, StringComparer.Ordinal)
            .ToDictionary(g => g.Key, g => g.Last(), StringComparer.Ordinal);

        var merged = new List<ManifestEntry>(incoming.Entries.Count + existing.Entries.Count);
        var consumed = new HashSet<string>(StringComparer.Ordinal);

        // Existing entries first, so a manifest keeps a stable order across runs rather than
        // reshuffling every time and producing a noisy diff in version history.
        foreach (var previous in existing.Entries)
        {
            if (incomingByIdentity.TryGetValue(previous.IdentityKey, out var fresh))
            {
                // Seen again this run: the fresh record wins, including a new name or path.
                merged.Add(fresh with { IsMissing = false });
                consumed.Add(previous.IdentityKey);
                continue;
            }

            switch (missingEntryPolicy)
            {
                case MissingEntryPolicy.Preserve:
                    merged.Add(previous);
                    break;

                case MissingEntryPolicy.Mark:
                    merged.Add(previous with { IsMissing = true });
                    break;

                case MissingEntryPolicy.Remove:
                    break;

                default:
                    merged.Add(previous);
                    break;
            }
        }

        // Newly discovered files, in discovery order.
        foreach (var fresh in incoming.Entries.Where(e => !consumed.Contains(e.IdentityKey)))
        {
            merged.Add(fresh);
        }

        // Counts must describe the merged document, not just this run's contribution.
        var header = incoming.Header with
        {
            SuccessfulFiles = merged.Count(e => !e.IsMissing),
            ReusedLinks = merged.Count(e =>
                string.Equals(e.Status, "Reused", StringComparison.OrdinalIgnoreCase)
                || string.Equals(e.Status, "Existing", StringComparison.OrdinalIgnoreCase)),
            SkippedFiles = incoming.Header.SkippedFiles,
            FailedFiles = incoming.Header.FailedFiles,
        };

        return new ManifestDocument
        {
            Header = header,
            Entries = merged,
            WasGeneratedByThisApplication = true,
        };
    }
}
