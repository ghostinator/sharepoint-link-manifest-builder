# ADR-0007: File identity is (driveId, itemId), never filename or path

- **Status:** Accepted
- **Date:** 2026-09-02

## Context
Manifests are regenerated over time. Files get renamed, moved between folders, and replaced by
different files with the same name. Overlapping targets can reach the same file by two routes.

## Decision
The identity of a file, everywhere in the system, is the ordered pair **(driveId, itemId)**.

- Deduplication across overlapping targets keys on it.
- Manifest update mode matches existing entries on it.
- Renames and moves are therefore detected as *the same file with new metadata*, not as a
  delete plus an add.

Filename and relative path are treated as **display metadata**, never as identity.

## Consequences
- Correct behaviour under rename, move, and same-name-different-file.
- A manifest entry can be updated in place across jobs, preserving history.
- Manifest formats must persist `DriveId` and `ItemId`, so the plain-text format carries both.
- An item genuinely deleted and re-uploaded gets a new `itemId` and is correctly treated as a
  new file.

## Alternatives considered
- **Filename** — breaks on rename and on duplicate names in different folders.
- **Relative path** — breaks on move.
- **`eTag`/`cTag`** — versions the content, not the identity.
