# Rollback

Two separate concerns: rolling back **source changes**, and rolling back an **installed
application**.

---

## Source

### Revert a commit

`git revert` creates a new commit undoing the change, preserving history. This is correct for
anything already pushed.

```bash
git revert <commit-sha>
git revert --no-commit <sha1> <sha2>   # several, as one commit
git revert --abort                     # bail out of a conflicted revert
```

### Revert a merge

A merge has two parents, so git needs to know which one to keep. `-m 1` keeps the branch you
merged **into**, which is almost always what you want.

```bash
git log --oneline --merges -5
git revert -m 1 <merge-commit-sha>
```

Note: reverting a merge makes it hard to re-merge that branch later. To bring the work back,
revert the revert.

### Restore a single file from an earlier commit

```bash
git log --oneline -- path/to/file
git checkout <commit-sha> -- path/to/file
git commit -m "fix: restore path/to/file to its state at <sha>"
```

### Check out a release tag

```bash
git tag -l 'v*'
git checkout v1.2.3                      # detached HEAD, for inspection
git checkout -b investigate-1.2.3 v1.2.3 # a branch, to work from
```

### Create a hotfix branch from a release

```bash
git checkout -b hotfix/1.2.4 v1.2.3
# fix, with a regression test
git commit -m "fix: <what>"
git tag -a v1.2.4 -m "SharePoint Link Manifest Builder 1.2.4"
git checkout main && git merge hotfix/1.2.4
```

### Compare releases

```bash
git diff v1.2.2..v1.2.3                       # full diff
git diff v1.2.2..v1.2.3 --stat                # summary
git log v1.2.2..v1.2.3 --oneline              # commits
git diff v1.2.2..v1.2.3 -- docs/GRAPH-OPERATIONS.md   # did Graph usage change?
```

### What not to do on a shared branch

Do **not** use `git reset --hard` followed by a force-push on `main` or any branch someone else
may have pulled. It rewrites history under their feet, and anyone who already pulled will
re-introduce the removed commits on their next push.

The one exception is removing a leaked secret from history, which genuinely requires a rewrite.
Even then: revoke and rotate the credential **first**, because rewriting history does not
un-disclose anything.

---

## Installed application

### Roll back to an earlier version

1. Download the earlier release artifact.
2. **Verify its SHA-256 checksum.** Artifacts are unsigned, so this is the only integrity check.
3. Close the application.
4. Replace the program directory with the older build.
5. Start it.

Settings, profiles, job history and sign-in details live in the application-data directory and
are untouched by this.

### Compatibility across versions

| Data | Newer reads older | Older reads newer |
|---|---|---|
| Settings | Yes | Yes; unknown fields ignored |
| Tenant configuration | Yes | Yes |
| Saved profiles | Yes | Usually; a profile using a newer option loses that option |
| Job history | Yes | Yes; unknown fields ignored |
| Plain-text manifests | Yes, same major schema | **No** — refused, not mis-parsed |

The manifest rule is deliberate. An older build encountering a newer **major** schema version
refuses to parse it and writes a timestamped copy instead, rather than silently reading it
wrongly and then overwriting it with a downgraded document.

Local state files are read defensively: a file that cannot be parsed is treated as absent rather
than blocking startup, so a rollback cannot leave the application unable to launch.

### Rolling back does not undo tenant changes

Sharing links already created stay created. Manifests already written stay written.

To undo a job's effects:

- **Sharing links** — remove them in SharePoint or OneDrive (*Manage access*), or with
  PowerShell. This application does not delete links it created, deliberately: a bulk
  permission-removal tool is a different and more dangerous product.
- **Manifests** — delete the manifest files, or restore an earlier version from SharePoint's
  version history.

Consult **Job History** for exactly what a run did: targets, counts and manifest locations.

### Rolling back a tenant configuration change

| Change | How to undo |
|---|---|
| Local configuration | *Remove local configuration* on the Permissions page |
| Consent | Revoke it in the Microsoft Entra admin center (enterprise applications) |
| App registration created by the application | Delete it in Entra, or use the guarded delete on the Permissions page |

The Permissions page shows a local audit history of every tenant change the application made,
which is the fastest way to see what needs undoing.
