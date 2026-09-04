# Publishing this repository to GitHub

## Read this first

This repository has never been pushed anywhere. Before it is, run the publication safety gate:

```bash
./scripts/scan-secrets.sh
```

It scans the working tree **and the full git history**, and it fails closed. History matters
because `.gitignore` does nothing about content that was already committed, and deleting a file
in a later commit does not remove it from the repository.

**Do not push until it passes.**

## Pre-publication checklist

- [ ] `./scripts/scan-secrets.sh` passes over full history
- [ ] `./scripts/format.sh --check` passes
- [ ] `./scripts/build.sh` passes
- [ ] `./scripts/test.sh` passes
- [ ] `git status` is clean
- [ ] `git log --stat` reviewed — no file that should not be there was ever committed
- [ ] A licence has been chosen, or the repository stays private ([LICENSE-SELECTION.md](../LICENSE-SELECTION.md))
- [ ] `SECURITY.md` contact is real, or private vulnerability reporting is enabled
- [ ] No screenshot shows real tenant data
- [ ] Placeholders are acceptable for now, or replaced

## Manually reviewing history

The scanner is a safety net, not a substitute for looking.

```bash
git log --oneline --stat | less           # every file ever touched
git log --diff-filter=A --name-only --format='%h' | sort -u | less   # every file ever added
git count-objects -vH                     # an unexpectedly large repo can indicate a committed binary
```

## Creating the repository with the GitHub CLI

```bash
gh auth status                            # confirm the intended account
gh api user --jq .login                   # confirm the owner
gh api user/orgs --jq '.[].login'         # organizations available
```

Create it **private** first. Making a repository public is effectively irreversible: anything
visible while public may have been cloned, forked, cached or indexed.

```bash
gh repo create <OWNER>/sharepoint-link-manifest-builder \
  --private \
  --source . \
  --remote origin \
  --description "Build explicit manifests of SharePoint and OneDrive file links for Microsoft Copilot"
```

`--source .` sets the remote without pushing. Verify, then push:

```bash
git remote -v
git push -u origin main
git push -u origin feature/initial-implementation
```

## Without the GitHub CLI

Create the repository through the web interface — **private**, with no README, `.gitignore` or
licence, since this repository already has them. Then:

```bash
git remote add origin https://github.com/<OWNER>/sharepoint-link-manifest-builder.git
git remote -v
git push -u origin main
```

## After the first push

1. **Settings → Security** — enable secret scanning, push protection, Dependabot alerts and
   Dependabot security updates.

   **CodeQL is skipped while the repository is private.** Analysis runs, but uploading results
   needs GitHub Advanced Security, a paid add-on for private repositories and free for public
   ones. The `codeql` job is conditioned on `visibility == 'public'`, so it begins running
   automatically when the repository is made public — no workflow edit needed. The rest of the
   Security workflow (publication safety scan, dependency audit, SBOM) runs either way.
2. **Settings → Security → Advisories** — enable private vulnerability reporting, which
   `SECURITY.md` points at.
3. **Settings → Actions** — confirm the workflows appear and pass. The `workflow` token scope is
   needed to push `.github/workflows`.
4. **Settings → Branches** — protect `main`: require a pull request, require CI to pass, and
   disallow force-pushes.
5. **CODEOWNERS** — copy `.github/CODEOWNERS.example` to `.github/CODEOWNERS` and set real
   owners. It ships as `.example` because committing a guessed owner would request review from
   someone who never agreed to it.

## Going public

Only after:

- [ ] A licence file exists
- [ ] The safety scan passes over full history
- [ ] Placeholders are replaced, or their presence is acceptable and documented
- [ ] `SECURITY.md` has a working private reporting channel
- [ ] Someone other than the author has reviewed the diff

```bash
./scripts/scan-secrets.sh                 # one more time
gh repo edit <OWNER>/sharepoint-link-manifest-builder --visibility public
```

Making a repository public is a one-way door in practice. Treat it as such.

## If a secret is pushed

**Revoke and rotate first.** The moment it reached a remote it must be treated as compromised.
Removing it from git afterwards does not un-disclose it: clones, forks, caches and search
indexes may already hold it.

```bash
# 1. REVOKE AND ROTATE THE CREDENTIAL. Do this before anything else.

# 2. Remove it from history (git-filter-repo, not filter-branch)
pip install git-filter-repo
git filter-repo --path path/to/leaked-file --invert-paths

# 3. Re-add the remote (filter-repo removes it deliberately) and force-update
git remote add origin https://github.com/<OWNER>/<REPO>.git
git push --force --all
git push --force --tags

# 4. Tell anyone with a clone or fork to re-clone; their copy still has it.
# 5. Ask GitHub Support to purge cached views of the affected commits.
# 6. Re-run the scan.
./scripts/scan-secrets.sh
```

Never skip step 1 in favour of "just removing the commit". Rewriting history is cleanup;
rotation is the actual fix.

## Publishing a release

Releases are produced by the tagged workflow — see [RELEASE-PROCESS.md](RELEASE-PROCESS.md).

Release artifacts are **unsigned and un-notarized**, and the notes say so. Do not remove that
statement. It is what tells a user the SHA-256 checksum is their only integrity check.

## Current state of this repository

| Item | State |
|---|---|
| Local repository | Initialized, `main` plus `feature/initial-implementation` |
| Commits | Focused conventional commits |
| Remote | Set only if the steps above were run |
| Pushed | Only if the steps above were run |
| Visibility | Private unless explicitly changed |
| Releases | None |

Nothing in this document happens automatically. Every push, and every visibility change, is a
deliberate action a human takes.
