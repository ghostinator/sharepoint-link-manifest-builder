# User guide

## What this application is for

It builds **manifests**: plain lists of files with sharing links, written back into SharePoint or
OneDrive. The point is to hand Microsoft Copilot, or another authorized AI system, an explicit
list of files rather than relying on it to discover them.

It acts as **you**. It can never see or share anything your own account cannot already open.

---

## First run

The setup wizard opens automatically when no tenant is configured.

1. **Welcome** — what the application does, what it accesses, and where it stores data locally.
2. **Choose a setup method** — automatic, existing registration, or manual instructions.
3. **Sign in** — opens your normal web browser. Multi-factor authentication and Conditional
   Access work exactly as your organization configured them. This application never sees your
   password.
4. **Review permissions** — every permission, with why it is needed and what it practically
   allows. Read this page.
5. **Provisioning** — only for automatic setup. Shows the exact changes before making them.
6. **Consent** — opens Microsoft's own consent experience.
7. **Verification** — confirms the result by obtaining a real token, not by assuming.
8. **Complete** — summary and links into the Microsoft Entra admin center.

If your organization does not allow you to create app registrations, choose **Use an existing
app registration** and ask an administrator for the client ID. Everything else works the same.

---

## Selecting what to process

Choosing locations happens inside a job rather than on separate pages. Open **New Link Job**,
stay on **1. Targets**, and use **Browse SharePoint** or **Browse OneDrive**. Each opens as a
step of the job, and **Done — back to targets** returns you to the target list with your
selections added.

### SharePoint

Choose **Browse SharePoint** from the Targets step. You can:

- **Search** by name. Search shows what Microsoft 365 returns for your account; it is not
  necessarily every site in the organization.
- **Paste a URL** — a site, a document library or a specific folder. This is the reliable way to
  reach something search does not surface. URLs copied from the SharePoint address bar work,
  including the long `_layouts` form.

Expand a site to see its document libraries, and a library to see its folders. Tick anything you
want to process, then choose **Add selected as targets**.

### OneDrive

Choose **Browse OneDrive** from the Targets step.

- **Open My OneDrive** for your own files.
- **Open another user's OneDrive** under the expander. Search for the person, select them, and
  open their drive.

Opening a drive collapses the others and expands the new one, so the drive you just asked for is
the one on screen. Opening a drive that is already open reveals it rather than adding a second
copy. **Close** on a drive row closes that one; **Close other users' drives** closes all but your
own. Neither removes targets you have already added — those are listed on the Targets step and
are removed there.

Finding a user does **not** mean their OneDrive can be opened. It may not exist yet, or your
permissions may not reach it. The application tells you which of those it is, and never creates
a OneDrive on someone's behalf.

### Recursion

Each target carries its own **Include subfolders** setting.

- **Off** — only the files directly inside the folder you selected.
- **On** — that folder and everything beneath it.

You can mix both in one job. Change it per target on the Targets tab.

### Overlapping selections

If you select a folder and also its parent, the application notices. By default it keeps the
broader target and processes each file once. You can instead keep the narrower target, or keep
both and rely on deduplication.

Recursion matters here: a **non-recursive** parent does not reach into a subfolder, so those two
targets do not actually overlap and both are kept.

---

## Choosing what to request

### Link permission

| Permission | Effect |
|---|---|
| **View** | Recipients can open and read. Right for a Copilot manifest. |
| **Edit** | Recipients can open and change. Only when collaboration is intended. |

### Link audience

| Audience | Who can use the link |
|---|---|
| **People in the organization** | Anyone signed in to your Microsoft 365 tenant. Usually the right choice: access stays inside the organization. |
| **Specific people** | Only the people you name. |
| **Anyone with the link** | Anyone at all, without signing in, including outside your organization. |

**Specific people** works in two steps, because the Microsoft Graph v1.0 `createLink` action does
not accept recipients. The application creates the link, then separately grants each named person
access. **No email is sent** unless you tick the invitation option.

**Anyone with the link** is the highest-risk option and many organizations disable it. If yours
does, the request is refused and recorded as *blocked by policy*. This application cannot
override your organization's policy and does not pretend to.

### Expiry and other options

- **Expiry** — support depends on your organization's policy and licensing. The expiry Microsoft
  actually applied is recorded in the results, which may differ from what you asked for.
- **Reuse an equivalent existing link** — Microsoft Graph returns an existing link rather than
  creating a duplicate. Recorded as *Reused*.
- **Skip when an equivalent link exists** — checks first and skips the write entirely. Costs one
  extra read per file but avoids a write. Useful on repeat runs over large libraries.
- **Retain inherited permissions** — leave on unless you specifically intend to strip existing
  permissions when sharing an item for the first time.

---

## Manifests

### Per-folder manifests

One manifest inside each folder that contained processed files, listing the files in that
folder. Default name `_sharepoint-links.txt`.

### Master manifests

One manifest listing every successful file beneath the target.

- For a **folder** target, written into the starting folder.
- For a **library** target, written at the library root.
- For a **site** spanning several libraries, you must choose a destination. The application will
  not guess, because guessing could put data somewhere unrelated.

### Formats

Plain text is the default and the only format the application reads back to update in place.
Markdown, CSV and JSON are also available. CSV is protected against formula injection, so a file
named `=cmd|...` cannot execute when the export is opened in a spreadsheet.

### Updating an existing manifest

By default the application **updates safely**: it merges into a manifest it recognises as its
own. Files are matched on drive and item identity, never on name, so a renamed or moved file
updates in place instead of appearing twice.

If the file at that path is one the application did not write, it is **left untouched** and a
timestamped copy is written alongside. Your own document at that path is never overwritten.

If the manifest changed in SharePoint between being read and written, the write is refused and
reported rather than clobbering the change.

---

## Filters

All optional. Excluded by default: temporary and lock files (`~$...`), manifests this
application generates, hidden and system items, and packages such as OneNote notebooks.

You can filter by extension, name pattern (`*` and `?`), modified date and size.

The Filters tab shows a plain-language summary of what is active.

---

## Preview, then run

Always preview first. The preview shows:

- **Checked and confirmed** — things actually verified, such as targets being reachable.
- **Expected** — things that should hold but were not directly confirmed.
- **Cannot be known until the job runs** — principally whether your organization's sharing
  policy permits the link you asked for. Microsoft 365 decides that when each link is requested,
  so no preflight check can answer it honestly.

**Dry run** is on by default. A dry run enumerates, filters and validates while creating no link
and writing no file. Run it first on anything unfamiliar.

**Start** only becomes available after a preview has been built.

---

## While a job runs

Progress shows counts for created, reused, skipped, failed and retried, plus the file currently
being processed.

- **Pause** stops at the next safe point; it never interrupts a request mid-flight.
- **Cancel** stops safely. Results already produced are kept, and manifests are still written for
  them, because those links were really created.

If Microsoft 365 asks the application to slow down, it waits and resumes automatically.

---

## Results

Each file shows an outcome:

| Outcome | Meaning |
|---|---|
| **Created** | A new sharing link was created. |
| **Reused** | An equivalent link already existed and was returned. |
| **Existing** | A matching link was found without requesting a new one. |
| **Skipped** | Not processed, by configuration or filtering. |
| **Policy blocked** | Your organization's policy refused this kind of link. |
| **Access denied** | You do not have permission to share this item. |
| **Unsupported** | This item type cannot be shared this way. |
| **Failed** | Something else went wrong; the detail column explains. |

**Reused is never reported as Created.** If a manifest says a link was reused, it genuinely
already existed.

### Retrying

**Retry failures** re-runs only the failures that could plausibly succeed — throttling and
transient network faults. Failures caused by policy or permissions are not offered, because
retrying them cannot change the outcome.

---

## Saved profiles and history

**Saved Profiles** stores a job configuration for reuse. **Job History** records every run, with
counts, manifest locations and sanitized errors, and can load a past configuration back into the
job page. Neither stores any credential.

---

## Getting help

The **Diagnostics** page has a connectivity test and a sanitized bundle export that shows you
exactly which categories it will include before writing anything. See
[TROUBLESHOOTING.md](TROUBLESHOOTING.md) for specific errors.
