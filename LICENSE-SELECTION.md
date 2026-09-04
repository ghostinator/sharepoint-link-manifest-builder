# Licence selection

**Decided: GNU General Public License, version 3 or later (`GPL-3.0-or-later`).**

The full text is in [LICENSE](LICENSE), taken verbatim from
<https://www.gnu.org/licenses/gpl-3.0.txt>.

## Why version 3, and why "or later"

Version 3 rather than 2 because two dependencies are Apache-2.0, which is compatible with GPLv3
but **not** with GPLv2 — its patent-termination clause counts as an additional restriction under
version 2. Choosing version 2 would have made the dependency set unusable.

"Or later" follows the boilerplate the FSF recommends and keeps the door open to a future
version without needing every contributor's agreement to relicense.

## What this means in practice

- Anyone may use, study, modify and redistribute the software.
- A distributed derivative must also be GPL-3.0-or-later and must ship its source.
- Running it inside an organisation is not distribution, so internal modifications carry no
  publication obligation.
- There is no warranty. Sections 15 and 16 disclaim it.

## Dependency compatibility

Every third-party dependency is MIT, BSD-3-Clause or Apache-2.0. All three are one-way
compatible with GPL-3.0: their code may be combined into a GPL-3.0 work. None imposes a
condition that GPL-3.0 cannot satisfy. The inventory is in
[THIRD-PARTY-NOTICES.md](THIRD-PARTY-NOTICES.md).

## Not done

Per-file copyright headers. The GPL recommends them but does not require them, and the licence
applies to the work as a whole through this file and `LICENSE`. Adding them across the tree is a
mechanical change that can be made later without affecting the licence's validity.

---

## Original analysis, retained for the record

**No licence has been chosen yet.** Until a `LICENSE` file exists in the repository root,
default copyright applies: nobody is granted permission to use, copy, modify or distribute this
software. That is deliberate — a licence is a decision for the publisher, and guessing one would
be worse than leaving it open.

## Why this matters before publishing

A public repository with no licence is *not* open source. People can read it, but they cannot
legally use it, and most organizations will not touch it. Choose before making the repository
public.

## Options

### MIT — recommended default

Short, permissive, universally understood. Anyone may use, modify and redistribute, including
commercially, provided the copyright notice is retained.

**Choose if** you want the widest possible adoption and do not mind proprietary derivatives.

### Apache License 2.0 — recommended where patents matter

Permissive like MIT, plus an express patent grant and a termination clause if a user brings a
patent suit. Also requires stating significant changes.

**Choose if** you or your users care about patent protection. Many enterprises prefer this.

Notably, every dependency this project uses is MIT-licensed, which is compatible with Apache-2.0
for a combined work.

### GNU GPL v3 — copyleft

Derivatives must also be GPL v3 and publish source.

**Choose if** you want improvements to stay open. Note this substantially reduces enterprise
adoption, which matters for an administrator tool.

### Mozilla Public License 2.0 — file-level copyleft

A middle ground: modifications to MPL files stay MPL, but the work can be combined with
proprietary code.

**Choose if** you want the files you wrote to stay open without restricting the wider system.

### Proprietary / all rights reserved

Keep the current state, or add an explicit proprietary notice. The repository should then be
private, or public read-only with the restriction stated clearly.

## Dependency compatibility

Every dependency is MIT-licensed, so all of the options above are viable.

| Dependency | Licence |
|---|---|
| Avalonia UI | MIT |
| CommunityToolkit.Mvvm | MIT |
| Microsoft.Identity.Client (MSAL) | MIT |
| Microsoft.Identity.Client.Extensions.Msal | MIT |
| Microsoft.Extensions.* | MIT |
| .NET runtime and SDK | MIT |
| xunit | Apache-2.0 |
| NSubstitute | BSD-3-Clause |

The test-only dependencies are not redistributed in a release artifact.

## How to apply a licence

1. Copy the full licence text to `LICENSE` in the repository root. Use the canonical text from
   <https://choosealicense.com> or <https://spdx.org/licenses/>; do not retype it.
2. Set the copyright line to the real holder and year.
3. Update `Directory.Build.props`: replace the `Copyright` placeholder and add
   `<PackageLicenseExpression>` with the SPDX identifier, for example `MIT`.
4. Add a Licence section to `README.md`.
5. Update `THIRD-PARTY-NOTICES.md` if the combined licensing changes.
6. Commit as `docs: add <name> licence`.

## Recommendation

**Apache-2.0.** This is an administrator tool that touches tenant configuration and content
permissions. Enterprises are its audience, and enterprise legal review is markedly more
comfortable with an express patent grant than with MIT's silence on the subject. The additional
obligations are trivial for both publisher and user.
