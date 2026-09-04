# ADR-0009: Tiered, opt-in permission scopes

- **Status:** Accepted
- **Date:** 2026-09-02

## Context
It is tempting to request `Sites.ReadWrite.All` plus `Application.ReadWrite.All` once and
never think about permissions again. That maximizes blast radius and gets applications
blocked by security review.

## Decision
Scopes are grouped into tiers and requested incrementally.

**Operation tiers**
| Tier | Scopes | When requested |
|---|---|---|
| Discovery (read-only) | `User.Read`, `Sites.Read.All`, `Files.Read.All` | Browsing and dry runs |
| Standard | Discovery + `Files.ReadWrite.All` | Creating links and writing manifests |
| People picker | + `User.ReadBasic.All` | Only if the User OneDrive source is enabled |
| Broad (flagged) | + `Sites.ReadWrite.All` | Only where a library rejects the standard write |

**Bootstrap tiers**
| Tier | Scopes | When requested |
|---|---|---|
| Create-only (default) | `User.Read`, `AppRegistration.Create` | Automatic setup happy path |
| Manage (opt-in) | `Application.ReadWrite.All` | Repair / Replace / delete / deep verification |

`Directory.ReadWrite.All` and `Directory.Read.All` are documented as higher-privileged
alternatives that this application deliberately never requests.

## Consequences
- The consent prompt an administrator sees is defensible line by line, and the UI states the
  purpose and practical data-access impact of each scope.
- Broad scopes are visibly marked as broad, with the least-privilege alternative named.
- More code: the app must handle "this action needs a scope you have not granted" gracefully
  and request incrementally rather than assuming.

## Alternatives considered
- **One maximal scope set** — simpler code, far worse security posture, and likely to fail
  customer security review.
