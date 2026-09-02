# ADR-0003: A centralized raw-HTTP Graph client instead of the Microsoft Graph SDK

- **Status:** Accepted
- **Date:** 2026-09-02

## Context
The brief allows the Graph SDK "where it improves reliability" and raw HTTP "behind a
centralized service where the SDK is unsuitable". Two requirements dominate:

1. The standard test suite must not need a live tenant, yet must cover pagination,
   throttling with `Retry-After`, ETag conflicts, `207 Multi-Status`, and cancellation.
2. Graph semantics this product depends on are subtle — `createLink` returning **200** for a
   reused link versus **201** for a new one is a product-visible distinction.

## Decision
Implement `IGraphApiClient` over `HttpClient` (via `IHttpClientFactory`), with
`System.Text.Json` request/response models, inside `SharePointLinkManifestBuilder.Graph`.
All Graph traffic goes through it. **No `HttpClient` or Graph type is referenced from a
view model.** Identity remains MSAL — OAuth is never hand-rolled.

## Consequences
- Tests inject a fake `HttpMessageHandler` and exercise the real client: real pagination
  loops, real retry policy, real status-code handling. The mock is at the transport boundary,
  which is the only place a mock is honest.
- Status codes, `Retry-After`, `ETag`/`If-Match`, and `@odata.nextLink` are handled explicitly
  and visibly, rather than inside generated request builders.
- Smaller self-contained publish; no Kiota abstraction layer.
- **Cost:** request/response shapes are maintained by hand. Mitigated by keeping DTOs minimal
  (only fields actually consumed) and documenting every operation in
  `docs/GRAPH-OPERATIONS.md`.

## Alternatives considered
- **Microsoft.Graph v5 (Kiota)** — request builders are difficult to fake without testing the
  mock rather than the code, and it obscures exactly the status-code nuances this product
  must surface.
