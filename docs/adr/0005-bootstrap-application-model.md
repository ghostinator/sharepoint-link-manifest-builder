# ADR-0005: A separate, publisher-owned bootstrap application for automatic tenant setup

- **Status:** Accepted
- **Date:** 2026-09-02

## Context
Automatic setup must create a tenant-specific app registration — but creating one requires an
identity that already exists in the tenant. This is a chicken-and-egg problem.

## Decision
Two distinct identities:

1. **Bootstrap application** — publisher-owned, multitenant, public client, no secret. Used
   *only* during onboarding. It creates and configures the tenant-specific registration and
   verifies the outcome. It never enumerates SharePoint content and never creates a
   sharing link.
2. **Permanent tenant-specific application** — created in the customer tenant, single-tenant
   (`AzureADMyOrg`), public client. All normal operation uses this client ID.

The bootstrap client ID is **configuration, not source**. It is supplied via build
configuration, environment variable, or the wizard itself. When absent, automatic setup is
disabled with a clear explanation and the existing-registration path remains fully functional.
No client ID is fabricated or committed.

Bootstrap scopes are tiered — see ADR-0009.

## Consequences
- Onboarding works without asking an administrator to hand-build a registration.
- The blast radius of the bootstrap identity is small and time-bounded, and it is visible in
  the tenant as a distinct enterprise application that can be reviewed or blocked.
- A publisher must own and publish the bootstrap app before automatic setup can ship.
- Every tenant modification is previewed before execution and written to a local sanitized
  audit history.

## Alternatives considered
- **Ship a single multitenant app and skip per-tenant registration** — simpler, but the
  customer then cannot own, audit, restrict, or delete the app identity that touches their
  content. Rejected.
- **Require manual registration always** — safe but fails the brief's core UX goal. Retained
  as Path B, not as the only path.
