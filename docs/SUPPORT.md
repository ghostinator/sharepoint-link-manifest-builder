# Support

## Before asking

1. **[TROUBLESHOOTING.md](TROUBLESHOOTING.md)** covers most errors and what to do about each.
2. **Diagnostics tab in Settings** — run the connectivity test and read the sanitized recent errors.
3. **Permissions tab in Settings** — a surprising number of problems are a missing scope.
4. **Try a dry run** — it isolates enumeration and permission problems from link creation.

## Where to ask

| Situation | Where |
|---|---|
| Something does not work as documented | Open a bug report |
| A capability is missing | Open a feature request |
| A question about permissions, consent or tenant configuration | Open a security-configuration question |
| **A security vulnerability** | **Do not open an issue.** See [SECURITY.md](../SECURITY.md) |

Issue templates: <https://github.com/ghostinator/sharepoint-link-manifest-builder/issues>

> Placeholder. A publisher replaces this before distribution.

## What to include

- Application version, from the About page
- Operating system and version
- What you did, what happened, what you expected
- The error message
- **The correlation ID** — the single most useful item for a Graph problem
- Whether the registration was created automatically or supplied by your organization

## What never to include

This is a public issue tracker.

- Access tokens, refresh tokens, authorization codes
- Client secrets, certificates, private keys
- Sharing links produced by this or any application
- Real SharePoint or OneDrive URLs
- Email addresses or user principal names
- Your tenant ID or client ID
- Tenant configuration exports, audit logs, or an unreviewed diagnostic bundle

Substitute `contoso.sharepoint.com`, `user@contoso.com` and
`00000000-0000-0000-0000-000000000000`.

If you have already posted something sensitive: delete the comment, then **revoke and rotate**
the exposed value. Deleting a comment does not un-disclose it.

## Response expectations

This is provided as-is with no service-level agreement. Maintainers respond as capacity allows.
Security reports are prioritised — see [SECURITY.md](../SECURITY.md) for those targets.

## Things support cannot fix

Some outcomes are your organization's policy working correctly, not defects:

- **A link refused by policy.** If external sharing is disabled for a site, no setting in this
  application will produce an external link. Ask your administrator.
- **Access denied to content.** Delegated access is bounded by your own permissions. Administrator
  consent does not widen it.
- **A user's OneDrive that does not exist.** The user must open OneDrive once. This application
  will not provision one for them.
- **A site missing from search.** The index is not exhaustive. Paste the URL instead.

## Commercial support

None is offered. A publisher distributing this software may offer their own; see the About page
for their support URL.
