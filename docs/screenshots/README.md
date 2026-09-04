# Screenshots

Placeholders. No screenshots are committed yet, because a real screenshot of this application
necessarily shows a real tenant: site names, library names, folder structure, file names, user
principal names and sometimes sharing links.

## Before adding one

Capture against a **dedicated test tenant with synthetic content**, never a production tenant.
Then check the image for:

- Tenant name and tenant ID in the header or the Permissions tab in Settings
- The signed-in account's user principal name
- Real site, library, folder and file names
- Sharing links in the results grid
- Client ID on the Permissions or About page
- Anything visible in a background window or the OS title bar

Redacting by blurring is unreliable at low resolution. Prefer synthetic content over redaction.

## Expected files

| File | Page | Notes |
|---|---|---|
| `01-home.png` | Home | Connected state with synthetic tenant |
| `02-tenant-setup.png` | Tenant Setup | Permission review page is the most useful one |
| `03-sharepoint-browser.png` | SharePoint Sites | Tree expanded two levels |
| `04-onedrive-browser.png` | OneDrive | My OneDrive expanded |
| `05-job-targets.png` | New Link Job | Several mixed targets |
| `06-job-preview.png` | New Link Job | Preview showing validated / expected / unknown |
| `07-job-results.png` | New Link Job | Results with mixed Created and Reused |
| `08-permissions.png` | Permissions | Granted and missing permissions |
| `09-diagnostics.png` | Diagnostics | Bundle category disclosure |

Reference them from `README.md` once they exist.
