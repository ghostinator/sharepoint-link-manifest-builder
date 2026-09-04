# Third-party notices

This product includes the third-party components listed below. A machine-readable inventory,
including transitive dependencies and their resolved versions, is produced by
`./scripts/sbom.sh` and attached to every release.

Regenerate the inventory with:

```bash
dotnet list SharePointLinkManifestBuilder.slnx package --include-transitive
```

---

## Redistributed in release artifacts

### Avalonia UI

- **Licence:** MIT
- **Project:** <https://avaloniaui.net>
- **Source:** <https://github.com/AvaloniaUI/Avalonia>
- **Used for:** the cross-platform user interface
- **Packages:** `Avalonia`, `Avalonia.Desktop`, `Avalonia.Themes.Fluent`, `Avalonia.Fonts.Inter`,
  `Avalonia.Controls.DataGrid`

### CommunityToolkit.Mvvm

- **Licence:** MIT
- **Source:** <https://github.com/CommunityToolkit/dotnet>
- **Used for:** MVVM source generators for observable properties and commands

### Microsoft Authentication Library for .NET (MSAL)

- **Licence:** MIT
- **Source:** <https://github.com/AzureAD/microsoft-authentication-library-for-dotnet>
- **Used for:** all Microsoft Entra authentication
- **Packages:** `Microsoft.Identity.Client`, `Microsoft.Identity.Client.Extensions.Msal`

### Microsoft.Extensions.*

- **Licence:** MIT
- **Source:** <https://github.com/dotnet/runtime>
- **Used for:** dependency injection, configuration, logging and HTTP client factory
- **Packages:** `Microsoft.Extensions.DependencyInjection`, `Microsoft.Extensions.Configuration`
  (and `.Binder`, `.Json`, `.EnvironmentVariables`), `Microsoft.Extensions.Logging`,
  `Microsoft.Extensions.Http`, `Microsoft.Extensions.Options`

### .NET runtime

- **Licence:** MIT
- **Source:** <https://github.com/dotnet/runtime>
- **Used for:** the application runtime, embedded in self-contained builds

### Inter typeface

- **Licence:** SIL Open Font License 1.1
- **Source:** <https://github.com/rsms/inter>
- **Used for:** the bundled UI typeface, via `Avalonia.Fonts.Inter`

---

## Build and test only — not redistributed

### xunit / xunit.v3

- **Licence:** Apache-2.0
- **Source:** <https://github.com/xunit/xunit>

### NSubstitute

- **Licence:** BSD-3-Clause
- **Source:** <https://github.com/nsubstitute/NSubstitute>

### coverlet

- **Licence:** MIT
- **Source:** <https://github.com/coverlet-coverage/coverlet>

### Microsoft.NET.Test.Sdk

- **Licence:** MIT
- **Source:** <https://github.com/microsoft/vstest>

---

## Licence texts

Full licence texts are distributed with each package and are available at the source URLs above.

- MIT: <https://opensource.org/licenses/MIT>
- Apache-2.0: <https://www.apache.org/licenses/LICENSE-2.0>
- BSD-3-Clause: <https://opensource.org/licenses/BSD-3-Clause>
- SIL OFL 1.1: <https://openfontlicense.org>

---

## Microsoft services

This application communicates with Microsoft identity and Microsoft Graph endpoints. Those
services are governed by the agreement between your organization and Microsoft, not by this
project's licence. This project is not affiliated with or endorsed by Microsoft.

Microsoft, Microsoft 365, SharePoint, OneDrive, Microsoft Entra, Microsoft Graph and Microsoft
Copilot are trademarks of the Microsoft group of companies.
