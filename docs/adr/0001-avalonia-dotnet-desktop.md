# ADR-0001: Avalonia UI on .NET 10 for the cross-platform desktop client

- **Status:** Accepted
- **Date:** 2026-09-02

## Context
The product must run on Windows, macOS and Linux with a native-feeling, accessible,
high-DPI GUI. It must talk to Microsoft identity and Microsoft Graph, store tokens in
OS-native secure storage, and ship as a self-contained artifact.

## Decision
Use **Avalonia UI 12.1.2** on **.NET 10** (the newest stable release; .NET 8 is the stated
floor). MVVM via `CommunityToolkit.Mvvm` source generators. Compiled XAML bindings are on
by default (`AvaloniaUseCompiledBindingsByDefault`).

Avalonia 12 was verified by compiling a representative spike — `TreeDataTemplate`,
three-state `CheckBox`, `TabControl`, `Expander`, compiled bindings — before adoption,
rather than assuming v11 idioms carry forward.

## Consequences
- One C# codebase and one UI toolkit for all three desktop platforms.
- Direct access to MSAL and `Microsoft.Extensions.*` without an interop bridge.
- `dotnet publish -r <rid>` produces self-contained single-file artifacts.
- Avalonia is not the OS-native toolkit, so platform look-and-feel is approximated by the
  Fluent theme rather than inherited.

## Alternatives considered
- **Electron** — explicitly disfavoured by the brief. Would require a second language
  runtime, a separate Node dependency tree, and an out-of-process identity story. No
  documented blocker makes Avalonia unsuitable, so Electron is not used.
- **MAUI** — no first-class Linux desktop target.
- **WPF / WinUI** — Windows only.
- **Avalonia 11.3.20** — safe and well-known, but 12.1.2 is stable and the spike passed.
