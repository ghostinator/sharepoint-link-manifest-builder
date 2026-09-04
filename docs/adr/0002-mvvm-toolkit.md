# ADR-0002: CommunityToolkit.Mvvm for MVVM plumbing

- **Status:** Accepted
- **Date:** 2026-09-02

## Context
Roughly twenty view models need change notification, commands, async command state, and
validation, without hand-written `INotifyPropertyChanged` boilerplate.

## Decision
Use `CommunityToolkit.Mvvm` (`ObservableObject`, `[ObservableProperty]`, `[RelayCommand]`).
It is source-generator based, so there is no runtime reflection cost and no
`INotifyPropertyChanged` weaving.

## Consequences
- View models stay short and readable.
- `[RelayCommand]` with `CancellationToken` parameters gives async commands cancellation
  for free, which the job screens rely on.
- Generated members are invisible in source, which can surprise readers; property names are
  therefore always referenced through the generated PascalCase form.

## Alternatives considered
- **ReactiveUI** — powerful, but a large conceptual surface and heavier dependency for what
  is mostly request/response UI.
- **Hand-rolled `INotifyPropertyChanged`** — needless boilerplate across ~20 view models.
