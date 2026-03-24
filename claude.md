# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

Excelsior is a .NET library for Excel spreadsheet generation using a data-driven approach. It provides a common abstraction (`BookBuilderBase`) with three backend implementations for different spreadsheet libraries.

## Build & Test Commands

All commands must be run from the repository root.

```bash
# Build
dotnet build src --configuration Release

# Run all tests
dotnet test src --configuration Release

# Run a single test project
dotnet test src/ExcelsiorClosedXml.Tests

# Run a specific test
dotnet test src/ExcelsiorClosedXml.Tests --filter "FullyQualifiedName~UsageTests"
```

CI uses AppVeyor (config at `src/appveyor.yml`).

## Architecture

### Core library (`src/Excelsior/`)
- `BookBuilderBase<TBook,TSheet,TStyle,TCell,TColor,TColumn>` — abstract generic base handling sheet registration, build orchestration, and stream/file output
- `RendererBase` — abstract base for writing headers and data rows to a sheet
- `SheetBuilder` / `ISheetBuilder<TModel,TStyle>` — fluent API for per-column configuration (heading, order, width, render, cell style, format)
- `ColumnConfig` — holds per-column settings; `Columns` manages the ordered list
- `Property` / `Properties` — reflection-based property discovery with attribute support
- `ValueRenderer` — static class handling type-to-cell-value conversion (dates, enums, strings, complex types, whitespace trimming)
- Attributes: `ColumnAttribute`, `SplitAttribute`, `IgnoreAttribute` (in `Attributes/`)

### Backend implementations
Each backend has a `BookBuilder` (extends `BookBuilderBase`) and a `Renderer` (extends `RendererBase`):
- **ExcelsiorClosedXml** (`src/ExcelsiorClosedXml/`) — wraps ClosedXML
- **ExcelsiorAspose** (`src/ExcelsiorAspose/`) — wraps Aspose.Cells
- **ExcelsiorSyncfusion** (`src/ExcelsiorSyncfusion/`) — wraps Syncfusion XlsIO

### Test projects
- `Excelsior.Tests` — core library unit tests
- `ExcelsiorClosedXml.Tests` — ClosedXml-specific tests; also compiles shared test files from `ExcelsiorAspose.Tests` via `<Compile Include>` links
- `ExcelsiorAspose.Tests` — Aspose tests (many test files are shared with ClosedXml tests)
- `ExcelsiorSyncfusion.Tests` — Syncfusion tests
- `StaticSettingsTests` — tests for global static settings (date formats, whitespace trimming, enum rendering)
- `SheetRender` — .NET Framework 4.8 utility using Excel Interop. Contains `RenderExcel` which opens all `.verified.xlsx` files, renders each sheet's used range to a bitmap, and saves as `_SheetName.png` alongside the xlsx. These pngs are referenced in the readme. Tests are `[Explicit]` (manual run only).
- `Model` — shared model classes used by tests

Tests use **NUnit** with **Verify** for snapshot testing. Verified snapshots are `.verified.png` and `.verified.txt` files alongside test classes.

## Key Conventions

- Target framework: `net10.0` (with `LangVersion` set to `preview`)
- Central package management via `src/Directory.Packages.props`
- `TreatWarningsAsErrors` is enabled
- Test files are shared between backend test projects using `<Compile Include>` links rather than duplication
- `ValueRenderer` static methods (e.g., `For<T>`, `ForEnums`, `DisableWhitespaceTrimming`) use `[ModuleInitializer]` for global configuration
