# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

Excelsior is a .NET library for Excel spreadsheet generation using a data-driven approach. It uses [DocumentFormat.OpenXml](https://github.com/dotnet/Open-XML-SDK) for spreadsheet creation and [OpenXmlHtml](https://github.com/SimonCropp/OpenXmlHtml) for HTML cell rendering.

## Build & Test Commands

All commands must be run from the repository root.

```bash
# Build
dotnet build src --configuration Release

# Run all tests
dotnet test src --configuration Release

# Run a single test project
dotnet test src/Excelsior.Tests

# Run a specific test
dotnet test src/Excelsior.Tests --filter "FullyQualifiedName~UsageTests"
```

CI uses AppVeyor (config at `src/appveyor.yml`).

## Architecture

### Core library (`src/Excelsior/`)
- `BookBuilder` — entry point; handles sheet registration, build orchestration, and stream/file output
- `Renderer` — writes headers and data rows to a sheet
- `SheetBuilder` / `ISheetBuilder<TModel>` — fluent API for per-column configuration (heading, order, width, render, cell style, format)
- `ColumnConfig` — holds per-column settings; `Columns` manages the ordered list
- `CellStyle` — style container for cell formatting (font, fill, alignment)
- `StyleManager` — deduplicates and caches OpenXml stylesheet elements
- `SheetContext` — wraps WorksheetPart for cell access and column letter conversion
- `Property` / `Properties` — reflection-based property discovery with attribute support
- `ValueRenderer` — static class handling type-to-cell-value conversion (dates, enums, strings, complex types, whitespace trimming)
- Attributes: `ColumnAttribute`, `SplitAttribute`, `IgnoreAttribute` (in `Attributes/`)

### Test projects
- `Excelsior.Tests` — all tests including integration tests with snapshot verification
- `Excelsior.SourceGenerator.Tests` — source generator tests
- `StaticSettingsTests` — tests for global static settings (date formats, whitespace trimming, enum rendering)
- `SheetRender` — .NET Framework 4.8 utility using Excel Interop. Contains `RenderExcel` which opens all `.verified.xlsx` files, renders each sheet's used range to a bitmap, and saves as `_SheetName.png` alongside the xlsx. These pngs are referenced in the readme. Tests are `[Explicit]` (manual run only).
- `Model` — shared model classes used by tests

Tests use **NUnit** with **Verify** for snapshot testing. Verified snapshots are `.verified.png` and `.verified.txt` files alongside test classes.

## Documentation

The `readme.md` uses [MarkdownSnippets](https://github.com/SimonCropp/MarkdownSnippets) with `InPlaceOverwrite` convention (configured in `src/mdsnippets.json`). Code samples in the readme are pulled from test files via `#region SnippetName` / `#endregion` markers. The snippets are populated automatically during build (via `MarkdownSnippets.MsBuild`). When adding new code samples to the readme, use `<!-- snippet: SnippetName -->` / `<!-- endSnippet -->` references backed by region markers in test code rather than inline code blocks.

## Key Conventions

- Target framework: `net10.0` (with `LangVersion` set to `preview`)
- Central package management via `src/Directory.Packages.props`
- `TreatWarningsAsErrors` is enabled
- `ValueRenderer` static methods (e.g., `For<T>`, `ForEnums`, `DisableWhitespaceTrimming`) use `[ModuleInitializer]` for global configuration
