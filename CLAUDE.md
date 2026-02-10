# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Build Commands

```bash
dotnet restore              # Restore dependencies
dotnet build                # Build all projects
dotnet test                 # Run all tests
dotnet test --filter "FullyQualifiedName~UtilityTests"  # Run specific test class
```

Run the web app:
```bash
cd src/ConvertLearnToDoc && dotnet run
# Access at https://localhost:5001/
```

Run CLI tool:
```bash
dotnet run --project src/ConvertDocx -- <input> <output> [options]
```

## Architecture Overview

This is a .NET 8.0 solution for bidirectional conversion between Microsoft Word documents and Microsoft Learn modules/articles.

### Core Conversion Flow

**Word → Learn Module:**
1. `DocxToLearn.ConvertAsync()` orchestrates the conversion
2. `MarkdownRenderer` converts .docx → intermediate Markdown using registry of `IMarkdownObjectRenderer` implementations
3. `DocToMarkdownRenderer.PostProcessMarkdown()` cleans up the output
4. `ModuleBuilder` splits Markdown into YAML + unit files based on H1 headers

**Learn → Word:**
1. `LearnToDocx.ConvertFromRepoAsync()` or `ConvertFromFolderAsync()`
2. `ModuleCombiner` downloads/combines Learn module files
3. `MarkdownToDocConverter` uses Markdig to parse, then `DocxRenderer` with `IDocxObjectRenderer` registry renders to Word

### Renderer Pattern

Both conversion directions use a registry pattern:

- **Docx.Renderer.Markdown**: `IMarkdownObjectRenderer` implementations (ParagraphRenderer, TableRenderer, RunRenderer, IgnoredBlocks)
- **Markdig.Renderer.Docx**: `IDocxObjectRenderer` implementations in `Blocks/` and `Inlines/` directories

To add rendering for a new element type:
1. Implement the appropriate interface
2. Implement `CanRender(object element)` for type matching
3. Register in the renderer constructor

### Key Libraries

| Library | Purpose |
|---------|---------|
| `LearnDocUtils` | Core conversion logic, module building |
| `Docx.Renderer.Markdown` | Word → Markdown rendering |
| `Markdig.Renderer.Docx` | Markdown → Word rendering |
| `DocsToMarkdown` | Downloads Learn content from URLs |
| `ConvertLearnToDoc.Shared` | DTOs shared between web app and libraries |

### External Dependencies

- `Julmar.DxPlus` - Word document I/O
- `Markdig` - Markdown parsing
- `Microsoft.DocAsCode.MarkdigEngine.Extensions` - DocFX triple-colon syntax (`:::type`)
- `YamlDotNet` - YAML serialization

### Web App Structure

Blazor Server app with API controllers:
- `DocConverterController` (`/api/docconverter`) - Word ↔ Learn conversions
- `ContentConverterController` (`/api/contentconverter`) - URL-based conversions
- Pages: `DocToModule.razor`, `DocToArticle.razor`, `ContentToDoc.razor`, `ContentToMarkdown.razor`

Authentication via GitHub OAuth with email allowlist validation.

## Important Implementation Details

- Module units are split on H1 headers (`#`) in the Markdown
- Post-processing removes: `\r\n`, `\xA0` (NBSP), `\u200b` (ZWSP), `\u202f` (NNBSP)
- Markdown formatting options are stored in Word custom properties (UseAsterisksForBullets, UseAsterisksForEmphasis)
- YAML templates are embedded resources in `LearnDocUtils/templates/`
- Debug mode (`-d` flag or `options.Debug`) preserves intermediate files for troubleshooting
- Max upload size: 1GB (configurable in `ArticleOrModuleRef.cs`)
