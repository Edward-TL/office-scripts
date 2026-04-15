# Office Scripts Support

Professional language support for [Office Scripts](https://learn.microsoft.com/office/dev/scripts/) — the TypeScript-based automation runtime for Excel on the web. Write, lint, and get full IntelliSense for `.osts` files in VS Code without having to cut-and-paste into the Excel Online code editor.

## What it does

This extension treats `.osts` files as first-class citizens in VS Code:

- **ExcelScript type resolution** — `workbook.`, `sheet.getRange()`, `table.addRow()`, and the rest of the `ExcelScript` namespace autocomplete and type-check as if you were writing them inside the Excel editor. Types are injected automatically by a TypeScript Server Plugin; no `import` or `/// <reference>` lines needed.
- **Strict linting rules** specific to Office Scripts:
  - `any` type is forbidden — use `unknown` or a concrete type.
  - `console.warn` / `console.error` are flagged — only `console.log` is supported by the Office Scripts runtime.
- **Quick-fixes** (Cmd+. / Ctrl+.) for each lint rule — one click to replace `any` with `unknown` or rewrite `console.warn` as `console.log`.
- **Snippets** for common patterns — type `osmain` to expand the full `main(workbook: ExcelScript.Workbook)` skeleton.
- **Context-aware completions** inside string-literal arguments where the type system alone can't help:
  - A1 notation inside `sheet.getRange("…")` (`A1`, `A1:B2`, `A:A`, `1:1`, …).
  - Hex colors inside `.setColor("…")` with human-readable names, including Excel's brand green.
  - Alignment enum values inside `.setHorizontalAlignment("…")`.
- **Hover docs** — hovering any `ExcelScript.*` type shows a link to the Microsoft Learn reference page.
- **Custom file icon** — `.osts` files get a distinct Office-orange icon in the explorer, independent of your active icon theme.

## Installation

### From a local `.vsix` (current distribution method)

```bash
git clone <this-repo>
cd office-scripts
npm install
npm run package
vsce package --allow-missing-repository
code --install-extension office-scripts-1.1.1.vsix
```

Fully quit and relaunch VS Code after install so the TypeScript server restarts with the plugin loaded.

### Prerequisites

- VS Code `≥ 1.75`
- Node.js `≥ 18` (only if building from source)

## Usage

Any file with the `.osts` extension is treated as an Office Script. The `ExcelScript` namespace is available globally — no setup per-project required.

### Type resolution

```typescript
function main(workbook: ExcelScript.Workbook) {
    const sheet = workbook.getActiveWorksheet();   // autocompletes from ExcelScript.Worksheet
    const range = sheet.getRange("A1:B10");        // autocompletes A1-notation options
    range.setValues([[1, 2], [3, 4]]);             // type-checks the 2D array shape
}
```

### Snippets

| Prefix       | Expands to                                                                |
| ------------ | ------------------------------------------------------------------------- |
| `osmain`     | Full `main(workbook: ExcelScript.Workbook)` function skeleton              |
| `ossheet`    | `const sheet = workbook.getActiveWorksheet();`                            |
| `ostable`    | `workbook.getTable(...)` with null-guard                                  |
| `osforrows`  | `for` loop over `table.getRange().getValues()`                            |
| `osrange`    | Range + `setValue` + font/format chain                                    |
| `osaddrow`   | `table.addRow(undefined, [...])`                                          |
| `osrangeidx` | `sheet.getRangeByIndexes(row, col, height, width)`                        |

All prefixes start with `os` so they don't clutter autocomplete in unrelated TypeScript projects.

### Quick-fixes

Place the cursor on a diagnostic and press `Cmd+.` (macOS) or `Ctrl+.` (Windows/Linux):

| Diagnostic                            | Quick-fix                          |
| ------------------------------------- | ---------------------------------- |
| `"any" type is forbidden`             | Replace with `unknown`             |
| `console.warn is not supported`       | Replace with `console.log`         |
| `console.error is not supported`      | Replace with `console.log`         |

## Architecture

The extension has two halves:

1. **VS Code-side** ([src/extension.ts](src/extension.ts)) — registers diagnostics, code actions, completion provider, and hover provider against the `office-script` language selector. Force-activates `vscode.typescript-language-features` on startup so tsserver runs for our custom language id.

2. **TypeScript-server-side** ([src/plugin.ts](src/plugin.ts)) — a tsserver plugin that injects [types/excel-script.d.ts](types/excel-script.d.ts) as an ambient root file in any project that contains at least one `.osts` file. The project-level gating prevents the `ExcelScript` namespace and the ambient `main` function from polluting unrelated TypeScript projects.

The plugin is bundled to `dist/plugin.js` by esbuild. A tiny shim in [plugin-package/](plugin-package/) is published as a `file:` dependency so tsserver can resolve `require('office-scripts-plugin')` from the extension's `node_modules`.

```
office-scripts/
├── src/
│   ├── extension.ts              VS Code activation + provider wiring
│   ├── plugin.ts                 TS Server Plugin (ambient type injection)
│   ├── diagnostics.ts            Custom lint rules
│   ├── codeActions.ts            Quick-fix provider
│   ├── completionProvider.ts     Context-aware string-literal completions
│   ├── hoverProvider.ts          Microsoft Learn docs links
│   └── test/
├── types/
│   └── excel-script.d.ts         Ambient ExcelScript namespace declarations
├── snippets/
│   └── office-scripts.json       Snippet definitions
├── syntaxes/
│   └── office-script.tmLanguage.json   Delegates to TypeScript grammar
├── icons/
│   └── osts.svg                  File icon shown in the explorer
├── plugin-package/               file: dependency shim for the TS plugin
├── language-configuration.json   Brackets, auto-close, comments
├── tsconfig.osts.json            Template tsconfig for user projects
├── package.json                  Extension manifest
└── esbuild.js                    Bundles src/ → dist/
```

## Development

```bash
npm install            # Installs deps and wires up the plugin shim
npm run watch          # esbuild watch mode for src/ → dist/
npm run check-types    # tsc --noEmit
npm run lint           # eslint src
```

Press `F5` in VS Code to launch the Extension Development Host with the extension loaded. Open [src/test-usage.osts](src/test-usage.osts) to exercise every feature.

### Running tests

The mocha suite in `src/test/` uses `@vscode/test-electron`. Run from the command palette via the *Extension Tests* launch configuration, or:

```bash
npm run compile
# then run the VS Code test CLI via .vscode-test.mjs
```

## Known limitations

- **Partial type coverage.** [types/excel-script.d.ts](types/excel-script.d.ts) currently declares only a subset of the `ExcelScript` namespace. To get full fidelity, replace it with the authoritative declarations extracted from the Excel Online Monaco editor (DevTools Console: `monaco.languages.typescript.typescriptDefaults.getExtraLibs()`).
- **No table-name resolution.** `workbook.getTable("…")` shows placeholder table names in completion; we can't know the real ones without connecting to a live workbook.
- **First-open latency.** The first `.osts` file opened after a VS Code restart takes a moment while the TypeScript extension force-activates. Subsequent files are instant.

## Release notes

### 1.1.1

- New Office-orange file icon, TypeScript-logo-style silhouette.

### 1.1.0

- Split `.osts` into its own `office-script` language id so the file icon doesn't collide with regular `.ts` files.
- Grammar inherits from TypeScript (`source.ts`); full syntax highlighting preserved.
- Added language-configuration (brackets, auto-close, region folding).
- Added context-aware completions (A1 notation, colors, alignment, table hints).
- Added hover provider with Microsoft Learn documentation links.

### 1.0.0

- Initial release: TS Server Plugin injecting `ExcelScript` types, custom diagnostics for `any` and `console.warn`/`console.error`, quick-fix code actions, snippet contributions.

## License

TBD.
