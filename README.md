# Office Scripts Support

Professional language support for [Office Scripts](https://learn.microsoft.com/office/dev/scripts/) — the TypeScript-based automation runtime for Excel on the web. Write, lint, and get full IntelliSense in VS Code for scripts living either in `.osts` files or in plain `.ts` files tagged with `/** @OfficeScript */`.

- **Version:** 1.3.0
- **Author:** Edward-TL
- **License:** MIT

## What it does

This extension treats Office Scripts as first-class citizens in VS Code:

- **Works on both `.osts` and `.ts` files.** You write standard TypeScript in either extension — the file type is a packaging decision, not a language one. See [.osts vs .ts](#osts-vs-ts--which-should-i-use) below.
- **`/** @OfficeScript *\/` marker** — tag the `main` function in any `.ts` file with this JSDoc hint and the extension gives it the full Office Scripts treatment (type injection, module isolation, Excel-matching diagnostics). `.osts` files are recognized automatically with no tag required.
- **ExcelScript type resolution** — `workbook.`, `sheet.getRange()`, `table.addRow()`, and the rest of the `ExcelScript` namespace autocomplete and type-check as if you were writing them inside the Excel editor. Types are injected automatically by a TypeScript Server Plugin; no `import` or `/// <reference>` lines needed.
- **Multi-script projects.** Each qualifying file is treated as an isolated module in tsserver's in-memory view, so you can keep dozens of scripts for one client in the same folder without `main`-function collisions or other top-level-symbol conflicts. The file on disk is untouched — only the language service sees the module boundary.
- **In-Excel-editor-matching diagnostics.** Errors that Microsoft's in-Excel Office Scripts editor doesn't surface are suppressed to match that experience:
  - "Object is possibly `null` / `undefined`" (TS2531/2532/2533/18047/18048/18049)
  - "Element implicitly has an `any` type because expression of type … can't be used to index type …" (TS7053)
  Regular (non-Office-Script) `.ts` files in the same project keep full strictness. Can be re-enabled per user via [`officeScripts.strictDiagnostics`](#strict-diagnostics-setting).
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
- **TypeScript Command Palette access.** With an Office Script focused, `Cmd+Shift+P` surfaces the TypeScript commands that VS Code normally hides for non-`typescript` language ids:
  - Office Scripts: Restart TS Server
  - Office Scripts: Reload Projects
  - Office Scripts: Select TypeScript Version
  - Office Scripts: Open TS Server Log
  - Office Scripts: Go to Project Configuration
  - Office Scripts: Go to Source Definition
- **Custom file icon** — `.osts` files get a distinct Office-orange icon in the explorer, independent of your active icon theme.
- **Harvest Core Library** — point the command at a folder of downloaded scripts and it builds a reusable helper library. See [Harvest Core Library](#harvest-core-library) below.
- **Export to OSTS (JSON for upload)** — takes a `.ts` file tagged with `/** @OfficeScript */`, inlines its imports, and writes a sibling `.osts` JSON envelope ready to drop into OneDrive → Scripts for upload.

## Installation

### From a local `.vsix` (current distribution method)

```bash
git clone <this-repo>
cd office-scripts
npm install
npm run package
vsce package
code --install-extension office-scripts-*.vsix
```

Fully quit and relaunch VS Code after install so the TypeScript server restarts with the plugin loaded.

### Prerequisites

- VS Code `≥ 1.75`
- Node.js `≥ 18` (only if building from source)

## Usage

Office Scripts in this extension are any file the extension recognizes as one: **every `.osts` file**, plus **any `.ts` file tagged with `/** @OfficeScript *\/`**. Both are plain TypeScript under the hood — the `ExcelScript` namespace is available globally, no setup per-project required.

### `.osts` vs `.ts` — which should I use?

Both file types accept the same TypeScript code. Pick by workflow:

| Scenario                                              | Use             |
| ----------------------------------------------------- | --------------- |
| File downloaded from OneDrive → Scripts (Power Automate format, JSON envelope) | `.osts` |
| Ready-to-upload artifact produced by **Export to OSTS** | `.osts`       |
| Day-to-day authoring, especially multi-file projects with shared helpers | `.ts` with `/** @OfficeScript */` |
| Files that need to live alongside a normal TypeScript codebase | `.ts` with `/** @OfficeScript */` |

The practical difference: `.osts` is the "on-disk upload format" — it carries the JSON envelope Excel Online / Power Automate expects. `.ts` is the "authoring format" — easier to diff, easier to grep, integrates cleanly with git and with tsserver. The extension recognizes both, and commands like **Inline Imports**, **Split Flows**, and **Export to OSTS** bridge between them.

### The `/** @OfficeScript */` marker

A JSDoc tag placed on the line immediately above the `main` function of any `.ts` file. It's an authoring-time hint that tells this extension "treat me like an Office Script" — so the file gets ambient `ExcelScript` types, module isolation, and the same relaxed diagnostics `.osts` files get. The marker is stripped automatically by **Export to OSTS**, so it never ships inside the uploaded script.

```typescript
/**
 * Rebuilds the month-end plaza summary tables from the raw sales data.
 */
/** @OfficeScript */
function main(workbook: ExcelScript.Workbook): void {
    // ...
}
```

Any existing JSDoc describing `main` can stay above the marker; the extension only cares that the tag exists somewhere as a `/** @OfficeScript */` block in the file. Snippet `osofficescript` expands to the marker.

Without the marker, a plain `.ts` file is treated by tsserver as ordinary TypeScript — strict null checks, no `ExcelScript` namespace, no module isolation. That's intentional: it keeps the extension invisible in regular TypeScript codebases.

### Strict diagnostics setting

By default the extension suppresses the strict-null / possibly-undefined / index-any codes that Microsoft's in-Excel editor doesn't surface, to match that runtime's behavior. If you want VS Code to be *more* rigorous than the Excel editor (catching every `Object is possibly 'undefined'` in your authoring view), opt back in via VS Code settings:

```json
"officeScripts.strictDiagnostics": true
```

Takes effect live — no TS Server restart needed. The suppressed codes that get re-enabled: `2531`, `2532`, `2533`, `18047`, `18048`, `18049`, `7053`.

### Type resolution

```typescript
function main(workbook: ExcelScript.Workbook) {
    const sheet = workbook.getActiveWorksheet();   // autocompletes from ExcelScript.Worksheet
    const range = sheet.getRange("A1:B10");        // autocompletes A1-notation options
    range.setValues([[1, 2], [3, 4]]);             // type-checks the 2D array shape
}
```

### Sharing code across scripts

Excel's runtime runs one `.osts` per invocation — it doesn't support `import`. To share helper functions across multiple scripts during authoring:

```
client-acme/
├── shared/
│   └── tableUtils.ts         ← helpers live here
├── sales-report.osts
└── inventory-update.osts
```

`shared/tableUtils.ts`:

```typescript
export function getColumnIndex(table: ExcelScript.Table, name: string): number {
    const col = table.getColumnByName(name);
    if (!col) throw new Error(`Column "${name}" not found`);
    return col.getIndex();
}
```

`sales-report.ts` (authoring view — note the `/** @OfficeScript */` tag so the extension treats this `.ts` file as a script):

```typescript
import { getColumnIndex } from './shared/tableUtils';

/** @OfficeScript */
function main(workbook: ExcelScript.Workbook) {
    const sales = workbook.getTable('Sales')!;
    const plazaIdx = getColumnIndex(sales, 'PLAZA');
    // ...
}
```

When ready to deploy: run **Office Scripts: Export to OSTS (JSON for upload)** to produce `sales-report.osts` with imports inlined and the JSON envelope already wrapped around the body. Or, for a quick preview, run **Office Scripts: Inline Imports for Excel Upload** to get an in-memory buffer you can paste into the Excel editor directly.

### Snippets

| Prefix       | Expands to                                                                |
| ------------ | ------------------------------------------------------------------------- |
| `osmain`     | Full `main(workbook: ExcelScript.Workbook)` function skeleton             |
| `ossheet`    | `const sheet = workbook.getActiveWorksheet();`                            |
| `ostable`    | `workbook.getTable(...)` with null-guard                                  |
| `osforrows`  | `for` loop over `table.getRange().getValues()`                            |
| `osrange`    | Range + `setValue` + font/format chain                                    |
| `osaddrow`   | `table.addRow(undefined, [...])`                                          |
| `osrangeidx` | `sheet.getRangeByIndexes(row, col, height, width)`                        |
| `osflowsplit`| `/** @FlowSplit */` marker for the Split Flows command                    |
| `osofficescript` | `/** @OfficeScript */` marker that tags a `.ts` file's `main` function |

All prefixes start with `os` so they don't clutter autocomplete in unrelated TypeScript projects.

### Quick-fixes

Place the cursor on a diagnostic and press `Cmd+.` (macOS) or `Ctrl+.` (Windows/Linux):

| Diagnostic                            | Quick-fix                          |
| ------------------------------------- | ---------------------------------- |
| `"any" type is forbidden`             | Replace with `unknown`             |
| `console.warn is not supported`       | Replace with `console.log`         |
| `console.error is not supported`      | Replace with `console.log`         |

## Harvest Core Library

A command for consolidating scattered Office Scripts into a reusable helper library. Useful when you have a folder full of scripts downloaded from OneDrive (typically via a Power Automate flow targeting `~/Documents/Scripts`) and want to extract the helper functions that keep getting copy-pasted across them.

### Running it

Command palette → **Office Scripts: Harvest Core Library**. You'll be asked for two folders:

1. **Source folder** — where your downloaded `.osts` files live. They can be either raw TypeScript or the JSON envelope format Power Automate produces; both are handled.
2. **Destination folder** — where the harvested library gets written.

### Output layout

```
<destination>/
├── script-osts-version/          Plain .ts copies of every source script,
│   ├── sales-report.ts           JSON envelopes stripped, `/** @OfficeScript */`
│   ├── inventory-update.ts       injected above each `main`.
│   └── ...
├── core/
│   ├── getColumnIndex.ts         One file per helper function that is either
│   ├── normalizeStoreName.ts     unique across the harvest, OR whose duplicates
│   └── ...                       had byte-identical bodies (whitespace-normalized).
├── conflict/
│   ├── formatHeader/
│   │   ├── sales-report.ts       Same-name functions whose bodies DIFFER —
│   │   └── inventory-update.ts   one file per distinct variant, named after
│   └── ...                       the source script, so you can diff & reconcile.
├── interface/
│   ├── core/
│   │   └── ReportRow.ts          Same rules, applied to `interface` declarations.
│   └── conflict/
│       └── StoreRow/
│           ├── sales-report.ts
│           └── inventory-update.ts
```

### The two folders explained

- **`core/`** — your working library. Each file is a standalone `export function` (or `export interface`) you can import from future scripts with a normal `import { getColumnIndex } from '../core/getColumnIndex'` and have the extension inline it automatically via **Inline Imports for Excel Upload** before you paste the script into Excel.
- **`conflict/`** — anything that needs human judgment. If three scripts all defined `formatHeader` but one used a different date format, they land here side-by-side so you can decide which variant becomes the canonical one to promote into `core/`.
- **`interface/`** — same split applied to TypeScript `interface` declarations, because shapes drift too (`StoreRow` with a `store` field here, with `storeName` there, etc.).
- **`script-osts-version/`** — a throwaway intermediate: plain-TS copies of your source scripts with the JSON envelope unwrapped and `/** @OfficeScript */` injected above each `main`. Useful as a grep-friendly backup of what the harvest saw.

### What qualifies as a helper

Top-level function declarations and `interface` declarations — not methods, not nested helpers, not arrow-function consts. The `main` function is excluded by definition.

## Architecture

The extension has two halves:

1. **VS Code-side** ([src/extension.ts](src/extension.ts)) — registers diagnostics, code actions, completion provider, and hover provider against the `office-script` language selector. Force-activates `vscode.typescript-language-features` on startup so tsserver runs for our custom language id. Also registers the `officeScripts.*` Command Palette proxies that forward to the built-in `typescript.*` commands so TypeScript tooling is reachable with an `.osts` file focused.

2. **TypeScript-server-side** ([src/plugin.ts](src/plugin.ts)) — a tsserver plugin that:
   - Injects [types/excel-script.d.ts](types/excel-script.d.ts) as a root file (via a `getScriptFileNames` proxy on the `LanguageServiceHost`) in any project that contains at least one `.osts` file. Project-level gating prevents the `ExcelScript` namespace from polluting unrelated TypeScript projects.
   - Wraps each `.osts` file's snapshot with a trailing `export {};` so tsserver treats it as a module. This isolates top-level declarations (`function main`, helper `const`s, etc.) per file, letting one folder hold many standalone scripts.
   - Proxies `getSemanticDiagnostics` to drop the diagnostic codes that Microsoft's in-Excel editor doesn't raise.

The plugin is bundled to `dist/plugin.js` by esbuild. A tiny shim in [plugin-package/](plugin-package/) is published as a `file:` dependency so tsserver can resolve `require('office-scripts-plugin')` from the extension's `node_modules`.

```
office-scripts/
├── src/
│   ├── extension.ts              VS Code activation + provider wiring + command proxies
│   ├── plugin.ts                 TS Server Plugin (type injection, module wrapping, diagnostic filtering)
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
npm run watch:esbuild  # Bundles src/ → dist/ on change
npm run watch:tsc      # Live type-checks the whole project (tsc --watch)
npm run check-types    # One-shot tsc --noEmit
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
- **Excel upload is manual.** Excel Online's runtime is single-file and doesn't support `import`. Helpers authored in shared `.ts` files must be inlined into the `.osts` before pasting into the in-Excel editor.

## Release notes

### 1.3.0

- **Extension now also works on `.ts` files**, not only `.osts`. Tag the `main` function of a `.ts` file with `/** @OfficeScript */` and it gets the full Office Scripts treatment (ambient `ExcelScript` types, module isolation, Excel-matching diagnostics). `.osts` continues to work with no tag required. New snippet `osofficescript` expands to the marker.
- New command: **Office Scripts: Export to OSTS (JSON for upload)** — takes a `.ts` file tagged with `/** @OfficeScript */`, inlines its relative imports, strips the marker, wraps the body in the Power Automate / OneDrive → Scripts JSON envelope, and writes `<name>.osts` next to the source. One click from authoring to an artifact you can drop straight into OneDrive.
- New command: **Office Scripts: Harvest Core Library** — point it at a folder of downloaded scripts and it extracts every non-`main` function and top-level `interface` into `core/` (unique or byte-identical duplicates) and `conflict/` (same name, different body), with a separate `interface/` tree for shape declarations. JSON envelopes are unwrapped on the fly into `script-osts-version/*.ts`, each tagged with `/** @OfficeScript */` above `main`. See [Harvest Core Library](#harvest-core-library).
- New setting: `officeScripts.strictDiagnostics` (default `false`). Set to `true` to re-enable TypeScript's strict-null / possibly-undefined / implicit-any-index codes on Office Scripts — useful if you want VS Code to flag issues the in-Excel editor lets through. Changes apply live; no TS Server restart needed.
- All commands that previously required an `.osts` file (Inline Imports, Split Flows) now also accept a `.ts` file tagged with `/** @OfficeScript */`. Split Flows output matches the source extension so a `.ts`-based flow produces `.ts` splits (each carrying the marker).

### 1.2.0

- Multi-script projects: each `.osts` file is treated as an isolated module so multiple `main` functions and other top-level declarations coexist in one folder.
- In-Excel-editor-matching diagnostics: strict-null-check and implicit-any-on-index errors are suppressed for `.osts` files to mirror the Microsoft online editor.
- TypeScript commands (Restart TS Server, Reload Projects, Select Version, Open TS Server Log, Go to Project Configuration, Go to Source Definition) now surface in the Command Palette when an `.osts` file is focused, under the `Office Scripts` category.
- New command: **Office Scripts: Inline Imports for Excel Upload** resolves relative imports (including chained `.osts`/`.ts` helpers) and inlines their bodies into a new editor ready to paste into Excel.
- New command: **Office Scripts: Split Flows (@FlowSplit)** — functions annotated with `/** @FlowSplit */` are split into their own `.osts` files under a sibling folder. Each split file carries only the helpers it uses, with imports resolved and inlined automatically.
- `osflowsplit` snippet expands to `/** @FlowSplit */`.
- Type injection reworked to use a `getScriptFileNames` proxy on the `LanguageServiceHost`, avoiding the `addRoot` assertion failure on inferred projects.
- Multi-strategy activation of the TypeScript feature extension so the plugin loads even on forks/Nightly builds.
- Output channel "Office Scripts" surfaces plugin activation logs for troubleshooting.

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

MIT © Edward TL
