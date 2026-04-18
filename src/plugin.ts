import * as ts from 'typescript/lib/tsserverlibrary';
import * as path from 'path';
import * as fs from 'fs';

/**
 * TypeScript Server Plugin that injects the ExcelScript ambient declarations
 * into TypeScript projects containing Office Scripts (.osts) files.
 *
 * The plugin is GATED: we only inject types when the project actually has
 * .osts files. This matters because types/excel-script.d.ts declares ambient
 * globals (notably `main` and the `ExcelScript` namespace), which would
 * pollute every unrelated TypeScript project if injected unconditionally.
 *
 * Injection strategy: we proxy `getScriptFileNames` on the language service
 * host to append our ambient .d.ts to the file list. This is the documented
 * TS plugin pattern and works for both configured projects and the inferred
 * project that tsserver creates for standalone .osts files. (The previous
 * approach of calling `project.addRoot(scriptInfo)` hits an internal
 * `Debug Failure. False expression.` assertion when the project is inferred.)
 *
 * Layout at runtime:
 *   <ext-root>/dist/plugin.js         (this file, after esbuild)
 *   <ext-root>/types/excel-script.d.ts (ambient declarations we inject)
 */
function init(mod: { typescript: typeof ts }) {
    const tsLib = mod.typescript;

    function create(info: ts.server.PluginCreateInfo) {
        const log = (msg: string) =>
            info.project.projectService.logger.info(`[office-scripts] ${msg}`);

        const declarationPath = path.resolve(__dirname, '..', 'types', 'excel-script.d.ts');
        const host = info.languageServiceHost;

        // Inject the ambient .d.ts by proxying getScriptFileNames on the
        // LanguageServiceHost. This is the supported pattern for adding
        // ambient declarations from a tsserver plugin — it avoids
        // project.addRoot(), which asserts the ScriptInfo is client-opened
        // on InferredProjects and throws "Debug Failure. False expression.".
        const originalGetScriptFileNames = host.getScriptFileNames.bind(host);
        host.getScriptFileNames = () => {
            const files = originalGetScriptFileNames();
            if (!projectHasOsts(info.project)) return files;
            return files.includes(declarationPath) ? files : [...files, declarationPath];
        };

        // Make each .osts file a module in tsserver's in-memory view by
        // appending a trailing `export {};`. Office Scripts are authored as
        // script-style files (top-level `function main`), so across multiple
        // .osts files the `main` symbols collide as TS2393 "Duplicate
        // function implementation". A client can ship many .osts files in
        // one folder — we want them all to typecheck independently.
        //
        // The appended text is invisible to the user (file on disk is
        // untouched) and appended AFTER the original content so line/column
        // positions for diagnostics, go-to-def, etc. stay correct. The
        // Office Scripts runtime reads the disk file, so runtime is
        // unaffected.
        const originalGetScriptSnapshot = host.getScriptSnapshot!.bind(host);
        host.getScriptSnapshot = (fileName: string) => {
            const snap = originalGetScriptSnapshot(fileName);
            if (!snap) return snap;
            if (!fileName.toLowerCase().endsWith('.osts')) return snap;
            const text = snap.getText(0, snap.getLength());
            return tsLib.ScriptSnapshot.fromString(text + '\nexport {};\n');
        };

        // Teach tsserver how to resolve imports that land on an .osts file.
        // Default module resolution only considers .ts/.tsx/.d.ts/.js/.jsx,
        // so `import { foo } from './helper'` where helper.osts exists would
        // otherwise report TS2307 "Cannot find module". We fall through to
        // the host's default resolver first, then check for a sibling .osts
        // when nothing else matched. The resolved file is classified as
        // Extension.Ts because tsserver has no "Extension.Osts" enum — the
        // classification is only used for diagnostics routing; the actual
        // parse uses our getScriptSnapshot proxy, which already wraps .osts
        // files as modules.
        const originalResolveModuleNameLiterals = host.resolveModuleNameLiterals?.bind(host);
        host.resolveModuleNameLiterals = (
            moduleLiterals,
            containingFile,
            redirectedReference,
            options,
            containingSourceFile,
            reusedNames,
        ) => {
            const defaults = originalResolveModuleNameLiterals
                ? originalResolveModuleNameLiterals(
                      moduleLiterals,
                      containingFile,
                      redirectedReference,
                      options,
                      containingSourceFile,
                      reusedNames,
                  )
                : moduleLiterals.map(() => ({ resolvedModule: undefined }));

            return defaults.map((res, i) => {
                if (res.resolvedModule) return res;
                const spec = moduleLiterals[i].text;
                if (!spec.startsWith('.')) return res;
                const containingDir = path.dirname(containingFile);
                const resolved = resolveOstsPath(containingDir, spec);
                if (!resolved) return res;
                return {
                    ...res,
                    resolvedModule: {
                        resolvedFileName: resolved,
                        extension: tsLib.Extension.Ts,
                        isExternalLibraryImport: false,
                    },
                };
            });
        };

        // Suppress "possibly null/undefined" diagnostics in .osts files to
        // match Microsoft's in-Excel Office Scripts editor behavior, which
        // doesn't surface these as errors. We keep the types accurate
        // (getColumnByName still returns TableColumn | undefined) so users
        // who want strict null handling still get correct inference, but
        // the red squigglies are removed for the common case. Filtering
        // only the diagnostic codes below — not disabling strictNullChecks
        // wholesale — avoids changing type inference for helper .ts files
        // in the same project.
        const originalGetSemanticDiagnostics =
            info.languageService.getSemanticDiagnostics.bind(info.languageService);
        info.languageService.getSemanticDiagnostics = (fileName: string) => {
            const diags = originalGetSemanticDiagnostics(fileName);
            if (!fileName.toLowerCase().endsWith('.osts')) return diags;
            return diags.filter(d => !SUPPRESSED_CODES.has(d.code));
        };

        log(`Plugin initialized for ${info.project.getProjectName()}, injecting ${declarationPath}`);
        return info.languageService;
    }

    return { create };
}

// TS codes for diagnostics Microsoft's Office Scripts editor doesn't
// surface; we filter them for .osts files to match. Keep this list
// narrow — only add codes that are both (a) common in Office Scripts
// idioms and (b) confirmed not raised by the in-Excel editor.
const SUPPRESSED_CODES = new Set<number>([
    2531, // Object is possibly 'null'.
    2532, // Object is possibly 'undefined'.
    2533, // Object is possibly 'null' or 'undefined'.
    18047, // 'X' is possibly 'null'.
    18048, // 'X' is possibly 'undefined'.
    18049, // 'X' is possibly 'null' or 'undefined'.
    7053, // Element implicitly has 'any' type because expression of type 'X' can't be used to index type 'Y'.
]);

function projectHasOsts(project: ts.server.Project): boolean {
    // Fast path: any already-known file with the .osts extension.
    const files = project.getFileNames(/* excludeFilesFromExternalLibraries */ true);
    return files.some(f => f.toLowerCase().endsWith('.osts'));
}

function resolveOstsPath(fromDir: string, specifier: string): string | undefined {
    const base = path.resolve(fromDir, specifier);
    const candidates = [base + '.osts', path.join(base, 'index.osts')];
    return candidates.find(p => fs.existsSync(p));
}

export = init;
