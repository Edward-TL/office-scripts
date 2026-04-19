import * as ts from 'typescript/lib/tsserverlibrary';
import * as path from 'path';
import * as fs from 'fs';

/**
 * TypeScript Server Plugin that applies Office Scripts treatment to:
 *   1. All `.osts` files (language id `office-script`).
 *   2. All `.ts` files containing a `/** @OfficeScript *\/` JSDoc tag.
 *
 * Treatment per qualifying file:
 *   - Snapshot is wrapped with a trailing `export {};` so each file becomes
 *     its own module. This prevents TS2393 "Duplicate function implementation"
 *     when multiple Office Scripts with `function main` live in one folder.
 *   - `getSemanticDiagnostics` filters the codes Microsoft's in-Excel editor
 *     demonstrably doesn't raise (possibly-null/undefined, implicit-any-index).
 *
 * The plugin is project-gated: types/excel-script.d.ts and the snapshot
 * proxies only activate for projects that contain at least one Office
 * Scripts file. Regular TypeScript projects pay the cost of the plugin
 * being loaded but see no behavioral change.
 */
const WRAP_SUFFIX = '\nexport {};\n';
const WRAPPED_HOSTS = new WeakSet<object>();

interface PluginConfig {
    strictDiagnostics?: boolean;
}

function init(mod: { typescript: typeof ts }) {
    const tsLib = mod.typescript;

    function create(info: ts.server.PluginCreateInfo) {
        const log = (msg: string) =>
            info.project.projectService.logger.info(`[office-scripts] ${msg}`);

        const state: { strictDiagnostics: boolean } = {
            strictDiagnostics: Boolean((info.config as PluginConfig | undefined)?.strictDiagnostics),
        };
        LIVE_STATES.add(state);

        const host = info.languageServiceHost;

        // Guard against double-registration: if the same host ever makes it
        // here twice (e.g. plugin gets loaded via multiple language ids or
        // re-init), stack only one layer of proxies. Otherwise snapshots
        // accumulate `export {};` suffixes and semantic diagnostics get
        // filtered twice — both observable as duplicated behavior.
        if (WRAPPED_HOSTS.has(host)) {
            log(`Skipping duplicate init for ${info.project.getProjectName()}`);
            return info.languageService;
        }
        WRAPPED_HOSTS.add(host);

        const declarationPath = path.resolve(__dirname, '..', 'types', 'excel-script.d.ts');

        // Inject the ambient .d.ts when the project contains any Office
        // Script. The presence check is cached per project and invalidated
        // when the project's file count changes.
        const originalGetScriptFileNames = host.getScriptFileNames.bind(host);
        host.getScriptFileNames = () => {
            const files = originalGetScriptFileNames();
            if (!projectHasOfficeScripts(info.project, files)) return files;
            return files.includes(declarationPath) ? files : [...files, declarationPath];
        };

        // Module-wrap qualifying files so top-level `main` symbols don't
        // collide across files in the same folder. Untagged .ts files are
        // passed through untouched.
        const originalGetScriptSnapshot = host.getScriptSnapshot!.bind(host);
        host.getScriptSnapshot = (fileName: string) => {
            const snap = originalGetScriptSnapshot(fileName);
            if (!snap) return snap;
            if (!isOfficeScriptPath(fileName)) return snap;
            const text = snap.getText(0, snap.getLength());
            if (!isOfficeScriptContent(fileName, text)) return snap;
            if (text.endsWith(WRAP_SUFFIX)) return snap;
            return tsLib.ScriptSnapshot.fromString(text + WRAP_SUFFIX);
        };

        // Resolve relative imports that land on an `.osts` file, since
        // default module resolution doesn't consider that extension.
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

        // Suppress "possibly null/undefined" and index-any diagnostics to
        // match Microsoft's in-Excel editor behavior. Applied only to
        // qualifying files — untagged .ts files keep strict checking.
        const originalGetSemanticDiagnostics =
            info.languageService.getSemanticDiagnostics.bind(info.languageService);
        info.languageService.getSemanticDiagnostics = (fileName: string) => {
            const diags = originalGetSemanticDiagnostics(fileName);
            if (state.strictDiagnostics) return diags;
            if (!isOfficeScriptPath(fileName)) return diags;
            const snap = originalGetScriptSnapshot(fileName);
            const text = snap ? snap.getText(0, snap.getLength()) : '';
            if (!isOfficeScriptContent(fileName, text)) return diags;
            return diags.filter(d => !SUPPRESSED_CODES.has(d.code));
        };

        log(`Plugin initialized for ${info.project.getProjectName()}, injecting ${declarationPath}`);
        return info.languageService;
    }

    // Called when the editor pushes new plugin configuration via
    // `typescript.tsserver.configurePlugin`. We route strictDiagnostics
    // changes in without needing a TS Server restart.
    function onConfigurationChanged(config: PluginConfig) {
        // No per-instance state here — the project's `create` closure owns
        // its own `state`. In practice there's one project closure alive;
        // we reach it via the WeakMap below so config flips re-take effect.
        for (const state of LIVE_STATES) {
            state.strictDiagnostics = Boolean(config.strictDiagnostics);
        }
    }

    return { create, onConfigurationChanged };
}

const LIVE_STATES = new Set<{ strictDiagnostics: boolean }>();

const SUPPRESSED_CODES = new Set<number>([
    2531, 2532, 2533,
    18047, 18048, 18049,
    7053,
]);

const MARKER_RE = /\/\*\*[\s\S]*?@OfficeScript\b/;

/** Cheap extension check. Says whether a file is *eligible* for Office
 *  Scripts treatment; content still has to be verified for `.ts`. */
function isOfficeScriptPath(fileName: string): boolean {
    const lower = fileName.toLowerCase();
    if (lower.endsWith('.osts')) return true;
    return lower.endsWith('.ts') && !lower.endsWith('.d.ts');
}

/** Final verdict combining extension + content. `.osts` always qualifies;
 *  `.ts` only if the text contains the `@OfficeScript` marker. */
function isOfficeScriptContent(fileName: string, text: string): boolean {
    if (fileName.toLowerCase().endsWith('.osts')) return true;
    return MARKER_RE.test(text);
}

/**
 * Keyed by file count: if files are added or removed, the cache is
 * recomputed. For changes WITHIN existing files (adding/removing the
 * marker), users can run "Restart TS Server" to refresh.
 */
const projectHasOSCache = new WeakMap<ts.server.Project, { fileCount: number; hit: boolean }>();

function projectHasOfficeScripts(project: ts.server.Project, files: readonly string[]): boolean {
    // Fast path: any `.osts` file is an instant yes. No file reads needed.
    if (files.some(f => f.toLowerCase().endsWith('.osts'))) return true;

    const cached = projectHasOSCache.get(project);
    if (cached && cached.fileCount === files.length) return cached.hit;

    let hit = false;
    for (const f of files) {
        const lower = f.toLowerCase();
        if (!lower.endsWith('.ts') || lower.endsWith('.d.ts')) continue;
        if (f.includes('node_modules')) continue;
        try {
            if (MARKER_RE.test(fs.readFileSync(f, 'utf8'))) {
                hit = true;
                break;
            }
        } catch {
            // ignore unreadable file
        }
    }
    projectHasOSCache.set(project, { fileCount: files.length, hit });
    return hit;
}

function resolveOstsPath(fromDir: string, specifier: string): string | undefined {
    const base = path.resolve(fromDir, specifier);
    const candidates = [base + '.osts', path.join(base, 'index.osts')];
    return candidates.find(p => fs.existsSync(p));
}

export = init;
