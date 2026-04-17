import * as ts from 'typescript/lib/tsserverlibrary';
import * as path from 'path';

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
function init({ typescript }: { typescript: typeof ts }) {
    function create(info: ts.server.PluginCreateInfo) {
        const log = (msg: string) =>
            info.project.projectService.logger.info(`[office-scripts] ${msg}`);

        try {
            log(`Plugin initialized for project: ${info.project.getProjectName()}`);
            log(`Project root: ${info.project.getCurrentDirectory()}`);
            log(`__dirname: ${__dirname}`);

            const declarationPath = path.resolve(__dirname, '..', 'types', 'excel-script.d.ts');
            log(`Resolved declaration path: ${declarationPath}`);

            // Proxy getScriptFileNames on the language service host so our
            // ambient declaration is always presented to the compiler whenever
            // the current project contains at least one .osts file. tsserver
            // re-queries this on every refresh, so late-opened .osts files are
            // picked up automatically without any explicit re-injection hook.
            const host = info.languageServiceHost;
            const originalGetScriptFileNames = host.getScriptFileNames.bind(host);
            host.getScriptFileNames = () => {
                const files = originalGetScriptFileNames();
                try {
                    if (files.some(f => f.toLowerCase().endsWith('.osts'))) {
                        if (!files.includes(declarationPath)) {
                            return [...files, declarationPath];
                        }
                    }
                } catch (err) {
                    log(`Error in getScriptFileNames proxy: ${err}`);
                }
                return files;
            };

            log(`Plugin create() installed getScriptFileNames proxy`);
            // Return the untouched language service — no method-level proxying
            // needed since the host-level injection handles everything.
            return info.languageService;
        } catch (err) {
            log(`FATAL: Error in plugin create(): ${err}`);
            throw err;
        }
    }

    return { create };
}

export = init;
