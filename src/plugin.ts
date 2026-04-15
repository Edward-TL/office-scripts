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
 * Layout at runtime:
 *   <ext-root>/dist/plugin.js         (this file, after esbuild)
 *   <ext-root>/types/excel-script.d.ts (ambient declarations we inject)
 */
function init({ typescript }: { typescript: typeof ts }) {
    function create(info: ts.server.PluginCreateInfo) {
        const log = (msg: string) =>
            info.project.projectService.logger.info(`[office-scripts] ${msg}`);

        const declarationPath = path.resolve(__dirname, '..', 'types', 'excel-script.d.ts');
        const normalizedPath = typescript.server.toNormalizedPath(declarationPath);

        let injected = false;

        const tryInject = () => {
            if (injected) return;
            if (!projectHasOsts(info.project)) return;

            const scriptInfo = info.project.projectService.getOrCreateScriptInfoForNormalizedPath(
                normalizedPath,
                /* openedByClient */ false,
                /* fileContent */ undefined,
                typescript.ScriptKind.TS,
                /* hasMixedContent */ false
            );

            if (!scriptInfo) {
                log(`Failed to load ExcelScript types from ${declarationPath}`);
                return;
            }

            if (!info.project.containsScriptInfo(scriptInfo)) {
                info.project.addRoot(scriptInfo);
                info.project.updateGraph();
                log(`Loaded ExcelScript types for project ${info.project.getProjectName()}`);
            }
            injected = true;
        };

        // Try immediately — covers the common case where the user opened an
        // .osts file first.
        tryInject();

        // Proxy the language service so that late-opened .osts files still
        // trigger injection on the next query.
        const proxy: ts.LanguageService = Object.create(null);
        for (const k of Object.keys(info.languageService) as (keyof ts.LanguageService)[]) {
            const original = info.languageService[k];
            (proxy as any)[k] = (...args: unknown[]) => {
                tryInject();
                return (original as Function).apply(info.languageService, args);
            };
        }

        return proxy;
    }

    return { create };
}

function projectHasOsts(project: ts.server.Project): boolean {
    // Fast path: any already-known file with the .osts extension.
    const files = project.getFileNames(/* excludeFilesFromExternalLibraries */ true);
    return files.some(f => f.toLowerCase().endsWith('.osts'));
}

export = init;
