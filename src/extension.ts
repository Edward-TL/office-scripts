import * as vscode from 'vscode';
import * as fs from 'fs';
import * as path from 'path';
import { subscribeToDocumentChanges } from './diagnostics';
import { OfficeScriptCodeActionProvider } from './codeActions';
import { OfficeScriptCompletionProvider } from './completionProvider';
import { OfficeScriptHoverProvider } from './hoverProvider';
import { inlineImports } from './inlineImports';
import { splitFlows } from './flowSplit';
import { harvestCoreLibrary } from './harvestCoreLibrary';
import { exportToOsts } from './exportToOsts';
import { exportAllToOsts } from './exportAllToOsts';

/**
 * Activates the Office Scripts extension.
 *
 * Type resolution (ExcelScript.*) is handled by the TS Server Plugin
 * registered in package.json under `contributes.typescriptServerPlugins`.
 * The plugin module lives at dist/plugin.js and is resolved via the
 * `office-scripts-plugin` file: dependency declared in package.json.
 *
 * Because .osts files use a custom language id ("office-script") rather
 * than "typescript", we must force-activate VS Code's TypeScript feature
 * extension manually — its own activation events only cover typescript /
 * typescriptreact language ids, so tsserver wouldn't start otherwise.
 *
 * This function wires up the VS Code-side providers that augment tsserver:
 *   - Custom diagnostics           (src/diagnostics.ts)
 *   - Quick-fix code actions       (src/codeActions.ts)
 *   - Context-aware completions    (src/completionProvider.ts)
 *   - Hover with docs links        (src/hoverProvider.ts)
 */
export async function activate(context: vscode.ExtensionContext) {
    const selector: vscode.DocumentSelector = [
        { language: 'office-script' },
        { language: 'typescript' },
    ];

    // Force a TypeScript feature extension to activate so tsserver runs with
    // our plugin loaded for office-script files. We try three strategies in
    // order: (1) the known built-in ID, (2) the Nightly replacement, and
    // (3) any installed extension whose ID looks TypeScript-related. The
    // fallback catches forks or renamed extensions we don't know about.
    const knownIds = [
        'vscode.typescript-language-features',
        'ms-vscode.vscode-typescript-next',
    ];
    const tsLike = (id: string) =>
        /typescript/i.test(id) && /language|features|nightly|next/i.test(id);
    const candidates: vscode.Extension<unknown>[] = [];
    for (const id of knownIds) {
        const ext = vscode.extensions.getExtension(id);
        if (ext) candidates.push(ext);
    }
    for (const ext of vscode.extensions.all) {
        if (tsLike(ext.id) && !candidates.some(c => c.id === ext.id)) {
            candidates.push(ext);
        }
    }

    const logChannel = vscode.window.createOutputChannel('Office Scripts');
    context.subscriptions.push(logChannel);
    logChannel.appendLine(
        `[activate] TypeScript-like extension candidates: ${candidates.map(c => c.id).join(', ') || '(none)'}`
    );

    let tsActivated = false;
    for (const ext of candidates) {
        try {
            if (!ext.isActive) await ext.activate();
            logChannel.appendLine(`[activate] activated ${ext.id}`);
            tsActivated = true;
            break;
        } catch (err) {
            logChannel.appendLine(`[activate] failed to activate ${ext.id}: ${err}`);
        }
    }
    if (!tsActivated) {
        logChannel.appendLine(
            `[activate] no TS extension activated. Total extensions visible: ${vscode.extensions.all.length}`
        );
        logChannel.show(true);
        void vscode.window.showWarningMessage(
            'Office Scripts: could not activate a TypeScript language-features extension. ' +
            'Check the "Office Scripts" output channel for details.'
        );
    }

    const officeScriptsDiagnostics = vscode.languages.createDiagnosticCollection("office-scripts");
    context.subscriptions.push(officeScriptsDiagnostics);
    subscribeToDocumentChanges(context, officeScriptsDiagnostics);

    context.subscriptions.push(
        vscode.languages.registerCodeActionsProvider(
            selector,
            new OfficeScriptCodeActionProvider(),
            { providedCodeActionKinds: OfficeScriptCodeActionProvider.providedCodeActionKinds }
        ),
        vscode.languages.registerCompletionItemProvider(
            selector,
            new OfficeScriptCompletionProvider(),
            '"',
            "'"
        ),
        vscode.languages.registerHoverProvider(
            selector,
            new OfficeScriptHoverProvider()
        )
    );

    // Proxy built-in TypeScript commands so they surface in the Command
    // Palette for .osts files. The built-in commands are contributed with
    // `when: editorLangId == typescript|javascript|...`, which excludes our
    // `office-script` language id. We expose equivalents under an
    // "Office Scripts" category, gated on `editorLangId == office-script`.
    const proxy = (from: string, to: string) =>
        vscode.commands.registerCommand(from, (...args: unknown[]) =>
            vscode.commands.executeCommand(to, ...args)
        );
    context.subscriptions.push(
        proxy('officeScripts.restartTsServer', 'typescript.restartTsServer'),
        proxy('officeScripts.reloadProjects', 'typescript.reloadProjects'),
        proxy('officeScripts.selectTypeScriptVersion', 'typescript.selectTypeScriptVersion'),
        proxy('officeScripts.openTsServerLog', 'typescript.openTsServerLog'),
        proxy('officeScripts.goToProjectConfig', 'typescript.goToProjectConfig'),
        proxy('officeScripts.goToSourceDefinition', 'typescript.goToSourceDefinition')
    );

    context.subscriptions.push(
        vscode.commands.registerCommand('officeScripts.inlineImports', async () => {
            const editor = vscode.window.activeTextEditor;
            if (!editor) {
                vscode.window.showErrorMessage('Open an .osts file first.');
                return;
            }
            await inlineImports(editor.document);
        }),
        vscode.commands.registerCommand('officeScripts.splitFlows', async () => {
            const editor = vscode.window.activeTextEditor;
            if (!editor) {
                vscode.window.showErrorMessage('Open an .osts file first.');
                return;
            }
            await splitFlows(editor.document);
        }),
        vscode.commands.registerCommand('officeScripts.harvestCoreLibrary', async () => {
            await harvestCoreLibrary();
        }),
        vscode.commands.registerCommand('officeScripts.exportToOsts', async () => {
            const editor = vscode.window.activeTextEditor;
            if (!editor) {
                vscode.window.showErrorMessage('Open a .ts Office Script first.');
                return;
            }
            await exportToOsts(editor.document);
        }),
        vscode.commands.registerCommand('officeScripts.exportAllToOsts', async (folderUri?: vscode.Uri) => {
            await exportAllToOsts(folderUri);
        }),
    );

    // Push the `officeScripts.strictDiagnostics` setting to our TS Server
    // plugin. VS Code forwards this via tsserver's `configurePlugin`
    // request so the plugin receives it in `onConfigurationChanged`.
    const pushPluginConfig = async () => {
        const cfg = vscode.workspace.getConfiguration('officeScripts');
        try {
            await vscode.commands.executeCommand('typescript.tsserverRequest', 'configurePlugin', {
                pluginName: 'office-scripts-plugin',
                configuration: { strictDiagnostics: cfg.get<boolean>('strictDiagnostics', false) },
            });
        } catch (err) {
            logChannel.appendLine(`[config] failed to push plugin config: ${err}`);
        }
    };
    await pushPluginConfig();
    context.subscriptions.push(
        vscode.workspace.onDidChangeConfiguration(e => {
            if (e.affectsConfiguration('officeScripts.strictDiagnostics')) {
                void pushPluginConfig();
            }
        }),
    );

    console.log('Office Scripts Support is now active.');
}

export function deactivate() {}
