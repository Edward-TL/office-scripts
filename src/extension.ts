import * as vscode from 'vscode';
import * as fs from 'fs';
import * as path from 'path';
import { subscribeToDocumentChanges } from './diagnostics';
import { OfficeScriptCodeActionProvider } from './codeActions';
import { OfficeScriptCompletionProvider } from './completionProvider';
import { OfficeScriptHoverProvider } from './hoverProvider';

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
    const selector: vscode.DocumentSelector = { language: 'office-script' };

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

    // Diagnostic command — invokable from Cmd+Shift+P regardless of the
    // active editor's language. Dumps plugin paths, whether each exists,
    // what the installed extension layout looks like, and a few commands
    // users may want to run manually (TS server commands are filtered out
    // of the palette when the active editor is office-script).
    context.subscriptions.push(
        vscode.commands.registerCommand('office-scripts.showDiagnostics', async () => {
            const extRoot = context.extensionPath;
            const probes = [
                ['extension root', extRoot],
                ['dist/plugin.js', path.join(extRoot, 'dist', 'plugin.js')],
                ['types/excel-script.d.ts', path.join(extRoot, 'types', 'excel-script.d.ts')],
                ['node_modules/office-scripts-plugin/package.json',
                    path.join(extRoot, 'node_modules', 'office-scripts-plugin', 'package.json')],
                ['node_modules/office-scripts-plugin/index.js',
                    path.join(extRoot, 'node_modules', 'office-scripts-plugin', 'index.js')],
            ];

            logChannel.appendLine('--- diagnostics ---');
            for (const [label, p] of probes) {
                const exists = fs.existsSync(p);
                logChannel.appendLine(`  ${exists ? 'OK ' : 'MISS'} ${label}: ${p}`);
            }

            const tsExt = vscode.extensions.getExtension('vscode.typescript-language-features');
            logChannel.appendLine(
                `  TS extension: id=${tsExt?.id ?? '(null)'} active=${tsExt?.isActive ?? false}`
            );

            const doc = vscode.window.activeTextEditor?.document;
            logChannel.appendLine(
                `  active doc: uri=${doc?.uri.toString() ?? '(none)'} languageId=${doc?.languageId ?? '(none)'}`
            );

            // List TypeScript-related commands that actually exist so the
            // user has a fallback for palette-filtered ones.
            const allCommands = await vscode.commands.getCommands(true);
            const tsCommands = allCommands.filter(c => c.startsWith('typescript.')).sort();
            logChannel.appendLine(`  typescript.* commands available: ${tsCommands.length}`);
            for (const c of tsCommands.slice(0, 20)) {
                logChannel.appendLine(`    ${c}`);
            }

            logChannel.show(true);
        }),
        vscode.commands.registerCommand('office-scripts.restartTsServer', async () => {
            try {
                await vscode.commands.executeCommand('typescript.restartTsServer');
                logChannel.appendLine('[command] typescript.restartTsServer invoked');
            } catch (err) {
                logChannel.appendLine(`[command] restart failed: ${err}`);
            }
            logChannel.show(true);
        }),
        vscode.commands.registerCommand('office-scripts.openTsServerLog', async () => {
            try {
                await vscode.commands.executeCommand('typescript.openTsServerLog');
            } catch (err) {
                logChannel.appendLine(`[command] openTsServerLog failed: ${err}`);
                logChannel.show(true);
            }
        }),
    );

    console.log('Office Scripts Support is now active.');
}

export function deactivate() {}
