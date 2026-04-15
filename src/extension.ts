import * as vscode from 'vscode';
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

    // Force the built-in TypeScript extension to activate so tsserver runs
    // with our plugin loaded for office-script files.
    const tsExtension = vscode.extensions.getExtension('vscode.typescript-language-features');
    if (tsExtension && !tsExtension.isActive) {
        try {
            await tsExtension.activate();
        } catch (err) {
            console.warn('Office Scripts: failed to activate TypeScript language features', err);
        }
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

    console.log('Office Scripts Support is now active.');
}

export function deactivate() {}
