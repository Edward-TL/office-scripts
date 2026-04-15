import * as vscode from 'vscode';
import { subscribeToDocumentChanges } from './diagnostics';
import { OfficeScriptCodeActionProvider } from './codeActions';

/**
 * Activates the Office Scripts extension.
 *
 * Type resolution (ExcelScript.*) is handled by the TS Server Plugin
 * registered in package.json under `contributes.typescriptServerPlugins`.
 * The plugin module lives at dist/plugin.js and is resolved via the
 * `office-scripts-plugin` file: dependency declared in package.json.
 *
 * This function wires up:
 *   - Custom diagnostics for Office-Scripts-specific rules (src/diagnostics.ts)
 *   - Quick-fix code actions for those diagnostics (src/codeActions.ts)
 */
export function activate(context: vscode.ExtensionContext) {
    const officeScriptsDiagnostics = vscode.languages.createDiagnosticCollection("office-scripts");
    context.subscriptions.push(officeScriptsDiagnostics);

    subscribeToDocumentChanges(context, officeScriptsDiagnostics);

    // Register quick-fixes for .osts files only.
    context.subscriptions.push(
        vscode.languages.registerCodeActionsProvider(
            { language: 'typescript', pattern: '**/*.osts' },
            new OfficeScriptCodeActionProvider(),
            { providedCodeActionKinds: OfficeScriptCodeActionProvider.providedCodeActionKinds }
        )
    );

    console.log('Office Scripts Support is now active.');
}

export function deactivate() {}
