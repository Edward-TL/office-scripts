import * as vscode from 'vscode';
import { isOfficeScriptFile } from './marker';

/**
 * Provides quick-fixes (lightbulb actions) for diagnostics emitted by
 * src/diagnostics.ts. Each diagnostic message here must match the string
 * produced in diagnostics.ts — keep the two in sync.
 */
export class OfficeScriptCodeActionProvider implements vscode.CodeActionProvider {
    public static readonly providedCodeActionKinds = [
        vscode.CodeActionKind.QuickFix
    ];

    provideCodeActions(
        document: vscode.TextDocument,
        _range: vscode.Range | vscode.Selection,
        context: vscode.CodeActionContext
    ): vscode.CodeAction[] {
        if (!isOfficeScriptFile(document.fileName, document.getText())) {
            return [];
        }

        const actions: vscode.CodeAction[] = [];

        for (const diagnostic of context.diagnostics) {
            if (diagnostic.message.includes('"any" type is forbidden')) {
                actions.push(this.replaceText(
                    document,
                    diagnostic,
                    'unknown',
                    'Replace "any" with "unknown"'
                ));
            }

            if (diagnostic.message.includes('console.warn is not supported')) {
                actions.push(this.replaceText(
                    document,
                    diagnostic,
                    'log',
                    'Replace console.warn with console.log'
                ));
            }

            if (diagnostic.message.includes('console.error is not supported')) {
                actions.push(this.replaceText(
                    document,
                    diagnostic,
                    'log',
                    'Replace console.error with console.log'
                ));
            }
        }

        return actions;
    }

    private replaceText(
        document: vscode.TextDocument,
        diagnostic: vscode.Diagnostic,
        replacement: string,
        title: string
    ): vscode.CodeAction {
        const action = new vscode.CodeAction(title, vscode.CodeActionKind.QuickFix);
        action.edit = new vscode.WorkspaceEdit();
        action.edit.replace(document.uri, diagnostic.range, replacement);
        action.diagnostics = [diagnostic];
        action.isPreferred = true;
        return action;
    }
}
