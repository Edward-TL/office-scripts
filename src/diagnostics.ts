import * as vscode from 'vscode';
import * as ts from 'typescript';
import { isOfficeScriptFile } from './marker';

/**
 * Analyzes the TypeScript AST of an Office Script file to enforce strict
 * rules: No 'any' type, no 'console.warn', no 'console.error'.
 *
 * Applies to `.osts` files and to `.ts` files tagged with
 * `/** @OfficeScript *\/`. Untagged `.ts` files are ignored.
 */
export function refreshDiagnostics(doc: vscode.TextDocument, collection: vscode.DiagnosticCollection): void {
    const diagnostics: vscode.Diagnostic[] = [];

    if (!isOfficeScriptFile(doc.fileName, doc.getText())) {
        collection.delete(doc.uri);
        return;
    }

    const sourceFile = ts.createSourceFile(
        doc.fileName,
        doc.getText(),
        ts.ScriptTarget.Latest,
        true
    );

    /**
     * Recursive function to walk the AST and find violations.
     */
    function visit(node: ts.Node) {
        // 1. Prohibit explicit 'any' type
        if (node.kind === ts.SyntaxKind.AnyKeyword) {
            diagnostics.push(createDiagnostic(node, doc, 'Strict Error: The "any" type is forbidden in Office Scripts. Use specific interfaces or "unknown".'));
        }

        // 2. Prohibit console.warn and console.error (only console.log is supported)
        if (ts.isCallExpression(node) && ts.isPropertyAccessExpression(node.expression)) {
            const expression = node.expression;
            if (expression.expression.getText() === 'console') {
                const methodName = expression.name.getText();
                if (methodName === 'warn' || methodName === 'error') {
                    diagnostics.push(createDiagnostic(
                        expression.name, 
                        doc, 
                        `Strict Error: console.${methodName} is not supported in Office Scripts. Use console.log instead.`
                    ));
                }
            }
        }

        ts.forEachChild(node, visit);
    }

    visit(sourceFile);
    collection.set(doc.uri, diagnostics);
}

/**
 * Creates a VS Code Diagnostic object for a given AST node.
 */
function createDiagnostic(node: ts.Node, doc: vscode.TextDocument, message: string): vscode.Diagnostic {
    const start = doc.positionAt(node.getStart());
    const end = doc.positionAt(node.getEnd());
    const range = new vscode.Range(start, end);

    return new vscode.Diagnostic(
        range,
        message,
        vscode.DiagnosticSeverity.Error
    );
}

/**
 * Sets up listeners for document changes to trigger validation.
 */
export function subscribeToDocumentChanges(context: vscode.ExtensionContext, collection: vscode.DiagnosticCollection): void {
    if (vscode.window.activeTextEditor) {
        refreshDiagnostics(vscode.window.activeTextEditor.document, collection);
    }

    context.subscriptions.push(
        vscode.window.onDidChangeActiveTextEditor(editor => {
            if (editor) {
                refreshDiagnostics(editor.document, collection);
            }
        })
    );

    context.subscriptions.push(
        vscode.workspace.onDidChangeTextDocument(e => refreshDiagnostics(e.document, collection))
    );

    context.subscriptions.push(
        vscode.workspace.onDidCloseTextDocument(doc => collection.delete(doc.uri))
    );
}
