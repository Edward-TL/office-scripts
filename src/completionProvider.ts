import * as vscode from 'vscode';
import { isOfficeScriptFile } from './marker';

/**
 * Context-aware completions for Office Scripts.
 *
 * These complement tsserver's built-in type-based IntelliSense by filling in
 * string-literal arguments where tsserver has no signal — e.g. the A1
 * notation expected by Range.getRange(), hex colors for setColor(), or
 * enum-like string values for alignment setters.
 *
 * Triggered on " and ' so the list appears as soon as the user opens a
 * string literal in one of the recognized call contexts.
 */
export class OfficeScriptCompletionProvider implements vscode.CompletionItemProvider {
    provideCompletionItems(
        document: vscode.TextDocument,
        position: vscode.Position
    ): vscode.CompletionItem[] {
        if (!isOfficeScriptFile(document.fileName, document.getText())) {
            return [];
        }

        const line = document.lineAt(position.line).text;
        const beforeCursor = line.substring(0, position.character);

        if (/\.getRange\s*\(\s*["'][^"']*$/.test(beforeCursor)) {
            return this.rangeCompletions();
        }

        if (/\.setColor\s*\(\s*["'][^"']*$/.test(beforeCursor)) {
            return this.colorCompletions();
        }

        if (/\.(set|get)HorizontalAlignment\s*\(\s*["'][^"']*$/.test(beforeCursor)) {
            return this.alignmentCompletions();
        }

        if (/\.getTable\s*\(\s*["'][^"']*$/.test(beforeCursor)) {
            return this.tableNameHints();
        }

        return [];
    }

    private rangeCompletions(): vscode.CompletionItem[] {
        const entries: Array<{ label: string; detail: string }> = [
            { label: 'A1', detail: 'Single cell (top-left)' },
            { label: 'A1:B2', detail: 'Small range (2 cols x 2 rows)' },
            { label: 'A1:D10', detail: 'Data range (4 cols x 10 rows)' },
            { label: 'A1:Z100', detail: 'Large range (26 cols x 100 rows)' },
            { label: 'A:A', detail: 'Entire column A' },
            { label: 'A:D', detail: 'Columns A through D' },
            { label: '1:1', detail: 'Entire row 1' },
            { label: '1:10', detail: 'Rows 1 through 10' }
        ];
        return entries.map(e => {
            const item = new vscode.CompletionItem(e.label, vscode.CompletionItemKind.Value);
            item.detail = e.detail;
            item.insertText = e.label;
            item.sortText = '0_' + e.label;
            return item;
        });
    }

    private colorCompletions(): vscode.CompletionItem[] {
        const colors: Array<{ hex: string; name: string }> = [
            { hex: '#FF0000', name: 'Red' },
            { hex: '#00FF00', name: 'Green' },
            { hex: '#0000FF', name: 'Blue' },
            { hex: '#FFFF00', name: 'Yellow' },
            { hex: '#FFA500', name: 'Orange' },
            { hex: '#800080', name: 'Purple' },
            { hex: '#000000', name: 'Black' },
            { hex: '#FFFFFF', name: 'White' },
            { hex: '#808080', name: 'Gray' },
            { hex: '#217346', name: 'Excel Green' },
            { hex: '#B7DEE8', name: 'Light Blue (Excel theme)' },
            { hex: '#E2EFDA', name: 'Light Green (Excel theme)' }
        ];
        return colors.map(c => {
            const item = new vscode.CompletionItem(c.hex, vscode.CompletionItemKind.Color);
            item.detail = c.name;
            item.insertText = c.hex;
            item.documentation = new vscode.MarkdownString(`Hex color \`${c.hex}\``);
            item.sortText = '0_' + c.hex;
            return item;
        });
    }

    private alignmentCompletions(): vscode.CompletionItem[] {
        const values = [
            'General', 'Left', 'Center', 'Right',
            'Fill', 'Justify', 'CenterAcrossSelection', 'Distributed'
        ];
        return values.map(v => {
            const item = new vscode.CompletionItem(v, vscode.CompletionItemKind.EnumMember);
            item.insertText = v;
            item.sortText = '0_' + v;
            return item;
        });
    }

    private tableNameHints(): vscode.CompletionItem[] {
        // We can't know user table names without running Excel, so this is a
        // helpful placeholder set. Users typically rename these immediately.
        const placeholders = ['Table1', 'SalesData', 'Inventory', 'Orders'];
        return placeholders.map(p => {
            const item = new vscode.CompletionItem(p, vscode.CompletionItemKind.Value);
            item.detail = 'Common table-name placeholder';
            item.insertText = p;
            item.sortText = '9_' + p; // De-prioritize below real matches.
            return item;
        });
    }
}
