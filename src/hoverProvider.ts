import * as vscode from 'vscode';

/**
 * Adds a Microsoft Learn documentation link on hover when the user points
 * at a known ExcelScript type name. Merges with (does not replace) tsserver's
 * native hover output.
 *
 * Detection is deliberately narrow: we only activate when the hovered word
 * is a known ExcelScript type AND is preceded on the same line by
 * `ExcelScript.`. This avoids polluting hovers on unrelated identifiers that
 * happen to share a name (e.g. a local variable named `Range`).
 */
const EXCEL_SCRIPT_TYPES = new Set<string>([
    'Workbook',
    'Worksheet',
    'Range',
    'Table',
    'TableColumn',
    'RangeFormat',
    'RangeFill',
    'RangeFont',
    'Chart',
    'PivotTable',
    'DataValidation',
    'Comment',
    'Shape',
    'Image',
    'ConditionalFormat',
    'NamedItem',
    'Slicer'
]);

export class OfficeScriptHoverProvider implements vscode.HoverProvider {
    provideHover(
        document: vscode.TextDocument,
        position: vscode.Position
    ): vscode.Hover | undefined {
        if (!document.fileName.endsWith('.osts')) {
            return undefined;
        }

        const wordRange = document.getWordRangeAtPosition(position);
        if (!wordRange) return undefined;

        const word = document.getText(wordRange);
        if (!EXCEL_SCRIPT_TYPES.has(word)) return undefined;

        const lineText = document.lineAt(position.line).text;
        const beforeWord = lineText.substring(0, wordRange.start.character);
        if (!/ExcelScript\s*\.\s*$/.test(beforeWord)) {
            return undefined;
        }

        const slug = word.toLowerCase();
        const docsUrl = `https://learn.microsoft.com/javascript/api/office-scripts/excelscript/excelscript.${slug}`;

        const md = new vscode.MarkdownString();
        md.isTrusted = true;
        md.appendMarkdown(`**Office Scripts** — \`ExcelScript.${word}\`\n\n`);
        md.appendMarkdown(`[Open Microsoft Learn reference](${docsUrl})`);
        return new vscode.Hover(md, wordRange);
    }
}
