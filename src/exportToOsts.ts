import * as vscode from 'vscode';
import * as fs from 'fs';
import * as path from 'path';
import { inlineImportsToText } from './inlineImports';
import { hasOfficeScriptMarker, OFFICE_SCRIPT_MARKER } from './marker';

/**
 * Converts a `.ts` Office Script (tagged with `/** @OfficeScript *\/`) into
 * a `.osts` file ready to upload to OneDrive → Scripts. Imports are
 * inlined, the `@OfficeScript` marker is stripped from the output, and
 * the resulting TypeScript body is wrapped in the JSON envelope that
 * Power Automate and Excel Online expect.
 *
 * Output filename: same basename as source, with `.osts` extension, in the
 * same folder as the source.
 */
export async function exportToOsts(doc: vscode.TextDocument): Promise<void> {
    const source = doc.getText();
    const lower = doc.fileName.toLowerCase();
    if (!lower.endsWith('.ts') || lower.endsWith('.d.ts') || !hasOfficeScriptMarker(source)) {
        vscode.window.showErrorMessage(
            'Export to OSTS requires a .ts file tagged with /** @OfficeScript */.',
        );
        return;
    }

    const { json, warnings } = await buildOstsJson(doc.fileName, source);
    const outPath = doc.fileName.replace(/\.ts$/i, '.osts');
    fs.writeFileSync(outPath, json);

    const relPath = vscode.workspace.asRelativePath(outPath);
    vscode.window.showInformationMessage(`Wrote ${relPath}. Ready to upload to OneDrive → Scripts.`);

    if (warnings.length > 0) {
        vscode.window.showWarningMessage(`Export to OSTS: ${warnings.join('; ')}`);
    }
}

/**
 * Reusable core of the single-file command. Given the source of a `.ts`
 * Office Script, returns the serialized OSTS envelope (JSON text) that
 * would be written to disk. Does not touch the filesystem. Used by the
 * bulk **Export all TS to OSTS** command, which writes each envelope to
 * a mirrored output folder.
 */
export async function buildOstsJson(
    fileName: string,
    source: string,
): Promise<{ json: string; warnings: string[] }> {
    const { output, warnings } = await inlineImportsToText(fileName, source);
    const scriptBody = stripOfficeScriptMarker(output);
    const envelope = buildEnvelope(scriptBody);
    return { json: JSON.stringify(envelope), warnings };
}

/**
 * Strips every `/** @OfficeScript *\/` JSDoc block from the text. The
 * marker is an authoring-time hint and shouldn't ship inside the uploaded
 * script. A JSDoc block that contains BOTH `@OfficeScript` and other
 * documentation is left alone — only drop the tag line within it.
 */
function stripOfficeScriptMarker(text: string): string {
    return text.replace(/\/\*\*\s*@OfficeScript\b[^*]*\*\/\s*\n?/g, '')
        .replace(/^[\s\t]*\*\s*@OfficeScript\b.*\n?/gm, '');
}

interface OstsEnvelope {
    version: string;
    body: string;
    description: string;
    noCodeMetadata: string;
    parameterInfo: string;
    apiInfo: string;
}

function buildEnvelope(body: string): OstsEnvelope {
    return {
        version: '0.3.0',
        body,
        description: '',
        noCodeMetadata: '',
        parameterInfo: JSON.stringify({
            version: 1,
            originalParameterOrder: [],
            parameterSchema: { type: 'object', default: {}, 'x-ms-visibility': 'internal' },
            returnSchema: { type: 'object', properties: {} },
            signature: { comment: '', parameters: [{ name: 'workbook', comment: '' }] },
        }),
        apiInfo: JSON.stringify({ variant: 'synchronous', variantVersion: 2 }),
    };
}

// Re-exported so callers that want the raw regex can reuse it.
export { OFFICE_SCRIPT_MARKER };
