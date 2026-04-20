import * as vscode from 'vscode';
import * as fs from 'fs';
import * as path from 'path';
import { hasOfficeScriptMarker } from './marker';
import { buildOstsJson } from './exportToOsts';

/**
 * Bulk counterpart to **Export to OSTS**. Walks the given folder for every
 * `.ts` file tagged with `/** @OfficeScript *\/` and writes the resulting
 * OSTS envelope to a sibling `<folder>-osts/` directory, preserving the
 * nested structure. Untagged `.ts` files and non-`.ts` files are skipped.
 *
 * Invoked either from the Command Palette (prompts for a folder) or from
 * the Explorer context menu on a folder (uses the selected URI directly).
 */
export async function exportAllToOsts(folderUri?: vscode.Uri): Promise<void> {
    const folder = folderUri ?? (await pickFolder());
    if (!folder) return;

    const folderPath = folder.fsPath;
    let stat: fs.Stats;
    try {
        stat = fs.statSync(folderPath);
    } catch {
        vscode.window.showErrorMessage(`Folder not found: ${folderPath}`);
        return;
    }
    if (!stat.isDirectory()) {
        vscode.window.showErrorMessage(`Not a folder: ${folderPath}`);
        return;
    }

    const outDir = path.join(path.dirname(folderPath), `${path.basename(folderPath)}-osts`);
    fs.mkdirSync(outDir, { recursive: true });

    const tsFiles = listTaggedTsFiles(folderPath);
    if (tsFiles.length === 0) {
        vscode.window.showInformationMessage(
            `No .ts files with /** @OfficeScript */ were found in ${path.basename(folderPath)}/.`,
        );
        return;
    }

    const written: string[] = [];
    const skipped: string[] = [];
    const allWarnings: string[] = [];
    for (const { file, source } of tsFiles) {
        try {
            const { json, warnings } = await buildOstsJson(file, source);
            const rel = path.relative(folderPath, file);
            const outPath = path.join(outDir, rel.replace(/\.ts$/i, '.osts'));
            fs.mkdirSync(path.dirname(outPath), { recursive: true });
            fs.writeFileSync(outPath, json);
            written.push(path.basename(outPath));
            if (warnings.length > 0) {
                allWarnings.push(`${path.basename(file)}: ${warnings.join('; ')}`);
            }
        } catch (err) {
            skipped.push(`${path.basename(file)} (${(err as Error).message})`);
        }
    }

    const outRel = vscode.workspace.asRelativePath(outDir);
    vscode.window.showInformationMessage(
        `Exported ${written.length} script(s) to ${outRel}/: ${written.join(', ')}`,
    );
    if (skipped.length > 0) {
        vscode.window.showWarningMessage(`Export all TS to OSTS — skipped: ${skipped.join('; ')}`);
    }
    if (allWarnings.length > 0) {
        vscode.window.showWarningMessage(`Export all TS to OSTS: ${allWarnings.join(' | ')}`);
    }
}

async function pickFolder(): Promise<vscode.Uri | undefined> {
    const picks = await vscode.window.showOpenDialog({
        canSelectFiles: false,
        canSelectFolders: true,
        canSelectMany: false,
        openLabel: 'Select folder to export',
        defaultUri: vscode.workspace.workspaceFolders?.[0]?.uri,
    });
    return picks?.[0];
}

interface TaggedFile {
    file: string;
    source: string;
}

function listTaggedTsFiles(root: string): TaggedFile[] {
    const out: TaggedFile[] = [];
    const walk = (dir: string) => {
        for (const entry of fs.readdirSync(dir, { withFileTypes: true })) {
            if (entry.name.startsWith('.') || entry.name === 'node_modules') continue;
            const full = path.join(dir, entry.name);
            if (entry.isDirectory()) {
                walk(full);
                continue;
            }
            if (!entry.isFile()) continue;
            const lower = entry.name.toLowerCase();
            if (!lower.endsWith('.ts') || lower.endsWith('.d.ts')) continue;
            let source: string;
            try {
                source = fs.readFileSync(full, 'utf8');
            } catch {
                continue;
            }
            if (!hasOfficeScriptMarker(source)) continue;
            out.push({ file: full, source });
        }
    };
    walk(root);
    return out;
}
