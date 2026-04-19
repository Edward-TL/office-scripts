import * as vscode from 'vscode';
import * as fs from 'fs';
import * as path from 'path';
import { harvestItems } from './harvest';
import { distillItems, DistilledResult, DistilledVariant } from './distill';
import { convertScriptsToOsts } from './convertScripts';

/**
 * Scans a folder of `.osts` files, extracts every non-`main` top-level
 * function and every interface declaration, then writes them to a
 * destination directory split by duplicate status:
 *
 *   script-osts-version/*.osts  — plain-script copies of the source files,
 *                                  with any JSON envelope stripped
 *   core/<name>.ts              — unique functions, or dupes whose bodies match
 *   conflict/<name>/<n>.ts      — same-name functions, different bodies
 *   interface/core/<name>.ts    — unique interfaces (same rules)
 *   interface/conflict/<name>/<n>.ts
 *
 * Power Automate downloads each Office Script as a JSON envelope around
 * the actual TypeScript body. We normalize every source file through
 * `script-osts-version/` first so the harvester can parse raw script.
 *
 * Each extracted declaration is written as a standalone `export` so it can
 * be imported from other `.osts`/`.ts` files and re-inlined via
 * "Inline Imports for Excel Upload" before uploading to Excel.
 */
export async function harvestCoreLibrary(): Promise<void> {
    const sourceFolders = await vscode.window.showOpenDialog({
        canSelectFiles: false,
        canSelectFolders: true,
        canSelectMany: false,
        openLabel: 'Select folder of .osts scripts to harvest',
    });
    if (!sourceFolders || sourceFolders.length === 0) return;
    const sourceDir = sourceFolders[0].fsPath;

    const destFolders = await vscode.window.showOpenDialog({
        canSelectFiles: false,
        canSelectFolders: true,
        canSelectMany: false,
        openLabel: 'Select destination folder for core/, conflict/, and interface/',
    });
    if (!destFolders || destFolders.length === 0) return;
    const destDir = destFolders[0].fsPath;

    const scriptDir = path.join(destDir, 'script-osts-version');
    const converted = convertScriptsToOsts(sourceDir, scriptDir);
    if (converted.length === 0) {
        vscode.window.showInformationMessage('No .osts files found in the source folder.');
        return;
    }

    const harvested = await harvestItems(scriptDir);
    if (harvested.length === 0) {
        vscode.window.showInformationMessage(
            `Converted ${converted.length} file(s) to script-osts-version/, but found no non-main functions or interfaces.`,
        );
        return;
    }

    const functions = distillItems(harvested.filter(h => h.kind === 'function'));
    const interfaces = distillItems(harvested.filter(h => h.kind === 'interface'));

    writeGroup(path.join(destDir, 'core'), path.join(destDir, 'conflict'), functions, scriptDir);
    writeGroup(
        path.join(destDir, 'interface', 'core'),
        path.join(destDir, 'interface', 'conflict'),
        interfaces,
        scriptDir,
    );

    vscode.window.showInformationMessage(
        `Harvested ${harvested.length} decl(s): ` +
        `functions ${functions.core.size} core / ${functions.conflict.size} conflict, ` +
        `interfaces ${interfaces.core.size} core / ${interfaces.conflict.size} conflict.`,
    );
}

function writeGroup(
    coreDir: string,
    conflictDir: string,
    distilled: DistilledResult,
    sourceDir: string,
): void {
    if (distilled.core.size > 0) fs.mkdirSync(coreDir, { recursive: true });
    if (distilled.conflict.size > 0) fs.mkdirSync(conflictDir, { recursive: true });

    for (const [name, variant] of distilled.core) {
        writeVariant(path.join(coreDir, `${name}.ts`), variant, sourceDir);
    }
    for (const [name, variants] of distilled.conflict) {
        const dir = path.join(conflictDir, name);
        fs.mkdirSync(dir, { recursive: true });
        variants.forEach((variant, i) => {
            const slug = variantSlug(variant, i);
            writeVariant(path.join(dir, `${slug}.ts`), variant, sourceDir);
        });
    }
}

function writeVariant(outPath: string, variant: DistilledVariant, sourceDir: string): void {
    const header = buildOriginHeader(variant.originFiles, sourceDir);
    fs.writeFileSync(outPath, `${header}export ${variant.source}\n`);
}

function variantSlug(variant: DistilledVariant, index: number): string {
    // Prefer the first origin file's basename so conflict variants are
    // human-identifiable by where they came from. Fall back to an index if
    // two variants happen to share an origin filename.
    const first = variant.originFiles[0];
    const base = path.basename(first, '.osts').replace(/[^\w.-]+/g, '_');
    return `${base}-${index + 1}`;
}

function buildOriginHeader(originFiles: string[], sourceDir: string): string {
    const rels = originFiles.map(f => path.relative(sourceDir, f));
    return `// Harvested from: ${rels.join(', ')}\n`;
}
