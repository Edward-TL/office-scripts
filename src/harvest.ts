import * as ts from 'typescript';
import * as fs from 'fs';
import * as path from 'path';

export type HarvestedKind = 'function' | 'interface';

/**
 * One top-level declaration extracted from an `.osts` file. `normalized`
 * strips whitespace variations so two declarations with the same logic but
 * different formatting hash to the same value for duplicate detection.
 */
export interface HarvestedItem {
    kind: HarvestedKind;
    name: string;
    sourceFile: string;
    source: string;
    normalized: string;
}

/**
 * Walks `folder` recursively, parses every `.osts` file, and returns each
 * top-level function declaration (excluding `main`) and interface
 * declaration. Source text preserves any attached JSDoc; a leading
 * `export` keyword, if present, is stripped so callers can re-add it
 * consistently.
 */
export async function harvestItems(folder: string): Promise<HarvestedItem[]> {
    const files = listOstsFiles(folder);
    const harvested: HarvestedItem[] = [];

    for (const file of files) {
        const text = fs.readFileSync(file, 'utf8');
        const sf = ts.createSourceFile(file, text, ts.ScriptTarget.ES2020, true);
        for (const stmt of sf.statements) {
            if (ts.isFunctionDeclaration(stmt) && stmt.name && stmt.name.text !== 'main') {
                const source = extractDeclarationText(stmt, text);
                harvested.push({
                    kind: 'function',
                    name: stmt.name.text,
                    sourceFile: file,
                    source,
                    normalized: normalize(source),
                });
            } else if (ts.isInterfaceDeclaration(stmt)) {
                const source = extractDeclarationText(stmt, text);
                harvested.push({
                    kind: 'interface',
                    name: stmt.name.text,
                    sourceFile: file,
                    source,
                    normalized: normalize(source),
                });
            }
        }
    }

    return harvested;
}

function listOstsFiles(root: string): string[] {
    const out: string[] = [];
    const walk = (dir: string) => {
        for (const entry of fs.readdirSync(dir, { withFileTypes: true })) {
            if (entry.name.startsWith('.') || entry.name === 'node_modules') continue;
            const full = path.join(dir, entry.name);
            if (entry.isDirectory()) walk(full);
            else if (entry.isFile() && (entry.name.endsWith('.osts') || entry.name.endsWith('.ts'))) {
                out.push(full);
            }
        }
    };
    walk(root);
    return out;
}

function extractDeclarationText(
    stmt: ts.FunctionDeclaration | ts.InterfaceDeclaration,
    source: string,
): string {
    const jsDocs = (stmt as unknown as { jsDoc?: ts.JSDoc[] }).jsDoc;
    const start = jsDocs && jsDocs.length > 0 ? jsDocs[0].getStart() : stmt.getStart();
    let text = source.slice(start, stmt.getEnd());
    text = text.replace(/^(\s*)export\s+/, '$1');
    return text;
}

/**
 * Collapses runs of whitespace to a single space and trims. Used only for
 * equality comparison between two harvested items — never written to disk.
 */
function normalize(text: string): string {
    return text.replace(/\s+/g, ' ').trim();
}
