import * as vscode from 'vscode';
import * as ts from 'typescript';
import * as fs from 'fs';
import * as path from 'path';

/**
 * Produces an Excel-ready version of an .osts file with all relative imports
 * replaced by inlined copies of the referenced declarations. The result is
 * opened in a new untitled editor so the user can review and paste into the
 * Office Scripts editor in Excel Online.
 *
 * Supported today:
 *   - Named imports from relative paths (`import { foo } from './helpers'`)
 *   - Nested helpers (a helper file importing from another helper file)
 *   - Exported functions, classes, interfaces, type aliases, enums, and
 *     top-level `const`/`let`/`var` declarations
 *
 * Not supported (warnings issued):
 *   - `import * as ns from ...` namespace imports
 *   - Default imports
 *   - Non-relative imports (node_modules, bare specifiers)
 */
export async function inlineImports(doc: vscode.TextDocument): Promise<void> {
    if (doc.languageId !== 'office-script') {
        vscode.window.showErrorMessage('Inline Imports only works on .osts files.');
        return;
    }

    const source = doc.getText();
    const collected = new Map<string, string>();
    const visited = new Set<string>();
    const warnings: string[] = [];

    const topImports = collectImports(doc.fileName, source, warnings);
    const docDir = path.dirname(doc.fileName);

    for (const imp of topImports) {
        await resolveAndCollect(docDir, imp.specifier, imp.names, collected, visited, warnings);
    }

    const output = stripImportsAndAppend(source, topImports, collected);

    const newDoc = await vscode.workspace.openTextDocument({
        content: output,
        language: 'office-script',
    });
    await vscode.window.showTextDocument(newDoc);

    if (warnings.length > 0) {
        vscode.window.showWarningMessage(`Inline Imports: ${warnings.join('; ')}`);
    }
}

interface ImportInfo {
    specifier: string;
    names: string[];
    start: number;
    end: number;
}

/**
 * Resolve a set of relative imports to the source text of their exported
 * declarations. Shared with the Split Flows command so a split file can
 * inline imported helpers instead of carrying the import line forward.
 *
 * Walks transitive relative imports — if helper A imports from helper B,
 * B's declarations are resolved too.
 */
export async function resolveImportsToDeclarations(
    fromDir: string,
    imports: { specifier: string; names: string[] }[],
    warnings: string[],
): Promise<Map<string, string>> {
    const collected = new Map<string, string>();
    const visited = new Set<string>();
    for (const imp of imports) {
        await resolveAndCollect(fromDir, imp.specifier, imp.names, collected, visited, warnings);
    }
    return collected;
}

function collectImports(fileName: string, text: string, warnings: string[]): ImportInfo[] {
    const sf = ts.createSourceFile(fileName, text, ts.ScriptTarget.ES2020, true);
    const imports: ImportInfo[] = [];

    for (const stmt of sf.statements) {
        if (!ts.isImportDeclaration(stmt)) continue;
        if (!ts.isStringLiteral(stmt.moduleSpecifier)) continue;

        const specifier = stmt.moduleSpecifier.text;
        if (!specifier.startsWith('.')) {
            warnings.push(`Skipped non-relative import "${specifier}"`);
            continue;
        }

        const clause = stmt.importClause;
        if (!clause) continue;

        if (clause.name) {
            warnings.push(`Skipped default import from "${specifier}"`);
        }

        const names: string[] = [];
        if (clause.namedBindings) {
            if (ts.isNamespaceImport(clause.namedBindings)) {
                warnings.push(`Skipped namespace import from "${specifier}"`);
            } else if (ts.isNamedImports(clause.namedBindings)) {
                for (const el of clause.namedBindings.elements) {
                    names.push(el.propertyName?.text ?? el.name.text);
                }
            }
        }

        if (names.length > 0) {
            imports.push({ specifier, names, start: stmt.getStart(sf), end: stmt.getEnd() });
        }
    }

    return imports;
}

async function resolveAndCollect(
    fromDir: string,
    specifier: string,
    names: string[],
    collected: Map<string, string>,
    visited: Set<string>,
    warnings: string[],
): Promise<void> {
    const resolved = resolveModulePath(fromDir, specifier);
    if (!resolved) {
        warnings.push(`Cannot resolve "${specifier}" from ${fromDir}`);
        return;
    }
    if (visited.has(resolved)) return;
    visited.add(resolved);

    const text = fs.readFileSync(resolved, 'utf8');
    const sf = ts.createSourceFile(resolved, text, ts.ScriptTarget.ES2020, true);

    for (const stmt of sf.statements) {
        const exportKw = getExportKeyword(stmt);
        if (!exportKw) continue;

        if (ts.isVariableStatement(stmt)) {
            const declNames = stmt.declarationList.declarations
                .map(d => (ts.isIdentifier(d.name) ? d.name.text : undefined))
                .filter((n): n is string => !!n);
            if (declNames.some(n => names.includes(n))) {
                const snippet = extractWithoutExport(stmt, text, exportKw);
                for (const n of declNames) collected.set(n, snippet);
            }
            continue;
        }

        const declName = getDeclarationName(stmt);
        if (declName && names.includes(declName)) {
            collected.set(declName, extractWithoutExport(stmt, text, exportKw));
        }
    }

    for (const nested of collectImports(resolved, text, warnings)) {
        await resolveAndCollect(path.dirname(resolved), nested.specifier, nested.names, collected, visited, warnings);
    }
}

function resolveModulePath(fromDir: string, specifier: string): string | undefined {
    const base = path.resolve(fromDir, specifier);
    const candidates = [
        base + '.ts',
        base + '.tsx',
        base + '.osts',
        base + '.d.ts',
        path.join(base, 'index.ts'),
        path.join(base, 'index.tsx'),
        path.join(base, 'index.osts'),
    ];
    return candidates.find(p => fs.existsSync(p));
}

function getExportKeyword(stmt: ts.Statement): ts.Modifier | undefined {
    if (!ts.canHaveModifiers(stmt)) return undefined;
    const mods = ts.getModifiers(stmt);
    return mods?.find(m => m.kind === ts.SyntaxKind.ExportKeyword) as ts.Modifier | undefined;
}

function getDeclarationName(stmt: ts.Statement): string | undefined {
    if (ts.isFunctionDeclaration(stmt)) return stmt.name?.text;
    if (ts.isClassDeclaration(stmt)) return stmt.name?.text;
    if (ts.isInterfaceDeclaration(stmt)) return stmt.name.text;
    if (ts.isTypeAliasDeclaration(stmt)) return stmt.name.text;
    if (ts.isEnumDeclaration(stmt)) return stmt.name.text;
    return undefined;
}

function extractWithoutExport(stmt: ts.Statement, text: string, exportKw: ts.Modifier): string {
    const fullStart = stmt.getFullStart();
    const end = stmt.getEnd();
    let snippet = text.slice(fullStart, end);

    const relStart = exportKw.getStart() - fullStart;
    let relEnd = exportKw.getEnd() - fullStart;
    while (snippet[relEnd] === ' ' || snippet[relEnd] === '\t') relEnd++;

    snippet = snippet.slice(0, relStart) + snippet.slice(relEnd);
    return snippet.replace(/^\s*\n/, '');
}

function stripImportsAndAppend(source: string, imports: ImportInfo[], collected: Map<string, string>): string {
    let output = source;

    for (const imp of [...imports].sort((a, b) => b.start - a.start)) {
        const lineStart = output.lastIndexOf('\n', imp.start - 1) + 1;
        let lineEnd = output.indexOf('\n', imp.end);
        if (lineEnd === -1) lineEnd = output.length;
        else lineEnd += 1;
        output = output.slice(0, lineStart) + output.slice(lineEnd);
    }

    if (collected.size === 0) return output;

    const trailer = Array.from(collected.values()).join('\n\n');
    return output.trimEnd() + '\n\n' + trailer + '\n';
}
