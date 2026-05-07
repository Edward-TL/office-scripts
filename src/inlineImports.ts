import * as vscode from 'vscode';
import * as ts from 'typescript';
import * as fs from 'fs';
import * as path from 'path';
import { isOfficeScriptFile } from './marker';

/**
 * Produces an Excel-ready version of an Office Script with all relative
 * imports replaced by inlined copies of the referenced declarations.
 *
 * Beyond just the explicitly imported names, the inliner also pulls in
 * non-exported helpers and types that the imported declarations depend on,
 * walking transitively across files. Type aliases and interfaces are
 * hoisted to the top of the output; other declarations are appended after
 * the original source.
 *
 * Accepts `.osts` files and `.ts` files tagged with `/** @OfficeScript *\/`.
 *
 * Not supported (warnings issued):
 *   - `import * as ns from ...` namespace imports
 *   - Default imports
 *   - Non-relative imports (node_modules, bare specifiers)
 */
export async function inlineImports(doc: vscode.TextDocument): Promise<void> {
    const source = doc.getText();
    if (!isOfficeScriptFile(doc.fileName, source)) {
        vscode.window.showErrorMessage(
            'Inline Imports requires an .osts file or a .ts file tagged with /** @OfficeScript */.',
        );
        return;
    }

    const { output, warnings } = await inlineImportsToText(doc.fileName, source);

    const newDoc = await vscode.workspace.openTextDocument({
        content: output,
        language: doc.languageId,
    });
    await vscode.window.showTextDocument(newDoc);

    if (warnings.length > 0) {
        vscode.window.showWarningMessage(`Inline Imports: ${warnings.join('; ')}`);
    }
}

/**
 * Reusable core of the Inline Imports command: returns the fully-inlined
 * source text plus any warnings, without touching the editor.
 */
export async function inlineImportsToText(
    fileName: string,
    source: string,
): Promise<{ output: string; warnings: string[] }> {
    const warnings: string[] = [];
    const collected = createCollected();
    const filesProcessed = new Map<string, Set<string>>();

    const sourceSf = ts.createSourceFile(fileName, source, ts.ScriptTarget.ES2020, true);
    const reservedNames = collectTopLevelNames(sourceSf);

    const topImports = collectImports(fileName, source, warnings);
    const fromDir = path.dirname(fileName);

    for (const imp of topImports) {
        await resolveAndCollect(fromDir, imp.specifier, imp.names, collected, filesProcessed, reservedNames, warnings);
    }

    const output = stripImportsAndAssemble(source, topImports, collected);
    return { output, warnings };
}

interface ImportInfo {
    specifier: string;
    names: string[];
    start: number;
    end: number;
}

interface Collected {
    types: Map<string, string>;
    values: Map<string, string>;
}

function createCollected(): Collected {
    return { types: new Map(), values: new Map() };
}

function allCollectedNames(c: Collected): Set<string> {
    return new Set([...c.types.keys(), ...c.values.keys()]);
}

/**
 * Resolve a set of relative imports to the source text of their exported
 * declarations (and any non-exported / type dependencies they pull in).
 * Shared with the Split Flows command. Returned Map iterates types/interfaces
 * first, then other declarations.
 */
export async function resolveImportsToDeclarations(
    fromDir: string,
    imports: { specifier: string; names: string[] }[],
    warnings: string[],
): Promise<Map<string, string>> {
    const collected = createCollected();
    const filesProcessed = new Map<string, Set<string>>();
    for (const imp of imports) {
        await resolveAndCollect(fromDir, imp.specifier, imp.names, collected, filesProcessed, new Set(), warnings);
    }
    const merged = new Map<string, string>();
    for (const [k, v] of collected.types) merged.set(k, v);
    for (const [k, v] of collected.values) merged.set(k, v);
    return merged;
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

interface SymbolEntry {
    stmt: ts.Statement;
    isExported: boolean;
    isType: boolean;
}

function buildSymbolTable(sf: ts.SourceFile): Map<string, SymbolEntry> {
    const table = new Map<string, SymbolEntry>();
    for (const stmt of sf.statements) {
        const isExported = !!getExportKeyword(stmt);
        if (ts.isVariableStatement(stmt)) {
            for (const d of stmt.declarationList.declarations) {
                if (ts.isIdentifier(d.name)) {
                    table.set(d.name.text, { stmt, isExported, isType: false });
                }
            }
            continue;
        }
        const name = getDeclarationName(stmt);
        if (!name) continue;
        const isType = ts.isTypeAliasDeclaration(stmt) || ts.isInterfaceDeclaration(stmt);
        table.set(name, { stmt, isExported, isType });
    }
    return table;
}

function collectTopLevelNames(sf: ts.SourceFile): Set<string> {
    return new Set(buildSymbolTable(sf).keys());
}

/**
 * Collect identifiers referenced inside a node, skipping property names so
 * `obj.foo` only counts `obj` (not `foo`). Doesn't try to honor lexical
 * scoping — local shadowing leads to a slight over-pull, never an under-pull.
 */
function collectIdentifierRefs(node: ts.Node): Set<string> {
    const refs = new Set<string>();
    function visit(n: ts.Node) {
        if (ts.isIdentifier(n)) {
            refs.add(n.text);
            return;
        }
        if (ts.isPropertyAccessExpression(n)) {
            visit(n.expression);
            return;
        }
        if (ts.isQualifiedName(n)) {
            visit(n.left);
            return;
        }
        if (ts.isPropertyAssignment(n)) {
            visit(n.initializer);
            return;
        }
        if (ts.isPropertySignature(n) || ts.isPropertyDeclaration(n) || ts.isMethodSignature(n) || ts.isMethodDeclaration(n)) {
            if (n.type) visit(n.type);
            if ('initializer' in n && n.initializer) visit(n.initializer as ts.Node);
            if ('body' in n && n.body) visit(n.body as ts.Node);
            if ('parameters' in n && n.parameters) {
                for (const p of n.parameters) visit(p);
            }
            return;
        }
        ts.forEachChild(n, visit);
    }
    visit(node);
    return refs;
}

async function resolveAndCollect(
    fromDir: string,
    specifier: string,
    requestedNames: string[],
    collected: Collected,
    filesProcessed: Map<string, Set<string>>,
    reservedNames: Set<string>,
    warnings: string[],
): Promise<void> {
    const resolved = resolveModulePath(fromDir, specifier);
    if (!resolved) {
        warnings.push(`Cannot resolve "${specifier}" from ${fromDir}`);
        return;
    }

    let processed = filesProcessed.get(resolved);
    if (!processed) {
        processed = new Set();
        filesProcessed.set(resolved, processed);
    }
    const seedNames = requestedNames.filter(n => !processed!.has(n));
    if (seedNames.length === 0) return;

    const text = fs.readFileSync(resolved, 'utf8');
    const sf = ts.createSourceFile(resolved, text, ts.ScriptTarget.ES2020, true);
    const fileLabel = path.basename(resolved);

    const symbols = buildSymbolTable(sf);

    const nestedImports = collectImports(resolved, text, warnings);
    const importedNameToSpec = new Map<string, string>();
    for (const ni of nestedImports) {
        for (const n of ni.names) importedNameToSpec.set(n, ni.specifier);
    }

    // Closure walk: starting from new requested (exported) names, expand to
    // any identifier referenced by an included decl that resolves to another
    // top-level symbol in this file. Names already processed in a prior call
    // are skipped — they were emitted (and their refs walked) before.
    const closureOrder: string[] = [];
    const closure = new Set<string>();
    const queue: string[] = [];
    const nestedNeeded = new Map<string, Set<string>>();

    for (const n of seedNames) {
        const sym = symbols.get(n);
        if (!sym) {
            warnings.push(`"${n}" not found in ${fileLabel}`);
            continue;
        }
        if (!sym.isExported) {
            warnings.push(`"${n}" is not exported from ${fileLabel}`);
            continue;
        }
        queue.push(n);
    }

    while (queue.length) {
        const name = queue.shift()!;
        if (closure.has(name) || processed.has(name)) continue;
        closure.add(name);
        processed.add(name);
        closureOrder.push(name);

        const sym = symbols.get(name);
        if (!sym) continue;
        const refs = collectIdentifierRefs(sym.stmt);
        for (const r of refs) {
            if (r === name) continue;
            if (symbols.has(r)) {
                if (!closure.has(r) && !processed.has(r)) queue.push(r);
                continue;
            }
            if (importedNameToSpec.has(r)) {
                const spec = importedNameToSpec.get(r)!;
                let bucket = nestedNeeded.get(spec);
                if (!bucket) {
                    bucket = new Set();
                    nestedNeeded.set(spec, bucket);
                }
                bucket.add(r);
            }
        }
    }

    // Decide renames for non-exported decls that conflict with already-collected
    // names or with the importer's own top-level names.
    const taken = new Set<string>([...allCollectedNames(collected), ...reservedNames]);
    const renames = new Map<string, string>();
    for (const name of closureOrder) {
        const sym = symbols.get(name)!;
        if (sym.isExported) continue;
        if (!taken.has(name)) {
            taken.add(name);
            continue;
        }
        let i = 1;
        let candidate = `${name}_${i}`;
        while (taken.has(candidate) || symbols.has(candidate)) {
            i++;
            candidate = `${name}_${i}`;
        }
        renames.set(name, candidate);
        taken.add(candidate);
        warnings.push(`Renamed non-exported "${name}" → "${candidate}" (collision while inlining ${fileLabel})`);
    }

    // Emit snippets, types first within the file. Dedup by stmt so
    // multi-declarator VariableStatements aren't repeated.
    const sortedClosure = [...closureOrder].sort((a, b) => {
        const at = symbols.get(a)!.isType;
        const bt = symbols.get(b)!.isType;
        return at === bt ? 0 : at ? -1 : 1;
    });
    const emittedStmts = new Set<ts.Statement>();
    for (const name of sortedClosure) {
        const sym = symbols.get(name)!;
        if (emittedStmts.has(sym.stmt)) continue;

        // Skip exported decls already collected from another file.
        const targetMap = sym.isType ? collected.types : collected.values;
        const otherMap = sym.isType ? collected.values : collected.types;
        const finalName = renames.get(name) ?? name;
        if (sym.isExported && (targetMap.has(finalName) || otherMap.has(finalName))) {
            emittedStmts.add(sym.stmt);
            continue;
        }

        emittedStmts.add(sym.stmt);
        let snippet = extractDecl(sym.stmt, text);
        for (const [oldN, newN] of renames) {
            snippet = renameInSnippet(snippet, oldN, newN);
        }
        targetMap.set(finalName, snippet);
    }

    // Recurse into nested imports — only with names actually referenced.
    for (const [spec, names] of nestedNeeded) {
        await resolveAndCollect(
            path.dirname(resolved),
            spec,
            [...names],
            collected,
            filesProcessed,
            reservedNames,
            warnings,
        );
    }
}

function resolveModulePath(fromDir: string, specifier: string): string | undefined {
    const base = path.resolve(fromDir, specifier);
    const candidates = [
        base,
        base + '.ts',
        base + '.tsx',
        base + '.osts',
        base + '.d.ts',
        path.join(base, 'index.ts'),
        path.join(base, 'index.tsx'),
        path.join(base, 'index.osts'),
    ];
    return candidates.find(p => {
        try {
            return fs.statSync(p).isFile();
        } catch {
            return false;
        }
    });
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

/**
 * Extract a top-level declaration's text, including any attached JSDoc,
 * with the leading `export` keyword stripped if present.
 */
function extractDecl(stmt: ts.Statement, source: string): string {
    const jsDocs = (stmt as unknown as { jsDoc?: ts.JSDoc[] }).jsDoc;
    const start = jsDocs && jsDocs.length > 0 ? jsDocs[0].getStart() : stmt.getStart();
    const end = stmt.getEnd();
    let snippet = source.slice(start, end);

    const exportKw = getExportKeyword(stmt);
    if (exportKw) {
        const relStart = exportKw.getStart() - start;
        let relEnd = exportKw.getEnd() - start;
        while (snippet[relEnd] === ' ' || snippet[relEnd] === '\t') relEnd++;
        snippet = snippet.slice(0, relStart) + snippet.slice(relEnd);
    }
    return snippet;
}

function escapeRegex(s: string): string {
    return s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

function renameInSnippet(snippet: string, oldName: string, newName: string): string {
    return snippet.replace(new RegExp(`\\b${escapeRegex(oldName)}\\b`, 'g'), newName);
}

function stripImportsAndAssemble(source: string, imports: ImportInfo[], collected: Collected): string {
    let body = source;

    for (const imp of [...imports].sort((a, b) => b.start - a.start)) {
        const lineStart = body.lastIndexOf('\n', imp.start - 1) + 1;
        let lineEnd = body.indexOf('\n', imp.end);
        if (lineEnd === -1) lineEnd = body.length;
        else lineEnd += 1;
        body = body.slice(0, lineStart) + body.slice(lineEnd);
    }

    const typeBlock = Array.from(collected.types.values()).join('\n\n').trim();
    const valueBlock = Array.from(collected.values.values()).join('\n\n').trim();
    const bodyTrimmed = body.trim();

    const parts: string[] = [];
    if (typeBlock) parts.push(typeBlock);
    if (bodyTrimmed) parts.push(bodyTrimmed);
    if (valueBlock) parts.push(valueBlock);

    return parts.join('\n\n') + '\n';
}
