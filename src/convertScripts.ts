import * as fs from 'fs';
import * as path from 'path';

/**
 * Copies every `.osts` file from `sourceDir` into `outDir`, converting JSON
 * payloads into plain script text along the way. Power Automate-downloaded
 * Office Scripts are wrapped in a JSON envelope — we parse it and extract
 * the TypeScript body so the harvester can walk a folder of raw scripts.
 *
 * Returns the list of files written.
 */
export function convertScriptsToOsts(sourceDir: string, outDir: string): string[] {
    fs.mkdirSync(outDir, { recursive: true });
    const written: string[] = [];
    for (const file of listOstsFiles(sourceDir)) {
        const text = fs.readFileSync(file, 'utf8');
        const script = extractScriptFromJson(text) ?? text;
        const marked = injectOfficeScriptMarker(script);
        const outPath = path.join(outDir, path.basename(file, '.osts') + '.ts');
        fs.writeFileSync(outPath, ensureTrailingNewline(marked));
        written.push(outPath);
    }
    return written;
}

function listOstsFiles(root: string): string[] {
    const out: string[] = [];
    const walk = (dir: string) => {
        for (const entry of fs.readdirSync(dir, { withFileTypes: true })) {
            if (entry.name.startsWith('.') || entry.name === 'node_modules') continue;
            const full = path.join(dir, entry.name);
            if (entry.isDirectory()) walk(full);
            else if (entry.isFile() && entry.name.endsWith('.osts')) out.push(full);
        }
    };
    walk(root);
    return out;
}

/**
 * Attempts to parse `text` as JSON and pull a TypeScript body out of it.
 * Returns null if the text isn't JSON or no `function main(...)` string is
 * found anywhere in the parsed value.
 *
 * The heuristic tolerates unknown schemas: the downloaded envelope varies
 * across Power Automate / Office Scripts versions, so we scan commonly
 * named fields first and fall back to any string that contains the
 * `function main` signature.
 */
function extractScriptFromJson(text: string): string | null {
    let parsed: unknown;
    try {
        parsed = JSON.parse(text);
    } catch {
        return null;
    }
    return findScriptString(parsed);
}

function findScriptString(value: unknown): string | null {
    if (typeof value === 'string') {
        return /\bfunction\s+main\s*\(/.test(value) ? value : null;
    }
    if (Array.isArray(value)) {
        for (const v of value) {
            const hit = findScriptString(v);
            if (hit) return hit;
        }
        return null;
    }
    if (value && typeof value === 'object') {
        const obj = value as Record<string, unknown>;
        const preferred = ['body', 'script', 'code', 'content', 'source'];
        for (const key of preferred) {
            if (key in obj) {
                const hit = findScriptString(obj[key]);
                if (hit) return hit;
            }
        }
        for (const v of Object.values(obj)) {
            const hit = findScriptString(v);
            if (hit) return hit;
        }
    }
    return null;
}

function ensureTrailingNewline(text: string): string {
    return text.endsWith('\n') ? text : text + '\n';
}

/**
 * Inserts `/** @OfficeScript *\/` on the line directly above `function main`.
 * If the script already carries the marker anywhere, leaves the text alone.
 * Any existing JSDoc above `main` is preserved — the marker lands between
 * the docstring and the function declaration, so descriptive comments stay
 * attached to `main` and the marker remains recognizable to the plugin.
 */
function injectOfficeScriptMarker(text: string): string {
    if (/\/\*\*[\s\S]*?@OfficeScript\b/.test(text)) return text;
    const mainRe = /^([ \t]*)((?:export\s+)?(?:async\s+)?function\s+main\s*\()/m;
    return text.replace(mainRe, (_match, indent: string, rest: string) =>
        `${indent}/** @OfficeScript */\n${indent}${rest}`,
    );
}
