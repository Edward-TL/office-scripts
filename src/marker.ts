/**
 * Regex matching a JSDoc block that carries the `@OfficeScript` tag. Used
 * to identify `.ts` files that should receive Office Scripts treatment
 * (module-wrap in the TS server plugin, diagnostic suppression, palette
 * commands). `.osts` files always qualify regardless of the tag.
 */
export const OFFICE_SCRIPT_MARKER = /\/\*\*[\s\S]*?@OfficeScript\b/;

export function hasOfficeScriptMarker(text: string): boolean {
    return OFFICE_SCRIPT_MARKER.test(text);
}

/**
 * Returns true if the given file + content pair should be treated as an
 * Office Script. `.osts` files always qualify; `.ts` files only if they
 * contain a `/** @OfficeScript *\/` JSDoc tag.
 */
export function isOfficeScriptFile(fileName: string, text: string): boolean {
    const lower = fileName.toLowerCase();
    if (lower.endsWith('.osts')) return true;
    if (lower.endsWith('.ts') && !lower.endsWith('.d.ts')) return hasOfficeScriptMarker(text);
    return false;
}
