import type { HarvestedItem } from './harvest';

export interface DistilledVariant {
    source: string;
    originFiles: string[];
}

export interface DistilledResult {
    /** Declarations whose name is unique across the harvest, or whose duplicates all had identical normalized bodies. One entry per name. */
    core: Map<string, DistilledVariant>;
    /** Declarations that appeared under the same name in multiple scripts with differing bodies. One entry per name, holding every distinct variant. */
    conflict: Map<string, DistilledVariant[]>;
}

/**
 * Groups harvested items by name, then splits each group:
 *   - all variants share the same normalized body → single `core` entry
 *   - variants differ → `conflict` entry listing each distinct body
 */
export function distillItems(harvested: HarvestedItem[]): DistilledResult {
    const byName = new Map<string, HarvestedItem[]>();
    for (const item of harvested) {
        const bucket = byName.get(item.name);
        if (bucket) bucket.push(item);
        else byName.set(item.name, [item]);
    }

    const core = new Map<string, DistilledVariant>();
    const conflict = new Map<string, DistilledVariant[]>();

    for (const [name, items] of byName) {
        const variants = new Map<string, DistilledVariant>();
        for (const item of items) {
            const existing = variants.get(item.normalized);
            if (existing) existing.originFiles.push(item.sourceFile);
            else variants.set(item.normalized, { source: item.source, originFiles: [item.sourceFile] });
        }

        if (variants.size === 1) {
            core.set(name, variants.values().next().value!);
        } else {
            conflict.set(name, [...variants.values()]);
        }
    }

    return { core, conflict };
}
