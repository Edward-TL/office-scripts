import * as assert from 'assert';
import * as path from 'path';
import * as vscode from 'vscode';

suite('Office Scripts extension', () => {
    suiteSetup(() => {
        vscode.window.showInformationMessage('Starting Office Scripts tests.');
    });

    test('ExcelScript types resolve in .osts files', async () => {
        // Fixture lives next to the compiled tests in dist/; source-of-truth is src/test-usage.osts.
        const fixture = path.resolve(__dirname, '..', '..', 'src', 'test-usage.osts');
        const doc = await vscode.workspace.openTextDocument(fixture);
        await vscode.window.showTextDocument(doc);

        // Give tsserver time to load the plugin and index the ambient .d.ts.
        await new Promise(resolve => setTimeout(resolve, 2500));

        const diagnostics = vscode.languages.getDiagnostics(doc.uri);

        // TS2339 = "Property X does not exist on type Y". The fixture has exactly one
        // such call (guarded by @ts-expect-error), which tsserver SUPPRESSES — so we
        // expect zero surfaced TS2339 diagnostics. A failure here means types didn't load.
        const ts2339 = diagnostics.filter(d => d.code === 2339);
        assert.strictEqual(
            ts2339.length,
            0,
            `Expected no TS2339 diagnostics, got: ${ts2339.map(d => d.message).join(' | ')}`
        );

        // Custom rule from diagnostics.ts must fire on `const x: any`.
        const anyRuleFired = diagnostics.some(d =>
            typeof d.message === 'string' && d.message.includes('"any" type is forbidden')
        );
        assert.ok(anyRuleFired, 'Expected the "any is forbidden" custom diagnostic to fire.');

        // Custom rule from diagnostics.ts must fire on console.warn / console.error.
        const consoleRuleFired = diagnostics.some(d =>
            typeof d.message === 'string' && d.message.includes('is not supported in Office Scripts')
        );
        assert.ok(consoleRuleFired, 'Expected the console.warn/error custom diagnostic to fire.');
    });
});
