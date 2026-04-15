const esbuild = require("esbuild");

const production = process.argv.includes('--production');
const watch = process.argv.includes('--watch');

/**
 * esbuild plugin to match VS Code problems
 */
const esbuildProblemMatcherPlugin = {
	name: 'esbuild-problem-matcher',
	setup(build) {
		build.onStart(() => console.log('[watch] build started'));
		build.onEnd((result) => {
			result.errors.forEach(({ text, location }) => {
				console.error(`✘ [ERROR] ${text}`);
				console.error(`    ${location.file}:${location.line}:${location.column}:`);
			});
			console.log('[watch] build finished');
		});
	},
};

async function main() {
	// Build for Extension logic
	const extensionCtx = await esbuild.context({
		entryPoints: ['src/extension.ts'],
		bundle: true,
		format: 'cjs',
		minify: production,
		sourcemap: !production,
		platform: 'node',
		outfile: 'dist/extension.js',
		external: ['vscode'],
		plugins: [esbuildProblemMatcherPlugin],
	});

	// Build for TS Server Plugin logic
	const pluginCtx = await esbuild.context({
		entryPoints: ['src/plugin.ts'],
		bundle: true,
		format: 'cjs',
		minify: production,
		sourcemap: !production,
		platform: 'node',
		outfile: 'dist/plugin.js',
		external: ['typescript'], // TS is provided by the server environment
	});

	if (watch) {
		await Promise.all([extensionCtx.watch(), pluginCtx.watch()]);
	} else {
		await Promise.all([extensionCtx.rebuild(), pluginCtx.rebuild()]);
		await Promise.all([extensionCtx.dispose(), pluginCtx.dispose()]);
	}
}

main().catch(e => {
	console.error(e);
	process.exit(1);
});
