// Bundles the extension host and webview code and generates the
// languageModelTools contribution from the agent tool definitions.
import { build, context } from 'esbuild';
import { readFile, writeFile, rm, mkdtemp } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join, dirname } from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { writeNotices } from './notices.mjs';

const root = join(dirname(fileURLToPath(import.meta.url)), '..');
const watch = process.argv.includes('--watch');
const common = { bundle: true, sourcemap: true, minify: !watch, conditions: ['source'], logLevel: 'info', absWorkingDir: root };

async function generateToolManifest() {
  const dir = await mkdtemp(join(tmpdir(), 'sheets-tools-'));
  const out = join(dir, 'tools.mjs');
  await build({ ...common, logLevel: 'warning', entryPoints: [join(root, 'src/tool-manifest.ts')], format: 'esm', platform: 'node', outfile: out, sourcemap: false });
  const { languageModelTools } = await import(pathToFileURL(out).href);
  await rm(dir, { recursive: true, force: true });
  const pkgPath = join(root, 'package.json');
  const pkg = JSON.parse(await readFile(pkgPath, 'utf8'));
  pkg.contributes.languageModelTools = languageModelTools;
  await writeFile(pkgPath, JSON.stringify(pkg, null, 2) + '\n');
  console.log(`languageModelTools: ${languageModelTools.length} tools written to package.json`);
}

await generateToolManifest();
const builds = [
  { ...common, entryPoints: ['src/extension.ts'], platform: 'node', format: 'cjs', target: 'node20', external: ['vscode'], outfile: 'dist/extension.js' },
  { ...common, entryPoints: ['src/webview/main.ts'], platform: 'browser', format: 'iife', target: 'es2022', outfile: 'dist/webview.js' },
];
if (watch) {
  // Report readiness only after both host and webview bundles finish. F5 must
  // never open a new host bundle with an outdated/missing webview bundle.
  let watching = false;
  const building = new Set();
  let failed = false;
  const contexts = await Promise.all(builds.map((b) => context({ ...b, logLevel: 'warning', plugins: [{
    name: 'watch-ready',
    setup(build) {
      build.onStart(() => {
        if (!building.size) failed = false;
        building.add(b.outfile);
        console.log('[watch] build started');
      });
      build.onEnd((result) => {
        failed ||= result.errors.length > 0;
        building.delete(b.outfile);
        if (watching && !building.size && !failed) console.log('[watch] build finished');
      });
    },
  }] })));
  await Promise.all(contexts.map((c) => c.rebuild()));
  console.log('[watch] build finished');
  watching = true;
  await Promise.all(contexts.map((c) => c.watch()));
} else {
  await Promise.all(builds.map((b) => build(b)));
  await writeNotices(root);
}
