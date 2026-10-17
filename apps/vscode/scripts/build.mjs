// Bundles the extension host and webview code and generates the
// languageModelTools contribution from the agent tool definitions.
import { build, context } from 'esbuild';
import { readFile, writeFile, rm, mkdtemp } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join, dirname } from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';

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
  for (const b of builds) await (await context(b)).watch();
} else {
  await Promise.all(builds.map((b) => build(b)));
}
