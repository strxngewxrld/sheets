import { mkdir, readFile } from 'node:fs/promises';
import { spawnSync } from 'node:child_process';
import { fileURLToPath } from 'node:url';
import { join } from 'node:path';
import { checkPackage } from './check-package.mjs';

const root = fileURLToPath(new URL('..', import.meta.url));
const pkg = JSON.parse(await readFile(join(root, 'package.json'), 'utf8'));
const output = join(root, 'artifacts', `${pkg.name}-${pkg.version}.vsix`);
await mkdir(join(root, 'artifacts'), { recursive: true });
// vsce executes vscode:prepublish, which performs the typecheck and build.
const result = spawnSync(process.execPath, [join(root, 'node_modules/@vscode/vsce/vsce'), 'package', '--no-dependencies', '--allow-missing-repository', '--no-rewrite-relative-links', '--out', output], { cwd: root, stdio: 'inherit' });
if (result.error) throw result.error;
if (result.status !== 0) process.exit(result.status ?? 1);
await checkPackage(output);
