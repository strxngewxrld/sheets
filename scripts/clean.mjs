import { readdir, rm, rmdir } from 'node:fs/promises';
import { join } from 'node:path';
import { fileURLToPath } from 'node:url';

const root = fileURLToPath(new URL('..', import.meta.url));
const projects = ['.', 'packages/core', 'packages/ui', 'packages/agent', 'packages/mcp'];
for (const project of [...projects, 'apps/vscode']) {
  const dir = join(root, project);
  for (const output of ['dist', 'dist-types']) await rm(join(dir, output), { recursive: true, force: true });
  for (const file of await readdir(dir).catch((e) => { if (e.code === 'ENOENT') return []; throw e; })) {
    if (file.endsWith('.tsbuildinfo')) await rm(join(dir, file));
  }
}
// Remove the former extension location only when its source folders are empty.
// Unexpected files are preserved, so clean never erases new work there.
for (const sub of ['src', 'test', 'scripts', 'media', '']) {
  await rmdir(join(root, 'apps/vscode', sub)).catch((e) => {
    if (!['ENOENT', 'ENOTEMPTY'].includes(e.code)) throw e;
  });
}
console.log('Build outputs and TypeScript caches removed. Source files and VSIX artifacts preserved.');
