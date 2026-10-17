import { readFile, readdir, writeFile } from 'node:fs/promises';
import { join } from 'node:path';

/** Include the actual bundled third-party licenses, without shipping node_modules. */
export async function writeNotices(root) {
  const packages = new Set();
  for (const bundle of ['extension', 'webview']) {
    const map = JSON.parse(await readFile(join(root, `dist/${bundle}.js.map`), 'utf8'));
    for (const source of map.sources) {
      if (!source.includes('node_modules/')) continue;
      const relative = source.split('node_modules/').at(-1);
      const parts = relative.split('/');
      packages.add(parts.slice(0, relative.startsWith('@') ? 2 : 1).join('/'));
    }
  }
  let text = '# Third-party notices\n\nThe extension bundles the following packages. Their licenses follow.\n';
  for (const name of [...packages].sort()) {
    const dir = join(root, 'node_modules', name);
    const pkg = JSON.parse(await readFile(join(dir, 'package.json'), 'utf8'));
    const license = (await readdir(dir)).find((file) => /^(?:licen[cs]e(?:\.(?:md|txt))?|copying)$/i.test(file));
    // @nodable/entities 3.1.0 declares MIT but omits LICENSE from its npm tarball.
    // Preserve the upstream repository license locally for offline builds.
    const fallback = name === '@nodable/entities' && pkg.version === '3.1.0'
      ? join(root, 'scripts/licenses/nodable-entities.LICENSE') : undefined;
    if (!license && !fallback) throw new Error(`Missing license file for bundled dependency ${name}`);
    text += `\n## ${name} ${pkg.version} (${pkg.license})\n\n`;
    if (fallback && !license) text += 'Source: https://github.com/nodable/val-parsers/blob/main/LICENSE\n\n';
    text += (await readFile(license ? join(dir, license) : fallback, 'utf8')).trim() + '\n';
  }
  await writeFile(join(root, 'THIRD_PARTY_NOTICES.md'), text);
  console.log(`Third-party licenses: ${packages.size} bundled dependencies`);
}
