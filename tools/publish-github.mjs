#!/usr/bin/env node
/**
 * Sheets als geprüftes VSIX in einer npm-Hülle auf GitHub Packages bereitstellen.
 * Aufruf: npm run publish:github -- [--owner <user-oder-org>]
 *   [--bump patch|minor|major|none] [--dry-run]
 *
 * Standard: Patch-Version anheben. --bump none nutzt die vorbereitete Version.
 * Ein Dry-Run verändert weder Version noch Git-Historie und lädt nichts hoch.
 * Es werden keine Commits, Tags, GitHub Releases oder Marketplace-Releases angelegt.
 */
import { execFileSync } from 'node:child_process';
import { copyFileSync, mkdirSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { parseArgs } from 'node:util';

const root = fileURLToPath(new URL('..', import.meta.url));
const output = join(root, 'artifacts', 'github-package');
const registry = 'https://npm.pkg.github.com';
const manifestPath = join(root, 'package.json');
const readManifest = () => JSON.parse(readFileSync(manifestPath, 'utf8'));

function run(command, args, options = {}) {
  return execFileSync(command, args, { cwd: root, stdio: 'inherit', ...options });
}
function npm(args, options = {}) {
  // npm run provides the CLI path; using Node avoids Windows .cmd quoting issues.
  if (process.env.npm_execpath) return run(process.execPath, [process.env.npm_execpath, ...args], options);
  return run(process.platform === 'win32' ? 'npm.cmd' : 'npm', args, {
    shell: process.platform === 'win32', ...options,
  });
}
function step(title) { console.log(`\n${title}`); }

function main() {
  const { values } = parseArgs({ options: {
    owner: { type: 'string' }, bump: { type: 'string', default: 'patch' },
    'dry-run': { type: 'boolean', default: false }, help: { type: 'boolean', short: 'h' },
  } });
  if (values.help) {
    console.log('npm run publish:github -- [--owner <user-oder-org>] [--bump patch|minor|major|none] [--dry-run]\n' +
      'Owner: --owner, GITHUB_OWNER oder Repository-Owner. Bump: patch (Standard).\n' +
      'Token: NODE_AUTH_TOKEN oder GITHUB_TOKEN; für lokale Uploads PAT (classic) mit write:packages.\n' +
      'Dry-Run: lokale Tests, VSIX-Prüfung und npm publish --dry-run; keine Versionsänderung und kein Upload.');
    return;
  }
  const bump = values.bump;
  const dryRun = values['dry-run'];
  if (!['patch', 'minor', 'major', 'none'].includes(bump)) throw new Error('Ungültiger --bump (patch, minor, major oder none).');

  const initial = readManifest();
  const source = typeof initial.repository === 'string' ? initial.repository : initial.repository?.url;
  const remote = source || run('git', ['remote', 'get-url', 'origin'], { stdio: 'pipe', encoding: 'utf8' }).trim();
  const match = /^(?:git\+https:\/\/github\.com\/|https:\/\/github\.com\/|git@github\.com:|ssh:\/\/git@github\.com\/)([a-z\d-]+)\/([a-z\d_.-]+?)\/?$/i.exec(remote.replace(/\.git\/?$/, ''));
  if (!match) throw new Error('Kein GitHub-Repository gefunden. repository.url oder origin auf das GitHub-Repo setzen.');
  const repository = { type: 'git', url: `git+https://github.com/${match[1]}/${match[2]}.git` };
  const owner = (values.owner ?? process.env.GITHUB_OWNER ?? match[1]).trim().toLowerCase();
  if (!/^[a-z\d](?:[a-z\d-]{0,37}[a-z\d])?$/.test(owner)) throw new Error('Ungültiger GitHub-Owner.');
  const token = process.env.NODE_AUTH_TOKEN || process.env.GITHUB_TOKEN || '';
  if (!token && !dryRun) throw new Error('NODE_AUTH_TOKEN oder GITHUB_TOKEN für GitHub Packages setzen (write:packages).');

  step('Arbeitsverzeichnis prüfen');
  const dirty = run('git', ['status', '--porcelain'], { stdio: 'pipe', encoding: 'utf8' }).trim();
  if (dirty && !dryRun) throw new Error('Es gibt nicht committete Änderungen. Vor der Veröffentlichung committen oder zunächst --dry-run verwenden.');
  if (dirty) console.log('Dry-Run mit lokalen Änderungen; keine Versionsänderung und kein Upload.');

  step('Tests ausführen');
  npm(['test']);
  if (bump !== 'none' && !dryRun) {
    step(`Root-Version anheben (${bump})`);
    npm(['version', bump, '--no-git-tag-version', '--ignore-scripts', '--workspaces=false']);
    console.log('Bei späterem Abbruch bleibt die angehobene Version bestehen. Wiederholung mit --bump none.');
  }
  const manifest = readManifest();
  const version = manifest.version;
  const vsixName = `${manifest.name}-${version}.vsix`;
  const packageName = `@${owner}/${manifest.name}`;

  step(`VSIX bauen und prüfen (${vsixName})`);
  // Runs the existing typecheck/build/vsce pipeline and the actual VSIX verifier.
  npm(['run', 'package']);

  step(`GitHub-Package vorbereiten (${packageName}@${version})`);
  rmSync(output, { recursive: true, force: true });
  mkdirSync(output, { recursive: true });
  copyFileSync(join(root, 'artifacts', vsixName), join(output, vsixName));
  copyFileSync(join(root, 'LICENSE'), join(output, 'LICENSE'));
  copyFileSync(join(root, 'THIRD_PARTY_NOTICES.md'), join(output, 'THIRD_PARTY_NOTICES.md'));
  writeFileSync(join(output, 'package.json'), JSON.stringify({
    name: packageName, version,
    description: `${manifest.displayName} – VS-Code-Extension als VSIX`,
    license: manifest.license,
    ...(manifest.author ? { author: manifest.author } : {}),
    repository,
    files: [vsixName, 'README.md', 'LICENSE', 'THIRD_PARTY_NOTICES.md'],
    publishConfig: { registry },
  }, null, 2) + '\n');
  writeFileSync(join(output, 'README.md'), `# ${manifest.displayName} ${version}\n\n` +
    'Dieses Paket enthält die Extension als VSIX, keine JavaScript-Bibliothek.\n' +
    `Quellcode: https://github.com/${match[1]}/${match[2]}\n\n` +
    'Mit einem GitHub-Token mit read:packages anmelden und installieren:\n\n```sh\n' +
    `npm login --scope=@${owner} --auth-type=legacy --registry=${registry}\n` +
    `npm pack ${packageName}@${version} --registry=${registry}\n` +
    `tar -xzf ${owner}-${manifest.name}-${version}.tgz package/${vsixName}\n` +
    `code --install-extension package/${vsixName}\n` + '```\n');
  // Only a placeholder is written. Never persist the credential itself.
  writeFileSync(join(output, '.npmrc'),
    `@${owner}:registry=${registry}\n//npm.pkg.github.com/:_authToken=\${NODE_AUTH_TOKEN}\n`);

  step(dryRun ? 'Probelauf: npm publish --dry-run' : `Veröffentlichen: ${packageName}@${version}`);
  npm(['publish', '--registry', registry, '--workspaces=false', '--ignore-scripts', ...(dryRun ? ['--dry-run'] : [])], {
    cwd: output, env: { ...process.env, NODE_AUTH_TOKEN: dryRun ? 'dry-run' : token },
  });
  console.log(`\nFertig: ${packageName}@${version}${dryRun ? ' (Dry-Run, nichts hochgeladen)' : ''}`);
  console.log(`Paketinhalt: ${output}`);
  if (bump !== 'none' && !dryRun) console.log(
    `Version ${version} steht in package.json und package-lock.json. README-Installationsbeispiel und CHANGELOG aktualisieren, danach committen und bei Bedarf taggen.`,
  );
}

try { main(); }
catch (error) {
  console.error(`\nAbbruch: ${error instanceof Error ? error.message : String(error)}`);
  process.exitCode = 1;
}
