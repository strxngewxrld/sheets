#!/usr/bin/env node
/** Build and verify Sheets, then attach the VSIX directly to a GitHub Release. */
import { execFileSync } from 'node:child_process';
import { createHash } from 'node:crypto';
import { mkdirSync, readFileSync, writeFileSync } from 'node:fs';
import { join, resolve } from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { parseArgs } from 'node:util';

const root = fileURLToPath(new URL('..', import.meta.url));

function run(command, args, options = {}) {
  return execFileSync(command, args, { cwd: root, stdio: 'inherit', ...options });
}
function output(command, args) {
  return run(command, args, { stdio: 'pipe', encoding: 'utf8' }).trim();
}
function npm(args) {
  if (process.env.npm_execpath) return run(process.execPath, [process.env.npm_execpath, ...args]);
  return run(process.platform === 'win32' ? 'npm.cmd' : 'npm', args, { shell: process.platform === 'win32' });
}

export function releaseRepository(source, ownerOverride) {
  const match = /^(?:git\+https:\/\/github\.com\/|https:\/\/github\.com\/|git@github\.com:|ssh:\/\/git@github\.com\/)([a-z\d-]+)\/([a-z\d_.-]+?)\/?$/i.exec(source.replace(/\.git\/?$/, ''));
  if (!match) throw new Error('Kein GitHub-Repository gefunden. repository.url oder origin prüfen.');
  const owner = (ownerOverride ?? match[1]).trim().toLowerCase();
  if (!/^[a-z\d](?:[a-z\d-]{0,37}[a-z\d])?$/.test(owner)) throw new Error('Ungültiger GitHub-Owner.');
  return `${owner}/${match[2]}`;
}

export function releaseArgs({ repository, tag, asset, notes, title, draft }) {
  return ['release', 'create', tag, asset, '--repo', repository, '--verify-tag', '--title', title,
    '--notes-file', notes, ...(draft ? ['--draft'] : [])];
}

export function releaseNotes(changelog, version, vsixName) {
  const lines = changelog.split(/\r?\n/);
  const start = lines.findIndex((line) => line.trim() === `## ${version}`);
  const end = start < 0 ? -1 : lines.findIndex((line, index) => index > start && /^## /.test(line));
  const changes = start < 0 ? '' : lines.slice(start + 1, end < 0 ? undefined : end).join('\n').trim();
  return (changes ? `${changes}\n\n` : '') +
    `Die Datei \`${vsixName}\` unter Assets herunterladen und in VS Code über „Extensions: Install from VSIX…“ installieren.\n`;
}

/** Verify lightweight and annotated tags against the exact local build commit. */
export function verifyRemoteTag(repository, tag, head, read = output) {
  let object = JSON.parse(read('gh', ['api', `repos/${repository}/git/ref/tags/${tag}`, '--jq', '.object']));
  for (let depth = 0; object.type === 'tag' && depth < 16; depth++) {
    object = JSON.parse(read('gh', ['api', `repos/${repository}/git/tags/${object.sha}`, '--jq', '.object']));
  }
  if (object.type !== 'commit' || object.sha !== head) {
    throw new Error(`Der GitHub-Tag ${tag} zeigt nicht auf den aktuellen Commit. Passenden Release-Commit auschecken.`);
  }
}

export function main() {
  const { values } = parseArgs({ options: {
    owner: { type: 'string' },
    // Retain the previously documented --bump none invocation.
    bump: { type: 'string', default: 'none' },
    'dry-run': { type: 'boolean', default: false }, draft: { type: 'boolean', default: false },
    help: { type: 'boolean', short: 'h' },
  } });
  if (values.help) {
    console.log('npm run publish:github -- [--owner <user-oder-org>] [--dry-run] [--draft]\n' +
      'Lädt die geprüfte VSIX direkt als GitHub-Release-Asset hoch. Keine npm-Veröffentlichung.\n' +
      'Voraussetzung: GitHub CLI (gh), Anmeldung und gepushter Tag v<package.json-Version> auf HEAD.\n' +
      '--dry-run baut und prüft lokal, schreibt einen Release-Plan und ruft GitHub nicht auf.\n' +
      '--bump none bleibt kompatibel; Versionsänderungen vorab committen und taggen.');
    return;
  }
  if (values.bump !== 'none') throw new Error(
    'Versionen vorab mit npm version patch|minor|major --no-git-tag-version --workspaces=false anheben, Release-Dateien aktualisieren, committen und taggen. Zum Veröffentlichen --bump none verwenden (Standard).',
  );
  const dryRun = values['dry-run'];
  const manifest = JSON.parse(readFileSync(join(root, 'package.json'), 'utf8'));
  const source = typeof manifest.repository === 'string' ? manifest.repository : manifest.repository?.url;
  const repository = releaseRepository(source || output('git', ['remote', 'get-url', 'origin']), values.owner ?? process.env.GITHUB_OWNER);
  const version = manifest.version, tag = `v${version}`;
  const head = output('git', ['rev-parse', 'HEAD']);
  const dirty = () => output('git', ['status', '--porcelain']);
  if (!dryRun) {
    if (dirty()) throw new Error('Vor der Veröffentlichung alle Änderungen committen und den Release-Tag pushen. Für lokale Prüfungen --dry-run verwenden.');
    try { output('gh', ['--version']); }
    catch { throw new Error('GitHub CLI (gh) installieren und gh auth login ausführen oder GH_TOKEN/GITHUB_TOKEN setzen.'); }
    run('gh', ['auth', 'status', '--hostname', 'github.com']);
    try { verifyRemoteTag(repository, tag, head); }
    catch (error) { throw new Error(`Release-Tag prüfen: ${error.message}\nTag ${tag} muss in ${repository} existieren und auf HEAD zeigen.`); }
  } else {
    console.log('Dry-Run: keine Versionsänderung, kein GitHub-Aufruf und kein Upload.');
  }

  console.log('\nTests ausführen');
  npm(['test']);
  console.log('\nVSIX bauen und prüfen');
  npm(['run', 'package']);
  if (!dryRun && (dirty() || output('git', ['rev-parse', 'HEAD']) !== head)) {
    throw new Error('Der Arbeitsstand wurde während des Builds verändert. Vor Veröffentlichung prüfen und erneut starten.');
  }

  const vsixName = `${manifest.name}-${version}.vsix`;
  const asset = join(root, 'artifacts', vsixName);
  const notes = join(root, 'artifacts', `${tag}-release-notes.md`);
  const title = `${manifest.displayName} ${version}`;
  mkdirSync(join(root, 'artifacts'), { recursive: true });
  writeFileSync(notes, releaseNotes(readFileSync(join(root, 'CHANGELOG.md'), 'utf8'), version, vsixName));
  const args = releaseArgs({ repository, tag, asset, notes, title, draft: values.draft });
  const plan = join(root, 'artifacts', `${tag}-release-plan.json`);
  writeFileSync(plan, JSON.stringify({ repository, tag, commit: head, version, asset,
    sha256: createHash('sha256').update(readFileSync(asset)).digest('hex'),
    draft: values.draft, command: ['gh', ...args],
  }, null, 2) + '\n');
  console.log(`\nRelease-Plan: ${plan}`);
  if (dryRun) {
    console.log(`Geplant: ${repository} / ${tag}${values.draft ? ' (Entwurf)' : ''}, Asset ${vsixName}. Nichts veröffentlicht.`);
    return;
  }
  // gh uploads assets to a draft before publishing; existing releases are not overwritten.
  run('gh', args);
  console.log(`\nFertig: ${repository} / ${tag}${values.draft ? ' (Entwurf)' : ''}`);
}

if (process.argv[1] && import.meta.url === pathToFileURL(resolve(process.argv[1])).href) {
  try { main(); }
  catch (error) {
    console.error(`\nAbbruch: ${error instanceof Error ? error.message : String(error)}`);
    process.exitCode = 1;
  }
}
