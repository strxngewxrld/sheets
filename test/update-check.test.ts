import { describe, expect, it, vi } from 'vitest';
import { checkForUpdates, githubRepository, isNewerVersion, latestRelease, type UpdateCheckDeps } from '../src/update-check.js';

const REPO = 'git+https://github.com/example/sheets.git';

function respond(body: unknown, ok = true): typeof fetch {
  return vi.fn(async () => ({ ok, json: async () => body }) as Response);
}

function deps(overrides: Partial<UpdateCheckDeps>): UpdateCheckDeps {
  return {
    current: '0.2.3', repositoryUrl: REPO,
    fetch: respond({ tag_name: 'v0.3.0', html_url: 'https://github.com/example/sheets/releases/tag/v0.3.0' }),
    notify: vi.fn(async () => undefined), open: vi.fn(), skip: vi.fn(),
    ...overrides,
  };
}

describe('update check', () => {
  it('finds the GitHub repository in package.json URLs', () => {
    expect(githubRepository(REPO)).toBe('example/sheets');
    expect(githubRepository('https://github.com/example/sheets')).toBe('example/sheets');
    expect(githubRepository('git@github.com:example/sheets.git')).toBe('example/sheets');
    expect(githubRepository('https://gitlab.com/example/sheets')).toBeUndefined();
  });

  it('compares versions numerically', () => {
    expect(isNewerVersion('v0.10.0', '0.9.9')).toBe(true);
    expect(isNewerVersion('0.2.4', '0.2.3')).toBe(true);
    expect(isNewerVersion('1.0.0', '0.99.99')).toBe(true);
    expect(isNewerVersion('0.2.3', '0.2.3')).toBe(false);
    expect(isNewerVersion('0.2.2', '0.2.3')).toBe(false);
    expect(isNewerVersion('0.3.0-beta.1', '0.2.3')).toBe(false);
  });

  it('asks the latest release endpoint and rejects foreign release pages', async () => {
    const fetchImpl = respond({ tag_name: 'v1.0.0', html_url: 'https://github.com/example/sheets/releases/tag/v1.0.0' });
    expect(await latestRelease('example/sheets', fetchImpl)).toEqual({ version: '1.0.0', url: 'https://github.com/example/sheets/releases/tag/v1.0.0' });
    expect(vi.mocked(fetchImpl).mock.calls[0]![0]).toBe('https://api.github.com/repos/example/sheets/releases/latest');
    expect(await latestRelease('example/sheets', respond({ tag_name: 'v1.0.0', html_url: 'https://evil.example/releases/x' }))).toBeUndefined();
    expect(await latestRelease('example/sheets', respond({ message: 'Not Found' }, false))).toBeUndefined();
  });

  it('opens the release page when the user chooses it', async () => {
    const d = deps({ notify: vi.fn(async () => 'open' as const) });
    expect(await checkForUpdates(d)).toMatchObject({ version: '0.3.0' });
    expect(d.notify).toHaveBeenCalledWith({ version: '0.3.0', url: 'https://github.com/example/sheets/releases/tag/v0.3.0' });
    expect(d.open).toHaveBeenCalledWith('https://github.com/example/sheets/releases/tag/v0.3.0');
    expect(d.skip).not.toHaveBeenCalled();
  });

  it('remembers a skipped version and stays quiet for it', async () => {
    const d = deps({ notify: vi.fn(async () => 'skip' as const) });
    await checkForUpdates(d);
    expect(d.skip).toHaveBeenCalledWith('0.3.0');
    const again = deps({ skipped: '0.3.0' });
    expect(await checkForUpdates(again)).toBeUndefined();
    expect(again.notify).not.toHaveBeenCalled();
  });

  it('stays quiet when up to date or offline', async () => {
    const current = deps({ current: '0.3.0' });
    expect(await checkForUpdates(current)).toBeUndefined();
    const offline = deps({ fetch: vi.fn(async () => { throw new TypeError('fetch failed'); }) });
    expect(await checkForUpdates(offline)).toBeUndefined();
    expect(current.notify).not.toHaveBeenCalled();
    expect(offline.notify).not.toHaveBeenCalled();
  });
});
