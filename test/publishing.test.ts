import { describe, expect, it, vi } from 'vitest';
import { releaseArgs, releaseNotes, releaseRepository, verifyRemoteTag } from '../tools/publish-github.mjs';

describe('GitHub VSIX releases', () => {
  it('resolves HTTPS and SSH repositories without creating an npm scope', () => {
    for (const source of ['git+https://github.com/Example/sheets.git', 'https://github.com/Example/sheets', 'git@github.com:Example/sheets.git', 'ssh://git@github.com/Example/sheets.git']) {
      expect(releaseRepository(source)).toBe('example/sheets');
    }
    expect(releaseRepository('https://github.com/example/sheets.git', 'Other')).toBe('other/sheets');
    expect(() => releaseRepository('https://example.org/example/sheets')).toThrow('Kein GitHub-Repository');
    expect(() => releaseRepository('https://github.com/example/sheets', '../bad')).toThrow('Ungültiger');
  });

  it('requires an existing remote tag and uploads the VSIX directly, optionally as a draft', () => {
    const input = { repository: 'example/sheets', tag: 'v1.0.0', asset: '/path with spaces/sheets.vsix', notes: '/tmp/notes.md', title: 'Sheets 1.0.0', draft: false };
    const args = releaseArgs(input);
    expect(args.slice(0, 4)).toEqual(['release', 'create', 'v1.0.0', input.asset]);
    expect(args).toContain('--verify-tag');
    expect(args).not.toContain('--clobber');
    expect(args).not.toContain('--draft');
    expect(releaseArgs({ ...input, draft: true })).toContain('--draft');
  });

  it('accepts a lightweight tag only when it matches the built commit', () => {
    const read = vi.fn(() => JSON.stringify({ type: 'commit', sha: 'head' }));
    expect(() => verifyRemoteTag('example/sheets', 'v1.0.0', 'head', read)).not.toThrow();
    expect(read).toHaveBeenCalledWith('gh', ['api', 'repos/example/sheets/git/ref/tags/v1.0.0', '--jq', '.object']);
    expect(() => verifyRemoteTag('example/sheets', 'v1.0.0', 'different-head', read)).toThrow('nicht auf den aktuellen Commit');
  });

  it('resolves annotated tags before comparing the built commit', () => {
    const read = vi.fn()
      .mockReturnValueOnce(JSON.stringify({ type: 'tag', sha: 'annotation' }))
      .mockReturnValueOnce(JSON.stringify({ type: 'commit', sha: 'head' }));
    verifyRemoteTag('example/sheets', 'v1.0.0', 'head', read);
    expect(read).toHaveBeenLastCalledWith('gh', ['api', 'repos/example/sheets/git/tags/annotation', '--jq', '.object']);
  });

  it('rejects missing or unresolvable tags before release creation', () => {
    expect(() => verifyRemoteTag('example/sheets', 'v1.0.0', 'head', () => { throw new Error('HTTP 404'); })).toThrow('404');
    const cycle = vi.fn(() => JSON.stringify({ type: 'tag', sha: 'cycle' }));
    expect(() => verifyRemoteTag('example/sheets', 'v1.0.0', 'head', cycle)).toThrow('nicht auf den aktuellen Commit');
    expect(cycle).toHaveBeenCalledTimes(17);
  });

  it('uses only the matching changelog section and includes direct VSIX installation', () => {
    const notes = releaseNotes('# Changelog\n\n## 1.0.1\n- Future\n\n## 1.0.0\n- Current\n\n## 0.9.0\n- Old\n', '1.0.0', 'sheets-1.0.0.vsix');
    expect(notes).toContain('- Current');
    expect(notes).not.toMatch(/Future|Old/);
    expect(notes).toContain('sheets-1.0.0.vsix');
    expect(notes).toContain('Install from VSIX');
    expect(releaseNotes('# Empty', '1.0.0', 'sheets.vsix')).toContain('sheets.vsix');
  });
});
