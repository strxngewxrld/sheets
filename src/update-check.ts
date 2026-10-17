/** Checks GitHub Releases for a newer version of the extension. Kept free of the VS Code API for tests. */

export interface LatestRelease {
  version: string;
  url: string;
}

export interface UpdateCheckDeps {
  /** Installed version, e.g. "0.2.3". */
  current: string;
  /** repository.url from package.json. */
  repositoryUrl: string;
  fetch: typeof fetch;
  /** Version the user chose to skip, if any. */
  skipped?: string;
  /** Shows the notification and resolves with the chosen action. */
  notify(release: LatestRelease): Promise<'open' | 'skip' | undefined>;
  open(url: string): void;
  skip(version: string): void;
}

/** "owner/repo" of a GitHub repository URL, or undefined for other hosts. */
export function githubRepository(url: string): string | undefined {
  const match = /^(?:git\+)?(?:https:\/\/github\.com\/|git@github\.com:|ssh:\/\/git@github\.com\/)([\w.-]+)\/([\w.-]+?)(?:\.git)?\/?$/i.exec(url.trim());
  return match ? `${match[1]}/${match[2]}` : undefined;
}

function parseVersion(text: string): number[] | undefined {
  const match = /^v?(\d+)\.(\d+)\.(\d+)$/.exec(text.trim());
  return match ? match.slice(1).map(Number) : undefined;
}

/** True when `candidate` is a higher x.y.z version than `current`. */
export function isNewerVersion(candidate: string, current: string): boolean {
  const a = parseVersion(candidate), b = parseVersion(current);
  if (!a || !b) return false;
  for (let i = 0; i < 3; i++) if (a[i] !== b[i]) return a[i]! > b[i]!;
  return false;
}

/** Latest published (non-draft, non-prerelease) release of a GitHub repository. */
export async function latestRelease(repository: string, fetchImpl: typeof fetch): Promise<LatestRelease | undefined> {
  const response = await fetchImpl(`https://api.github.com/repos/${repository}/releases/latest`, {
    headers: { Accept: 'application/vnd.github+json', 'X-GitHub-Api-Version': '2022-11-28' },
    signal: AbortSignal.timeout(10_000),
  });
  if (!response.ok) return undefined;
  const data = await response.json() as { tag_name?: unknown; html_url?: unknown };
  if (typeof data.tag_name !== 'string' || typeof data.html_url !== 'string') return undefined;
  // Only ever open the release page of this repository.
  if (!data.html_url.toLowerCase().startsWith(`https://github.com/${repository.toLowerCase()}/releases/`)) return undefined;
  return { version: data.tag_name.replace(/^v/, ''), url: data.html_url };
}

/** Notifies about a newer release; network and API errors are ignored. Returns the release that was offered. */
export async function checkForUpdates(deps: UpdateCheckDeps): Promise<LatestRelease | undefined> {
  const repository = githubRepository(deps.repositoryUrl);
  if (!repository) return undefined;
  const release = await latestRelease(repository, deps.fetch).catch(() => undefined);
  if (!release || !isNewerVersion(release.version, deps.current) || release.version === deps.skipped) return undefined;
  const choice = await deps.notify(release);
  if (choice === 'open') deps.open(release.url);
  else if (choice === 'skip') deps.skip(release.version);
  return release;
}
