import { defaultClientConditions, defaultServerConditions } from 'vite';
import { defineConfig } from 'vitest/config';

// The "source" export condition lets workspace packages be consumed from
// TypeScript sources without a build step.
export default defineConfig({
  resolve: { conditions: ['source', ...defaultClientConditions] },
  ssr: { resolve: { conditions: ['source', ...defaultServerConditions], externalConditions: ['source'] } },
  test: { include: ['packages/*/test/**/*.test.ts'] },
});
