import { defaultClientConditions, defineConfig } from 'vite';

export default defineConfig({
  // Consume workspace packages straight from their TypeScript sources.
  resolve: { conditions: ['source', ...defaultClientConditions] },
  build: { target: 'es2022', sourcemap: true },
});
