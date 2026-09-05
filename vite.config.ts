import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'
import { sites } from '@openai/sites-vite-plugin'

export default defineConfig({
  base: process.env.GITHUB_ACTIONS ? '/orgnize/' : '/',
  plugins: [
    react(),
    sites(),
    {
      name: 'sites-static-worker',
      generateBundle() {
        this.emitFile({
          type: 'asset',
          fileName: 'server/index.js',
          source: `export default { fetch(request, env) { return env.ASSETS.fetch(request) } }\n`,
        })
      },
    },
  ],
})
