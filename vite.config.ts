import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'
import devCerts from 'office-addin-dev-certs'
import { resolve } from 'path'
import { fileURLToPath } from 'url'

const __dirname = fileURLToPath(new URL('.', import.meta.url))

export default defineConfig(async () => {
  const httpsOptions = await devCerts.getHttpsServerOptions()

  return {
    plugins: [react()],

    server: {
      port: 3000,
      https: httpsOptions,
    },

    resolve: {
      alias: {
        '@': resolve(__dirname, 'src'),
      },
    },

    build: {
      rollupOptions: {
        input: {
          taskpane: resolve(__dirname, 'src/taskpane/taskpane.html'),
          commands: resolve(__dirname, 'src/commands/commands.html'),
          // auth/callback.html is served as a static file from public/ — no bundling needed
        },
      },
    },

    // Map process.env.NODE_ENV for any Office.js internals that reference it
    define: {
      'process.env.NODE_ENV': JSON.stringify(process.env.NODE_ENV ?? 'development'),
    },

    test: {
      globals: true,
      environment: 'jsdom',
      setupFiles: ['./tests/setup.ts'],
      coverage: {
        reporter: ['text', 'html'],
        include: ['src/**'],
      },
    },
  }
})
