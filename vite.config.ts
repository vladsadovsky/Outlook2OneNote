import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'
import devCerts from 'office-addin-dev-certs'
import { resolve } from 'path'
import { fileURLToPath } from 'url'

const __dirname = fileURLToPath(new URL('.', import.meta.url))

export default defineConfig(async () => {
  const httpsOptions = await devCerts.getHttpsServerOptions()
  
  // Use process.cwd() to work with junction points, force consistent path resolution
  const rootDir = process.cwd().replace(/\\/g, '/')

  return {
    plugins: [react()],
    
    // Explicitly set root to current working directory
    root: rootDir,

    server: {
      port: 3000,
      https: httpsOptions,
    },

    resolve: {
      alias: {
        '@': resolve(rootDir, 'src'),
      },
      // Preserve symlinks to maintain consistent paths
      preserveSymlinks: true,
    },

    build: {
      rollupOptions: {
        input: {
          taskpane: resolve(rootDir, 'src/taskpane/taskpane.html'),
          commands: resolve(rootDir, 'src/commands/commands.html'),
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
      setupFiles: [resolve(rootDir, 'tests/setup.ts')],
      coverage: {
        reporter: ['text', 'html'],
        include: ['src/**'],
      },
    },
  }
})
