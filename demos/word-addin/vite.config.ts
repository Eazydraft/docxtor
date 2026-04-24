import { defineConfig } from 'vite';
import { resolve } from 'path';
import { fileURLToPath } from 'url';
import { getHttpsServerOptions } from 'office-addin-dev-certs';

const __dirname = fileURLToPath(new URL('.', import.meta.url));

const httpsOptions = await getHttpsServerOptions();

export default defineConfig({
  server: {
    port: 3000,
    https: httpsOptions,
  },
  resolve: {
    alias: [
      {
        find: '@eazydraft/docxtor',
        replacement: resolve(__dirname, '../../src/index.ts'),
      },
      {
        find: 'node:fs/promises',
        replacement: resolve(__dirname, 'src/stubs/node-fs-promises.ts'),
      },
    ],
  },
  build: {
    rollupOptions: {
      input: { taskpane: resolve(__dirname, 'index.html') },
    },
  },
});
