import { resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

import { defineConfig } from 'vite';
import vue from '@vitejs/plugin-vue';
import ui from '@nuxt/ui/vite';

const playgroundDir = fileURLToPath(new URL('.', import.meta.url));

export default defineConfig({
  plugins: [vue(), ui()],
  define: {
    __VUE_OPTIONS_API__: true,
    __VUE_PROD_DEVTOOLS__: false,
    __VUE_PROD_HYDRATION_MISMATCH_DETAILS__: false,
  },
  build: {
    outDir: resolve(playgroundDir, '..', 'dist-playground'),
    emptyOutDir: true,
    rollupOptions: {
      input: {
        playground: resolve(playgroundDir, 'index.html'),
      },
    },
  },
});
