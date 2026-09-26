import path from 'node:path';
import react from '@vitejs/plugin-react';
import { defineConfig } from 'vite';

export default defineConfig({
  root: path.resolve(import.meta.dirname, 'src/renderer'),
  build: {
    outDir: path.resolve(import.meta.dirname, '.vite/renderer/main_window'),
  },
  plugins: [react()],
});
