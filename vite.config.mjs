import { defineConfig } from 'vite';
import { readFileSync } from 'node:fs';
const { version } = JSON.parse(readFileSync(new URL('./package.json', import.meta.url), 'utf8'));

export default defineConfig({
  base: './',
  plugins: [{ name: 'sheet-studio-build-version', generateBundle() {
    this.emitFile({ type: 'asset', fileName: 'app-version.json', source: JSON.stringify({ version })+'\n' });
  } }],
  server: { host: '127.0.0.1', port: 5192, strictPort: true },
  worker: {format: 'es'},
  build: {
    target: 'es2022',
    outDir: 'dist',
    chunkSizeWarningLimit: 1400,
    rollupOptions: { output: { manualChunks: { spreadsheet: ['exceljs/dist/exceljs.min.js'] } } },
  },
});
