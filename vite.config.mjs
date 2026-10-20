import { defineConfig } from 'vite';

export default defineConfig({
  base: './',
  server: { host: '127.0.0.1', port: 5192, strictPort: true },
  worker: {format: 'es'},
  build: {
    target: 'es2022',
    outDir: 'dist',
    chunkSizeWarningLimit: 1400,
    rollupOptions: { output: { manualChunks: { spreadsheet: ['exceljs/dist/exceljs.min.js'] } } },
  },
});
