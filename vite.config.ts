import { defineConfig } from 'vite';
import react from '@vitejs/plugin-react';

// base './' damit die App auch unter einem Unterpfad (GitHub Pages) läuft
export default defineConfig({
  base: './',
  plugins: [react()],
});
