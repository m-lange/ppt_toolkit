import { defineConfig } from 'vite';
import react from '@vitejs/plugin-react';
import basicSsl from '@vitejs/plugin-basic-ssl';

export default defineConfig({
  base: '/ppt_toolkit/',
  plugins: [react(), basicSsl()],
  server: {
    port: 3000,
    https: true,
  },
});