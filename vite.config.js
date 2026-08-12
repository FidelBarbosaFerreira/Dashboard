import { defineConfig } from 'vite';
import react from '@vitejs/plugin-react';

// https://vite.dev/config/
export default defineConfig({
  plugins: [react()],
  base: './', // 👈 Define caminhos relativos para funcionar em subpastas do SharePoint/OneDrive
});
