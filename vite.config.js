import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'

// L'aplicació ja no es desplega a GitHub Pages — només es treballa amb
// el servidor de desenvolupament local. Si en algun moment es vol tornar
// a publicar, restaurar `base: '/radar-gestor-app/'` per al mode build.
export default defineConfig({
  plugins: [react()],
  base: '/',
  server: { port: 5173, open: true }
})
