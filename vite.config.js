import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'

// L'aplicació ja no es desplega a GitHub Pages — només es treballa amb
// el servidor de desenvolupament local. Si en algun moment es vol tornar
// a publicar, restaurar `base: '/radar-gestor-app/'` per al mode build.
export default defineConfig({
  plugins: [react()],
  base: '/',
  // host:'127.0.0.1' força IPv4: evita que Vite escolti només a [::1] (IPv6),
  // cosa que feia que Firefox (que resol localhost a 127.0.0.1) no s'hi
  // pogués connectar. strictPort evita que salti al 5174 silenciosament:
  // si el 5173 està ocupat, fallarà amb un error clar en comptes de canviar.
  server: { host: '127.0.0.1', port: 5173, strictPort: true, open: true }
})
