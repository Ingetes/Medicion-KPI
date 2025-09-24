import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'

// Cambia '/ingetes-kpi/' por el nombre de tu repo de GitHub Pages.
// Si vas a usar dominio propio o Vercel, puedes dejar base: './'.
export default defineConfig({
  plugins: [react()],
  base: '/ingetes-kpi/',
})
