import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'

// Si vas a usar dominio propio o Vercel, puedes dejar base: './'.
export default defineConfig({
  plugins: [react()],
  base: '/Medicion-KPI/',
})
