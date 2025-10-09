import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'

export default defineConfig({
  plugins: [react()],
  base: '/Medicion-KPI/',   // <-- usa EXACTAMENTE el nombre de tu repo
})
