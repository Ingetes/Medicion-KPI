export default defineConfig({
  base: '/Medicion-KPI/',
  build: {
    sourcemap: false,      // evita subir mapas (pesan)
    cssCodeSplit: true,
    brotliSize: false,     // evita cálculo de tamaños
    chunkSizeWarningLimit: 1200
  }
})
