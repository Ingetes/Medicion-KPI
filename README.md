# INGETES • KPI (Starter)

Proyecto base con **Vite + React + Tailwind + XLSX** para el portal de KPI.
Incluye el componente `ingetes_kpi_pipeline_mvp_web_app.jsx` listo para usar.

## Cómo ejecutar
```bash
npm install
npm run dev
```

## Build
```bash
npm run build
npm run preview
```

## Deploy con GitHub Pages (rama gh-pages)
1. Cambia `base` en `vite.config.js` por `'/NOMBRE-DEL-REPO/'`.
2. En `package.json`, puedes añadir `homepage` si quieres:
   `https://TU_USUARIO.github.io/NOMBRE-DEL-REPO`
3. Publica:
```bash
npm run deploy
```

## Deploy automático (GitHub Actions)
- Activa Pages con **GitHub Actions** en Settings → Pages.
- Cada push a `main` compilará y publicará.

## Notas
- Si deseas abrir `dist/index.html` sin servidor, usa `base: './'`.
- El portal carga archivos Excel RESUMEN y DETALLE y ofrece KPIs y fixtures para pruebas.
