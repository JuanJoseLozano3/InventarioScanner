# Inventario Web

Versión web responsive para teléfonos Android. Se publica como sitio estático en Vercel; no necesita API ni base de datos.

## Datos y privacidad

- El Excel `.xls` o `.xlsx` se lee dentro del navegador.
- El conteo se guarda en IndexedDB local, en el teléfono y navegador donde se abrió el sitio. Cerrar la pestaña y volver a entrar conserva la sesión.
- La sesión puede perderse si se borran los datos del sitio, se usa incógnito o se cambia de navegador o teléfono. Al terminar, exporta el CSV como copia.
- La foto se procesa en el teléfono. Tesseract descarga su modelo de idioma la primera vez que se usa OCR, así que esa primera lectura necesita conexión a Internet.

## Publicar en Vercel

1. Sube esta carpeta a un repositorio Git.
2. Importa el repositorio en Vercel.
3. Vercel detecta Vite. Los ajustes equivalentes están en `vercel.json` (`npm run build`, salida `dist`).
4. Abre la URL HTTPS en Chrome Android y permite la cámara.

La cámara requiere un contexto seguro HTTPS y permiso del usuario. Vercel entrega despliegues HTTPS.

## Desarrollo local

```sh
npm install
npm run dev -- --host 0.0.0.0
```

Para probar la cámara desde otro dispositivo, la página debe abrirse con HTTPS; `localhost` solo sirve para pruebas en el mismo equipo.
