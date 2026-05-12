# SMTO — Reporte de Gastos

Aplicación web para gestión y conciliación de facturas CFDI (XML) con estados de cuenta bancarios.

## Funcionalidades

- 📂 **Cargar Carpeta** — Importa todos los XML de CFDI de una carpeta (y sus PDFs asociados)
- 🏦 **Validar Banco** — Concilia facturas contra un estado de cuenta CSV
- ✚ **Manual** — Agrega filas de gastos manuales (tickets sin factura)
- 📋 **Copiar a Excel** — Copia la tabla en formato TSV listo para pegar en Excel
- 📦 **Exportar a ZIP** — Genera un ZIP con el CSV + facturas renombradas

## Cómo desplegar

### Opción 1: Vercel (recomendado)

1. Sube este proyecto a un repositorio en GitHub
2. Ve a [vercel.com](https://vercel.com) → **Add New Project**
3. Importa el repositorio desde GitHub
4. Vercel detecta Vite automáticamente — haz clic en **Deploy**

### Opción 2: Local

```bash
npm install
npm run dev
```

Abre [http://localhost:5173](http://localhost:5173)

### Build de producción

```bash
npm run build
# Los archivos quedan en /dist listo para cualquier hosting estático
```

## Compatibilidad

La carga de carpetas (`webkitdirectory`) funciona en:
- ✅ Chrome / Edge (recomendado)
- ✅ Firefox
- ✅ Safari 14+

## Tecnologías

- [React 18](https://react.dev/)
- [Vite 5](https://vitejs.dev/)
- [JSZip](https://stuk.github.io/jszip/) — para generar el ZIP de exportación
