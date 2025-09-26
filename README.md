## Conversor Excel a VCF (Node + Express)

Aplicación web sencilla para:

- Subir archivos Excel (`.xlsx`, `.xls`) o CSV
- Listarlos y eliminarlos
- Generar un archivo `.vcf` indicando índice de hoja y columnas de teléfono y nombre, con opción de prefijo de país

### Requisitos

- Node.js 16+

### Instalación

```bash
npm install
```

Si aún no tienes el proyecto inicializado:

```bash
npm init -y
npm i express multer xlsx
```

### Ejecutar

```bash
npm run start
```

Abre `http://localhost:3000` en tu navegador.

### Estructura

```
.
├─ public/           # interfaz web (index.html)
├─ uploads/          # archivos subidos (se crea automáticamente)
├─ exports/          # vcf generados (se crea automáticamente)
├─ server.js         # servidor Express
├─ package.json
└─ README.md
```

### Uso de la interfaz

1. Sube un archivo `.xlsx`, `.xls` o `.csv`.
2. En el listado, ajusta:
   - Hoja (idx): índice 0-based de la hoja de Excel (0 = primera).
   - Tel (col idx): índice 0-based de la columna teléfono (ej. 1 equivale a columna 2).
   - Nombre (col idx): índice 0-based de la columna nombre (ej. 2 equivale a columna 3).
   - Header: auto/sí/no para indicar si la primera fila es encabezado.
   - Prefijo país: por ejemplo `+34`, `+52`, `+54`, etc.
3. Clic en "Generar VCF" para descargar el archivo `.vcf`.

### API (opcional)

- `GET /api/files`: lista archivos subidos.
- `POST /api/upload` (form-data `file`): sube un archivo.
- `DELETE /api/files/:name`: elimina un archivo por nombre.
- `POST /api/generate` (JSON):
  ```json
  {
    "file": "<nombre en uploads>",
    "sheetIndex": 0,
    "phoneCol": 1,
    "nameCol": 2,
    "hasHeader": true,
    "dialCode": "+34"
  }
  ```
  Respuesta exitosa:
  ```json
  {
    "ok": true,
    "file": "/exports/<archivo>.vcf",
    "count": 123
  }
  ```

### Notas

- La normalización de teléfono conserva `+` y dígitos, remueve separadores. Si no hay `+` y especificas `dialCode`, se antepone.
- El `.vcf` se genera en formato vCard 3.0 con campos `FN` y `TEL;TYPE=CELL`.

### Licencia

ISC


