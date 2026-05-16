# Web Control Policial

Dashboard React + Tailwind para control policial operativo del SubTE de Buenos Aires, con estética de centro de monitoreo EMOVA/SubTE.

## Ejecutar

```bash
npm install
npm run dev
```

Luego abrir la URL que muestre Vite, normalmente `http://127.0.0.1:5173`.

## Funciones

- Dashboard oscuro responsive.
- Tarjetas estadísticas de efectivos, líneas, observaciones e incidentes críticos.
- Tabla de registros con filtros por línea, estado y turno.
- Buscador global por estación, fuerza, puesto o novedad.
- Importación de Excel `.xlsx`, `.xls` o `.csv`.
- Normalización de planillas con varias hojas por día.

## Importación Excel

El importador detecta columnas habituales como fecha, hora, línea, estación, turno, dependencia, dotación, puesto, estado y novedades. Si la planilla tiene formato visual en lugar de tabla plana, intenta convertir filas con contenido en registros operativos usando la fecha de la hoja.

## Actualización automática desde Excel

Para que la web se actualice cuando cambie el Excel o cuando agregues uno nuevo a la carpeta, dejar corriendo:

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\watch-excel.ps1
```

Ese proceso vigila todos los `.xlsx/.xls` de la carpeta, ignora temporales `~$...`, usa automáticamente el Excel más reciente y regenera:

- `control-policial-data.js`
- `src/controlData.js`

La vista `preview.html` se recarga cada 30 segundos para tomar los cambios nuevos.
