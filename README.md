# 💍 Nuestra Boda · Organizador — Claudia & Jorge

Web app para organizar la boda, **conectada a Google Sheets** (la hoja es la base de datos)
y a **Google Calendar** (para las citas/atenciones de proveedores). Construida como
**Google Apps Script** y con una paleta tomada del moodboard de la boda: marfil cálido,
taupe champagne, espresso y salvia.

📅 **Fecha:** sábado 12 de diciembre de 2026 · Lima, Perú

## Módulos
- **Inicio / Resumen** — días para la boda, progreso, KPIs y alertas.
- **Presupuesto** — tope total, estimado vs. cotizado vs. pagado por categoría (PEN, TC S/3.40).
- **Proveedores** — contactos, estado, cotizado y pagado (conectado al presupuesto).
- **Invitados y mesas** — lista, grupos, confirmación, restricciones y asignación de mesas.
- **Cronograma del día** — línea de tiempo por bloques (mañana/tarde/noche).
- **Citas & Gantt** — diagrama de Gantt + sincronización con Google Calendar.

## Archivos
| Archivo | Descripción |
|---|---|
| [`Code.gs`](Code.gs) | Backend Apps Script (Sheets + Calendar) |
| [`Index.html`](Index.html) | Interfaz web (SPA); también corre standalone con datos de ejemplo |
| [`appsscript.json`](appsscript.json) | Manifiesto (zona horaria Lima + permisos) |
| [`INSTRUCCIONES.md`](INSTRUCCIONES.md) | Guía de instalación paso a paso |

## Instalación
Ver **[INSTRUCCIONES.md](INSTRUCCIONES.md)**. En resumen: crear un Google Sheet →
Extensiones ▸ Apps Script → pegar `Code.gs` y `Index.html` → ejecutar `setup()` →
Implementar como aplicación web.

## Vista previa del diseño
Abrí `Index.html` directo en el navegador (o serví la carpeta con
`python -m http.server`) para ver el diseño con datos de ejemplo, sin tocar el Sheet.
