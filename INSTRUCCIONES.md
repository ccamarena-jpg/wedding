# 💍 Organizador de Boda — Claudia & Jorge

Web app conectada a **Google Sheets** (la hoja es la base de datos) y a **Google Calendar**
(para las citas/atenciones de proveedores). Diseño inspirado en tu moodboard: verde oliva,
crema marfil y dorado.

## Qué incluye
- **Inicio / Resumen** — días para la boda, progreso, KPIs y alertas (vista general).
- **Presupuesto** — tope total, estimado vs. cotizado vs. pagado por categoría.
- **Proveedores** — contactos, estado, cotizado y pagado (conectado al presupuesto).
- **Invitados y mesas** — personas, mesa, confirmación e intolerancias.
- **Cronograma del día** — línea de tiempo por bloques (mañana/tarde/noche).
- **Citas & Gantt** — diagrama de Gantt de atenciones + **sincronización con Google Calendar**.

---

## Instalación (una sola vez, ~5 min)

1. Entrá a **[sheets.new](https://sheets.new)** para crear una hoja nueva (o abrí una existente).
   Ponele nombre, p. ej. *"Boda Claudia y Jorge"*.
2. Menú **Extensiones ▸ Apps Script**. Se abre el editor.
3. Borrá el contenido de `Código.gs` y **pegá todo el archivo [`Code.gs`](Code.gs)**.
4. Clic en **＋ ▸ HTML**, nombralo exactamente **`Index`** (sin `.html`).
   Borrá su contenido y **pegá todo [`Index.html`](Index.html)**.
5. *(Opcional pero recomendado)* Clic en el ⚙️ **Configuración del proyecto** y activá
   *"Mostrar el archivo de manifiesto appsscript.json"*. Abrí `appsscript.json`,
   borralo y pegá el de esta carpeta (fija zona horaria Lima y permisos).
6. Arriba, elegí la función **`setup`** y presioná **▶ Ejecutar**.
   Google te va a pedir **autorizar permisos** (Sheets + Calendar) → aceptá con tu cuenta.
   Esto crea las pestañas y carga datos de ejemplo.
7. Botón **Implementar ▸ Nueva implementación ▸ Aplicación web**.
   - *Ejecutar como:* **Yo**
   - *Quién tiene acceso:* **Solo yo** (o "Cualquiera con el vínculo" si querés compartirla con Jorge).
   - **Implementar** → copiá la **URL de la app web**. Esa es tu organizador. 🎉

> Cada vez que edites datos en la web, se guardan en las pestañas del Sheet.
> Cada cita que "Sincronizás" crea/actualiza un evento en tu Google Calendar
> (con recordatorio 1 día antes).

---

## Cómo se guardan los datos (pestañas del Sheet)

| Pestaña | Contenido |
|---|---|
| `Config` | Pareja, fecha, lugar, moneda, presupuesto total |
| `Presupuesto` | Categorías y montos estimados |
| `Proveedores` | Nombre, categoría, estado, contacto, cotizado, pagado |
| `Invitados` | Invitación, grupo, lado, personas, mesa, confirmación, restricciones |
| `Cronograma` | Bloque, hora, duración, título, lugar, responsable |
| `Citas` | Proveedor, título, inicio, fin, estado, notas, ID del evento de Calendar |

Podés editar directamente en el Sheet **o** desde la web; siempre están sincronizados.

---

## Notas
- **Moneda:** PEN (S/). Se puede cambiar desde *Inicio ▸ Editar datos*.
- **Vista previa del diseño:** abrí `Index.html` directo en el navegador para ver el diseño
  con datos de ejemplo (en ese modo no toca tu Sheet/Calendar; sirve para mostrar cómo se ve).
- **Compartir con tu pareja:** en el paso 7 elegí "Cualquiera con el vínculo" y pasale la URL.
- **Actualizar el código:** si cambiás algo, *Implementar ▸ Gestionar implementaciones ▸
  editar (lápiz) ▸ Versión nueva*.
- **Permisos:** el script solo accede a *esta* hoja (`spreadsheets.currentonly`) y a tu calendario.

---

## Estructura de archivos de esta carpeta
```
wedding-planner/
├─ Code.gs           → backend (pegar en Apps Script)
├─ Index.html        → interfaz web (pegar como archivo HTML "Index")
├─ appsscript.json   → manifiesto (zona horaria + permisos)
└─ INSTRUCCIONES.md  → este archivo
```
