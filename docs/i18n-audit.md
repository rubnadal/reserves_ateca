# i18n Audit (UI hardcoded strings)

Date: 2026-03-06
Scope: UI-only, no behavioral changes. Inventory only.
Sources: `scripts.html`, `admin-scripts.html`, `index.html`, `admin-panel.html`.

## Toasts in `scripts.html`
- Line ~3158: "Actualizando recurso..."
- Line ~3190: "Recurso ... bloqueado/activado correctamente"
- Line ~3193: "Error: ..."
- Line ~3198: "Error de conexion: ..."
- Line ~3526: "Nota guardada correctamente"
- Line ~3528: "Error al guardar"
- Line ~3534: "Error de conexion"
- Line ~3127: comment "Prioridad actualizada" (disabled)

## Toasts in `admin-scripts.html`
- Line ~694: "Reserva cancelada correctamente"
- Line ~701: "Error al cancelar"
- Line ~706: "Error de conexion: ..."
- Line ~1164: "Falta el nombre en X recurso(s)."
- Line ~1173: "Recursos guardados"
- Line ~1202: "Cambios descartados"
- Line ~1339: "Error al cargar datos"
- Line ~1368: "Error al cargar: ..."
- Line ~1521: "Solicitud no encontrada"
- Line ~1545: "Solicitud aprobada"
- Line ~1561: "Error al aprobar"
- Line ~1566: "Error: ..."
- Line ~1622: "X solicitud(es) aprobada(s)"
- Line ~1625: "X error(es) al aprobar"
- Line ~2104: "Error al eliminar tramo"
- Line ~2135: "Motivo actualizado"

## Hardcoded UI labels in `scripts.html` via textContent
- Line ~1153: "Cursos de ...:"
- Line ~1174: "Selecciona el curso:"
- Line ~1296: "Reservando..."
- Line ~1306: "Confirmar Reserva"
- Line ~1904: "Recurso"
- Line ~2097: "0 tramos seleccionados"
- Line ~2289: "Selecciona un recurso"

## Notes
- This is an inventory only; no translations applied.
- Next step: map each entry to a `data-i18n` key and add to `i18n-ca.html` and `i18n-es.html` in small batches.

