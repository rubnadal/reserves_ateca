# i18n Audit (UI hardcoded strings)

Date: 2026-03-08
Scope: UI-only, no behavioral changes. Inventory only.
Sources: `scripts.html`, `admin-scripts.html`.

## Toasts pending in `scripts.html`
- Line ~1840: "Se cancelaron X reservas del grupo."
- Line ~1857: "Error al cancelar el grupo"
- Line ~1862: "Error al cancelar: ..."
- Line ~1829 (info banner): "Cancelando reservas del grupo..." (uses `showInfo`, not `showToast`)

## Toasts pending in `admin-scripts.html`
- Line ~694: "Reserva cancelada correctamente"
- Line ~701: "Error al cancelar"
- Line ~706: "Error de conexión: ..."
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
- Line ~2110: "Error: ..."

## Notes
- This is an inventory only; no translations applied.
- Other UI hardcoded strings may exist outside toast flows and are not included in this pass.
- Next step: map each entry to a `data-i18n` or `window.t(...)` key and add to `i18n-ca.html` and `i18n-es.html` in small batches.
