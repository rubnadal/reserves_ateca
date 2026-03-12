# i18n Audit (UI hardcoded strings)

Updated: 2026-03-12
Scope: Cadenes visibles per l'usuari que NO utilitzen `window.t()`. S'exclouen logs, comentaris i missatges interns.

**Estat de la migració:**
- `scripts.html`: parcialment migrat (la majoria de toasts principals ✅, resten blocs específics ❌)
- `admin-scripts.html`: **0% migrat** — cap crida a `window.t()`. Tots els toasts i texts UI són hardcoded en castellà.

---

## `scripts.html`

### Toasts pendents

| Línia | Cadena hardcoded | Clau i18n suggerida |
|-------|-----------------|---------------------|
| 1829 | `'Cancelando reservas del grupo...'` (via `showInfo`) | `toast.cancelingGroup` |
| 1840 | `` `Se cancelaron ${n} reservas del grupo.` `` | `toast.groupCanceled` (amb `{n}`) |
| 1857 | `'Error al cancelar el grupo'` | `toast.errorCancelGroup` |
| 1862 | `'Error al cancelar: ' + msg` | `toast.errorCancelPrefix` (ja existent) |

### Texts UI (no toasts)

| Línia | Cadena hardcoded | Clau i18n suggerida |
|-------|-----------------|---------------------|
| 1003 | `'Confirmar'` (botó inside `handleTramoClick`) | `modalReservation.confirm` (ja existent) |
| 1482 | `'Recurrencias'` (separador visual) | `myReservations.recurringLabel` |
| 1504 | `'Recurrencia'` (text fallback) | `myReservations.recurringFallback` |
| 1557 | `'Cancelar toda la recurrencia'` (botó) | `myReservations.cancelGroup` |
| 1585 | `'Recurso eliminado'` (recurs no trobat) | `myReservations.deletedResource` |
| 1594 | `'Hora desconocida'` (tram no trobat) | `myReservations.unknownTime` |
| 1703 | `'Cancelando...'` (overlay de la targeta) | `myReservations.canceling` |
| 1807 | `'¿Cancelar TODAS las reservas...'` (confirm natiu, fallback) | `myReservations.confirmCancelGroup` |
| 2099–2100 | `'0 tramos seleccionados'` / `'1 tramo seleccionado'` / `'${n} tramos seleccionados'` | `recurring.slotsSelected` (amb `{n}`) |
| 2327 | `'-- Selecciona un recurso --'` (optgroup placeholder) | `incidencias.modalReport.resourcePlaceholder` (ja existent) |
| 2336 | `'🚪 Salas y Espacios'` (optgroup) | `sections.rooms` (ja existent) |
| 2349 | `'💻 Dispositivos y Carritos'` (optgroup) | `sections.devices` (ja existent) |
| 2619–2621 | `` `${n} activa${...} de ${t} total${...}` `` / `` `${t} incidencia${...} (todas resueltas ✅)` `` | `incidencias.counterActive` / `incidencias.counterResolved` |

### Locale hardcoded (`es-ES`)

Les següents crides a `toLocaleDateString` ignoren l'idioma de la UI:

| Línia | Context |
|-------|---------|
| 591 | Títol del calendari (mes/any) |
| 681 | Data seleccionada sota el calendari |
| 715 | Etiqueta de data als trams |
| 963 | Data al modal de confirmació de reserva |
| 1529 | Data a les targetes de recurrències |
| 1597 | Data a les targetes de reserves individuals |

---

## `admin-scripts.html`

> **Cap crida `window.t()` a tot el fitxer.** Tot el text UI és hardcoded.

### Reservas

| Línia | Cadena | Clau i18n suggerida |
|-------|--------|---------------------|
| 694 | `'Reserva cancelada correctamente'` | `adminToast.reservas.canceled` |
| 701 | `'Error al cancelar'` | `adminToast.reservas.errorCancel` |
| 706 | `'Error de conexión: ' + err` | `adminToast.connectionError` |
| 426 | `'No hay reservas registradas'` (empty state) | `admin.reservas.emptyState` |
| 546 | `'No hay coincidencias con la búsqueda'` (empty state) | `admin.reservas.noResults` |

### Recursos

| Línia | Cadena | Clau i18n suggerida |
|-------|--------|---------------------|
| 850 | `'✨ No hay recursos. Crea el primero arriba.'` (empty state) | `admin.recursos.emptyState` |
| 1164 | `` `Falta el nombre en ${n} recurso(s).` `` | `adminToast.recursos.missingName` |
| 1173 | `'Recursos guardados'` | `adminToast.recursos.saved` |
| 1202 | `'Cambios descartados'` | `adminToast.changesDiscarded` |
| 1248 | `'No hay recursos disponibles.'` (selector buit) | `admin.recursos.noneAvailable` |

### Disponibilitat

| Línia | Cadena | Clau i18n suggerida |
|-------|--------|---------------------|
| 1339 | `'Error al cargar datos'` | `adminToast.errorLoadData` |
| 1368 | `'Error al cargar: ' + msg` | `adminToast.errorLoad` |
| 1521 | `'Solicitud no encontrada'` | `adminToast.disp.requestNotFound` |
| 1545 | `'Solicitud aprobada'` | `adminToast.disp.requestApproved` |
| 1561 | `'Error al aprobar'` | `adminToast.disp.errorApprove` |
| 1566 | `'Error: ' + msg` | `adminToast.genericError` |
| 1596 | `` `¿Aprobar las ${n} solicitudes pendientes?` `` (confirm natiu) | `admin.disp.confirmApproveAll` |
| 1622 | `` `${n} solicitud(es) aprobada(s)` `` | `adminToast.disp.batchApproved` |
| 1625 | `` `${n} error(es) al aprobar` `` | `adminToast.disp.batchErrors` |
| 2124/2144 | `'Guardando...'` / `'Guardar'` (botó) | `adminToast.disp.savingMotivo` / `admin.disp.saveMotivo` |
| 2135 | `'Motivo actualizado'` | `adminToast.disp.motivoSaved` |
| 2140/2147 | `'Error al guardar'` / `'Error: ' + msg` | `adminToast.errorSave` / `adminToast.genericError` |
| 2174/2214 | `'Revocando...'` / `'Revocar'` (botó) | `adminToast.disp.revoking` / `admin.disp.revoke` |
| 2185 | `'Recurrencia revocada'` | `adminToast.disp.recurringRevoked` |
| 2210 | `'Error al revocar'` | `adminToast.disp.errorRevoke` |
| 2240 | `'Solicitud aprobada correctamente'` | `adminToast.disp.requestApprovedOk` |
| 2338 | `'Error: No hay recurso seleccionado'` | `adminToast.disp.noResource` |
| 2364 | `'Disponibilidad guardada'` | `adminToast.disp.dispSaved` |
| 2371/2500 | `'Error al guardar'` | `adminToast.errorSave` |
| 2494 | `'Disponibilidad guardada y usuarios notificados'` | `adminToast.disp.dispSavedNotified` |
| 2506 | `'Error: ' + err` | `adminToast.genericError` |
| 2516 | `'Cambios revertidos'` | `adminToast.changesReverted` |

### Usuaris

| Línia | Cadena | Clau i18n suggerida |
|-------|--------|---------------------|
| 2750 | `'No se encontraron usuarios'` (empty state) | `admin.usuarios.emptyState` |
| 2860/2865 | `'Admin'`/`'Usuario'`, `'Activo'`/`'Inactivo'` (toggles) | `admin.usuarios.roleAdmin` etc. |
| 2964 | `'Usuarios actualizados'` | `adminToast.usuarios.saved` |
| 2983 | `'Cambios descartados'` | `adminToast.changesDiscarded` |

### Cursos

| Línia | Cadena | Clau i18n suggerida |
|-------|--------|---------------------|
| 3044 | `'No hay cursos'` (empty state) | `admin.cursos.emptyState` |
| 3221 | `'Hay cursos sin nombre. Por favor, revísalos (marcados en rojo).'` | `adminToast.cursos.missingName` |
| 3237 | `'Cursos actualizados correctamente'` | `adminToast.cursos.saved` |
| 3273 | `'Cambios descartados'` | `adminToast.changesDiscarded` |

### Trams

| Línia | Cadena | Clau i18n suggerida |
|-------|--------|---------------------|
| 3372 | `'No hay tramos definidos'` (empty state) | `admin.tramos.emptyState` |
| 2104 | `'Error al eliminar tramo'` | `adminToast.tramos.errorDelete` |
| 2110 | `'Error: ' + msg` | `adminToast.genericError` |
| 3572 | `'Hay tramos con datos incompletos (en rojo). Revísalos.'` | `adminToast.tramos.missingData` |
| 3592 | `'Tramos guardados correctamente'` | `adminToast.tramos.saved` |
| 3622 | `'Cambios en tramos descartados'` | `adminToast.changesDiscarded` |

### Configuració

| Línia | Cadena | Clau i18n suggerida |
|-------|--------|---------------------|
| 3849 | `'Permisos de imagen actualizados automáticamente'` | `adminToast.config.drivePermissions` |
| 3871 | `'Configuración guardada'` | `adminToast.config.saved` |
| 3923 | `'Cambios descartados'` | `adminToast.changesDiscarded` |

### Global / Càrrega de dades

| Línia | Cadena | Clau i18n suggerida |
|-------|--------|---------------------|
| 4199 | `'No se encontraron iconos'` (cerca d'icones) | `admin.icons.noResults` |
| 4352 | `'Error al cargar datos'` / `'Error servidor'` | `adminToast.errorLoadData` |
| 4422 | `'Error de conexión: ' + msg` | `adminToast.connectionError` |

### Recurrències (panell d'admin)

| Línia | Cadena | Clau i18n suggerida |
|-------|--------|---------------------|
| 4645 | `'No hay solicitudes pendientes'` | `adminToast.rec.noPending` |
| 4812 | `'No hay solicitudes pendientes para aprobar'` | `adminToast.rec.noPendingApprove` |
| 4854 | `` `${n} solicitud(es) aprobada(s). ${r} reservas generadas.` `` | `adminToast.rec.batchApproved` |
| 4856 | `` `Aprobadas: ${n}, Errores: ${e}. Revisa los detalles.` `` | `adminToast.rec.batchPartial` |
| 4869/4908/5069/5326/5486 | `'Solicitud no encontrada'` | `adminToast.rec.notFound` |
| 4877 | `'Cancelando recurrencia...'` | `adminToast.rec.canceling` |
| 4888 | `'Recurrencia cancelada'` | `adminToast.rec.canceled` |
| 4892/4896 | `'Error al cancelar'` / `'Error: ' + msg` | `adminToast.rec.errorCancel` |
| 4926 | `'Esta recurrencia no tiene tramos editables'` | `adminToast.rec.noEditableSlots` |
| 5084 | `'No hay cambios para guardar'` | `adminToast.rec.noChanges` |
| 5120 | `'Tramos actualizados correctamente'` | `adminToast.rec.slotsUpdated` |
| 5124/5129 | `'Error al actualizar tramos'` / `'Error: ' + msg` | `adminToast.rec.errorUpdateSlots` |
| 5394/5447 | `'Error interno: botón no encontrado'` | `adminToast.internalError` |
| 5422/5426 | `'Error al aprobar'` / `'Error al aprobar la solicitud: ' + msg` | `adminToast.rec.errorApprove` |
| 5441 | `'Debes indicar un motivo para rechazar'` | `adminToast.rec.rejectReasonRequired` |
| 5464 | `'Solicitud rechazada correctamente'` | `adminToast.rec.rejected` |
| 5468/5474 | `'Error al rechazar'` / `'Error al rechazar la solicitud: ' + msg` | `adminToast.rec.errorReject` |

### Recurrència directa (admin)

| Línia | Cadena | Clau i18n suggerida |
|-------|--------|---------------------|
| 5838 | `'Indica la fecha de fin'` | `adminToast.rec.direct.noEndDate` |
| 5847 | `'Selecciona un recurso'` | `adminToast.rec.direct.noResource` |
| 5848 | `'Selecciona un usuario'` | `adminToast.rec.direct.noUser` |
| 5849 | `'Selecciona al menos un tramo en el calendario'` | `adminToast.rec.direct.noSlots` |
| 5882/5886 | `'Error al crear reservas'` / `'Error: ' + msg` | `adminToast.rec.direct.errorCreate` |

---

## Resum

| Fitxer | Toasts pendents | Texts UI pendents | Totals |
|--------|----------------|-------------------|--------|
| `scripts.html` | 4 | 13 + 6 locales | ~23 |
| `admin-scripts.html` | ~65 | ~15 | ~80 |
| **Total** | | | **~103** |

## Notes per a la migració

- Les claus d'admin podrien agrupar-se en un nou namespace `adminToast.*` als diccionaris.
- Moltes cadenes d'error genèriques es repeteixen (`'Error: ' + msg`, `'Error al guardar'`, `'Cambios descartados'`) → candidats a reutilitzar claus existents o crear-ne una de genèrica.
- Les cadenes pluralitzades (e.g. `solicitud(es)`, `${n} tramos`) requeriran una estratègia de pluralització (funció auxiliar o claus separades per 0/1/n).
- Els `confirm()` natius (línies 1596 i 1807) haurien de migrar a modals personalitzats amb `window.t()`.
- Els `toLocaleDateString('es-ES', ...)` de `scripts.html` haurien de rebre el locale dinàmic (`window.APP_LANG === 'ca' ? 'ca-ES' : 'es-ES'`).
