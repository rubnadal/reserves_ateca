# i18n Audit (UI hardcoded strings)

Updated: 2026-03-13 (post Fases 0–6: migració completa)
Scope: Cadenes visibles per l'usuari que NO utilitzen `window.t()`. S'exclouen logs de consola, comentaris, missatges interns i errors passats directament del servidor (`res.error`, `err.message`).

**Estat de la migració:**
- `scripts.html`: 100% migrat ✅
- `admin-scripts.html`: 100% migrat ✅

---

## `scripts.html`

### Ja migrat ✅
- Tots els toasts principals (`toast.*`, `myReservations.*`, `incidencias.*`, `recurring.*`)
- Texts de les targetes de reserves individuals i recurrents
- Botons de cancel·lació de grup
- Optgroups del selector de recursos d'incidències
- Comptadors d'incidències actives/resoltes
- Tots els `toLocaleDateString` ara usen `window.getLocale()`
- `alert('Error: ...')` → `showToast` amb `window.t('adminToast.genericError', {msg})`

---

## `admin-scripts.html`

### Ja migrat ✅
- Tots els toasts de Reserves, Recursos, Disponibilitat, Usuaris, Cursos, Trams, Configuració, Recurrències
- Tots els `mostrarConfirmacion()` (confirm.*)
- Tots els textos de `SaveBarManager` (incl. `saveBar.processing`)
- Estats buits de totes les taules (`emptyState`, `noResults`, `admin.reservas.noResultsPeriod`)
- Textos de la barra de capçalera dels modals
- Botó "Cancel·lar" de la columna Accions → Reserves (`admin.reservas.cancelBtn`)
- Validació usuaris: `adminToast.usuarios.validationError`
- Validació config: `adminToast.config.validationError`
- Comptador trams historial recurrències: `admin.historial.slotsCounter`
- Botó aprovar sol·licitud: `adminModal.gestionar.approving`
- Botó rebutjar sol·licitud: `adminModal.gestionar.rejecting`
- Placeholder select recurs modal recurrent: `adminModal.recDirecta.resourcePlaceholder`
- Placeholder select usuari modal recurrent: `adminModal.recDirecta.userPlaceholder`
- Botó generar reserves recurrents: `adminModal.recDirecta.generating`

---

## Resum

| Fitxer | Pendent | Estat |
|--------|---------|-------|
| `scripts.html` | 0 | ✅ Complet |
| `admin-scripts.html` | 0 | ✅ Complet |

## Fora d'abast

- `Sidebar.html` i `ActivacionSistema.html`: no tenen sistema i18n (molt baixa prioritat, no estan en el flux principal de l'aplicació).
- Missatges d'error del servidor (`res.error`, `err.message`): passen directament del backend, no traduïbles sense refactor del backend.
