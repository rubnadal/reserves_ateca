# i18n Audit (UI hardcoded strings)

Updated: 2026-03-12 (post Fases 0–5)
Scope: Cadenes visibles per l'usuari que NO utilitzen `window.t()`. S'exclouen logs de consola, comentaris, missatges interns i errors passats directament del servidor (`res.error`, `err.message`).

**Estat de la migració:**
- `scripts.html`: ~95% migrat ✅ — resta 1 `alert` i 1 text hardcoded menor
- `admin-scripts.html`: ~95% migrat ✅ — resten textos de loading d'alguns botons i 2–3 missatges de validació interns

---

## `scripts.html`

### Pendent

| Línia | Cadena hardcoded | Clau i18n suggerida |
|-------|-----------------|---------------------|
| 3231 | `alert('Error: ' + res.error)` (actualitzar incidència) | substituir per `showToast` amb `window.t('adminToast.genericError', {msg})` |

### Ja migrat ✅
- Tots els toasts principals (`toast.*`, `myReservations.*`, `incidencias.*`, `recurring.*`)
- Texts de les targetes de reserves individuals i recurrents
- Botons de cancel·lació de grup
- Optgroups del selector de recursos d'incidències
- Comptadors d'incidències actives/resoltes
- Tots els `toLocaleDateString` ara usen `window.getLocale()`

---

## `admin-scripts.html`

### Pendent

| Línia | Cadena hardcoded | Clau i18n suggerida |
|-------|-----------------|---------------------|
| 2965 | `` `Revisa los datos: ${errores.slice(0,3).join('. ')}` `` (validació usuaris) | `adminToast.usuarios.validationError` (amb `{errors}`) |
| 3848 | `` `Revisa: ${errores.slice(0,2).join('. ')}` `` (validació config) | `adminToast.config.validationError` (amb `{errors}`) |
| 4845 | `'Procesando...'` (botó "Aprovar totes" en recurrències) | `saveBar.processing` |
| 5048 | `` `${seleccionados} de ${total} tramos seleccionados` `` (historial recurrències) | `admin.historial.slotsCounter` (amb `{sel}`, `{total}`) |
| 5423 | `'Aprobando...'` (botó aprovar sol·licitud) | `adminModal.gestionar.approving` |
| 5475 | `'Rechazando...'` (botó rebutjar sol·licitud) | `adminModal.gestionar.rejecting` |
| 5651 | `'Selecciona un recurso'` (placeholder select modal recurrent directa) | `adminModal.recDirecta.resourcePlaceholder` |
| 5668 | `'Selecciona un usuario'` (placeholder select modal recurrent directa) | `adminModal.recDirecta.userPlaceholder` |
| 5877 | `'Generando...'` (botó crear reserves recurrents directes) | `adminModal.recDirecta.generating` |

### Ja migrat ✅
- Tots els toasts de Reserves, Recursos, Disponibilitat, Usuaris, Cursos, Trams, Configuració, Recurrències
- Tots els `mostrarConfirmacion()` (confirm.*)
- Tots els textos de `SaveBarManager`
- Estats buits de totes les taules (`emptyState`, `noResults`)
- Textos de la barra de capçalera dels modals
- Botó "Cancel·lar" de la columna Accions → Reserves

---

## Resum

| Fitxer | Pendent | Prioritat |
|--------|---------|-----------|
| `scripts.html` | 1 | Baixa — `alert` intern rar |
| `admin-scripts.html` | 9 | Baixa — textos de loading efímers + validacions internes |
| **Total** | **10** | |

## Notes

- Els textos de loading (`Procesando...`, `Aprobando...`, `Generando...`) són efímers i de molt baixa visibilitat — prioritat mínima.
- Els missatges de validació (línies 2965/3848) concatenen errors del servidor, difícilment 100% traduïbles sense refactor de la lògica de validació.
- Els placeholders de `<select>` del modal recurrent directa (línies 5651/5668) podrien reutilitzar claus `adminModal.recDirecta.*` ja existents si s'afegissin les sub-claus `resourcePlaceholder` i `userPlaceholder`.
- `Sidebar.html` i `ActivacionSistema.html` no tenen sistema i18n i no s'han migrat (molt baixa prioritat).
