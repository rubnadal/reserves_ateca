# Sistema de Reserves ATECA

[![License: CC BY-NC-SA 4.0](https://img.shields.io/badge/License-CC%20BY--NC--SA%204.0-lightgrey.svg)](https://creativecommons.org/licenses/by-nc-sa/4.0/)
[![Google Apps Script](https://img.shields.io/badge/Google%20Apps%20Script-4285F4?logo=google&logoColor=white)](https://script.google.com/)
[![Tailwind CSS](https://img.shields.io/badge/Tailwind%20CSS-38B2AC?logo=tailwind-css&logoColor=white)](https://tailwindcss.com/)

Sistema de gestió de reserves de recursos educatius per a l'aula ATECA (Aula de Tecnologia Aplicada).

## Estat del projecte

Aquest projecte està en procés de modificació i en fase de proves per aconseguir que la interfície sigui multiidioma (ES/CA).

## Control de canvis

| Canvi                                                     | Estat |
|-----------------------------------------------------------|-------|
| Internacionalització UI: pantalla principal               | ✅     |
| Internacionalització UI: Admin > Recursos                 | ✅     | 
| Internacionalització UI: Admin > Reserves                 | ✅     | 
| Internacionalització UI: Admin > Trams                    | ✅     | 
| Internacionalització UI: Admin > Disponibilitat           | ✅     | 
| Internacionalització UI: Admin > Cursos                   | ✅     | 
| Internacionalització UI: Admin > Usuaris                  | ✅     | 
| Internacionalització UI: Admin > Configuració del sistema | ✅     | 
| Diccionari i18n ES/CA i helpers UI                        | ✅     |
| Internacionalització UI: gestió d'incidències             | ✅     |
| Internacionalització UI: modal Reserves                   | ✅     |
| Internacionalització UI: missatges Toast                  | 🔄    |


## Sobre aquest projecte

Aquest projecte és un **fork adaptat** del sistema original de reserves creat per [@maestroseb](https://github.com/maestroseb/reservasrecursosysalas). El projecte original proporciona una solució completa per a la gestió de reserves de recursos i sales en centres educatius.

### Adaptacions per a ATECA

Aquest fork ha estat adaptat específicament per gestionar les reserves d'espais i recursos dins de l'aula ATECA de l'[Institut Eugeni d'Ors](https://www.ies-eugeni.cat) de Vilafranca del Penedès, optimitzant-lo per a les necessitats d'aquest entorn.
És, per tant, una solució personalitzada que manté les funcionalitats principals del projecte original, però amb modificacions específiques per a l'aula ATECA, com ara la internacionalització de la interfície i l'adaptació dels recursos i trams horaris disponibles.

## Característiques principals

- Gestió de reserves per franges horàries.
- Multirecurs: sales, carros de portàtils i equipament.
- Control d'accés amb autorització.
- Reserves recurrents setmanals amb aprovació administrativa.
- Sistema d'incidències per reportar problemes.
- Panell d'administració complet.
- Notificacions per correu electrònic.
- Disseny adaptatiu.
- Sistema de memòria cau per optimitzar rendiment.
- Multidioma (ES/CA) en procés de desenvolupament.

## Tecnologies utilitzades

**Backend**
- Google Apps Script (JavaScript).
- Google Sheets (base de dades).
- Gmail API (notificacions).
- Google Workspace (autenticació SSO).

**Frontend**
- HTML5.
- Tailwind CSS 3.x.
- JavaScript (vanilla).
- Iconify (Material Design Icons).

## Estructura del projecte

```
├── Codigo.gs                 # Motor principal de l'aplicació
├── AdminFunctions.gs         # Funcions d'administració
├── Setup.gs                  # Instal·lador inicial
├── ReservasRecurrentes.gs    # Gestió de reserves recurrents
├── Incidencias.gs            # Sistema d'incidències
├── index.html                # Interfície principal
├── admin-panel.html          # Panell d'administració
├── styles.html               # Estils CSS
├── scripts.html              # JavaScript client
├── admin-scripts.html        # JavaScript admin
├── i18n.html                 # Diccionari i18n UI (ES/CA)
├── registro.html             # Formulari de registre
├── ActivacionSistema.html    # Assistència d'instal·lació
├── Sidebar.html              # Instruccions
└── appsscript.json           # Configuració del projecte
```

## Ús

**Usuaris**
- Selecció del recurs, data i franja horària disponible.
- Confirmació de la reserva.

**Administradors**
- Gestió d'usuaris, recursos, trams horaris i reserves.
- Configuració de paràmetres del sistema.

## Crèdits i autoria

**Projecte original**
- Autor: [Sebastián Giraldo(maestroseb)](https://github.com/maestroseb)
- Repositori: [reservasrecursosysalas](https://github.com/maestroseb/reservasrecursosysalas)

**Aquest fork**
- Adaptat per a l'aula ATECA amb modificacions específiques.

## Llicència

Aquest projecte està llicenciat sota **CC BY-NC-SA 4.0**.

Per a més informació: https://creativecommons.org/licenses/by-nc-sa/4.0/

## Reportar problemes

Si detectes errors o suggeriments de millora, obre un issue amb una descripció clara i, si és possible, captures de pantalla.
