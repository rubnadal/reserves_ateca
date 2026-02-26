# 🗓️ Sistema de Reserves ATECA

[![License: CC BY-NC-SA 4.0](https://img.shields.io/badge/License-CC%20BY--NC--SA%204.0-lightgrey.svg)](https://creativecommons.org/licenses/by-nc-sa/4.0/)
[![Google Apps Script](https://img.shields.io/badge/Google%20Apps%20Script-4285F4?logo=google&logoColor=white)](https://script.google.com/)
[![Made with Tailwind CSS](https://img.shields.io/badge/Tailwind%20CSS-38B2AC?logo=tailwind-css&logoColor=white)](https://tailwindcss.com/)

Sistema de gestió de reserves de recursos educatius per a l'aula ATECA (Aula de Tecnologia Aplicada).

## ⚠️ Estat del projecte

Aquest projecte està en procés de modificació i en fase de proves per aconseguir que la interfície sigui multi-idioma (ES/CA).

## 📋 Sobre aquest projecte

Aquest projecte és un **fork adaptat** del sistema original de reserves creat per [**@maestroseb**](https://github.com/maestroseb/reservasrecursosysalas). El projecte original proporciona una solució completa per a la gestió de reserves de recursos i sales en centres educatius.

### Adaptacions per a ATECA

Aquest fork ha estat adaptat específicament per gestionar les reserves d'espais i recursos dins d'una **aula ATECA** (Aula de Tecnologia Aplicada) d'un centre educatiu, optimitzant-lo per a les necessitats específiques d'aquest entorn.

## ✨ Característiques principals

- 📅 **Gestió de reserves** per franges horàries
- 🏫 **Multi-recurs**: Sales, carros de portàtils, equips informàtics
- 👥 **Control d'accés** amb sistema d'autorització
- 🔁 **Reserves recurrents** setmanals amb aprovació administrativa
- 🚨 **Sistema d'incidències** per reportar problemes
- 👨‍💼 **Panel d'administració** complet
- 📧 **Notificacions per email** automàtiques
- 📱 **Disseny responsive** (mòbil, tauleta, escriptori)
- ⚡ **Sistema de caché** per optimitzar rendiment
- 🎨 **Interfície moderna** amb Tailwind CSS

## 🛠️ Tecnologies utilitzades

### Backend
- **Google Apps Script** (JavaScript)
- **Google Sheets** (base de dades)
- **Gmail API** (notificacions)
- **Google Workspace** (autenticació SSO)

### Frontend
- **HTML5**
- **Tailwind CSS 3.x**
- **JavaScript Vanilla**
- **Iconify** (Material Design Icons)

## 📦 Estructura del projecte

```
├── Codigo.gs                 # Motor principal de l'aplicació
├── AdminFunctions.gs         # Funcions d'administració
├── Setup.gs                  # Instal·lador inicial
├── ReservasRecurrentes.gs    # Gestió de reserves recurrents
├── Incidencias.gs            # Sistema d'incidències
├── index.html                # Interfície principal
├── admin-panel.html          # Panel d'administració
├── styles.html               # Estils CSS
├── scripts.html              # JavaScript client
├── admin-scripts.html        # JavaScript admin
├── registro.html             # Formulari de registre
├── ActivacionSistema.html    # Wizard d'instal·lació
├── Sidebar.html              # Instruccions
└── appsscript.json           # Configuració del projecte
```

## 🚀 Instal·lació

### Prerequisits
- Compte de Google Workspace
- Permisos per crear Google Sheets i desplegaments web

### Passos

1. **Crear una còpia del Google Sheet**
   - Crea un nou Google Sheet al teu Drive
   - Obre l'editor d'scripts: `Extensions > Apps Script`

2. **Copiar el codi**
   - Copia tots els fitxers `.gs` i `.html` al projecte
   - Assegura't que `appsscript.json` estigui configurat correctament

3. **Desplegar com a Web App**
   - Al menú de l'editor: `Deploy > New deployment`
   - Tipus: **Web App**
   - Execute as: **User accessing the web app**
   - Who has access: **Anyone within [el teu domini]**
   - Copia la URL generada

4. **Activar el sistema**
   - Accedeix a la URL de la Web App
   - El wizard d'instal·lació es llançarà automàticament
   - Fes clic a "Inicializar Sistema"
   - El sistema crearà totes les pestanyes necessàries i dades d'exemple

5. **Configurar l'accés**
   - El teu usuari serà automàticament administrador
   - Configura els recursos, tramos horaris i cursos segons les teves necessitats

## 📖 Ús

### Per a usuaris

1. Accedeix a la URL de l'aplicació
2. Si és la primera vegada, omple el formulari de registre
3. Espera l'aprovació de l'administrador
4. Un cop aprovat:
   - Selecciona un recurs (sala, carro, etc.)
   - Tria una data al calendari
   - Selecciona una franja horària disponible
   - Confirma la reserva

### Per a administradors

Accedeix al **Panel Admin** per:
- ✅ Aprovar/rebutjar usuaris nous
- 📦 Gestionar recursos (afegir, editar, eliminar)
- ⏰ Configurar tramos horaris
- 🚫 Bloquejar disponibilitat per manteniment
- 📊 Veure totes les reserves
- 🔧 Configurar paràmetres del sistema
- 🚨 Gestionar incidències reportades
- 🔁 Aprovar reserves recurrents

## 🔧 Configuració

Al panel d'administració, secció **Config**, pots configurar:

- `dias_vista_maximo`: Dies futurs disponibles per reservar
- `limite_reservas`: Màxim de reserves actives per usuari
- `horas_cancelacion`: Hores mínimes per cancel·lar
- `exigir_motivo`: Obligar a especificar motiu de la reserva
- `modo_mantenimiento`: Bloquejar noves reserves temporalment
- `permitir_multitramo`: Permetre reserves de múltiples franges consecutives

## 📄 Model de dades

El sistema utilitza 8 pestanyes al Google Sheet:

| Pestanya | Funció |
|----------|--------|
| **Recursos** | Sales, carros, equipament |
| **Tramos** | Franges horàries |
| **Disponibilidad** | Control de disponibilitat per dia/hora |
| **Reservas** | Registre de reserves |
| **Usuarios** | Control d'accés |
| **Cursos** | Cursos acadèmics |
| **Incidencias** | Reportes de problemes |
| **SolicitudesRecurrentes** | Peticions de reserves setmanals |
| **Config** | Paràmetres del sistema |

## 🤝 Crèdits

### Projecte Original
- **Autor**: [Sebastián (maestroseb)](https://github.com/maestroseb)
- **Repository**: [reservasrecursosysalas](https://github.com/maestroseb/reservasrecursosysalas)

Aquest projecte no seria possible sense l'excel·lent treball de l'autor original. Gràcies per compartir el teu codi amb la comunitat educativa! 🙏

### Aquest Fork
Adaptat per a l'aula ATECA amb modificacions específiques per a l'entorn del nostre centre educatiu.

## 📜 Llicència

Aquest projecte està llicenciat sota la llicència **Creative Commons Attribution-NonCommercial-ShareAlike 4.0 International (CC BY-NC-SA 4.0)**.

[![CC BY-NC-SA 4.0](https://licensebuttons.net/l/by-nc-sa/4.0/88x31.png)](https://creativecommons.org/licenses/by-nc-sa/4.0/)

### Això significa que pots:

✅ **Compartir** — copiar i redistribuir el material en qualsevol mitjà o format  
✅ **Adaptar** — remesclar, transformar i crear a partir del material

### Sota les següents condicions:

📝 **Atribució (BY)** — Has de donar crèdit apropiat, proporcionar un enllaç a la llicència i indicar si s'han realitzat canvis.

🚫 **No Comercial (NC)** — No pots utilitzar el material amb finalitats comercials.

🔄 **Compartir Igual (SA)** — Si remescles, transformes o crees a partir del material, has de distribuir les teves contribucions sota la mateixa llicència que l'original.

Per més informació, consulta el text complet de la llicència: [CC BY-NC-SA 4.0](https://creativecommons.org/licenses/by-nc-sa/4.0/legalcode)

## 🐛 Reportar problemes

Si trobes algun error o tens suggeriments de millora, si us plau:
1. Comprova que no estigui ja reportat
2. Obre un **Issue** amb una descripció detallada
3. Inclou captures de pantalla si és possible

## 💡 Contribucions

Les contribucions són benvingudes! Si vols millorar aquest projecte:

1. Fes un fork del repositori
2. Crea una branca per a la teva característica (`git checkout -b feature/nova-caracteristica`)
3. Fes commit dels teus canvis (`git commit -am 'Afegida nova característica'`)
4. Puja la branca (`git push origin feature/nova-caracteristica`)
5. Obre un Pull Request

## ❓ Preguntes freqüents (FAQ)

### Com puc canviar el logo del sistema?
Al panel Admin > Config, modifica el valor de `url_logo` amb la URL de la teva imatge.

### Els usuaris no reben emails de confirmació
Comprova que el script tingui permisos per enviar emails i que l'adreça `email_admin` estigui configurada correctament.

### Com puc fer backup de les dades?
Google Sheets té historial de versions automàtic. També pots exportar cada pestanya a CSV des del menú File > Download.

### Puc usar-lo en un centre amb múltiples aules ATECA?
Sí! Només has de configurar cada aula com un recurs diferent al panel d'administració.

## 📞 Contacte

Per qüestions específiques sobre aquest fork adaptat per a ATECA, obre un Issue al repositori.

Per qüestions sobre el projecte original, visita el [repositori de maestroseb](https://github.com/maestroseb/reservasrecursosysalas).

---

**Fet amb IA i ❤️ per a la comunitat educativa**
