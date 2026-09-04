# Jira to Teams/Outlook Sync (JiraTeamsSync)

Herramienta de sincronización en segundo plano para Windows que integra tickets asignados y sprints de Jira Software con el calendario de Outlook/Teams mediante la API local MAPI COM.

Incluye interfaz web local construida sobre FastAPI y SQLite para administrar múltiples clientes, mapear campos dinámicos mediante introspección de Jira, definir ventanas horarias de red/VPN y controlar la ejecución desde la bandeja del sistema (System Tray).

---

## 🚀 Características Principales

- **Arquitectura Multicliente (Multi-Tenant):** Configuración independiente de entornos Jira con intervalos de sincronización propios (`sync_interval_minutes`) y ventanas horarias específicas por cliente.
- **Control de Conectividad y VPN (`NetworkGuard`):** Valida ventanas horarias operativas (ej. 07:00 a 20:00) y verifica conectividad/VPN antes de disparar consultas para evitar bloqueos y logs de error.
- **Identificación Unívoca en Calendario:** Eventos prefijados con la conexión del cliente y etiquetas normalizadas de estado (ej. `[FedPat] [CBS-5060] [ESTADO: EN CURSO]`) para facilitar filtros y coloreado automático.
- **Introspección y Mapeo Dinámico de Campos:** Detección automática del catálogo de campos de Jira con asignación interactiva de roles (`Fecha Inicio Real`, `Fecha Fin Real`, `Fecha Inicio Esperada`, `Fecha Fin Esperada`, `Story Points`).
- **Cálculo de Esfuerzo en Horas:** Conversión automática de Story Points a horas de trabajo reales/estimadas proyectadas en el asunto y cuerpo del evento.
- **Métricas de Sprints:** Generación de eventos de Sprint con estado explícito (`[ESTADO: ACTIVO]` o `[ESTADO: CERRADO]`), progreso de Story Points (estimados vs. completados) y desglose de tickets resueltos vs. pendientes.
- **Consola de Logs Integrada:** Pestaña web con visor en vivo y buscador/filtro en tiempo real por cliente o ticket.
- **Sin permisos de Azure AD:** Conexión nativa directa vía MAPI COM con el cliente Outlook instalado en la máquina.

---

## 📁 Estructura del Proyecto

```text
JiraTeamSync/
├── resource/
│   └── images.png              # Icono para la bandeja del sistema
├── src/
│   ├── db/
│   │   ├── database.py         # Motor y sesión de SQLAlchemy
│   │   └── models.py           # Modelos: Client, BoardConfig, AppSettings
│   ├── server/
│   │   └── web_server.py       # FastAPI UI multisolapa, endpoints REST y visor de logs
│   ├── services/
│   │   ├── jira_service.py     # Integración REST con Jira API / Agile API
│   │   ├── outlook_service.py  # Manipulación del calendario vía MAPI COM
│   │   └── network_guard.py    # Validación de ventanas horarias y host de VPN
│   └── utils/
│       └── logger.py           # Logger rotativo con persistencia en app.log
├── tests/
│   ├── conftest.py             # Fixtures para pruebas aisladas
│   ├── test_network_guard.py   # Pruebas de validación de red y horarios
│   ├── test_jira_service.py    # Pruebas de filtrado y parsing de Jira
│   └── test_web_server.py      # Pruebas de endpoints FastAPI y SQLite
├── app.db                      # Base de datos SQLite (generada en runtime)
├── app.log                     # Archivo de registro local
├── main.py                     # Entry point y orquestador del System Tray
├── pytest.ini                  # Configuración y filtros de la suite de pruebas
├── README.md                   # Documentación técnica
├── requirements.txt            # Dependencias base de producción / ejecución
└── requirements-dev.txt        # Dependencias de desarrollo, testing y compilación
```

---

## ⚙️ Requisitos Previos

- Windows 10/11 x64.
- Python 3.10+.
- Microsoft Outlook instalado y configurado con la cuenta corporativa.
- Token de API de Atlassian generado desde id.atlassian.com.

---

## 🛠️ Instalación y Entorno

PowerShell

```text
# 1. Crear y activar entorno virtual
python -m venv venv
.\venv\Scripts\activate

# 2. Instalar dependencias completas de desarrollo y testing
pip install -r requirements-dev.txt
```

---

## 💻 Ejecución en Desarrollo

PowerShell

```text
python .\main.py
```

---

Al iniciar:

1. Se crean automáticamente las tablas en app.db si no existen.

2. Se levanta el servidor web local en <http://127.0.0.1:8765>

3. Se inicia el bucle de sincronización periódica en segundo plano.

4. Se instancia el icono interactivo en la bandeja del sistema (System Tray).

---

## 📌 Menú de la Bandeja del Sistema (System Tray)

Al hacer clic derecho sobre el icono en la barra de tareas de Windows:

- Configuración ⚙️: Abre el navegador en la interfaz web de gestión (solapas General, Conexiones y Logs).

- Sincronizar 🔄: Submenú dinámico con opciones para:

  - Todos los clientes 🌐: Fuerza la sincronización completa de todos los clientes activos.

  - [Cliente individual]: Dispara la sincronización inmediata del cliente seleccionado ignorando su intervalo.

- Ver Logs 📄: Abre el archivo de registro app.log local en el editor por defecto.

- Pausar / Reanudar ⏸️: Detiene o reanuda las consultas automáticas de segundo plano.

- Salir ❌: Cierra el servidor y finaliza la aplicación de forma limpia.

---

## 🧪 Ejecución de Pruebas Unitarias

PowerShell

```text
# Correr toda la suite de pruebas
pytest tests/ -v
```

---

## 🎨 Configuración de Colores en Microsoft Teams (Formato Condicional)

Para que Teams resalte los eventos con la paleta de colores por estado, ve a **Teams > Calendario > Configuración (⚙️) > Conditional formatting (Formato condicional)** y da de alta las siguientes reglas mediante la opción **Subject includes**:

| Regla | Condición (*Subject includes*) | Color Sugerido |
| :--- | :--- | :--- |
| **Sprint Cerrado** | `[ESTADO: CERRADO]` | Verde oscuro (*Dark green*) |
| **Sprint Activo** | `[ESTADO: ACTIVO]` | Amarillo (*Yellow*) |
| **Ticket Cerrado** | `[ESTADO: FINALIZADO]`, `[ESTADO: CERRADO]`, `[ESTADO: RESUELTO]`, `[ESTADO: DONE]` | Verde (*Dark green* / *Green*) |
| **Ticket Bloqueado / Espera** | `[ESTADO: EN ESPERA]`, `[ESTADO: BLOQUEADO]` | Naranja (*Peach*) |
| **Ticket En Curso** | `[ESTADO: EN CURSO]`, `[ESTADO: EN DESARROLLO]`, `[ESTADO: IN PROGRESS]` | Púrpura (*Plum*) |

> **Nota de prioridad:** En el listado de reglas, sitúa **Sprint Cerrado** por encima de **Sprint Activo**. Teams evaluará las condiciones en cascada y aplicará el color exacto sobre el calendario.

---

## 📦 Compilación a Ejecutable (.exe)

Para compilar el proyecto en un binario portable e independiente:

PowerShell

- Ejecuta el siguiente comando para borrar las compilaciones previas.

```text
Remove-Item -Recurse -Force build, dist -ErrorAction SilentlyContinue
```

- Ejecuta este comando para compilar el proyecto en un archivo exe,

```text
pyinstaller --noconsole --onefile `
  --collect-all uvicorn `
  --collect-all sqlalchemy `
  --collect-all fastapi `
  --collect-all pydantic `
  --add-data "resource;resource" `
  --name JiraTeamsSync main.py
```

El ejecutable quedará disponible en /dist/JiraTeamsSync.exe
