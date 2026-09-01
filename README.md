# Jira to Teams/Outlook Sync (JiraTeamsSync)

Herramienta de sincronización en segundo plano para Windows que integra tickets asignados y sprints de Jira Software con el calendario de Outlook/Teams mediante la API local MAPI COM.

Incluye interfaz web local construida sobre FastAPI y SQLite para administrar múltiples clientes, mapear campos dinámicos mediante introspección de Jira, definir ventanas horarias de red/VPN y controlar la ejecución desde la bandeja del sistema (System Tray).

---

## 🚀 Características Principales

- **Arquitectura Multicliente (Multi-Tenant):** Configuración independiente para múltiples entornos Jira (ej. FedPat, Supervielle, Citi).
- **Control de Conectividad y VPN (`NetworkGuard`):** Valida ventanas horarias operativas (ej. 07:00 a 20:00) y verifica conectividad/VPN antes de disparar consultas para evitar bloqueos y logs de error.
- **Introspección y Mapeo Dinámico de Campos:** Detección automática del catálogo de campos de Jira con asignación interactiva de roles (`Fecha Inicio Real`, `Fecha Fin Real`, `Fecha Inicio Esperada`, `Fecha Fin Esperada`, `Story Points`).
- **Cálculo de Esfuerzo en Horas:** Conversión automática de Story Points a horas de trabajo reales/estimadas (ej. escala Fibonacci 1 SP = 4 hs) proyectadas en el asunto y cuerpo del evento.
- **Visualización y Métricas de Sprints:** Generación de eventos que abarcan la duración completa del Sprint con métricas personales filtradas (SP estimados vs. SP completados, horas consumidas y desglose de tickets resueltos vs. pendientes).
- **Persistencia en SQLite:** Almacenamiento local mediante SQLAlchemy sin necesidad de archivos de texto o YAML externos.
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
│   │   └── web_server.py       # FastAPI UI y endpoints REST
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

2. Se levanta el servidor web local en http://127.0.0.1:8765.

3. Se inicia el bucle de sincronización periódica en segundo plano.

4. Se instancia el icono de control en la bandeja del sistema (al lado del reloj de Windows).

---

## 📌 Menú de la Bandeja del Sistema (System Tray)
Al hacer clic derecho sobre el icono en la barra de tareas de Windows:

- Abrir Configuración Web: Abre el navegador en http://127.0.0.1:8765 para administrar clientes, tableros y mapeos.

- Sincronizar ahora 🔄: Dispara una comprobación manual inmediata contra los tableros configurados.

- Pausar / Reanudar ⏸️: Detiene o reanuda las consultas automáticas periódicas.

- Intervalo ⏱️: Permite ajustar el tiempo entre revisiones automáticas (5m, 15m, 30m, 60m).

- Salir ❌: Detiene el servidor y finaliza la aplicación de forma limpia.

---

## 🧪 Ejecución de Pruebas Unitarias
PowerShell
```text
# Correr toda la suite de pruebas
pytest tests/ -v
```

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
El ejecutable quedará disponible en dist/JiraTeamsSync.exe.