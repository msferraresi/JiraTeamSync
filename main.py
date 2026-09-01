import os
import sys
import time
import threading
import webbrowser
import pystray
from PIL import Image, ImageDraw
import pythoncom

from src.db.database import SessionLocal, init_db
from src.db.models import Client, BoardConfig, AppSettings
from src.server.web_server import run_server
from src.services.jira_service import JiraService
from src.services.outlook_service import OutlookService
from src.services.network_guard import NetworkGuard
from src.utils.logger import logger, LOG_FILE

is_syncing = False
is_paused = False


def get_app_settings():
    session = SessionLocal()
    try:
        settings = session.query(AppSettings).first()
        if not settings:
            settings = AppSettings(server_port=8765, interval_minutes=15)
            session.add(settings)
            session.commit()
            session.refresh(settings)
        return settings.server_port, settings.interval_minutes
    finally:
        session.close()


def execute_sync():
    global is_syncing
    if is_syncing:
        return
    is_syncing = True

    session = SessionLocal()
    pythoncom.CoInitialize()

    try:
        clients = session.query(Client).filter(Client.is_active == True).all()
        if not clients:
            logger.info("No hay clientes configurados o activos.")
            return

        outlook_svc = OutlookService()

        for client in clients:
            can_sync, reason = NetworkGuard.can_sync_client(client)
            if not can_sync:
                logger.info(f"Omitiendo cliente [{client.name}]: {reason}")
                continue

            logger.info(f"Procesando cliente: {client.name} ({client.domain})...")
            jira_svc = JiraService(client)

            try:
                all_fields = jira_svc.get_fields()
            except Exception as e:
                logger.error(f"Falla al conectar con Jira de {client.name}: {e}")
                continue

            for board in client.boards:
                if not board.is_enabled:
                    continue

                logger.info(f"Sincronizando tablero: {board.board_name}")
                issues, sprints_data, field_mapping = jira_svc.fetch_board_issues(board)

                created, updated, unchanged = outlook_svc.sync_board_issues(
                    issues=issues,
                    sprints_data=sprints_data,
                    field_mapping=field_mapping,
                    board_config=board,
                    domain=client.domain,
                )
                logger.info(
                    f"[{board.board_name}] Creados: {created} | Actualizados: {updated} | Sin cambios: {unchanged}"
                )
    except Exception as e:
        logger.error(
            f"Error general durante el ciclo de sincronización: {e}", exc_info=True
        )
    finally:
        session.close()
        pythoncom.CoUninitialize()
        is_syncing = False


def background_loop():
    while True:
        _, interval = get_app_settings()
        if not is_paused:
            execute_sync()
        for _ in range(interval * 60):
            time.sleep(1)


def open_settings(icon, item):
    port, _ = get_app_settings()
    webbrowser.open(f"http://127.0.0.1:{port}")


def open_logs(icon, item):
    if os.path.exists(LOG_FILE):
        os.startfile(LOG_FILE)
    else:
        logger.info("El archivo de log aún no se ha generado.")


def on_sync_now(icon, item):
    threading.Thread(target=execute_sync, daemon=True).start()


def toggle_pause(icon, item):
    global is_paused
    is_paused = not is_paused


def on_exit(icon, item):
    icon.stop()
    os._exit(0)


def create_tray_icon():
    if getattr(sys, "frozen", False):
        base_dir = getattr(sys, "_MEIPASS", os.path.dirname(sys.executable))
    else:
        base_dir = os.path.dirname(os.path.abspath(__file__))

    icon_path = os.path.join(base_dir, "resource", "images.png")
    if not os.path.exists(icon_path):
        icon_path = os.path.join(base_dir, "resource", "diagrama.gif")

    if os.path.exists(icon_path):
        try:
            return Image.open(icon_path)
        except Exception:
            pass

    # Fallback si no encuentra el archivo gráfico
    img = Image.new("RGB", (64, 64), color=(2, 132, 199))
    d = ImageDraw.Draw(img)
    d.rectangle([16, 16, 48, 48], fill=(255, 255, 255))
    return img


if __name__ == "__main__":
    # 1. Crear tablas en SQLite
    init_db()

    server_port, _ = get_app_settings()

    # 2. Iniciar servidor FastAPI en segundo plano
    server_thread = threading.Thread(
        target=run_server, args=(server_port,), daemon=True
    )
    server_thread.start()

    # 3. Iniciar ciclo de sincronización periódica
    sync_thread = threading.Thread(target=background_loop, daemon=True)
    sync_thread.start()

    # 4. Construir menú de la bandeja del sistema
    menu = pystray.Menu(
        pystray.MenuItem("Configuración ⚙️", open_settings, default=True),
        pystray.MenuItem("Sincronizar ahora 🔄", on_sync_now),
        pystray.MenuItem("Ver Logs 📄", open_logs),
        pystray.MenuItem(
            lambda text: "Reanudar ▶️" if is_paused else "Pausar ⏸️", toggle_pause
        ),
        pystray.Menu.SEPARATOR,
        pystray.MenuItem("Salir ❌", on_exit),
    )

    icon = pystray.Icon("JiraTeamsSync", create_tray_icon(), "Jira -> Teams Sync", menu)
    icon.run()
