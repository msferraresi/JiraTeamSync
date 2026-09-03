from datetime import datetime, timedelta
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


def should_sync_client_now(client, force: bool = False) -> bool:
    if not client.is_active:
        return False
    if force or not client.last_synced_at:
        return True

    last_synced = client.last_synced_at
    if isinstance(last_synced, str):
        try:
            last_synced = datetime.fromisoformat(last_synced)
        except Exception:
            return True

    elapsed = datetime.now() - last_synced
    interval_delta = timedelta(minutes=client.sync_interval_minutes or 60)
    return elapsed >= interval_delta


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


def execute_sync(force: bool = False, target_client_id: int = None):
    global is_syncing
    if is_syncing:
        logger.warning(
            "⏳ Ya hay un ciclo de sincronización en curso. Omitiendo solicitud."
        )
        return
    is_syncing = True

    session = SessionLocal()
    pythoncom.CoInitialize()

    try:
        logger.info(
            f"🚀 Iniciando ciclo de sincronización (force={force}, target_client_id={target_client_id})..."
        )
        query = session.query(Client).filter(Client.is_active == True)
        if target_client_id:
            query = query.filter(Client.id == target_client_id)

        clients = query.all()
        if not clients:
            logger.info("ℹ️ No hay clientes activos configurados para sincronizar.")
            return

        outlook_svc = OutlookService()

        for client in clients:
            if not should_sync_client_now(client, force=force):
                logger.info(
                    f"⏳ Cliente [{client.name}]: aún no transcurrió su intervalo de {client.sync_interval_minutes} min."
                )
                continue

            can_sync, reason = NetworkGuard.can_sync_client(client)
            if not can_sync:
                logger.info(f"⛔ Omitiendo cliente [{client.name}]: {reason}")
                continue

            logger.info(f"🔄 Procesando cliente: {client.name} ({client.domain})...")
            jira_svc = JiraService(client)

            try:
                all_fields = jira_svc.get_fields()
            except Exception as e:
                logger.error(f"❌ Falla al conectar con Jira de {client.name}: {e}")
                continue

            if not client.boards:
                logger.info(
                    f"⚠️ Cliente [{client.name}] no tiene tableros configurados."
                )
                continue

            for board in client.boards:
                if not board.is_enabled:
                    continue

                logger.info(f"📋 Sincronizando tablero: {board.board_name}")
                try:
                    issues, sprints_data, field_mapping = jira_svc.fetch_board_issues(
                        board, force=force
                    )

                    created, updated, unchanged = outlook_svc.sync_board_issues(
                        issues=issues,
                        sprints_data=sprints_data,
                        field_mapping=field_mapping,
                        board_config=board,
                        domain=client.domain,
                        client_name=client.name,
                    )
                    logger.info(
                        f"📊 [{board.board_name}] Creados: {created} | Actualizados: {updated} | Sin cambios: {unchanged}"
                    )
                except Exception as b_err:
                    logger.error(
                        f"❌ Error al procesar tablero {board.board_name}: {b_err}",
                        exc_info=True,
                    )

            client.last_synced_at = datetime.now()
            session.commit()

        logger.info("🏁 Ciclo de sincronización completado.")

    except Exception as e:
        logger.error(
            f"❌ Error general durante el ciclo de sincronización: {e}", exc_info=True
        )
    finally:
        session.close()
        pythoncom.CoUninitialize()
        is_syncing = False


def background_loop():
    while True:
        if not is_paused:
            execute_sync(force=False)
        time.sleep(60)


def open_settings(icon=None, item=None):
    port, _ = get_app_settings()
    webbrowser.open(f"http://127.0.0.1:{port}")


def open_logs(icon=None, item=None):
    if os.path.exists(LOG_FILE):
        os.startfile(LOG_FILE)
    else:
        logger.info("El archivo de log aún no se ha generado.")


def toggle_pause(icon=None, item=None):
    global is_paused
    is_paused = not is_paused


def on_exit(icon=None, item=None):
    if icon:
        icon.stop()
    os._exit(0)


def sync_specific_client(client_id: int):
    threading.Thread(
        target=execute_sync,
        kwargs={"force": True, "target_client_id": client_id},
        daemon=True,
    ).start()


def sync_all_clients(icon=None, item=None):
    logger.info("🖱️ Sincronización manual: todos los clientes")
    threading.Thread(
        target=execute_sync,
        kwargs={"force": True, "target_client_id": None},
        daemon=True,
    ).start()


def make_sync_handler(client_id: int):
    def _handler(icon=None, item=None):
        logger.info(f"🖱️ Sincronización manual: cliente ID {client_id}")
        sync_specific_client(client_id)

    return _handler


def get_sync_menu_items():
    yield pystray.MenuItem("Todos los clientes 🌐", sync_all_clients)
    yield pystray.Menu.SEPARATOR

    session = SessionLocal()
    try:
        clients = session.query(Client).filter(Client.is_active == True).all()
        for c in clients:
            yield pystray.MenuItem(f"{c.name} ({c.domain})", make_sync_handler(c.id))
    except Exception as e:
        logger.error(f"Error cargando clientes para el menú de bandeja: {e}")
    finally:
        session.close()


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

    img = Image.new("RGB", (64, 64), color=(2, 132, 199))
    d = ImageDraw.Draw(img)
    d.rectangle([16, 16, 48, 48], fill=(255, 255, 255))
    return img


if __name__ == "__main__":
    init_db()

    server_port, _ = get_app_settings()

    server_thread = threading.Thread(
        target=run_server, args=(server_port,), daemon=True
    )
    server_thread.start()

    sync_thread = threading.Thread(target=background_loop, daemon=True)
    sync_thread.start()

    menu = pystray.Menu(
        pystray.MenuItem("Configuración ⚙️", open_settings, default=True),
        pystray.MenuItem("Sincronizar 🔄", pystray.Menu(get_sync_menu_items)),
        pystray.MenuItem("Ver Logs 📄", open_logs),
        pystray.MenuItem(
            lambda text: "Reanudar ▶️" if is_paused else "Pausar ⏸️", toggle_pause
        ),
        pystray.Menu.SEPARATOR,
        pystray.MenuItem("Salir ❌", on_exit),
    )

    icon = pystray.Icon("JiraTeamsSync", create_tray_icon(), "Jira -> Teams Sync", menu)
    icon.run()
