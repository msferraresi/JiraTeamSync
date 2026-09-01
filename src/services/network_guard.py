import socket
from datetime import datetime
from src.utils.logger import logger


class NetworkGuard:
    @staticmethod
    def is_within_time_window(client) -> bool:
        now = datetime.now()

        # Validación de día hábil
        if (
            client.active_weekdays_only and now.weekday() >= 5
        ):  # 5 = Sábado, 6 = Domingo
            return False

        current_time = now.strftime("%H:%M")
        return client.time_window_start <= current_time <= client.time_window_end

    @staticmethod
    def is_vpn_reachable(client) -> bool:
        """Verifica conectividad contra el host de Jira o un endpoint interno de VPN."""
        host_to_check = client.vpn_check_host or client.domain
        try:
            # Intento de resolución DNS y conexión TCP corta en puerto HTTPS (443) o DNS (53)
            socket.setdefaulttimeout(3)
            socket.gethostbyname(host_to_check)
            s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
            s.connect((host_to_check, 443))
            s.close()
            return True
        except Exception:
            return False

    @classmethod
    def can_sync_client(cls, client) -> tuple[bool, str]:
        if not client.is_active:
            return False, "Cliente deshabilitado"
        if not cls.is_within_time_window(client):
            return (
                False,
                f"Fuera de franja horaria ({client.time_window_start} a {client.time_window_end})",
            )
        if not cls.is_vpn_reachable(client):
            return False, f"VPN desconectada o host inalcanzable ({client.domain})"
        return True, "OK"
