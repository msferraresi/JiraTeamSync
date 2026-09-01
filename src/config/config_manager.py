import os
import sys
import yaml


def get_base_dir():
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    # 2 niveles hacia arriba desde src/config/
    return os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))


CONFIG_FILE = os.path.join(get_base_dir(), "settings.yml")

DEFAULT_CONFIG = {
    "jira": {
        "domain": "",
        "email": "",
        "api_token": "",
        "jql": 'assignee = currentUser() AND status != "Cancelado" ORDER BY created ASC',
    },
    "fields": {
        "start_field": "Fecha de Inicio Esperada",
        "end_field": "Fecha de Finalizacion Esperada",
    },
    "app": {"interval_minutes": 15, "server_port": 8765},
}


def load_settings():
    if not os.path.exists(CONFIG_FILE):
        save_settings(DEFAULT_CONFIG)
        return DEFAULT_CONFIG
    try:
        with open(CONFIG_FILE, "r", encoding="utf-8") as f:
            cfg = yaml.safe_load(f) or DEFAULT_CONFIG
            # Asegurar fallback si falta la clave en settings existentes
            if "server_port" not in cfg.get("app", {}):
                cfg.setdefault("app", {})["server_port"] = 8765
            return cfg
    except Exception:
        return DEFAULT_CONFIG


def save_settings(data):
    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        yaml.dump(data, f, default_flow_style=False, allow_unicode=True)
