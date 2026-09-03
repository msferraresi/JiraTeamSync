import json
import requests
from src.utils.logger import logger


class JiraService:
    def __init__(self, client):
        self.client = client
        self.auth = (client.email, client.api_token)
        self.base_url = f"https://{client.domain}"

    def get_available_boards(self):
        """Obtiene la lista de tableros (Scrum/Kanban) a los que tiene acceso la cuenta."""
        url = f"{self.base_url}/rest/agile/1.0/board"
        try:
            res = requests.get(url, auth=self.auth, timeout=10)
            res.raise_for_status()
            boards = res.json().get("values", [])
            return [
                {"id": b["id"], "name": b["name"], "type": b.get("type")}
                for b in boards
            ]
        except Exception as e:
            logger.error(f"Error obteniendo tableros de {self.client.name}: {e}")
            return []

    def get_fields(self):
        """Reflection de campos disponibles clasificados por tipo."""
        url = f"{self.base_url}/rest/api/3/field"
        res = requests.get(url, auth=self.auth, timeout=10)
        res.raise_for_status()
        return res.json()

    def fetch_board_issues(self, board_config, force: bool = False):
        mapping = json.loads(board_config.fields_mapping_json or "[]")

        query_fields = [
            "summary",
            "key",
            "status",
            "duedate",
            "resolutiondate",
            "created",
            "customfield_10020",
        ]
        for m in mapping:
            fid = m["field_id"]
            if fid not in query_fields:
                query_fields.append(fid)

        # 1. Tickets: si es automático, solo no terminados o terminados hace <= 30 días
        effective_jql = board_config.custom_jql
        if not force:
            effective_jql = f"({effective_jql}) AND (statusCategory != Done OR resolutiondate >= -30d)"

        url = f"{self.base_url}/rest/api/3/search/jql"
        params = {
            "jql": effective_jql,
            "maxResults": 100,
            "fields": query_fields,
        }
        res = requests.get(url, auth=self.auth, params=params, timeout=15)
        res.raise_for_status()
        board_issues = res.json().get("issues", [])

        # 2. Sprints: si es automático solo trae el activo; si es force trae los últimos cerrados
        sprints_data = []
        raw_sprints = self.fetch_board_sprints(board_config.board_id, force=force)
        for s in raw_sprints:
            s_id = s.get("id")
            s_issues = self.fetch_sprint_my_issues(s_id, query_fields)
            sprints_data.append({"meta": s, "issues": s_issues})

        return board_issues, sprints_data, mapping

    def fetch_board_sprints(self, board_id, force: bool = False):
        if not board_id:
            return []

        # En corrida normal solo consulta el activo; si se fuerza trae cerrados recientes
        state_filter = "active,closed" if force else "active"
        url = f"{self.base_url}/rest/agile/1.0/board/{board_id}/sprint?state={state_filter}"
        try:
            res = requests.get(url, auth=self.auth, timeout=10)
            if res.status_code != 200:
                return []
            sprints = res.json().get("values", [])
            return sprints[-3:] if (force and len(sprints) > 3) else sprints
        except Exception as e:
            logger.warning(
                f"No se pudieron obtener sprints del tablero {board_id}: {e}"
            )
            return []

    def fetch_sprint_my_issues(self, sprint_id, query_fields):
        """Consulta tickets asignados al usuario filtrando por el ID específico de Sprint en JQL."""
        url = f"{self.base_url}/rest/api/3/search/jql"
        params = {
            "jql": f"sprint = {sprint_id} AND assignee = currentUser()",
            "maxResults": 100,
            "fields": query_fields,
        }
        try:
            res = requests.get(url, auth=self.auth, params=params, timeout=15)
            res.raise_for_status()
            return res.json().get("issues", [])
        except Exception as e:
            logger.error(f"Error consultando tickets del Sprint {sprint_id}: {e}")
            return []
