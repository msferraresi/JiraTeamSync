import re
from typing import List, Dict, Any, Optional


def sanitize_and_prepare_jql(raw_jql: str, force: bool = False) -> str:
    """Separa ORDER BY y añade la ventana temporal para sincronizaciones periódicas."""
    raw = (raw_jql or "").strip()
    order_by_clause = ""
    if "order by" in raw.lower():
        parts = re.split(r"(?i)\s+order\s+by\s+", raw, maxsplit=1)
        raw = parts[0].strip()
        order_by_clause = f" ORDER BY {parts[1].strip()}"

    if not force:
        return f"({raw}) AND (statusCategory != Done OR resolutiondate >= -30d){order_by_clause}"
    return f"{raw}{order_by_clause}"


def extract_story_points(fields: dict, role_map: dict, field_mapping: list) -> float:
    """Extrae puntos de historia buscando por rol asignado o heurística de nombres de campo."""
    point_field_ids = []
    primary_sp_fid = role_map.get("story_points")
    if primary_sp_fid:
        point_field_ids.append(primary_sp_fid)

    secondary_fids = []
    for m in field_mapping:
        fid = m.get("field_id")
        fname = m.get("field_name", "").lower()

        if not fid or fid in point_field_ids:
            continue

        if any(
            term in fname
            for term in ["point", "puntos", "story point", "bugpoint", "taskpoint"]
        ):
            if "final" in fname or "real" in fname:
                secondary_fids.insert(0, fid)
            else:
                secondary_fids.append(fid)

    point_field_ids.extend(secondary_fids)

    for fid in point_field_ids:
        val = fields.get(fid)
        if val is not None:
            try:
                num = float(val)
                if num > 0:
                    return num
            except (ValueError, TypeError):
                continue
    return 0.0


def extract_user_logged_hours(fields: dict, current_user_email: str = "") -> float:
    """Calcula las horas cargadas por el usuario actual sumando sus worklogs específicos."""
    worklog_data = fields.get("worklog") or {}
    worklogs = worklog_data.get("worklogs", [])

    my_seconds = 0
    if current_user_email and worklogs:
        target_email = current_user_email.lower().strip()
        target_user = target_email.split("@")[0]

        for entry in worklogs:
            author = entry.get("author", {})
            author_email = (author.get("emailAddress") or "").lower().strip()
            author_name = (author.get("displayName") or "").lower().strip()

            if (
                (target_email and target_email in author_email)
                or (target_user and target_user in author_email)
                or (target_user and target_user in author_name)
            ):
                my_seconds += entry.get("timeSpentSeconds", 0)

    if my_seconds == 0:
        my_seconds = fields.get("timespent") or 0

    return round(my_seconds / 3600.0, 2)
