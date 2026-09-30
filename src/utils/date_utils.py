from datetime import timedelta
from typing import Any, Optional

from dateutil import parser


def parse_effective_issue_dates(
    fields: dict, role_map: dict
) -> tuple[Optional[object], Optional[object], dict[str, Any]]:
    """Resuelve las fechas efectivas de inicio y fin junto con los valores crudos para el cuerpo."""
    real_start = (
        fields.get(role_map.get("real_start")) if "real_start" in role_map else None
    )
    real_end = (
        fields.get(role_map.get("real_end"))
        if "real_end" in role_map
        else fields.get("resolutiondate")
    )
    exp_start = (
        fields.get(role_map.get("exp_start")) if "exp_start" in role_map else None
    )
    exp_end = (
        fields.get(role_map.get("exp_end"))
        if "exp_end" in role_map
        else fields.get("duedate")
    )

    effective_start = real_start or exp_start or real_end or exp_end
    effective_end = real_end or exp_end or real_start or exp_start

    raw_dates = {
        "real_start": real_start,
        "real_end": real_end,
        "exp_start": exp_start,
        "exp_end": exp_end,
    }

    if not effective_end:
        return None, None, raw_dates

    dt_start = parser.parse(effective_start).date()
    dt_end = parser.parse(effective_end).date()
    return dt_start, dt_end, raw_dates


def format_all_day_range(dt_start, dt_end) -> tuple[str, str]:
    """Genera las marcas de tiempo con final inclusivo para citas de día completo en Outlook."""
    dt_end_inclusive = dt_end + timedelta(days=1)
    return dt_start.strftime("%Y-%m-%d 00:00"), dt_end_inclusive.strftime(
        "%Y-%m-%d 00:00"
    )
