import re
from datetime import timedelta
from dateutil import parser
import win32com.client
from src.utils.logger import logger


class OutlookService:
    def __init__(self):
        self.outlook = win32com.client.Dispatch("Outlook.Application")
        self.namespace = self.outlook.GetNamespace("MAPI")
        self.calendar_folder = self.namespace.GetDefaultFolder(9)

    def get_existing_events_map(self):
        items = self.calendar_folder.Items
        items.IncludeRecurrences = False
        events_map = {}
        for item in items:
            try:
                subj = getattr(item, "Subject", "")
                # Busca [SPRINT-123] o [ABC-123] en cualquier parte del asunto
                matches = re.findall(r"\[([A-Z0-9_\-]+)\]", subj)
                for m in matches:
                    # Si coincide con SPRINT- o tiene formato de key (ej: CBS-5060)
                    if m.startswith("SPRINT-") or "-" in m:
                        events_map[m] = item
                        break
            except Exception:
                continue
        return events_map

    @staticmethod
    def resolve_category_for_item(item_type: str, status: str) -> str:
        status_lower = (status or "").lower().strip()

        if item_type == "sprint":
            if status_lower in ["closed", "cerrado"]:
                return "Categoría verde, Green Category"
            return "Categoría amarilla, Yellow Category"

        # Lógica para tickets / issues
        if status_lower in [
            "done",
            "closed",
            "finalizado",
            "resuelto",
            "cerrado",
            "completado",
        ]:
            return "Categoría verde, Green Category"

        if status_lower in ["in progress", "en curso", "en desarrollo", "in dev"]:
            return "Categoría púrpura, Purple Category"

        if status_lower in [
            "en espera",
            "bloqueado",
            "blocked",
            "waiting",
            "dev implementación pendiente",
            "pendiente",
        ]:
            return "Categoría naranja, Orange Category"

        # Color por defecto (backlog / to do)
        return "Categoría azul, Blue Category"

    def sync_board_issues(
        self,
        issues,
        sprints_data,
        field_mapping,
        board_config,
        domain,
        client_name: str,
    ):
        events_map = self.get_existing_events_map()

        role_map = {}
        for m in field_mapping:
            if m.get("role") and m["role"] != "none":
                role_map[m["role"]] = m["field_id"]

        created_count = 0
        updated_count = 0
        unchanged_count = 0

        # --- A. Sincronización Estricta de Sprints ---
        for sprint_entry in sprints_data:
            sprint = sprint_entry["meta"]
            sprint_issues = sprint_entry["issues"]
            s_id_num = sprint.get("id")
            s_name = sprint.get("name", "Sprint")
            s_id = f"SPRINT-{s_id_num}"
            s_state = sprint.get("state", "").lower()

            s_start = sprint.get("startDate")
            s_end = sprint.get("endDate") or sprint.get("completeDate")

            if not s_start or not s_end:
                continue

            dt_s_start = parser.parse(s_start).date()
            dt_s_end = parser.parse(s_end).date()
            dt_s_end_inclusive = dt_s_end + timedelta(days=1)

            my_sp_estimated = 0.0
            my_sp_completed = 0.0
            my_closed_issues = []
            my_pending_issues = []

            for issue in sprint_issues:
                fields = issue["fields"]
                status_obj = fields.get("status", {})
                status_cat = status_obj.get("statusCategory", {}).get("key", "")
                status_name = status_obj.get("name", "")
                summary = fields.get("summary", "")
                resolution_date_str = fields.get("resolutiondate")
                k = issue["key"]

                sp_num = self._extract_points(fields, role_map, field_mapping)

                is_done = status_cat == "done"

                if is_done:
                    if resolution_date_str:
                        dt_res = parser.parse(resolution_date_str).date()
                        if dt_s_start <= dt_res <= dt_s_end:
                            my_sp_estimated += sp_num
                            my_sp_completed += sp_num
                            my_closed_issues.append(
                                f"  ✅ [{k}] ({sp_num:g} SP) {summary}"
                            )
                        elif s_state == "closed" and dt_res <= dt_s_end:
                            pass
                    else:
                        if s_state == "closed":
                            my_sp_estimated += sp_num
                            my_sp_completed += sp_num
                            my_closed_issues.append(
                                f"  ✅ [{k}] ({sp_num:g} SP) {summary}"
                            )
                else:
                    my_sp_estimated += sp_num
                    my_pending_issues.append(
                        f"  ⏳ [{k}] ({sp_num:g} SP | {status_name}) {summary}"
                    )

            hours_per_sp = board_config.hours_per_sp or 4
            my_hours_completed = my_sp_completed * hours_per_sp
            my_hours_estimated = my_sp_estimated * hours_per_sp
            pct = (
                int((my_sp_completed / my_sp_estimated) * 100) if my_sp_estimated else 0
            )

            # 1. Determinar el estado para la etiqueta
            sprint_state_tag = (
                "CERRADO" if s_state in ["closed", "cerrado"] else "ACTIVO"
            )

            # 2. Asunto con tag explícito [ESTADO: ...]
            sprint_subject = f"[{client_name}] [{s_id}] [ESTADO: {sprint_state_tag}] {s_name} | Mis SP: {my_sp_completed:g}/{my_sp_estimated:g} ({pct}%) [{my_hours_completed:g}h/{my_hours_estimated:g}h]"

            sprint_body = [
                f"Cliente: {client_name}",
                f"Tablero: {board_config.board_name}",
                f"Sprint: {s_name}",
                f"Estado del Sprint: {sprint.get('state', '').upper()}",
                f"Fechas: {dt_s_start} al {dt_s_end}",
                f"Mis Story Points : {my_sp_completed:g} finalizados / {my_sp_estimated:g} estimados",
                f"Mis Horas Totales: {my_hours_completed:g}h finalizadas / {my_hours_estimated:g}h estimadas",
                "====================================",
                f"Mis Tickets Cerrados en este Sprint ({len(my_closed_issues)}):",
                (
                    "\n".join(my_closed_issues)
                    if my_closed_issues
                    else "  (Ninguno cerrado en este período)"
                ),
                "------------------------------------",
                f"Mis Tickets Pendientes / En Curso ({len(my_pending_issues)}):",
                (
                    "\n".join(my_pending_issues)
                    if my_pending_issues
                    else "  (Todos completados)"
                ),
            ]
            sprint_body_content = "\n".join(sprint_body)
            sprint_category = self.resolve_category_for_item("sprint", s_state)

            start_str = dt_s_start.strftime("%Y-%m-%d 00:00")
            end_str = dt_s_end_inclusive.strftime("%Y-%m-%d 00:00")

            if s_id in events_map:
                ev = events_map[s_id]
                if (
                    getattr(ev, "Subject", "") != sprint_subject
                    or getattr(ev, "Body", "").strip() != sprint_body_content.strip()
                    or str(getattr(ev, "Start", ""))[:10] != start_str[:10]
                    or str(getattr(ev, "End", ""))[:10] != end_str[:10]
                    or getattr(ev, "Categories", "") != sprint_category
                ):
                    ev.Subject = sprint_subject
                    ev.Start = start_str
                    ev.End = end_str
                    ev.AllDayEvent = True
                    ev.ReminderSet = False
                    ev.Body = sprint_body_content
                    ev.Categories = sprint_category
                    ev.Save()
                    updated_count += 1
                    logger.info(f"🔄 Sprint personal actualizado: {s_name}")
                else:
                    unchanged_count += 1
            else:
                new_ev = self.outlook.CreateItem(1)
                new_ev.Subject = sprint_subject
                new_ev.Start = start_str
                new_ev.End = end_str
                new_ev.AllDayEvent = True
                new_ev.ReminderSet = False
                new_ev.Body = sprint_body_content
                new_ev.Categories = sprint_category
                new_ev.Save()
                created_count += 1
                logger.info(f"✅ Sprint personal creado en calendario: {s_name}")

        # --- B. Sincronizar Tickets Individuales ---
        for issue in issues:
            k = issue["key"]
            fields = issue["fields"]
            summary = fields.get("summary", "")
            status_name = fields.get("status", {}).get("name", "Desconocido")

            real_start = (
                fields.get(role_map.get("real_start"))
                if "real_start" in role_map
                else None
            )
            real_end = (
                fields.get(role_map.get("real_end"))
                if "real_end" in role_map
                else fields.get("resolutiondate")
            )

            exp_start = (
                fields.get(role_map.get("exp_start"))
                if "exp_start" in role_map
                else None
            )
            exp_end = (
                fields.get(role_map.get("exp_end"))
                if "exp_end" in role_map
                else fields.get("duedate")
            )

            effective_start = real_start or exp_start or real_end or exp_end
            effective_end = real_end or exp_end or real_start or exp_start

            if not effective_end:
                continue

            dt_start = parser.parse(effective_start).date()
            dt_end = parser.parse(effective_end).date()
            dt_end_inclusive = dt_end + timedelta(days=1)

            sp_num = self._extract_points(fields, role_map, field_mapping)

            hours_per_sp = board_config.hours_per_sp or 4
            total_hours = sp_num * hours_per_sp

            # Prefijo con cliente y tag explícito [ESTADO: ...]
            subject = f"[{client_name}] [{k}] [ESTADO: {status_name.upper()}] ({sp_num:g} SP / {total_hours:g}h) {summary}"

            body_lines = [
                f"Cliente: {client_name}",
                f"Ticket: https://{domain}/browse/{k}",
                f"Resumen: {summary}",
                f"Estado: {status_name}",
                f"Story Points: {sp_num:g} ({total_hours:g} hs de esfuerzo)",
                "------------------------------------",
                f"Inicio Real     : {real_start or 'Pendiente'}",
                f"Fin Real        : {real_end or 'Pendiente'}",
                f"Inicio Estimado : {exp_start or 'N/A'}",
                f"Fin Estimado    : {exp_end or 'N/A'}",
            ]

            extra_lines = []
            for m in field_mapping:
                if m.get("include_in_body") and m.get("role") not in [
                    "real_start",
                    "real_end",
                    "exp_start",
                    "exp_end",
                    "story_points",
                ]:
                    val = fields.get(m["field_id"])
                    if val:
                        extra_lines.append(f"{m['field_name']}: {val}")

            if extra_lines:
                body_lines.append("------------------------------------")
                body_lines.extend(extra_lines)

            body_content = "\n".join(body_lines)
            issue_category = self.resolve_category_for_item("ticket", status_name)
            start_str = dt_start.strftime("%Y-%m-%d 00:00")
            end_str = dt_end_inclusive.strftime("%Y-%m-%d 00:00")

            if k in events_map:
                ev = events_map[k]
                if (
                    getattr(ev, "Subject", "") != subject
                    or getattr(ev, "Body", "").strip() != body_content.strip()
                    or str(getattr(ev, "Start", ""))[:10] != start_str[:10]
                    or str(getattr(ev, "End", ""))[:10] != end_str[:10]
                    or getattr(ev, "Categories", "") != issue_category
                ):
                    ev.Subject = subject
                    ev.Start = start_str
                    ev.End = end_str
                    ev.AllDayEvent = True
                    ev.ReminderSet = False
                    ev.Body = body_content
                    ev.Categories = issue_category
                    ev.Save()
                    updated_count += 1
                    logger.info(
                        f"🔄 Ticket actualizado: {k} -> {dt_start} (Estado: {status_name})"
                    )
                else:
                    unchanged_count += 1
            else:
                new_event = self.outlook.CreateItem(1)
                new_event.Subject = subject
                new_event.Start = start_str
                new_event.End = end_str
                new_event.AllDayEvent = True
                new_event.ReminderSet = False
                new_event.Body = body_content
                new_event.Categories = issue_category
                new_event.Save()
                created_count += 1
                logger.info(f"✅ Ticket creado en calendario: {k} en {dt_start}")

        return created_count, updated_count, unchanged_count

    def _extract_points(
        self, fields: dict, role_map: dict, field_mapping: list
    ) -> float:
        """Obtiene Story Points o hace fallback a Bugpoints si el ticket es un bug."""
        # 1. Intentar por el rol estándar mapeado
        sp_fid = role_map.get("story_points")
        sp_val = fields.get(sp_fid) if sp_fid else None

        # 2. Si viene None o 0, buscar el campo Bugpoints
        if sp_val is None:
            # Buscar por nombre en el mapeo de campos guardado
            bp_fid = next(
                (
                    m["field_id"]
                    for m in field_mapping
                    if "bugpoint" in m.get("field_name", "").lower()
                ),
                None,
            )
            if bp_fid:
                sp_val = fields.get(bp_fid)

        # 3. Conversión segura a float
        try:
            return float(sp_val) if sp_val is not None else 0.0
        except (ValueError, TypeError):
            return 0.0
