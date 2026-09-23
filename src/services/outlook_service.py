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
                matches = re.findall(r"\[([A-Z0-9_\-]+)\]", subj)
                for m in matches:
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
                return "Green category"
            return "Yellow category"

        if status_lower in [
            "done",
            "closed",
            "finalizado",
            "resuelto",
            "cerrado",
            "completado",
        ]:
            return "Green category"

        if status_lower in ["in progress", "en curso", "en desarrollo", "in dev"]:
            return "Purple category"

        if status_lower in [
            "en espera",
            "bloqueado",
            "blocked",
            "waiting",
            "dev implementación pendiente",
            "pendiente",
        ]:
            return "Orange category"

        return "Blue category"

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
            sprint_issues = sprint_entry.get("issues", [])
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

            # 1. Unificar todos los candidatos (de la query de sprint y del tablero general)
            all_candidate_issues = {}
            for it in sprint_issues:
                all_candidate_issues[it["key"]] = it
            for it in issues:
                all_candidate_issues[it["key"]] = it

            # Prefijos de proyecto válidos para este sprint (ej: CBS)
            sprint_project_prefixes = {
                it["key"].split("-")[0]
                for it in sprint_issues
                if "-" in it.get("key", "")
            }

            # 2. Evaluar cada ticket para ver si realmente pertenece a este Sprint
            for k, issue in all_candidate_issues.items():
                fields = issue["fields"]
                status_obj = fields.get("status", {})
                status_cat = status_obj.get("statusCategory", {}).get("key", "")
                status_name = status_obj.get("name", "").strip()
                status_lower = status_name.lower()

                # Ignorar tickets cancelados/descartados
                if any(
                    term in status_lower
                    for term in ["cancelado", "cancelled", "descartado", "rechazado"]
                ):
                    continue

                # AISLAMIENTO DE PROYECTO (Detección dinámica de sprints sin IDs fijos):
                sprint_raw_val = fields.get("customfield_10020")
                if not sprint_raw_val:
                    for f_key, f_val in fields.items():
                        if "sprint" in f_key.lower() and f_val:
                            sprint_raw_val = f_val
                            break

                sprint_field_raw = str(sprint_raw_val or "")
                explicitly_in_sprint = (str(s_id_num) in sprint_field_raw) or (
                    s_name.lower() in sprint_field_raw.lower()
                )
                officially_in_sprint = k in {it["key"] for it in sprint_issues}

                summary = fields.get("summary", "")
                res_str = fields.get("resolutiondate")

                real_end = (
                    fields.get(role_map.get("real_end"))
                    if "real_end" in role_map
                    else None
                )
                real_start = (
                    fields.get(role_map.get("real_start"))
                    if "real_start" in role_map
                    else None
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

                effective_date_str = (
                    res_str or real_end or exp_end or real_start or exp_start
                )
                dt_ticket = (
                    parser.parse(effective_date_str).date()
                    if effective_date_str
                    else None
                )
                in_date_window = bool(
                    dt_ticket and (dt_s_start <= dt_ticket <= dt_s_end)
                )

                ticket_prefix = k.split("-")[0] if "-" in k else ""
                same_project_family = (
                    ticket_prefix in sprint_project_prefixes
                    if sprint_project_prefixes
                    else True
                )

                # Si no está oficialmente asignado al sprint, sólo entra si es del mismo proyecto Y cae en fecha
                if not explicitly_in_sprint and not officially_in_sprint:
                    if not (same_project_family and in_date_window):
                        continue

                # FILTRO DE EXCLUSIÓN PARA TICKETS VIEJOS:
                close_date_str = res_str or real_end
                if close_date_str:
                    dt_closed = parser.parse(close_date_str).date()
                    if dt_closed < dt_s_start:
                        continue

                # Extracción de puntos y tipo
                sp_num = self._extract_points(fields, role_map, field_mapping)
                issue_type_obj = fields.get("issuetype") or {}
                type_name = (
                    issue_type_obj.get("name", "").upper()
                    if isinstance(issue_type_obj, dict)
                    else ""
                )
                type_badge = f"[{type_name}] " if type_name else ""

                is_done = (status_cat == "done") or (
                    status_name.lower()
                    in ["cerrado", "closed", "resuelto", "finalizado"]
                )

                if is_done:
                    my_sp_estimated += sp_num
                    my_sp_completed += sp_num
                    my_closed_issues.append(
                        f"  ✅ {type_badge}[{k}] ({sp_num:g} SP) {summary}"
                    )
                else:
                    my_sp_estimated += sp_num
                    my_pending_issues.append(
                        f"  ⏳ {type_badge}[{k}] ({sp_num:g} SP | {status_name}) {summary}"
                    )

            hours_per_sp = board_config.hours_per_sp or 4
            my_hours_completed = my_sp_completed * hours_per_sp
            my_hours_estimated = my_sp_estimated * hours_per_sp
            pct = (
                int((my_sp_completed / my_sp_estimated) * 100) if my_sp_estimated else 0
            )

            sprint_state_tag = (
                "CERRADO" if s_state in ["closed", "cerrado"] else "ACTIVO"
            )

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
            status_name = fields.get("status", {}).get("name", "Desconocido")
            status_lower = status_name.lower().strip()

            # Descartar tickets cancelados
            if any(
                term in status_lower
                for term in ["cancelado", "cancelled", "descartado", "rechazado"]
            ):
                if k in events_map:
                    try:
                        events_map[k].Delete()
                    except Exception:
                        pass
                continue

            # Extracción del tipo de ticket (Bug, Subtarea, Historia, etc.)
            issue_type_obj = fields.get("issuetype") or {}
            type_name = (
                issue_type_obj.get("name", "").upper()
                if isinstance(issue_type_obj, dict)
                else ""
            )
            type_tag = f"[{type_name}] " if type_name else ""

            summary = fields.get("summary", "")

            real_start = (
                fields.get(role_map.get("real_start"))
                if "real_start" in role_map
                else None
            )

            real_end = None
            if "real_end" in role_map:
                real_end = fields.get(role_map.get("real_end"))
            if not real_end:
                real_end = fields.get("resolutiondate")

            exp_start = (
                fields.get(role_map.get("exp_start"))
                if "exp_start" in role_map
                else None
            )

            exp_end = None
            if "exp_end" in role_map:
                exp_end = fields.get(role_map.get("exp_end"))
            if not exp_end:
                exp_end = fields.get("duedate")

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

            # Prefijo con Cliente, Tipo de Incidencia y Tag explícito [ESTADO: ...]
            subject = f"[{client_name}] {type_tag}[{k}] [ESTADO: {status_name.upper()}] ({sp_num:g} SP / {total_hours:g}h) {summary}"

            body_lines = [
                f"Cliente: {client_name}",
                f"Ticket: https://{domain}/browse/{k}",
                f"Tipo: {type_name or 'N/A'}",
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
        """Busca el valor de puntos dinámicamente sin IDs harcodeados, compatible con cualquier cliente."""
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
