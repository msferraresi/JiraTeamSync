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
                match = re.match(r"^\[([A-Z0-9_\-]+)\]", subj)
                if match:
                    events_map[match.group(1)] = item
            except Exception:
                continue
        return events_map

    def sync_board_issues(
        self, issues, sprints_data, field_mapping, board_config, domain
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

                # SP del ticket
                sp_fid = role_map.get("story_points")
                sp_val = fields.get(sp_fid) if sp_fid else None
                try:
                    sp_num = float(sp_val) if sp_val is not None else 0.0
                except (ValueError, TypeError):
                    sp_num = 0.0

                # Validar pertenencia cronológica al Sprint
                is_done = status_cat == "done"

                if is_done:
                    # Si el ticket está cerrado, verificar cuándo se cerró
                    if resolution_date_str:
                        dt_res = parser.parse(resolution_date_str).date()
                        # Solo se cuenta como cerrado en este sprint si se cerró dentro de su ventana temporal
                        if dt_s_start <= dt_res <= dt_s_end:
                            my_sp_estimated += sp_num
                            my_sp_completed += sp_num
                            my_closed_issues.append(
                                f"  ✅ [{k}] ({sp_num:g} SP) {summary}"
                            )
                        elif s_state == "closed" and dt_res <= dt_s_end:
                            # Caso borde: cerrado antes o durante un sprint ya cerrado
                            # Si no pertenecía a este rango, se ignora para no duplicar en sprint nuevo
                            pass
                    else:
                        # Si está cerrado pero no tiene resolutiondate
                        if s_state == "closed":
                            my_sp_estimated += sp_num
                            my_sp_completed += sp_num
                            my_closed_issues.append(
                                f"  ✅ [{k}] ({sp_num:g} SP) {summary}"
                            )
                else:
                    # Ticket pendiente/en progreso en este Sprint
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

            sprint_subject = f"[{s_id}] {s_name} | Mis SP: {my_sp_completed:g}/{my_sp_estimated:g} ({pct}%) [{my_hours_completed:g}h/{my_hours_estimated:g}h]"

            sprint_body = [
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

            start_str = dt_s_start.strftime("%Y-%m-%d 00:00")
            end_str = dt_s_end_inclusive.strftime("%Y-%m-%d 00:00")

            if s_id in events_map:
                ev = events_map[s_id]
                if (
                    getattr(ev, "Subject", "") != sprint_subject
                    or getattr(ev, "Body", "").strip() != sprint_body_content.strip()
                    or str(getattr(ev, "Start", ""))[:10] != start_str[:10]
                    or str(getattr(ev, "End", ""))[:10] != end_str[:10]
                ):
                    ev.Subject = sprint_subject
                    ev.Start = start_str
                    ev.End = end_str
                    ev.AllDayEvent = True
                    ev.ReminderSet = False
                    ev.Body = sprint_body_content
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

            sp_fid = role_map.get("story_points")
            sp_val = fields.get(sp_fid) if sp_fid else None
            try:
                sp_num = float(sp_val) if sp_val is not None else 0.0
            except (ValueError, TypeError):
                sp_num = 0.0

            hours_per_sp = board_config.hours_per_sp or 4
            total_hours = sp_num * hours_per_sp

            subject = (
                f"[{k}] ({status_name} | {sp_num:g} SP / {total_hours:g}h) {summary}"
            )

            body_lines = [
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
            start_str = dt_start.strftime("%Y-%m-%d 00:00")
            end_str = dt_end_inclusive.strftime("%Y-%m-%d 00:00")

            if k in events_map:
                ev = events_map[k]
                if (
                    getattr(ev, "Subject", "") != subject
                    or getattr(ev, "Body", "").strip() != body_content.strip()
                    or str(getattr(ev, "Start", ""))[:10] != start_str[:10]
                    or str(getattr(ev, "End", ""))[:10] != end_str[:10]
                ):
                    ev.Subject = subject
                    ev.Start = start_str
                    ev.End = end_str
                    ev.AllDayEvent = True
                    ev.ReminderSet = False
                    ev.Body = body_content
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
                new_event.Save()
                created_count += 1
                logger.info(f"✅ Ticket creado en calendario: {k} en {dt_start}")

        return created_count, updated_count, unchanged_count
