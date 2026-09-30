import re
from dateutil import parser
import win32com.client
from src.utils.logger import logger
from src.utils.jira_utils import extract_story_points, extract_user_logged_hours
from src.utils.date_utils import parse_effective_issue_dates, format_all_day_range


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

    def _upsert_event(
        self,
        event_key: str,
        subject: str,
        body: str,
        start_str: str,
        end_str: str,
        category: str,
        events_map: dict,
    ) -> str:
        """Crea o actualiza un evento en Outlook si hubo cambios. Devuelve 'created', 'updated' o 'unchanged'."""
        if event_key in events_map:
            ev = events_map[event_key]
            has_changed = (
                getattr(ev, "Subject", "") != subject
                or getattr(ev, "Body", "").strip() != body.strip()
                or str(getattr(ev, "Start", ""))[:10] != start_str[:10]
                or str(getattr(ev, "End", ""))[:10] != end_str[:10]
                or getattr(ev, "Categories", "") != category
            )

            if has_changed:
                ev.Subject = subject
                ev.Start = start_str
                ev.End = end_str
                ev.AllDayEvent = True
                ev.ReminderSet = False
                ev.Body = body
                ev.Categories = category
                ev.Save()
                return "updated"
            return "unchanged"

        new_ev = self.outlook.CreateItem(1)
        new_ev.Subject = subject
        new_ev.Start = start_str
        new_ev.End = end_str
        new_ev.AllDayEvent = True
        new_ev.ReminderSet = False
        new_ev.Body = body
        new_ev.Categories = category
        new_ev.Save()
        return "created"

    def sync_board_issues(
        self,
        issues,
        sprints_data,
        field_mapping,
        board_config,
        domain,
        client_name: str,
        user_email: str = "",
    ):
        events_map = self.get_existing_events_map()
        role_map = {
            m["role"]: m["field_id"]
            for m in field_mapping
            if m.get("role") and m["role"] != "none"
        }

        created_count = 0
        updated_count = 0
        unchanged_count = 0

        # 1. Sincronización de Sprints
        for sprint_entry in sprints_data:
            result = self._sync_single_sprint(
                sprint_entry=sprint_entry,
                all_issues=issues,
                role_map=role_map,
                field_mapping=field_mapping,
                board_config=board_config,
                client_name=client_name,
                events_map=events_map,
            )
            if result == "created":
                created_count += 1
            elif result == "updated":
                updated_count += 1
            elif result == "unchanged":
                unchanged_count += 1

        # 2. Sincronización de Tickets Individuales
        for issue in issues:
            result = self._sync_single_issue(
                issue=issue,
                role_map=role_map,
                field_mapping=field_mapping,
                board_config=board_config,
                domain=domain,
                client_name=client_name,
                user_email=user_email,
                events_map=events_map,
            )
            if result == "created":
                created_count += 1
            elif result == "updated":
                updated_count += 1
            elif result == "unchanged":
                unchanged_count += 1

        return created_count, updated_count, unchanged_count

    def _is_issue_in_sprint(
        self,
        key: str,
        fields: dict,
        s_id_num: int,
        s_name: str,
        sprint_issues: list,
        sprint_project_prefixes: set,
        dt_s_start,
        dt_s_end,
        role_map: dict,
    ) -> bool:
        """Determina con guard clauses si un ticket pertenece al sprint evaluado."""
        status_name = fields.get("status", {}).get("name", "").strip().lower()
        if any(
            term in status_name
            for term in ["cancelado", "cancelled", "descartado", "rechazado"]
        ):
            return False

        sprint_raw_val = fields.get("customfield_10020")
        if not sprint_raw_val:
            sprint_raw_val = next(
                (v for k, v in fields.items() if "sprint" in k.lower() and v), ""
            )

        sprint_field_raw = str(sprint_raw_val or "").lower()
        explicitly_in_sprint = (str(s_id_num) in sprint_field_raw) or (
            s_name.lower() in sprint_field_raw
        )
        officially_in_sprint = any(it.get("key") == key for it in sprint_issues)

        res_str = fields.get("resolutiondate")
        real_end = (
            fields.get(role_map.get("real_end")) if "real_end" in role_map else None
        )
        real_start = (
            fields.get(role_map.get("real_start")) if "real_start" in role_map else None
        )
        exp_start = (
            fields.get(role_map.get("exp_start")) if "exp_start" in role_map else None
        )
        exp_end = (
            fields.get(role_map.get("exp_end"))
            if "exp_end" in role_map
            else fields.get("duedate")
        )

        effective_date_str = res_str or real_end or exp_end or real_start or exp_start
        dt_ticket = (
            parser.parse(effective_date_str).date() if effective_date_str else None
        )
        in_date_window = bool(dt_ticket and (dt_s_start <= dt_ticket <= dt_s_end))

        ticket_prefix = key.split("-")[0] if "-" in key else ""
        same_project = (
            ticket_prefix in sprint_project_prefixes
            if sprint_project_prefixes
            else True
        )

        if not explicitly_in_sprint and not officially_in_sprint:
            if not (same_project and in_date_window):
                return False

        close_date_str = res_str or real_end
        if close_date_str and parser.parse(close_date_str).date() < dt_s_start:
            return False

        return True

    def _aggregate_sprint_metrics(
        self,
        candidate_issues: dict,
        s_id_num: int,
        s_name: str,
        sprint_issues: list,
        sprint_project_prefixes: set,
        dt_s_start,
        dt_s_end,
        role_map: dict,
        field_mapping: list,
    ):
        sp_est, sp_comp = 0.0, 0.0
        closed_issues, pending_issues = [], []

        for k, issue in candidate_issues.items():
            fields = issue["fields"]
            if not self._is_issue_in_sprint(
                k,
                fields,
                s_id_num,
                s_name,
                sprint_issues,
                sprint_project_prefixes,
                dt_s_start,
                dt_s_end,
                role_map,
            ):
                continue

            status_obj = fields.get("status", {})
            status_cat = status_obj.get("statusCategory", {}).get("key", "")
            status_name = status_obj.get("name", "").strip()

            sp_num = extract_story_points(fields, role_map, field_mapping)
            issue_type_obj = fields.get("issuetype") or {}
            type_name = (
                issue_type_obj.get("name", "").upper()
                if isinstance(issue_type_obj, dict)
                else ""
            )
            type_badge = f"[{type_name}] " if type_name else ""
            summary = fields.get("summary", "")

            is_done = (status_cat == "done") or (
                status_name.lower() in ["cerrado", "closed", "resuelto", "finalizado"]
            )
            sp_est += sp_num

            if is_done:
                sp_comp += sp_num
                closed_issues.append(
                    f"  ✅ {type_badge}[{k}] ({sp_num:g} SP) {summary}"
                )
            else:
                pending_issues.append(
                    f"  ⏳ {type_badge}[{k}] ({sp_num:g} SP | {status_name}) {summary}"
                )

        return sp_est, sp_comp, closed_issues, pending_issues

    def _build_sprint_body(
        self,
        client_name: str,
        board_name: str,
        sprint_name: str,
        sprint_state: str,
        dt_start,
        dt_end,
        sp_comp: float,
        sp_est: float,
        hours_comp: float,
        hours_est: float,
        closed_issues: list,
        pending_issues: list,
    ) -> str:
        return "\n".join(
            [
                f"Cliente: {client_name}",
                f"Tablero: {board_name}",
                f"Sprint: {sprint_name}",
                f"Estado del Sprint: {sprint_state.upper()}",
                f"Fechas: {dt_start} al {dt_end}",
                f"Mis Story Points : {sp_comp:g} finalizados / {sp_est:g} estimados",
                f"Mis Horas Totales: {hours_comp:g}h finalizadas / {hours_est:g}h estimadas",
                "====================================",
                f"Mis Tickets Cerrados en este Sprint ({len(closed_issues)}):",
                (
                    "\n".join(closed_issues)
                    if closed_issues
                    else "  (Ninguno cerrado en este período)"
                ),
                "------------------------------------",
                f"Mis Tickets Pendientes / En Curso ({len(pending_issues)}):",
                (
                    "\n".join(pending_issues)
                    if pending_issues
                    else "  (Todos completados)"
                ),
            ]
        )

    def _sync_single_sprint(
        self,
        sprint_entry: dict,
        all_issues: list,
        role_map: dict,
        field_mapping: list,
        board_config,
        client_name: str,
        events_map: dict,
    ) -> str:
        sprint = sprint_entry["meta"]
        s_start, s_end = sprint.get("startDate"), sprint.get("endDate") or sprint.get(
            "completeDate"
        )
        if not s_start or not s_end:
            return "skipped"

        s_id_num, s_name = sprint.get("id"), sprint.get("name", "Sprint")
        s_id, s_state = f"SPRINT-{s_id_num}", sprint.get("state", "").lower()
        dt_s_start, dt_s_end = parser.parse(s_start).date(), parser.parse(s_end).date()

        candidate_issues = {it["key"]: it for it in sprint_entry.get("issues", [])}
        for it in all_issues:
            candidate_issues[it["key"]] = it

        prefixes = {
            it["key"].split("-")[0]
            for it in sprint_entry.get("issues", [])
            if "-" in it.get("key", "")
        }

        my_sp_est, my_sp_comp, closed_issues, pending_issues = (
            self._aggregate_sprint_metrics(
                candidate_issues,
                s_id_num,
                s_name,
                sprint_entry.get("issues", []),
                prefixes,
                dt_s_start,
                dt_s_end,
                role_map,
                field_mapping,
            )
        )

        hours_per_sp = board_config.hours_per_sp or 4
        my_h_comp, my_h_est = my_sp_comp * hours_per_sp, my_sp_est * hours_per_sp
        pct = int((my_sp_comp / my_sp_est) * 100) if my_sp_est else 0
        tag = "CERRADO" if s_state in ["closed", "cerrado"] else "ACTIVO"

        subject = f"[{client_name}] [{s_id}] [ESTADO: {tag}] {s_name} | Mis SP: {my_sp_comp:g}/{my_sp_est:g} ({pct}%) [{my_h_comp:g}h/{my_h_est:g}h]"
        body = self._build_sprint_body(
            client_name,
            board_config.board_name,
            s_name,
            s_state,
            dt_s_start,
            dt_s_end,
            my_sp_comp,
            my_sp_est,
            my_h_comp,
            my_h_est,
            closed_issues,
            pending_issues,
        )

        category = self.resolve_category_for_item("sprint", s_state)
        start_str, end_str = format_all_day_range(dt_s_start, dt_s_end)

        result = self._upsert_event(
            s_id, subject, body, start_str, end_str, category, events_map
        )
        if result == "updated":
            logger.info(f"🔄 Sprint personal actualizado: {s_name}")
        elif result == "created":
            logger.info(f"✅ Sprint personal creado en calendario: {s_name}")

        return result

    def _build_issue_effort_and_subject(
        self,
        client_name: str,
        key: str,
        summary: str,
        status_name: str,
        type_name: str,
        sp_num: float,
        hours_per_sp: float,
        logged_hours: float,
    ):
        total_hours = sp_num * hours_per_sp
        if total_hours > 0:
            hours_str = f"{sp_num:g} SP / {total_hours:g}h"
        elif logged_hours > 0:
            hours_str = f"{logged_hours:g}h logueadas"
        else:
            hours_str = "0 SP / 0h"

        type_tag = f"[{type_name}] " if type_name else ""
        subject = f"[{client_name}] {type_tag}[{key}] [ESTADO: {status_name.upper()}] ({hours_str}) {summary}"
        return total_hours, subject

    def _build_issue_body(
        self,
        client_name: str,
        domain: str,
        key: str,
        summary: str,
        status_name: str,
        type_name: str,
        sp_num: float,
        total_hours: float,
        logged_hours: float,
        raw_dates: dict,
        fields: dict,
        field_mapping: list,
    ) -> str:
        body_lines = [
            f"Cliente: {client_name}",
            f"Ticket: https://{domain}/browse/{key}",
            f"Tipo: {type_name or 'N/A'}",
            f"Resumen: {summary}",
            f"Estado: {status_name}",
            f"Story Points: {sp_num:g} ({total_hours:g} hs de esfuerzo estimado)",
            f"Horas registradas (Worklog): {logged_hours:g}h",
            "------------------------------------",
            f"Inicio Real     : {raw_dates.get('real_start') or 'Pendiente'}",
            f"Fin Real        : {raw_dates.get('real_end') or 'Pendiente'}",
            f"Inicio Estimado : {raw_dates.get('exp_start') or 'N/A'}",
            f"Fin Estimado    : {raw_dates.get('exp_end') or 'N/A'}",
        ]

        excluded_roles = {
            "real_start",
            "real_end",
            "exp_start",
            "exp_end",
            "story_points",
        }
        extra_lines = [
            f"{m['field_name']}: {fields.get(m['field_id'])}"
            for m in field_mapping
            if m.get("include_in_body")
            and m.get("role") not in excluded_roles
            and fields.get(m["field_id"])
        ]

        if extra_lines:
            body_lines.append("------------------------------------")
            body_lines.extend(extra_lines)

        return "\n".join(body_lines)

    def _sync_single_issue(
        self,
        issue: dict,
        role_map: dict,
        field_mapping: list,
        board_config,
        domain: str,
        client_name: str,
        user_email: str,
        events_map: dict,
    ) -> str:
        k = issue["key"]
        fields = issue["fields"]
        status_name = fields.get("status", {}).get("name", "Desconocido")
        status_lower = status_name.lower().strip()

        if any(
            term in status_lower
            for term in ["cancelado", "cancelled", "descartado", "rechazado"]
        ):
            if k in events_map:
                try:
                    events_map[k].Delete()
                except Exception:
                    pass
            return "skipped"

        dt_start, dt_end, raw_dates = parse_effective_issue_dates(fields, role_map)
        if not dt_end:
            return "skipped"

        issue_type_obj = fields.get("issuetype") or {}
        type_name = (
            issue_type_obj.get("name", "").upper()
            if isinstance(issue_type_obj, dict)
            else ""
        )
        summary = fields.get("summary", "")

        sp_num = extract_story_points(fields, role_map, field_mapping)
        logged_hours = extract_user_logged_hours(fields, user_email)
        hours_per_sp = board_config.hours_per_sp or 4

        total_hours, subject = self._build_issue_effort_and_subject(
            client_name,
            k,
            summary,
            status_name,
            type_name,
            sp_num,
            hours_per_sp,
            logged_hours,
        )

        body_content = self._build_issue_body(
            client_name,
            domain,
            k,
            summary,
            status_name,
            type_name,
            sp_num,
            total_hours,
            logged_hours,
            raw_dates,
            fields,
            field_mapping,
        )

        issue_category = self.resolve_category_for_item("ticket", status_name)
        start_str, end_str = format_all_day_range(dt_start, dt_end)

        result = self._upsert_event(
            k, subject, body_content, start_str, end_str, issue_category, events_map
        )
        if result == "updated":
            logger.info(
                f"🔄 Ticket actualizado: {k} -> {dt_start} (Estado: {status_name})"
            )
        elif result == "created":
            logger.info(f"✅ Ticket creado en calendario: {k} en {dt_start}")

        return result
