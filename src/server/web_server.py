import os
import sys
import json
import uvicorn
import requests
from typing import Optional, List, Dict, Any
from fastapi import FastAPI, HTTPException, Depends, Query
from fastapi.responses import HTMLResponse
from pydantic import BaseModel, Field
from sqlalchemy.orm import Session

from src.db.database import SessionLocal, init_db
from src.db.models import Client, BoardConfig, AppSettings
from src.utils.logger import LOG_FILE

app = FastAPI(title="JiraTeamsSync Admin")


def get_db():
    db = SessionLocal()
    try:
        yield db
    finally:
        db.close()


# --- Schemas ---
class AppSettingsSchema(BaseModel):
    interval_minutes: int
    server_port: int


class ClientSchema(BaseModel):
    name: str = Field(..., min_length=1)
    domain: str = Field(..., min_length=1)
    email: str = Field(..., min_length=1)
    api_token: str = Field(..., min_length=1)
    is_active: bool = True
    time_window_start: str = "07:00"
    time_window_end: str = "20:00"
    active_weekdays_only: bool = True
    vpn_check_host: Optional[str] = None
    sync_interval_minutes: int = 60


class BoardConfigSchema(BaseModel):
    id: Optional[int] = None
    client_id: int
    board_id: Optional[int] = None
    board_name: str
    custom_jql: str
    hours_per_sp: int = 4
    fields_mapping: List[Dict[str, Any]] = []


# --- Endpoints App Settings ---
@app.get("/api/settings")
def get_settings(db: Session = Depends(get_db)):
    s = db.query(AppSettings).first() or AppSettings()
    return {
        "interval_minutes": s.interval_minutes or 15,
        "server_port": s.server_port or 8765,
    }


@app.post("/api/settings")
def update_settings(data: AppSettingsSchema, db: Session = Depends(get_db)):
    s = db.query(AppSettings).first()
    if not s:
        s = AppSettings()
        db.add(s)
    s.interval_minutes = data.interval_minutes
    s.server_port = data.server_port
    db.commit()
    return {"success": True}


# --- Endpoints Clientes ---
@app.get("/api/clients")
def list_clients(db: Session = Depends(get_db)):
    clients = db.query(Client).all()
    res = []
    for c in clients:
        if not c.name and not c.domain:
            continue
        res.append(
            {
                "id": c.id,
                "name": c.name or "Sin Nombre",
                "domain": c.domain or "",
                "email": c.email or "",
                "is_active": c.is_active if c.is_active is not None else True,
                "time_window": f"{c.time_window_start or '07:00'} - {c.time_window_end or '20:00'}",
                "sync_interval_minutes": getattr(c, "sync_interval_minutes", 60) or 60,
                "boards_count": len(c.boards) if c.boards else 0,
            }
        )
    return res


@app.post("/api/clients")
def create_client(c_data: ClientSchema, db: Session = Depends(get_db)):
    if not c_data.name.strip() or not c_data.domain.strip():
        raise HTTPException(400, "Nombre y Dominio son obligatorios")
    c = Client(
        name=c_data.name.strip(),
        domain=c_data.domain.strip()
        .replace("https://", "")
        .replace("http://", "")
        .rstrip("/"),
        email=c_data.email.strip(),
        api_token=c_data.api_token.strip(),
        time_window_start=c_data.time_window_start,
        time_window_end=c_data.time_window_end,
        sync_interval_minutes=c_data.sync_interval_minutes,
        vpn_check_host=c_data.vpn_check_host.strip() if c_data.vpn_check_host else None,
    )
    db.add(c)
    db.commit()
    return {"success": True}


@app.delete("/api/clients/{client_id}")
def delete_client(client_id: int, db: Session = Depends(get_db)):
    c = db.query(Client).filter(Client.id == client_id).first()
    if c:
        db.delete(c)
        db.commit()
    return {"success": True}


# --- Endpoints Introspección Jira ---
@app.get("/api/clients/{client_id}/jira-boards")
def get_jira_boards(client_id: int, db: Session = Depends(get_db)):
    client = db.query(Client).filter(Client.id == client_id).first()
    if not client:
        raise HTTPException(404, "Cliente no encontrado")
    auth = (client.email, client.api_token)
    all_boards, start_at = [], 0
    try:
        while True:
            r = requests.get(
                f"https://{client.domain}/rest/agile/1.0/board?startAt={start_at}&maxResults=50",
                auth=auth,
                timeout=10,
            )
            r.raise_for_status()
            data = r.json()
            vals = data.get("values", [])
            all_boards.extend(vals)
            if data.get("isLast", True) or not vals:
                break
            start_at += 50
        return sorted(all_boards, key=lambda x: str(x.get("name", "")).lower())
    except Exception as e:
        raise HTTPException(400, f"Error consultando Jira: {str(e)}")


@app.get("/api/clients/{client_id}/jira-fields")
def get_jira_fields(client_id: int, db: Session = Depends(get_db)):
    client = db.query(Client).filter(Client.id == client_id).first()
    if not client:
        raise HTTPException(404, "Cliente no encontrado")
    try:
        r = requests.get(
            f"https://{client.domain}/rest/api/3/field",
            auth=(client.email, client.api_token),
            timeout=10,
        )
        r.raise_for_status()
        fields = r.json()
        result = []
        for f in fields:
            t = f.get("schema", {}).get("type", "custom")
            result.append({"id": f.get("id", ""), "name": f.get("name", ""), "type": t})
        return sorted(result, key=lambda x: str(x["name"]).lower())
    except Exception as e:
        raise HTTPException(400, f"Error obteniendo campos: {str(e)}")


# --- Endpoints Tableros ---
@app.get("/api/clients/{client_id}/boards")
def list_boards(client_id: int, db: Session = Depends(get_db)):
    boards = db.query(BoardConfig).filter(BoardConfig.client_id == client_id).all()
    res = []
    for b in boards:
        res.append(
            {
                "id": b.id,
                "board_id": b.board_id,
                "board_name": b.board_name,
                "custom_jql": b.custom_jql,
                "hours_per_sp": b.hours_per_sp,
                "fields_mapping": json.loads(b.fields_mapping_json or "[]"),
            }
        )
    return res


@app.post("/api/boards")
def save_or_update_board(data: BoardConfigSchema, db: Session = Depends(get_db)):
    if data.id:
        board = db.query(BoardConfig).filter(BoardConfig.id == data.id).first()
        if not board:
            raise HTTPException(404, "Tablero no encontrado")
        board.board_id = data.board_id
        board.board_name = data.board_name
        board.custom_jql = data.custom_jql
        board.hours_per_sp = data.hours_per_sp
        board.fields_mapping_json = json.dumps(data.fields_mapping)
    else:
        board = BoardConfig(
            client_id=data.client_id,
            board_id=data.board_id,
            board_name=data.board_name,
            custom_jql=data.custom_jql,
            hours_per_sp=data.hours_per_sp,
            fields_mapping_json=json.dumps(data.fields_mapping),
        )
        db.add(board)
    db.commit()
    return {"success": True}


@app.delete("/api/boards/{board_id}")
def delete_board(board_id: int, db: Session = Depends(get_db)):
    b = db.query(BoardConfig).filter(BoardConfig.id == board_id).first()
    if b:
        db.delete(b)
        db.commit()
    return {"success": True}


# --- Endpoint de Logs con Filtro ---
@app.get("/api/logs")
def get_logs(filter_text: str = Query("", alias="filter"), limit: int = 300):
    if not os.path.exists(LOG_FILE):
        return {"logs": ["El archivo de log aún no se ha generado."]}
    try:
        with open(LOG_FILE, "r", encoding="utf-8", errors="replace") as f:
            lines = f.readlines()
        if filter_text.strip():
            q = filter_text.lower().strip()
            lines = [line for line in lines if q in line.lower()]
        return {"logs": lines[-limit:]}
    except Exception as e:
        return {"logs": [f"Error leyendo logs: {str(e)}"]}


# --- UI HTML Multisolapa ---
HTML_TEMPLATE = """
<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8">
    <title>Jira Teams Sync - Panel de Control</title>
    <style>
        :root {
            --bg: #0f172a; --card-bg: #1e293b; --border: #334155;
            --text: #f8fafc; --muted: #94a3b8; --primary: #0284c7;
            --primary-hover: #0369a1; --accent: #38bdf8; --danger: #be123c;
        }
        body { font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif; background: var(--bg); color: var(--text); margin: 0; padding: 25px; }
        .container { max-width: 1100px; margin: 0 auto; }
        h1 { color: var(--accent); margin-bottom: 20px; font-size: 1.6rem; }
        h2, h3, h4 { color: var(--accent); margin-top: 0; }
        .card { background: var(--card-bg); border: 1px solid var(--border); border-radius: 8px; padding: 20px; margin-bottom: 25px; box-shadow: 0 4px 6px -1px rgba(0,0,0,0.2); }
        .grid { display: grid; grid-template-columns: 1fr 1fr; gap: 15px; }
        .grid-3 { display: grid; grid-template-columns: repeat(3, 1fr); gap: 15px; }
        label { display: block; font-size: 12px; color: var(--muted); font-weight: bold; margin-top: 10px; }
        input, select, textarea { width: 100%; padding: 8px; background: #0f172a; border: 1px solid var(--border); color: #fff; border-radius: 4px; box-sizing: border-box; }
        button { background: var(--primary); color: white; border: none; padding: 8px 14px; border-radius: 4px; cursor: pointer; font-weight: bold; }
        button:hover { background: var(--primary-hover); }
        .btn-edit { background: #0d9488; margin-right: 5px; }
        .btn-edit:hover { background: #0f766e; }
        .btn-danger { background: var(--danger); }
        .btn-danger:hover { background: #9f1239; }
        
        /* Solapas */
        .tabs { display: flex; gap: 5px; border-bottom: 2px solid var(--border); margin-bottom: 20px; }
        .tab-btn { background: transparent; border: none; color: var(--muted); font-size: 15px; font-weight: bold; padding: 10px 18px; border-bottom: 3px solid transparent; cursor: pointer; border-radius: 0; }
        .tab-btn:hover { color: var(--text); background: transparent; }
        .tab-btn.active { color: var(--accent); border-bottom-color: var(--accent); }
        .tab-content { display: none; }
        .tab-content.active { display: block; }

        table { width: 100%; border-collapse: collapse; margin-top: 15px; }
        th, td { border: 1px solid var(--border); padding: 8px 10px; text-align: left; font-size: 13px; }
        th { background: #0f172a; color: var(--muted); }
        .badge { display: inline-block; padding: 2px 6px; border-radius: 4px; font-size: 11px; background: #065f46; color: #6ee7b7; }
        .table-scroll { max-height: 400px; overflow-y: auto; border: 1px solid var(--border); margin-top: 10px; }
        .search-box { margin-bottom: 8px; }

        /* Visor de logs */
        .log-box { background: #050b14; border: 1px solid var(--border); border-radius: 4px; font-family: 'Consolas', monospace; font-size: 12px; height: 500px; overflow-y: auto; padding: 15px; white-space: pre-wrap; color: #e2e8f0; line-height: 1.4; }
    </style>
</head>
<body>
<div class="container">
    <h1>⚙️ Panel de Control - Jira Teams Sync</h1>

    <div class="tabs">
        <button class="tab-btn active" onclick="switchTab('tab-general', this)">⚙️ General</button>
        <button class="tab-btn" onclick="switchTab('tab-connections', this)">🏢 Conexiones y Clientes</button>
        <button class="tab-btn" onclick="switchTab('tab-logs', this)">📄 Logs de Ejecución</button>
    </div>

    <!-- SOLAPA 1: GENERAL -->
    <div id="tab-general" class="tab-content active">
        <div class="card">
            <h2>⏱️ Frecuencia Global del Bucle</h2>
            <div class="grid">
                <div><label>Frecuencia de ciclo global (minutos):</label><input type="number" id="app_interval" value="15"></div>
                <div><label>Puerto Web UI:</label><input type="number" id="app_port" value="8765"></div>
            </div>
            <button style="margin-top:15px;" onclick="saveAppSettings()">Guardar Frecuencia</button>
        </div>
    </div>

    <!-- SOLAPA 2: CONEXIONES Y TABLEROS -->
    <div id="tab-connections" class="tab-content">
        <div class="card">
            <h2>🏢 Clientes Configurados</h2>
            <div id="clientsList">Cargando...</div>

            <h3 style="margin-top: 25px;">➕ Agregar Cliente</h3>
            <div class="grid-3">
                <div>
                    <label>Nombre Empresa / Cliente:</label>
                    <input type="text" id="c_name" placeholder="Ej: MiEmpresa" required>
                </div>
                <div>
                    <label>Dominio Jira:</label>
                    <input type="text" id="c_domain" placeholder="ejemplo.atlassian.net" required>
                </div>
                <div>
                    <label>Email:</label>
                    <input type="email" id="c_email" placeholder="usuario@dominio.com" required>
                </div>
            </div>
            <div class="grid-3">
                <div>
                    <label>Host VPN (opcional):</label>
                    <input type="text" id="c_vpn" placeholder="vpn.ejemplo.com">
                </div>
                <div>
                    <label>Horario Operativo (HH:MM):</label>
                    <div style="display:flex; gap:10px;">
                        <input type="text" id="c_start" value="07:00">
                        <input type="text" id="c_end" value="20:00">
                    </div>
                </div>
                <div>
                    <label>Intervalo de Sincro (minutos):</label>
                    <input type="number" id="c_interval" value="60">
                </div>
            </div>
            <div>
                <label>API Token:</label>
                <input type="password" id="c_token" placeholder="Token de Jira" required>
            </div>
            <button style="margin-top:15px;" onclick="createClient()">Guardar Cliente</button>
        </div>

        <div class="card" id="boardsSection" style="display:none;">
            <h2 id="currentClientTitle">📋 Tableros del Cliente</h2>
            <div id="boardsList"></div>

            <h3 id="boardFormTitle" style="margin-top:25px;">➕ Configurar Tablero</h3>
            <input type="hidden" id="edit_board_db_id" value="">

            <div style="margin-bottom:15px;">
                <button onclick="fetchJiraCatalog()" style="background:#475569;">🔄 Traer / Actualizar Catálogo de Campos de Jira</button>
            </div>

            <div class="grid">
                <div>
                    <label>Tablero de Jira:</label>
                    <select id="b_select" onchange="onBoardSelect()"><option value="">-- Seleccionar --</option></select>
                    <label>Nombre Identificador:</label><input type="text" id="b_name">
                </div>
                <div>
                    <label>JQL Filter:</label><input type="text" id="b_jql" value="assignee = currentUser() AND statusCategory != Done">
                    <label>Horas por Story Point:</label><input type="number" id="b_hours" value="4">
                </div>
            </div>

            <h4 style="margin-top:20px;">🗂️ Mapeo y Uso de Campos de Jira</h4>
            <input type="text" class="search-box" id="fieldSearch" placeholder="🔍 Buscar campo por nombre..." onkeyup="filterFieldsTable()">

            <div class="table-scroll">
                <table id="fieldsMappingTable">
                    <thead>
                        <tr>
                            <th style="width:30%;">Campo en Jira</th>
                            <th style="width:20%;">ID Técnico</th>
                            <th style="width:30%;">Rol / Uso en la Sincronización</th>
                            <th style="width:20%; text-align:center;">¿Incluir en Outlook?</th>
                        </tr>
                    </thead>
                    <tbody id="fieldsTbody">
                        <tr><td colspan="4" style="text-align:center;">Haz clic en 'Traer / Actualizar Catálogo' para ver los campos.</td></tr>
                    </tbody>
                </table>
            </div>

            <div style="margin-top:15px; display:flex; gap:10px;">
                <button onclick="saveBoardConfig()">💾 Guardar Configuración de Tablero</button>
                <button onclick="cancelBoardEdit()" style="background:#64748b;">Cancelar</button>
            </div>
        </div>
    </div>

    <!-- SOLAPA 3: LOGS -->
    <div id="tab-logs" class="tab-content">
        <div class="card">
            <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:10px;">
                <h2>📄 Logs de Ejecución</h2>
                <div style="display:flex; gap:10px; width:60%;">
                    <input type="text" id="logFilter" placeholder="Filtrar por cliente, ticket o error..." onkeyup="filterLogs()">
                    <button onclick="loadLogs()">Refrescar</button>
                </div>
            </div>
            <div class="log-box" id="logBox">Cargando registros...</div>
        </div>
    </div>
</div>

<script>
    let activeClientId = null;
    let jiraFieldsCatalog = [];
    let currentBoardMapping = [];
    let logTimer = null;

    function switchTab(tabId, btn) {
        document.querySelectorAll('.tab-content').forEach(el => el.classList.remove('active'));
        document.querySelectorAll('.tab-btn').forEach(el => el.classList.remove('active'));
        document.getElementById(tabId).classList.add('active');
        btn.classList.add('active');

        if(tabId === 'tab-logs') {
            loadLogs();
        }
    }

    function loadAppSettings() {
        fetch('/api/settings').then(r => r.json()).then(d => {
            document.getElementById('app_interval').value = d.interval_minutes;
            document.getElementById('app_port').value = d.server_port;
        }).catch(e => console.error("Error settings:", e));
    }

    function saveAppSettings() {
        fetch('/api/settings', {
            method: 'POST', headers: {'Content-Type':'application/json'},
            body: JSON.stringify({
                interval_minutes: parseInt(document.getElementById('app_interval').value)||15,
                server_port: parseInt(document.getElementById('app_port').value)||8765
            })
        }).then(() => alert('Frecuencia guardada.'));
    }

    function loadClients() {
        fetch('/api/clients')
            .then(r => r.json())
            .then(data => {
                if(!data || !data.length) { 
                    document.getElementById('clientsList').innerHTML = '<p>No hay clientes configurados.</p>'; 
                    return; 
                }
                let h = '<table><tr><th>Nombre</th><th>Dominio</th><th>Horario</th><th>Intervalo</th><th>Tableros</th><th>Acciones</th></tr>';
                data.forEach(c => {
                    const cName = c.name || 'Sin Nombre';
                    h += `<tr>
                        <td><strong>${cName}</strong></td>
                        <td>${c.domain || ''}</td>
                        <td><span class="badge">${c.time_window || ''}</span></td>
                        <td>${c.sync_interval_minutes || 60} min</td>
                        <td>${c.boards_count || 0}</td>
                        <td>
                            <button onclick="manageBoards(${c.id}, '${cName}')">Gestionar Tableros</button>
                            <button class="btn-danger" onclick="deleteClient(${c.id})">Eliminar</button>
                        </td>
                    </tr>`;
                });
                document.getElementById('clientsList').innerHTML = h + '</table>';
            })
            .catch(err => {
                console.error("Error al cargar clientes:", err);
                document.getElementById('clientsList').innerHTML = '<p style="color:#f87171;">Error al cargar clientes.</p>';
            });
    }

    function createClient() {
        const name = document.getElementById('c_name').value.trim();
        const domain = document.getElementById('c_domain').value.trim();
        const email = document.getElementById('c_email').value.trim();
        const token = document.getElementById('c_token').value.trim();

        if(!name || !domain || !email || !token) {
            alert('Completa los campos requeridos (Nombre, Dominio, Email y Token).');
            return;
        }

        const payload = {
            name: name, domain: domain, email: email, api_token: token,
            time_window_start: document.getElementById('c_start').value, 
            time_window_end: document.getElementById('c_end').value,
            sync_interval_minutes: parseInt(document.getElementById('c_interval').value) || 60,
            vpn_check_host: document.getElementById('c_vpn').value.trim() || null
        };

        fetch('/api/clients', { 
            method: 'POST', 
            headers: {'Content-Type':'application/json'}, 
            body: JSON.stringify(payload) 
        })
        .then(r => {
            if(!r.ok) throw new Error("Error al guardar cliente");
            return r.json();
        })
        .then(() => { 
            document.getElementById('c_name').value = '';
            document.getElementById('c_domain').value = '';
            document.getElementById('c_email').value = '';
            document.getElementById('c_token').value = '';
            document.getElementById('c_vpn').value = '';
            loadClients(); 
            alert('Cliente guardado con éxito.'); 
        })
        .catch(err => alert('No se pudo guardar el cliente: ' + err.message));
    }

    function deleteClient(id) {
        if(confirm('¿Eliminar este cliente y todos sus tableros?')) {
            fetch('/api/clients/' + id, { method: 'DELETE' }).then(() => {
                if(activeClientId === id) {
                    document.getElementById('boardsSection').style.display = 'none';
                    activeClientId = null;
                }
                loadClients();
            });
        }
    }

    function manageBoards(id, name) {
        activeClientId = id;
        document.getElementById('boardsSection').style.display = 'block';
        document.getElementById('currentClientTitle').innerText = '📋 Tableros de: ' + name;
        cancelBoardEdit();
        loadBoards();
        fetchJiraCatalog();
        document.getElementById('boardsSection').scrollIntoView({ behavior: 'smooth' });
    }

    function loadBoards() {
        fetch(`/api/clients/${activeClientId}/boards`)
            .then(r => r.json())
            .then(boards => {
                if(!boards.length) {
                    document.getElementById('boardsList').innerHTML = '<p>Sin tableros configurados para este cliente.</p>';
                    return;
                }
                let h = '<table><tr><th>Tablero</th><th>Filtro JQL</th><th>Horas/SP</th><th>Acciones</th></tr>';
                boards.forEach(b => {
                    const bStr = encodeURIComponent(JSON.stringify(b));
                    h += `<tr>
                        <td><strong>${b.board_name}</strong></td>
                        <td><code>${b.custom_jql}</code></td>
                        <td>${b.hours_per_sp} hs/SP</td>
                        <td>
                            <button class="btn-edit" onclick="editBoard('${bStr}')">✏️ Editar</button>
                            <button class="btn-danger" onclick="deleteBoard(${b.id})">Eliminar</button>
                        </td>
                    </tr>`;
                });
                document.getElementById('boardsList').innerHTML = h + '</table>';
            });
    }

    function fetchJiraCatalog() {
        fetch(`/api/clients/${activeClientId}/jira-boards`)
            .then(r => r.json())
            .then(boards => {
                const sel = document.getElementById('b_select');
                sel.innerHTML = '<option value="">-- Seleccionar --</option>';
                boards.forEach(b => {
                    sel.innerHTML += `<option value="${b.id}" data-name="${b.name}">${b.name}</option>`;
                });
            })
            .catch(e => console.error("Error al traer boards:", e));

        fetch(`/api/clients/${activeClientId}/jira-fields`)
            .then(r => r.json())
            .then(fields => {
                jiraFieldsCatalog = fields || [];
                renderFieldsTable();
            })
            .catch(e => console.error("Error al traer fields:", e));
    }

    function renderFieldsTable() {
        const tbody = document.getElementById('fieldsTbody');
        if(!jiraFieldsCatalog.length) { 
            tbody.innerHTML = '<tr><td colspan="4" style="text-align:center;">No se cargaron campos de Jira todavía.</td></tr>'; 
            return; 
        }
        
        let h = '';
        jiraFieldsCatalog.forEach(f => {
            const existing = currentBoardMapping.find(m => m.field_id === f.id);
            const role = existing ? existing.role : 'none';
            const checked = existing && existing.include_in_body ? 'checked' : '';
            const fNameSafe = (f.name || '').toLowerCase();

            h += `<tr class="field-row" data-name="${fNameSafe}">
                <td><strong>${f.name}</strong></td>
                <td><code>${f.id}</code></td>
                <td>
                    <select class="field-role-select" data-fid="${f.id}" data-fname="${f.name}">
                        <option value="none" ${role==='none'?'selected':''}>-- No usar para cálculo --</option>
                        <option value="real_start" ${role==='real_start'?'selected':''}>🟢 Fecha Inicio Real (Start date)</option>
                        <option value="real_end" ${role==='real_end'?'selected':''}>🔴 Fecha Fin Real (End date)</option>
                        <option value="exp_start" ${role==='exp_start'?'selected':''}>📅 Fecha Inicio Esperada</option>
                        <option value="exp_end" ${role==='exp_end'?'selected':''}>📅 Fecha Fin Esperada</option>
                        <option value="story_points" ${role==='story_points'?'selected':''}>🔢 Story Points / Puntos</option>
                    </select>
                </td>
                <td style="text-align:center;">
                    <input type="checkbox" class="field-body-check" data-fid="${f.id}" ${checked}>
                </td>
            </tr>`;
        });
        tbody.innerHTML = h;
    }

    function filterFieldsTable() {
        const q = (document.getElementById('fieldSearch').value || '').toLowerCase();
        document.querySelectorAll('.field-row').forEach(tr => {
            const rowName = tr.getAttribute('data-name') || '';
            tr.style.display = rowName.includes(q) ? '' : 'none';
        });
    }

    function onBoardSelect() {
        const sel = document.getElementById('b_select');
        const opt = sel.options[sel.selectedIndex];
        if(opt && opt.value) {
            document.getElementById('b_name').value = opt.getAttribute('data-name') || '';
        }
    }

    function editBoard(bJsonStr) {
        const b = JSON.parse(decodeURIComponent(bJsonStr));
        document.getElementById('edit_board_db_id').value = b.id;
        document.getElementById('boardFormTitle').innerText = '✏️ Editando Tablero: ' + b.board_name;
        document.getElementById('b_name').value = b.board_name;
        document.getElementById('b_jql').value = b.custom_jql;
        document.getElementById('b_hours').value = b.hours_per_sp;
        document.getElementById('b_select').value = b.board_id || '';
        
        currentBoardMapping = b.fields_mapping || [];
        renderFieldsTable();
        document.getElementById('boardFormTitle').scrollIntoView({ behavior: 'smooth' });
    }

    function cancelBoardEdit() {
        document.getElementById('edit_board_db_id').value = '';
        document.getElementById('boardFormTitle').innerText = '➕ Configurar Tablero';
        document.getElementById('b_name').value = '';
        document.getElementById('b_jql').value = 'assignee = currentUser() AND statusCategory != Done';
        document.getElementById('b_hours').value = 4;
        currentBoardMapping = [];
        renderFieldsTable();
    }

    function saveBoardConfig() {
        const mapping = [];
        document.querySelectorAll('.field-role-select').forEach(sel => {
            const fid = sel.getAttribute('data-fid');
            const fname = sel.getAttribute('data-fname');
            const role = sel.value;
            const chk = document.querySelector(`.field-body-check[data-fid="${fid}"]`);
            const inc = chk ? chk.checked : false;

            if(role !== 'none' || inc) {
                mapping.push({ field_id: fid, field_name: fname, role: role, include_in_body: inc });
            }
        });

        const bName = document.getElementById('b_name').value.trim();
        if(!bName) {
            alert('Debes ingresar un nombre para el tablero.');
            return;
        }

        const editId = document.getElementById('edit_board_db_id').value;
        const payload = {
            id: editId ? parseInt(editId) : null,
            client_id: activeClientId,
            board_id: parseInt(document.getElementById('b_select').value) || null,
            board_name: bName,
            custom_jql: document.getElementById('b_jql').value,
            hours_per_sp: parseInt(document.getElementById('b_hours').value) || 4,
            fields_mapping: mapping
        };

        fetch('/api/boards', {
            method: 'POST', headers: {'Content-Type':'application/json'}, body: JSON.stringify(payload)
        }).then(() => {
            alert('Configuración del tablero guardada correctamente.');
            cancelBoardEdit();
            loadBoards();
        });
    }

    function deleteBoard(id) {
        if(confirm('¿Eliminar tablero?')) {
            fetch('/api/boards/' + id, { method: 'DELETE' }).then(() => loadBoards());
        }
    }

    function loadLogs() {
        const filterVal = document.getElementById('logFilter').value;
        fetch(`/api/logs?filter=${encodeURIComponent(filterVal)}`)
            .then(r => r.json())
            .then(data => {
                const box = document.getElementById('logBox');
                box.textContent = data.logs.join('');
                box.scrollTop = box.scrollHeight;
            })
            .catch(e => console.error("Error al cargar logs:", e));
    }

    function filterLogs() {
        clearTimeout(logTimer);
        logTimer = setTimeout(loadLogs, 300);
    }

    window.onload = function() {
        loadAppSettings();
        loadClients();
    };
</script>
</body>
</html>
"""


@app.get("/", response_class=HTMLResponse)
def index():
    return HTML_TEMPLATE


def run_server(port: int = 8765):
    if sys.stdout is None:
        sys.stdout = open(os.devnull, "w")
    if sys.stderr is None:
        sys.stderr = open(os.devnull, "w")
    init_db()
    uvicorn.run(app, host="127.0.0.1", port=port, log_config=None, access_log=False)
