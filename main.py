import io
import base64
import json
import os
import time
import unicodedata
from urllib.parse import urlparse, unquote
from datetime import datetime
from PIL import Image
from office365.sharepoint.client_context import ClientContext
from office365.runtime.auth.user_credential import UserCredential

# ==========================================
# 1. CONFIGURACIÓN PARA GITHUB ACTIONS
# ==========================================
SITE_URL = "https://teams.wal-mart.com/sites/EquipoPlanificacin"
LIST_NAME = "Seguimiento Infraestructura"

# Las credenciales ahora se jalan desde los Secrets de GitHub de forma segura
USERNAME = os.environ.get("SP_USER")
PASSWORD = os.environ.get("SP_PASS") 

# Para GitHub Pages, el archivo debe llamarse index.html y guardarse en la raíz
OUTPUT_HTML = "index.html"

# ==========================================
# 2. UTILIDADES Y "SABUESO DE FOTOS"
# ==========================================
def limpiar(val):
    if val is None: return ""
    s = str(val).strip()
    if s == "0" or s == "0.0": return "0"
    if s.lower() == "nan": return "" 
    return s.replace(".0", "")

def normalizar_texto(texto):
    if not texto: return ""
    s = str(texto).lower().strip()
    return ''.join(c for c in unicodedata.normalize('NFD', s) if unicodedata.category(c) != 'Mn')

def formatear_fecha(texto_fecha):
    if not texto_fecha: return "--"
    try:
        s_fecha = str(texto_fecha)
        if "T" in s_fecha: return datetime.strptime(s_fecha.split("T")[0], "%Y-%m-%d").strftime("%d-%m-%Y")
        if isinstance(texto_fecha, datetime): return texto_fecha.strftime("%d-%m-%Y")
        if " " in s_fecha: return s_fecha.split(" ")[0]
        return s_fecha
    except: return str(texto_fecha)

def descargar_foto_por_url(ctx, url):
    try:
        url = unquote(url)
        if url.startswith("http"): url = urlparse(url).path
        
        file_content = io.BytesIO()
        ctx.web.get_file_by_server_relative_url(url).download(file_content).execute_query()
        file_content.seek(0)
        
        if len(file_content.getvalue()) > 0:
            with Image.open(file_content) as img:
                if img.mode != "RGB": img = img.convert("RGB")
                # Escalado a 600x600
                img.thumbnail((600, 600))
                buf = io.BytesIO()
                # Compresión a 60 
                img.save(buf, format='JPEG', quality=60)
                return f"data:image/jpeg;base64,{base64.b64encode(buf.getvalue()).decode('utf-8')}"
    except Exception:
        pass
    return None

def extraer_foto_columna(ctx, p, col_name, item_id):
    """Extrae la imagen específicamente de la columna indicada (Antes o Despues)"""
    img_b64 = None
    json_raw = p.get(col_name)
    if json_raw:
        try:
            data = json.loads(json_raw) if isinstance(json_raw, str) else json_raw
            if isinstance(data, dict):
                url = data.get("serverRelativeUrl") or data.get("serverUrl") or data.get("Url")
                filename = data.get("fileName")
                if url: 
                    img_b64 = descargar_foto_por_url(ctx, url)
                if not img_b64 and filename:
                    rel_site = SITE_URL.replace("https://teams.wal-mart.com", "")
                    url_adj = f"{rel_site}/Lists/{LIST_NAME}/Attachments/{item_id}/{filename}"
                    img_b64 = descargar_foto_por_url(ctx, url_adj)
        except: pass
    return img_b64

# ==========================================
# 3. EXTRACCIÓN PRINCIPAL
# ==========================================
def main():
    try:
        print("🚀 INICIANDO EXTRACCIÓN MAESTRA...")
        ctx = ClientContext(SITE_URL).with_credentials(UserCredential(USERNAME, PASSWORD))
        sp_list = ctx.web.lists.get_by_title(LIST_NAME)
        
        print("   ⏳ Solicitando registros y adjuntos a SharePoint...")
        
        columnas_req = [
            "Id", "Title", "LinkTitle", "field_2", "field_3", "field_4", 
            "field_5", "field_6", "field_7", "Responsable", "field_10", 
            "field_11", "field_14", "field_15", "Antes", "Despues", 
            "field_1", "ClaseM", "Zona", "Attachments", "AttachmentFiles"
        ]
        
        try:
            items = sp_list.items.select(columnas_req).expand(["AttachmentFiles"]).top(5000).get().execute_query()
        except Exception:
            columnas_req.remove("AttachmentFiles")
            items = sp_list.items.select(columnas_req).top(5000).get().execute_query()
            
        total_main = len(items)
        print(f"   ✅ Se descargaron {total_main} registros brutos.")
        
        db_json = {}
        for idx, item in enumerate(items):
            print(f"      ... Procesando OT {idx+1} de {total_main}", end='\r')
            p = item.properties
            item_id = int(p.get("Id", 0))
            
            semana_val = limpiar(p.get("field_1"))
            if semana_val not in ["14"]:
                continue 

            act_str = limpiar(p.get("field_4")) 
            tag_id = limpiar(p.get("LinkTitle"))
            titulo_final = act_str if act_str else (tag_id or f"OT #{item_id}")

            status_txt = normalizar_texto(limpiar(p.get("field_11"))) 
            if any(k in status_txt for k in ['ok', 'listo', 'cerrad', 'realiza', 'complet']): status = "realizada"
            elif any(k in status_txt for k in ['prog', 'planif']): status = "programado"
            elif any(k in status_txt for k in ['proceso', 'tratando', 'curso']): status = "en proceso"
            else: status = "pendiente"

            prio_raw = normalizar_texto(limpiar(p.get("field_10"))) 
            if "calavera" in prio_raw or "0" in prio_raw: prio = "0"
            elif "alta" in prio_raw or "1" in prio_raw: prio = "1"
            elif "media" in prio_raw or "2" in prio_raw: prio = "2"
            else: prio = "3"

            img_antes = extraer_foto_columna(ctx, p, "Antes", item_id)
            img_despues = extraer_foto_columna(ctx, p, "Despues", item_id)

            key_id = f"MTTO_{item_id}"
            db_json[key_id] = {
                "key_id": key_id,
                "id_real": item_id,
                "titulo": titulo_final,
                "tag": tag_id,
                "semana": semana_val or "S/N",
                "ejecutor": limpiar(p.get("Responsable")) or "Sin Asignar",
                "prioridad": prio,
                "ubicacion": limpiar(p.get("field_5")),
                "sub_ubi": limpiar(p.get("field_6")),
                "ot": limpiar(p.get("field_7")),
                "zona": limpiar(p.get("Zona")),
                "f_lev": formatear_fecha(p.get("field_2")),
                "f_cie": formatear_fecha(p.get("field_3")),
                "actividad": act_str or "Sin descripción",
                "observacion": limpiar(p.get("field_14")),
                "obs2": limpiar(p.get("field_15")),
                "status": status,
                "clase": limpiar(p.get("ClaseM")).title() or "General",
                "origen": "act",
                "img_antes": img_antes,
                "img_despues": img_despues
            }
            
        print("\n   ✅ Procesamiento finalizado. Construyendo HTML...")
        generar_html_moderno(db_json)

    except Exception as e: print(f"\n❌ Error Fatal: {e}")

# ==========================================
# 4. GENERADOR HTML
# ==========================================
def generar_html_moderno(db_json):
    fecha_actual = datetime.now().strftime("%d/%m/%Y %H:%M")
    
    html_template = """<!DOCTYPE html><html lang="es"><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width, initial-scale=1.0"><title>Dashboard Mantenimiento</title>
    <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
    <script src="https://cdn.jsdelivr.net/npm/chartjs-plugin-datalabels@2.0.0"></script>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/html2canvas/1.4.1/html2canvas.min.js"></script>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
    <style>
        :root { --primary: #0f172a; --secondary: #334155; --accent: #2563eb; --bg: #f8fafc; --border: #e2e8f0; --text: #1e293b; --muted: #64748b; --success: #10b981; --warn: #f59e0b; --danger: #ef4444; --info: #3b82f6; }
        * { box-sizing: border-box; outline: none; font-family: 'Segoe UI', system-ui, sans-serif; }
        body { background: var(--bg); color: var(--text); margin: 0; height: 100vh; display: flex; flex-direction: column; overflow: hidden; }
        
        .top-bar { background: var(--primary); color: white; padding: 0 20px; height: 60px; display: flex; justify-content: space-between; align-items: center; flex-shrink: 0; z-index: 10; box-shadow: 0 2px 5px rgba(0,0,0,0.1); }
        .brand h2 { margin: 0; font-size: 1.2rem; display:flex; align-items:center; gap: 8px; } 
        .brand span { opacity: 0.7; font-weight: 300; font-size: 0.95rem; }
        
        .tabs-container { background: white; border-bottom: 1px solid var(--border); padding: 0 20px; flex-shrink: 0; display:flex; justify-content: space-between; align-items: center; z-index: 5; box-shadow: 0 1px 3px rgba(0,0,0,0.05); }
        .tabs-nav { display: flex; gap: 15px; }
        .tab-btn { background: none; border: none; padding: 15px 5px; font-weight: 600; color: var(--muted); cursor: pointer; border-bottom: 3px solid transparent; transition: 0.2s; font-size: 0.95rem; }
        .tab-btn:hover { color: var(--accent); } .tab-btn.active { color: var(--accent); border-bottom-color: var(--accent); }
        
        .app-layout { display: flex; height: calc(100vh - 110px); width: 100%; overflow: hidden; }
        
        .col-filters { width: 280px; background: #fff; border-right: 1px solid var(--border); display: flex; flex-direction: column; flex-shrink: 0; z-index: 5; }
        
        /* Modificado para alojar el botón de borrar de manera elegante */
        .filters-header { padding: 15px 20px; border-bottom: 1px solid var(--border); font-weight: 700; color: var(--primary); font-size: 0.9rem; text-transform: uppercase; background: #f8fafc; display: flex; justify-content: space-between; align-items: center; }
        
        .filters-body { flex: 1; overflow-y: auto; padding: 20px; min-height: 0; } 
        .filters-footer { padding: 20px; border-top: 1px solid var(--border); background: #f8fafc; flex-shrink: 0; }
        
        .f-group { margin-bottom: 15px; }
        .f-group label { font-size: 0.75rem; font-weight: 700; color: var(--muted); display: block; margin-bottom: 6px; text-transform: uppercase; }
        select, input { width: 100%; padding: 10px; border: 1px solid var(--border); border-radius: 6px; font-size: 0.85rem; color: var(--text); }
        select:focus, input:focus { border-color: var(--accent); box-shadow: 0 0 0 2px rgba(37, 99, 235, 0.1); }
        .range-box { display: flex; align-items: center; gap: 5px; }
        
        .btn-clean { background: white; border: 1px solid var(--danger); color: var(--danger); padding: 10px; border-radius: 6px; cursor: pointer; font-weight: 700; transition: 0.2s; width: 100%; text-transform: uppercase; font-size: 0.8rem; letter-spacing: 0.5px; }
        .btn-clean:hover { background: var(--danger); color: white; }
        
        .kpi-row-mini { display: flex; justify-content: space-between; margin-bottom: 15px; }
        .kpi-box { text-align: center; } .k-label { display: block; font-size: 0.7rem; color: var(--muted); font-weight: 700; }
        .k-num { display: block; font-size: 1.3rem; font-weight: 800; color: var(--primary); } .k-ok { color: var(--success); } .k-pend { color: var(--danger); }
        .prog-title { display: flex; justify-content: space-between; font-size: 0.75rem; font-weight: 700; color: var(--muted); margin-bottom: 6px; }
        .progress-bar-container { width: 100%; height: 10px; background: #e2e8f0; border-radius: 5px; overflow: hidden; }
        .progress-bar-fill { height: 100%; background: var(--success); width: 0%; transition: width 1s cubic-bezier(0.4, 0, 0.2, 1); }
        
        /* LISTA OT */
        .col-list { width: 380px; background: #fff; border-right: 1px solid var(--border); display: flex; flex-direction: column; flex-shrink: 0; }
        .list-header { padding: 20px; border-bottom: 1px solid var(--border); font-weight: 600; background: #f8fafc; color: var(--secondary); font-size: 0.9rem; flex-shrink: 0; display:flex; flex-direction:column; gap:12px; }
        .list-scroll-area { flex: 1; overflow-y: auto; min-height: 0; }
        
        .list-item { padding: 15px 20px; border-bottom: 1px solid var(--border); cursor: pointer; transition: 0.2s; border-left: 4px solid transparent; }
        .list-item:hover { background: #f8fafc; } .list-item.selected { background: #eff6ff; border-left-color: var(--accent); }
        .li-top { display: flex; justify-content: space-between; margin-bottom: 6px; font-size: 0.75rem; color: var(--muted); font-weight: 600; }
        .li-title { font-weight: 700; font-size: 0.95rem; color: var(--primary); margin-bottom: 10px; line-height: 1.4; }
        .li-btm { display: flex; justify-content: space-between; font-size: 0.75rem; align-items: center; }
        
        .tag { padding: 4px 8px; border-radius: 4px; font-weight: 700; font-size: 0.7rem; letter-spacing: 0.3px; }
        .st-ok { background: #dcfce7; color: #166534; } .st-pend { background: #fee2e2; color: #991b1b; } .st-prog { background: #e0f2fe; color: #075985; } .st-proc { background: #fef3c7; color: #92400e; }
        
        /* DETALLE */
        .col-detail { flex: 1; background: #f1f5f9; overflow-y: auto; padding: 40px; }
        .empty-state { display: flex; flex-direction: column; align-items: center; justify-content: center; height: 100%; color: var(--muted); opacity: 0.7; }
        .detail-content { background: white; border-radius: 12px; box-shadow: 0 10px 15px -3px rgba(0,0,0,0.1); overflow: hidden; max-width: 1000px; margin: 0 auto; border: 1px solid var(--border); }
        .detail-header { padding: 30px; border-bottom: 1px solid var(--border); background: #fff; }
        .dh-top { display: flex; justify-content: space-between; margin-bottom: 15px; align-items:center; }
        .detail-header h2 { margin: 0 0 5px 0; font-size: 1.6rem; color: var(--primary); }
        
        .data-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(160px, 1fr)); gap: 25px; padding: 30px; background: #fff; border-bottom: 1px solid var(--border); }
        .dg-item small { display: block; font-size: 0.7rem; color: var(--muted); font-weight: 700; margin-bottom: 6px; text-transform: uppercase; }
        .dg-item strong { font-size: 1rem; color: var(--text); }
        
        .obs-box { padding: 30px; border-bottom: 1px solid var(--border); }
        .obs-box h4 { margin: 0 0 12px; color: var(--secondary); font-size: 0.9rem; text-transform: uppercase; }
        .obs-box p { background: #f8fafc; padding: 20px; border-radius: 8px; border: 1px solid var(--border); margin: 0; line-height: 1.6; color: #334155; }
        
        /* GALERÍA ANTES/DESPUÉS */
        .gallery-section { padding: 30px; background: #f8fafc; display:flex; flex-direction: column; align-items: center; gap: 15px; }
        .gallery-section h4 { margin:0; color:var(--secondary); font-size:0.9rem; text-transform:uppercase; align-self: flex-start; }
        .gallery-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(300px, 1fr)); gap: 20px; width: 100%; }
        .gal-box { background: white; border: 1px solid var(--border); border-radius: 8px; padding: 15px; display: flex; flex-direction: column; align-items: center; gap: 10px; box-shadow: 0 2px 4px rgba(0,0,0,0.05); }
        .gal-box span { font-weight: 700; font-size: 0.85rem; color: var(--secondary); text-transform: uppercase; padding-bottom: 5px; border-bottom: 2px solid var(--accent); margin-bottom: 5px; }
        .gal-img { max-width: 100%; max-height: 350px; border-radius: 6px; cursor: zoom-in; box-shadow: 0 2px 5px rgba(0,0,0,0.1); transition: transform 0.2s; object-fit: contain; }
        .gal-img:hover { transform: scale(1.02); }
        
        /* CSS DE GRÁFICOS */
        .graficos-layout { flex: 1; padding: 30px; display: grid; grid-template-columns: repeat(auto-fit, minmax(400px, 1fr)); grid-auto-rows: min-content; gap: 25px; overflow-y: auto; background: #f1f5f9; align-content: start; }
        .chart-card { background: white; padding: 25px; border-radius: 12px; border: 1px solid var(--border); box-shadow: 0 4px 6px -1px rgba(0,0,0,0.05); display: flex; flex-direction: column; height: 400px; width: 100%; }
        .chart-card.wide { grid-column: 1 / -1; height: 450px; }
        .chart-title { font-size: 1rem; font-weight: 700; color: var(--secondary); margin-bottom: 15px; text-transform: uppercase; text-align: center; }
        .canvas-container { position: relative; flex: 1 1 auto; width: 100%; min-height: 0; }
        
        .prio-flag { padding: 4px 10px; border-radius: 6px; font-weight: 700; font-size: 0.75rem; }
        .p-crit { background: #fee2e2; color: #dc2626; border: 1px solid #f87171; }
        .p-alta { background: #ffedd5; color: #ea580c; border: 1px solid #fdba74; }
        .p-med { background: #fef3c7; color: #d97706; border: 1px solid #fcd34d; }
        .p-baja { background: #f1f5f9; color: #64748b; border: 1px solid #cbd5e1; }
        
        /* Modals */
        .modal { display: none; position: fixed; z-index: 2000; left: 0; top: 0; width: 100%; height: 100%; background: rgba(15, 23, 42, 0.85); align-items: center; justify-content: center; backdrop-filter: blur(4px); }
        .modal img { max-width: 90%; max-height: 90vh; border-radius: 8px; box-shadow: 0 25px 50px -12px rgba(0,0,0,0.5); }
        
        #data_modal_content { background: white; width: 90%; max-width: 1200px; max-height: 85vh; border-radius: 12px; display: flex; flex-direction: column; overflow: hidden; box-shadow: 0 25px 50px -12px rgba(0,0,0,0.5); }
        .dm-header { padding: 20px 25px; background: var(--primary); color: white; display: flex; justify-content: space-between; align-items: center; }
        .dm-header h3 { margin: 0; font-size: 1.2rem; font-weight: 600; }
        .dm-close { background: none; border: none; color: white; font-size: 1.8rem; cursor: pointer; opacity: 0.8; transition: 0.2s; line-height: 1; }
        .dm-close:hover { opacity: 1; transform: scale(1.1); }
        .dm-body { padding: 0; overflow-y: auto; flex: 1; background: var(--bg); }
        .dm-table { width: 100%; border-collapse: collapse; background: white; font-size: 0.9rem; text-align: left; }
        .dm-table th { background: #f8fafc; padding: 15px 20px; font-weight: 700; color: var(--secondary); border-bottom: 2px solid var(--border); position: sticky; top: 0; z-index: 10; text-transform: uppercase; font-size: 0.8rem; }
        .dm-table td { padding: 15px 20px; border-bottom: 1px solid var(--border); color: var(--text); }
        .dm-table tr { transition: background 0.2s; }
        .dm-table tr:hover td { background: #eff6ff; cursor: pointer; }
        
        .summary-block { background:#f8fafc; border:1px solid #e2e8f0; border-radius:8px; padding:15px; margin-bottom:12px; }
        .summary-header { display:flex; justify-content:space-between; align-items:center; margin-bottom:5px; }
        .summary-title { font-weight:700; font-size: 0.95rem; }
        .summary-perc { font-weight:800; font-size: 1.1rem; }
        .summary-sub { font-size:0.8rem; color:#64748b; }
        .summary-bar-bg { width:100%; height:6px; background:#e2e8f0; border-radius:3px; margin-top:8px; overflow:hidden; }
        .summary-bar-fill { height:100%; transition:width 1s cubic-bezier(0.4, 0, 0.2, 1); }
    </style>
</head>
<body>
    <div id="modal" class="modal" onclick="if(event.target===this) this.style.display='none'"><img id="modalImg"></div>
    
    <div id="data_modal" class="modal" onclick="if(event.target===this) this.style.display='none'">
        <div id="data_modal_content"></div>
    </div>

    <div class="top-bar">
        <div class="brand"><h2>⚙️ Panel Gestión de Actividades <span>SubGerencia de Mantenimiento</span></h2></div>
        <div style="font-size:0.85rem; font-weight:600; opacity:0.9;">Actualizado: __FECHA_ACTUAL__ | Semanas 14</div>
    </div>
    
    <div class="tabs-container">
        <div class="tabs-nav">
            <button class="tab-btn active" onclick="setView('list', this)" id="btn_tab_list">📋 Visor de OTs</button>
            <button class="tab-btn" onclick="setView('charts', this)">📊 Análisis y Tendencias</button>
            <button class="tab-btn" onclick="setView('row', this)">📈 ROW</button>
        </div>
        <div style="display:flex; gap:10px;">
            <button onclick="descargarExcel()" class="btn-clean" style="margin: 0; padding: 8px 15px; width: auto; border-color: #10b981; color: #10b981; display: flex; align-items: center; gap: 8px;" title="Descargar datos filtrados">
                <span style="font-size:1.2rem;">📊</span> Exportar Excel
            </button>
        </div>
    </div>
    
    <div class="app-layout">
        <div class="col-filters" id="main_filters">
            <div class="filters-header">
                <span>🔍 Filtros Principales</span>
                <button onclick="resetFilters()" class="btn-clean" style="margin: 0; padding: 4px 8px; width: auto; font-size: 0.7rem; border-color: #ef4444; color: #ef4444; display: flex; align-items: center; gap: 4px; text-transform: none; letter-spacing: normal;" title="Limpiar todos los filtros">
                    🧹 Borrar
                </button>
            </div>
            
            <div class="filters-body" id="filters_dynamic"></div>
            <div class="filters-footer">
                <div class="kpi-row-mini">
                    <div class="kpi-box"><span class="k-label">TOTAL OT</span><span class="k-num" id="k_total">0</span></div>
                    <div class="kpi-box"><span class="k-label">CERRADAS</span><span class="k-num k-ok" id="k_ok">0</span></div>
                    <div class="kpi-box"><span class="k-label">BACKLOG</span><span class="k-num k-pend" id="k_pend">0</span></div>
                </div>
                <div class="prog-title"><span>Cumplimiento Global</span><span id="k_perc">0%</span></div>
                <div class="progress-bar-container"><div id="bar_fill" class="progress-bar-fill"></div></div>
            </div>
        </div>

        <div id="view_list" style="display:flex; flex:1; overflow:hidden;">
            <div class="col-list">
                <div class="list-header">
                    <div>📋 Listado de Actividades</div>
                    <input type="text" id="search_input" placeholder="🔍 Buscar TAG, Título o OT..." onkeyup="applyFilters()">
                </div>
                <div id="list_container" class="list-scroll-area"></div>
            </div>
            <div class="col-detail">
                <div id="empty_state" class="empty-state"><div style="font-size:4rem; margin-bottom:15px;">📋</div><h3 style="margin:0">Selecciona una OT</h3><p>Usa la lista izquierda para ver detalles.</p></div>
                <div id="detail_view" class="detail-content" style="display:none">
                    <div class="detail-header">
                        <div class="dh-top">
                            <div><span id="d_status" class="tag st-ok">STATUS</span></div>
                            <div id="d_prio_lbl">PRIO</div>
                        </div>
                        <h2 id="d_title">Título de la Actividad</h2>
                        <p style="color:var(--accent); font-weight: 600; font-size: 1.05rem; margin:0;" id="d_tag">TAG</p>
                    </div>
                    <div class="data-grid" id="d_grid"></div>
                    <div class="obs-box" id="box_obs1"><h4 id="lbl_obs_title">📝 Observación Técnica</h4><p id="d_obs">--</p></div>
                    <div class="obs-box" id="box_obs2" style="display:none;"><h4 id="lbl_obs_title2">📝 Observación Adicional</h4><p id="d_obs2">--</p></div>
                    
                    <div class="gallery-section" id="d_gallery_sec">
                        <h4>📸 Registro Fotográfico</h4>
                        <div id="d_img_container" style="width: 100%;"></div>
                    </div>
                </div>
            </div>
        </div>

        <div id="view_charts" class="graficos-layout" style="display:none;">
            <div class="chart-card"><div class="chart-title">Status del Backlog</div><div class="canvas-container"><canvas id="chart1"></canvas></div></div>
            <div class="chart-card"><div class="chart-title">Clase de Mantenimiento</div><div class="canvas-container"><canvas id="chart2"></canvas></div></div>
            
            <div class="chart-card">
                <div class="chart-title">Resumen de Actividades</div>
                <div id="summary_content" style="display:flex; flex-direction:column; justify-content:space-around; flex:1;">
                </div>
            </div>
            
            <div class="chart-card wide"><div class="chart-title">Carga Laboral por Responsable</div><div class="canvas-container"><canvas id="chart3"></canvas></div></div>
            <div class="chart-card wide"><div class="chart-title">Top Ubicaciones Críticas</div><div class="canvas-container"><canvas id="chart4"></canvas></div></div>
        </div>
        
        <div id="view_row" style="display:none; flex:1; flex-direction:column; overflow-y:auto; padding:30px; background:#f1f5f9;">
            <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:20px; flex-wrap:wrap; gap:15px;">
                <h2 style="color:var(--primary); margin:0; font-size:1.8rem;">Planificación Mantenimiento <span id="row_week_title" style="color:var(--accent);">--</span></h2>
                <button id="btn_descargar_row" class="btn-clean" style="width:auto; margin:0; padding: 8px 15px; border-color:var(--accent); color:var(--accent); display:flex; align-items:center; gap:8px;" onclick="descargarROW()">
                    <span style="font-size:1.2rem;">📸</span> Descargar Dashboard ROW
                </button>
            </div>
            
            <div style="display:flex; gap:25px; margin-bottom:30px; flex-wrap:wrap;">
                <div class="chart-card" style="flex:1; height:350px; min-width:300px;"><div class="chart-title">Distribución Mantenimiento vs Aseo</div><div class="canvas-container"><canvas id="row_chart1"></canvas></div></div>
                <div class="chart-card" style="flex:1; height:350px; min-width:300px;"><div class="chart-title">Cumplimiento Mantenimiento General</div><div class="canvas-container"><canvas id="row_chart2"></canvas></div></div>
                <div class="chart-card" style="flex:1; height:350px; min-width:300px;"><div class="chart-title">Cumplimiento Aseo General</div><div class="canvas-container"><canvas id="row_chart3"></canvas></div></div>
            </div>
            
            <div style="display:flex; gap:25px; flex-wrap:wrap;">
                <div class="chart-card" style="flex:1; height:450px; min-width:400px;"><div class="chart-title">Panadería: Cumplimiento por Línea</div><div class="canvas-container"><canvas id="row_chart4"></canvas></div></div>
                <div class="chart-card" style="flex:1; height:450px; min-width:400px;"><div class="chart-title">Dely: Cumplimiento por Área</div><div class="canvas-container"><canvas id="row_chart5"></canvas></div></div>
            </div>
        </div>

    </div>

    <script>
    Chart.register(ChartDataLabels);
    Chart.defaults.plugins.datalabels.display = false; 

    const db = __DB_JSON_DATA__;
    const records = Object.values(db).sort((a,b) => b.id_real - a.id_real);
    const weeks = [...new Set(records.map(x=>x.semana).filter(x=>x!=="S/N"))].sort((a,b)=>{ let na=parseInt(a), nb=parseInt(b); return (isNaN(na)||isNaN(nb)) ? a.localeCompare(b) : na-nb; });
    
    let appState = { statusFilter: 'all', view: 'list' };
    let currentChartData = [];
    let chartInstances = {};
    
    Chart.defaults.font.family = "'Segoe UI', system-ui, sans-serif";
    Chart.defaults.color = '#64748b';

    function buildFilters() {
        const fDiv = document.getElementById('filters_dynamic');
        
        const createSelect = (id, label, options) => {
            let sel = `<div class="f-group"><label>${label}</label><select id="${id}" onchange="applyFilters()">`;
            sel += `<option value="ALL">Todos</option>`;
            options.forEach(o => { if(o) sel += `<option value="${o}">${o}</option>`; });
            sel += `</select></div>`;
            return sel;
        };

        let html = '';
        html += createSelect('f_semana', '📆 Semana', weeks);
        html += createSelect('f_zona', '📍 Zona', [...new Set(records.map(x=>x.zona))].filter(Boolean).sort());
        html += createSelect('f_clase', '🛠️ Clase MTTO', [...new Set(records.map(x=>x.clase))].sort());
        html += createSelect('f_exec', '👷 Responsable', [...new Set(records.map(x=>x.ejecutor))].sort());
        html += createSelect('f_ubi', '🏭 Línea / Área', [...new Set(records.map(x=>x.ubicacion))].sort());
        html += `<div class="f-group"><label>🚦 Estado</label><select id="f_status" onchange="applyFilters()"><option value="ALL">Todas las OTs</option><option value="pendientes">Pendientes / En Proceso</option><option value="realizada">Solo Cerradas (OK)</option></select></div>`;
        
        fDiv.innerHTML = html;
    }

    function resetFilters() {
        if(document.getElementById('search_input')) document.getElementById('search_input').value = '';
        document.querySelectorAll('.f-group select').forEach(sel => sel.value = "ALL");
        applyFilters();
    }

    function setView(view, btn) {
        appState.view = view;
        document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
        if(btn) btn.classList.add('active');
        else document.getElementById('btn_tab_list').classList.add('active');
        
        document.getElementById('view_list').style.display = 'none';
        document.getElementById('view_charts').style.display = 'none';
        document.getElementById('view_row').style.display = 'none';

        if (view === 'list') {
            document.getElementById('view_list').style.display = 'flex';
        } else if (view === 'charts') {
            document.getElementById('view_charts').style.display = 'grid';
            setTimeout(() => { drawCharts(currentChartData); }, 50);
        } else if (view === 'row') {
            document.getElementById('view_row').style.display = 'flex';
            setTimeout(() => { drawRowCharts(currentChartData); }, 50);
        }
        applyFilters();
    }

    function getFilteredData() {
        const eVal = document.getElementById('f_exec') ? document.getElementById('f_exec').value : 'ALL';
        const uVal = document.getElementById('f_ubi') ? document.getElementById('f_ubi').value : 'ALL';
        const cVal = document.getElementById('f_clase') ? document.getElementById('f_clase').value : 'ALL';
        const stVal = document.getElementById('f_status') ? document.getElementById('f_status').value : 'ALL';
        const semVal = document.getElementById('f_semana') ? document.getElementById('f_semana').value : 'ALL';
        const zVal = document.getElementById('f_zona') ? document.getElementById('f_zona').value : 'ALL';
        const searchVal = document.getElementById('search_input') ? document.getElementById('search_input').value.toLowerCase().trim() : '';

        return records.filter(d => {
            if (stVal !== 'ALL') {
                const isOk = (d.status === 'realizada');
                if (stVal === 'realizada' && !isOk) return false;
                if (stVal === 'pendientes' && isOk) return false;
            }
            
            if (searchVal !== '') {
                const text = `${d.titulo} ${d.ot} ${d.tag}`.toLowerCase();
                if (!text.includes(searchVal)) return false;
            }

            if (cVal !== 'ALL' && d.clase !== cVal) return false;
            if (eVal !== 'ALL' && d.ejecutor !== eVal) return false;
            if (uVal !== 'ALL' && d.ubicacion !== uVal) return false;
            if (semVal !== 'ALL' && d.semana !== semVal) return false;
            if (zVal !== 'ALL' && d.zona !== zVal) return false;
            
            return true;
        });
    }

    function applyFilters() {
        currentChartData = getFilteredData();
        
        let ok = 0;
        currentChartData.forEach(d => { if(d.status === 'realizada') ok++; });
        const total = currentChartData.length;
        
        document.getElementById('k_total').innerText = total;
        document.getElementById('k_ok').innerText = ok;
        document.getElementById('k_pend').innerText = total - ok;
        let perc = total > 0 ? Math.round((ok/total)*100) : 0;
        document.getElementById('k_perc').innerText = perc + '%';
        const bar = document.getElementById('bar_fill');
        bar.style.width = perc + '%';
        bar.style.backgroundColor = perc > 80 ? '#10b981' : (perc > 40 ? '#f59e0b' : '#ef4444');

        if(appState.view === 'list') renderList(currentChartData);
        else if (appState.view === 'charts') drawCharts(currentChartData);
        else if (appState.view === 'row') drawRowCharts(currentChartData);
    }

    function renderList(data) {
        const container = document.getElementById('list_container');
        container.innerHTML = '';
        
        data.forEach(d => {
            const item = document.createElement('div');
            item.className = 'list-item';
            item.onclick = function() { 
                renderDetail(d.key_id); 
                document.querySelectorAll('.list-item').forEach(i=>i.classList.remove('selected'));
                item.classList.add('selected');
            };
            
            let stText = '⚠️ PEND'; let stClass = 'st-pend';
            if (d.status === 'realizada') { stText='✅ CERRADA'; stClass='st-ok'; }
            else if (d.status === 'programado') { stText='📅 PROG'; stClass='st-prog'; }
            else if (d.status === 'en proceso') { stText='🔨 PROCESO'; stClass='st-proc'; }
            
            let idDisplay = d.ot ? `OT: ${d.ot}` : (d.tag ? d.tag : '#' + d.id_real);
            
            item.innerHTML = `
                <div class="li-top"><span>${idDisplay}</span><span>Sem: ${d.semana}</span></div>
                <div class="li-title">${d.titulo}</div>
                <div class="li-btm">
                    <span class="tag ${stClass}">${stText}</span>
                    <span style="color:var(--muted); font-weight:700;">👷 ${d.ejecutor.split(' ')[0]}</span>
                </div>
            `;
            container.appendChild(item);
        });
    }

    function renderDetail(key) {
        document.getElementById('empty_state').style.display='none';
        document.getElementById('detail_view').style.display='block';
        const d = db[key];
        
        document.getElementById('d_title').innerText = d.titulo;
        document.getElementById('d_tag').innerText = d.tag ? `TAG / Equipo: ${d.tag}` : (d.ot ? `OT: ${d.ot}` : 'Sin TAG');
        
        const stBadge = document.getElementById('d_status');
        if (d.status === 'realizada') { stBadge.innerText = '✅ CERRADA'; stBadge.className = 'tag st-ok'; }
        else if (d.status === 'programado') { stBadge.innerText = '📅 PROGRAMADA'; stBadge.className = 'tag st-prog'; }
        else if (d.status === 'en proceso') { stBadge.innerText = '🔨 EN PROCESO'; stBadge.className = 'tag st-proc'; }
        else { stBadge.innerText = '⚠️ PENDIENTE'; stBadge.className = 'tag st-pend'; }
        
        let pl = d.prioridad;
        if(pl==='0') pl='<span class="prio-flag p-crit">🚨 ALTA / CRÍTICA</span>';
        else if(pl==='1') pl='<span class="prio-flag p-alta">🔴 ALTA</span>';
        else if(pl==='2') pl='<span class="prio-flag p-med">🟡 MEDIA</span>';
        else pl='<span class="prio-flag p-baja">🟢 BAJA</span>';
        document.getElementById('d_prio_lbl').innerHTML = pl;

        const grid = document.getElementById('d_grid');
        grid.innerHTML = '';
        const createItem = (label, val) => `<div class="dg-item"><small>${label}</small><strong>${val||'--'}</strong></div>`;
        
        grid.innerHTML += createItem('🛠️ Clase MTTO', d.clase);
        grid.innerHTML += createItem('📍 Zona', d.zona);
        grid.innerHTML += createItem('👷 Responsable', d.ejecutor);
        grid.innerHTML += createItem('🏭 Línea / Área', d.ubicacion);
        grid.innerHTML += createItem('📌 Sub Ubicación', d.sub_ubi);
        grid.innerHTML += createItem('🟢 Levantamiento', d.f_lev);
        grid.innerHTML += createItem('🏁 Cierre', d.f_cie);
        grid.innerHTML += createItem('📆 Semana', d.semana);
        grid.innerHTML += createItem('🧾 OT SAP', d.ot);
        
        document.getElementById('box_obs1').style.display = 'block';
        document.getElementById('d_obs').innerText = d.observacion || 'Sin observaciones registradas.';
        
        if(d.obs2) { document.getElementById('box_obs2').style.display = 'block'; document.getElementById('d_obs2').innerText = d.obs2; }
        else { document.getElementById('box_obs2').style.display = 'none'; }
        
        const imgContainer = document.getElementById('d_img_container');
        let htmlImgs = '<div class="gallery-grid">';
        let hasImgs = false;

        if (d.img_antes) {
            htmlImgs += `<div class="gal-box"><span>📸 Antes</span><img src="${d.img_antes}" class="gal-img" onclick="openModal(this.src)"></div>`;
            hasImgs = true;
        } else {
            htmlImgs += `<div class="gal-box"><span>📸 Antes</span><div style="height:150px; display:flex; align-items:center; justify-content:center; color:#cbd5e1; font-style:italic; font-weight:600; font-size:0.9rem;">Sin foto "Antes"</div></div>`;
        }

        if (d.img_despues) {
            htmlImgs += `<div class="gal-box"><span>📸 Después</span><img src="${d.img_despues}" class="gal-img" onclick="openModal(this.src)"></div>`;
            hasImgs = true;
        } else {
            htmlImgs += `<div class="gal-box"><span>📸 Después</span><div style="height:150px; display:flex; align-items:center; justify-content:center; color:#cbd5e1; font-style:italic; font-weight:600; font-size:0.9rem;">Sin foto "Después"</div></div>`;
        }
        htmlImgs += '</div>';

        imgContainer.innerHTML = htmlImgs;
        document.getElementById('d_gallery_sec').style.display = 'flex';
    }

    function openModal(src) {
        document.getElementById('modalImg').src = src;
        document.getElementById('modal').style.display = 'flex';
    }

    function showDataModal(title, filterFn, colProp = 'ubicacion') {
        let colHeader = colProp === 'clase' ? 'Clase de Actividad' : 'Ubicación';
        
        let html = `<div class="dm-header">
            <h3>📊 Desglose: ${title}</h3>
            <button class="dm-close" onclick="document.getElementById('data_modal').style.display='none'">&times;</button>
        </div>
        <div class="dm-body">
            <table class="dm-table">
                <thead><tr><th>OT / TAG</th><th>${colHeader}</th><th>Título / Actividad</th><th>Responsable</th><th>Estado</th><th>Observación</th></tr></thead>
                <tbody>`;

        let datosFiltrados = currentChartData.filter(filterFn);
        
        datosFiltrados.sort((a, b) => {
            let valA = a[colProp] ? String(a[colProp]).toLowerCase() : "";
            let valB = b[colProp] ? String(b[colProp]).toLowerCase() : "";
            return valA.localeCompare(valB);
        });

        let found = datosFiltrados.length > 0;
        
        datosFiltrados.forEach(d => {
            let stColor = d.status === 'realizada' ? '#166534' : (d.status === 'pendiente' ? '#991b1b' : '#92400e');
            let idDisplay = d.ot ? d.ot : (d.tag ? d.tag : '#' + d.id_real);
            let obsText = d.observacion ? (d.observacion.length > 45 ? d.observacion.substring(0, 42) + '...' : d.observacion) : '-';
            let colText = d[colProp] || '-';

            html += `<tr onclick="document.getElementById('data_modal').style.display='none'; document.getElementById('btn_tab_list').click(); setTimeout(() => renderDetail('${d.key_id}'), 100);">
                <td style="font-weight:700;">${idDisplay}</td>
                <td>${colText}</td>
                <td>${d.titulo}</td>
                <td>${d.ejecutor.split(' ')[0]}</td>
                <td style="color:${stColor}; font-weight:700; text-transform:uppercase;">${d.status}</td>
                <td style="max-width: 250px; white-space: nowrap; overflow: hidden; text-overflow: ellipsis;" title="${d.observacion}">${obsText}</td>
            </tr>`;
        });

        if (!found) html += `<tr><td colspan="6" style="text-align:center; padding: 30px; color:var(--muted);">No hay OTs para esta selección</td></tr>`;
        html += `</tbody></table></div>`;
        document.getElementById('data_modal_content').innerHTML = html;
        document.getElementById('data_modal').style.display = 'flex';
    }

    function getFreshCanvas(id) {
        const old = document.getElementById(id);
        if(!old) return null;
        const container = old.parentElement;
        container.innerHTML = `<canvas id="${id}"></canvas>`;
        return document.getElementById(id);
    }

    function descargarROW() {
        const btn = document.getElementById('btn_descargar_row');
        const originalText = btn.innerHTML;
        btn.innerHTML = "⏳ Generando Imagen...";
        
        const container = document.getElementById('view_row');
        
        html2canvas(container, { scale: 2, backgroundColor: "#f1f5f9" }).then(canvas => {
            let link = document.createElement('a');
            link.download = 'Dashboard_ROW.png';
            link.href = canvas.toDataURL('image/png');
            link.click();
            btn.innerHTML = originalText;
        }).catch(err => {
            alert("Error al capturar la pantalla.");
            btn.innerHTML = originalText;
        });
    }

    function descargarExcel() {
        if (!currentChartData || currentChartData.length === 0) {
            alert("No hay datos para exportar con los filtros actuales.");
            return;
        }

        const datosExcel = currentChartData.map(d => ({
            "Levantamiento": d.f_lev,
            "Cierre": d.f_cie,
            "Actividad": d.actividad,
            "Clase": d.clase,
            "Zona": d.zona,
            "Ubicación": d.ubicacion,
            "Sub Ubicación": d.sub_ubi,
            "OT": d.ot,
            "Ejecutor": d.ejecutor,
            "Status": d.status.toUpperCase(),
            "Observación": d.observacion,
            "Semana": d.semana
        }));

        const worksheet = XLSX.utils.json_to_sheet(datosExcel);
        const workbook = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(workbook, worksheet, "Base de Datos");

        const anchos = [
            { wch: 12 }, { wch: 12 }, { wch: 40 }, { wch: 15 }, { wch: 15 }, { wch: 20 }, 
            { wch: 20 }, { wch: 15 }, { wch: 20 }, { wch: 15 }, { wch: 50 }, { wch: 10 }
        ];
        worksheet['!cols'] = anchos;

        let fechaEx = new Date().toISOString().split('T')[0];
        XLSX.writeFile(workbook, `Reporte_MTTO_${fechaEx}.xlsx`);
    }

    const isAseoAct = (d) => {
        let textMatch = (d.ubicacion + " " + (d.sub_ubi || "") + " " + (d.titulo || "")).toLowerCase();
        let claseL = (d.clase || '').toLowerCase();
        return textMatch.includes('aseo') || claseL.includes('aseo');
    };

    const getPLoc = (d) => {
        let textMatch = (d.ubicacion + " " + (d.sub_ubi || "") + " " + (d.titulo || "")).toLowerCase();
        if (textMatch.includes('l1') || textMatch.includes('panadería 1') || textMatch.includes('panaderia 1')) return 'L1';
        if (textMatch.includes('l2') || textMatch.includes('panadería 2') || textMatch.includes('panaderia 2')) return 'L2';
        if (textMatch.includes('l3') || textMatch.includes('panadería 3') || textMatch.includes('panaderia 3')) return 'L3';
        if (textMatch.includes('l4') || textMatch.includes('panadería 4') || textMatch.includes('panaderia 4')) return 'L4';
        if (textMatch.includes('l5') || textMatch.includes('panadería 5') || textMatch.includes('panaderia 5')) return 'L5';
        return null;
    };

    const getDLoc = (d) => {
        let textMatch = (d.ubicacion + " " + (d.sub_ubi || "") + " " + (d.titulo || "")).toLowerCase();
        if (textMatch.includes('pizza')) return 'Pizza';
        if (textMatch.includes('bollerí') || textMatch.includes('bolleri')) return 'Bolleria';
        if (textMatch.includes('empanada')) return 'Empanadas';
        return null;
    };

    function drawCharts(data) {
        if(!data) return;

        let stats = { ok:0, pend:0, prog:0, ex:{}, loc:{}, wCounts:{}, cCounts:{} };
        let totAseo = 0, okAseo = 0;
        let totMtto = 0, okMtto = 0;
        let totGen = data.length, okGen = 0;

        weeks.forEach(w => stats.wCounts[w] = {total:0, ok:0});
        
        data.forEach(d => {
            let isOk = (d.status === 'realizada');
            if(isOk) { stats.ok++; okGen++; }
            else if(d.status === 'programado') stats.prog++;
            else stats.pend++;
            
            stats.cCounts[d.clase] = (stats.cCounts[d.clase]||0)+1;
            
            let isAseo = isAseoAct(d);
            if(isAseo) {
                totAseo++;
                if(isOk) okAseo++;
            } else {
                totMtto++;
                if(isOk) okMtto++;
            }

            const e = d.ejecutor || 'Sin Asignar';
            if(!stats.ex[e]) stats.ex[e]={ok:0, pend:0};
            if(isOk) stats.ex[e].ok++; else stats.ex[e].pend++;

            const l = d.ubicacion || 'Sin Ubicación';
            if(!stats.loc[l]) stats.loc[l]=0;
            stats.loc[l]++;
            
            if(d.semana!=="S/N" && stats.wCounts[d.semana]) {
                stats.wCounts[d.semana].total++;
                if(isOk) stats.wCounts[d.semana].ok++;
            }
        });

        let percAseo = totAseo > 0 ? Math.round((okAseo/totAseo)*100) : 0;
        let percMtto = totMtto > 0 ? Math.round((okMtto/totMtto)*100) : 0;
        let percGen = totGen > 0 ? Math.round((okGen/totGen)*100) : 0;

        let colAseo = percAseo >= 80 ? '#10b981' : (percAseo >= 40 ? '#f59e0b' : '#ef4444');
        let colMtto = percMtto >= 80 ? '#10b981' : (percMtto >= 40 ? '#f59e0b' : '#ef4444');
        let colGen = percGen >= 80 ? '#1d4ed8' : (percGen >= 40 ? '#f59e0b' : '#ef4444');

        let summaryHtml = `
            <div class="summary-block">
                <div class="summary-header">
                    <span class="summary-title" style="color:#3b82f6;">🧹 Apoyo Aseo</span>
                    <span class="summary-perc" style="color:${colAseo};">${percAseo}%</span>
                </div>
                <div class="summary-sub">De un total de <b>${totAseo}</b>, <b>${okAseo}</b> realizadas</div>
                <div class="summary-bar-bg">
                    <div class="summary-bar-fill" style="width:${percAseo}%; background:${colAseo};"></div>
                </div>
            </div>

            <div class="summary-block">
                <div class="summary-header">
                    <span class="summary-title" style="color:#8b5cf6;">🔧 Mantenimiento</span>
                    <span class="summary-perc" style="color:${colMtto};">${percMtto}%</span>
                </div>
                <div class="summary-sub">De un total de <b>${totMtto}</b>, <b>${okMtto}</b> realizadas</div>
                <div class="summary-bar-bg">
                    <div class="summary-bar-fill" style="width:${percMtto}%; background:${colMtto};"></div>
                </div>
            </div>

            <div class="summary-block" style="background:#eff6ff; border-color:#bfdbfe; text-align:center; padding: 20px 15px; margin-top: auto; margin-bottom: 0;">
                <div style="font-size:0.8rem; color:#1e40af; font-weight:700; text-transform:uppercase; margin-bottom:5px;">Cumplimiento Plan FDS Total</div>
                <div style="font-size:2rem; font-weight:800; color:${colGen};">${percGen}%</div>
                <div style="font-size:0.85rem; color:#3b82f6; margin-top:5px;">De un total de <b>${totGen}</b> actividades</div>
            </div>
        `;
        document.getElementById('summary_content').innerHTML = summaryHtml;

        const chartOpts = { 
            maintainAspectRatio:false, 
            responsive:true, 
            animation: { duration: 1200, easing: 'easeOutQuart' },
            layout: { padding: 10 },
            plugins: { datalabels: { display: false } } 
        };
        const gridHideX = { x: { grid: { display: false }, ticks: { maxRotation: 0, autoSkip: false } }, y: { grid: { color: '#f1f5f9' } } };
        const gridHideY = { x: { grid: { color: '#f1f5f9' } }, y: { grid: { display: false } } };

        new Chart(getFreshCanvas('chart1'), { 
            type: 'doughnut', 
            data: { labels:['Cerradas','Pendientes','Programadas'], datasets:[{ data:[stats.ok, stats.pend, stats.prog], backgroundColor:['#10b981','#ef4444','#3b82f6'], borderWidth: 2, borderColor: '#fff', hoverOffset: 5 }] }, 
            options: { ...chartOpts, cutout: '65%', plugins: { legend: { position: 'bottom', labels: { padding: 20, usePointStyle: true } }, datalabels: {display:false} }, onClick: (e, els, ch) => { if(els.length>0) showDataModal(ch.data.labels[els[0].index], d => { let st = ch.data.labels[els[0].index]; if(st==='Cerradas') return d.status==='realizada'; if(st==='Programadas') return d.status==='programado'; return d.status==='pendiente' || d.status==='en proceso'; }); } }
        });
        
        new Chart(getFreshCanvas('chart2'), { 
            type: 'pie', 
            data: { labels:Object.keys(stats.cCounts), datasets:[{ data:Object.values(stats.cCounts), backgroundColor:['#3b82f6','#8b5cf6','#ec4899','#14b8a6','#f97316'], borderWidth: 2, borderColor: '#fff', hoverOffset: 5 }] }, 
            options: { ...chartOpts, plugins: { legend: { position: 'right', labels: { padding: 15, usePointStyle: true } }, datalabels: {display:false} }, onClick: (e, els, ch) => { if(els.length>0) showDataModal(ch.data.labels[els[0].index], d => d.clase === ch.data.labels[els[0].index]); } }
        });
        
        const sortedEx = Object.entries(stats.ex).sort((a,b)=>(b[1].ok+b[1].pend)-(a[1].ok+a[1].pend)).slice(0,12);
        new Chart(getFreshCanvas('chart3'), { 
            type: 'bar', 
            data: { labels: sortedEx.map(x=>x[0]), datasets: [ { label:'Pendientes', data:sortedEx.map(x=>x[1].pend), backgroundColor:'#ef4444', borderRadius: 4, barPercentage: 0.7 }, { label:'Cerradas', data:sortedEx.map(x=>x[1].ok), backgroundColor:'#10b981', borderRadius: 4, barPercentage: 0.7 } ]}, 
            options: { ...chartOpts, indexAxis: 'y', scales: { x: { stacked: true, grid: { color: '#f1f5f9' } }, y: { stacked: true, grid: { display: false } } }, plugins: { legend: { position: 'top', labels: { usePointStyle: true } }, datalabels: {display:false} }, onClick: (e, els, ch) => { if(els.length>0) showDataModal(ch.data.labels[els[0].index], d => d.ejecutor === ch.data.labels[els[0].index]); } }
        });

        const sortedLocs = Object.entries(stats.loc).sort((a,b)=>b[1]-a[1]).slice(0,12);
        new Chart(getFreshCanvas('chart4'), {
            type: 'bar',
            data: { labels: sortedLocs.map(x=>x[0]), datasets: [ { label: 'Total Hallazgos', data: sortedLocs.map(x=>x[1]), backgroundColor:'#3b82f6', borderRadius: 6, barPercentage: 0.6 } ]},
            options: { ...chartOpts, indexAxis: 'y', scales: gridHideY, plugins: { legend: { display: false }, datalabels: {display:false} }, onClick: (e, els, ch) => { if(els.length>0) showDataModal(ch.data.labels[els[0].index], d => d.ubicacion === ch.data.labels[els[0].index], 'clase'); } }
        });
    }

    function drawRowCharts(data) {
        if(!data) return;

        const semVal = document.getElementById('f_semana') ? document.getElementById('f_semana').value : 'ALL';
        document.getElementById('row_week_title').innerText = semVal === "ALL" ? "Semanas: " + weeks.join(' y ') : "Semana " + semVal;
        
        let stats = {
            mtto: { total: 0, ok: 0 },
            aseo: { total: 0, ok: 0 },
            panaderia: {
                'L1': { mtto: {tot:0, ok:0} },
                'L2': { mtto: {tot:0, ok:0} },
                'L3': { mtto: {tot:0, ok:0} },
                'L4': { mtto: {tot:0, ok:0} },
                'L5': { mtto: {tot:0, ok:0} }
            },
            dely: {
                'Pizza': { mtto: {tot:0, ok:0} },
                'Bolleria': { mtto: {tot:0, ok:0} },
                'Empanadas': { mtto: {tot:0, ok:0} }
            }
        };
        
        data.forEach(d => {
            let isOk = (d.status === 'realizada');
            let isAseo = isAseoAct(d);
            let isMtto = !isAseo; 
            
            if (isAseo) { stats.aseo.total++; if(isOk) stats.aseo.ok++; }
            if (isMtto) { stats.mtto.total++; if(isOk) stats.mtto.ok++; }
            
            let pLoc = getPLoc(d);
            if (pLoc && isMtto) { stats.panaderia[pLoc].mtto.tot++; if(isOk) stats.panaderia[pLoc].mtto.ok++; }
            
            let dLoc = getDLoc(d);
            if (dLoc && isMtto) { stats.dely[dLoc].mtto.tot++; if(isOk) stats.dely[dLoc].mtto.ok++; }
        });

        const getPerc = (ok, tot) => tot > 0 ? Math.round((ok/tot)*100) : 0;
        
        const chartIds = ['row_chart1', 'row_chart2', 'row_chart3', 'row_chart4', 'row_chart5'];
        chartIds.forEach(id => {
            if (chartInstances[id]) { chartInstances[id].destroy(); chartInstances[id] = null; }
        });

        const commonOptsRow = { 
            maintainAspectRatio: false, responsive: true, animation: { duration: 1000 },
            plugins: { 
                legend: { position: 'top', labels: { usePointStyle: true } },
                datalabels: { 
                    display: (ctx) => ctx.dataset.data[ctx.dataIndex] > 0, 
                    color: '#fff', font: { weight: 'bold', size: 13 },
                    formatter: (val) => val + '%'
                }
            }
        };

        let totalAct = stats.mtto.total + stats.aseo.total;
        let pMttoTot = getPerc(stats.mtto.total, totalAct);
        let pAseoTot = getPerc(stats.aseo.total, totalAct);
        
        chartInstances['row_chart1'] = new Chart(getFreshCanvas('row_chart1'), {
            type: 'pie',
            data: { labels: ['Mantenimiento', 'Aseo'], datasets: [{ data: [pMttoTot, pAseoTot], backgroundColor: ['#8b5cf6', '#3b82f6'], borderWidth: 2, borderColor: '#fff' }] },
            options: { 
                ...commonOptsRow, 
                plugins: { ...commonOptsRow.plugins, legend: { position: 'bottom', labels: { usePointStyle: true } } },
                onClick: (e, els, ch) => { 
                    if(els.length>0) {
                        let label = ch.data.labels[els[0].index];
                        showDataModal(label, d => label === 'Aseo' ? isAseoAct(d) : !isAseoAct(d));
                    }
                }
            }
        });

        let pMttoCump = getPerc(stats.mtto.ok, stats.mtto.total);
        chartInstances['row_chart2'] = new Chart(getFreshCanvas('row_chart2'), {
            type: 'bar',
            data: { labels: ['Cumplimiento MTTO'], datasets: [{ label: 'Cerradas', data: [pMttoCump], backgroundColor: '#8b5cf6', barPercentage: 0.5, borderRadius: 6 }] },
            options: { 
                ...commonOptsRow, indexAxis: 'y', scales: { x: { max: 100, grid: {color:'#f1f5f9'} }, y: { grid: {display:false} } }, 
                plugins: { ...commonOptsRow.plugins, legend: { display: false } },
                onClick: (e, els, ch) => { if(els.length>0) showDataModal('Mantenimiento (General)', d => !isAseoAct(d)); }
            }
        });

        let pAseoCump = getPerc(stats.aseo.ok, stats.aseo.total);
        chartInstances['row_chart3'] = new Chart(getFreshCanvas('row_chart3'), {
            type: 'bar',
            data: { labels: ['Cumplimiento ASEO'], datasets: [{ label: 'Cerradas', data: [pAseoCump], backgroundColor: '#3b82f6', barPercentage: 0.5, borderRadius: 6 }] },
            options: { 
                ...commonOptsRow, indexAxis: 'y', scales: { x: { max: 100, grid: {color:'#f1f5f9'} }, y: { grid: {display:false} } }, 
                plugins: { ...commonOptsRow.plugins, legend: { display: false } },
                onClick: (e, els, ch) => { if(els.length>0) showDataModal('Aseo (General)', d => isAseoAct(d)); }
            }
        });

        const pLabels = ['L1', 'L2', 'L3', 'L4', 'L5'];
        const pMttoData = pLabels.map(l => getPerc(stats.panaderia[l].mtto.ok, stats.panaderia[l].mtto.tot));
        
        chartInstances['row_chart4'] = new Chart(getFreshCanvas('row_chart4'), {
            type: 'bar',
            data: { 
                labels: pLabels, 
                datasets: [ { label: '% Cumpl. Mtto', data: pMttoData, backgroundColor: '#8b5cf6', borderRadius: 4, barPercentage: 0.8, categoryPercentage: 0.8 } ] 
            },
            options: { 
                ...commonOptsRow, indexAxis: 'y', scales: { x: { max: 100, grid: {color:'#f1f5f9'} }, y: { grid: {display:false} } }, 
                plugins: { ...commonOptsRow.plugins, legend: { display: false } },
                onClick: (e, els, ch) => { 
                    if(els.length>0) {
                        let label = ch.data.labels[els[0].index];
                        showDataModal('Panadería MTTO - ' + label, d => !isAseoAct(d) && getPLoc(d) === label);
                    }
                }
            }
        });

        const dLabels = ['Pizza', 'Bolleria', 'Empanadas'];
        const dMttoData = dLabels.map(l => getPerc(stats.dely[l].mtto.ok, stats.dely[l].mtto.tot));
        
        chartInstances['row_chart5'] = new Chart(getFreshCanvas('row_chart5'), {
            type: 'bar',
            data: { 
                labels: dLabels, 
                datasets: [ { label: '% Cumpl. Mtto', data: dMttoData, backgroundColor: '#8b5cf6', borderRadius: 4, barPercentage: 0.8, categoryPercentage: 0.8 } ] 
            },
            options: { 
                ...commonOptsRow, indexAxis: 'y', scales: { x: { max: 100, grid: {color:'#f1f5f9'} }, y: { grid: {display:false} } }, 
                plugins: { ...commonOptsRow.plugins, legend: { display: false } },
                onClick: (e, els, ch) => { 
                    if(els.length>0) {
                        let label = ch.data.labels[els[0].index];
                        showDataModal('Dely MTTO - ' + label, d => !isAseoAct(d) && getDLoc(d) === label);
                    }
                }
            }
        });
    }

    window.onload = () => {
        buildFilters();
        applyFilters();
    };
    </script>
</body></html>"""
    
    # Inyectamos las variables de Python usando replace() de forma segura
    full_html = html_template.replace("__FECHA_ACTUAL__", fecha_actual)
    full_html = full_html.replace("__DB_JSON_DATA__", json.dumps(db_json))
    
    with open(OUTPUT_HTML, "w", encoding="utf-8") as f: 
        f.write(full_html)
        
    print(f"\n✅ REPORTE GENERADO CON ÉXITO: {OUTPUT_HTML}")
    # Eliminamos webbrowser.open() porque en el servidor no hay pantalla

if __name__ == "__main__":
    main()