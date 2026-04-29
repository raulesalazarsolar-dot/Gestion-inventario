import io
import base64
import json
import os
import pandas as pd
from urllib.parse import unquote
from PIL import Image

# Librerías de SharePoint
from office365.sharepoint.client_context import ClientContext
from office365.runtime.auth.user_credential import UserCredential

# ==========================================
# 1. CONFIGURACIÓN
# ==========================================
SITE_URL = "https://teams.wal-mart.com/sites/MejoradedesempeoprocesosclaveP5-Carnes"

# Rutas relativas en SharePoint
EXCEL_URL = "/sites/MejoradedesempeoprocesosclaveP5-Carnes/Documentos compartidos/Asset Strategy/Levantamiento Repuestos Mantenimiento/Levantamiento de Repuestos Bodega Mantenimiento/Gestión de Inventario (solo datos) y ficha.xlsm"
FOTOS_FOLDER_URL = "/sites/MejoradedesempeoprocesosclaveP5-Carnes/Documentos compartidos/Asset Strategy/Levantamiento Repuestos Mantenimiento/Levantamiento de Repuestos Bodega Mantenimiento/Fotos Repuestos"

# Usando tus credenciales que ya sabemos que funcionan
USERNAME = os.environ.get("SP_USERNAME", "r0r0noi@cl.wal-mart.com")
PASSWORD = os.environ.get("SP_PASSWORD", "fiXed.sPout+8")

OUTPUT_HTML = "index.html"

# ==========================================
# 2. FUNCIONES DE IMÁGENES
# ==========================================
def obtener_mapa_fotos(ctx, folder_url):
    print("📸 Mapeando carpeta de fotos...")
    try:
        folder = ctx.web.get_folder_by_server_relative_url(folder_url)
        files = folder.files.get().execute_query()
        return {os.path.splitext(f.name)[0]: f.serverRelativeUrl for f in files}
    except Exception as e:
        print(f"⚠️ Error al mapear fotos: {e}")
        return {}

def descargar_y_comprimir_foto(ctx, relative_url):
    try:
        file_content = io.BytesIO()
        ctx.web.get_file_by_server_relative_url(relative_url).download(file_content).execute_query()
        file_content.seek(0)
        
        if len(file_content.getvalue()) > 0:
            with Image.open(file_content) as img:
                if img.mode != "RGB": img = img.convert("RGB")
                img.thumbnail((400, 400))
                buf = io.BytesIO()
                img.save(buf, format='JPEG', quality=60)
                return f"data:image/jpeg;base64,{base64.b64encode(buf.getvalue()).decode('utf-8')}"
    except Exception:
        pass
    return None

# ==========================================
# 3. EXTRACCIÓN Y PROCESAMIENTO
# ==========================================
def main():
    try:
        print("🚀 CONECTANDO A SHAREPOINT...")
        ctx = ClientContext(SITE_URL).with_credentials(UserCredential(USERNAME, PASSWORD))
        
        mapa_fotos = obtener_mapa_fotos(ctx, FOTOS_FOLDER_URL)
        
        print("📊 Descargando Base de Datos Excel...")
        response = io.BytesIO()
        ctx.web.get_file_by_server_relative_url(EXCEL_URL).download(response).execute_query()
        response.seek(0)
        
        df = pd.read_excel(response, sheet_name="Gestión Inventario")
        df = df.fillna("") 
        
        print(f"✅ Se leyeron {len(df)} registros. Procesando datos y fotos...")
        
        db_json = {}
        for index, row in df.iterrows():
            cod_interno = str(row.get('Código\n interno \nproyecto', '')).strip()
            if not cod_interno or cod_interno == "nan": continue 
            
            cod_foto = str(row.get('Código Fotografía asociada', '')).strip()
            img_b64 = None
            
            if cod_foto and cod_foto in mapa_fotos:
                print(f"   ... Descargando foto {cod_foto}", end='\r')
                img_b64 = descargar_y_comprimir_foto(ctx, mapa_fotos[cod_foto])
            
            db_json[cod_interno] = {
                "codigo_interno": cod_interno,
                "codigo_sap": str(row.get('Código SAP', '')).replace('.0', ''),
                "nombre": str(row.get('Nombre repuesto', '')),
                "ubicacion_fisica": str(row.get('Ubicación física', '')),
                "ubicacion_sap": str(row.get('Ubicacion SAP', '')),
                "dimensiones": str(row.get('Dimensiones (Alto – Largo – Ancho) cm ', '')),
                "peso": str(row.get('Peso total contenido (gr)', '')).replace('.0', ''),
                "unidad": str(row.get('Unidad de contenido', '')),
                "descripcion": str(row.get('Descripción técnica', '')),
                "categoria": str(row.get('Categoría/Familia', '')),
                "planta": str(row.get('Planta asociada', '')),
                "criticidad": str(row.get('Criticidad', '')),
                "stock": str(row.get('Repetido', '')).replace('.0', ''),
                "img_base64": img_b64
            }
            
        print("\n✅ Procesamiento finalizado. Construyendo HTML...")
        generar_html_inventario(db_json)

    except Exception as e: 
        print(f"\n❌ Error Fatal: {e}")
        import traceback
        traceback.print_exc()

# ==========================================
# 4. GENERADOR HTML
# ==========================================
def generar_html_inventario(db_json):
    html_template = """<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Inventario de Repuestos</title>
    <style>
        :root { --primary: #0f172a; --secondary: #334155; --accent: #0ea5e9; --bg: #f8fafc; --border: #e2e8f0; }
        body { font-family: 'Segoe UI', system-ui, sans-serif; background: var(--bg); margin: 0; display: flex; height: 100vh; overflow: hidden; }
        * { box-sizing: border-box; }
        
        .sidebar { width: 300px; background: white; border-right: 1px solid var(--border); display: flex; flex-direction: column; }
        .header { padding: 20px; background: var(--primary); color: white; }
        .header h2 { margin: 0; font-size: 1.2rem; }
        .filters { padding: 20px; overflow-y: auto; flex: 1; }
        .f-group { margin-bottom: 15px; }
        .f-group label { display: block; font-size: 0.8rem; font-weight: 700; color: var(--secondary); margin-bottom: 5px; text-transform: uppercase;}
        select, input { width: 100%; padding: 8px; border: 1px solid var(--border); border-radius: 4px; }
        
        .main-content { flex: 1; display: flex; flex-direction: column; overflow: hidden; }
        .top-bar { padding: 15px 20px; background: white; border-bottom: 1px solid var(--border); display: flex; justify-content: space-between; align-items: center; }
        
        .grid-container { padding: 20px; overflow-y: auto; flex: 1; display: grid; grid-template-columns: repeat(auto-fill, minmax(280px, 1fr)); gap: 20px; align-content: start; }
        .card { background: white; border: 1px solid var(--border); border-radius: 8px; overflow: hidden; cursor: pointer; transition: transform 0.2s, box-shadow 0.2s; display: flex; flex-direction: column; }
        .card:hover { transform: translateY(-3px); box-shadow: 0 10px 15px -3px rgba(0,0,0,0.1); }
        .card-img-wrapper { height: 200px; background: #e2e8f0; display: flex; align-items: center; justify-content: center; overflow: hidden; }
        .card-img { width: 100%; height: 100%; object-fit: cover; }
        .no-img { color: #94a3b8; font-style: italic; }
        .card-body { padding: 15px; display: flex; flex-direction: column; flex: 1; }
        .c-tag { background: #e0f2fe; color: #0284c7; padding: 3px 8px; border-radius: 4px; font-size: 0.7rem; font-weight: bold; align-self: flex-start; margin-bottom: 8px; }
        .c-title { font-weight: 700; font-size: 1rem; color: var(--primary); margin: 0 0 10px 0; }
        .c-info { font-size: 0.8rem; color: var(--secondary); margin: 2px 0; }
        
        .modal { display: none; position: fixed; top:0; left:0; width: 100%; height: 100%; background: rgba(15,23,42,0.8); z-index: 1000; align-items: center; justify-content: center; backdrop-filter: blur(4px); }
        .modal-content { background: white; width: 90%; max-width: 900px; border-radius: 12px; display: flex; overflow: hidden; max-height: 80vh; }
        .m-img-sec { width: 40%; background: #f1f5f9; display: flex; align-items: center; justify-content: center; }
        .m-img-sec img { max-width: 100%; max-height: 100%; object-fit: contain; }
        .m-data-sec { width: 60%; padding: 30px; overflow-y: auto; position: relative; }
        .close-btn { position: absolute; top: 15px; right: 20px; font-size: 1.5rem; cursor: pointer; border: none; background: none; color: var(--secondary); }
        .m-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 15px; margin-top: 20px; }
        .m-item small { display: block; color: #64748b; font-size: 0.75rem; text-transform: uppercase; font-weight: bold; }
        .m-item strong { color: var(--primary); font-size: 0.95rem; }
    </style>
</head>
<body>

    <div class="sidebar">
        <div class="header"><h2>⚙️ Inventario Repuestos</h2></div>
        <div class="filters" id="filters_container"></div>
    </div>

    <div class="main-content">
        <div class="top-bar">
            <div>Resultados: <strong id="k_count">0</strong> repuestos</div>
            <input type="text" id="search_input" placeholder="🔍 Buscar nombre, SAP, ubicación..." onkeyup="applyFilters()" style="width: 300px;">
        </div>
        <div class="grid-container" id="grid_container"></div>
    </div>

    <div class="modal" id="detail_modal" onclick="if(event.target===this) this.style.display='none'">
        <div class="modal-content" id="modal_body"></div>
    </div>

    <script>
        const db = __DB_JSON_DATA__;
        const records = Object.values(db);
        
        function getUnique(key) {
            return [...new Set(records.map(x => x[key]).filter(x => x && x !== 'nan'))].sort();
        }

        function buildFilters() {
            const container = document.getElementById('filters_container');
            const createSelect = (id, label, options) => {
                let sel = `<div class="f-group"><label>${label}</label><select id="${id}" onchange="applyFilters()">`;
                sel += `<option value="ALL">Todos</option>`;
                options.forEach(o => sel += `<option value="${o}">${o}</option>`);
                sel += `</select></div>`;
                return sel;
            };

            container.innerHTML = 
                createSelect('f_cat', 'Categoría / Familia', getUnique('categoria')) +
                createSelect('f_planta', 'Planta Asociada', getUnique('planta')) +
                createSelect('f_crit', 'Criticidad', getUnique('criticidad'));
        }

        function applyFilters() {
            const fCat = document.getElementById('f_cat').value;
            const fPlanta = document.getElementById('f_planta').value;
            const fCrit = document.getElementById('f_crit').value;
            const search = document.getElementById('search_input').value.toLowerCase();

            const filtered = records.filter(d => {
                if (fCat !== 'ALL' && d.categoria !== fCat) return false;
                if (fPlanta !== 'ALL' && d.planta !== fPlanta) return false;
                if (fCrit !== 'ALL' && d.criticidad !== fCrit) return false;
                if (search && !`${d.nombre} ${d.codigo_sap} ${d.ubicacion_fisica}`.toLowerCase().includes(search)) return false;
                return true;
            });

            document.getElementById('k_count').innerText = filtered.length;
            renderGrid(filtered);
        }

        function renderGrid(data) {
            const container = document.getElementById('grid_container');
            container.innerHTML = '';

            data.forEach(d => {
                const imgHtml = d.img_base64 ? `<img src="${d.img_base64}" class="card-img">` : `<span class="no-img">📷 Sin foto</span>`;
                const card = document.createElement('div');
                card.className = 'card';
                card.onclick = () => openModal(d);
                card.innerHTML = `
                    <div class="card-img-wrapper">${imgHtml}</div>
                    <div class="card-body">
                        <span class="c-tag">${d.codigo_sap || 'Sin SAP'}</span>
                        <h3 class="c-title">${d.nombre || 'Sin Nombre'}</h3>
                        <p class="c-info">📍 <b>Ubicación:</b> ${d.ubicacion_fisica}</p>
                        <p class="c-info">📦 <b>Stock:</b> ${d.stock || '0'}</p>
                    </div>
                `;
                container.appendChild(card);
            });
        }

        function openModal(d) {
            const modal = document.getElementById('detail_modal');
            const body = document.getElementById('modal_body');
            const imgHtml = d.img_base64 ? `<img src="${d.img_base64}">` : `<div style="color:#94a3b8; font-style:italic;">📷 Sin foto</div>`;

            body.innerHTML = `
                <div class="m-img-sec">${imgHtml}</div>
                <div class="m-data-sec">
                    <button class="close-btn" onclick="document.getElementById('detail_modal').style.display='none'">&times;</button>
                    <span style="background:#e0f2fe; color:#0284c7; padding:4px 8px; border-radius:4px; font-weight:bold; font-size:0.8rem;">SAP: ${d.codigo_sap || '-'}</span>
                    <h2 style="color:#0f172a; margin-top:10px;">${d.nombre}</h2>
                    <p style="color:#64748b; font-size:0.9rem;">${d.descripcion}</p>
                    <div class="m-grid">
                        <div class="m-item"><small>Ubicación Física</small><strong>${d.ubicacion_fisica || '-'}</strong></div>
                        <div class="m-item"><small>Ubicación SAP</small><strong>${d.ubicacion_sap || '-'}</strong></div>
                        <div class="m-item"><small>Categoría</small><strong>${d.categoria || '-'}</strong></div>
                        <div class="m-item"><small>Planta</small><strong>${d.planta || '-'}</strong></div>
                        <div class="m-item"><small>Criticidad</small><strong>${d.criticidad || '-'}</strong></div>
                        <div class="m-item"><small>Dimensiones</small><strong>${d.dimensiones || '-'}</strong></div>
                        <div class="m-item"><small>Peso / Unidad</small><strong>${d.peso} gr / ${d.unidad}</strong></div>
                    </div>
                </div>
            `;
            modal.style.display = 'flex';
        }

        window.onload = () => { buildFilters(); applyFilters(); };
    </script>
</body>
</html>"""

    full_html = html_template.replace("__DB_JSON_DATA__", json.dumps(db_json))
    with open(OUTPUT_HTML, "w", encoding="utf-8") as f: f.write(full_html)
    print(f"\n✅ DASHBOARD GENERADO CON ÉXITO: {OUTPUT_HTML}")

if __name__ == "__main__":
    main()
