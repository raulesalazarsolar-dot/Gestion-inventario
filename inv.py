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
# 1. CONFIGURACIÓN (CON CREDENCIALES)
# ==========================================
SITE_URL = "https://teams.wal-mart.com/sites/MejoradedesempeoprocesosclaveP5-Carnes"

# Rutas de búsqueda
RUTA_ES = "/sites/MejoradedesempeoprocesosclaveP5-Carnes/Documentos compartidos/Asset Strategy/Levantamiento Repuestos Mantenimiento/Levantamiento de Repuestos Bodega Mantenimiento/Fotos Repuestos"
RUTA_EN = "/sites/MejoradedesempeoprocesosclaveP5-Carnes/Shared Documents/Asset Strategy/Levantamiento Repuestos Mantenimiento/Levantamiento de Repuestos Bodega Mantenimiento/Fotos Repuestos"
EXCEL_URL = "/sites/MejoradedesempeoprocesosclaveP5-Carnes/Documentos compartidos/Asset Strategy/Levantamiento Repuestos Mantenimiento/Levantamiento de Repuestos Bodega Mantenimiento/Gestión de Inventario (solo datos) y ficha.xlsm"

USERNAME = "r0r0noi@cl.wal-mart.com"
PASSWORD = "fiXed.sPout+8"

OUTPUT_HTML = "index.html"

# ==========================================
# 2. FUNCIONES DE IMÁGENES
# ==========================================
def obtener_mapa_fotos(ctx):
    print("📸 Buscando carpeta de fotos en SharePoint...")
    mapa = {}
    rutas = [RUTA_ES, RUTA_EN]
    
    for folder_url in rutas:
        try:
            folder = ctx.web.get_folder_by_server_relative_url(folder_url)
            files = folder.files.get().execute_query()
            if len(files) > 0:
                print(f"✅ ¡ÉXITO! Encontrada la carpeta con {len(files)} fotos.")
                for f in files:
                    nombre_base = str(os.path.splitext(f.name)[0]).strip().lower()
                    mapa[nombre_base] = f.serverRelativeUrl
                return mapa
        except:
            pass
            
    print("⚠️ ADVERTENCIA: No se pudo encontrar la carpeta de fotos. Revisa la ruta en SharePoint.")
    return mapa

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
        
        mapa_fotos = obtener_mapa_fotos(ctx)
        
        print("📊 Descargando Base de Datos Excel...")
        try:
            response = io.BytesIO()
            ctx.web.get_file_by_server_relative_url(EXCEL_URL).download(response).execute_query()
        except:
            excel_en = EXCEL_URL.replace("Documentos compartidos", "Shared Documents")
            response = io.BytesIO()
            ctx.web.get_file_by_server_relative_url(excel_en).download(response).execute_query()
            
        response.seek(0)
        df = pd.read_excel(response, sheet_name="Gestión Inventario")
        
        # TRUCO ANTIMAGIA: Limpiar los nombres de las columnas del Excel
        # Esto elimina saltos de línea extraños o espacios dobles
        df.columns = [' '.join(str(c).split()).lower() for c in df.columns]
        df = df.fillna("") 
        
        # Buscar los nombres exactos de las columnas clave
        col_interno = next((c for c in df.columns if 'interno' in c and 'proyecto' in c), None)
        col_foto = next((c for c in df.columns if 'fotografía' in c or 'fotografia' in c), None)
        
        if not col_interno or not col_foto:
            print("❌ ERROR: No pude encontrar las columnas 'Código interno proyecto' o 'Código Fotografía asociada' en el Excel.")
            print(f"Columnas detectadas: {list(df.columns)}")
            return

        print(f"✅ Se leyeron {len(df)} filas del Excel. Iniciando cruce de datos...")
        
        db_json = {}
        fotos_exito = 0

        for index, row in df.iterrows():
            cod_interno = str(row.get(col_interno, '')).strip()
            if not cod_interno or cod_interno == "nan": continue 
            
            val_excel_foto = str(row.get(col_foto, '')).strip().lower()
            cod_foto = os.path.splitext(val_excel_foto)[0] if val_excel_foto != "nan" else ""
            
            img_b64 = None
            if cod_foto and cod_foto in mapa_fotos:
                print(f"   ... Procesando imagen para: {cod_foto}", end='\r')
                img_b64 = descargar_y_comprimir_foto(ctx, mapa_fotos[cod_foto])
                if img_b64: fotos_exito += 1
            
            # Buscar otras columnas con tolerancia a fallos
            def get_val(keywords):
                col = next((c for c in df.columns if any(k in c for k in keywords)), None)
                return str(row.get(col, '')).replace('.0', '') if col else ''

            db_json[cod_interno] = {
                "codigo_interno": cod_interno,
                "codigo_sap": get_val(['sap']),
                "nombre": get_val(['nombre repuesto']),
                "ubicacion_fisica": get_val(['ubicación física', 'ubicacion fisica']),
                "ubicacion_sap": get_val(['ubicacion sap', 'ubicación sap']),
                "dimensiones": get_val(['dimensiones']),
                "peso": get_val(['peso']),
                "unidad": get_val(['unidad']),
                "descripcion": get_val(['descripción técnica', 'descripcion tecnica']),
                "categoria": get_val(['categoría', 'categoria']),
                "planta": get_val(['planta']),
                "criticidad": get_val(['criticidad']),
                "stock": get_val(['repetido']),
                "img_base64": img_b64
            }
            
        print(f"\n✅ PROCESO COMPLETADO. Fotos vinculadas con éxito: {fotos_exito}")
        generar_html_inventario(db_json)

    except Exception as e: 
        print(f"\n❌ ERROR FATAL: {e}")
        import traceback
        traceback.print_exc()

# ==========================================
# 4. GENERADOR HTML (CATÁLOGO VISUAL)
# ==========================================
def generar_html_inventario(db_json):
    html_template = """<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Inventario de Repuestos | Mantenimiento</title>
    <style>
        :root { --primary: #0f172a; --secondary: #334155; --accent: #0ea5e9; --bg: #f8fafc; --border: #e2e8f0; }
        body { font-family: 'Segoe UI', sans-serif; background: var(--bg); margin: 0; display: flex; height: 100vh; overflow: hidden; }
        * { box-sizing: border-box; }
        
        .sidebar { width: 300px; background: white; border-right: 1px solid var(--border); display: flex; flex-direction: column; }
        .header { padding: 20px; background: var(--primary); color: white; }
        .header h2 { margin: 0; font-size: 1.1rem; }
        .filters { padding: 20px; overflow-y: auto; flex: 1; }
        .f-group { margin-bottom: 15px; }
        .f-group label { display: block; font-size: 0.75rem; font-weight: 700; color: var(--secondary); margin-bottom: 5px; text-transform: uppercase;}
        select, input { width: 100%; padding: 10px; border: 1px solid var(--border); border-radius: 6px; font-size: 0.85rem; }
        
        .main-content { flex: 1; display: flex; flex-direction: column; overflow: hidden; }
        .top-bar { padding: 15px 25px; background: white; border-bottom: 1px solid var(--border); display: flex; justify-content: space-between; align-items: center; }
        
        .grid-container { padding: 25px; overflow-y: auto; flex: 1; display: grid; grid-template-columns: repeat(auto-fill, minmax(260px, 1fr)); gap: 20px; align-content: start; }
        .card { background: white; border: 1px solid var(--border); border-radius: 10px; overflow: hidden; cursor: pointer; transition: 0.2s; display: flex; flex-direction: column; }
        .card:hover { transform: translateY(-4px); box-shadow: 0 12px 20px -5px rgba(0,0,0,0.1); border-color: var(--accent); }
        .card-img-wrapper { height: 180px; background: #f1f5f9; display: flex; align-items: center; justify-content: center; }
        .card-img { width: 100%; height: 100%; object-fit: contain; padding: 10px; }
        .no-img { color: #94a3b8; font-style: italic; font-size: 0.85rem; }
        .card-body { padding: 15px; display: flex; flex-direction: column; flex: 1; }
        .c-tag { background: #eff6ff; color: #0284c7; padding: 3px 8px; border-radius: 4px; font-size: 0.7rem; font-weight: 800; align-self: flex-start; margin-bottom: 10px; }
        .c-title { font-weight: 700; font-size: 0.95rem; color: var(--primary); margin: 0 0 12px 0; line-height: 1.3; }
        .c-info { font-size: 0.8rem; color: var(--secondary); margin: 3px 0; display: flex; justify-content: space-between; }
        
        .modal { display: none; position: fixed; top:0; left:0; width: 100%; height: 100%; background: rgba(15,23,42,0.85); z-index: 1000; align-items: center; justify-content: center; backdrop-filter: blur(5px); }
        .modal-content { background: white; width: 95%; max-width: 950px; border-radius: 16px; display: flex; overflow: hidden; max-height: 85vh; box-shadow: 0 25px 50px -12px rgba(0,0,0,0.5); }
        .m-img-sec { width: 45%; background: #f8fafc; display: flex; align-items: center; justify-content: center; border-right: 1px solid var(--border); }
        .m-img-sec img { max-width: 90%; max-height: 90%; object-fit: contain; }
        .m-data-sec { width: 55%; padding: 35px; overflow-y: auto; position: relative; }
        .close-btn { position: absolute; top: 20px; right: 25px; font-size: 2rem; cursor: pointer; border: none; background: none; color: #94a3b8; }
        .m-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 20px; margin-top: 25px; }
        .m-item small { display: block; color: #64748b; font-size: 0.7rem; text-transform: uppercase; font-weight: 800; margin-bottom: 4px; }
        .m-item strong { color: var(--primary); font-size: 1rem; }
    </style>
</head>
<body>

    <div class="sidebar">
        <div class="header"><h2>📦 Gestión de Repuestos</h2></div>
        <div class="filters" id="filters_container"></div>
    </div>

    <div class="main-content">
        <div class="top-bar">
            <div style="color: var(--secondary); font-weight: 600;">Encontrados: <span id="k_count" style="color: var(--accent); font-size: 1.1rem;">0</span></div>
            <input type="text" id="search_input" placeholder="🔍 Buscar por nombre, SAP o ubicación física..." onkeyup="applyFilters()" style="width: 400px;">
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
                sel += `<option value="ALL">TODOS</option>`;
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
                const matchBusqueda = `${d.nombre} ${d.codigo_sap} ${d.ubicacion_fisica}`.toLowerCase().includes(search);
                return search ? matchBusqueda : true;
            });

            document.getElementById('k_count').innerText = filtered.length;
            renderGrid(filtered);
        }

        function renderGrid(data) {
            const container = document.getElementById('grid_container');
            container.innerHTML = '';

            data.forEach(d => {
                const imgHtml = d.img_base64 
                    ? `<img src="${d.img_base64}" class="card-img">` 
                    : `<div class="no-img">📷 Sin fotografía</div>`;

                const card = document.createElement('div');
                card.className = 'card';
                card.onclick = () => openModal(d);
                card.innerHTML = `
                    <div class="card-img-wrapper">${imgHtml}</div>
                    <div class="card-body">
                        <span class="c-tag">SAP: ${d.codigo_sap || '---'}</span>
                        <h3 class="c-title">${d.nombre || 'Sin Nombre'}</h3>
                        <div class="c-info"><span>📍 Ubicación:</span> <b>${d.ubicacion_fisica}</b></div>
                        <div class="c-info"><span>📦 Stock:</span> <b>${d.stock}</b></div>
                    </div>
                `;
                container.appendChild(card);
            });
        }

        function openModal(d) {
            const modal = document.getElementById('detail_modal');
            const body = document.getElementById('modal_body');
            const imgHtml = d.img_base64 
                ? `<img src="${d.img_base64}">` 
                : `<div style="color:#94a3b8; font-style:italic; font-size:1.2rem;">📷 Fotografía no disponible</div>`;

            body.innerHTML = `
                <div class="m-img-sec">${imgHtml}</div>
                <div class="m-data-sec">
                    <button class="close-btn" onclick="document.getElementById('detail_modal').style.display='none'">&times;</button>
                    <span style="background:#0ea5e9; color:white; padding:4px 10px; border-radius:6px; font-weight:bold; font-size:0.75rem;">SAP: ${d.codigo_sap || '---'}</span>
                    <h2 style="color:var(--primary); margin:15px 0 10px 0; font-size:1.6rem;">${d.nombre}</h2>
                    <p style="color:#64748b; font-size:0.95rem; line-height:1.5; background:#f1f5f9; padding:15px; border-radius:8px;">${d.descripcion || 'Sin descripción técnica detallada.'}</p>
                    <div class="m-grid">
                        <div class="m-item"><small>Ubicación Física</small><strong>${d.ubicacion_fisica || '---'}</strong></div>
                        <div class="m-item"><small>Ubicación SAP</small><strong>${d.ubicacion_sap || '---'}</strong></div>
                        <div class="m-item"><small>Categoría</small><strong>${d.categoria || '---'}</strong></div>
                        <div class="m-item"><small>Planta</small><strong>${d.planta || '---'}</strong></div>
                        <div class="m-item"><small>Criticidad</small><strong>${d.criticidad || '---'}</strong></div>
                        <div class="m-item"><small>Dimensiones</small><strong>${d.dimensiones || '---'}</strong></div>
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

if __name__ == "__main__":
    main()
