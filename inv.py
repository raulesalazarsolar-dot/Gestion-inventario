import io
import json
import os
import pandas as pd

# Librerías de SharePoint
from office365.sharepoint.client_context import ClientContext
from office365.runtime.auth.user_credential import UserCredential

# ==========================================
# 1. CONFIGURACIÓN
# ==========================================
SITE_URL = "https://teams.wal-mart.com/sites/MejoradedesempeoprocesosclaveP5-Carnes"
EXCEL_URL = "/sites/MejoradedesempeoprocesosclaveP5-Carnes/Documentos compartidos/Asset Strategy/Levantamiento Repuestos Mantenimiento/Levantamiento de Repuestos Bodega Mantenimiento/Gestión de Inventario (solo datos) y ficha.xlsm"

USERNAME = "r0r0noi@cl.wal-mart.com"
PASSWORD = "fiXed.sPout+8"

OUTPUT_HTML = "index.html"

# ==========================================
# 2. EXTRACCIÓN DEL EXCEL
# ==========================================
def main():
    try:
        print("🚀 CONECTANDO A SHAREPOINT...")
        ctx = ClientContext(SITE_URL).with_credentials(UserCredential(USERNAME, PASSWORD))
        
        print("📊 Descargando Base de Datos...")
        try:
            response = io.BytesIO()
            ctx.web.get_file_by_server_relative_url(EXCEL_URL).download(response).execute_query()
        except:
            excel_en = EXCEL_URL.replace("Documentos compartidos", "Shared Documents")
            response = io.BytesIO()
            ctx.web.get_file_by_server_relative_url(excel_en).download(response).execute_query()
            
        response.seek(0)
        df = pd.read_excel(response, sheet_name="Gestión Inventario")
        
        # Limpiar nombres de columnas (quitar saltos de línea y espacios)
        df.columns = [' '.join(str(c).split()).lower() for c in df.columns]
        df = df.fillna("") 
        
        col_interno = next((c for c in df.columns if 'interno' in c and 'proyecto' in c), None)
        col_foto = next((c for c in df.columns if 'fotografía' in c or 'fotografia' in c), None)
        
        print(f"✅ Se leyeron {len(df)} repuestos. Procesando...")
        
        db_json = {}

        for index, row in df.iterrows():
            cod_interno = str(row.get(col_interno, '')).strip()
            if not cod_interno or cod_interno == "nan": continue 
            
            # Match de foto
            val_excel_foto = str(row.get(col_foto, '')).strip().upper()
            cod_foto = os.path.splitext(val_excel_foto)[0] if val_excel_foto != "NAN" else ""
            img_ruta = f"fotos/{cod_foto}.jpg" if cod_foto else None
            
            # Helper para buscar datos en columnas con nombres variables
            def get_val(keywords):
                col = next((c for c in df.columns if any(k in c for k in keywords)), None)
                res = str(row.get(col, '')).strip()
                return res.replace('.0', '') if col else ''

            db_json[cod_interno] = {
                "id": cod_interno,
                "sap": get_val(['sap']),
                "nombre": get_val(['nombre repuesto']),
                "ubi_fisica": get_val(['ubicación física', 'ubicacion fisica']),
                "ubi_sap": get_val(['ubicacion sap', 'ubicación sap']),
                "dimensiones": get_val(['dimensiones']),
                "peso": get_val(['peso']),
                "unidad": get_val(['unidad']),
                "fecha": get_val(['fecha de levantamiento']),
                "facilidad": get_val(['facilidad para encontrar']),
                "observaciones": get_val(['observaciones de su uso']),
                "descripcion": get_val(['descripción técnica', 'descripcion tecnica']),
                "categoria": get_val(['categoría', 'categoria']),
                "planta": get_val(['planta']),
                "equipo": get_val(['equipo asociado']),
                "funcion": get_val(['función', 'funcion']),
                "criticidad": get_val(['criticidad']),
                "sustitutos": get_val(['sustitutos']),
                "estandar": get_val(['estándar o a medida', 'estandar o a medida']),
                "compatibilidad": get_val(['compatibilidad']),
                "vida_util": get_val(['vida útil', 'vida util']),
                "almacenamiento": get_val(['condiciones de almacenamiento']),
                "stock": get_val(['repetido']),
                "img": img_ruta
            }
            
        print("✅ Generando HTML mejorado...")
        generar_html_inventario(db_json)

    except Exception as e: 
        print(f"\n❌ ERROR FATAL: {e}")

# ==========================================
# 3. GENERADOR HTML
# ==========================================
def generar_html_inventario(db_json):
    html_template = """<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Dashboard Repuestos Walmart</title>
    <style>
        :root { --primary: #0071ce; --secondary: #475569; --accent: #ffc220; --bg: #f1f5f9; --border: #cbd5e1; }
        body { font-family: 'Segoe UI', system-ui, sans-serif; background: var(--bg); margin: 0; display: flex; height: 100vh; overflow: hidden; }
        * { box-sizing: border-box; }
        
        /* Sidebar con Filtros */
        .sidebar { width: 320px; background: white; border-right: 2px solid var(--border); display: flex; flex-direction: column; box-shadow: 4px 0 10px rgba(0,0,0,0.05); }
        .sidebar-header { padding: 20px; background: var(--primary); color: white; text-align: center; }
        .sidebar-header h2 { margin: 0; font-size: 1.2rem; letter-spacing: 1px; }
        
        .filters-container { padding: 15px; overflow-y: auto; flex: 1; }
        .f-group { margin-bottom: 12px; }
        .f-group label { display: block; font-size: 0.7rem; font-weight: 800; color: var(--secondary); margin-bottom: 4px; text-transform: uppercase; }
        select, input { width: 100%; padding: 8px; border: 1px solid var(--border); border-radius: 5px; font-size: 0.85rem; background: #f8fafc; }
        
        .btn-clear { width: 100%; padding: 10px; background: #ef4444; color: white; border: none; border-radius: 5px; font-weight: bold; cursor: pointer; margin-top: 10px; transition: 0.3s; }
        .btn-clear:hover { background: #dc2626; }

        /* Area Principal */
        .main { flex: 1; display: flex; flex-direction: column; overflow: hidden; }
        .top-nav { padding: 15px 25px; background: white; border-bottom: 1px solid var(--border); display: flex; justify-content: space-between; align-items: center; }
        
        .grid { padding: 20px; overflow-y: auto; flex: 1; display: grid; grid-template-columns: repeat(auto-fill, minmax(450px, 1fr)); gap: 20px; align-content: start; }
        
        /* Tarjeta Estilo Ficha */
        .card { background: white; border: 1px solid var(--border); border-radius: 8px; overflow: hidden; display: flex; height: 180px; transition: 0.2s; cursor: pointer; }
        .card:hover { border-color: var(--primary); transform: translateY(-2px); box-shadow: 0 5px 15px rgba(0,0,0,0.1); }
        
        .card-img-box { width: 160px; background: #f8fafc; display: flex; align-items: center; justify-content: center; border-right: 1px solid var(--border); position: relative; }
        .card-img-box img { max-width: 100%; max-height: 100%; object-fit: contain; padding: 5px; }
        .card-stock { position: absolute; top: 5px; left: 5px; background: var(--accent); color: #000; font-size: 0.65rem; font-weight: 800; padding: 2px 6px; border-radius: 3px; }
        
        .card-info { flex: 1; padding: 12px; display: flex; flex-direction: column; justify-content: space-between; position: relative; }
        .card-sap { color: var(--primary); font-weight: 800; font-size: 0.85rem; }
        .card-title { font-weight: 700; font-size: 0.95rem; color: #1e293b; margin: 4px 0; display: -webkit-box; -webkit-line-clamp: 2; -webkit-box-orient: vertical; overflow: hidden; }
        .card-details { font-size: 0.75rem; color: #64748b; line-height: 1.4; }
        .card-details b { color: #334155; }
        .card-crit { position: absolute; top: 12px; right: 12px; font-size: 0.6rem; font-weight: 800; padding: 2px 5px; border-radius: 4px; text-transform: uppercase; }
        
        /* Colores Criticidad */
        .crit-alta { background: #fee2e2; color: #b91c1c; border: 1px solid #f87171; }
        .crit-media { background: #fef3c7; color: #b45309; border: 1px solid #fbbf24; }
        .crit-baja { background: #dcfce7; color: #15803d; border: 1px solid #4ade80; }

        /* Modal Ficha Completa */
        .modal { display: none; position: fixed; top:0; left:0; width: 100%; height: 100%; background: rgba(0,0,0,0.7); z-index: 1000; align-items: center; justify-content: center; backdrop-filter: blur(4px); }
        .modal-content { background: white; width: 90%; max-width: 1000px; border-radius: 12px; display: flex; overflow: hidden; max-height: 90vh; }
        .modal-img-sec { width: 40%; background: #fff; display: flex; align-items: center; justify-content: center; padding: 20px; border-right: 1px solid var(--border); }
        .modal-img-sec img { max-width: 100%; max-height: 100%; object-fit: contain; }
        .modal-data-sec { width: 60%; padding: 30px; overflow-y: auto; position: relative; }
        .modal-data-sec h2 { color: var(--primary); margin-top: 0; border-bottom: 2px solid var(--accent); padding-bottom: 10px; }
        
        .ficha-table { width: 100%; border-collapse: collapse; font-size: 0.85rem; }
        .ficha-table td { padding: 6px 0; border-bottom: 1px solid #f1f5f9; }
        .ficha-label { font-weight: 800; color: var(--secondary); width: 40%; text-transform: uppercase; font-size: 0.75rem; }
        .ficha-val { color: #000; font-weight: 500; }

        .close-btn { position: absolute; top: 15px; right: 20px; font-size: 2rem; cursor: pointer; color: #94a3b8; border:none; background:none; }
    </style>
</head>
<body>

    <div class="sidebar">
        <div class="sidebar-header"><h2>GESTIÓN INVENTARIO</h2></div>
        <div class="filters-container" id="filters_ui">
            </div>
        <div style="padding: 15px;"><button class="btn-clear" onclick="resetFilters()">🧹 BORRAR FILTROS</button></div>
    </div>

    <div class="main">
        <div class="top-nav">
            <div style="font-weight: 700; color: var(--secondary)">Repuestos encontrados: <span id="k_count" style="color: var(--primary); font-size: 1.2rem;">0</span></div>
            <img src="https://upload.wikimedia.org/wikipedia/commons/c/ca/Walmart_logo.svg" height="25">
        </div>
        <div class="grid" id="grid_container"></div>
    </div>

    <div class="modal" id="modal_ficha" onclick="if(event.target===this) this.style.display='none'">
        <div class="modal-content" id="modal_body"></div>
    </div>

    <script>
        const db = __DB_JSON_DATA__;
        const records = Object.values(db);
        
        function getUnique(key) {
            return [...new Set(records.map(x => x[key]).filter(x => x && x !== ''))].sort();
        }

        function initUI() {
            const container = document.getElementById('filters_ui');
            
            const createFilter = (id, label, options, isSearch = false) => {
                let html = `<div class="f-group"><label>${label}</label>`;
                if(isSearch) {
                    html += `<input type="text" id="${id}" placeholder="Escriba para buscar..." onkeyup="applyFilters()">`;
                } else {
                    html += `<select id="${id}" onchange="applyFilters()"><option value="ALL">-- TODOS --</option>`;
                    options.forEach(o => html += `<option value="${o}">${o}</option>`);
                    html += `</select>`;
                }
                html += `</div>`;
                return html;
            };

            container.innerHTML = 
                createFilter('f_planta', 'Planta Asociada', getUnique('planta')) +
                createFilter('f_cat', 'Categoría / Familia', getUnique('categoria')) +
                createFilter('f_crit_sel', 'Criticidad', getUnique('criticidad')) +
                createFilter('f_sap', 'Código SAP', [], true) +
                createFilter('f_nom', 'Nombre Repuesto', [], true) +
                createFilter('f_ubif', 'Ubicación Física', [], true) +
                createFilter('f_ubis', 'Ubicación SAP', [], true);

            applyFilters();
        }

        function resetFilters() {
            document.querySelectorAll('.f-group select').forEach(s => s.value = 'ALL');
            document.querySelectorAll('.f-group input').forEach(i => i.value = '');
            applyFilters();
        }

        function applyFilters() {
            const fPlanta = document.getElementById('f_planta').value;
            const fCat = document.getElementById('f_cat').value;
            const fCrit = document.getElementById('f_crit_sel').value;
            const fSap = document.getElementById('f_sap').value.toLowerCase();
            const fNom = document.getElementById('f_nom').value.toLowerCase();
            const fUbif = document.getElementById('f_ubif').value.toLowerCase();
            const fUbis = document.getElementById('f_ubis').value.toLowerCase();

            const filtered = records.filter(d => {
                if (fPlanta !== 'ALL' && d.planta !== fPlanta) return false;
                if (fCat !== 'ALL' && d.categoria !== fCat) return false;
                if (fCrit !== 'ALL' && d.criticidad !== fCrit) return false;
                if (fSap && !d.sap.toLowerCase().includes(fSap)) return false;
                if (fNom && !d.nombre.toLowerCase().includes(fNom)) return false;
                if (fUbif && !d.ubi_fisica.toLowerCase().includes(fUbif)) return false;
                if (fUbis && !d.ubi_sap.toLowerCase().includes(fUbis)) return false;
                return true;
            });

            document.getElementById('k_count').innerText = filtered.length;
            renderGrid(filtered);
        }

        function renderGrid(data) {
            const container = document.getElementById('grid_container');
            container.innerHTML = '';

            data.forEach(d => {
                const critClass = d.criticidad.toLowerCase().includes('alta') ? 'crit-alta' : 
                                 (d.criticidad.toLowerCase().includes('media') ? 'crit-media' : 'crit-baja');
                
                const imgError = "this.onerror=null; this.src='https://placehold.co/200?text=Sin+Foto';";
                
                const card = document.createElement('div');
                card.className = 'card';
                card.onclick = () => openFicha(d);
                card.innerHTML = `
                    <div class="card-img-box">
                        <span class="card-stock">Stock: ${d.stock}</span>
                        <img src="${d.img || ''}" onerror="${imgError}" loading="lazy">
                    </div>
                    <div class="card-info">
                        <div class="card-sap">SAP: ${d.sap}</div>
                        <div class="card-crit ${critClass}">${d.criticidad}</div>
                        <div class="card-title">${d.nombre}</div>
                        <div class="card-details">
                            <b>Ubicación:</b> ${d.ubi_fisica}<br>
                            <b>Dimensiones:</b> ${d.dimensiones} cm<br>
                            <b>Peso:</b> ${d.peso} ${d.unidad}
                        </div>
                    </div>
                `;
                container.appendChild(card);
            });
        }

        function openFicha(d) {
            const modal = document.getElementById('modal_ficha');
            const body = document.getElementById('modal_body');
            
            const row = (label, val) => `<tr><td class="ficha-label">${label}</td><td class="ficha-val">: ${val || '---'}</td></tr>`;

            body.innerHTML = `
                <div class="modal-img-sec">
                    <img src="${d.img || ''}" onerror="this.src='https://placehold.co/400?text=Sin+Imagen+Disponible'">
                </div>
                <div class="modal-data-sec">
                    <button class="close-btn" onclick="document.getElementById('modal_ficha').style.display='none'">&times;</button>
                    <h2>FICHA TÉCNICA REPUESTO</h2>
                    <table class="ficha-table">
                        ${row('Código SAP', d.sap)}
                        ${row('Nombre Repuesto', d.nombre)}
                        ${row('Ubicación Física', d.ubi_fisica)}
                        ${row('Ubicación SAP', d.ubi_sap)}
                        ${row('Dimensiones', d.dimensiones + ' cm')}
                        ${row('Peso Total', d.peso + ' ' + d.unidad)}
                        ${row('Fecha Levantamiento', d.fecha)}
                        ${row('Facilidad Encontrar', d.facilidad)}
                        ${row('Criticidad', d.criticidad)}
                        ${row('Categoría/Familia', d.categoria)}
                        ${row('Planta Asociada', d.planta)}
                        ${row('Equipo Asociado', d.equipo)}
                        ${row('Función', d.funcion)}
                        ${row('Sustitutos', d.sustitutos)}
                        ${row('Estándar / Medida', d.estandar)}
                        ${row('Vida Útil', d.vida_util)}
                        ${row('Almacenamiento', d.almacenamiento)}
                    </table>
                    <div style="margin-top:15px; padding:10px; background:#fefce8; border-radius:5px; font-size:0.8rem;">
                        <b>OBSERVACIONES:</b><br>${d.observaciones || 'Sin observaciones adicionales.'}
                    </div>
                </div>
            `;
            modal.style.display = 'flex';
        }

        window.onload = initUI;
    </script>
</body>
</html>"""

    full_html = html_template.replace("__DB_JSON_DATA__", json.dumps(db_json))
    with open(OUTPUT_HTML, "w", encoding="utf-8") as f: f.write(full_html)

if __name__ == "__main__":
    main()
