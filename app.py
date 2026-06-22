import logging
import os
from flask import Flask, render_template_string, request, jsonify, Response
from flask_httpauth import HTTPBasicAuth
from datetime import datetime, timedelta
from io import BytesIO
import libsql_client
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment

# Configurar logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

app = Flask(__name__)
auth = HTTPBasicAuth()

USUARIOS = {
    "admin": "vacunas2025",
    "consulta": "buscar123"
}

@auth.verify_password
def verify_password(usuario, password):
    return usuario in USUARIOS and USUARIOS[usuario] == password

def get_turso_client():
    url = os.getenv("TURSO_URL")
    token = os.getenv("TURSO_TOKEN")
    if not url or not token:
        logger.error("❌ Faltan variables de entorno TURSO_URL o TURSO_TOKEN")
        raise Exception("Faltan variables de entorno")
    logger.info(f"✅ Conectando a Turso: {url}")
    return libsql_client.create_client_sync(url=url, auth_token=token)

def get_distritos(client):
    """Obtiene lista única de distritos ordenada alfabéticamente"""
    try:
        result = client.execute("SELECT DISTINCT distrito FROM datos_completos WHERE distrito IS NOT NULL AND distrito != '' ORDER BY distrito")
        if hasattr(result, 'rows') and callable(result.rows):
            rows = result.rows()
        else:
            rows = list(result)
        return [row[0] for row in rows]
    except Exception as e:
        logger.error(f"❌ Error al obtener distritos: {str(e)}")
        return []

# ==================================================
# RENOMBRES (nombre actual en BD → título bonito)
# ==================================================

RENOMBRES = {
    "rowid": "ID",
    "año": "Año",
    "no.": "No",
    "área": "Área Salud",
    "distrito": "Distrito Salud",
    "servicio": "Servicio Salud",
    "departamento": "Departamento",
    "municipio": "Municipio",
    "cui_nino": "CUI Niño",
    "nombre_nino": "Nombre Niño",
    "cui_responsable": "CUI Responsable",
    "nombre_responsable": "Nombre Responsable",
    "telefono": "Teléfono",
    "falleció": "Falleció",
    # Vacunas
    "hep._b": "Hepatitis B (<1 año)",
    "bcg": "BCG (<1 año)",
    "1a._(fipv)": "Polio 1 (<1 año)",
    "2a._(fipv)": "Polio 2 (<1 año)",
    "1a._(ipv)": "Polio IPV (<1 año)",
    "1a._(historico)": "Polio Histórico (<1 año)",
    "2a._(opv)": "OPV 2 (<1 año)",
    "3a._(opv)": "OPV 3 (<1 año)",
    "1a.": "Pentavalente 1 (<1 año)",
    "2a.": "Pentavalente 2 (<1 año)",
    "3a.": "Pentavalente 3 (<1 año)",
    "1a..1": "Rotavirus 1 (<1 año)",
    "2a..1": "Rotavirus 2 (<1 año)",
    "1a..2": "Neumococo 1 (<1 año)",
    "2a..2": "Neumococo 2 (<1 año)",
    "spr_1": "SPR 1 (12 meses)",
    "neumo-_r1": "Neumococo R1 (12 meses)",
    "spr_2": "SPR 2 (18 meses)",
    "r1_(opv)": "Refuerzo OPV (18 meses)",
    "r1_(dpt)": "Refuerzo DPT (18 meses)",
    "r2_(opv)": "Refuerzo OPV (4-7 años)",
    "r2_(dpt)": "Refuerzo DPT (4-7 años)",
    "1a..3": "Influenza 1 (6-11 meses)",
    "2a..3": "Influenza 2 (6-11 meses)",
    "1a..4": "Influenza 1 (12-23 meses)",
    "2a..4": "Influenza 2 (12-23 meses)",
    "1a..5": "Influenza 1 (24-35 meses)",
    "2a..5": "Influenza 2 (24-35 meses)",
    "1a._(fipv).1": "Polio fIPV 1 (1-7 años)",
    "2a._(fipv).1": "Polio fIPV 2 (1-7 años)",
    "1a._(ipv).1": "Polio IPV Refuerzo (1-7 años)",
    "1a._(historico).1": "Polio Histórico Refuerzo (1-7 años)",
    "2a._(opv).1": "OPV 2 Refuerzo (1-7 años)",
    "3a._(opv).1": "OPV 3 Refuerzo (1-7 años)",
    "r1_(opv).1": "Refuerzo OPV 1 (1-7 años)",
    "r2_(opv).1": "Refuerzo OPV 2 (1-7 años)",
    "1a..6": "Pentavalente Ref 1 (1-7 años)",
    "2a..6": "Pentavalente Ref 2 (1-7 años)",
    "3a..1": "Pentavalente Ref 3 (1-7 años)",
    "r1": "Refuerzo 1 (1-7 años)",
    "r2": "Refuerzo 2 (1-7 años)",
    "spr_1.1": "SPR Refuerzo 1 (1-7 años)",
    "spr_2.1": "SPR Refuerzo 2 (1-7 años)",
    "1a..7": "Otras Vacunas 1 (1-7 años)",
    "2a..7": "Otras Vacunas 2 (1-7 años)",
    "3a..2": "Otras Vacunas 3 (1-7 años)"
}

# ==================================================
# COLUMNAS FECHA (nombres actuales en BD)
# ==================================================

COLUMNAS_FECHA = {
    "hep._b", "bcg",
    "1a._(fipv)", "2a._(fipv)", "1a._(ipv)",
    "1a._(historico)", "2a._(opv)", "3a._(opv)",
    "1a.", "2a.", "3a.",
    "1a..1", "2a..1",
    "1a..2", "2a..2",
    "spr_1", "neumo-_r1", "spr_2",
    "r1_(opv)", "r1_(dpt)",
    "r2_(opv)", "r2_(dpt)",
    "1a..3", "2a..3",
    "1a..4", "2a..4",
    "1a..5", "2a..5",
    "1a._(fipv).1", "2a._(fipv).1", "1a._(ipv).1",
    "1a._(historico).1",
    "2a._(opv).1", "3a._(opv).1",
    "r1_(opv).1", "r2_(opv).1",
    "1a..6", "2a..6", "3a..1",
    "r1", "r2",
    "spr_1.1", "spr_2.1",
    "1a..7", "2a..7", "3a..2"
}

# ==================================================
# FUNCIONES DE PROCESAMIENTO
# ==================================================

def limpiar_numero(valor):
    try:
        if valor is None:
            return ""
        numero = float(valor)
        if numero.is_integer():
            return str(int(numero))
        return str(numero)
    except:
        return str(valor)

def convertir_fecha_excel(valor):
    try:
        if valor is None or valor == "":
            return ""
        numero = float(valor)
        if 20000 <= numero <= 60000:
            fecha_base = datetime(1899, 12, 30)
            fecha = fecha_base + timedelta(days=int(numero))
            return fecha.strftime("%d/%m/%Y")
        return limpiar_numero(valor)
    except:
        return str(valor)

def procesar_valor(columna, valor):
    if columna in COLUMNAS_FECHA:
        return convertir_fecha_excel(valor)
    return limpiar_numero(valor)

# ==================================================
# COLUMNAS REALES (lista fija para evitar PRAGMA)
# ==================================================

COLUMNAS_REALES = [
    "rowid", "año", "no.", "área", "distrito", "servicio",
    "departamento", "municipio", "cui_nino", "nombre_nino",
    "día", "mes", "año_1", "departamento.1", "municipio.1",
    "comunidad", "hombre", "mujer", "pueblo", "comunidad.1",
    "cui_responsable", "nombre_responsable", "día.1", "mes.1",
    "año.1", "departamento.2", "municipio.2", "comunidad.2",
    "calle,_avenida,_zona,_lote,", "telefono", "falleció",
    "hep._b", "bcg", "1a._(fipv)", "2a._(fipv)", "1a._(ipv)",
    "1a._(historico)", "2a._(opv)", "3a._(opv)", "1a.", "2a.",
    "3a.", "1a..1", "2a..1", "1a..2", "2a..2", "spr_1",
    "neumo-_r1", "spr_2", "r1_(opv)", "r1_(dpt)", "r2_(opv)",
    "r2_(dpt)", "1a..3", "2a..3", "1a..4", "2a..4", "1a..5",
    "2a..5", "1a._(fipv).1", "2a._(fipv).1", "1a._(ipv).1",
    "1a._(historico).1", "2a._(opv).1", "3a._(opv).1",
    "r1_(opv).1", "r2_(opv).1", "1a..6", "2a..6", "3a..1",
    "r1", "r2", "spr_1.1", "spr_2.1", "1a..7", "2a..7", "3a..2"
]

# ==================================================
# HTML (con marcador para distritos)
# ==================================================

HTML = """
<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<title>Buscador SIGSA</title>
<style>
body{ font-family:Arial; background:#eef2f7; padding:20px; }
.container{ background:white; padding:20px; border-radius:10px; }
input,select,button{ padding:10px; margin:5px; }
table{ width:100%; border-collapse:collapse; margin-top:20px; font-size:12px; }
th,td{ border:1px solid #ccc; padding:5px; }
th{ background:#1e466e; color:white; position:sticky; top:0; }
</style>
</head>
<body>
<div class="container">
<h2>🔍 Buscador SIGSA</h2>
<select id="distrito">
<!-- DISTRITOS_AQUI -->
</select>
<select id="tipo">
<option value="nombre_nino">Nombre Niño</option>
<option value="nombre_responsable">Nombre Responsable</option>
<option value="cui_nino">CUI Niño</option>
<option value="cui_responsable">CUI Responsable</option>
</select>
<input type="text" id="search">
<button onclick="buscar()">Buscar</button>
<button onclick="exportarExcel()">Exportar Excel</button>
<div id="resultado"></div>
</div>
<script>
function buscar(){
    let q = document.getElementById('search').value;
    let tipo = document.getElementById('tipo').value;
    let distrito = document.getElementById('distrito').value;
    fetch(`/buscar?q=${encodeURIComponent(q)}&tipo=${tipo}&distrito=${encodeURIComponent(distrito)}`)
    .then(r => r.json())
    .then(data => {
        if (data.error) {
            document.getElementById('resultado').innerHTML = `<p style="color:red;">❌ Error del servidor:</p><pre style="background:#fdd;padding:10px;border-radius:5px;white-space:pre-wrap;">${data.error}</pre>`;
            return;
        }
        if (!data.columnas || !data.rows) {
            document.getElementById('resultado').innerHTML = '<p>⚠️ No se encontraron resultados o la respuesta es inválida.</p>';
            return;
        }
        let html = '';
        html += `<p>Total resultados: ${data.total}</p>`;
        html += '<div style="overflow:auto; max-height:700px;">';
        html += '<table><tr>';
        data.columnas.forEach(c => { html += `<th>${c}</th>`; });
        html += '</tr>';
        data.rows.forEach(r => {
            html += '<tr>';
            r.forEach(c => { html += `<td>${c ?? ''}</td>`; });
            html += '</tr>';
        });
        html += '</table></div>';
        document.getElementById('resultado').innerHTML = html;
    })
    .catch(err => {
        document.getElementById('resultado').innerHTML = `<p style="color:red;">❌ Error de red o servidor: ${err}</p>`;
    });
}
function exportarExcel(){
    let q = document.getElementById('search').value;
    let tipo = document.getElementById('tipo').value;
    let distrito = document.getElementById('distrito').value;
    window.location.href = `/exportar?q=${encodeURIComponent(q)}&tipo=${tipo}&distrito=${encodeURIComponent(distrito)}`;
}
</script>
</body>
</html>
"""

# ==================================================
# RUTAS
# ==================================================

@app.route('/')
@auth.login_required
def index():
    try:
        client = get_turso_client()
        distritos = get_distritos(client)
        client.close()
        
        # Generar opciones HTML
        opciones = '<option value="">TODOS LOS DISTRITOS</option>'
        for d in distritos:
            opciones += f'<option value="{d}">{d}</option>'
        
        # Reemplazar en HTML
        html_con_distritos = HTML.replace('<!-- DISTRITOS_AQUI -->', opciones)
        return render_template_string(html_con_distritos)
    except Exception as e:
        logger.error(f"❌ Error al cargar distritos: {str(e)}")
        # Fallback: usar lista manual simple
        html_fallback = HTML.replace('<!-- DISTRITOS_AQUI -->', '<option value="">TODOS LOS DISTRITOS</option>')
        return render_template_string(html_fallback)

@app.route('/buscar')
@auth.login_required
def buscar():
    query = request.args.get('q', '').strip()
    tipo = request.args.get('tipo', 'nombre_nino')
    distrito = request.args.get('distrito', '').strip()
    
    logger.info(f"🔍 Búsqueda: query='{query}', tipo='{tipo}', distrito='{distrito}'")
    
    try:
        client = get_turso_client()
        
        mapa = {
            "nombre_nino": "nombre_nino",
            "nombre_responsable": "nombre_responsable",
            "cui_nino": "cui_nino",
            "cui_responsable": "cui_responsable"
        }
        columna_busqueda = mapa.get(tipo, "nombre_nino")
        logger.info(f"📌 Columna de búsqueda: '{columna_busqueda}'")
        
        condiciones = []
        parametros = []
        if query:
            if "cui" in tipo:
                condiciones.append(f'"{columna_busqueda}" = ?')
                parametros.append(query)
            else:
                condiciones.append(f'"{columna_busqueda}" LIKE ?')
                parametros.append(f"%{query}%")
        if distrito:
            condiciones.append('"distrito" = ?')
            parametros.append(distrito)
        
        where = " AND ".join(condiciones) if condiciones else ""
        sql = f"""
        SELECT * FROM datos_completos
        {f'WHERE {where}' if where else ''}
        LIMIT 30
        """
        logger.info(f"📝 SQL: {sql}")
        logger.info(f"📦 Parámetros: {parametros}")
        
        result = client.execute(sql, parametros)
        if hasattr(result, 'rows') and callable(result.rows):
            rows = result.rows()
        else:
            rows = list(result)
        
        logger.info(f"📊 Filas obtenidas: {len(rows)}")
        
        if not rows:
            client.close()
            return jsonify({"rows": [], "columnas": [], "total": 0})
        
        rows_dict = [dict(zip(COLUMNAS_REALES, row)) for row in rows]
        
        # Filtrar columnas que no se quieren mostrar
        columnas_excluir = ["día", "mes", "año_1", "hombre", "mujer"]
        columnas_finales = [c for c in COLUMNAS_REALES if c not in columnas_excluir]
        titulos_columnas = [RENOMBRES.get(c, c) for c in columnas_finales]
        
        resultados = []
        for row in rows_dict:
            fila = [procesar_valor(c, row.get(c)) for c in columnas_finales]
            resultados.append(fila)
        
        client.close()
        logger.info(f"✅ Resultados devueltos: {len(resultados)}")
        return jsonify({
            "rows": resultados,
            "columnas": titulos_columnas,
            "total": len(resultados)
        })
    
    except Exception as e:
        logger.error(f"❌ Error en /buscar: {str(e)}", exc_info=True)
        return jsonify({
            "rows": [],
            "columnas": [],
            "total": 0,
            "error": str(e)
        })

MAX_EXPORT_ROWS = 1000

@app.route('/exportar')
@auth.login_required
def exportar():
    query = request.args.get('q', '').strip()
    tipo = request.args.get('tipo', 'nombre_nino')
    distrito = request.args.get('distrito', '').strip()
    
    try:
        client = get_turso_client()
        
        mapa = {
            "nombre_nino": "nombre_nino",
            "nombre_responsable": "nombre_responsable",
            "cui_nino": "cui_nino",
            "cui_responsable": "cui_responsable"
        }
        columna_busqueda = mapa.get(tipo, "nombre_nino")
        
        condiciones = []
        parametros = []
        if query:
            if "cui" in tipo:
                condiciones.append(f'"{columna_busqueda}" = ?')
                parametros.append(query)
            else:
                condiciones.append(f'"{columna_busqueda}" LIKE ?')
                parametros.append(f"%{query}%")
        if distrito:
            condiciones.append('"distrito" = ?')
            parametros.append(distrito)
        
        where = " AND ".join(condiciones) if condiciones else ""
        sql = f"""
        SELECT * FROM datos_completos
        {f'WHERE {where}' if where else ''}
        LIMIT {MAX_EXPORT_ROWS}
        """
        logger.info(f"📝 SQL export: {sql}")
        
        result = client.execute(sql, parametros)
        if hasattr(result, 'rows') and callable(result.rows):
            rows = result.rows()
        else:
            rows = list(result)
        
        if not rows:
            client.close()
            wb = Workbook()
            ws = wb.active
            ws.title = "Resultados"
            ws.append(["No se encontraron registros"])
            output = BytesIO()
            wb.save(output)
            output.seek(0)
            return Response(
                output.getvalue(),
                mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                headers={'Content-Disposition': 'attachment; filename=resultados.xlsx'}
            )
        
        rows_dict = [dict(zip(COLUMNAS_REALES, row)) for row in rows]
        
        wb = Workbook()
        ws = wb.active
        ws.title = "Resultados"
        
        columnas_excluir = ["día", "mes", "año_1", "hombre", "mujer"]
        columnas_finales = [c for c in COLUMNAS_REALES if c not in columnas_excluir]
        titulos_columnas = [RENOMBRES.get(c, c) for c in columnas_finales]
        ws.append(titulos_columnas)
        for cell in ws[1]:
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = PatternFill(start_color="1e466e", end_color="1e466e", fill_type="solid")
            cell.alignment = Alignment(horizontal="center")
        
        for row in rows_dict:
            fila = [procesar_valor(c, row.get(c)) for c in columnas_finales]
            ws.append(fila)
        
        output = BytesIO()
        wb.save(output)
        output.seek(0)
        client.close()
        
        return Response(
            output.getvalue(),
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            headers={'Content-Disposition': 'attachment; filename=resultados.xlsx'}
        )
    
    except Exception as e:
        logger.error(f"❌ Error en /exportar: {str(e)}", exc_info=True)
        return jsonify({"error": str(e)}), 500

if __name__ == '__main__':
    logger.info("🚀 Servidor iniciado...")
    app.run(host='0.0.0.0', port=5000, debug=True, threaded=True)