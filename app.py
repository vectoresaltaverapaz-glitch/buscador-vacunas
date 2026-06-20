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
logging.basicConfig(level=logging.DEBUG)
logger = logging.getLogger(__name__)

# ==================================================
# APP
# ==================================================

app = Flask(__name__)
auth = HTTPBasicAuth()

# ==================================================
# USUARIOS
# ==================================================

USUARIOS = {
    "admin": "vacunas2025",
    "consulta": "buscar123"
}

@auth.verify_password
def verify_password(usuario, password):
    return usuario in USUARIOS and USUARIOS[usuario] == password

# ==================================================
# TURSO - CONEXIÓN CON libsql-client (VERSIÓN HTTP)
# ==================================================

def get_turso_client():
    url = os.getenv("TURSO_URL")
    token = os.getenv("TURSO_TOKEN")
    
    if not url or not token:
        logger.error("❌ Variables de entorno TURSO_URL o TURSO_TOKEN no están definidas")
        raise Exception("Faltan variables de entorno TURSO_URL o TURSO_TOKEN")
    
    logger.info(f"✅ Conectando a Turso: {url}")
    return libsql_client.create_client_sync(url=url, auth_token=token)

# ==================================================
# RENOMBRES (se mantiene igual)
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
    "código_único_de_identificación": "CUI Niño",
    "nombre_de_la_niña_o_del_niño": "Nombre Niño",
    "cui": "CUI Responsable",
    "nombre_de_la_madre_padre_o_responsable": "Nombre Responsable",
    "teléfono": "Teléfono",
    "falleció": "Falleció",
    "hep._b": "Hepatitis B",
    "bcg": "BCG",
    "1a._(fipv)": "Polio 1",
    "2a._(fipv)": "Polio 2",
    "1a._(ipv)": "Polio 3",
    "1a._(historico)": "Polio Histórico",
    "2a._(opv)": "OPV 2",
    "3a._(opv)": "OPV 3",
    "1a.": "Pentavalente 1",
    "2a.": "Pentavalente 2",
    "3a.": "Pentavalente 3",
    "1a..1": "Rotavirus 1",
    "2a..1": "Rotavirus 2",
    "1a..2": "Neumococo 1",
    "2a..2": "Neumococo 2",
    "spr_1": "SPR 1",
    "neumo-_r1": "Neumo R1",
    "spr_2": "SPR 2",
    "r1_(opv)": "Refuerzo 1 OPV",
    "r1_(dpt)": "Refuerzo 1 DPT",
    "r2_(opv)": "Refuerzo 2 OPV",
    "r2_(dpt)": "Refuerzo 2 DPT",
    "1a..3": "Vitamina A 1",
    "2a..3": "Vitamina A 2",
    "1a..4": "SRP 1",
    "2a..4": "SRP 2",
    "1a..5": "TD 1",
    "2a..5": "TD 2",
    "1a._(fipv).1": "Polio 1 Refuerzo",
    "2a._(fipv).1": "Polio 2 Refuerzo",
    "1a._(ipv).1": "Polio IPV Refuerzo",
    "1a._(historico).1": "Histórico Refuerzo",
    "2a._(opv).1": "OPV 2 Refuerzo",
    "3a._(opv).1": "OPV 3 Refuerzo",
    "r1_(opv).1": "Refuerzo OPV 1",
    "r2_(opv).1": "Refuerzo OPV 2",
    "1a..6": "Esquema 1",
    "2a..6": "Esquema 2",
    "3a..1": "Esquema 3",
    "r1": "Refuerzo 1",
    "r2": "Refuerzo 2",
    "spr_1.1": "SPR Refuerzo 1",
    "spr_2.1": "SPR Refuerzo 2",
    "1a..7": "Dosis 1",
    "2a..7": "Dosis 2",
    "3a..": "Dosis 3"
}

# ==================================================
# COLUMNAS FECHA
# ==================================================

COLUMNAS_FECHA = {
    "hep._b", "bcg", "1a._(fipv)", "2a._(fipv)", "1a._(ipv)",
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
    "1a..7", "2a..7", "3a.."
}

# ==================================================
# FUNCIONES
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
# HTML (se mantiene igual)
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
<option value="">TODOS LOS DISTRITOS</option>
<option value="COBAN">COBAN</option>
<option value="SAN PEDRO CARCHA">SAN PEDRO CARCHA</option>
<option value="TACTIC">TACTIC</option>
<option value="LANQUIN">LANQUIN</option>
<option value="CHISEC">CHISEC</option>
<option value="CAHABON">CAHABON</option>
<option value="SENAHU">SENAHU</option>
<option value="FRAY BARTOLOME">FRAY BARTOLOME</option>
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
    .then(r=>r.json())
    .then(data=>{
        let html = '';
        html += `<p>Total resultados: ${data.total}</p>`;
        html += '<div style="overflow:auto; max-height:700px;">';
        html += '<table><tr>';
        data.columnas.forEach(c=>{ html += `<th>${c}</th>`; });
        html += '</tr>';
        data.rows.forEach(r=>{
            html += '<tr>';
            r.forEach(c=>{ html += `<td>${c ?? ''}</td>`; });
            html += '</tr>';
        });
        html += '</table></div>';
        document.getElementById('resultado').innerHTML = html;
    })
    .catch(err => {
        document.getElementById('resultado').innerHTML = '<p style="color:red;">❌ Error al buscar: ' + err + '</p>';
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
    return render_template_string(HTML)

# ==================================================
# BUSQUEDA - CORREGIDA CON API DE libsql-client
# ==================================================

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
            "nombre_nino": "nombre_de_la_niña_o_del_niño",
            "nombre_responsable": "nombre_de_la_madre_padre_o_responsable",
            "cui_nino": "código_único_de_identificación",
            "cui_responsable": "cui"
        }
        columna = mapa[tipo]
        logger.info(f"📌 Columna mapeada: '{columna}'")

        condiciones = []
        parametros = []

        if query:
            if "cui" in tipo:
                condiciones.append(f'"{columna}" = ?')
                parametros.append(query)
            else:
                condiciones.append(f'"{columna}" LIKE ?')
                parametros.append(f"%{query}%")
        if distrito:
            condiciones.append('UPPER("distrito") LIKE ?')
            parametros.append(f"%{distrito.upper()}%")

        where = " AND ".join(condiciones) if condiciones else ""
        sql = f"""
        SELECT rowid,* FROM datos_completos
        {f'WHERE {where}' if where else ''}
        LIMIT 300
        """
        logger.info(f"📝 SQL: {sql}")
        logger.info(f"📦 Parámetros: {parametros}")

        # CORRECCIÓN: Obtener filas correctamente
        result = client.execute(sql, parametros)
        
        # Verificar método de obtención de filas
        if hasattr(result, 'rows'):
            rows = result.rows()
        else:
            # Fallback si no tiene rows()
            rows = result.fetchall() if hasattr(result, 'fetchall') else []
        
        logger.info(f"📊 Filas obtenidas: {len(rows)}")
        
        # Obtener nombres de columna
        if hasattr(result, 'columns'):
            col_names = result.columns()
        else:
            # Fallback con description
            col_names = [desc[0] for desc in result.description] if hasattr(result, 'description') else []
        
        logger.info(f"📋 Columnas: {col_names}")

        rows_dict = [dict(zip(col_names, row)) for row in rows]

        client.close()

        if not rows_dict:
            logger.warning("⚠️ No se encontraron resultados")
            return jsonify({"rows": [], "columnas": [], "total": 0})

        columnas = list(rows_dict[0].keys())
        columnas_finales = [c for c in columnas if c not in ["día", "mes", "año_1", "hombre", "mujer"]]
        titulos_columnas = [RENOMBRES.get(c, c) for c in columnas_finales]

        resultados = []
        for row in rows_dict:
            fila = [procesar_valor(c, row.get(c)) for c in columnas_finales]
            resultados.append(fila)

        logger.info(f"✅ Resultados devueltos: {len(resultados)}")
        return jsonify({"rows": resultados, "columnas": titulos_columnas, "total": len(resultados)})

    except Exception as e:
        logger.error(f"❌ Error en /buscar: {str(e)}", exc_info=True)
        return jsonify({"error": str(e)}), 500

# ==================================================
# EXPORTAR - CORREGIDA CON API DE libsql-client
# ==================================================

@app.route('/exportar')
@auth.login_required
def exportar():
    query = request.args.get('q', '').strip()
    tipo = request.args.get('tipo', 'nombre_nino')
    distrito = request.args.get('distrito', '').strip()
    
    logger.info(f"📤 Exportando: query='{query}', tipo='{tipo}', distrito='{distrito}'")

    try:
        client = get_turso_client()

        mapa = {
            "nombre_nino": "nombre_de_la_niña_o_del_niño",
            "nombre_responsable": "nombre_de_la_madre_padre_o_responsable",
            "cui_nino": "código_único_de_identificación",
            "cui_responsable": "cui"
        }
        columna = mapa[tipo]

        condiciones = []
        parametros = []

        if query:
            if "cui" in tipo:
                condiciones.append(f'"{columna}" = ?')
                parametros.append(query)
            else:
                condiciones.append(f'"{columna}" LIKE ?')
                parametros.append(f"%{query}%")
        if distrito:
            condiciones.append('UPPER("distrito") LIKE ?')
            parametros.append(f"%{distrito.upper()}%")

        where = " AND ".join(condiciones) if condiciones else ""
        sql = f"SELECT rowid,* FROM datos_completos {f'WHERE {where}' if where else ''}"
        logger.info(f"📝 SQL export: {sql}")

        result = client.execute(sql, parametros)
        
        # Obtener filas correctamente
        if hasattr(result, 'rows'):
            rows = result.rows()
        else:
            rows = result.fetchall() if hasattr(result, 'fetchall') else []
        
        if rows:
            if hasattr(result, 'columns'):
                col_names = result.columns()
            else:
                col_names = [desc[0] for desc in result.description] if hasattr(result, 'description') else []
        else:
            col_names = []
        
        rows_dict = [dict(zip(col_names, row)) for row in rows]

        wb = Workbook()
        ws = wb.active
        ws.title = "Resultados"

        if rows_dict:
            columnas = list(rows_dict[0].keys())
            columnas_finales = [c for c in columnas if c not in ["día", "mes", "año_1", "hombre", "mujer"]]
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

# ==================================================
# INICIO
# ==================================================

if __name__ == '__main__':
    logger.info("🚀 Servidor iniciado...")
    app.run(host='0.0.0.0', port=5000, debug=True, threaded=True)