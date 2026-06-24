import logging
import os
from flask import Flask, render_template_string, request, jsonify, Response
from flask_httpauth import HTTPBasicAuth
from datetime import datetime, timedelta
from io import BytesIO
import libsql_client
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment

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
        logger.error("❌ Faltan variables de entorno")
        raise Exception("Faltan variables de entorno")
    logger.info(f"✅ Conectando a Turso: {url}")
    # Forzar HTTPS para evitar problemas de WebSocket en Render
    if url.startswith("libsql://"):
        url = url.replace("libsql://", "https://")
    return libsql_client.create_client_sync(url=url, auth_token=token)

# ==================================================
# HTML (simplificado, puedes usar el mismo de antes)
# ==================================================

HTML = """
<!DOCTYPE html>
<html>
<head><title>Buscador de Vacunas</title>
<style>
body{font-family:Arial;background:#eef2f7;padding:20px;}
.container{background:white;padding:20px;border-radius:10px;}
input,select,button{padding:10px;margin:5px;}
table{width:100%;border-collapse:collapse;margin-top:20px;font-size:12px;}
th,td{border:1px solid #ccc;padding:5px;}
th{background:#1e466e;color:white;position:sticky;top:0;}
</style>
</head>
<body>
<div class="container">
<h2>🔍 Buscador de Vacunas</h2>
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
<option value="PANZÓS">PANZÓS</option>
<option value="SAN JUAN CHAMELCO">SAN JUAN CHAMELCO</option>
<option value="TUCURÚ">TUCURÚ</option>
<option value="TAMAHÚ">TAMAHÚ</option>
<option value="SANTA CRUZ VERAPAZ">SANTA CRUZ VERAPAZ</option>
<option value="CAMPUR">CAMPUR</option>
<option value="COBÁN">COBÁN</option>
<option value="RAXRUHA">RAXRUHA</option>
<option value="CHAHAL">CHAHAL</option>
<option value="LA TINTA">LA TINTA</option>
<option value="TELEMÁN">TELEMÁN</option>
</select>
<select id="tipo">
<option value="nombre_nino">Nombre Niño</option>
<option value="nombre_responsable">Nombre Responsable</option>
<option value="cui_nino">CUI Niño</option>
<option value="cui_responsable">CUI Responsable</option>
</select>
<select id="genero">
<option value="">Género (todos)</option>
<option value="Hombre">Hombre</option>
<option value="Mujer">Mujer</option>
<option value="No especificado">No especificado</option>
</select>
<input type="text" id="search" placeholder="Buscar..." style="width:200px;">
<div style="margin-top:10px;">
<label>Fecha Nacimiento Niño:</label>
<input type="number" id="dia_nac" placeholder="Día" style="width:60px;">
<input type="number" id="mes_nac" placeholder="Mes" style="width:60px;">
<input type="number" id="ano_nac" placeholder="Año" style="width:80px;">
<label style="margin-left:15px;">Fecha Nacimiento Responsable:</label>
<input type="number" id="dia_madre" placeholder="Día" style="width:60px;">
<input type="number" id="mes_madre" placeholder="Mes" style="width:60px;">
<input type="number" id="ano_madre" placeholder="Año" style="width:80px;">
</div>
<div style="margin-top:10px;">
<button onclick="buscar()">🔍 Buscar</button>
<button onclick="exportarExcel()">📊 Exportar Excel</button>
<button onclick="limpiar()" style="background:#e74c3c;color:white;border:none;">🔄 Nueva búsqueda</button>
<span style="margin-left:15px;font-size:14px;color:#555;">Mostrando hasta 30 registros</span>
</div>
<div id="resultado"></div>
</div>
<script>
function buscar(){
    let q = document.getElementById('search').value;
    let tipo = document.getElementById('tipo').value;
    let distrito = document.getElementById('distrito').value;
    let genero = document.getElementById('genero').value;
    let dia_nac = document.getElementById('dia_nac').value;
    let mes_nac = document.getElementById('mes_nac').value;
    let ano_nac = document.getElementById('ano_nac').value;
    let dia_madre = document.getElementById('dia_madre').value;
    let mes_madre = document.getElementById('mes_madre').value;
    let ano_madre = document.getElementById('ano_madre').value;
    let url = `/buscar?q=${encodeURIComponent(q)}&tipo=${tipo}&distrito=${encodeURIComponent(distrito)}&genero=${encodeURIComponent(genero)}`;
    if(dia_nac) url+=`&dia_nac=${dia_nac}`;
    if(mes_nac) url+=`&mes_nac=${mes_nac}`;
    if(ano_nac) url+=`&ano_nac=${ano_nac}`;
    if(dia_madre) url+=`&dia_madre=${dia_madre}`;
    if(mes_madre) url+=`&mes_madre=${mes_madre}`;
    if(ano_madre) url+=`&ano_madre=${ano_madre}`;
    fetch(url).then(r=>r.json()).then(data=>{
        if(data.error){ document.getElementById('resultado').innerHTML = `<p style="color:red;">❌ Error: ${data.error}</p>`; return; }
        if(!data.columnas || !data.rows){ document.getElementById('resultado').innerHTML = '<p>⚠️ No se encontraron resultados.</p>'; return; }
        let html = `<p>Total resultados: ${data.total}</p><div style="overflow:auto;max-height:700px;"><table><tr>`;
        data.columnas.forEach(c=> html+=`<th>${c}</th>`);
        html+='</tr>';
        data.rows.forEach(r=>{ html+='<tr>'; r.forEach(c=> html+=`<td>${c??''}</td>`); html+='</tr>'; });
        html+='</table></div>';
        document.getElementById('resultado').innerHTML = html;
    }).catch(err=>{ document.getElementById('resultado').innerHTML = `<p style="color:red;">❌ Error: ${err}</p>`; });
}
function limpiar(){
    document.getElementById('search').value='';
    document.getElementById('tipo').value='nombre_nino';
    document.getElementById('distrito').value='';
    document.getElementById('genero').value='';
    document.getElementById('dia_nac').value='';
    document.getElementById('mes_nac').value='';
    document.getElementById('ano_nac').value='';
    document.getElementById('dia_madre').value='';
    document.getElementById('mes_madre').value='';
    document.getElementById('ano_madre').value='';
    document.getElementById('resultado').innerHTML='';
}
function exportarExcel(){
    let q = document.getElementById('search').value;
    let tipo = document.getElementById('tipo').value;
    let distrito = document.getElementById('distrito').value;
    let genero = document.getElementById('genero').value;
    let dia_nac = document.getElementById('dia_nac').value;
    let mes_nac = document.getElementById('mes_nac').value;
    let ano_nac = document.getElementById('ano_nac').value;
    let dia_madre = document.getElementById('dia_madre').value;
    let mes_madre = document.getElementById('mes_madre').value;
    let ano_madre = document.getElementById('ano_madre').value;
    let url = `/exportar?q=${encodeURIComponent(q)}&tipo=${tipo}&distrito=${encodeURIComponent(distrito)}&genero=${encodeURIComponent(genero)}`;
    if(dia_nac) url+=`&dia_nac=${dia_nac}`;
    if(mes_nac) url+=`&mes_nac=${mes_nac}`;
    if(ano_nac) url+=`&ano_nac=${ano_nac}`;
    if(dia_madre) url+=`&dia_madre=${dia_madre}`;
    if(mes_madre) url+=`&mes_madre=${mes_madre}`;
    if(ano_madre) url+=`&ano_madre=${ano_madre}`;
    window.location.href = url;
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

@app.route('/buscar')
@auth.login_required
def buscar():
    query = request.args.get('q', '').strip()
    tipo = request.args.get('tipo', 'nombre_nino')
    distrito = request.args.get('distrito', '').strip()
    genero = request.args.get('genero', '').strip()
    dia_nac = request.args.get('dia_nac', '').strip()
    mes_nac = request.args.get('mes_nac', '').strip()
    ano_nac = request.args.get('ano_nac', '').strip()
    dia_madre = request.args.get('dia_madre', '').strip()
    mes_madre = request.args.get('mes_madre', '').strip()
    ano_madre = request.args.get('ano_madre', '').strip()
    
    try:
        client = get_turso_client()
        
        # Construir condiciones
        condiciones = []
        parametros = []
        # Mapeo de columnas según tipo
        col_map = {
            "nombre_nino": "nombre_nino",
            "nombre_responsable": "nombre_responsable",
            "cui_nino": "cui_nino",
            "cui_responsable": "cui_responsable"
        }
        col = col_map.get(tipo, "nombre_nino")
        
        if query:
            if tipo in ["cui_nino", "cui_responsable"]:
                condiciones.append(f'"{col}" = ?')
                parametros.append(query)
            else:
                condiciones.append(f'"{col}" LIKE ?')
                parametros.append(f"%{query}%")
        if distrito:
            condiciones.append('UPPER("distrito") LIKE ?')
            parametros.append(f"%{distrito.upper()}%")
        if genero:
            if genero == "Hombre":
                condiciones.append('"hombre" = "X"')
            elif genero == "Mujer":
                condiciones.append('"mujer" = "X"')
            elif genero == "No especificado":
                condiciones.append('"hombre" != "X" AND "mujer" != "X"')
        if dia_nac:
            condiciones.append('"día" = ?')
            parametros.append(dia_nac)
        if mes_nac:
            condiciones.append('"mes" = ?')
            parametros.append(mes_nac)
        if ano_nac:
            condiciones.append('"año_1" = ?')
            parametros.append(ano_nac)
        if dia_madre:
            condiciones.append('"día.1" = ?')
            parametros.append(dia_madre)
        if mes_madre:
            condiciones.append('"mes.1" = ?')
            parametros.append(mes_madre)
        if ano_madre:
            condiciones.append('"año.1" = ?')
            parametros.append(ano_madre)
        
        where = " AND ".join(condiciones) if condiciones else ""
        sql = f"SELECT * FROM datos_completos {f'WHERE {where}' if where else ''} LIMIT 30"
        logger.info(f"SQL: {sql}")
        logger.info(f"Params: {parametros}")
        
        result = client.execute(sql, parametros)
        # Extraer filas y columnas
        if hasattr(result, 'rows') and callable(result.rows):
            rows = result.rows()
            columns = result.columns() if hasattr(result, 'columns') and callable(result.columns) else []
        else:
            rows = list(result)
            columns = [f"col_{i}" for i in range(len(rows[0]))] if rows else []
        
        # Convertir a lista de diccionarios
        data = [dict(zip(columns, row)) for row in rows]
        
        # Mapeo de nombres bonitos
        renombres = {
            "rowid": "ID",
            "año": "Año",
            "no.": "No",
            "área": "Área Salud",
            "distrito": "Distrito",
            "servicio": "Servicio Salud",
            "departamento": "Departamento",
            "municipio": "Municipio",
            "cui_nino": "CUI Niño",
            "nombre_nino": "Nombre Niño",
            "cui_responsable": "CUI Responsable",
            "nombre_responsable": "Nombre Responsable",
            "telefono": "Teléfono",
            "falleció": "Falleció",
            "genero": "Género",
        }
        # Columnas a mostrar (excluir columnas internas)
        excluir = ["día", "mes", "año_1", "día.1", "mes.1", "año.1", "hombre", "mujer", "comunidad", "pueblo", "rowid", "no.", "área"]
        columnas_finales = [c for c in columns if c not in excluir]
        # Agregar columna genero calculada
        for row in data:
            row["genero"] = "Hombre" if row.get("hombre") == 'X' else "Mujer" if row.get("mujer") == 'X' else "No especificado"
        columnas_finales.append("genero")
        titulos = [renombres.get(c, c) for c in columnas_finales]
        
        resultados = []
        for row in data:
            fila = []
            for c in columnas_finales:
                val = row.get(c, "")
                # Intentar convertir fechas de Excel si es columna de fecha
                if c in ["hep._b", "bcg", "1a._(fipv)", "2a._(fipv)", "1a._(ipv)", "1a._(historico)", "2a._(opv)", "3a._(opv)", "1a.", "2a.", "3a.", "1a..1", "2a..1", "1a..2", "2a..2", "spr_1", "neumo-_r1", "spr_2", "r1_(opv)", "r1_(dpt)", "r2_(opv)", "r2_(dpt)", "1a..3", "2a..3", "1a..4", "2a..4", "1a..5", "2a..5", "1a._(fipv).1", "2a._(fipv).1", "1a._(ipv).1", "1a._(historico).1", "2a._(opv).1", "3a._(opv).1", "r1_(opv).1", "r2_(opv).1", "1a..6", "2a..6", "3a..1", "r1", "r2", "spr_1.1", "spr_2.1", "1a..7", "2a..7", "3a..2"]:
                    if val and isinstance(val, (int, float)) and val > 20000:
                        try:
                            fecha = datetime(1899, 12, 30) + timedelta(days=val)
                            val = fecha.strftime("%d/%m/%Y")
                        except:
                            pass
                fila.append(val)
            resultados.append(fila)
        
        client.close()
        return jsonify({"rows": resultados, "columnas": titulos, "total": len(resultados)})
    except Exception as e:
        logger.error(f"Error: {e}")
        return jsonify({"error": str(e)})

@app.route('/exportar')
@auth.login_required
def exportar():
    # Similar a buscar pero genera Excel
    try:
        client = get_turso_client()
        query = request.args.get('q', '').strip()
        tipo = request.args.get('tipo', 'nombre_nino')
        distrito = request.args.get('distrito', '').strip()
        genero = request.args.get('genero', '').strip()
        dia_nac = request.args.get('dia_nac', '').strip()
        mes_nac = request.args.get('mes_nac', '').strip()
        ano_nac = request.args.get('ano_nac', '').strip()
        dia_madre = request.args.get('dia_madre', '').strip()
        mes_madre = request.args.get('mes_madre', '').strip()
        ano_madre = request.args.get('ano_madre', '').strip()
        
        condiciones = []
        parametros = []
        col_map = {
            "nombre_nino": "nombre_nino",
            "nombre_responsable": "nombre_responsable",
            "cui_nino": "cui_nino",
            "cui_responsable": "cui_responsable"
        }
        col = col_map.get(tipo, "nombre_nino")
        if query:
            if tipo in ["cui_nino", "cui_responsable"]:
                condiciones.append(f'"{col}" = ?')
                parametros.append(query)
            else:
                condiciones.append(f'"{col}" LIKE ?')
                parametros.append(f"%{query}%")
        if distrito:
            condiciones.append('UPPER("distrito") LIKE ?')
            parametros.append(f"%{distrito.upper()}%")
        if genero:
            if genero == "Hombre":
                condiciones.append('"hombre" = "X"')
            elif genero == "Mujer":
                condiciones.append('"mujer" = "X"')
            elif genero == "No especificado":
                condiciones.append('"hombre" != "X" AND "mujer" != "X"')
        if dia_nac:
            condiciones.append('"día" = ?')
            parametros.append(dia_nac)
        if mes_nac:
            condiciones.append('"mes" = ?')
            parametros.append(mes_nac)
        if ano_nac:
            condiciones.append('"año_1" = ?')
            parametros.append(ano_nac)
        if dia_madre:
            condiciones.append('"día.1" = ?')
            parametros.append(dia_madre)
        if mes_madre:
            condiciones.append('"mes.1" = ?')
            parametros.append(mes_madre)
        if ano_madre:
            condiciones.append('"año.1" = ?')
            parametros.append(ano_madre)
        
        where = " AND ".join(condiciones) if condiciones else ""
        sql = f"SELECT * FROM datos_completos {f'WHERE {where}' if where else ''} LIMIT 30"
        result = client.execute(sql, parametros)
        if hasattr(result, 'rows') and callable(result.rows):
            rows = result.rows()
            columns = result.columns() if hasattr(result, 'columns') and callable(result.columns) else []
        else:
            rows = list(result)
            columns = [f"col_{i}" for i in range(len(rows[0]))] if rows else []
        data = [dict(zip(columns, row)) for row in rows]
        for row in data:
            row["genero"] = "Hombre" if row.get("hombre") == 'X' else "Mujer" if row.get("mujer") == 'X' else "No especificado"
        
        wb = Workbook()
        ws = wb.active
        ws.title = "Resultados"
        renombres = {
            "rowid": "ID",
            "año": "Año",
            "no.": "No",
            "área": "Área Salud",
            "distrito": "Distrito",
            "servicio": "Servicio Salud",
            "departamento": "Departamento",
            "municipio": "Municipio",
            "cui_nino": "CUI Niño",
            "nombre_nino": "Nombre Niño",
            "cui_responsable": "CUI Responsable",
            "nombre_responsable": "Nombre Responsable",
            "telefono": "Teléfono",
            "falleció": "Falleció",
            "genero": "Género",
        }
        excluir = ["día", "mes", "año_1", "día.1", "mes.1", "año.1", "hombre", "mujer", "comunidad", "pueblo", "rowid", "no.", "área"]
        columnas_finales = [c for c in columns if c not in excluir]
        columnas_finales.append("genero")
        titulos = [renombres.get(c, c) for c in columnas_finales]
        ws.append(titulos)
        for cell in ws[1]:
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = PatternFill(start_color="1e466e", end_color="1e466e", fill_type="solid")
            cell.alignment = Alignment(horizontal="center")
        for row in data:
            fila = []
            for c in columnas_finales:
                val = row.get(c, "")
                if c in ["hep._b", "bcg", "1a._(fipv)", "2a._(fipv)", "1a._(ipv)", "1a._(historico)", "2a._(opv)", "3a._(opv)", "1a.", "2a.", "3a.", "1a..1", "2a..1", "1a..2", "2a..2", "spr_1", "neumo-_r1", "spr_2", "r1_(opv)", "r1_(dpt)", "r2_(opv)", "r2_(dpt)", "1a..3", "2a..3", "1a..4", "2a..4", "1a..5", "2a..5", "1a._(fipv).1", "2a._(fipv).1", "1a._(ipv).1", "1a._(historico).1", "2a._(opv).1", "3a._(opv).1", "r1_(opv).1", "r2_(opv).1", "1a..6", "2a..6", "3a..1", "r1", "r2", "spr_1.1", "spr_2.1", "1a..7", "2a..7", "3a..2"]:
                    if val and isinstance(val, (int, float)) and val > 20000:
                        try:
                            fecha = datetime(1899, 12, 30) + timedelta(days=val)
                            val = fecha.strftime("%d/%m/%Y")
                        except:
                            pass
                fila.append(val)
            ws.append(fila)
        output = BytesIO()
        wb.save(output)
        output.seek(0)
        client.close()
        return Response(output.getvalue(), mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', headers={'Content-Disposition': 'attachment; filename=resultados.xlsx'})
    except Exception as e:
        return jsonify({"error": str(e)})

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)