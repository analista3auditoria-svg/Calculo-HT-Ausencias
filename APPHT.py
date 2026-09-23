import os
import io
import re
import datetime
import warnings
import pandas as pd
import numpy as np
import openpyxl
from openpyxl import Workbook, load_workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.utils.dataframe import dataframe_to_rows
import streamlit as st

# Ocultar advertencias no críticas
warnings.filterwarnings('ignore', category=UserWarning)

# Configuración de página en Streamlit
st.set_page_config(
    page_title="Modelo de Auditoria T.S. - Casalimpia", 
    page_icon="🧼", 
    layout="wide",
    initial_sidebar_state="expanded"
)

# ─── ESTILOS Y PALETA DE COLORES CORPORATIVOS CASALIMPIA ──────────────────
st.markdown("""
    <style>
        @import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700;800&display=swap');
        
        * {
            font-family: 'Plus Jakarta Sans', -apple-system, BlinkMacSystemFont, sans-serif;
        }

        #MainMenu {visibility: hidden;}
        footer {visibility: hidden;}
        
        [data-testid="stSidebar"] {
            display: block !important;
            visibility: visible !important;
        }
        [data-testid="stSidebarNav"] {
            display: block !important;
        }
        
        .main {
            background-color: #f8fafc;
        }

        .header-brand {
            background: linear-gradient(135deg, #ffffff 0%, #f0f7ff 100%);
            border-left: 6px solid #00529B;
            border-top: 1px solid #e2e8f0;
            border-right: 1px solid #e2e8f0;
            border-bottom: 1px solid #e2e8f0;
            border-radius: 16px;
            padding: 24px 32px;
            margin-bottom: 24px;
            box-shadow: 0 4px 20px -2px rgba(0, 82, 155, 0.06);
        }
        .header-brand-content {
            display: flex;
            align-items: center;
            gap: 24px;
        }
        .header-brand-content img {
            height: 58px;
            width: auto;
            object-fit: contain;
        }
        .title-text h1 {
            color: #00529B !important;
            font-size: 24px !important;
            font-weight: 700 !important;
            margin: 0 !important;
            padding: 0 !important;
            line-height: 1.2 !important;
            letter-spacing: -0.02em;
        }
        .title-text p {
            color: #64748b !important;
            font-size: 14px !important;
            margin: 4px 0 0 0 !important;
            padding: 0 !important;
            font-weight: 500;
        }

        .kpi-card {
            background-color: #ffffff;
            border-radius: 16px;
            padding: 20px 24px;
            box-shadow: 0 4px 12px rgba(0, 0, 0, 0.03);
            border: 1px solid #e2e8f0;
            display: flex;
            flex-direction: column;
            justify-content: space-between;
            transition: transform 0.2s ease, box-shadow 0.2s ease;
        }
        .kpi-card:hover {
            transform: translateY(-2px);
            box-shadow: 0 6px 16px rgba(0, 0, 0, 0.06);
        }
        .kpi-card-danger {
            border-left: 6px solid #dc2626;
            background: linear-gradient(135deg, #ffffff 0%, #fef2f2 100%);
        }
        .kpi-card-warning {
            border-left: 6px solid #d97706;
            background: linear-gradient(135deg, #ffffff 0%, #fffbeb 100%);
        }
        .kpi-card-info {
            border-left: 6px solid #00529B;
            background: linear-gradient(135deg, #ffffff 0%, #f0f7ff 100%);
        }
        .kpi-card-success {
            border-left: 6px solid #16a34a;
            background: linear-gradient(135deg, #ffffff 0%, #f0fdf4 100%);
        }
        .kpi-title {
            font-size: 13px;
            font-weight: 700;
            text-transform: uppercase;
            letter-spacing: 0.05em;
            margin-bottom: 8px;
        }
        .kpi-title-danger { color: #991b1b; }
        .kpi-title-warning { color: #92400e; }
        .kpi-title-info { color: #00529B; }
        .kpi-title-success { color: #14532d; }

        .kpi-value {
            font-size: 36px;
            font-weight: 800;
            line-height: 1;
            margin-bottom: 0px;
        }
        .kpi-value-danger { color: #dc2626; }
        .kpi-value-warning { color: #d97706; }
        .kpi-value-info { color: #00529B; }
        .kpi-value-success { color: #16a34a; }

        [data-testid="stFileUploader"] {
            padding: 0px;
            margin-bottom: 12px;
        }
        [data-testid="stFileUploaderDropzone"] {
            padding: 10px 14px !important;
            border: 1.5px dashed #cbd5e1 !important;
            border-radius: 10px !important;
            background-color: #f8fafc !important;
            display: flex !important;
            flex-direction: row !important;
            align-items: center !important;
            justify-content: flex-start !important;
            gap: 12px !important;
            min-height: 46px !important;
            transition: all 0.2s ease;
            position: relative;
        }
        [data-testid="stFileUploaderDropzone"]:hover {
            background-color: #f0f7ff !important;
            border-color: #00529B !important;
        }
        [data-testid="stFileUploaderDropzone"] section,
        [data-testid="stFileUploaderDropzone"] small,
        [data-testid="stFileUploaderDropzone"] svg {
            display: none !important;
        }
        [data-testid="stFileUploaderDropzone"]::before {
            content: "📄 Cargar" !important;
            display: inline-block !important;
            background-color: #00529B !important;
            color: #ffffff !important;
            border: none !important;
            border-radius: 6px !important;
            padding: 5px 12px !important;
            font-size: 12.5px !important;
            font-weight: 600 !important;
            cursor: pointer !important;
        }
        [data-testid="stFileUploaderDropzone"] button {
            display: none !important;
        }

        /* Cuadros de archivos cargados en color verde esmeralda */
        [data-testid="stFileUploaderFileData"] > div:first-child,
        [data-testid="stFileUploaderFileData"] svg,
        div[data-testid="stFileUploaderFileData"] > span:first-child,
        [data-testid="stFileUploaderFileData"] [data-testid="stFileUploaderDeleteBtn"] ~ div,
        [data-testid="stFileUploaderFileData"] > div {
            background-color: #059669 !important;
            color: #ffffff !important;
            border-radius: 8px !important;
            fill: #ffffff !important;
        }

        [data-testid="stFileUploaderFileData"] > div:first-child::after {
            content: "✓" !important;
            display: flex !important;
            align-items: center !important;
            justify-content: center !important;
            background-color: #059669 !important;
            color: #ffffff !important;
            font-weight: 800 !important;
            font-size: 16px !important;
            width: 32px !important;
            height: 32px !important;
            border-radius: 8px !important;
            position: absolute !important;
            left: 0 !important;
            top: 0 !important;
        }

        [data-testid="stFileUploaderFileData"] svg {
            display: none !important;
        }

        /* Ventana desplegable del selectbox redimensionable */
        div[data-baseweb="popover"],
        div[data-baseweb="menu"] {
            resize: both !important;
            overflow: auto !important;
            min-width: 380px !important;
            min-height: 200px !important;
            max-width: 90vw !important;
            max-height: 80vh !important;
        }

        ul[data-testid="stSelectboxVirtualDropdown"] {
            resize: both !important;
            overflow: auto !important;
            min-width: 380px !important;
            width: 100% !important;
        }

        ul[data-testid="stSelectboxVirtualDropdown"] li {
            white-space: normal !important;
            word-break: break-word !important;
            padding-top: 8px !important;
            padding-bottom: 8px !important;
            line-height: 1.3 !important;
        }

        div.stButton > button:first-child {
            background: linear-gradient(135deg, #00529B 0%, #003366 100%) !important;
            color: white !important;
            border-radius: 10px !important;
            padding: 12px 28px !important;
            font-weight: 700 !important;
            font-size: 16px !important;
            border: none !important;
            box-shadow: 0 4px 12px rgba(0, 82, 155, 0.25) !important;
            transition: all 0.2s ease !important;
            width: 100%;
        }
        div.stButton > button:first-child:hover {
            background: linear-gradient(135deg, #003366 0%, #002244 100%) !important;
            box-shadow: 0 6px 16px rgba(0, 82, 155, 0.35) !important;
            transform: translateY(-1px);
        }

        div.stDownloadButton > button:first-child {
            background: linear-gradient(135deg, #059669 0%, #047857 100%) !important;
            color: white !important;
            border-radius: 10px !important;
            padding: 10px 20px !important;
            font-weight: 700 !important;
            font-size: 14px !important;
            border: none !important;
            box-shadow: 0 4px 12px rgba(5, 150, 105, 0.25) !important;
            width: 100%;
        }
        div.stDownloadButton > button:first-child:hover {
            background: linear-gradient(135deg, #047857 0%, #065f46 100%) !important;
            box-shadow: 0 6px 16px rgba(5, 150, 105, 0.35) !important;
        }
    </style>

    <div class="header-brand">
        <div class="header-brand-content">
            <img src="https://cdn1.totalcommerce.cloud/casalimpia/web_content/assets/logo-casa-limpia.svg" alt="Casalimpia Logo" />
            <div class="title-text">
                <h1>Modelo de Auditoria T.S.</h1>
                <p>Auditoría de Tiempos y Análisis de Compensatorios</p>
            </div>
        </div>
    </div>
""", unsafe_allow_html=True)


# ─── CONSTANTES Y MAPEOS GLOBALES ───────────────────────────────────────────

MAPA_ausencias = {
    "ninguno":                     None,
    "ausencia":                  "A",
    "ausencia sin justa causa sd": "FSJ",
    "enfermedad comun":          "EC",
    "accidente de trabajo":      "AT",
    "descanso":                  "DES",
    "dia de la familia":         "DF",
    "licencia no remunerada":    "LNR",
    "licencia remunerada":       "LR",
    "parada de salario":         "PS",
    "suspensiones":              "S",
    "descanso compensatorio":    "C",
    "festivo compensatorio":     "FC",
    "vacaciones":                "V",
    "incapacidad":               "INC",
    "licencia de maternidad":    "LM",
    "remunerada":                "LR",
    "no remunerada":             "LNR",
    "licencia por luto":         "LT",
    "calamidad doméstica":       "CD",
    "falta sin justificar":      "FSJ",
    "suspensión":                "S",
    "CCCO":                      "CCCO",
    "retiros":                   "R",
    "perdida dominical":         "PD",
    "p":                         "P"   
}

AUSENCIAS_ROJO_SET = {
    "A", "INC", "AT", "LM", "LR", "LNR", "DF", "CD",
    "S", "FC", "DES", "PS", "FSJ", "PD"
}
AUSENCIAS_AMARILLO_SET = {"P"}
MESES_ES = {1:"ene",2:"feb",3:"mar",4:"abr",5:"may",6:"jun",7:"jul",8:"ago",9:"sep",10:"oct",11:"nov",12:"dic"}

PERIODOS = {
    "Marzo":      (pd.Timestamp(2026, 2, 15), pd.Timestamp(2026, 3, 21)),
    "Abril":      (pd.Timestamp(2026, 3, 22), pd.Timestamp(2026, 4, 18)),
    "Mayo":       (pd.Timestamp(2026, 4, 19), pd.Timestamp(2026, 5, 16)),
    "Junio":      (pd.Timestamp(2026, 5, 17), pd.Timestamp(2026, 6, 13)),
    "Julio":      (pd.Timestamp(2026, 6, 14), pd.Timestamp(2026, 7, 11)),
    "Agosto":     (pd.Timestamp(2026, 7, 12), pd.Timestamp(2026, 8, 8)),
    "Septiembre": (pd.Timestamp(2026, 8, 9), pd.Timestamp(2026, 9, 19)),
    "Octubre":    (pd.Timestamp(2026, 9, 20), pd.Timestamp(2026, 10, 17)),
    "Noviembre":  (pd.Timestamp(2026, 10, 18), pd.Timestamp(2026, 11, 14)),
    "Diciembre":  (pd.Timestamp(2026, 11, 15), pd.Timestamp(2026, 12, 12))
}

MAPA_CONCEPTOS = {
    "HT Normales":                        "ht normales",
    "Recargo Nocturno 0.35%":             "recargo nocturno 0.35%",
    "Recargo Dominical Compensado":      "recargo dominical compensado",
    "Recargo Dominical No Compensado": "recargo dominical no compensado",
    "Horas Extras Diurnas 1.25%":        "horas extras diurnas 1.25%",
    "Recargo Festivo":                  "recargo festivo",
    "Hora Extra Diurna Dom/Fest":        "hora extra diurna dom/fest",
    "Recargo Festivo Adicional":        "recargo festivo adicional",
}

COLORES_LETRAS = {
    "A":   {"bg": "FF0000", "fg": "FFFFFF"},
    "S":   {"bg": "FF0000", "fg": "FFFFFF"},
    "INC": {"bg": "FF0000", "fg": "FFFFFF"},
    "C":   {"bg": "F4B942", "fg": "000000"},
    "P":   {"bg": "C00000", "fg": "FFFFFF"},
    "LIC": {"bg": "FFC000", "fg": "000000"},
    "DIA": {"bg": "92D050", "fg": "000000"},
}

# ─── FUNCIONES AUXILIARES ─────────────────────────────────────────────────

def obtener_val_iloc(row, index_col):
    if isinstance(row, dict):
        keys = list(row.keys())
        if len(keys) > index_col:
            val = row[keys[index_col]]
            return str(val).strip() if pd.notna(val) else ""
        return ""
    if hasattr(row, 'iloc') and len(row) > index_col:
        val = row.iloc[index_col]
        return str(val).strip() if pd.notna(val) else ""
    return ""

def convertir_a_hora(val):
    if not val or pd.isna(val):
        return None
    try:
        dt = pd.to_datetime(val, format='%H:%M', errors='coerce')
        if pd.isna(dt):
            dt = pd.to_datetime(val, errors='coerce')
        if pd.notna(dt):
            return dt.time()
    except Exception:
        pass
    return None

def estilo_etiqueta_ausentismo(val):
    if pd.isna(val) or not str(val).strip():
        return ''
    val_str = str(val).strip().lower()
    if 'ausencia' in val_str or 'requiere' in val_str:
        return 'background-color: #FCE8E6; color: #991B1B; border: 1.5px solid #FDBA74; border-radius: 12px; font-weight: bold; text-align: center; padding: 4px 8px;'
    elif val_str == 'p':
        return 'background-color: #FEF3C7; color: #92400E; border: 1.5px solid #FCD34D; border-radius: 12px; font-weight: bold; text-align: center; padding: 4px 8px;'
    else:
        return 'background-color: #E0F2FE; color: #0369A1; border: 1.5px solid #7DD3FC; border-radius: 12px; font-weight: bold; text-align: center; padding: 4px 8px;'

def df_a_excel_bytes(df, sheet_name="Detalle_Novedades"):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)
    output.seek(0)
    return output.getvalue()

def hhmm_a_decimal(texto):
    if pd.isna(texto): return None
    t = str(texto).strip()
    if t in ("", "0", "nan", "None"): return None
    m = re.match(r'^(\d+):(\d{2})$', t)
    if m:
        val = int(m.group(1)) + int(m.group(2)) / 60.0
        return round(val, 2) if val > 0 else None
    try:
        val = float(t)
        return round(val, 2) if val > 0 else None
    except ValueError: return None

def num_limpio(val):
    if pd.isna(val): return None
    try:
        v = float(val)
        return round(v, 2) if v > 0 else None
    except (ValueError, TypeError): return None

def resolver_ausencia(texto_ausent):
    if pd.isna(texto_ausent): return None
    t = str(texto_ausent).strip()
    if not t or t.lower() in ('nan', 'none'): return None
    clave = t.lower()
    if clave in MAPA_ausencias: return MAPA_ausencias[clave]
    for k, v in MAPA_ausencias.items():
        if k and k in clave: return v
    return t

def celda(valor, ausent_raw):
    v = valor
    a = resolver_ausencia(ausent_raw)
    if a == 'C': return 'C'
    if v is not None and a is not None:
        v_str = str(int(v)) if v % 1 == 0 else str(round(v, 2))
        return f'{v_str} {a}'
    if v is not None: return int(v) if v % 1 == 0 else round(v, 2)
    return a

def limpiar_id_a_texto(valor):
    if isinstance(valor, pd.Series):
        temp = pd.to_numeric(valor, errors='coerce')
        return temp.fillna(0).astype(int).astype(str).replace('0', '')
    if pd.isna(valor): return ""
    s = str(valor).strip()
    if s.lower() in ("nan", "none", "0", ""): return ""
    if s.endswith(".0"): s = s[:-2]
    try: return str(int(float(s)))
    except: return s

def normalizar_concepto_txt(texto):
    if not texto: return ""
    t = str(texto).strip().lower()
    t = re.sub(r'\s+', ' ', t)
    return t

def fecha_a_clave_corta(val_fecha):
    if pd.isna(val_fecha) or val_fecha is None: return None
    if isinstance(val_fecha, (pd.Timestamp, pd.DatetimeIndex)):
        return f"{val_fecha.day}-{MESES_ES[val_fecha.month]}"
    try:
        d = pd.to_datetime(val_fecha, errors='coerce')
        if pd.notna(d): return f"{d.day}-{MESES_ES[d.month]}"
    except: pass
    s = str(val_fecha).strip().lower()
    m = re.search(r'(\d{1,2})[-_\s/]+([a-z]{3})', s)
    if m: return f"{int(m.group(1))}-{m.group(2)}"
    return s

def obtener_dt_fecha(val_fecha):
    if pd.isna(val_fecha) or val_fecha is None: return None
    if isinstance(val_fecha, pd.Timestamp): return val_fecha.normalize()
    try:
        d = pd.to_datetime(val_fecha, errors='coerce')
        if pd.notna(d): return d.normalize()
    except: pass
    return None

def valores_son_diferentes(v1, v2):
    if (v1 is None or str(v1).strip() in ("", "None", "nan", "0", "0.0", "0,00")) and \
       (v2 is None or str(v2).strip() in ("", "None", "nan", "0", "0.0", "0,00")):
        return False
    try:
        n1 = round(float(v1), 2) if v1 is not None else 0.0
        n2 = round(float(v2), 2) if v2 is not None else 0.0
        return abs(n1 - n2) > 0.01
    except:
        s1 = str(v1).strip().upper() if v1 is not None else ""
        s2 = str(v2).strip().upper() if v2 is not None else ""
        return s1 != s2


# ─── INTERFAZ PRINCIPAL Y CARGA DE ARCHIVOS ─────────────────────────────

with st.expander("📁 Bases de datos", expanded=True):
    col1, col2 = st.columns(2, gap="large")

    with col1:
        file_entrada = st.file_uploader("1. BBDD Marcaciones Geovictoria (.xlsx)", type=["xlsx"], help="Origen Geovictoria")
        file_operativa = st.file_uploader("2. BBDD Nómina Compensación de tiempo (.xlsx)", type=["xlsx"], help="BD que el supervisor envía a nómina con los compensatorios")
        file_novasoft = st.file_uploader("3. BBDD Ausentismos Novasoft (.xlsx)", type=["xlsx"], help="Archivos descargados por el usuario de Novasoft")
        file_supernumerario = st.file_uploader("7. BBDD Ubicaciones (.xlsx)", type=["xlsx"], help="Ubicaciones descargadas del módulo de supernumerarios")

    with col2:
        file_sic = st.file_uploader("4. BBDD Gestión de personal SIC (.xlsx)", type=["xlsx"], help="Archivo descargado por el usuario del módulo SIC")
        file_maestro = st.file_uploader("5. BBDD Maestro de empleados (.xlsx)", type=["xlsx"], help="BD maestro del personal de la compañía")
        file_historial = st.file_uploader("6. BBDD Historia laboral de empleados (.xlsx)", type=["xlsx"], help="BD descargada del SIC")
        file_nomina = st.file_uploader("8. Nómina (.xlsx)", type=["xlsx"], help="Plantilla de Nómina para consolidación HTCC")

lista_cc = ["FUNDACION HOSPITAL DE LA MISERICORDIA"]

if file_historial:
    try:
        excel_hist_temp = pd.ExcelFile(file_historial)
        target_data_sheet = None
        for sheet_name in excel_hist_temp.sheet_names:
            if sheet_name.strip().lower() == "data":
                target_data_sheet = sheet_name
                break
        
        if target_data_sheet:
            df_cc_data = pd.read_excel(file_historial, sheet_name=target_data_sheet)
            if not df_cc_data.empty:
                centros_extraidos = df_cc_data.iloc[:, 0].dropna().astype(str).str.strip().unique().tolist()
                centros_extraidos = [c for c in centros_extraidos if c.lower() != "centro de costos"]
                if centros_extraidos:
                    lista_cc = sorted(list(set(centros_extraidos)))
    except Exception as e:
        st.warning(f"⚠️ No se pudo leer la hoja 'data' del Historial Laboral: {e}")

st.sidebar.markdown("## ⚙️ Parámetros de Configuración")
contrato_principal = st.sidebar.selectbox("Contrato / CC Principal", options=lista_cc, index=0)

st.sidebar.markdown("---")
st.sidebar.markdown("### 📅 Filtro Rango de Fechas")
fecha_ini_sup = st.sidebar.date_input("Fecha Inicial", value=datetime.date(2026, 2, 1))
fecha_fin_sup = st.sidebar.date_input("Fecha Final", value=datetime.date(2026, 2, 28))

st.sidebar.markdown("---")
st.sidebar.markdown("""
<div style="background-color: #f0f7ff; padding: 12px; border-radius: 8px; border-left: 4px solid #00529B;">
    <small style="color: #00529B; font-weight: 600;">💡 Instrucciones</small><br>
    <small style="color: #475569;">1. Carga los archivos requeridos.<br>2. Selecciona el Centro de Costos.<br>3. Ajusta las fechas y ejecuta la auditoría.</small>
</div>
""", unsafe_allow_html=True)

with st.expander("🛠️ Configuración Avanzada de Pestañas (Opcional)"):
    st.caption("Solo modifica estos campos si los libros de Excel tienen nombres de hoja diferentes a los estándar.")
    c_a, c_b = st.columns(2)
    with c_a:
        hoja_entrada = st.text_input("1. Hoja Marcaciones", value="Marcaciones")
        hoja_festivos = st.text_input("Hoja Festivos", value="Festivos")
        hoja_operativa = st.text_input("2. Hoja Operativa", value="CONSOLIDADO")
        hoja_novasoft = st.text_input("3. Hoja Novasoft", value="BBDD_Novasof")
    with c_b:
        hoja_sic = st.text_input("4. Hoja SIC", value="Datos")
        hoja_maestro = st.text_input("5. Hoja Maestro", value="NOM1911")
        hoja_historial = st.text_input("6. Hoja Historial", value="Hoja 1")
        hoja_supernumerario = st.text_input("7. Hoja Supernumerario", value="Base")

st.markdown("<br>", unsafe_allow_html=True)

# ─── PROCESAMIENTO GENERAL AL PRESIONAR EL BOTÓN ─────────────────────────

if st.button("⚡ Ejecutar Auditoría TS y Procesar Marcaciones", type="primary"):
    if not file_entrada:
        st.error("⚠️ Es obligatorio cargar el archivo principal de Marcaciones (GeoVictoria).")
    elif not contrato_principal:
        st.error("⚠️ Por favor, selecciona el Contrato / Centro de Costo Principal en el panel izquierdo.")
    else:
        try:
            with st.spinner("Procesando marcaciones y generando auditoría unificada..."):
                df_marc_raw = pd.read_excel(file_entrada, sheet_name=hoja_entrada)
                
                operativa_dict = {}
                if file_operativa:
                    try:
                        excel_op = pd.ExcelFile(file_operativa)
                        target_sheet = hoja_operativa
                        for name in excel_op.sheet_names:
                            if name.strip().upper() == hoja_operativa.strip().upper():
                                target_sheet = name
                                break
                        df_op = pd.read_excel(file_operativa, sheet_name=target_sheet)
                        
                        if not df_op.empty:
                            df_op['Cédula_Str'] = df_op.apply(lambda r: obtener_val_iloc(r, 0).replace('.0', ''), axis=1)
                            df_op['Fecha_Dt'] = pd.to_datetime(df_op.iloc[:, 2], dayfirst=True, errors='coerce')
                            df_op['Val_C'] = df_op.apply(lambda r: obtener_val_iloc(r, 8), axis=1)
                            
                            for _, row_op in df_op.iterrows():
                                ced_op = row_op['Cédula_Str']
                                f_op = row_op['Fecha_Dt']
                                val_c = row_op['Val_C']
                                if ced_op and pd.notna(f_op):
                                    operativa_dict[(ced_op, f_op.date())] = val_c
                    except Exception as e:
                        st.warning(f"⚠️ No se pudo procesar la hoja '{hoja_operativa}' de la Base Operativa: {e}")

                df_super_filtrado = pd.DataFrame()
                if file_supernumerario:
                    try:
                        excel_sup = pd.ExcelFile(file_supernumerario)
                        target_sheet_sup = hoja_supernumerario
                        for name in excel_sup.sheet_names:
                            if name.strip().upper() == hoja_supernumerario.strip().upper():
                                target_sheet_sup = name
                                break
                        df_sup_raw = pd.read_excel(file_supernumerario, sheet_name=target_sheet_sup)
                        
                        if not df_sup_raw.empty and df_sup_raw.shape[1] > 12:
                            col_m_val = df_sup_raw.iloc[:, 12].astype(str).str.strip().str.upper()
                            contrato_target = str(contrato_principal).strip().upper()
                            mask = col_m_val.str.contains(contrato_target, regex=False, na=False)

                            if fecha_ini_sup and fecha_fin_sup:
                                fechas_col_b = pd.to_datetime(df_sup_raw.iloc[:, 1], dayfirst=True, errors='coerce').dt.date
                                mask = mask & (fechas_col_b >= fecha_ini_sup) & (fechas_col_b <= fecha_fin_sup)

                            df_super_filtrado = df_sup_raw[mask].copy()
                    except Exception as e:
                        st.warning(f"⚠️ No se pudo procesar la BD Supernumerario: {e}")

                df_nova = pd.read_excel(file_novasoft, sheet_name=hoja_novasoft) if file_novasoft else pd.DataFrame()
                df_sic = pd.read_excel(file_sic, sheet_name=hoja_sic) if file_sic else pd.DataFrame()
                df_maestro = pd.read_excel(file_maestro, sheet_name=hoja_maestro) if file_maestro else pd.DataFrame()
                df_hist = pd.read_excel(file_historial, sheet_name=hoja_historial) if file_historial else pd.DataFrame()

                set_festivos = set()
                if file_entrada:
                    try:
                        df_festivos = pd.read_excel(file_entrada, sheet_name=hoja_festivos)
                        if not df_festivos.empty:
                            fechas_fest = pd.to_datetime(df_festivos.iloc[:, 0], dayfirst=True, errors='coerce').dropna()
                            set_festivos = set(fechas_fest.dt.date)
                    except Exception as e:
                        st.warning(f"⚠️ Nota: No se pudo cargar la hoja '{hoja_festivos}'. Se continuará sin marcar festivos.")

                df_marc_raw['Cédula_Str'] = df_marc_raw.apply(lambda r: obtener_val_iloc(r, 2).replace('.0', ''), axis=1)
                
                if not df_hist.empty:
                    df_hist['Cédula_Str'] = df_hist.apply(lambda r: obtener_val_iloc(r, 0).replace('.0', ''), axis=1)
                    df_hist['Centro_Costo'] = df_hist.apply(lambda r: obtener_val_iloc(r, 2), axis=1)
                    df_hist['Fecha_Inicio'] = pd.to_datetime(df_hist.iloc[:, 3], dayfirst=True, errors='coerce') if df_hist.shape[1] > 3 else pd.NaT
                    if df_hist.shape[1] > 4:
                        df_hist['Fecha_Fin_Raw'] = pd.to_datetime(df_hist.iloc[:, 4], dayfirst=True, errors='coerce')
                        fecha_dummy = pd.to_datetime('2001-01-01')
                        df_hist['Fecha_Fin'] = df_hist['Fecha_Fin_Raw'].apply(
                            lambda x: pd.to_datetime('2099-12-31') if (pd.isna(x) or x == fecha_dummy) else x
                        )
                    else:
                        df_hist['Fecha_Fin'] = pd.to_datetime('2099-12-31')
                    df_hist['Frente_Trabajo'] = df_hist.apply(lambda r: obtener_val_iloc(r, 5), axis=1) if df_hist.shape[1] > 5 else df_hist['Centro_Costo']

                if not df_nova.empty:
                    df_nova['Cédula_Str'] = df_nova.apply(lambda r: obtener_val_iloc(r, 0).replace('.0', ''), axis=1)
                    df_nova['Concepto'] = df_nova.apply(lambda r: obtener_val_iloc(r, 2), axis=1)
                    df_nova['Fecha_Inicio'] = pd.to_datetime(df_nova.iloc[:, 3], dayfirst=True, errors='coerce') if df_nova.shape[1] > 3 else pd.NaT
                    df_nova['Fecha_Fin'] = pd.to_datetime(df_nova.iloc[:, 4], dayfirst=True, errors='coerce') if df_nova.shape[1] > 4 else pd.NaT
                    df_nova['Codigo_Novasoft'] = df_nova.apply(lambda r: obtener_val_iloc(r, 8), axis=1)

                if not df_sic.empty:
                    df_sic['Cédula_Str'] = df_sic.apply(lambda r: obtener_val_iloc(r, 9).replace('.0', ''), axis=1)
                    df_sic['Proceso'] = df_sic.apply(lambda r: obtener_val_iloc(r, 1), axis=1)
                    df_sic['Estado'] = df_sic.apply(lambda r: obtener_val_iloc(r, 32), axis=1)
                    df_sic = df_sic[df_sic['Estado'].str.lower() == 'nomina'].copy()
                    if df_sic.shape[1] > 4:
                        df_sic['Fecha_Inicio'] = pd.to_datetime(df_sic.iloc[:, 4], dayfirst=True, errors='coerce')
                    if df_sic.shape[1] > 5:
                        df_sic['Fecha_Fin'] = pd.to_datetime(df_sic.iloc[:, 5], dayfirst=True, errors='coerce')

                maestro_dict = {}
                if not df_maestro.empty:
                    df_maestro['Cédula_Str'] = df_maestro.apply(lambda r: obtener_val_iloc(r, 1).replace('.0', ''), axis=1)
                    df_maestro['F_INGRESO'] = pd.to_datetime(df_maestro.iloc[:, 26], dayfirst=True, errors='coerce') if df_maestro.shape[1] > 26 else pd.NaT
                    df_maestro['F_RETIRO'] = pd.to_datetime(df_maestro.iloc[:, 27], dayfirst=True, errors='coerce') if df_maestro.shape[1] > 27 else pd.NaT
                    for _, row_m in df_maestro.iterrows():
                        if row_m['Cédula_Str']:
                            maestro_dict[row_m['Cédula_Str']] = (row_m['F_INGRESO'], row_m['F_RETIRO'])

                df_marc_raw['Fecha_Ori_Dt'] = df_marc_raw.apply(
                    lambda r: pd.to_datetime(str(r.iloc[4])[-10:], dayfirst=True, errors='coerce') if len(str(r.iloc[4])) >= 10 else pd.NaT,
                    axis=1
                )

                rango_dias = pd.date_range(start=fecha_ini_sup, end=fecha_fin_sup).date if (fecha_ini_sup and fecha_fin_sup) else []
                empleados_unicos = df_marc_raw['Cédula_Str'].unique()
                filas_construidas = []

                cols_keys = [c for c in list(df_marc_raw.columns) if c != 'Cédula_Str']

                for ced in empleados_unicos:
                    if not ced:
                        continue
                    df_emp = df_marc_raw[df_marc_raw['Cédula_Str'] == ced]
                    dict_fechas_emp = {row['Fecha_Ori_Dt'].date(): row.to_dict() for _, row in df_emp.iterrows() if pd.notna(row['Fecha_Ori_Dt'])}
                    row_base = df_emp.iloc[0].to_dict()

                    if len(rango_dias) > 0:
                        for f_dia in rango_dias:
                            if f_dia in dict_fechas_emp:
                                filas_construidas.append(dict_fechas_emp[f_dia])
                            else:
                                new_row = row_base.copy()
                                new_row['Fecha_Ori_Dt'] = pd.Timestamp(f_dia)
                                if len(cols_keys) > 4:
                                    new_row[cols_keys[4]] = f_dia.strftime('%d/%m/%Y')
                                
                                for c_idx in [7, 9, 10, 12]:
                                    if c_idx < len(cols_keys):
                                        new_row[cols_keys[c_idx]] = None
                                filas_construidas.append(new_row)
                    else:
                        filas_construidas.extend(df_emp.to_dict('records'))

                df_marc = pd.DataFrame(filas_construidas)

                file_entrada.seek(0)
                wb = openpyxl.load_workbook(file_entrada, data_only=False)
                ws = wb[hoja_entrada]
                ws.views.sheetView[0].showGridLines = True

                encabezados_estilos = [
                    ("AY1", "Fecha Ori", "D0E1F9", "002244", True),
                    ("AZ1", "Dia", "D0E1F9", "002244", True),
                    ("BA1", "Entrada2", "FCE8E6", "A80000", True),
                    ("BB1", "Salida2", "FCE8E6", "A80000", True),
                    ("BC1", "Cant HT", "E2E8F0", "1E293B", True),
                    ("BD1", "HT", "E2E8F0", "1E293B", True),
                    ("BE1", "Compensatorio", "00529B", "FFFFFF", True),
                    ("BF1", "Ausencias / Marcaciones Erroneas", "990000", "FFFFFF", True),
                    ("BG1", "Recargo Dominical No Compensado", "DCFCE7", "14532D", True),
                    ("BH1", "Recargo Dominical Compensado", "FEF08A", "713F12", True),
                    ("BI1", "Recargo Festivo", "FFEDD5", "7C2D12", True),
                    ("BJ1", "Recargo Nocturno 0.35%", "BAE6FD", "0369A1", True),
                    ("BK1", "Horas Extras Diurnas 1.25%", "E9D5FF", "581C87", True),
                    ("BL1", "Hora Extra Diurna Dom/Fest", "BBF7D0", "166534", True),
                    ("BM1", "Horas Extras Nocturnas 1.75%", "FECDD3", "9F1239", True),
                    ("BN1", "Hora Extra Dominical o Festiva Nocturna", "C6EFCE", "064E3B", True),
                    ("BO1", "CCCO", "003366", "FFFFFF", True),
                    ("BP1", "Frente de trabajo", "003366", "FFFFFF", True),
                    ("BQ1", "Centro de costos", "003366", "FFFFFF", True),
                    ("BR1", "Fecha Ingreso", "00529B", "FFFFFF", True),
                    ("BS1", "Validar Fingreso", "00529B", "FFFFFF", True),
                    ("BT1", "F Retiro", "00529B", "FFFFFF", True),
                    ("BU1", "Validar F Retiro", "00529B", "FFFFFF", True),
                    ("BV1", "Ausentismo Novasoft", "990000", "FFFFFF", True),
                    ("BW1", "Codigo novasoft", "00529B", "FFFFFF", True),
                    ("BX1", "Ausentismo Sic", "00529B", "FFFFFF", True),
                    ("BY1", "Ausentismo", "C00000", "FFFFFF", True),
                    ("BZ1", "Compensado", "00529B", "FFFFFF", True),
                    ("CA1", "Ausentismo Real", "1E4620", "FFFFFF", True)
                ]

                thin_border = Border(
                    left=Side(style='thin', color='CBD5E1'),
                    right=Side(style='thin', color='CBD5E1'),
                    top=Side(style='thin', color='CBD5E1'),
                    bottom=Side(style='thin', color='CBD5E1')
                )

                fill_ausencia_rojo = PatternFill(start_color="C00000", end_color="C00000", fill_type="solid")
                font_ausencia_blanco = Font(name="Calibri", size=10, bold=True, color="FFFFFF")

                for celda_ref, titulo, color_bg, color_fg, es_negrita in encabezados_estilos:
                    celda = ws[celda_ref]
                    celda.value = titulo
                    celda.fill = PatternFill(start_color=color_bg, end_color=color_bg, fill_type="solid")
                    celda.font = Font(name="Calibri", size=10, bold=es_negrita, color=color_fg)
                    celda.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
                    celda.border = thin_border

                dias_semana_es = {0: "lunes", 1: "martes", 2: "miércoles", 3: "jueves", 4: "viernes", 5: "sábado", 6: "domingo"}
                hist_dict = {ced: grp for ced, grp in df_hist.groupby('Cédula_Str')} if not df_hist.empty else {}
                nova_dict = {ced: grp for ced, grp in df_nova.groupby('Cédula_Str')} if not df_nova.empty else {}
                sic_dict = {ced: grp for ced, grp in df_sic.groupby('Cédula_Str')} if not df_sic.empty else {}
                fecha_minima_valida = pd.to_datetime('1900-01-01')

                total_filas = len(df_marc)
                conteo_ausencias = 0
                conteo_p = 0
                registros_ausencias = []
                registros_novedades = []
                marcaciones_ht_dict = {}

                hora_corte_nocturna = datetime.time(20, 0)

                for idx, row in df_marc.iterrows():
                    i = idx + 2

                    cols_df = list(df_marc.columns)
                    for col_i, col_name in enumerate(cols_df, start=1):
                        val_col = row[col_name]
                        if str(col_name) != 'Fecha_Ori_Dt':
                            ws.cell(row=i, column=col_i, value=val_col if pd.notna(val_col) else "")

                    val_e = obtener_val_iloc(row, 4)
                    fecha_ori = None
                    if isinstance(row.get('Fecha_Ori_Dt'), (datetime.date, datetime.datetime, pd.Timestamp)):
                        fecha_ori = pd.to_datetime(row['Fecha_Ori_Dt'])
                    elif len(val_e) >= 10:
                        try:
                            fecha_ori = pd.to_datetime(val_e[-10:], dayfirst=True)
                        except Exception:
                            fecha_ori = None
                    
                    celda_ay = ws[f'AY{i}']
                    if fecha_ori is not pd.NaT and fecha_ori is not None:
                        celda_ay.value = fecha_ori.date()
                        celda_ay.number_format = 'DD/MM/YYYY'
                    else:
                        celda_ay.value = val_e

                    dia_nombre = ""
                    if fecha_ori and pd.notna(fecha_ori):
                        dia_nombre = dias_semana_es[fecha_ori.weekday()]
                        if fecha_ori.date() in set_festivos:
                            dia_nombre += " Festivo"
                    
                    ws[f'AZ{i}'].value = dia_nombre

                    h_val, j_val = obtener_val_iloc(row, 7), obtener_val_iloc(row, 9)
                    k_val, m_val = obtener_val_iloc(row, 10), obtener_val_iloc(row, 12)

                    hora_h = convertir_a_hora(h_val)
                    hora_j = convertir_a_hora(j_val)
                    hora_k = convertir_a_hora(k_val)
                    hora_m = convertir_a_hora(m_val)

                    celda_ba = ws[f'BA{i}']
                    celda_bb = ws[f'BB{i}']

                    if hora_k is not None and hora_k > hora_corte_nocturna:
                        celda_ba.value = hora_k
                        celda_ba.number_format = 'hh:mm:ss AM/PM'

                        if hora_j is not None:
                            celda_bb.value = hora_j
                            celda_bb.number_format = 'hh:mm:ss AM/PM'
                        else:
                            celda_bb.value = ""
                    else:
                        horas_validas = [dt for dt in [hora_h, hora_j, hora_k, hora_m] if dt is not None]

                        if len(horas_validas) > 0:
                            celda_ba.value = min(horas_validas)
                            celda_ba.number_format = 'hh:mm:ss AM/PM'
                        else:
                            celda_ba.value = ""

                        if len(horas_validas) > 1:
                            celda_bb.value = max(horas_validas)
                            celda_bb.number_format = 'hh:mm:ss AM/PM'
                        else:
                            celda_bb.value = ""

                    val_ba = celda_ba.value
                    val_bb = celda_bb.value

                    ws[f'BC{i}'].value = f'=IF(OR(BA{i}="",BB{i}=""),"",MOD(BB{i}-BA{i},1))'
                    ws[f'BC{i}'].number_format = '[h]:mm'
                    ws[f'BD{i}'] = f'=IFERROR(ROUND(BC{i}*24,1),"")'

                    f_val = obtener_val_iloc(row, 5)
                    ws[f'BE{i}'] = "C" if f_val == "Descanso compensatorio" else ""

                    has_h = h_val != ""
                    has_j = j_val != ""
                    has_k = k_val != ""
                    has_m = m_val != ""
                    count_marcas = sum([has_h, has_j, has_k, has_m])

                    if count_marcas >= 2:
                        val_bf = ""
                    elif count_marcas == 0:
                        if dia_nombre.lower() == "domingo" or "festivo" in dia_nombre.lower():
                            val_bf = "Descanso"
                        else:
                            val_bf = "Ausencia"
                    else:
                        val_bf = "P"

                    ws[f'BF{i}'] = val_bf

                    ws[f'BG{i}'] = f'=AM{i}+AO{i}'
                    ws[f'BH{i}'] = f'=AI{i}+AK{i}'
                    ws[f'BI{i}'] = f'=AQ{i}+AS{i}+AU{i}+AW{i}'
                    ws[f'BJ{i}'] = f'=AG{i}+AK{i}+AS{i}+AO{i}+AW{i}'
                    ws[f'BK{i}'] = f'=U{i}'
                    ws[f'BL{i}'] = f'=Y{i}+AC{i}'
                    ws[f'BM{i}'] = f'=W{i}'
                    ws[f'BN{i}'] = f'=AA{i}+AE{i}'

                    cedula_emp = obtener_val_iloc(row, 2).replace('.0', '')
                    val_bo = ""
                    if cedula_emp in hist_dict and fecha_ori and pd.notna(fecha_ori):
                        sub_hist = hist_dict[cedula_emp]
                        match_ccco = sub_hist[
                            (sub_hist['Fecha_Inicio'] <= fecha_ori) &
                            (sub_hist['Fecha_Fin'] >= fecha_ori) &
                            (sub_hist['Centro_Costo'] != contrato_principal)
                        ]
                        if not match_ccco.empty:
                            val_bo = "CCCO"

                    ws[f'BO{i}'] = val_bo

                    val_bp_cc, val_bq_cc = "", ""
                    if cedula_emp in hist_dict and fecha_ori and pd.notna(fecha_ori):
                        sub_hist = hist_dict[cedula_emp]
                        match_hist = sub_hist[
                            (sub_hist['Fecha_Inicio'] <= fecha_ori) &
                            (sub_hist['Fecha_Fin'] >= fecha_ori)
                        ]
                        if not match_hist.empty:
                            val_bp_cc = match_hist.iloc[0]['Frente_Trabajo']
                            val_bq_cc = match_hist.iloc[0]['Centro_Costo']
                        else:
                            ult_registro = sub_hist.sort_values(by='Fecha_Inicio', ascending=False)
                            if not ult_registro.empty and fecha_ori > ult_registro.iloc[0]['Fecha_Fin']:
                                val_bp_cc = ult_registro.iloc[0]['Frente_Trabajo']
                                val_bq_cc = ult_registro.iloc[0]['Centro_Costo']

                    ws[f'BP{i}'] = val_bp_cc
                    ws[f'BQ{i}'] = val_bq_cc

                    celda_br, celda_bt = ws[f'BR{i}'], ws[f'BT{i}']
                    val_bs, val_bu = "", ""
                    datos_maestro = maestro_dict.get(cedula_emp, (pd.NaT, pd.NaT))
                    fecha_ing, fecha_ret = datos_maestro[0], datos_maestro[1]

                    if pd.notna(fecha_ing):
                        celda_br.value = fecha_ing.date()
                        celda_br.number_format = 'DD/MM/YYYY'
                        if fecha_ing > fecha_minima_valida and fecha_ori and pd.notna(fecha_ori) and fecha_ing > fecha_ori:
                            val_bs = "Revisar"
                    else:
                        celda_br.value = ""

                    if pd.notna(fecha_ret):
                        celda_bt.value = fecha_ret.date()
                        celda_bt.number_format = 'DD/MM/YYYY'
                        if fecha_ret > fecha_minima_valida and fecha_ori and pd.notna(fecha_ori) and fecha_ret < fecha_ori:
                            val_bu = "Retirado"
                    else:
                        celda_bt.value = ""

                    ws[f'BS{i}'] = val_bs
                    ws[f'BU{i}'] = val_bu

                    val_bv_aus, val_bw_nova = val_bf, ""
                    if val_bu == "Retirado":
                        val_bv_aus = "Retiro"
                    elif val_bs == "Revisar":
                        val_bv_aus = "Ingreso"
                    elif val_bo == "CCCO":
                        val_bv_aus = "CCCO"
                    elif f_val == "Descanso compensatorio":
                        val_bv_aus = "C"
                    elif cedula_emp in nova_dict and fecha_ori and pd.notna(fecha_ori):
                        sub_nova = nova_dict[cedula_emp]
                        match_nova = sub_nova[
                            (sub_nova['Fecha_Inicio'] <= fecha_ori) &
                            (sub_nova['Fecha_Fin'] >= fecha_ori)
                        ]
                        if not match_nova.empty:
                            val_bv_aus = match_nova.iloc[0]['Concepto']
                            val_bw_nova = match_nova.iloc[0]['Codigo_Novasoft']

                    ws[f'BV{i}'] = val_bv_aus
                    ws[f'BW{i}'] = val_bw_nova

                    val_bx_sic = ""
                    if cedula_emp in sic_dict and fecha_ori and pd.notna(fecha_ori):
                        sub_sic = sic_dict[cedula_emp]
                        match_sic = sub_sic[
                            (sub_sic['Fecha_Inicio'] <= fecha_ori) &
                            (sub_sic['Fecha_Fin'] >= fecha_ori)
                        ]
                        if not match_sic.empty:
                            val_bx_sic = match_sic.iloc[0]['Proceso']

                    ws[f'BX{i}'] = val_bx_sic

                    val_by_consolidado = val_bv_aus
                    if val_bv_aus == "Ausencia" and val_bx_sic != "":
                        val_by_consolidado = val_bx_sic

                    ws[f'BY{i}'] = val_by_consolidado

                    val_bz_comp = ""
                    if cedula_emp and fecha_ori and pd.notna(fecha_ori):
                        key_op = (cedula_emp, fecha_ori.date())
                        val_bz_comp = operativa_dict.get(key_op, "")

                    ws[f'BZ{i}'] = val_bz_comp

                    val_ca_aus_real = val_by_consolidado
                    if str(val_by_consolidado).strip().lower() in ["ausencia", "descanso", "p"] and str(val_bz_comp).strip().upper() == "C":
                        val_ca_aus_real = "C"

                    celda_ca = ws[f'CA{i}']
                    celda_ca.value = val_ca_aus_real

                    if cedula_emp and fecha_ori and pd.notna(fecha_ori):
                        marcaciones_ht_dict[(cedula_emp, fecha_ori.date())] = f'=IFERROR(ROUND(BC{i}*24,1),"")'

                    ca_val_clean = str(val_ca_aus_real).strip()
                    nombre_emp = f"{obtener_val_iloc(row, 0)} {obtener_val_iloc(row, 1)}".strip()
                    f_str = fecha_ori.strftime('%Y-%m-%d') if fecha_ori and pd.notna(fecha_ori) else str(val_e)

                    if ca_val_clean and ca_val_clean != "":
                        registros_novedades.append({
                            "Cédula": cedula_emp,
                            "Nombre": nombre_emp,
                            "Fecha": f_str,
                            "Novedad Ausentismo": ca_val_clean
                        })

                    if ca_val_clean.lower() == "ausencia":
                        celda_ca.fill = fill_ausencia_rojo
                        celda_ca.font = font_ausencia_blanco
                        celda_ca.alignment = Alignment(horizontal="center", vertical="center")
                        conteo_ausencias += 1

                        registros_ausencias.append({
                            "Identificador": cedula_emp,
                            "Nombres": nombre_emp,
                            "Fecha Ori": f_str,
                            "Dia": dia_nombre,
                            "Ausentismo Real": "Ausencia"
                        })
                    elif ca_val_clean.upper() == "P":
                        conteo_p += 1

                    for col_letra in ['AY', 'AZ', 'BA', 'BB', 'BC', 'BD', 'BE', 'BF', 'BG', 'BH', 'BI', 'BJ', 'BK', 'BL', 'BM', 'BN', 'BO', 'BP', 'BQ', 'BR', 'BS', 'BT', 'BU', 'BV', 'BW', 'BX', 'BY', 'BZ', 'CA']:
                        ws[f'{col_letra}{i}'].border = thin_border

                output = io.BytesIO()
                wb.save(output)
                output.seek(0)
                excel_salida = output

                # ── PROCESAMIENTO SEGUNDA FASE (PLANTILLA NÓMINA HTCC) ──
                kpi_total_proc_m2 = 0
                kpi_validos_m2 = 0
                kpi_revisar_m2 = 0
                htcc_buffer = None

                if file_nomina:
                    file_nomina.seek(0)
                    wb_htcc = load_workbook(file_nomina)
                    HOJA_LIBRO3 = wb_htcc.sheetnames[0]
                    ws_htcc = wb_htcc[HOJA_LIBRO3]

                    FILA_ENCABEZADO = 4
                    COL_INICIO_FECHAS = 12

                    col_periodo_libro3 = col_id_libro3 = col_concepto_libro3 = col_cantidad_libro3 = col_nombre_libro3 = None
                    for cell in ws_htcc[FILA_ENCABEZADO]:
                        if cell.value is None: continue
                        val = str(cell.value).strip().lower()
                        if val == "periodo": col_periodo_libro3 = cell.column
                        if "identificador" in val: col_id_libro3 = cell.column
                        if val == "nombre": col_nombre_libro3 = cell.column
                        if "nombre concepto" in val: col_concepto_libro3 = cell.column
                        if val == "cantidad": col_cantidad_libro3 = cell.column

                    if col_cantidad_libro3 is None:
                        col_cantidad_libro3 = 7

                    col_cantidad_letra = get_column_letter(col_cantidad_libro3)

                    col_fecha_libro3 = {}
                    max_col_fecha = COL_INICIO_FECHAS
                    for col_idx in range(COL_INICIO_FECHAS, ws_htcc.max_column + 1):
                        val = ws_htcc.cell(row=FILA_ENCABEZADO, column=col_idx).value
                        if val is None: continue
                        try:
                            fecha = pd.Timestamp(val).normalize()
                            col_fecha_libro3[fecha] = col_idx
                            if col_idx > max_col_fecha: max_col_fecha = col_idx
                        except: pass

                    COL_TOTAL_IDX = max_col_fecha + 1
                    COL_DIF_IDX = COL_TOTAL_IDX + 1

                    indice_filas = {}
                    mapa_nombres_ht = {}
                    for row in ws_htcc.iter_rows(min_row=FILA_ENCABEZADO + 1):
                        periodo_val = id_val = conc_val = nom_val = None
                        fila_num = row[0].row
                        for cell in row:
                            if cell.column == col_periodo_libro3: periodo_val = str(cell.value).strip().upper() if cell.value else None
                            if cell.column == col_id_libro3: id_val = limpiar_id_a_texto(cell.value)
                            if col_nombre_libro3 and cell.column == col_nombre_libro3: nom_val = str(cell.value).strip() if cell.value else None
                            if cell.column == col_concepto_libro3: conc_val = str(cell.value).strip().lower() if cell.value else None
                        if periodo_val and id_val and conc_val and id_val not in ("None", "nan", ""):
                            indice_filas[(periodo_val, id_val, conc_val)] = fila_num
                            if id_val and nom_val:
                                mapa_nombres_ht[id_val] = nom_val

                    cols_por_periodo = {p.upper(): [] for p in PERIODOS}
                    for f_dt, c_idx in col_fecha_libro3.items():
                        for per, (inicio, fin) in PERIODOS.items():
                            if inicio.normalize() <= f_dt <= fin.normalize():
                                cols_por_periodo[per.upper()].append(c_idx)
                                break

                    # Evaluación directa de la Hoja HT en memoria
                    for (periodo, id_str, conc_libro3), fila_excel in indice_filas.items():
                        cols_periodo = cols_por_periodo.get(periodo, [])
                        if not cols_periodo: continue

                        kpi_total_proc_m2 += 1
                        val_cant_nom = ws_htcc.cell(row=fila_excel, column=col_cantidad_libro3).value if col_cantidad_libro3 else 0
                        
                        suma_fechas_nom = 0.0
                        for c_f_idx in cols_periodo:
                            v_f = ws_htcc.cell(row=fila_excel, column=c_f_idx).value
                            try:
                                if v_f is not None:
                                    suma_fechas_nom += float(v_f)
                            except (ValueError, TypeError):
                                pass

                        try:
                            val_cant_num = float(val_cant_nom) if val_cant_nom is not None else 0.0
                        except (ValueError, TypeError):
                            val_cant_num = 0.0

                        dif_calculada = round(val_cant_num - suma_fechas_nom, 2)

                        if abs(dif_calculada) < 0.01:
                            kpi_validos_m2 += 1
                        else:
                            kpi_revisar_m2 += 1

                    htcc_buffer_out = io.BytesIO()
                    wb_htcc.save(htcc_buffer_out)
                    htcc_buffer_out.seek(0)
                    htcc_bytes = htcc_buffer_out.getvalue()
                else:
                    htcc_bytes = None

                st.session_state["procesado_exitoso"] = True
                st.session_state["excel_salida"] = excel_salida
                st.session_state["kpi_ausencias"] = conteo_ausencias
                st.session_state["kpi_p"] = conteo_p
                st.session_state["total_filas"] = total_filas
                st.session_state["df_novedades_res"] = pd.DataFrame(registros_novedades)
                st.session_state["kpi_total_proc_m2"] = kpi_total_proc_m2
                st.session_state["kpi_validos_m2"] = kpi_validos_m2
                st.session_state["kpi_revisar_m2"] = kpi_revisar_m2
                st.session_state["htcc_bytes"] = htcc_bytes

        except Exception as e:
            st.error(f"❌ Ocurrió un error durante el procesamiento: {str(e)}")

# ─── DESPLIEGUE PERSISTENTE DE RESULTADOS ─────────────────────────────────

if st.session_state.get("procesado_exitoso", False):
    st.success("✨ ¡Auditoría finalizada con éxito!")
    
    col_d1, col_d2 = st.columns(2)
    with col_d1:
        st.download_button(
            label="📥 Descargar Marcaciones Calculadas (.xlsx)",
            data=st.session_state["excel_salida"],
            file_name="Calculado_GeoVictoria_Casalimpia.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    with col_d2:
        if st.session_state.get("htcc_bytes"):
            st.download_button(
                label="📥 Descargar Plantilla HTCC Consolidada (.xlsx)",
                data=st.session_state["htcc_bytes"],
                file_name="HTCC_Consolidado_Casalimpia.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    st.markdown("<br><h3 style='color: #00529B; font-weight: 700;'>📊 Resumen Ejecutivo de Marcaciones y Ausentismos</h3>", unsafe_allow_html=True)
    kpi_col1, kpi_col2, kpi_col3, kpi_col4 = st.columns(4)

    total_proc = st.session_state["total_filas"]
    ausencias_val = st.session_state["kpi_ausencias"]
    erroneas_val = st.session_state["kpi_p"]
    exito_val = max(0, total_proc - (ausencias_val + erroneas_val))

    with kpi_col1:
        st.markdown(f"""
            <div class="kpi-card kpi-card-info">
                <div class="kpi-title kpi-title-info">📋 Total Registros Procesados</div>
                <div class="kpi-value kpi-value-info">{total_proc:,}</div>
            </div>
        """, unsafe_allow_html=True)

    with kpi_col2:
        st.markdown(f"""
            <div class="kpi-card kpi-card-success">
                <div class="kpi-title kpi-title-success">✅ Procesados con Éxito</div>
                <div class="kpi-value kpi-value-success">{exito_val:,}</div>
            </div>
        """, unsafe_allow_html=True)

    with kpi_col3:
        st.markdown(f"""
            <div class="kpi-card kpi-card-danger">
                <div class="kpi-title kpi-title-danger">🚨 Ausencias Reales (CA)</div>
                <div class="kpi-value kpi-value-danger">{ausencias_val:,}</div>
            </div>
        """, unsafe_allow_html=True)

    with kpi_col4:
        st.markdown(f"""
            <div class="kpi-card kpi-card-warning">
                <div class="kpi-title kpi-title-warning">⚠️ Marcaciones Incompletas</div>
                <div class="kpi-value kpi-value-warning">{erroneas_val:,}</div>
            </div>
        """, unsafe_allow_html=True)

    if st.session_state.get("kpi_total_proc_m2", 0) > 0:
        st.markdown("<br><h3 style='color: #00529B; font-weight: 700;'>📊 Resumen Ejecutivo de Auditoría de Nómina (HTCC)</h3>", unsafe_allow_html=True)
        kpi_m2_1, kpi_m2_2, kpi_m2_3 = st.columns(3)

        val_total_proc = st.session_state.get("kpi_total_proc_m2", 0)
        val_validos = st.session_state.get("kpi_validos_m2", 0)
        val_revisar = st.session_state.get("kpi_revisar_m2", 0)

        with kpi_m2_1:
            st.markdown(f"""
                <div class="kpi-card kpi-card-info">
                    <div class="kpi-title kpi-title-info">📋 Total Registros Nómina</div>
                    <div class="kpi-value kpi-value-info">{val_total_proc:,}</div>
                </div>
            """, unsafe_allow_html=True)

        with kpi_m2_2:
            st.markdown(f"""
                <div class="kpi-card kpi-card-success">
                    <div class="kpi-title kpi-title-success">✅ Registros Válidos</div>
                    <div class="kpi-value kpi-value-success">{val_validos:,}</div>
                </div>
            """, unsafe_allow_html=True)

        with kpi_m2_3:
            st.markdown(f"""
                <div class="kpi-card kpi-card-warning">
                    <div class="kpi-title kpi-title-warning">⚠️ Registros a Revisar</div>
                    <div class="kpi-value kpi-value-warning">{val_revisar:,}</div>
                </div>
            """, unsafe_allow_html=True)

    if "df_novedades_res" in st.session_state and not st.session_state["df_novedades_res"].empty:
        df_nov_full = st.session_state["df_novedades_res"]

        st.markdown("<br><h3 style='color: #00529B; font-weight: 700;'>📋 Detalle de Novedades y Ausentismos</h3>", unsafe_allow_html=True)
        
        alertas_disponibles = ["Todas las Alertas"] + sorted(list(df_nov_full["Novedad Ausentismo"].dropna().unique()))

        col_f1, col_f2 = st.columns([2, 1])
        with col_f1:
            filtro_alerta_sel = st.selectbox(
                "Filtrar por Tipo de Alerta / Novedad:",
                options=alertas_disponibles,
                index=0
            )

        if filtro_alerta_sel != "Todas las Alertas":
            df_nov_display = df_nov_full[df_nov_full["Novedad Ausentismo"] == filtro_alerta_sel]
        else:
            df_nov_display = df_nov_full

        with col_f2:
            st.markdown("<div style='padding-top: 28px;'></div>", unsafe_allow_html=True)
            bytes_excel_tabla = df_a_excel_bytes(df_nov_display, sheet_name="Novedades_Filtradas")
            st.download_button(
                label="📊 Exportar Tabla a Excel (.xlsx)",
                data=bytes_excel_tabla,
                file_name=f"Detalle_Novedades_{filtro_alerta_sel.replace(' ', '_')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        try:
            df_styled = df_nov_display.style.map(estilo_etiqueta_ausentismo, subset=["Novedad Ausentismo"])
        except AttributeError:
            df_styled = df_nov_display.style.applymap(estilo_etiqueta_ausentismo, subset=["Novedad Ausentismo"])

        st.dataframe(
            df_styled,
            use_container_width=True,
            hide_index=True
        )
