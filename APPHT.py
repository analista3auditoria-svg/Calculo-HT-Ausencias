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
    page_title="Auditor Corporativo - Casalimpia", 
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
                <h1>Auditor TS Corporativo</h1>
                <p>Plataforma Integral de Procesamiento, Auditoría de Tiempos y Análisis de Compensatorios</p>
            </div>
        </div>
    </div>
""", unsafe_allow_html=True)

# ─── NAVEGACIÓN ENTRE MÓDULOS DE LA APLICACIÓN ────────────────────────────
st.sidebar.markdown("## 🧭 Navegación de Módulos")
modulo_seleccionado = st.sidebar.radio(
    "Selecciona el Módulo de Trabajo:",
    options=["1. Auditor TS & GeoVictoria", "2. Análisis Auditoría TS"],
    index=0
)
st.sidebar.markdown("---")

# ==============================================================================
# MÓDULO 1: AUDITOR TS & GEOVICTORIA
# ==============================================================================

if modulo_seleccionado == "1. Auditor TS & GeoVictoria":
    
    # --- FUNCIONES DE APOYO DEL MÓDULO 1 ---
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

    def procesar_plantilla_geovictoria(
        file_entrada, sheet_entrada, sheet_festivos,
        file_operativa, sheet_operativa,
        file_novasoft, sheet_novasoft,
        file_sic, sheet_sic,
        file_maestro, sheet_maestro,
        file_historial, sheet_historial,
        file_supernumerario, sheet_supernumerario,
        contrato_principal,
        fecha_ini_sup, fecha_fin_sup
    ):
        df_marc_raw = pd.read_excel(file_entrada, sheet_name=sheet_entrada)
        
        operativa_dict = {}
        if file_operativa:
            try:
                excel_op = pd.ExcelFile(file_operativa)
                target_sheet = sheet_operativa
                for name in excel_op.sheet_names:
                    if name.strip().upper() == sheet_operativa.strip().upper():
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
                st.warning(f"⚠️ No se pudo procesar la hoja '{sheet_operativa}' de la Base Operativa: {e}")

        df_super_filtrado = pd.DataFrame()
        if file_supernumerario:
            try:
                excel_sup = pd.ExcelFile(file_supernumerario)
                target_sheet_sup = sheet_supernumerario
                for name in excel_sup.sheet_names:
                    if name.strip().upper() == sheet_supernumerario.strip().upper():
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
                st.warning(f"⚠️ No se pudo procesar la BD Supernumerario (hoja '{sheet_supernumerario}'): {e}")

        df_nova = pd.read_excel(file_novasoft, sheet_name=sheet_novasoft) if file_novasoft else pd.DataFrame()
        df_sic = pd.read_excel(file_sic, sheet_name=sheet_sic) if file_sic else pd.DataFrame()
        df_maestro = pd.read_excel(file_maestro, sheet_name=sheet_maestro) if file_maestro else pd.DataFrame()
        df_hist = pd.read_excel(file_historial, sheet_name=sheet_historial) if file_historial else pd.DataFrame()

        set_festivos = set()
        if file_entrada:
            try:
                df_festivos = pd.read_excel(file_entrada, sheet_name=sheet_festivos)
                if not df_festivos.empty:
                    fechas_fest = pd.to_datetime(df_festivos.iloc[:, 0], dayfirst=True, errors='coerce').dropna()
                    set_festivos = set(fechas_fest.dt.date)
            except Exception as e:
                st.warning(f"⚠️ Nota: No se pudo cargar la hoja '{sheet_festivos}' ({e}). Se continuará sin marcar festivos.")

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
        ws = wb[sheet_entrada]
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

        progress_bar = st.progress(0)
        status_text = st.empty()
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

            pct = (idx + 1) / total_filas
            progress_bar.progress(pct)
            status_text.caption(f"⚡ Procesando fila {idx + 1} de {total_filas} ({int(pct*100)}%)")

        nombre_hoja_nov = "Novedades Ausentismo"
        if nombre_hoja_nov in wb.sheetnames:
            ws_nov = wb[nombre_hoja_nov]
            ws_nov.delete_rows(1, ws_nov.max_row + 1)
        else:
            ws_nov = wb.create_sheet(title=nombre_hoja_nov)
        
        ws_nov.views.sheetView[0].showGridLines = True

        encabezados_nov = ["Cédula", "Nombre", "Fecha", "Novedad Ausentismo"]
        bg_azul_header = PatternFill(start_color="00529B", end_color="00529B", fill_type="solid")
        font_header = Font(name="Calibri", size=11, bold=True, color="FFFFFF")

        fill_nov_rosado = PatternFill(start_color="FCE8E6", end_color="FCE8E6", fill_type="solid")
        font_nov_rojo = Font(name="Calibri", size=10, bold=True, color="991B1B")

        for col_idx, text_h in enumerate(encabezados_nov, start=1):
            c = ws_nov.cell(row=1, column=col_idx, value=text_h)
            c.fill = bg_azul_header
            c.font = font_header
            c.alignment = Alignment(horizontal="center", vertical="center")
            c.border = thin_border

        for row_idx, r_nov in enumerate(registros_novedades, start=2):
            c1 = ws_nov.cell(row=row_idx, column=1, value=r_nov["Cédula"])
            c2 = ws_nov.cell(row=row_idx, column=2, value=r_nov["Nombre"])
            c3 = ws_nov.cell(row=row_idx, column=3, value=r_nov["Fecha"])
            c4 = ws_nov.cell(row=row_idx, column=4, value=r_nov["Novedad Ausentismo"])

            c1.alignment = Alignment(horizontal="center", vertical="center")
            c2.alignment = Alignment(horizontal="left", vertical="center")
            c3.alignment = Alignment(horizontal="center", vertical="center")
            c4.alignment = Alignment(horizontal="center", vertical="center")

            c4.fill = fill_nov_rosado
            c4.font = font_nov_rojo

            for c_tmp in [c1, c2, c3, c4]:
                c_tmp.border = thin_border

        if "Ausencias" in wb.sheetnames:
            ws_aus = wb["Ausencias"]
            ws_aus.delete_rows(1, ws_aus.max_row + 1)
        else:
            ws_aus = wb.create_sheet(title="Ausencias")
        
        ws_aus.views.sheetView[0].showGridLines = True

        ws_aus["A1"] = "Identificador"
        ws_aus["B1"] = "Nombres"
        ws_aus["C1"] = "Fecha Ori"
        ws_aus["D1"] = "Dia"
        ws_aus["E1"] = "Ausentismo Real"

        bg_azul_header_aus = PatternFill(start_color="1B5E82", end_color="1B5E82", fill_type="solid")
        bg_verde_header = PatternFill(start_color="1E4620", end_color="1E4620", fill_type="solid")

        for col in ["A1", "B1", "C1", "D1"]:
            c = ws_aus[col]
            c.fill = bg_azul_header_aus
            c.font = font_header
            c.alignment = Alignment(horizontal="center", vertical="center")
            c.border = thin_border

        c_e1 = ws_aus["E1"]
        c_e1.fill = bg_verde_header
        c_e1.font = font_header
        c_e1.alignment = Alignment(horizontal="center", vertical="center")
        c_e1.border = thin_border

        df_aus_det = pd.DataFrame(registros_ausencias)
        for row_idx, r in enumerate(registros_ausencias, start=2):
            ws_aus[f"A{row_idx}"] = r["Identificador"]
            ws_aus[f"B{row_idx}"] = r["Nombres"]
            ws_aus[f"C{row_idx}"] = r["Fecha Ori"]
            ws_aus[f"D{row_idx}"] = r["Dia"]
            
            c_aus_val = ws_aus[f"E{row_idx}"]
            c_aus_val.value = r["Ausentismo Real"]
            c_aus_val.fill = fill_ausencia_rojo
            c_aus_val.font = font_ausencia_blanco
            c_aus_val.alignment = Alignment(horizontal="center", vertical="center")

            for col_l in ["A", "B", "C", "D", "E"]:
                ws_aus[f"{col_l}{row_idx}"].border = thin_border

        ws_aus["H1"] = "Identificador"
        ws_aus["I1"] = "Cantidad"

        for col in ["H1", "I1"]:
            c = ws_aus[col]
            c.fill = bg_azul_header
            c.font = font_header
            c.alignment = Alignment(horizontal="center", vertical="center")
            c.border = thin_border

        if not df_aus_det.empty:
            df_resumen = df_aus_det.groupby("Identificador").size().reset_index(name="Cantidad")
            df_resumen = df_resumen.sort_values(by="Cantidad", ascending=False)

            for r_idx, r in enumerate(df_resumen.itertuples(), start=2):
                c_h = ws_aus[f"H{r_idx}"]
                c_i = ws_aus[f"I{r_idx}"]
                
                c_h.value = r.Identificador
                c_i.value = r.Cantidad
                
                c_h.alignment = Alignment(horizontal="center", vertical="center")
                c_i.alignment = Alignment(horizontal="right", vertical="center")
                c_h.border = thin_border
                c_i.border = thin_border

        if not df_super_filtrado.empty:
            if "Supernumerario" in wb.sheetnames:
                ws_sup = wb["Supernumerario"]
                ws_sup.delete_rows(1, ws_sup.max_row + 1)
            else:
                ws_sup = wb.create_sheet(title="Supernumerario")
            
            ws_sup.views.sheetView[0].showGridLines = True
            
            for r_idx, r in enumerate(dataframe_to_rows(df_super_filtrado, index=False, header=True), start=1):
                ws_sup.append(r)
                for c_idx in range(1, len(r) + 1):
                    cell = ws_sup.cell(row=r_idx, column=c_idx)
                    cell.border = thin_border
                    if r_idx == 1:
                        cell.fill = PatternFill(start_color="00529B", end_color="00529B", fill_type="solid")
                        cell.font = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
                        cell.alignment = Alignment(horizontal="center", vertical="center")

            ws_sup.cell(row=1, column=23, value="HT")
            ws_sup.cell(row=1, column=23).fill = PatternFill(start_color="00529B", end_color="00529B", fill_type="solid")
            ws_sup.cell(row=1, column=23).font = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
            ws_sup.cell(row=1, column=23).alignment = Alignment(horizontal="center", vertical="center")

            max_r_sup = ws_sup.max_row
            for r_sup in range(2, max_r_sup + 1):
                cell_f = ws_sup.cell(row=r_sup, column=2)
                f_val_sup = cell_f.value
                dt_sup = None
                
                if pd.notna(f_val_sup):
                    try:
                        dt_sup = pd.to_datetime(f_val_sup, dayfirst=True, errors='coerce')
                        if pd.notna(dt_sup):
                            cell_f.value = dt_sup.date()
                            cell_f.number_format = 'DD/MM/YYYY'
                    except Exception:
                        pass

                cell_ced = ws_sup.cell(row=r_sup, column=3)
                ced_sup = str(cell_ced.value).strip().replace('.0', '') if cell_ced.value else ""

                cell_ht = ws_sup.cell(row=r_sup, column=23)
                cell_ht.border = thin_border
                cell_ht.alignment = Alignment(horizontal="center", vertical="center")

                if ced_sup and dt_sup and pd.notna(dt_sup):
                    key_sup = (ced_sup, dt_sup.date())
                    ht_encontrado = marcaciones_ht_dict.get(key_sup, "")
                    cell_ht.value = ht_encontrado
                else:
                    cell_ht.value = ""

        status_text.empty()
        
        output = io.BytesIO()
        wb.save(output)
        output.seek(0)
        return output, conteo_ausencias, conteo_p, total_filas, pd.DataFrame(registros_novedades)

    # ── CARGA DE ARCHIVOS ──
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

    st.sidebar.markdown("## ⚙️ Parámetros")
    contrato_principal = st.sidebar.selectbox("Contrato / CC Principal", options=lista_cc, index=0)

    st.sidebar.markdown("---")
    st.sidebar.markdown("### 📅 Filtro Rango de Fechas")
    fecha_ini_sup = st.sidebar.date_input("Fecha Inicial", value=datetime.date(2026, 2, 1))
    fecha_fin_sup = st.sidebar.date_input("Fecha Final", value=datetime.date(2026, 2, 28))

    st.sidebar.markdown("---")
    st.sidebar.markdown("""
    <div style="background-color: #f0f7ff; padding: 12px; border-radius: 8px; border-left: 4px solid #00529B;">
        <small style="color: #00529B; font-weight: 600;">💡 Instrucciones</small><br>
        <small style="color: #475569;">1. Carga el archivo <b>6. BBDD Historia laboral de empleados</b>.<br>2. Selecciona el Centro de Costos.<br>3. Ajusta las fechas y ejecuta la auditoría.</small>
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

    if st.button("⚡ Ejecutar Auditoría TS y Procesar Marcaciones", type="primary"):
        if not file_entrada:
            st.error("⚠️ Es obligatorio cargar el archivo principal de Marcaciones (GeoVictoria).")
        elif not contrato_principal:
            st.error("⚠️ Por favor, selecciona el Contrato / Centro de Costo Principal en el panel izquierdo.")
        else:
            try:
                with st.spinner("Procesando marcaciones y generando reporte consolidado de Novedades de Ausentismo..."):
                    excel_salida, kpi_ausencias, kpi_p, total_filas, df_novedades_res = procesar_plantilla_geovictoria(
                        file_entrada, hoja_entrada, sheet_festivos=hoja_festivos,
                        file_operativa=file_operativa, sheet_operativa=hoja_operativa,
                        file_novasoft=file_novasoft, sheet_novasoft=hoja_novasoft,
                        file_sic=file_sic, sheet_sic=hoja_sic,
                        file_maestro=file_maestro, sheet_maestro=hoja_maestro,
                        file_historial=file_historial, sheet_historial=hoja_historial,
                        file_supernumerario=file_supernumerario, sheet_supernumerario=hoja_supernumerario,
                        contrato_principal=contrato_principal,
                        fecha_ini_sup=fecha_ini_sup, fecha_fin_sup=fecha_fin_sup
                    )

                st.session_state["procesado_exitoso"] = True
                st.session_state["excel_salida"] = excel_salida
                st.session_state["kpi_ausencias"] = kpi_ausencias
                st.session_state["kpi_p"] = kpi_p
                st.session_state["total_filas"] = total_filas
                st.session_state["df_novedades_res"] = df_novedades_res

            except Exception as e:
                st.error(f"❌ Ocurrió un error durante el procesamiento: {str(e)}")

    if st.session_state.get("procesado_exitoso", False):
        st.success("✨ ¡Auditoría finalizada con éxito!")
        
        st.download_button(
            label="📥 Descargar Resultado Calculado Completo (Excel Completo)",
            data=st.session_state["excel_salida"],
            file_name="Calculado_GeoVictoria_Casalimpia.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        st.markdown("<br><h3 style='color: #00529B; font-weight: 700;'>📊 Resumen Ejecutivo de Auditoría</h3>", unsafe_allow_html=True)
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


# ==============================================================================
# MÓDULO 2: ANÁLISIS AUDITORÍA TS
# ==============================================================================

elif modulo_seleccionado == "2. Análisis Auditoría TS":

    st.markdown("<h2 style='color: #00529B; font-weight: 700;'>📊 Módulo de Análisis Auditoría TS & Compensatorios</h2>", unsafe_allow_html=True)

    # ── Configuración Inicial Interna del Módulo 2 ──
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
        
        if pd.isna(valor): 
            return ""
        s = str(valor).strip()
        if s.lower() in ("nan", "none", "0", ""): 
            return ""
        if s.endswith(".0"): 
            s = s[:-2]
        try:
            return str(int(float(s)))
        except:
            return s

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
            if pd.notna(d):
                return f"{d.day}-{MESES_ES[d.month]}"
        except: pass
        
        s = str(val_fecha).strip().lower()
        m = re.search(r'(\d{1,2})[-_\s/]+([a-z]{3})', s)
        if m:
            return f"{int(m.group(1))}-{m.group(2)}"
        return s

    def obtener_dt_fecha(val_fecha):
        if pd.isna(val_fecha) or val_fecha is None: return None
        if isinstance(val_fecha, pd.Timestamp):
            return val_fecha.normalize()
        try:
            d = pd.to_datetime(val_fecha, errors='coerce')
            if pd.notna(d):
                return d.normalize()
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

    # ── Paso 1: Carga de Archivos ──
    st.header("📁 1. Carga de Archivos Base")

    col_file1, col_file2, col_file3 = st.columns(3)
    with col_file1:
        archivo_cargado = st.file_uploader("Subir plantilla de Asistencia (.xlsx)", type=["xlsx"], key="u_asistencia")
    with col_file2:
        archivo_htcc = st.file_uploader("Subir plantilla de Consolidación HTCC (.xlsx)", type=["xlsx"], key="u_htcc")
    with col_file3:
        archivo_operativo = st.file_uploader("Subir Reporte Operativo (.xlsx)", type=["xlsx"], key="u_operativo")

    if archivo_cargado is not None and archivo_htcc is not None and archivo_operativo is not None:
        xl = pd.ExcelFile(archivo_cargado)
        hojas_disponibles = xl.sheet_names
        
        xl_htcc = pd.ExcelFile(archivo_htcc)
        hojas_htcc = xl_htcc.sheet_names

        xl_operativo = pd.ExcelFile(archivo_operativo)
        hojas_operativo = xl_operativo.sheet_names
        
        col_sheet1, col_sheet2, col_sheet3 = st.columns(3)
        with col_sheet1:
            HOJA_ENTRADA = st.selectbox("Seleccione la hoja de marcación:", hojas_disponibles, key="s_hoja1")
        with col_sheet2:
            HOJA_LIBRO3 = st.selectbox("Seleccione la hoja de destino en HTCC:", hojas_htcc, key="s_hoja2")
        with col_sheet3:
            HOJA_OPERATIVO = st.selectbox("Seleccione la hoja del reporte operativo:", hojas_operativo, key="s_hoja3")
        
        st.header("📅 2. Parámetros de Filtrado y Fechas")
        fecha_inicio_input = st.date_input("Fecha Inicial de Semanas para Análisis de Compensatorios", value=datetime.date(2026, 2, 15))
        
        if "procesado_m2" not in st.session_state:
            st.session_state.procesado_m2 = False
            st.session_state.output_bytes_m2 = None
            st.session_state.htcc_bytes_m2 = None
            st.session_state.listado_m2 = None
            st.session_state.resumen_m2 = None
            st.session_state.df_c_m2 = None
            st.session_state.resultados_c_m2 = None

        if st.button("🚀 Procesar Información y Generar Análisis", type="primary"):
            with st.spinner("Procesando datos y estructurando archivos de Excel..."):
                try:
                    df_origen = pd.read_excel(archivo_cargado, sheet_name=HOJA_ENTRADA)
                    df = df_origen.copy()
                    
                    df_operativo = pd.read_excel(archivo_operativo, sheet_name=HOJA_OPERATIVO)
                    
                    def extraer_nombre_limpio(dataframe):
                        cols_map = {str(c).strip().lower(): c for c in dataframe.columns}
                        
                        def buscar_c(patron):
                            for k, orig in cols_map.items():
                                if re.search(patron, k, re.IGNORECASE):
                                    return orig
                            return None

                        c_nom_comp  = buscar_c(r'nombre.*completo|^empleado|^usuario')
                        c_nombres   = buscar_c(r'^nombre|^nombres')
                        c_apellidos = buscar_c(r'^apellido|^apellidos')

                        if c_nombres and c_apellidos and c_nombres != c_apellidos:
                            n = dataframe[c_nombres].fillna('').astype(str).str.strip().replace('nan', '')
                            a = dataframe[c_apellidos].fillna('').astype(str).str.strip().replace('nan', '')
                            res = (n + " " + a).str.strip()
                            if res.str.len().sum() > 0:
                                return res

                        col_unica = c_nom_comp or c_nombres or c_apellidos
                        if col_unica:
                            res = dataframe[col_unica].fillna('').astype(str).str.strip().replace('nan', '')
                            if res.str.len().sum() > 0:
                                return res

                        return pd.Series(["Sin Nombre"] * len(dataframe))

                    df['Nombre'] = extraer_nombre_limpio(df)
                    df_origen['Nombre'] = df['Nombre']

                    col_map = {str(c).strip().lower(): c for c in df.columns}
                    def buscar_col(patrones):
                        for pat in patrones:
                            for k, orig in col_map.items():
                                if re.search(pat, k, re.IGNORECASE): return orig
                        return None

                    col_id     = buscar_col([r'^identificador$'])
                    col_fecha  = buscar_col([r'^fecha ori', r'^fecha$'])
                    col_dia    = buscar_col([r'^dia$'])
                    col_ht     = buscar_col([r'^ht$'])
                    col_ausent = buscar_col([r'^ausentismo$'])

                    col_ciudad = buscar_col([r'^ciudad$', r'municipio', r'regional'])
                    col_sede   = buscar_col([r'^sede$', r'^grupo$', r'^contrato$', r'ubicacion', r'lugar'])

                    df['Ciudad'] = df[col_ciudad].fillna('No Registra').astype(str).str.strip() if col_ciudad else 'No Registra'
                    df['Sede']   = df[col_sede].fillna('No Registra').astype(str).str.strip() if col_sede else 'No Registra'
                    
                    df_origen['Ciudad'] = df['Ciudad']
                    df_origen['Sede']   = df['Sede']

                    col_rno   = buscar_col([r'recargo nocturno'])
                    col_rdc   = buscar_col([r'dominical compensado'])
                    col_rdnc  = buscar_col([r'dominical no compensado'])
                    col_rf    = buscar_col([r'recargo festivo'])
                    col_hed   = buscar_col([r'extras diurnas.*1[\.,\s]?25'])
                    col_hedf  = buscar_col([r'extra diurna dom'])
                    col_hen   = buscar_col([r'extras nocturnas.*1[\.,\s]?75'])
                    col_hendf = buscar_col([r'dominical o festiva nocturna'])

                    criticas = {'Identificador':col_id,'Fecha':col_fecha,'Dia':col_dia,'HT':col_ht,'Ausentismo':col_ausent}
                    faltantes = [k for k,v in criticas.items() if v is None]
                    if faltantes:
                        st.error(f"❌ Columnas críticas no detectadas en la hoja elegida: {faltantes}.")
                        st.stop()

                    df['FechaReal'] = pd.to_datetime(df[col_fecha], errors='coerce')
                    df['FechaCorta'] = df['FechaReal'].apply(lambda d: f"{d.day}-{MESES_ES[d.month]}" if pd.notna(d) else 'sin-fecha')
                    
                    def limpiar_dia(texto):
                        if pd.isna(texto): return ''
                        for p in str(texto).strip().split():
                            pl = p.lower()
                            if pl in {'lunes','martes','miércoles','miercoles','jueves','viernes','sábado','sabado','domingo'}:
                                return {'miercoles': 'miércoles', 'sabado': 'sábado'}.get(pl, pl)
                        return str(texto).strip().lower()

                    df['dia']      = df[col_dia].apply(limpiar_dia)
                    df['_festivo'] = df[col_dia].apply(lambda x: 'festivo' in str(x).strip().lower())

                    aus = df[col_ausent]
                    df['_HTn']    = [celda(hhmm_a_decimal(h), a) for h, a in zip(df[col_ht], aus)]
                    df['_RNn']    = [celda(num_limpio(v), a) for v, a in zip(df[col_rno], aus)]
                    df['_RDCn']   = [celda(num_limpio(v), a) for v, a in zip(df[col_rdc], aus)]
                    df['_RDNCn']  = [celda(num_limpio(v), a) for v, a in zip(df[col_rdnc], aus)]
                    df['_RFn']    = [celda(num_limpio(v), a) for v, a in zip(df[col_rf], aus)]
                    df['_HEDn']   = [celda(num_limpio(v), a) for v, a in zip(df[col_hed], aus)]
                    df['_HEDFn']  = [celda(num_limpio(v), a) for v, a in zip(df[col_hedf], aus)]
                    df['_HENn']   = [celda(num_limpio(v), a) for v, a in zip(df[col_hen], aus)]
                    df['_HENDFn'] = [celda(num_limpio(v), a) for v, a in zip(df[col_hendf], aus)]

                    CONCEPTOS_COLS = ['_HTn','_RNn','_RDCn','_RDNCn','_RFn','_HEDn','_HEDFn','_HENn','_HENDFn']
                    ETIQUETAS = {
                        '_HTn': 'HT Normales', '_RNn': 'Recargo Nocturno 0.35%', '_RDCn': 'Recargo Dominical Compensado',
                        '_RDNCn': 'Recargo Dominical No Compensado', '_RFn': 'Recargo Festivo', '_HEDn': 'Horas Extras Diurnas 1.25%',
                        '_HEDFn': 'Hora Extra Diurna Dom/Fest', '_HENn': 'Horas Extras Nocturnas 1.75%', '_HENDFn': 'Hora Extra Dominical o Festiva Nocturna'
                    }
                    ORDEN_CONCEPTOS = {v: i+1 for i, v in enumerate(ETIQUETAS.values())}
                    COLS_FIJAS = [c for c in [col_id, 'FechaReal', 'FechaCorta', 'dia', '_festivo'] if c in df.columns]

                    unpivoted = df[COLS_FIJAS + CONCEPTOS_COLS].melt(id_vars=COLS_FIJAS, value_vars=CONCEPTOS_COLS, var_name='_col', value_name='_valor').reset_index(drop=True)
                    unpivoted = unpivoted[~(unpivoted['_valor'].isna() & (unpivoted['_col'] != '_HTn'))].reset_index(drop=True)
                    unpivoted['Concepto'] = unpivoted['_col'].map(ETIQUETAS)
                    unpivoted['_Orden']   = unpivoted['Concepto'].map(ORDEN_CONCEPTOS).fillna(10).astype(int)
                    unpivoted = unpivoted.drop(columns='_col').sort_values([col_id, '_Orden', 'FechaReal']).reset_index(drop=True)

                    mapa_fechas = unpivoted[['FechaReal', 'FechaCorta', 'dia', '_festivo']].drop_duplicates(subset=['FechaCorta']).dropna(subset=['FechaReal']).sort_values('FechaReal')
                    fechas_unicas  = mapa_fechas['FechaCorta'].tolist()
                    mapa_dia       = dict(zip(mapa_fechas['FechaCorta'], mapa_fechas['dia']))
                    mapa_festivo   = dict(zip(mapa_fechas['FechaCorta'], mapa_fechas['_festivo']))

                    pivotados = []
                    for nombre in ETIQUETAS.values():
                        sub = unpivoted[unpivoted['Concepto'] == nombre].sort_values('FechaReal').drop_duplicates(subset=[col_id, 'FechaCorta'], keep='first')
                        piv = sub.pivot(index=[col_id], columns='FechaCorta', values='_valor').reset_index()
                        cols_f = [c for c in piv.columns if c != col_id]
                        piv[cols_f] = piv[cols_f].fillna(0)
                        piv['Concepto'] = nombre
                        piv['_Orden']   = ORDEN_CONCEPTOS[nombre]
                        pivotados.append(piv)

                    pivotado = pd.concat(pivotados, ignore_index=True)
                    cols_fecha_ok = [f for f in fechas_unicas if f in pivotado.columns]
                    pivotado = pivotado[[c for c in pivotado.columns if c not in fechas_unicas] + cols_fecha_ok].fillna(0)

                    mapa_nombres = df_origen[[col_id, 'Nombre']].drop_duplicates(subset=[col_id]).copy()
                    mapa_nombres[col_id] = limpiar_id_a_texto(mapa_nombres[col_id])

                    final = pivotado.drop(columns=['_Orden'], errors='ignore').rename(columns={col_id: 'Identificador'})
                    final['Identificador'] = limpiar_id_a_texto(final['Identificador'])
                    final = pd.merge(final, mapa_nombres, left_on='Identificador', right_on=col_id, how='left')
                    if col_id != 'Identificador' and col_id in final.columns: final = final.drop(columns=[col_id])

                    COLS_INICIO_PRES = [c for c in ['Identificador', 'Nombre', 'Concepto'] if c in final.columns]
                    date_cols = [c for c in final.columns if c not in COLS_INICIO_PRES]
                    final = final[COLS_INICIO_PRES + date_cols]
                    final['_ord_c']  = final['Concepto'].map(ORDEN_CONCEPTOS).fillna(10).astype(int)
                    final['_ord_id'] = pd.to_numeric(final['Identificador'], errors='coerce')
                    final = final.sort_values(['_ord_id', '_ord_c']).reset_index(drop=True).drop(columns=['_ord_c', '_ord_id'])
                    final['Identificador'] = final['Identificador'].astype(str)
                    final[date_cols] = final[date_cols].fillna('')

                    output_buffer = io.BytesIO()
                    wb = Workbook()
                    ws = wb.active
                    ws.title = 'Reporte_Horizontal'

                    C_NAVY, C_ROJO, C_VRD_DOM, C_VRD_FES, C_VRD_HDR, C_FES_HDR, C_GRIS, C_NARANJA, C_RDC, C_HE, C_AMBAR, C_AMARILLO = '1F4E79', 'FF0000', 'C6EFCE', '92D050', '538135', '375623', 'D9D9D9', 'FF8C00', 'FFCCCC', 'BDD7EE', 'F4B942', 'FFEB9C'
                    thin = Side(style='thin', color='CCCCCC')
                    brd  = Border(left=thin, right=thin, top=thin, bottom=thin)

                    all_cols = final.columns.tolist()
                    for ci, col in enumerate(all_cols, start=1):
                        c = ws.cell(row=1, column=ci, value=col)
                        c.font, c.fill, c.border = Font(name='Arial', bold=True, color='FFFFFF', size=9), PatternFill('solid', fgColor=C_NAVY), brd
                        c.alignment = Alignment(horizontal='center', vertical='center', text_rotation=90 if col in date_cols else 0, wrap_text=(col not in date_cols))
                    ws.row_dimensions[1].height = 55

                    for ci, col in enumerate(all_cols, start=1):
                        dia, es_fest = mapa_dia.get(col, ''), mapa_festivo.get(col, False)
                        texto_dia = f"{dia} Fest" if es_fest and col in date_cols else dia if col in date_cols else ''
                        bg, fg = C_FES_HDR if es_fest else C_VRD_HDR if 'domingo' in dia else C_GRIS, 'FFFFFF' if (es_fest or 'domingo' in dia) else '000000'
                        c = ws.cell(row=2, column=ci, value=texto_dia)
                        c.font, c.fill, c.alignment, c.border = Font(name='Arial', bold=True, color=fg, size=8), PatternFill('solid', fgColor=bg), Alignment(horizontal='center', vertical='center'), brd
                    ws.row_dimensions[2].height = 15

                    for ri in range(len(final)):
                        concepto = str(final.iloc[ri]['Concepto']).strip()
                        for ci, col in enumerate(all_cols, start=1):
                            val = final.iloc[ri][col]
                            val_str = str(val).strip() if val is not None else ''
                            abrev = resolver_ausencia(val_str) if col in date_cols and val_str else None
                            
                            if col not in date_cols:
                                if concepto == 'Recargo Dominical No Compensado': bg, fg = C_NARANJA, 'FFFFFF'
                                elif concepto == 'Recargo Dominical Compensado': bg, fg = C_RDC, '000000'
                                elif concepto in {'Horas Extras Diurnas 1.25%', 'Hora Extra Diurna Dom/Fest', 'Horas Extras Nocturnas 1.75%', 'Hora Extra Dominical o Festiva Nocturna'}: bg, fg = C_HE, '000000'
                                else: bg, fg = ('F2F2F2' if ri % 2 == 0 else 'FFFFFF'), '000000'
                            elif col in date_cols and abrev in AUSENCIAS_ROJO_SET: bg, fg = C_ROJO, 'FFFFFF'
                            elif col in date_cols and abrev in AUSENCIAS_AMARILLO_SET: bg, fg = C_AMARILLO, '000000'
                            elif col in date_cols and val_str.lower() == 'c': bg, fg = C_AMBAR, '000000'
                            elif mapa_festivo.get(col, False): bg, fg = C_VRD_FES, '000000'
                            elif 'domingo' in mapa_dia.get(col, ''): bg, fg = C_VRD_DOM, '000000'
                            else: bg, fg = ('FFFFFF' if ri % 2 == 0 else 'F5F5F5'), '000000'

                            c = ws.cell(row=ri + 3, column=ci, value=val)
                            if col in date_cols and isinstance(val, (int, float)): c.number_format = '#,##0.00'
                            c.font, c.fill, c.alignment, c.border = Font(name='Arial', size=9, bold=(col not in date_cols), color=fg), PatternFill('solid', fgColor=bg), Alignment(horizontal='center' if col in date_cols else 'left', vertical='center'), brd
                        ws.row_dimensions[ri + 3].height = 14

                    ws.column_dimensions['A'].width, ws.column_dimensions['B'].width = 14, 42
                    for ci in range(3, len(all_cols) + 1): ws.column_dimensions[get_column_letter(ci)].width = 6
                    ws.freeze_panes, ws.auto_filter.ref = 'D3', f"A2:{get_column_letter(final.shape[1])}{final.shape[0] + 2}"
                    wb.save(output_buffer)

                    mask_aus = (df[col_ausent].notna() & (df[col_ausent].astype(str).str.strip().str.lower() == 'ausencia'))
                    df_aus = df[mask_aus][[col_id, 'Nombre', 'Ciudad', 'Sede', 'FechaReal', col_ausent]].copy()
                    df_aus['Fecha'] = df_aus['FechaReal'].dt.strftime('%d/%m/%Y')
                    df_aus['_ord']  = pd.to_numeric(df_aus[col_id], errors='coerce')
                    df_aus = df_aus.sort_values(['_ord', 'FechaReal']).reset_index(drop=True)
                    
                    listado = df_aus[[col_id, 'Nombre', 'Ciudad', 'Sede', 'Fecha', col_ausent]].copy()
                    listado.columns = ['Identificador', 'Nombre', 'Ciudad', 'Sede', 'Fecha', 'Novedad']
                    
                    resumen = listado.groupby(['Identificador', 'Nombre', 'Ciudad', 'Sede'], as_index=False).size().rename(columns={'size': 'Cantidad'}).sort_values('Cantidad', ascending=False).reset_index(drop=True)

                    wb2 = load_workbook(output_buffer)
                    ws2 = wb2.create_sheet('Ausencias')
                    
                    headers_listado = ['Identificador', 'Nombre', 'Ciudad', 'Sede', 'Fecha', 'Novedad']
                    for ci, n in enumerate(headers_listado, start=1):
                        c = ws2.cell(row=1, column=ci, value=n)
                        c.font, c.fill, c.alignment, c.border = Font(name='Arial', bold=True, color='FFFFFF', size=10), PatternFill('solid', fgColor='1F4E79'), Alignment(horizontal='center', vertical='center'), brd
                    
                    for ri, row in listado.iterrows():
                        for ci, val in enumerate(row, start=1):
                            c = ws2.cell(row=ri + 2, column=ci, value=val)
                            c.font, c.fill, c.alignment, c.border = Font(name='Arial', size=9), PatternFill('solid', fgColor='FFFFFF' if ri % 2 == 0 else 'F2F2F2'), Alignment(horizontal='left' if ci in [2, 3, 4] else 'center', vertical='center'), brd
                    
                    ws2.column_dimensions['A'].width = 15
                    ws2.column_dimensions['B'].width = 35
                    ws2.column_dimensions['C'].width = 20
                    ws2.column_dimensions['D'].width = 25
                    ws2.column_dimensions['E'].width = 15
                    ws2.column_dimensions['F'].width = 25

                    headers_resumen = ['Identificador', 'Nombre', 'Ciudad', 'Sede', 'Cantidad']
                    start_col_resumen = 8
                    for ci, n in enumerate(headers_resumen, start=start_col_resumen):
                        c = ws2.cell(row=1, column=ci, value=n)
                        c.font, c.fill, c.alignment, c.border = Font(name='Arial', bold=True, color='FFFFFF', size=10), PatternFill('solid', fgColor='1F4E79'), Alignment(horizontal='center', vertical='center'), brd
                    
                    for ri, row in resumen.iterrows():
                        for ci, col_n in enumerate(headers_resumen, start=start_col_resumen):
                            val = row[col_n]
                            c = ws2.cell(row=ri + 2, column=ci, value=val)
                            c.font, c.fill, c.alignment, c.border = Font(name='Arial', size=9), PatternFill('solid', fgColor='FFFFFF' if ri % 2 == 0 else 'F2F2F2'), Alignment(horizontal='left' if col_n in ['Nombre', 'Ciudad', 'Sede'] else 'center', vertical='center'), brd
                    
                    ws2.column_dimensions['H'].width = 15
                    ws2.column_dimensions['I'].width = 35
                    ws2.column_dimensions['J'].width = 20
                    ws2.column_dimensions['K'].width = 25
                    ws2.column_dimensions['L'].width = 12
                    
                    ws2.freeze_panes, ws2.auto_filter.ref = 'A2', f"A1:L{max(len(listado), len(resumen)) + 1}"
                    
                    output_buffer = io.BytesIO()
                    wb2.save(output_buffer)

                    output_buffer.seek(0)
                    df_rep = pd.read_excel(output_buffer, sheet_name='Reporte_Horizontal')
                    COLS_FIJAS_REP = ["Identificador", "Concepto"]
                    cols_fecha_rep = [c for c in df_rep.columns if c not in COLS_FIJAS_REP]
                    
                    df_long = df_rep.melt(id_vars=COLS_FIJAS_REP, value_vars=cols_fecha_rep, var_name="FechaCorta", value_name="Valor")
                    
                    def fecha_corta_a_dt(texto):
                        meses_inv = {v: k for k, v in MESES_ES.items()}
                        try:
                            partes = str(texto).split("-")
                            return pd.Timestamp(2026, meses_inv[partes[1].lower()], int(partes[0]))
                        except: return pd.NaT

                    df_long["FechaReal"] = df_long["FechaCorta"].apply(fecha_corta_a_dt)
                    df_long = df_long.dropna(subset=["FechaReal"])
                    df_long = df_long[df_long["Valor"].notna() & df_long["Valor"].astype(str).str.strip().ne("") & df_long["Valor"].astype(str).str.strip().ne("nan")]
                    df_long["Identificador"] = df_long["Identificador"].apply(limpiar_id_a_texto)

                    def asignar_periodo(fecha):
                        for per, (inicio, fin) in PERIODOS.items():
                            if inicio <= fecha <= fin: return per
                        return None

                    df_long["Periodo"] = df_long["FechaReal"].apply(asignar_periodo)
                    df_long = df_long[df_long["Periodo"].notna()]

                    archivo_htcc.seek(0)
                    wb_htcc = load_workbook(archivo_htcc)
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

                    for _, fila_long in df_long.iterrows():
                        id_str = fila_long["Identificador"]
                        concepto_long = str(fila_long["Concepto"]).strip()
                        fecha_real_long = fila_long["FechaReal"].normalize()
                        valor_raw = fila_long["Valor"]
                        periodo_long = str(fila_long["Periodo"]).strip().upper()

                        concepto_libro3 = MAPA_CONCEPTOS.get(concepto_long)
                        if concepto_libro3 is None: continue

                        clave = (periodo_long, id_str, concepto_libro3)
                        fila_excel = indice_filas.get(clave)
                        if fila_excel is None: continue

                        col_excel = col_fecha_libro3.get(fecha_real_long)
                        if col_excel is None: continue

                        try:
                            valor_final = round(float(str(valor_raw)), 2)
                            es_numero = True
                        except:
                            valor_final = str(valor_raw).strip().upper()
                            es_numero = False

                        cell = ws_htcc.cell(row=fila_excel, column=col_excel, value=valor_final)
                        cell.alignment = Alignment(horizontal="center", vertical="center")
                        if es_numero:
                            cell.number_format = '#,##0.00'
                            cell.font = Font(name="Arial", size=9)
                            cell.fill = PatternFill(fill_type="solid", fgColor="FFFFFF")
                        else:
                            colores = COLORES_LETRAS.get(valor_final, {"bg":"FFFFFF","fg":"000000"})
                            cell.font = Font(name="Arial", size=9, bold=True, color=colores["fg"])
                            cell.fill = PatternFill(fill_type="solid", fgColor=colores["bg"])

                    cols_por_periodo = {p.upper(): [] for p in PERIODOS}
                    for f_dt, c_idx in col_fecha_libro3.items():
                        for per, (inicio, fin) in PERIODOS.items():
                            if inicio.normalize() <= f_dt <= fin.normalize():
                                cols_por_periodo[per.upper()].append(c_idx)
                                break

                    ws_htcc.cell(row=FILA_ENCABEZADO, column=COL_TOTAL_IDX, value="Total").font = Font(name="Arial", size=10, bold=True, color="FFFFFF")
                    ws_htcc.cell(row=FILA_ENCABEZADO, column=COL_TOTAL_IDX).fill = PatternFill(fill_type="solid", fgColor="000000")
                    ws_htcc.cell(row=FILA_ENCABEZADO, column=COL_TOTAL_IDX).alignment = Alignment(horizontal="center", vertical="center")

                    ws_htcc.cell(row=FILA_ENCABEZADO, column=COL_DIF_IDX, value="Diferencia").font = Font(name="Arial", size=10, bold=True, color="FFFFFF")
                    ws_htcc.cell(row=FILA_ENCABEZADO, column=COL_DIF_IDX).fill = PatternFill(fill_type="solid", fgColor="000000")
                    ws_htcc.cell(row=FILA_ENCABEZADO, column=COL_DIF_IDX).alignment = Alignment(horizontal="center", vertical="center")

                    for (periodo, id_str, conc_libro3), fila_excel in indice_filas.items():
                        cols_periodo = cols_por_periodo.get(periodo, [])
                        if not cols_periodo: continue
                        col_ini = get_column_letter(min(cols_periodo))
                        col_fin = get_column_letter(max(cols_periodo))
                        
                        formula_total = f"=SUM({col_ini}{fila_excel}:{col_fin}{fila_excel})"
                        color_fondo = "FFFFFF"
                        if conc_libro3 == "recargo nocturno 0.35%": color_fondo = "F2F2F2"
                        elif conc_libro3 == "recargo dominical compensado": color_fondo = "E2EFDA"
                        elif conc_libro3 == "recargo festivo": color_fondo = "FFF2CC"

                        c_total = ws_htcc.cell(row=fila_excel, column=COL_TOTAL_IDX, value=formula_total)
                        c_total.number_format, c_total.font, c_total.alignment, c_total.fill = '#,##0.00', Font(name="Arial", size=9, bold=True), Alignment(horizontal="center", vertical="center"), PatternFill(fill_type="solid", fgColor=color_fondo)

                        formula_dif = f"=ROUND({col_cantidad_letra}{fila_excel}-{get_column_letter(COL_TOTAL_IDX)}{fila_excel},0)"
                        c_dif = ws_htcc.cell(row=fila_excel, column=COL_DIF_IDX, value=formula_dif)
                        c_dif.number_format, c_dif.font, c_dif.alignment, c_dif.fill = '#,##0.00', Font(name="Arial", size=9, bold=True), Alignment(horizontal="center", vertical="center"), PatternFill(fill_type="solid", fgColor="FFFFFF")

                    ws_comp = wb_htcc.create_sheet('Comparaciones')

                    cols_excluir_ht = []
                    for cell in ws_htcc[FILA_ENCABEZADO]:
                        if cell.value:
                            v_lower = str(cell.value).strip().lower()
                            if "fecha de ingreso" in v_lower or "fecha de retiro" in v_lower or "fecha_ingreso" in v_lower or "fecha_retiro" in v_lower:
                                cols_excluir_ht.append(cell.column)

                    mapa_cols_ht_a_comp = {}
                    col_comp_idx = 1

                    for col_ht_idx in range(1, ws_htcc.max_column + 1):
                        if col_ht_idx in cols_excluir_ht:
                            continue

                        if col_ht_idx == col_id_libro3:
                            mapa_cols_ht_a_comp['NOM/Oper'] = col_comp_idx
                            col_comp_idx += 1

                        mapa_cols_ht_a_comp[col_ht_idx] = col_comp_idx
                        col_comp_idx += 1

                    col_obs_idx = col_comp_idx
                    mapa_cols_ht_a_comp['Observacion'] = col_obs_idx

                    col_cant_ht_idx = col_cantidad_libro3 if col_cantidad_libro3 else 7
                    col_cant_comp_idx = mapa_cols_ht_a_comp.get(col_cant_ht_idx, 8)
                    col_cant_comp_letra = get_column_letter(col_cant_comp_idx)

                    for col_ht_idx, col_c_idx in mapa_cols_ht_a_comp.items():
                        if col_ht_idx == 'NOM/Oper':
                            for r_idx in range(1, FILA_ENCABEZADO + 1):
                                c_dest = ws_comp.cell(row=r_idx, column=col_c_idx)
                                c_ref_orig = ws_htcc.cell(row=r_idx, column=1)
                                c_dest.font = Font(name=c_ref_orig.font.name, size=c_ref_orig.font.size, bold=True, color=c_ref_orig.font.color)
                                c_dest.fill = PatternFill(fill_type=c_ref_orig.fill.fill_type, fgColor=c_ref_orig.fill.fgColor)
                                c_dest.alignment = Alignment(horizontal="center", vertical="center")
                                if r_idx == FILA_ENCABEZADO:
                                    c_dest.value = "NOM/Oper"
                            continue

                        if col_ht_idx == 'Observacion':
                            for r_idx in range(1, FILA_ENCABEZADO + 1):
                                c_dest = ws_comp.cell(row=r_idx, column=col_c_idx)
                                c_dest.font = Font(name="Arial", size=10, bold=True, color="FFFFFF")
                                c_dest.fill = PatternFill(fill_type="solid", fgColor="000000")
                                c_dest.alignment = Alignment(horizontal="center", vertical="center")
                                if r_idx == FILA_ENCABEZADO:
                                    c_dest.value = "Observacion"
                            continue

                        for r_idx in range(1, FILA_ENCABEZADO + 1):
                            c_orig = ws_htcc.cell(row=r_idx, column=col_ht_idx)
                            c_dest = ws_comp.cell(row=r_idx, column=col_c_idx, value=c_orig.value)
                            
                            if c_orig.has_style:
                                c_dest.font = Font(name=c_orig.font.name, size=c_orig.font.size, bold=c_orig.font.bold, color=c_orig.font.color)
                                c_dest.fill = PatternFill(fill_type=c_orig.fill.fill_type, fgColor=c_orig.fill.fgColor)
                                c_dest.alignment = Alignment(horizontal=c_orig.alignment.horizontal, vertical=c_orig.alignment.vertical, text_rotation=c_orig.alignment.text_rotation)
                                c_dest.number_format = c_orig.number_format
                                
                                if c_orig.border:
                                    c_dest.border = Border(
                                        left=Side(style=c_orig.border.left.style, color=c_orig.border.left.color) if c_orig.border.left else None,
                                        right=Side(style=c_orig.border.right.style, color=c_orig.border.right.color) if c_orig.border.right else None,
                                        top=Side(style=c_orig.border.top.style, color=c_orig.border.top.color) if c_orig.border.top else None,
                                        bottom=Side(style=c_orig.border.bottom.style, color=c_orig.border.bottom.color) if c_orig.border.bottom else None
                                    )

                    cols_op_map = {str(c).strip().lower(): c for c in df_operativo.columns}
                    col_id_op = next((orig for k, orig in cols_op_map.items() if 'identificador' in k or 'cedula' in k or 'id' in k), None)
                    col_conc_op = next((orig for k, orig in cols_op_map.items() if 'concepto' in k), None)
                    
                    dict_operativo_cruce = {}
                    
                    if col_id_op and col_conc_op:
                        df_op_calc = df_operativo.copy()
                        df_op_calc['_id_clean'] = limpiar_id_a_texto(df_op_calc[col_id_op])
                        df_op_calc['_conc_norm'] = df_op_calc[col_conc_op].apply(normalizar_concepto_txt)
                        
                        cols_fecha_op_map = {}
                        for c_col in df_op_calc.columns:
                            if c_col in ('_id_clean', '_conc_norm', col_id_op, col_conc_op):
                                continue
                            f_clave = fecha_a_clave_corta(c_col)
                            if f_clave:
                                cols_fecha_op_map[c_col] = f_clave

                        for _, row_op in df_op_calc.iterrows():
                            id_op_val = row_op['_id_clean']
                            conc_op_norm = row_op['_conc_norm']
                            
                            for orig_col, f_clave in cols_fecha_op_map.items():
                                val_op = row_op[orig_col]
                                if pd.notna(val_op) and str(val_op).strip() not in ("", "nan", "None"):
                                    dict_operativo_cruce[(id_op_val, conc_op_norm, f_clave)] = val_op

                    fill_operativo = PatternFill(fill_type="solid", fgColor="FFFF99")
                    font_operativo = Font(name="Arial", size=9, color="000000", bold=True)
                    fill_rojo_alerta = PatternFill(fill_type="solid", fgColor="FF0000")
                    font_rojo_alerta = Font(name="Arial", size=9, color="FFFFFF", bold=True)

                    fila_comp = FILA_ENCABEZADO + 1

                    for (periodo, id_str, conc_libro3), fila_excel in indice_filas.items():
                        nombre_emp = ws_htcc.cell(row=fila_excel, column=col_nombre_libro3).value if col_nombre_libro3 else mapa_nombres_ht.get(id_str, "")

                        periodo_cap = periodo.title()
                        rango_periodo = PERIODOS.get(periodo_cap)

                        fila_nom_idx = fila_comp
                        fila_op_idx = fila_comp + 1

                        conc_libro3_norm = normalizar_concepto_txt(conc_libro3)

                        for col_ht_idx, col_c_idx in mapa_cols_ht_a_comp.items():
                            if col_ht_idx == 'Observacion':
                                continue

                            val_orig = ws_htcc.cell(row=fila_excel, column=col_ht_idx).value if col_ht_idx != 'NOM/Oper' else "Nomina"
                            c_nom = ws_comp.cell(row=fila_nom_idx, column=col_c_idx, value=val_orig)
                            c_orig = ws_htcc.cell(row=fila_excel, column=col_ht_idx) if col_ht_idx != 'NOM/Oper' else None
                            
                            if c_orig and c_orig.has_style:
                                c_nom.font = Font(name=c_orig.font.name, size=c_orig.font.size, bold=c_orig.font.bold, color=c_orig.font.color)
                                c_nom.fill = PatternFill(fill_type=c_orig.fill.fill_type, fgColor=c_orig.fill.fgColor)
                                c_nom.alignment = Alignment(horizontal="center" if col_ht_idx >= COL_INICIO_FECHAS or col_ht_idx == 'NOM/Oper' else "left", vertical="center")
                                c_nom.number_format = c_orig.number_format
                                c_nom.border = brd

                            c_op = ws_comp.cell(row=fila_op_idx, column=col_c_idx)
                            c_op.fill, c_op.font, c_op.border = fill_operativo, font_operativo, brd
                            c_op.alignment = Alignment(horizontal="center" if isinstance(col_ht_idx, int) and col_ht_idx >= COL_INICIO_FECHAS else "left", vertical="center")

                            if col_ht_idx == 'NOM/Oper':
                                c_op.value = "Operativo"
                                c_op.alignment = Alignment(horizontal="center", vertical="center")
                            elif col_ht_idx == col_periodo_libro3:
                                c_op.value = periodo.upper()
                            elif col_ht_idx == col_id_libro3:
                                c_op.value = id_str
                            elif col_nombre_libro3 and col_ht_idx == col_nombre_libro3:
                                c_op.value = nombre_emp
                            elif col_ht_idx == col_concepto_libro3:
                                c_op.value = conc_libro3
                            elif col_ht_idx == COL_TOTAL_IDX:
                                cols_periodo = cols_por_periodo.get(periodo.upper(), [])
                                if cols_periodo:
                                    col_ini_comp = mapa_cols_ht_a_comp.get(min(cols_periodo))
                                    col_fin_comp = mapa_cols_ht_a_comp.get(max(cols_periodo))
                                    if col_ini_comp and col_fin_comp:
                                        letra_ini = get_column_letter(col_ini_comp)
                                        letra_fin = get_column_letter(col_fin_comp)
                                        
                                        c_nom.value = f"=SUM({letra_ini}{fila_nom_idx}:{letra_fin}{fila_nom_idx})"
                                        c_nom.number_format = '#,##0.00'

                                        c_op.value = f"=SUM({letra_ini}{fila_op_idx}:{letra_fin}{fila_op_idx})"
                                        c_op.number_format = '#,##0.00'
                            elif col_ht_idx == COL_DIF_IDX:
                                col_tot_comp_idx = mapa_cols_ht_a_comp.get(COL_TOTAL_IDX)
                                if col_tot_comp_idx:
                                    letra_col_total = get_column_letter(col_tot_comp_idx)
                                    c_nom.value = f"=ROUND({col_cant_comp_letra}{fila_nom_idx}-{letra_col_total}{fila_nom_idx},0)"
                                    c_nom.number_format = '#,##0.00'
                                    c_op.value = ""
                            elif isinstance(col_ht_idx, int) and col_ht_idx >= COL_INICIO_FECHAS:
                                fecha_header_val = ws_htcc.cell(row=FILA_ENCABEZADO, column=col_ht_idx).value
                                dt_header = obtener_dt_fecha(fecha_header_val)
                                
                                esta_en_periodo = True
                                if dt_header and rango_periodo:
                                    inicio_per, fin_per = rango_periodo
                                    esta_en_periodo = (inicio_per <= dt_header <= fin_per)

                                if esta_en_periodo:
                                    f_clave_header = fecha_a_clave_corta(fecha_header_val)
                                    val_operativo_encontrado = None
                                    
                                    if f_clave_header:
                                        val_operativo_encontrado = dict_operativo_cruce.get((id_str, conc_libro3_norm, f_clave_header))

                                        if val_operativo_encontrado is None:
                                            for key_etiqueta, val_mapeado in MAPA_CONCEPTOS.items():
                                                if val_mapeado.lower() == conc_libro3_norm:
                                                    key_norm = normalizar_concepto_txt(key_etiqueta)
                                                    val_operativo_encontrado = dict_operativo_cruce.get((id_str, key_norm, f_clave_header))
                                                    if val_operativo_encontrado is not None:
                                                        break

                                    if val_operativo_encontrado is not None and str(val_operativo_encontrado).strip() not in ("", "nan", "None"):
                                        try:
                                            c_op.value = round(float(val_operativo_encontrado), 2)
                                            c_op.number_format = '#,##0.00'
                                        except:
                                            c_op.value = str(val_operativo_encontrado).strip()

                        fechas_con_diferencia = []

                        for col_ht_idx, col_c_idx in mapa_cols_ht_a_comp.items():
                            if isinstance(col_ht_idx, int) and col_ht_idx >= COL_INICIO_FECHAS and col_ht_idx not in (COL_TOTAL_IDX, COL_DIF_IDX):
                                fecha_header_val = ws_htcc.cell(row=FILA_ENCABEZADO, column=col_ht_idx).value
                                dt_header = obtener_dt_fecha(fecha_header_val)
                                
                                esta_en_periodo = True
                                if dt_header and rango_periodo:
                                    inicio_per, fin_per = rango_periodo
                                    esta_en_periodo = (inicio_per <= dt_header <= fin_per)

                                if esta_en_periodo:
                                    cell_nom = ws_comp.cell(row=fila_nom_idx, column=col_c_idx)
                                    cell_op = ws_comp.cell(row=fila_op_idx, column=col_c_idx)

                                    val_n = cell_nom.value
                                    val_o = cell_op.value

                                    if valores_son_diferentes(val_n, val_o):
                                        cell_nom.fill = fill_rojo_alerta
                                        cell_nom.font = font_rojo_alerta

                                        cell_op.fill = fill_rojo_alerta
                                        cell_op.font = font_rojo_alerta

                                        f_corta_txt = fecha_a_clave_corta(fecha_header_val)
                                        if f_corta_txt and f_corta_txt not in fechas_con_diferencia:
                                            fechas_con_diferencia.append(f_corta_txt)

                        c_obs_nom = ws_comp.cell(row=fila_nom_idx, column=col_obs_idx)
                        c_obs_op = ws_comp.cell(row=fila_op_idx, column=col_obs_idx)

                        c_obs_nom.border, c_obs_op.border = brd, brd
                        c_obs_op.fill = fill_operativo

                        if fechas_con_diferencia:
                            txt_fechas = ", ".join(fechas_con_diferencia)
                            c_obs_nom.value = f"Diferencia en fecha(s): {txt_fechas}"
                            c_obs_nom.font = Font(name="Arial", size=9, bold=True, color="C62828")
                            c_obs_nom.alignment = Alignment(horizontal="left", vertical="center")
                        else:
                            c_obs_nom.value = "Sin diferencias"
                            c_obs_nom.font = Font(name="Arial", size=9, bold=False, color="2E7D32")
                            c_obs_nom.alignment = Alignment(horizontal="center", vertical="center")

                        ws_comp.column_dimensions[get_column_letter(col_obs_idx)].width = 35

                        fila_comp += 2

                    ws_comp.freeze_panes, ws_comp.auto_filter.ref = f"A{FILA_ENCABEZADO + 1}", f"A{FILA_ENCABEZADO}:{get_column_letter(ws_comp.max_column)}{ws_comp.max_row}"

                    htcc_buffer = io.BytesIO()
                    wb_htcc.save(htcc_buffer)

                    fecha_inicio_sem = pd.Timestamp(fecha_inicio_input)
                    wb_c = load_workbook(output_buffer)
                    ws_c = wb_c['Reporte_Horizontal']
                    
                    actual_col_inicio = 3
                    for col in range(1, ws_c.max_column + 1):
                        val_header = str(ws_c.cell(row=1, column=col).value or "")
                        if "-" in val_header and any(m in val_header.lower() for m in ["ene","feb","mar","abr","may","jun","jul","ago","sep","oct","nov","dic"]):
                            actual_col_inicio = col
                            break

                    columnas_c = []
                    MESES_ES_INV = {v: k for k, v in MESES_ES.items()}
                    for col in range(actual_col_inicio, ws_c.max_column + 1):
                        fecha_str = ws_c.cell(row=1, column=col).value
                        dia_str   = ws_c.cell(row=2, column=col).value
                        if fecha_str is None: continue
                        try:
                            pf = str(fecha_str).strip().split("-")
                            fecha_dt = pd.Timestamp(fecha_inicio_sem.year, MESES_ES_INV[pf[1].lower()], int(pf[0]))
                            columnas_c.append({"col": col, "fecha_str": fecha_str, "dia": str(dia_str).strip().lower().split()[0] if dia_str else "", "fecha_dt": fecha_dt})
                        except: pass

                    columnas_c = [c for c in columnas_c if c["fecha_dt"] >= fecha_inicio_sem]
                    semanas_c = []
                    fecha_sem = fecha_inicio_sem
                    num_sem = 1
                    while columnas_c and fecha_sem <= columnas_c[-1]["fecha_dt"]:
                        fecha_fin_sem = fecha_sem + pd.Timedelta(days=6)
                        cols_semana = [c for c in columnas_c if fecha_sem <= c["fecha_dt"] <= fecha_fin_sem]
                        if cols_semana:
                            semanas_c.append({"num": num_sem, "label": f"Sem{num_sem}", "inicio": fecha_sem, "fin": fecha_fin_sem, "columnas": cols_semana})
                        fecha_sem = fecha_fin_sem + pd.Timedelta(days=1)
                        num_sem += 1

                    df_id_clean = pd.to_numeric(df[col_id], errors='coerce').fillna(0).astype(int).astype(str)
                    
                    mapa_nombres_c = dict(zip(df_id_clean, df['Nombre'].astype(str).str.strip()))
                    mapa_ciudad_c  = dict(zip(df_id_clean, df['Ciudad'].astype(str).str.strip()))
                    mapa_sede_c    = dict(zip(df_id_clean, df['Sede'].astype(str).str.strip()))

                    empleados_c = []
                    for row in range(3, ws_c.max_row + 1):
                        id_val = ws_c.cell(row=row, column=1).value
                        conc_val = None
                        for c_idx in range(2, actual_col_inicio):
                            c_val = str(ws_c.cell(row=row, column=c_idx).value).strip()
                            if c_val == "HT Normales":
                                conc_val = c_val
                                break
                        if id_val is None or conc_val is None: continue
                        try: id_clean = str(int(float(id_val)))
                        except: id_clean = str(id_val).strip()
                        
                        valores = {col_info["col"]: str(ws_c.cell(row=row, column=col_info["col"]).value).strip() if ws_c.cell(row=row, column=col_info["col"]).value is not None else None for col_info in columnas_c}
                        
                        empleados_c.append({
                            "id": id_clean, 
                            "nombre": mapa_nombres_c.get(id_clean, "No Encontrado"), 
                            "ciudad": mapa_ciudad_c.get(id_clean, "No Registra"), 
                            "sede": mapa_sede_c.get(id_clean, "No Registra"), 
                            "concepto": conc_val, 
                            "fila": row, 
                            "valores": valores
                        })

                    resultados_c = []
                    max_fechas_c = 0
                    for emp in empleados_c:
                        for semana in semanas_c:
                            cols_semana = semana["columnas"]
                            cols_domingo = [d for d in cols_semana if d["dia"] == "domingo"]
                            cols_c_val = [d for d in cols_semana if emp["valores"].get(d["col"]) == "C"]
                            trabajo_domingo = any(emp["valores"].get(d["col"]) not in (None, "A","V","S","INC","REN","LR","PD") for d in cols_domingo)
                            cantidad_c = len(cols_c_val)
                            
                            if len(cols_domingo) == 0: estado = "-" if cantidad_c == 0 else "NO CUMPLE"
                            elif trabajo_domingo: estado = "CUMPLE" if cantidad_c == 1 else "-" if cantidad_c == 0 else "NO CUMPLE"
                            else: estado = "-" if cantidad_c == 0 else "NO CUMPLE"

                            if estado != "-" and cantidad_c > 0:
                                fechas_c_list = [c["fecha_str"] for c in cols_c_val]
                                max_fechas_c = max(max_fechas_c, len(fechas_c_list))
                                
                                fila = {
                                    "Identificador": emp["id"], 
                                    "Nombre": emp["nombre"], 
                                    "Ciudad": emp["ciudad"], 
                                    "Sede": emp["sede"], 
                                    "Concepto": emp["concepto"], 
                                    "Trabajo Domingo": "SI" if trabajo_domingo else "NO", 
                                    "Cant. C": cantidad_c, 
                                    "Estado": estado
                                }
                                for i, f in enumerate(fechas_c_list, 1): fila[f"Fecha {i}"] = f
                                resultados_c.append(fila)

                    df_c = pd.DataFrame(resultados_c) if resultados_c else pd.DataFrame()

                    if not df_c.empty:
                        ws_out_c = wb_c.create_sheet("Analisis C")
                        headers_base = ["Identificador", "Nombre", "Ciudad", "Sede", "Concepto", "Trabajo Domingo", "Cant. C", "Estado"]
                        headers_fechas = [f"Fecha {i}" for i in range(1, max_fechas_c + 1)]
                        headers_c = headers_base + headers_fechas
                        brd_c = Border(left=Side(style="thin", color="CCCCCC"), right=Side(style="thin", color="CCCCCC"), top=Side(style="thin", color="CCCCCC"), bottom=Side(style="thin", color="CCCCCC"))

                        for col_i, h in enumerate(headers_c, 1):
                            cell = ws_out_c.cell(row=1, column=col_i, value=h)
                            cell.font, cell.fill, cell.alignment, cell.border = Font(name="Arial", size=10, bold=True, color="FFFFFF"), PatternFill("solid", fgColor='1F4E79'), Alignment(horizontal="center", vertical="center"), brd_c

                        n_cumple = len(df_c[df_c["Estado"] == "CUMPLE"])
                        n_no_cumple = len(df_c[df_c["Estado"] == "NO CUMPLE"])

                        for row_i, (_, row) in enumerate(df_c.iterrows(), 2):
                            color_fondo = 'D9E1F2' if row_i % 2 == 0 else 'FFFFFF'
                            for col_i, h in enumerate(headers_c, 1):
                                val = row.get(h, "")
                                cell = ws_out_c.cell(row=row_i, column=col_i, value=val)
                                cell.alignment = Alignment(horizontal="left" if h in ["Nombre", "Ciudad", "Sede"] else "center", vertical="center")
                                cell.border = brd_c
                                if h == "Estado":
                                    cell.fill = PatternFill("solid", fgColor='92D050' if val == "CUMPLE" else 'FF0000')
                                    cell.font = Font(name="Arial", size=9, bold=True, color='000000' if val == "CUMPLE" else 'FFFFFF')
                                elif h.startswith("Fecha ") and val:
                                    cell.fill, cell.font = PatternFill("solid", fgColor='E2EFDA'), Font(name="Arial", size=9, bold=True, color="1A5C2A")
                                else:
                                    cell.fill, cell.font = PatternFill("solid", fgColor=color_fondo), Font(name="Arial", size=9)

                        for col_i, ancho in enumerate([15, 35, 20, 25, 20, 16, 10, 12] + [12] * max_fechas_c, 1):
                            ws_out_c.column_dimensions[get_column_letter(col_i)].width = ancho
                        
                        fila_total_c = len(df_c) + 2
                        ws_out_c.cell(row=fila_total_c, column=1, value="TOTALES").font = Font(bold=True)
                        ws_out_c.cell(row=fila_total_c, column=7, value=f"CUMPLE: {n_cumple}").font = Font(bold=True, color="2E7D32")
                        ws_out_c.cell(row=fila_total_c, column=8, value=f"NO CUMPLE: {n_no_cumple}").font = Font(bold=True, color="C62828")
                        ws_out_c.freeze_panes, ws_out_c.auto_filter.ref = 'A2', f"A1:{get_column_letter(len(headers_c))}{fila_total_c - 1}"
                    
                    output_buffer = io.BytesIO()
                    wb_c.save(output_buffer)
                    
                    st.session_state.output_bytes_m2 = output_buffer.getvalue()
                    st.session_state.htcc_bytes_m2 = htcc_buffer.getvalue()
                    st.session_state.listado_m2 = listado
                    st.session_state.resumen_m2 = resumen
                    st.session_state.df_c_m2 = df_c
                    st.session_state.resultados_c_m2 = resultados_c
                    st.session_state.procesado_m2 = True

                except Exception as e:
                    st.error(f"❌ Ocurrió un error inesperado al procesar: {e}")

        # ── BLOQUE DE RENDERIZADO VISUAL DEL MÓDULO 2 ──
        if st.session_state.get("procesado_m2", False):
            st.success("🎉 ¡Reporte y Consolidación Multi-Periodo procesados exitosamente!")
            
            col_down1, col_down2 = st.columns(2)
            with col_down1:
                st.download_button(
                    label="📥 Descargar Reporte Ausencias Y Compensados Procesado",
                    data=st.session_state.output_bytes_m2,
                    file_name="Reporte_Ausencias_Compensados.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="d_btn_m2_1"
                )
            with col_down2:
                st.download_button(
                    label="📥 Descargar Plantilla HTCC Consolidada",
                    data=st.session_state.htcc_bytes_m2,
                    file_name="HTCC_Consolidado_2026.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="d_btn_m2_2"
                )

            st.markdown("---")
            st.header("📋 Vista Previa de Resultados")
            
            tab_aus, tab_comp = st.tabs([
                "📄 Hoja de Ausencias", 
                "🔍 Análisis de Compensatorios"
            ])

            with tab_aus:
                col_titulo, col_metrica = st.columns([3, 1])
                with col_titulo:
                    st.subheader("Registros Detallados de Ausencias")
                with col_metrica:
                    total_ausencias = len(st.session_state.listado_m2)
                    st.markdown(f"""
                        <div style="background-color:#FFEB9C; padding:5px 15px; border-radius:15px; text-align:center; border:1px solid #FFC7CE; margin-top:5px;">
                            <strong style="color:#9C0006; font-size:16px;">Total: {total_ausencias}</strong>
                        </div>
                    """, unsafe_allow_html=True)
                st.dataframe(st.session_state.listado_m2, use_container_width=True, hide_index=True)
                st.subheader("Resumen Consolidado por Persona")
                st.dataframe(st.session_state.resumen_m2, use_container_width=True, hide_index=True)
                
            with tab_comp:
                st.subheader("Validación de Compensatorios (Analisis C)")
                if st.session_state.resultados_c_m2:
                    st.write("Filtrar por Estado:")
                    col_cumple, col_nocumple, _ = st.columns([1, 1, 3])
                    with col_cumple:
                        chk_cumple = st.checkbox("CUMPLE", value=True, key="filtro_cumple_m2")
                    with col_nocumple:
                        chk_nocumple = st.checkbox("NO CUMPLE", value=True, key="filtro_nocumple_m2")
                    
                    estados_activos = []
                    if chk_cumple:
                        estados_activos.append("CUMPLE")
                    if chk_nocumple:
                        estados_activos.append("NO CUMPLE")
                    
                    df_c_filtrado = st.session_state.df_c_m2[st.session_state.df_c_m2["Estado"].isin(estados_activos)]
                    st.dataframe(df_c_filtrado, use_container_width=True, hide_index=True)
                else:
                    st.info("No se encontraron registros de compensatorios que requieran validación para el periodo seleccionado.")
