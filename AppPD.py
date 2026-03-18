import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import urllib.parse
import requests
import zipfile
import io
import os
import plotly.express as px

# --- 1. CONFIGURACIÓN DE PÁGINA ---
st.set_page_config(page_title="Supervisión SEIN - Osinergmin", layout="wide")
st.title("⚡ Dashboard de Supervisión de Mantenimientos - SEIN")
st.markdown("Supervisión COES: Mantenimientos Programado vs. Ejecutado + Resumen Operativo")

MESES = {
    1: "ENERO", 2: "FEBRERO", 3: "MARZO", 4: "ABRIL",
    5: "MAYO", 6: "JUNIO", 7: "JULIO", 8: "AGOSTO",
    9: "SETIEMBRE", 10: "OCTUBRE", 11: "NOVIEMBRE", 12: "DICIEMBRE"
}

# --- ARCHIVOS DE PERSISTENCIA ---
ARCHIVO_POTENCIAS_LOCAL = "potencias_historicas.csv"
ARCHIVO_GITHUB_SEMILLA = "pOTENCIAS.csv" # Archivo del repositorio provisto

def cargar_potencias_guardadas():
    if os.path.exists(ARCHIVO_POTENCIAS_LOCAL):
        df = pd.read_csv(ARCHIVO_POTENCIAS_LOCAL)
    elif os.path.exists(ARCHIVO_GITHUB_SEMILLA):
        df = pd.read_csv(ARCHIVO_GITHUB_SEMILLA)
        if 'Empresa' not in df.columns:
            df['Empresa'] = 'NO ESPECIFICADO'
    else:
        df = pd.DataFrame(columns=['Empresa', 'Central/Ubicacion', 'Equipo', 'Potencia_Indisponible_MW'])
    
    return df.drop_duplicates(subset=['Central/Ubicacion', 'Equipo'], keep='last')

def guardar_potencias_asignadas(df_nuevas):
    df_historico = cargar_potencias_guardadas()
    if not df_historico.empty:
        df_final = pd.concat([df_nuevas, df_historico]).drop_duplicates(subset=['Central/Ubicacion', 'Equipo'], keep='first')
    else:
        df_final = df_nuevas
    df_final.to_csv(ARCHIVO_POTENCIAS_LOCAL, index=False)

def generar_urls_coes(fecha):
    año = fecha.strftime("%Y")
    mes_num = fecha.strftime("%m")
    dia = fecha.strftime("%d")
    mes_mayus = MESES[fecha.month]
    mes_titulo = MESES[fecha.month].capitalize()
    
    fecha_str_prog = fecha.strftime("%Y%m%d")
    path_prog = f"Operación/Programa de Mantenimiento/Programa Diario/{año}/{mes_num}_{mes_mayus}/Día {dia}/Anexo1_Intervenciones_{fecha_str_prog}.zip"
    url_prog = f"https://www.coes.org.pe/portal/browser/download?url={urllib.parse.quote(path_prog)}"
    
    fecha_str_ejec = fecha.strftime("%d%m")
    rutas_ejec_opciones = [
        (f"Post Operación/Reportes/IEOD/{año}/{mes_num}_{mes_titulo}/{dia}/AnexoA_{fecha_str_ejec}.xlsx", 5),
        (f"Post Operación/Reportes/IEOD/{año}/{mes_num}_{mes_titulo}/{dia}/Anexo5_Manttoejec_{fecha_str_ejec}.xls", 8),
        (f"Post Operación/Reportes/IEOD/{año}/{mes_num}_{mes_titulo}/{dia}/Anexo5_Manttoejec_{fecha_str_ejec}.xlsx", 8)
    ]
    urls_ejec = [(f"https://www.coes.org.pe/portal/browser/download?url={urllib.parse.quote(path)}", skip) for path, skip in rutas_ejec_opciones]
    
    # URL de Costos Marginales
    fecha_str_cmg = fecha.strftime("%Y%m%d")
    path_cmg = f"Post Operación/Reportes/IEOD/{año}/{mes_num}_{mes_titulo}/{dia}/CMg{fecha_str_cmg}.zip"
    url_cmg = f"https://www.coes.org.pe/portal/browser/download?url={urllib.parse.quote(path_cmg)}"
    
    return url_prog, urls_ejec, url_cmg

def col2idx(col_str):
    """Convierte letras de columna Excel a índice base 0."""
    expn = 0
    col_num = 0
    for char in reversed(col_str):
        col_num += (ord(char.upper()) - ord('A') + 1) * (26 ** expn)
        expn += 1
    return col_num - 1

# --- 2. EXTRACCIÓN Y LIMPIEZA ESPEJO (ETL UNIFICADO) ---
@st.cache_data(show_spinner=False)
def extraer_datos_coes(fecha):
    url_prog, urls_ejec, url_cmg = generar_urls_coes(fecha)
    headers = {'User-Agent': 'Mozilla/5.0'}
    
    df_prog, df_ejec, df_rf, df_cmg = pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), pd.DataFrame()
    columnas_estandar = ['Empresa', 'Ubicacion', 'Equipo', 'Inicio', 'Fin', 'Descripcion', 'MW_Indisponibles', 'Es_Programado', 'Disponibilidad_Equipo', 'Ocasiona_Interrupciones', 'Tipo_Mantenimiento', 'Codigo_Equipo', 'Tipo_Equipo']
    
    # 2.1 EXTRACCIÓN DEL PROGRAMADO
    try:
        res_prog = requests.get(url_prog, headers=headers, timeout=20)
        if res_prog.status_code == 200:
            with zipfile.ZipFile(io.BytesIO(res_prog.content)) as z:
                archivo_objetivo = None
                for f in z.namelist():
                    nombre_archivo = f.split('/')[-1]
                    if nombre_archivo.endswith(('.xls', '.xlsx')) and not nombre_archivo.startswith('~$') and '__MACOSX' not in f:
                        if 'Anexo1_Intervenciones_(Osinergmin)' in nombre_archivo:
                            archivo_objetivo = f
                            break
                            
                if archivo_objetivo:
                    with z.open(archivo_objetivo) as f:
                        file_ext = archivo_objetivo.split('.')[-1]
                        motor_excel = 'xlrd' if file_ext == 'xls' else 'openpyxl'
                        archivo_excel_prog = pd.ExcelFile(io.BytesIO(f.read()), engine=motor_excel)
                        sheet_names_prog = archivo_excel_prog.sheet_names
                        
                        hoja_prog = next((h for h in sheet_names_prog if 'PROGRAMADO' in h.upper() and 'DIARIO' in h.upper()), None)
                        if not hoja_prog: hoja_prog = next((h for h in sheet_names_prog if 'PROGRAMADO' in h.upper()), None)
                        if not hoja_prog: hoja_prog = next((h for h in sheet_names_prog if 'ANEXO' in h.upper()), None)
                        if not hoja_prog: hoja_prog = sheet_names_prog[0] 
                            
                        df_prog = pd.read_excel(archivo_excel_prog, sheet_name=hoja_prog, skiprows=8, usecols="B:N", names=columnas_estandar)
                        df_prog = df_prog.dropna(subset=['Empresa', 'Equipo'], how='all')
                        df_prog = df_prog[~df_prog['Empresa'].astype(str).str.contains('TOTAL|NOTA|ELABORADO|FUENTE', case=False, na=False)]
    except Exception:
        pass 

    # 2.2 EXTRACCIÓN DEL EJECUTADO Y RESERVA FRÍA
    exito_ejecutado = False
    for url_ejec, skiprows_val in urls_ejec:
        if exito_ejecutado: break
        try:
            res_ejec = requests.get(url_ejec, headers=headers, timeout=15)
            if res_ejec.status_code == 200:
                motor_excel = 'xlrd' if 'xls' in urllib.parse.unquote(url_ejec).split('.')[-1] and not 'xlsx' in urllib.parse.unquote(url_ejec) else 'openpyxl'
                archivo_excel = pd.ExcelFile(io.BytesIO(res_ejec.content), engine=motor_excel)
                
                sheet_names_ejec = archivo_excel.sheet_names
                hoja_buscada = next((h for h in sheet_names_ejec if 'EJECUTADO' in h.upper()), None)
                if not hoja_buscada: hoja_buscada = sheet_names_ejec[0]
                
                if hoja_buscada:
                    df_ejec = pd.read_excel(archivo_excel, sheet_name=hoja_buscada, skiprows=skiprows_val, usecols="B:N", names=columnas_estandar)
                    df_ejec = df_ejec.dropna(subset=['Empresa', 'Equipo'], how='all')
                    df_ejec = df_ejec[~df_ejec['Empresa'].astype(str).str.contains('TOTAL|NOTA|ELABORADO|FUENTE', case=False, na=False)]
                    exito_ejecutado = True
                    
                # Extracción anidada de Reserva Fría
                hojas_limpias = {h.strip().upper(): h for h in sheet_names_ejec}
                nombres_posibles_rf = [h for h in hojas_limpias.keys() if "RESERVA" in h and "FR" in h]
                if nombres_posibles_rf:
                    hoja_rf = hojas_limpias[nombres_posibles_rf[0]]
                    df_raw_rf = pd.read_excel(archivo_excel, sheet_name=hoja_rf, header=None)
                    
                    codigos_restriccion = [239, 263, 265, 240, 241, 242, 924, 926, 786, 787, 788, 789, 995, 996, 997, 758, 42667, 42688, 756, 156]
                    codigos_nodo = [240, 241, 242, 995, 996, 997, 926, 924, 786, 787, 788, 789, 239, 263, 265]
                    
                    col_rf = None
                    cols_restriccion = []
                    cols_nodo = []
                    fila_lista = []
                    
                    for idx_fila in range(3, 7): 
                        fila_vals = df_raw_rf.iloc[idx_fila].values
                        if 7000 in fila_vals:
                            fila_lista = list(fila_vals)
                            col_rf = fila_lista.index(7000)
                            for cod in codigos_restriccion:
                                if cod in fila_lista: cols_restriccion.append(fila_lista.index(cod))
                            for cod in codigos_nodo:
                                if cod in fila_lista: cols_nodo.append(fila_lista.index(cod))
                            break
                    
                    if col_rf is not None:
                        data_rf = df_raw_rf.iloc[6:54, col_rf].values
                        reserva_fria_series = pd.to_numeric(pd.Series(data_rf), errors='coerce').fillna(0)
                        
                        reglas_cc = [
                            {'tv': 56677, 'tgs': [209]}, {'tv': 250, 'tgs': [252, 249]},
                            {'tv': 236, 'tgs': [194, 196, 207]}, {'tv': 285, 'tgs': [795]},
                            {'tv': 193, 'tgs': [113, 114]}, {'tv': 2159, 'tgs': [248]}
                        ]
                        for regla in reglas_cc:
                            tv_cod = regla['tv']
                            tgs_cods = regla['tgs']
                            if tv_cod in fila_lista:
                                idx_tv = fila_lista.index(tv_cod)
                                tv_series = pd.to_numeric(pd.Series(df_raw_rf.iloc[6:54, idx_tv].values), errors='coerce').fillna(0)
                                sum_tgs = pd.Series(np.zeros(48))
                                for tg in tgs_cods:
                                    if tg in fila_lista:
                                        idx_tg = fila_lista.index(tg)
                                        tg_series = pd.to_numeric(pd.Series(df_raw_rf.iloc[6:54, idx_tg].values), errors='coerce').fillna(0)
                                        sum_tgs += tg_series
                                reserva_fria_series = pd.Series(np.where(sum_tgs <= 0, reserva_fria_series - tv_series, reserva_fria_series))
                        
                        reserva_fria_series = reserva_fria_series.clip(lower=0) 
                        
                        restriccion_total = pd.Series(np.zeros(48))
                        for col_idx in cols_restriccion:
                            data_res = df_raw_rf.iloc[6:54, col_idx].values
                            restriccion_total += pd.to_numeric(pd.Series(data_res), errors='coerce').fillna(0)
                            
                        reserva_eficiente_series = reserva_fria_series - restriccion_total
                        reserva_eficiente_series = reserva_eficiente_series.clip(lower=0) 
                        
                        reserva_nodo_series = pd.Series(np.zeros(48))
                        for col_idx in cols_nodo:
                            data_nodo = df_raw_rf.iloc[6:54, col_idx].values
                            reserva_nodo_series += pd.to_numeric(pd.Series(data_nodo), errors='coerce').fillna(0)
                        
                        dt_fecha = datetime.combine(fecha, datetime.min.time())
                        fechas_horas = [dt_fecha + timedelta(minutes=30 * (i + 1)) for i in range(48)]
                        
                        df_rf = pd.DataFrame({
                            'FECHA_HORA': fechas_horas,
                            'RESERVA_FRIA_MW': reserva_fria_series,
                            'RESERVA_EFICIENTE_MW': reserva_eficiente_series,
                            'RESERVA_NODO_MW': reserva_nodo_series
                        })

        except Exception:
            continue

    # 2.3 EXTRACCIÓN DE COSTOS MARGINALES
    try:
        res_cmg = requests.get(url_cmg, headers=headers, timeout=20)
        if res_cmg.status_code == 200 and b"html" not in res_cmg.content[:100].lower():
            with zipfile.ZipFile(io.BytesIO(res_cmg.content)) as z:
                target_file = None
                for fname in z.namelist():
                    if "CMgCP" in fname and fname.endswith((".xlsx", ".xls")):
                        target_file = fname
                        break
                if target_file:
                    with z.open(target_file) as f:
                        file_bytes = io.BytesIO(f.read())
                        engine = 'openpyxl' if target_file.endswith('.xlsx') else None
                        df_raw_cmg = pd.read_excel(file_bytes, header=None, engine=engine)
                        
                        idx_sr = col2idx('GV')  # SANTA ROSA 220
                        idx_ta = col2idx('HH')  # TALARA 220
                        idx_mo = col2idx('EP')  # MOQUEGUA 220
                        
                        val_sr = pd.to_numeric(df_raw_cmg.iloc[3:51, idx_sr], errors='coerce').values
                        val_ta = pd.to_numeric(df_raw_cmg.iloc[3:51, idx_ta], errors='coerce').values
                        val_mo = pd.to_numeric(df_raw_cmg.iloc[3:51, idx_mo], errors='coerce').values
                        
                        dt_fecha = datetime.combine(fecha, datetime.min.time())
                        fechas_horas = [dt_fecha + timedelta(minutes=30 * (i + 1)) for i in range(48)]
                        
                        df_cmg_tmp = pd.DataFrame({
                            'FECHA_HORA': fechas_horas,
                            'SANTA ROSA 220': val_sr,
                            'TALARA 220': val_ta,
                            'MOQUEGUA 220': val_mo
                        })
                        df_cmg = df_cmg_tmp.melt(id_vars=['FECHA_HORA'], var_name='NODO', value_name='COSTO_MARGINAL_SOLES')
    except Exception:
        pass

    return df_prog, df_ejec, df_rf, df_cmg

def obtener_datos_rango(fecha_inicio, fecha_fin):
    dfs_prog, dfs_ejec, dfs_rf, dfs_cmg = [], [], [], []
    rango_dias = pd.date_range(fecha_inicio, fecha_fin)
    total_dias = len(rango_dias)
    
    barra_progreso = st.progress(0)
    texto_progreso = st.empty()
    
    for i, d in enumerate(rango_dias):
        texto_progreso.text(f"⏳ Extrayendo Despachos, Reservas y CMg del {d.strftime('%d/%m/%Y')} ({i+1}/{total_dias})...")
        df_p, df_e, df_r, df_c = extraer_datos_coes(d)
        
        if df_p is not None and not df_p.empty:
            df_p.insert(0, 'Fecha_Operacion', d.date())
            dfs_prog.append(df_p)
        if df_e is not None and not df_e.empty:
            df_e.insert(0, 'Fecha_Operacion', d.date())
            dfs_ejec.append(df_e)
        if df_r is not None and not df_r.empty:
            dfs_rf.append(df_r)
        if df_c is not None and not df_c.empty:
            dfs_cmg.append(df_c)
            
        barra_progreso.progress((i + 1) / total_dias)
    
    texto_progreso.empty()
    barra_progreso.empty()
        
    df_prog_final = pd.concat(dfs_prog, ignore_index=True) if dfs_prog else pd.DataFrame()
    df_ejec_final = pd.concat(dfs_ejec, ignore_index=True) if dfs_ejec else pd.DataFrame()
    df_rf_final = pd.concat(dfs_rf, ignore_index=True) if dfs_rf else pd.DataFrame()
    df_cmg_final = pd.concat(dfs_cmg, ignore_index=True) if dfs_cmg else pd.DataFrame()
    
    return df_prog_final, df_ejec_final, df_rf_final, df_cmg_final

# --- 3. FUNCIONES GLOBALES DE NORMALIZACIÓN ---
def normalizar_texto(serie):
    return serie.astype(str).str.strip().str.upper().str.normalize('NFKD').str.encode('ascii', errors='ignore').str.decode('utf-8')

def determinar_sector(row):
    tipo = str(row.get('Tipo_Equipo', ''))
    equipo = str(row.get('Equipo', ''))
    if tipo.startswith('G'): return 'GENERACIÓN'
    if tipo.startswith('T') or tipo.startswith('L'): return 'TRANSMISIÓN'
    if equipo.startswith('L-') or equipo.startswith('TR') or equipo.startswith('AT') or equipo.startswith('SE '): return 'TRANSMISIÓN'
    if equipo.startswith('G-') or equipo.startswith('TV') or equipo.startswith('TG') or equipo.startswith('CH '): return 'GENERACIÓN'
    return 'OTROS'

# --- 4. MOTOR DE CONCILIACIÓN ---
def conciliar_datos(df_prog, df_ejec):
    if not df_prog.empty:
        df_prog = df_prog.drop_duplicates(subset=['Fecha_Operacion', 'Empresa', 'Ubicacion', 'Equipo', 'Inicio', 'Fin'])
    if not df_ejec.empty:
        df_ejec = df_ejec.drop_duplicates(subset=['Fecha_Operacion', 'Empresa', 'Ubicacion', 'Equipo', 'Inicio', 'Fin'])

    for df in [df_prog, df_ejec]:
        if not df.empty:
            df['Empresa'] = normalizar_texto(df['Empresa'])
            df['Equipo'] = normalizar_texto(df['Equipo'])
            df.replace(['NAN', 'NAT', ''], np.nan, inplace=True)
            df['Seq_Mantenimiento'] = df.groupby(['Fecha_Operacion', 'Empresa', 'Equipo']).cumcount()
            
    if df_prog.empty: df_prog['Seq_Mantenimiento'] = []
    if df_ejec.empty: df_ejec['Seq_Mantenimiento'] = []

    df_merged = pd.merge(df_prog, df_ejec, on=['Fecha_Operacion', 'Empresa', 'Equipo', 'Seq_Mantenimiento'], how='outer', suffixes=('_Prog', '_Ejec'), indicator=True)
    df_merged.drop(columns=['Seq_Mantenimiento'], inplace=True, errors='ignore')
    
    def estado_supervision(row):
        if row['_merge'] == 'left_only': return 'Programado NO Ejecutado'
        elif row['_merge'] == 'right_only': return 'Ejecutado NO Programado'
        else: return 'Programado y Ejecutado'
            
    df_merged['Estado_Supervision'] = df_merged.apply(estado_supervision, axis=1)
    
    df_merged['Disponibilidad_Equipo'] = df_merged['Disponibilidad_Equipo_Ejec'].fillna(df_merged['Disponibilidad_Equipo_Prog']).fillna('NO ESPECIFICADO').astype(str).str.upper().str.strip()
    df_merged['Tipo_Equipo'] = df_merged['Tipo_Equipo_Ejec'].fillna(df_merged['Tipo_Equipo_Prog']).fillna('').astype(str).str.upper().str.strip()
    df_merged['Tipo_Mantenimiento'] = df_merged['Tipo_Mantenimiento_Ejec'].fillna(df_merged['Tipo_Mantenimiento_Prog']).fillna('NO ESPECIFICADO').astype(str).str.upper().str.strip()
    df_merged['Tipo_Mantenimiento'] = df_merged['Tipo_Mantenimiento'].replace('NAN', 'NO ESPECIFICADO')
    df_merged['Descripcion_Ejec'] = df_merged['Descripcion_Ejec'].fillna(df_merged['Descripcion_Prog']).fillna('-')
    
    df_merged['Central/Ubicacion'] = df_merged['Ubicacion_Ejec'].fillna(df_merged['Ubicacion_Prog']).fillna('-').astype(str).str.upper().str.strip()
    df_merged['Central/Ubicacion'] = normalizar_texto(df_merged['Central/Ubicacion'])
    df_merged['Sector'] = df_merged.apply(determinar_sector, axis=1)

    for col in ['Inicio_Prog', 'Fin_Prog', 'Inicio_Ejec', 'Fin_Ejec']:
        df_merged[col] = pd.to_datetime(df_merged[col], dayfirst=True, errors='coerce')
    df_merged['Horas_Prog'] = np.round((df_merged['Fin_Prog'] - df_merged['Inicio_Prog']).dt.total_seconds() / 3600, 2)
    df_merged['Horas_Ejec'] = np.round((df_merged['Fin_Ejec'] - df_merged['Inicio_Ejec']).dt.total_seconds() / 3600, 2)
    df_merged['Desviacion_Horas'] = np.round(df_merged['Horas_Ejec'].fillna(0) - df_merged['Horas_Prog'].fillna(0), 2)
    
    for col in ['Inicio_Prog', 'Fin_Prog', 'Inicio_Ejec', 'Fin_Ejec']:
        df_merged[col] = df_merged[col].dt.strftime('%d/%m/%Y %H:%M')
        df_merged[col] = df_merged[col].fillna('-')

    df_merged['MW_Indisponibles_Prog'] = pd.to_numeric(df_merged['MW_Indisponibles_Prog'], errors='coerce').fillna(0)
    df_merged['MW_Indisponibles_Ejec'] = pd.to_numeric(df_merged['MW_Indisponibles_Ejec'], errors='coerce').fillna(0)
    df_merged['Desviacion_MW'] = df_merged['MW_Indisponibles_Ejec'] - df_merged['MW_Indisponibles_Prog']

    return df_merged

# --- 5. INTERFAZ GRÁFICA Y MANEJO DE ESTADO ---
st.sidebar.header("Informe de Programación de Mantenimientos")

hoy = datetime.today()
rango_fechas = st.sidebar.date_input("Seleccione Rango de Fechas Operativas", value=(hoy, hoy))

if len(rango_fechas) == 2:
    fecha_inicio, fecha_fin = rango_fechas
else:
    fecha_inicio = fecha_fin = rango_fechas[0]

if 'dashboard_activo' not in st.session_state:
    st.session_state.dashboard_activo = False

if st.sidebar.button("Procesar Información"):
    st.session_state.dashboard_activo = True

if st.session_state.dashboard_activo:
    with st.spinner("Compilando bases de datos y sincronizando métricas operativas..."):
        df_prog_raw, df_ejec_raw, df_rf_raw, df_cmg_raw = obtener_datos_rango(fecha_inicio, fecha_fin)
        
        if not df_prog_raw.empty or not df_ejec_raw.empty:
            df_conciliado = conciliar_datos(df_prog_raw.copy(), df_ejec_raw.copy())
            
            st.markdown("### 🎛️ Filtros Dinámicos")
            col_f1, col_f2, col_f3, col_f4, col_f5 = st.columns(5)
            
            empresas_disp = sorted(df_conciliado[df_conciliado['Empresa'] != 'NAN']['Empresa'].dropna().unique())
            empresa_sel = col_f1.multiselect("Empresa Concesionaria:", empresas_disp, default=[])
            
            centrales_disp = sorted(df_conciliado[df_conciliado['Central/Ubicacion'] != '-']['Central/Ubicacion'].dropna().unique())
            central_sel = col_f2.multiselect("Central / Ubicación:", centrales_disp, default=[])
            
            sectores_disp = sorted(df_conciliado['Sector'].unique())
            default_sector = ['GENERACIÓN'] if 'GENERACIÓN' in sectores_disp else []
            sector_sel = col_f3.multiselect("Sector (Gen/Trans):", sectores_disp, default=default_sector)
            
            disp_equipos = sorted(df_conciliado[df_conciliado['Disponibilidad_Equipo'] != 'NAN']['Disponibilidad_Equipo'].unique())
            disp_sel = col_f4.multiselect("Estado (E/S o F/S):", disp_equipos, default=disp_equipos)
            
            tipos_disp = sorted(df_conciliado[df_conciliado['Tipo_Mantenimiento'] != 'NO ESPECIFICADO']['Tipo_Mantenimiento'].unique())
            tipo_sel = col_f5.multiselect("Tipo de Mantenimiento:", tipos_disp, default=[])
            
            df_filtrado = df_conciliado.copy()
            if len(empresa_sel) > 0: df_filtrado = df_filtrado[df_filtrado['Empresa'].isin(empresa_sel)]
            if len(central_sel) > 0: df_filtrado = df_filtrado[df_filtrado['Central/Ubicacion'].isin(central_sel)]
            if len(sector_sel) > 0: df_filtrado = df_filtrado[df_filtrado['Sector'].isin(sector_sel)]
            if len(disp_sel) > 0: df_filtrado = df_filtrado[df_filtrado['Disponibilidad_Equipo'].isin(disp_sel)]
            if len(tipo_sel) > 0: df_filtrado = df_filtrado[df_filtrado['Tipo_Mantenimiento'].isin(tipo_sel)]

            df_prog_raw_f = df_prog_raw.copy()
            if not df_prog_raw_f.empty:
                df_prog_raw_f['Empresa_Norm'] = normalizar_texto(df_prog_raw_f['Empresa'])
                df_prog_raw_f['Sector'] = df_prog_raw_f.apply(determinar_sector, axis=1)
                df_prog_raw_f['Disp_Norm'] = df_prog_raw_f['Disponibilidad_Equipo'].fillna('NO ESPECIFICADO').astype(str).str.upper().str.strip()
                df_prog_raw_f['Tipo_Norm'] = df_prog_raw_f['Tipo_Mantenimiento'].fillna('NO ESPECIFICADO').astype(str).str.upper().str.strip().replace('NAN', 'NO ESPECIFICADO')
                
                if len(empresa_sel) > 0: df_prog_raw_f = df_prog_raw_f[df_prog_raw_f['Empresa_Norm'].isin(empresa_sel)]
                if len(sector_sel) > 0: df_prog_raw_f = df_prog_raw_f[df_prog_raw_f['Sector'].isin(sector_sel)]
                if len(disp_sel) > 0: df_prog_raw_f = df_prog_raw_f[df_prog_raw_f['Disp_Norm'].isin(disp_sel)]
                if len(tipo_sel) > 0: df_prog_raw_f = df_prog_raw_f[df_prog_raw_f['Tipo_Norm'].isin(tipo_sel)]
                df_prog_raw_f = df_prog_raw_f.drop(columns=['Empresa_Norm', 'Sector', 'Disp_Norm', 'Tipo_Norm'])
                df_prog_raw_f.index = np.arange(1, len(df_prog_raw_f) + 1)

            df_ejec_raw_f = df_ejec_raw.copy()
            if not df_ejec_raw_f.empty:
                df_ejec_raw_f['Empresa_Norm'] = normalizar_texto(df_ejec_raw_f['Empresa'])
                df_ejec_raw_f['Sector'] = df_ejec_raw_f.apply(determinar_sector, axis=1)
                df_ejec_raw_f['Disp_Norm'] = df_ejec_raw_f['Disponibilidad_Equipo'].fillna('NO ESPECIFICADO').astype(str).str.upper().str.strip()
                df_ejec_raw_f['Tipo_Norm'] = df_ejec_raw_f['Tipo_Mantenimiento'].fillna('NO ESPECIFICADO').astype(str).str.upper().str.strip().replace('NAN', 'NO ESPECIFICADO')
                
                if len(empresa_sel) > 0: df_ejec_raw_f = df_ejec_raw_f[df_ejec_raw_f['Empresa_Norm'].isin(empresa_sel)]
                if len(sector_sel) > 0: df_ejec_raw_f = df_ejec_raw_f[df_ejec_raw_f['Sector'].isin(sector_sel)]
                if len(disp_sel) > 0: df_ejec_raw_f = df_ejec_raw_f[df_ejec_raw_f['Disp_Norm'].isin(disp_sel)]
                if len(tipo_sel) > 0: df_ejec_raw_f = df_ejec_raw_f[df_ejec_raw_f['Tipo_Norm'].isin(tipo_sel)]
                df_ejec_raw_f = df_ejec_raw_f.drop(columns=['Empresa_Norm', 'Sector', 'Disp_Norm', 'Tipo_Norm'])
                df_ejec_raw_f.index = np.arange(1, len(df_ejec_raw_f) + 1)

            st.markdown("---")

            # --- AGREGADO PESTAÑA 6 ---
            tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs([
                "📊 1. Resumen y Métricas",
                "✅ 2. MATCH: Tiempos",
                "⚠️ 3. Ejecutados NO Programados", 
                "❌ 4. Programados NO Ejecutados",
                "🗄️ 5. Datos Originales (Raw)",
                "🔌 6. Potencia Indisponible (Gen F/S)"
            ])
            
            # --- PESTAÑA 1: RESUMEN EJECUTIVO ---
            with tab1:
                st.header(f"📊 Informe de Supervisión del Sistema Interconectado Nacional")
                
                total_prog_raw_count = len(df_prog_raw_f) if df_prog_raw_f is not None else 0
                total_ejec_raw_count = len(df_ejec_raw_f) if df_ejec_raw_f is not None else 0
                
                total_match = len(df_filtrado[df_filtrado['Estado_Supervision'] == 'Programado y Ejecutado'])
                total_no_ejec = len(df_filtrado[df_filtrado['Estado_Supervision'] == 'Programado NO Ejecutado'])
                total_forzados = len(df_filtrado[df_filtrado['Estado_Supervision'] == 'Ejecutado NO Programado'])
                
                total_programados = total_match + total_no_ejec
                total_ejecutados = total_match + total_forzados
                total_universo = len(df_filtrado)
                
                desviacion_neta = df_filtrado['Desviacion_Horas'].sum()
                porcentaje_ejec_prog = (total_match / total_ejecutados * 100) if total_ejecutados > 0 else 0
                
                texto_datos_criticos = ""
                df_ejecutados_total = df_filtrado[df_filtrado['Estado_Supervision'].isin(['Programado y Ejecutado', 'Ejecutado NO Programado'])]
                if not df_ejecutados_total.empty:
                    empresa_max_registros = df_ejecutados_total['Empresa'].value_counts().idxmax()
                    max_registros = df_ejecutados_total['Empresa'].value_counts().max()
                    
                    df_empresa_horas = df_ejecutados_total.groupby('Empresa')['Horas_Ejec'].sum().reset_index()
                    empresa_max_horas = df_empresa_horas.loc[df_empresa_horas['Horas_Ejec'].idxmax(), 'Empresa']
                    max_horas = df_empresa_horas['Horas_Ejec'].max()
                    
                    texto_datos_criticos = f"\n\n🚨 *Datos Críticos:* La empresa que registró la mayor cantidad de intervenciones ejecutadas fue **{empresa_max_registros}** ({max_registros} maniobras). Asimismo, la empresa que acumuló el mayor tiempo operativo de mantenimiento fue **{empresa_max_horas}** con un total de **{max_horas:.2f} horas**."

                texto_diagnostico = f"""
                **📌 Resumen Ejecutivo de Operaciones:** En la ventana de supervisión, los documentos del COES reportaron un consolidado de **{total_prog_raw_count} mantenimientos programados** y **{total_ejec_raw_count} mantenimientos ejecutados** (bajo los filtros aplicados). 
                Se determinó que el **{porcentaje_ejec_prog:.1f}% de los mantenimientos ejecutados fueron programados previamente**. 
                Asimismo, se registraron **{total_forzados} mantenimientos ejecutados no programados** y **{total_no_ejec} mantenimientos programados no ejecutados**.{texto_datos_criticos}
                """
                st.info(texto_diagnostico)
                
                st.markdown("*💡 **Nota sobre la Desviación Neta (Horas):*** Un valor **positivo (+)** indica un **retraso neto** en el sistema (las maniobras tomaron más tiempo del planificado), mientras que un valor **negativo (-)** indica un **ahorro operativo** (las unidades retornaron al servicio antes de lo previsto).")
                
                c_doc1, c_doc2, c_doc3 = st.columns(3)
                c_doc1.metric("Mantenimientos Programados", total_prog_raw_count, help="Volumen del Anexo Osinergmin (Programados y Cancelados).")
                c_doc2.metric("Mantenimientos Ejecutados", total_ejec_raw_count, help="Volumen del Anexo A (Programados/Ejecutados y Forzados).")
                c_doc3.metric("Universo Total Único", total_universo, delta="Eventos unificados", delta_color="normal")
                
                c_res1, c_res2, c_res3, c_res4 = st.columns(4)
                c_res1.metric("Programado y Ejecutado", total_match, help="Cumplieron con planificar y ejecutar.")
                c_res2.metric("Ejecutado NO Programado", total_forzados, help="Eventos intempestivos o de emergencia.")
                c_res3.metric("Programado NO Ejecutado", total_no_ejec, help="Se programaron pero no se realizaron.")
                c_res4.metric("🎯 Ejecutados que fueron Programados", f"{porcentaje_ejec_prog:.1f}%", help="Porcentaje de la ejecución total que sí estuvo planificada.")
                
                st.markdown("<br>", unsafe_allow_html=True)
                if desviacion_neta > 0 or total_forzados > 0:
                    st.warning(f"**⚠️ Alerta de Supervisión:** Se registraron **{total_forzados}** mantenimientos ejecutados no programados y un retraso neto de **{desviacion_neta:.2f} horas** en el sistema operativo.")
                else:
                    st.success("**✅ Operación Óptima:** Las empresas han operado respetando los márgenes de tiempo estipulados y no se detectaron ejecuciones sin programación.")
                
                col_g1, col_g2 = st.columns(2)
                with col_g1:
                    df_estado_counts = df_filtrado['Estado_Supervision'].value_counts().reset_index()
                    df_estado_counts.columns = ['Estado Operativo', 'Cantidad']
                    fig_pie_univ = px.pie(df_estado_counts, values='Cantidad', names='Estado Operativo', hole=0.4, title="Distribución de Estados Operativos del SEIN", color='Estado Operativo', color_discrete_map={'Programado y Ejecutado': '#2ca02c', 'Ejecutado NO Programado': '#d62728', 'Programado NO Ejecutado': '#ff7f0e'})
                    fig_pie_univ.update_traces(textinfo='value+percent', textfont_size=14, hoverinfo='label+percent+value')
                    st.plotly_chart(fig_pie_univ, use_container_width=True)

                with col_g2:
                    df_excesos = df_filtrado[df_filtrado['Desviacion_Horas'] > 0]
                    if not df_excesos.empty:
                        df_agrupado = df_excesos.groupby('Empresa')['Desviacion_Horas'].sum().reset_index()
                        df_agrupado = df_agrupado.sort_values('Desviacion_Horas', ascending=False).head(10)
                        fig_bar_tiempo = px.bar(df_agrupado, x='Desviacion_Horas', y='Empresa', orientation='h', title="Top 10 Empresas con Mayor Retraso en Ejecución (Horas)", text_auto='.2f', color='Desviacion_Horas', color_continuous_scale='Reds')
                        fig_bar_tiempo.update_layout(yaxis={'categoryorder':'total ascending'})
                        fig_bar_tiempo.update_traces(textposition='outside', textfont_size=12)
                        st.plotly_chart(fig_bar_tiempo, use_container_width=True)

                st.markdown("---")
                
                st.markdown("#### 📅 3. Cronograma de Indisponibilidades Operativas (Gantt)")
                
                df_gantt = df_filtrado[df_filtrado['Estado_Supervision'].isin(['Programado y Ejecutado', 'Ejecutado NO Programado'])].copy()
                
                if not df_gantt.empty:
                    df_gantt['Inicio_DT'] = pd.to_datetime(df_gantt['Inicio_Ejec'], format='%d/%m/%Y %H:%M', errors='coerce')
                    df_gantt['Fin_DT'] = pd.to_datetime(df_gantt['Fin_Ejec'], format='%d/%m/%Y %H:%M', errors='coerce')
                    df_gantt = df_gantt.dropna(subset=['Inicio_DT', 'Fin_DT'])
                    
                    if not df_gantt.empty:
                        df_gantt['Central_Unidad'] = df_gantt['Central/Ubicacion'] + " | " + df_gantt['Equipo']
                        num_y_items = len(df_gantt['Central_Unidad'].unique())
                        altura_dinamica = max(400, num_y_items * 35) 
                        
                        fig_gantt = px.timeline(
                            df_gantt, 
                            x_start="Inicio_DT", 
                            x_end="Fin_DT", 
                            y="Central_Unidad", 
                            color="Disponibilidad_Equipo",
                            hover_name="Empresa",
                            hover_data={"Tipo_Mantenimiento": True, "Estado_Supervision": True, "Disponibilidad_Equipo": False},
                            title="Línea de Tiempo por Estado (E/S - F/S)",
                            color_discrete_map={'F/S': '#d62728', 'E/S': '#2ca02c', 'NO ESPECIFICADO': '#7f7f7f'},
                            height=altura_dinamica 
                        )
                        fig_gantt.update_yaxes(autorange="reversed") 
                        fig_gantt.update_layout(xaxis_title="Fechas y Horas de Operación", yaxis_title="Centrales / Unidades")
                        st.plotly_chart(fig_gantt, use_container_width=True)
                    else:
                        st.info("⚠️ Los registros de ejecución actuales no poseen un formato de fecha y hora válido para diagramar el cronograma.")
                else:
                    st.info("⚠️ No hay eventos ejecutados bajo los filtros actuales para graficar el Cronograma.")

                st.markdown("---")
                
                st.markdown("#### ⏱️ 4. Análisis de Tiempos Exclusivo (Programado y Ejecutado)")
                df_match_kpi = df_filtrado[df_filtrado['Estado_Supervision'] == 'Programado y Ejecutado'].copy()
                
                if not df_match_kpi.empty:
                    df_match_kpi['Desempeño_Tiempo'] = np.where(df_match_kpi['Desviacion_Horas'] > 0, 'Excedieron Programación (Retraso)', 
                                                               np.where(df_match_kpi['Desviacion_Horas'] < 0, 'Terminaron Antes (Ahorro)', 'Ejecución Exacta'))
                    
                    col_t1, col_t2, col_t3, col_t4 = st.columns(4)
                    cant_mayor = len(df_match_kpi[df_match_kpi['Desviacion_Horas'] > 0])
                    cant_menor = len(df_match_kpi[df_match_kpi['Desviacion_Horas'] < 0])
                    cant_exacto = len(df_match_kpi[df_match_kpi['Desviacion_Horas'] == 0])
                    
                    if desviacion_neta > 0: col_t1.metric("Desviación Neta (Horas)", f"+{desviacion_neta:.2f} h", delta="Retraso Acumulado", delta_color="inverse")
                    elif desviacion_neta < 0: col_t1.metric("Desviación Neta (Horas)", f"{desviacion_neta:.2f} h", delta="Ahorro Operativo", delta_color="normal")
                    else: col_t1.metric("Desviación Neta (Horas)", "0.00 h", delta="Sincronización Exacta", delta_color="off")
                    
                    col_t2.metric("Excedieron Programación", cant_mayor)
                    col_t3.metric("Terminaron Antes", cant_menor)
                    col_t4.metric("Ejecución Exacta", cant_exacto)

                    col_gt1, col_gt2 = st.columns(2)
                    with col_gt1:
                        tiempos_counts = df_match_kpi['Desempeño_Tiempo'].value_counts().reset_index()
                        tiempos_counts.columns = ['Eficiencia', 'Cantidad']
                        fig_pie_tiempos = px.pie(tiempos_counts, values='Cantidad', names='Eficiencia', hole=0.4, title="Eficiencia en Tiempos de Maniobra", color='Eficiencia', color_discrete_map={'Excedieron Programación (Retraso)': '#d62728', 'Terminaron Antes (Ahorro)': '#2ca02c', 'Ejecución Exacta': '#1f77b4'})
                        fig_pie_tiempos.update_traces(textinfo='value+percent', textfont_size=14, hoverinfo='label+percent+value')
                        st.plotly_chart(fig_pie_tiempos, use_container_width=True)
                        
                    with col_gt2:
                        df_sector_tiempos = df_match_kpi.groupby('Sector')[['Horas_Prog', 'Horas_Ejec']].sum().reset_index()
                        df_sector_tiempos_melted = df_sector_tiempos.melt(id_vars='Sector', value_vars=['Horas_Prog', 'Horas_Ejec'], var_name='Tipo', value_name='Horas')
                        df_sector_tiempos_melted['Tipo'] = df_sector_tiempos_melted['Tipo'].replace({'Horas_Prog': 'H. Programadas', 'Horas_Ejec': 'H. Ejecutadas'})
                        fig_bar_tiempos = px.bar(df_sector_tiempos_melted, x='Sector', y='Horas', color='Tipo', barmode='group', title="Volumen de Horas Operativas por Sector", color_discrete_map={'H. Programadas': '#1f77b4', 'H. Ejecutadas': '#ff7f0e'}, text_auto='.1f')
                        fig_bar_tiempos.update_traces(textposition='outside', textfont_size=12)
                        st.plotly_chart(fig_bar_tiempos, use_container_width=True)
                else:
                    st.info("No hay mantenimientos ejecutados bajo programación (Match) para generar métricas de eficiencia.")
                
                st.markdown("---")
                
                st.markdown("#### 🛠️ 5. Estadísticas Operativas por Tipo de Mantenimiento")
                col_tm1, col_tm2 = st.columns(2)
                with col_tm1:
                    tipo_counts = df_filtrado['Tipo_Mantenimiento'].value_counts().reset_index()
                    tipo_counts.columns = ['Tipo de Mantenimiento', 'Cantidad de Registros']
                    fig_pie_tipo = px.pie(tipo_counts, values='Cantidad de Registros', names='Tipo de Mantenimiento', hole=0.4, title="Volumen de Maniobras por Tipo (Preventivo/Correctivo)", color_discrete_sequence=px.colors.qualitative.Prism)
                    fig_pie_tipo.update_traces(textinfo='value+percent', textfont_size=14, hoverinfo='label+percent+value')
                    st.plotly_chart(fig_pie_tipo, use_container_width=True)

                with col_tm2:
                    df_tipo_horas = df_filtrado.groupby('Tipo_Mantenimiento')[['Horas_Prog', 'Horas_Ejec']].sum().reset_index()
                    df_tipo_horas_melted = df_tipo_horas.melt(id_vars='Tipo_Mantenimiento', value_vars=['Horas_Prog', 'Horas_Ejec'], var_name='Fase', value_name='Total de Horas')
                    df_tipo_horas_melted['Fase'] = df_tipo_horas_melted['Fase'].replace({'Horas_Prog': 'H. Programadas', 'Horas_Ejec': 'H. Ejecutadas'})
                    fig_bar_tipo = px.bar(df_tipo_horas_melted, x='Tipo_Mantenimiento', y='Total de Horas', color='Fase', barmode='group', title="Consumo de Horas (Prog vs Ejec) por Tipo de Maniobra", color_discrete_map={'H. Programadas': '#1f77b4', 'H. Ejecutadas': '#ff7f0e'}, text_auto='.1f')
                    fig_bar_tipo.update_traces(textposition='outside', textfont_size=12)
                    st.plotly_chart(fig_bar_tipo, use_container_width=True)

            # --- PESTAÑA 2: MATCH ---
            with tab2:
                st.subheader("✅ Mantenimientos Programados y Ejecutados")
                st.caption("Esta vista muestra el cruce exacto de los eventos del COES.")
                df_tab2 = df_filtrado[df_filtrado['Estado_Supervision'] == 'Programado y Ejecutado'].copy()
                if not df_tab2.empty:
                    df_tab2['Estado_Tiempo'] = np.where(df_tab2['Desviacion_Horas'] > 0, '🔴 Mayor Tiempo', 
                                                       np.where(df_tab2['Desviacion_Horas'] < 0, '🟠 Menor Tiempo', '🟢 Tiempo Exacto'))
                    columnas_tab2 = ['Fecha_Operacion', 'Estado_Tiempo', 'Empresa', 'Central/Ubicacion', 'Equipo', 'Sector', 'Tipo_Mantenimiento', 'Disponibilidad_Equipo', 
                                     'Inicio_Prog', 'Fin_Prog', 'Horas_Prog', 
                                     'Inicio_Ejec', 'Fin_Ejec', 'Horas_Ejec', 'Desviacion_Horas',
                                     'MW_Indisponibles_Prog', 'MW_Indisponibles_Ejec', 'Desviacion_MW', 'Descripcion_Ejec']
                    df_tab2_show = df_tab2[columnas_tab2].copy()
                    df_tab2_show.index = np.arange(1, len(df_tab2_show) + 1) 
                    st.dataframe(df_tab2_show, use_container_width=True)
                else:
                    st.info("No hay coincidencias perfectas bajo los filtros actuales.")

            # --- PESTAÑA 3: FORZADOS ---
            with tab3:
                st.subheader("⚠️ Mantenimientos Ejecutados NO Programados")
                df_tab3 = df_filtrado[df_filtrado['Estado_Supervision'] == 'Ejecutado NO Programado'].copy()
                if not df_tab3.empty:
                    columnas_tab3 = ['Fecha_Operacion', 'Empresa', 'Central/Ubicacion', 'Equipo', 'Sector', 'Tipo_Mantenimiento', 'Disponibilidad_Equipo', 'Inicio_Ejec', 'Fin_Ejec', 'Horas_Ejec', 'MW_Indisponibles_Ejec', 'Descripcion_Ejec']
                    df_tab3_show = df_tab3[columnas_tab3].copy()
                    df_tab3_show.index = np.arange(1, len(df_tab3_show) + 1) 
                    st.dataframe(df_tab3_show, use_container_width=True)
                else:
                    st.info("No hay eventos forzados bajo los filtros actuales.")

            # --- PESTAÑA 4: NO EJECUTADOS ---
            with tab4:
                st.subheader("❌ Mantenimientos Programados NO Ejecutados")
                df_tab4 = df_filtrado[df_filtrado['Estado_Supervision'] == 'Programado NO Ejecutado'].copy()
                if not df_tab4.empty:
                    columnas_tab4 = ['Fecha_Operacion', 'Empresa', 'Central/Ubicacion', 'Equipo', 'Sector', 'Tipo_Mantenimiento', 'Disponibilidad_Equipo', 'Inicio_Prog', 'Fin_Prog', 'Horas_Prog', 'MW_Indisponibles_Prog', 'Descripcion_Prog']
                    df_tab4_show = df_tab4[columnas_tab4].copy()
                    df_tab4_show.index = np.arange(1, len(df_tab4_show) + 1) 
                    st.dataframe(df_tab4_show, use_container_width=True)
                else:
                    st.info("No hay eventos cancelados bajo los filtros actuales.")

            # --- PESTAÑA 5: RAW DATA ---
            with tab5:
                st.subheader("🗄️ Trazabilidad: Archivos Crudos (Raw Data)")
                col_raw1, col_raw2 = st.columns(2)
                with col_raw1:
                    st.markdown(f"**📁 Anexo Osinergmin (Programado Diario) - {total_prog_raw_count} Registros**")
                    if not df_prog_raw_f.empty:
                        st.dataframe(df_prog_raw_f, use_container_width=True)
                    else:
                        st.info("Sin registros tras aplicar filtros.")
                with col_raw2:
                    st.markdown(f"**📁 Anexo A (Mantenimientos Ejecutados) - {total_ejec_raw_count} Registros**")
                    if not df_ejec_raw_f.empty:
                        st.dataframe(df_ejec_raw_f, use_container_width=True)
                    else:
                        st.info("Sin registros tras aplicar filtros.")
                        
            # --- PESTAÑA 6: POTENCIA INDISPONIBLE, RESERVA Y CMg ---
            with tab6:
                st.subheader("🔌 Gestión de Potencia Indisponible - Generación (F/S)")
                st.caption("Módulo de asignación para la Potencia Restada al SEIN. La base se autocompleta con el repositorio histórico de Osinergmin.")

                df_gen_fs = df_filtrado[
                    (df_filtrado['Sector'] == 'GENERACIÓN') &
                    (df_filtrado['Disponibilidad_Equipo'] == 'F/S') &
                    (df_filtrado['Estado_Supervision'].isin(['Programado y Ejecutado', 'Ejecutado NO Programado']))
                ].copy()

                if not df_gen_fs.empty:
                    df_unidades = df_gen_fs[['Empresa', 'Central/Ubicacion', 'Equipo']].drop_duplicates().reset_index(drop=True)
                    df_historico_pot = cargar_potencias_guardadas()
                    df_editor = pd.merge(df_unidades, df_historico_pot[['Central/Ubicacion', 'Equipo', 'Potencia_Indisponible_MW']], on=['Central/Ubicacion', 'Equipo'], how='left')
                    df_editor['Potencia_Indisponible_MW'] = df_editor['Potencia_Indisponible_MW'].fillna(0.0)

                    st.markdown("#### 📝 Asignación de Potencia Indisponible")
                    df_editado = st.data_editor(
                        df_editor,
                        column_config={
                            "Empresa": st.column_config.TextColumn("Empresa Concesionaria", disabled=True),
                            "Central/Ubicacion": st.column_config.TextColumn("Central de Generación", disabled=True),
                            "Equipo": st.column_config.TextColumn("Unidad/Equipo F/S", disabled=True),
                            "Potencia_Indisponible_MW": st.column_config.NumberColumn(
                                "Potencia Indisponible (MW)", min_value=0.0, format="%.2f"
                            )
                        },
                        use_container_width=True, hide_index=True, key="editor_mw"
                    )

                    if st.button("💾 Guardar Potencias Asignadas al Repositorio", type="primary"):
                        guardar_potencias_asignadas(df_editado)
                        st.success("¡Valores almacenados en el búfer con éxito!")

                    st.markdown("---")
                    st.markdown("#### 📈 Impacto Operativo y Perfil de Indisponibilidad")
                    
                    df_grafica_base = pd.merge(df_gen_fs[['Empresa', 'Central/Ubicacion', 'Equipo', 'Inicio_Ejec', 'Fin_Ejec', 'Horas_Ejec']], 
                                          df_editado, on=['Empresa', 'Central/Ubicacion', 'Equipo'], how='inner')
                    
                    if not df_grafica_base.empty and df_grafica_base['Potencia_Indisponible_MW'].sum() > 0:
                        col_filt_1, col_filt_2 = st.columns(2)
                        empresas_graf = sorted(df_grafica_base['Empresa'].unique())
                        empresa_graf_sel = col_filt_1.multiselect("Filtrar Análisis por Empresa:", empresas_graf, default=[], key="flt_empresa_impacto")
                        
                        if empresa_graf_sel:
                            centrales_graf = sorted(df_grafica_base[df_grafica_base['Empresa'].isin(empresa_graf_sel)]['Central/Ubicacion'].unique())
                        else:
                            centrales_graf = sorted(df_grafica_base['Central/Ubicacion'].unique())
                            
                        central_graf_sel = col_filt_2.multiselect("Filtrar Análisis por Central/Ubicación:", centrales_graf, default=[], key="flt_central_impacto")
                        
                        df_grafica = df_grafica_base.copy()
                        if empresa_graf_sel: df_grafica = df_grafica[df_grafica['Empresa'].isin(empresa_graf_sel)]
                        if central_graf_sel: df_grafica = df_grafica[df_grafica['Central/Ubicacion'].isin(central_graf_sel)]
                        
                        if not df_grafica.empty:
                            df_grafica['Inicio_DT'] = pd.to_datetime(df_grafica['Inicio_Ejec'], format='%d/%m/%Y %H:%M', errors='coerce')
                            df_grafica['Fin_DT'] = pd.to_datetime(df_grafica['Fin_Ejec'], format='%d/%m/%Y %H:%M', errors='coerce')
                            df_grafica = df_grafica.dropna(subset=['Inicio_DT', 'Fin_DT'])
                            df_grafica['Central_Equipo'] = df_grafica['Central/Ubicacion'] + " - " + df_grafica['Equipo']
                            
                            if not df_grafica.empty:
                                min_dt = df_grafica['Inicio_DT'].min()
                                max_dt = df_grafica['Fin_DT'].max()
                                
                                if min_dt < max_dt:
                                    time_grid = pd.date_range(start=min_dt, end=max_dt, freq='h')
                                    series_list, series_detail_list, series_eq_list = [], [], []
                                    
                                    centrales_involucradas = df_grafica['Central/Ubicacion'].unique()
                                    equipos_involucrados = df_grafica['Central_Equipo'].unique()
                                    
                                    for t in time_grid:
                                        mask = (df_grafica['Inicio_DT'] <= t) & (df_grafica['Fin_DT'] > t)
                                        df_t = df_grafica.loc[mask]
                                        df_t_unique = df_t.drop_duplicates(subset=['Central_Equipo'])
                                        
                                        centrales_completas = df_t_unique[df_t_unique['Equipo'] == 'CENTRAL'][['Empresa', 'Central/Ubicacion']].drop_duplicates()
                                        for _, row in centrales_completas.iterrows():
                                            condicion_remover = (df_t_unique['Empresa'] == row['Empresa']) & (df_t_unique['Central/Ubicacion'] == row['Central/Ubicacion']) & (df_t_unique['Equipo'] != 'CENTRAL')
                                            df_t_unique = df_t_unique[~condicion_remover]
                                        
                                        mw_sum = df_t_unique['Potencia_Indisponible_MW'].sum()
                                        series_list.append({'Fecha_Hora': t, 'MW_Total_Indisponible': mw_sum})
                                        
                                        agg_central = df_t_unique.groupby('Central/Ubicacion')['Potencia_Indisponible_MW'].sum()
                                        for c in centrales_involucradas:
                                            series_detail_list.append({
                                                'Fecha_Hora': t, 'Central': c, 'MW_Indisponible': agg_central.get(c, 0.0), 'MW_Total_Sistema': mw_sum
                                            })
                                            
                                        agg_eq = df_t_unique.groupby('Central_Equipo')['Potencia_Indisponible_MW'].sum()
                                        for eq in equipos_involucrados:
                                            series_eq_list.append({
                                                'Fecha_Hora': t, 'Central_Equipo': eq, 'MW_Indisponible': agg_eq.get(eq, 0.0), 'MW_Total_Sistema': mw_sum
                                            })
                                    
                                    df_area = pd.DataFrame(series_list)
                                    df_area_detail = pd.DataFrame(series_detail_list)
                                    df_area_eq = pd.DataFrame(series_eq_list)
                                    
                                    max_potencia_aislada = df_grafica['Potencia_Indisponible_MW'].max()
                                    energia_ns_total = df_area['MW_Total_Indisponible'].sum()
                                    unidades_afectadas = len(df_grafica[df_grafica['Potencia_Indisponible_MW'] > 0]['Central_Equipo'].unique())
                                    
                                    col_st1, col_st2, col_st3 = st.columns(3)
                                    col_st1.metric("Máxima Potencia Unitaria F/S", f"{max_potencia_aislada:.2f} MW")
                                    col_st2.metric("Equipos en Indisponibilidad", f"{unidades_afectadas} unidades")
                                    col_st3.metric("Energía No Suministrada (Est.)", f"{energia_ns_total:,.2f} MWh")

                                    colores_solidos_centrales = px.colors.qualitative.Vivid + px.colors.qualitative.Dark24
                                    colores_solidos_equipos = px.colors.qualitative.Alphabet + px.colors.qualitative.Dark24 + px.colors.qualitative.Set1
                                    
                                    fig_area = px.area(df_area, x='Fecha_Hora', y='MW_Total_Indisponible', 
                                                       title="Perfil Evolutivo de Potencia Indisponible Global (MW) del SEIN",
                                                       color_discrete_sequence=['#d62728'])
                                    fig_area.update_xaxes(tickmode='linear', dtick=86400000, tickformat="%d/%m/%Y", title_text="Fecha de Operación")
                                    fig_area.update_yaxes(title_text="Demanda Indisponible (MW)")
                                    fig_area.update_traces(line=dict(width=0), hovertemplate="<b>%{y:,.2f} MW</b>")
                                    st.plotly_chart(fig_area, use_container_width=True)
                                    
                                    fig_area_detail = px.area(df_area_detail, x='Fecha_Hora', y='MW_Indisponible', color='Central',
                                                       title="Desglose de Potencia Indisponible (MW) por Central de Generación",
                                                       color_discrete_sequence=colores_solidos_centrales,
                                                       hover_data={'Fecha_Hora': '|%d/%m/%Y %H:%M', 'MW_Indisponible': ':.2f', 'MW_Total_Sistema': ':.2f'})
                                    fig_area_detail.update_xaxes(tickmode='linear', dtick=86400000, tickformat="%d/%m/%Y", title_text="Fecha de Operación")
                                    fig_area_detail.update_yaxes(title_text="Demanda Indisponible (MW)")
                                    fig_area_detail.update_traces(line=dict(width=0))
                                    st.plotly_chart(fig_area_detail, use_container_width=True)
                                    
                                    fig_area_eq = px.area(df_area_eq, x='Fecha_Hora', y='MW_Indisponible', color='Central_Equipo',
                                                       title="Desglose Extendido de Potencia Indisponible (MW) por Unidad/Equipo",
                                                       color_discrete_sequence=colores_solidos_equipos,
                                                       hover_data={'Fecha_Hora': '|%d/%m/%Y %H:%M', 'MW_Indisponible': ':.2f', 'MW_Total_Sistema': ':.2f'})
                                    fig_area_eq.update_xaxes(tickmode='linear', dtick=86400000, tickformat="%d/%m/%Y", title_text="Fecha de Operación")
                                    fig_area_eq.update_yaxes(title_text="Demanda Indisponible (MW)")
                                    fig_area_eq.update_traces(line=dict(width=0))
                                    st.plotly_chart(fig_area_eq, use_container_width=True)

                                    st.markdown("#### 🗃️ Base de Datos Analítica: Perfil de Indisponibilidad")
                                    df_pivot = df_area_detail.pivot(index='Fecha_Hora', columns='Central', values='MW_Indisponible').reset_index()
                                    df_pivot['Total_SEIN (MW)'] = df_pivot.drop(columns=['Fecha_Hora']).sum(axis=1)
                                    df_pivot['Fecha_Hora'] = df_pivot['Fecha_Hora'].dt.strftime('%d/%m/%Y %H:%M')
                                    st.dataframe(df_pivot, use_container_width=True)
                                else:
                                    st.info("El intervalo de tiempo es demasiado estrecho para construir el área continua.")
                        else:
                            st.info("Los filtros aplicados no arrojaron resultados para graficar el perfil.")
                else:
                    st.info("No se registraron maniobras EJECUTADAS en el sector GENERACIÓN con un estado operativo Fuera de Servicio (F/S).")

                # ========================================================
                # INYECCIÓN DE SECCIÓN: RESERVA FRÍA Y EFICIENTE
                # ========================================================
                st.markdown("---")
                st.markdown("### ❄️ Fiscalización de Reserva Operativa del SEIN")
                
                if df_rf_raw is None or df_rf_raw.empty:
                    st.warning("⚠️ No se encontró información de Reserva Fría para las fechas seleccionadas.")
                else:
                    df_rf = df_rf_raw.copy()
                    fecha_min_rf = df_rf['FECHA_HORA'].min()
                    fecha_max_rf = df_rf['FECHA_HORA'].max()
                    
                    st.markdown("#### 1. Disponibilidad de Reserva Fría Total (Ajustada por TV/TG)")
                    limite_superior_rf = df_rf['RESERVA_FRIA_MW'].max() * 1.10

                    fig_rf = px.area(
                        df_rf, x="FECHA_HORA", y="RESERVA_FRIA_MW", 
                        title="Curva de Reserva Fría Total (MW)",
                        color_discrete_sequence=["#00BFFF"] 
                    )
                    fig_rf.update_traces(hovertemplate="<b>%{y:,.2f} MW</b>", line=dict(width=0))
                    fig_rf.update_layout(
                        hovermode="x unified",
                        xaxis=dict(tickformat="%d/%m\n%H:%M", title="Fecha Operativa", range=[fecha_min_rf, fecha_max_rf], tickmode="linear", dtick=86400000),
                        yaxis=dict(title="Reserva Total (MW)", range=[0, limite_superior_rf]),
                        height=350, margin=dict(t=30, b=40, l=50, r=20)
                    )
                    st.plotly_chart(fig_rf, use_container_width=True)
                    
                    col_rf1, col_rf2, col_rf3 = st.columns(3)
                    col_rf1.metric("Promedio - Reserva Fría", f"{df_rf['RESERVA_FRIA_MW'].mean():.2f} MW")
                    col_rf2.metric("Máxima - Reserva Fría", f"{df_rf['RESERVA_FRIA_MW'].max():.2f} MW")
                    col_rf3.metric("Mínima - Reserva Fría", f"{df_rf['RESERVA_FRIA_MW'].min():.2f} MW")
                    
                    st.markdown("---")

                    st.markdown("#### 2. Disponibilidad de Reserva Eficiente (Descontando Restricciones)")
                    limite_superior_ef = df_rf['RESERVA_EFICIENTE_MW'].max() * 1.10

                    fig_ef = px.area(
                        df_rf, x="FECHA_HORA", y="RESERVA_EFICIENTE_MW", 
                        title="Curva de Reserva Eficiente (MW)",
                        color_discrete_sequence=["#32CD32"]
                    )
                    fig_ef.update_traces(hovertemplate="<b>%{y:,.2f} MW</b>", line=dict(width=0))
                    fig_ef.update_layout(
                        hovermode="x unified",
                        xaxis=dict(tickformat="%d/%m\n%H:%M", title="Fecha Operativa", range=[fecha_min_rf, fecha_max_rf], tickmode="linear", dtick=86400000),
                        yaxis=dict(title="Reserva Eficiente (MW)", range=[0, limite_superior_ef]),
                        height=350, margin=dict(t=30, b=40, l=50, r=20)
                    )
                    st.plotly_chart(fig_ef, use_container_width=True)
                    
                    col_ef1, col_ef2, col_ef3 = st.columns(3)
                    col_ef1.metric("Promedio - Reserva Eficiente", f"{df_rf['RESERVA_EFICIENTE_MW'].mean():.2f} MW")
                    col_ef2.metric("Máxima - Reserva Eficiente", f"{df_rf['RESERVA_EFICIENTE_MW'].max():.2f} MW")
                    col_ef3.metric("Mínima - Reserva Eficiente", f"{df_rf['RESERVA_EFICIENTE_MW'].min():.2f} MW")

                    st.markdown("---")

                    st.markdown("#### 3. Disponibilidad de Reserva Fría y Nodo Energético")
                    limite_superior_nodo = df_rf['RESERVA_NODO_MW'].max() * 1.10

                    fig_nodo = px.area(
                        df_rf, x="FECHA_HORA", y="RESERVA_NODO_MW", 
                        title="Curva de Reserva Fría y Nodo Energético (MW)",
                        color_discrete_sequence=["#FF8C00"] 
                    )
                    fig_nodo.update_traces(hovertemplate="<b>%{y:,.2f} MW</b>", line=dict(width=0))
                    fig_nodo.update_layout(
                        hovermode="x unified",
                        xaxis=dict(tickformat="%d/%m\n%H:%M", title="Fecha Operativa", range=[fecha_min_rf, fecha_max_rf], tickmode="linear", dtick=86400000),
                        yaxis=dict(title="Reserva Nodo (MW)", range=[0, limite_superior_nodo]),
                        height=350, margin=dict(t=30, b=40, l=50, r=20)
                    )
                    st.plotly_chart(fig_nodo, use_container_width=True)
                    
                    col_n1, col_n2, col_n3 = st.columns(3)
                    col_n1.metric("Promedio - Reserva Nodo", f"{df_rf['RESERVA_NODO_MW'].mean():.2f} MW")
                    col_n2.metric("Máxima - Reserva Nodo", f"{df_rf['RESERVA_NODO_MW'].max():.2f} MW")
                    col_n3.metric("Mínima - Reserva Nodo", f"{df_rf['RESERVA_NODO_MW'].min():.2f} MW")

                # ========================================================
                # INYECCIÓN DE SECCIÓN: COSTOS MARGINALES
                # ========================================================
                st.markdown("---")
                st.markdown("### 💰 Evolución de Costos Marginales de Corto Plazo (CMg)")
                
                if df_cmg_raw is None or df_cmg_raw.empty:
                    st.warning("⚠️ No se encontraron archivos de Costos Marginales (.zip) para las fechas seleccionadas o el formato ha cambiado.")
                else:
                    df_cmg = df_cmg_raw.copy()
                    fecha_min_cmg = df_cmg['FECHA_HORA'].min()
                    fecha_max_cmg = df_cmg['FECHA_HORA'].max()
                    
                    colores_nodos = {
                        "SANTA ROSA 220": "#E91E63", 
                        "TALARA 220": "#8B4513",     
                        "MOQUEGUA 220": "#4169E1"    
                    }

                    st.markdown("#### Dinámica de Precios (Soles/MWh) en Barras Estratégicas")
                    
                    fig_cmg = px.line(
                        df_cmg, 
                        x="FECHA_HORA", 
                        y="COSTO_MARGINAL_SOLES", 
                        color="NODO",
                        title="Curvas de Costo Marginal - SEIN (Norte, Centro, Sur)",
                        color_discrete_map=colores_nodos
                    )
                    
                    fig_cmg.update_traces(hovertemplate="<b>%{y:,.2f} Soles/MWh</b>", connectgaps=True)
                    
                    fig_cmg.update_layout(
                        hovermode="x unified",
                        xaxis=dict(tickformat="%d/%m\n%H:%M", title="Fecha Operativa", range=[fecha_min_cmg, fecha_max_cmg], tickmode="linear", dtick=86400000),
                        yaxis=dict(title="Costo Marginal (Soles/MWh)"),
                        height=500, margin=dict(t=30, b=40, l=50, r=20),
                        legend_title="Nodo (Barra)"
                    )
                    st.plotly_chart(fig_cmg, use_container_width=True)
                    
                    st.markdown("---")
                    st.markdown("#### Estadísticas Operativas por Nodo (excluyendo vacíos)")
                    
                    col_cmg1, col_cmg2, col_cmg3 = st.columns(3)
                    
                    sr_data = df_cmg[df_cmg['NODO'] == 'SANTA ROSA 220']['COSTO_MARGINAL_SOLES']
                    with col_cmg1:
                        st.markdown("**🔴 SANTA ROSA 220 (Centro)**")
                        st.metric("CMg Promedio", f"{sr_data.mean():.2f} Soles")
                        st.metric("CMg Máximo", f"{sr_data.max():.2f} Soles")
                        st.metric("CMg Mínimo", f"{sr_data.min():.2f} Soles")
                    
                    ta_data = df_cmg[df_cmg['NODO'] == 'TALARA 220']['COSTO_MARGINAL_SOLES']
                    with col_cmg2:
                        st.markdown("**🟤 TALARA 220 (Norte)**")
                        st.metric("CMg Promedio", f"{ta_data.mean():.2f} Soles")
                        st.metric("CMg Máximo", f"{ta_data.max():.2f} Soles")
                        st.metric("CMg Mínimo", f"{ta_data.min():.2f} Soles")
                        
                    mo_data = df_cmg[df_cmg['NODO'] == 'MOQUEGUA 220']['COSTO_MARGINAL_SOLES']
                    with col_cmg3:
                        st.markdown("**🔵 MOQUEGUA 220 (Sur)**")
                        st.metric("CMg Promedio", f"{mo_data.mean():.2f} Soles")
                        st.metric("CMg Máximo", f"{mo_data.max():.2f} Soles")
                        st.metric("CMg Mínimo", f"{mo_data.min():.2f} Soles")

                    st.markdown("---")
                    st.markdown("#### 📋 Trazabilidad y Auditoría de Datos")
                    
                    df_cmg_format = df_cmg.copy()
                    df_cmg_format['FECHA'] = df_cmg_format['FECHA_HORA'].dt.strftime('%d/%m/%Y')
                    df_cmg_format['HORA'] = df_cmg_format['FECHA_HORA'].dt.strftime('%H:%M')
                    
                    col_tab1, col_tab2 = st.columns(2)
                    
                    with col_tab1:
                        st.markdown("**✔️ Registros Válidos (Utilizados para gráficas y promedios)**")
                        df_validos = df_cmg_format.dropna(subset=['COSTO_MARGINAL_SOLES'])[['FECHA', 'HORA', 'NODO', 'COSTO_MARGINAL_SOLES']]
                        st.dataframe(df_validos, use_container_width=True, hide_index=True)
                        
                    with col_tab2:
                        st.markdown("**❌ Tramos con Falta de Datos (Omitidos en cálculo)**")
                        df_faltantes = df_cmg_format[df_cmg_format['COSTO_MARGINAL_SOLES'].isna()][['FECHA', 'HORA', 'NODO']]
                        if df_faltantes.empty:
                            st.success("✅ No se detectaron vacíos de información en este periodo operativo.")
                        else:
                            st.dataframe(df_faltantes, use_container_width=True, hide_index=True)

        else:
            st.warning("No se pudieron extraer datos operacionales de los anexos del COES para la ventana de tiempo estipulada.")