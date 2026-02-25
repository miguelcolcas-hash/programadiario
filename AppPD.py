import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime
import urllib.parse
import requests
import zipfile
import io
import plotly.express as px
import difflib
import re

# --- 1. CONFIGURACIÓN DE PÁGINA ---
st.set_page_config(page_title="Supervisión SEIN - Osinergmin", layout="wide")
st.title("⚡ Dashboard de Supervisión de Mantenimientos - SEIN")
st.markdown("Supervisión COES: Mantenimientos Programado vs. Ejecutado")

MESES = {
    1: "ENERO", 2: "FEBRERO", 3: "MARZO", 4: "ABRIL",
    5: "MAYO", 6: "JUNIO", 7: "JULIO", 8: "AGOSTO",
    9: "SETIEMBRE", 10: "OCTUBRE", 11: "NOVIEMBRE", 12: "DICIEMBRE"
}

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
    
    return url_prog, urls_ejec

# --- 2. EXTRACCIÓN Y LIMPIEZA ESPEJO (ETL) ---
@st.cache_data(show_spinner=False)
def extraer_datos_coes(fecha):
    url_prog, urls_ejec = generar_urls_coes(fecha)
    headers = {'User-Agent': 'Mozilla/5.0'}
    
    df_prog, df_ejec = None, None
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
                else:
                    st.sidebar.warning("Se descargó el ZIP programado pero no se encontró el archivo 'Anexo1_Intervenciones_(Osinergmin)'.")
    except Exception as e:
        st.sidebar.error(f"Error extrayendo Programado: {e}")

    # 2.2 EXTRACCIÓN DEL EJECUTADO
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
        except Exception:
            continue
            
    if not exito_ejecutado:
        st.sidebar.error("⚠️ No se pudo descargar el archivo de Mantenimientos Ejecutados. Verifique la fecha.")

    return df_prog, df_ejec

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
    for df in [df_prog, df_ejec]:
        if df is not None:
            df['Empresa'] = normalizar_texto(df['Empresa'])
            df['Equipo'] = normalizar_texto(df['Equipo'])
            df.replace(['NAN', 'NAT', ''], np.nan, inplace=True)
            df['Seq_Mantenimiento'] = df.groupby(['Empresa', 'Equipo']).cumcount()
            
    df_merged = pd.merge(df_prog, df_ejec, on=['Empresa', 'Equipo', 'Seq_Mantenimiento'], how='outer', suffixes=('_Prog', '_Ejec'), indicator=True)
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
st.sidebar.header("Informe de Programación Diaria")
fecha_seleccionada = st.sidebar.date_input("Seleccione Fecha de Operación", datetime.today())

if 'dashboard_activo' not in st.session_state:
    st.session_state.dashboard_activo = False

if st.sidebar.button("Procesar Información"):
    st.session_state.dashboard_activo = True

if st.session_state.dashboard_activo:
    with st.spinner("Construyendo Tablero Gerencial y sincronizando métricas..."):
        df_prog_raw, df_ejec_raw = extraer_datos_coes(fecha_seleccionada)
        
        if df_prog_raw is not None and df_ejec_raw is not None:
            df_conciliado = conciliar_datos(df_prog_raw.copy(), df_ejec_raw.copy())
            
            st.markdown("### 🎛️ Filtros Dinámicos")
            col_f1, col_f2, col_f3, col_f4 = st.columns(4)
            
            empresas_disp = sorted(df_conciliado[df_conciliado['Empresa'] != 'NAN']['Empresa'].dropna().unique())
            empresa_sel = col_f1.multiselect("Empresa Concesionaria:", empresas_disp, default=[])
            
            sectores_disp = sorted(df_conciliado['Sector'].unique())
            sector_sel = col_f2.multiselect("Sector (Gen/Trans):", sectores_disp, default=sectores_disp)
            
            disp_equipos = sorted(df_conciliado[df_conciliado['Disponibilidad_Equipo'] != 'NAN']['Disponibilidad_Equipo'].unique())
            disp_sel = col_f3.multiselect("Estado (E/S o F/S):", disp_equipos, default=disp_equipos)
            
            tipos_disp = sorted(df_conciliado[df_conciliado['Tipo_Mantenimiento'] != 'NO ESPECIFICADO']['Tipo_Mantenimiento'].unique())
            tipo_sel = col_f4.multiselect("Tipo de Mantenimiento:", tipos_disp, default=[])
            
            # 🔴 APLICACIÓN DE FILTROS AL UNIVERSO CONCILIADO
            df_filtrado = df_conciliado.copy()
            if len(empresa_sel) > 0: df_filtrado = df_filtrado[df_filtrado['Empresa'].isin(empresa_sel)]
            if len(sector_sel) > 0: df_filtrado = df_filtrado[df_filtrado['Sector'].isin(sector_sel)]
            if len(disp_sel) > 0: df_filtrado = df_filtrado[df_filtrado['Disponibilidad_Equipo'].isin(disp_sel)]
            if len(tipo_sel) > 0: df_filtrado = df_filtrado[df_filtrado['Tipo_Mantenimiento'].isin(tipo_sel)]

            # 🔴 APLICACIÓN DE FILTROS A LOS ARCHIVOS RAW (BASE DOCUMENTAL EXACTA)
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
                df_prog_raw_f.index = np.arange(1, len(df_prog_raw_f) + 1) # Índice Inicia en 1

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
                df_ejec_raw_f.index = np.arange(1, len(df_ejec_raw_f) + 1) # Índice Inicia en 1

            st.markdown("---")

            tab1, tab2, tab3, tab4, tab5 = st.tabs([
                "📊 1. Resumen y Métricas",
                "✅ 2. MATCH: Tiempos y MW",
                "⚠️ 3. Ejecutados NO Programados", 
                "❌ 4. Programados NO Ejecutados",
                "🗄️ 5. Datos Originales (Raw)"
            ])
            
            # --- PESTAÑA 1: RESUMEN EJECUTIVO ---
            with tab1:
                st.header("📊 Informe de la Supervisión de la Programación del Mantenimiento Diario")
                
                # 🔴 Conteos Matemáticos Precisos con las bases filtradas
                total_prog_raw_count = len(df_prog_raw_f) if df_prog_raw_f is not None else 0
                total_ejec_raw_count = len(df_ejec_raw_f) if df_ejec_raw_f is not None else 0
                
                total_match = len(df_filtrado[df_filtrado['Estado_Supervision'] == 'Programado y Ejecutado'])
                total_no_ejec = len(df_filtrado[df_filtrado['Estado_Supervision'] == 'Programado NO Ejecutado'])
                total_forzados = len(df_filtrado[df_filtrado['Estado_Supervision'] == 'Ejecutado NO Programado'])
                
                total_programados = total_match + total_no_ejec
                total_ejecutados = total_match + total_forzados
                total_universo = len(df_filtrado)
                
                mw_perdidos_forzados = df_filtrado[df_filtrado['Estado_Supervision'] == 'Ejecutado NO Programado']['MW_Indisponibles_Ejec'].sum()
                desviacion_neta = df_filtrado['Desviacion_Horas'].sum()
                tasa_cumplimiento = (total_match / total_programados * 100) if total_programados > 0 else 0
                
                texto_diagnostico = f"""
                **📌 Resumen Ejecutivo de Operaciones:** En la ventana de supervisión, se procesaron y consolidaron **{total_prog_raw_count} mantenimientos programados** y **{total_ejec_raw_count} mantenimientos ejecutados** (bajo los filtros aplicados).
                Se determinó una **Tasa de Cumplimiento Normativo del {tasa_cumplimiento:.1f}%** de los eventos planificados. 
                Sin embargo, la confiabilidad del sistema fue impactada por la presencia de **{total_forzados} mantenimientos ejecutados no programados** (salidas forzadas), los cuales comprometieron **{mw_perdidos_forzados:.2f} MW** de la capacidad del SEIN.
                """
                st.info(texto_diagnostico)
                
                # ---------------------------------------------------------
                # SECCIÓN 1: INTEGRIDAD DE LA BASE DOCUMENTAL
                # ---------------------------------------------------------
                st.markdown("#### 📑 1. Integridad de la Base Documental")
                st.caption("Volúmenes de registros extraídos directamente de los archivos del COES, ajustados a los filtros actuales.")
                
                c_doc1, c_doc2, c_doc3 = st.columns(3)
                c_doc1.metric("Mantenimientos Programados", total_prog_raw_count, help="Volumen del Anexo Osinergmin (Programados y Cancelados).")
                c_doc2.metric("Mantenimientos Ejecutados", total_ejec_raw_count, help="Volumen del Anexo A (Programados/Ejecutados y Forzados).")
                c_doc3.metric("Universo Total Único", total_universo, delta="Cruce Consolidado", delta_color="normal", help="Base de datos unificada sin duplicados.")
                
                col_g1, col_g2 = st.columns(2)
                with col_g1:
                    df_doc = pd.DataFrame({'Origen Operativo': ['Total Programados', 'Total Ejecutados'], 'Volumen': [total_prog_raw_count, total_ejec_raw_count]})
                    fig_bar_doc = px.bar(df_doc, x='Origen Operativo', y='Volumen', color='Origen Operativo', text_auto=True, title="Comparativa de Volúmenes Documentales", color_discrete_sequence=['#1f77b4', '#ff7f0e'])
                    fig_bar_doc.update_layout(showlegend=False)
                    st.plotly_chart(fig_bar_doc, use_container_width=True)
                    
                with col_g2:
                    df_estado_counts = df_filtrado['Estado_Supervision'].value_counts().reset_index()
                    df_estado_counts.columns = ['Estado Operativo', 'Cantidad']
                    fig_pie_univ = px.pie(df_estado_counts, values='Cantidad', names='Estado Operativo', hole=0.4, title="Estructura del Universo de Mantenimientos", color='Estado Operativo', color_discrete_map={'Programado y Ejecutado': '#2ca02c', 'Ejecutado NO Programado': '#d62728', 'Programado NO Ejecutado': '#ff7f0e'})
                    fig_pie_univ.update_traces(textinfo='value+percent', textfont_size=14, hoverinfo='label+percent+value')
                    st.plotly_chart(fig_pie_univ, use_container_width=True)

                st.markdown("---")

                # ---------------------------------------------------------
                # SECCIÓN 2: RESULTADOS DE LA SUPERVISIÓN
                # ---------------------------------------------------------
                st.markdown("#### 🔍 2. Resultados de la Supervisión Operativa")
                st.caption("Evaluación del impacto en los tiempos de ejecución y capacidad (MW).")
                
                c_res1, c_res2, c_res3, c_res4 = st.columns(4)
                c_res1.metric("Programado y Ejecutado", total_match, help="Cumplieron con planificar y ejecutar.")
                c_res2.metric("Ejecutado NO Programado", total_forzados, delta=f"-{mw_perdidos_forzados:.1f} MW", delta_color="inverse")
                c_res3.metric("Programado NO Ejecutado", total_no_ejec, help="Se programaron pero no se realizaron.")
                
                if desviacion_neta > 0:
                    c_res4.metric("Desviación Neta (Horas)", f"+{desviacion_neta:.2f} h", delta="Exceso / Retraso", delta_color="inverse")
                elif desviacion_neta < 0:
                    c_res4.metric("Desviación Neta (Horas)", f"{desviacion_neta:.2f} h", delta="Ahorro Operativo", delta_color="normal")
                else:
                    c_res4.metric("Desviación Neta (Horas)", "0.00 h", delta="Ejecución Perfecta", delta_color="off")
                
                st.markdown("<br>", unsafe_allow_html=True)
                if desviacion_neta > 0 or total_forzados > 0:
                    st.warning(f"**⚠️ Alerta Directiva de Supervisión:** El SEIN registró **{total_forzados}** mantenimientos no programados y un retraso acumulado de **{desviacion_neta:.2f} horas**. Por favor, analice las gráficas inferiores de impacto.")
                else:
                    st.success("**✅ Operación Óptima:** Las empresas concesionarias han operado el SEIN respetando los márgenes de tiempo estipulados y no se detectaron salidas intempestivas.")
                
                col_g3, col_g4 = st.columns(2)
                with col_g3:
                    df_mw_melt = df_filtrado.groupby('Estado_Supervision')[['MW_Indisponibles_Prog', 'MW_Indisponibles_Ejec']].sum().reset_index()
                    df_mw_melt = df_mw_melt.melt(id_vars='Estado_Supervision', var_name='Fase', value_name='Megavatios (MW)')
                    df_mw_melt['Fase'] = df_mw_melt['Fase'].replace({'MW_Indisponibles_Prog': 'MW Programado', 'MW_Indisponibles_Ejec': 'MW Ejecutado'})
                    
                    fig_bar_mw = px.bar(df_mw_melt, x='Estado_Supervision', y='Megavatios (MW)', color='Fase', barmode='group', text_auto='.1f', title="Impacto de Indisponibilidad (MW) por Evento", color_discrete_map={'MW Programado': '#1f77b4', 'MW Ejecutado': '#ff7f0e'})
                    st.plotly_chart(fig_bar_mw, use_container_width=True)

                with col_g4:
                    df_excesos = df_filtrado[df_filtrado['Desviacion_Horas'] > 0]
                    if not df_excesos.empty:
                        df_agrupado = df_excesos.groupby('Empresa')['Desviacion_Horas'].sum().reset_index()
                        df_agrupado = df_agrupado.sort_values('Desviacion_Horas', ascending=False).head(10)
                        fig_bar_tiempo = px.bar(df_agrupado, x='Desviacion_Horas', y='Empresa', orientation='h', title="Top 10 Empresas con Mayor Retraso en Ejecución (Horas)", text_auto='.2f', color='Desviacion_Horas', color_continuous_scale='Reds')
                        fig_bar_tiempo.update_layout(yaxis={'categoryorder':'total ascending'})
                        st.plotly_chart(fig_bar_tiempo, use_container_width=True)
                    else:
                        st.success("✅ Excelente: No se registraron empresas con excesos de tiempo de mantenimiento.")

                st.markdown("---")
                
                st.markdown("#### ⏱️ Análisis de Tiempos Exclusivo (Programado y Ejecutado)")
                df_match_kpi = df_filtrado[df_filtrado['Estado_Supervision'] == 'Programado y Ejecutado'].copy()
                
                if not df_match_kpi.empty:
                    df_match_kpi['Desempeño_Tiempo'] = np.where(df_match_kpi['Desviacion_Horas'] > 0, 'Excedieron Programación (Retraso)', 
                                                   np.where(df_match_kpi['Desviacion_Horas'] < 0, 'Terminaron Antes (Ahorro)', 'Ejecución Exacta'))
                    
                    col_gt1, col_gt2 = st.columns(2)
                    with col_gt1:
                        tiempos_counts = df_match_kpi['Desempeño_Tiempo'].value_counts().reset_index()
                        tiempos_counts.columns = ['Eficiencia', 'Cantidad']
                        fig_pie_tiempos = px.pie(tiempos_counts, values='Cantidad', names='Eficiencia', hole=0.4, title="Proporción de Eficiencia en Tiempos de Maniobra", color='Eficiencia', color_discrete_map={'Excedieron Programación (Retraso)': '#d62728', 'Terminaron Antes (Ahorro)': '#2ca02c', 'Ejecución Exacta': '#1f77b4'})
                        fig_pie_tiempos.update_traces(textinfo='value+percent', textfont_size=14, hoverinfo='label+percent+value')
                        st.plotly_chart(fig_pie_tiempos, use_container_width=True)
                        
                    with col_gt2:
                        df_sector_tiempos = df_match_kpi.groupby('Sector')[['Horas_Prog', 'Horas_Ejec']].sum().reset_index()
                        df_sector_tiempos_melted = df_sector_tiempos.melt(id_vars='Sector', value_vars=['Horas_Prog', 'Horas_Ejec'], var_name='Tipo', value_name='Horas')
                        df_sector_tiempos_melted['Tipo'] = df_sector_tiempos_melted['Tipo'].replace({'Horas_Prog': 'Horas Programadas', 'Horas_Ejec': 'Horas Ejecutadas'})
                        fig_bar_tiempos = px.bar(df_sector_tiempos_melted, x='Sector', y='Horas', color='Tipo', barmode='group', title="Volumen de Horas Operativas por Sector", color_discrete_map={'Horas Programadas': '#1f77b4', 'Horas Ejecutadas': '#ff7f0e'}, text_auto='.1f')
                        st.plotly_chart(fig_bar_tiempos, use_container_width=True)
                else:
                    st.info("No hay mantenimientos ejecutados bajo programación (Match) para generar métricas de eficiencia.")
                
                st.markdown("---")
                st.markdown("#### 🛠️ Estadísticas Operativas por Tipo de Mantenimiento")
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
                    st.plotly_chart(fig_bar_tipo, use_container_width=True)

                st.markdown("---")
                st.markdown("#### Análisis Gráfico de Supervisión Global")
                col_graf1, col_graf2 = st.columns(2)
                with col_graf1:
                    estado_counts = df_filtrado['Estado_Supervision'].value_counts().reset_index()
                    estado_counts.columns = ['Estado', 'Cantidad']
                    fig_pie = px.pie(estado_counts, values='Cantidad', names='Estado', hole=0.4, title="Distribución de Todos los Eventos Operativos", color='Estado', color_discrete_map={'Programado y Ejecutado': '#2ca02c', 'Ejecutado NO Programado': '#d62728', 'Programado NO Ejecutado': '#ff7f0e'})
                    fig_pie.update_traces(textinfo='value+percent', textfont_size=14, hoverinfo='label+percent+value')
                    st.plotly_chart(fig_pie, use_container_width=True)
                    
                with col_graf2:
                    df_excesos = df_filtrado[df_filtrado['Desviacion_Horas'] > 0]
                    if not df_excesos.empty:
                        df_agrupado = df_excesos.groupby('Empresa')['Desviacion_Horas'].sum().reset_index()
                        df_agrupado = df_agrupado.sort_values('Desviacion_Horas', ascending=False).head(10)
                        fig_bar_tiempo = px.bar(df_agrupado, x='Desviacion_Horas', y='Empresa', orientation='h', title="Top 10: Empresas con Mayores Retrasos (Horas Acumuladas)", text_auto='.2f', color='Desviacion_Horas', color_continuous_scale='Reds')
                        fig_bar_tiempo.update_layout(yaxis={'categoryorder':'total ascending'})
                        st.plotly_chart(fig_bar_tiempo, use_container_width=True)
                    else:
                        st.success("✅ Excelente: No se registraron empresas con retrasos operacionales.")

            # --- PESTAÑA 2: MATCH ---
            with tab2:
                st.subheader("✅ Mantenimientos Programados y Ejecutados")
                st.caption("Esta vista muestra el cruce exacto de los eventos del COES.")
                df_tab2 = df_filtrado[df_filtrado['Estado_Supervision'] == 'Programado y Ejecutado'].copy()
                if not df_tab2.empty:
                    df_tab2['Estado_Tiempo'] = np.where(df_tab2['Desviacion_Horas'] > 0, '🔴 Mayor Tiempo', 
                                               np.where(df_tab2['Desviacion_Horas'] < 0, '🟠 Menor Tiempo', '🟢 Tiempo Exacto'))
                    columnas_tab2 = ['Estado_Tiempo', 'Empresa', 'Central/Ubicacion', 'Equipo', 'Sector', 'Tipo_Mantenimiento', 'Disponibilidad_Equipo', 
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
                    columnas_tab3 = ['Empresa', 'Central/Ubicacion', 'Equipo', 'Sector', 'Tipo_Mantenimiento', 'Disponibilidad_Equipo', 'Inicio_Ejec', 'Fin_Ejec', 'Horas_Ejec', 'MW_Indisponibles_Ejec', 'Descripcion_Ejec']
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
                    columnas_tab4 = ['Empresa', 'Central/Ubicacion', 'Equipo', 'Sector', 'Tipo_Mantenimiento', 'Disponibilidad_Equipo', 'Inicio_Prog', 'Fin_Prog', 'Horas_Prog', 'MW_Indisponibles_Prog', 'Descripcion_Prog']
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

        else:
            st.warning("No se pudieron obtener ambos archivos. Verifique que la fecha seleccionada ya tenga los reportes publicados.")