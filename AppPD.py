import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime
import urllib.parse
import requests
import zipfile
import io
import plotly.express as px

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
    
    # 2.1 EXTRACCIÓN DEL PROGRAMADO (CON ESCUDO ANTI "Hoja1")
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
                        if not hoja_prog:
                            hoja_prog = next((h for h in sheet_names_prog if 'PROGRAMADO' in h.upper()), None)
                        if not hoja_prog:
                            hoja_prog = next((h for h in sheet_names_prog if 'ANEXO' in h.upper()), None)
                        if not hoja_prog:
                            hoja_prog = sheet_names_prog[0] 
                            
                        df_prog = pd.read_excel(
                            archivo_excel_prog, 
                            sheet_name=hoja_prog, 
                            skiprows=8, 
                            usecols="B:N", 
                            names=columnas_estandar
                        )
                        df_prog = df_prog.dropna(subset=['Empresa', 'Equipo'], how='all')
                        df_prog = df_prog[~df_prog['Empresa'].astype(str).str.contains('TOTAL|NOTA|ELABORADO|FUENTE', case=False, na=False)]
                else:
                    st.sidebar.warning("Se descargó el ZIP programado pero no se encontró el archivo 'Anexo1_Intervenciones_(Osinergmin)'.")
    except Exception as e:
        st.sidebar.error(f"Error extrayendo Programado: {e}")

    # 2.2 EXTRACCIÓN DEL EJECUTADO (CON ESCUDO DE HOJA)
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
                if not hoja_buscada:
                    hoja_buscada = sheet_names_ejec[0]
                
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

# --- 3. FUNCIONES DE NORMALIZACIÓN ---
def normalizar_texto(serie):
    return serie.astype(str).str.strip().str.upper().str.normalize('NFKD').str.encode('ascii', errors='ignore').str.decode('utf-8')

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
    
    def estado_auditoria(row):
        if row['_merge'] == 'left_only': return 'Programado NO Ejecutado'
        elif row['_merge'] == 'right_only': return 'Ejecutado NO Programado'
        else: return 'Match: Programado Y Ejecutado'
            
    df_merged['Estado_Auditoria'] = df_merged.apply(estado_auditoria, axis=1)
    
    df_merged['Disponibilidad_Equipo'] = df_merged['Disponibilidad_Equipo_Ejec'].fillna(df_merged['Disponibilidad_Equipo_Prog']).fillna('NO ESPECIFICADO').astype(str).str.upper().str.strip()
    df_merged['Tipo_Equipo'] = df_merged['Tipo_Equipo_Ejec'].fillna(df_merged['Tipo_Equipo_Prog']).fillna('').astype(str).str.upper().str.strip()
    df_merged['Tipo_Mantenimiento'] = df_merged['Tipo_Mantenimiento_Ejec'].fillna(df_merged['Tipo_Mantenimiento_Prog']).fillna('NO ESPECIFICADO').astype(str).str.upper().str.strip()
    df_merged['Tipo_Mantenimiento'] = df_merged['Tipo_Mantenimiento'].replace('NAN', 'NO ESPECIFICADO')
    df_merged['Descripcion_Ejec'] = df_merged['Descripcion_Ejec'].fillna(df_merged['Descripcion_Prog']).fillna('-')
    
    df_merged['Central/Ubicacion'] = df_merged['Ubicacion_Ejec'].fillna(df_merged['Ubicacion_Prog']).fillna('-').astype(str).str.upper().str.strip()
    df_merged['Central/Ubicacion'] = normalizar_texto(df_merged['Central/Ubicacion'])

    def determinar_sector(row):
        tipo, equipo = str(row.get('Tipo_Equipo', '')), str(row.get('Equipo', ''))
        if tipo.startswith('G'): return 'GENERACIÓN'
        if tipo.startswith('T') or tipo.startswith('L'): return 'TRANSMISIÓN'
        if equipo.startswith('L-') or equipo.startswith('TR') or equipo.startswith('AT') or equipo.startswith('SE '): return 'TRANSMISIÓN'
        if equipo.startswith('G-') or equipo.startswith('TV') or equipo.startswith('TG') or equipo.startswith('CH '): return 'GENERACIÓN'
        return 'OTROS'
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
    with st.spinner("Descargando anexos y realizando validaciones de la base documental..."):
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
            
            df_filtrado = df_conciliado.copy()
            if len(empresa_sel) > 0: df_filtrado = df_filtrado[df_filtrado['Empresa'].isin(empresa_sel)]
            if len(sector_sel) > 0: df_filtrado = df_filtrado[df_filtrado['Sector'].isin(sector_sel)]
            if len(disp_sel) > 0: df_filtrado = df_filtrado[df_filtrado['Disponibilidad_Equipo'].isin(disp_sel)]
            if len(tipo_sel) > 0: df_filtrado = df_filtrado[df_filtrado['Tipo_Mantenimiento'].isin(tipo_sel)]

            st.markdown("---")

            # Removida la Pestaña 6
            tab1, tab2, tab3, tab4, tab5 = st.tabs([
                "📊 1. Resumen y Métricas",
                "✅ 2. MATCH: Tiempos y MW",
                "⚠️ 3. Ejecutados NO Programados", 
                "❌ 4. Programados NO Ejecutados",
                "🗄️ 5. Datos Originales (Raw)"
            ])
            
            # --- PESTAÑA 1: RESUMEN Y MÉTRICAS GERENCIALES ---
            with tab1:
                st.header("📊 Informe de la Supervisión de la Programación del Mantenimiento Diario")
                
                total_prog_raw = len(df_prog_raw)
                total_ejec_raw = len(df_ejec_raw)
                
                total_registros = len(df_filtrado)
                total_match = len(df_filtrado[df_filtrado['Estado_Auditoria'] == 'Match: Programado Y Ejecutado'])
                total_forzados = len(df_filtrado[df_filtrado['Estado_Auditoria'] == 'Ejecutado NO Programado'])
                total_ejecutados = total_match + total_forzados 
                total_no_ejec = len(df_filtrado[df_filtrado['Estado_Auditoria'] == 'Programado NO Ejecutado'])
                
                mw_perdidos_forzados = df_filtrado[df_filtrado['Estado_Auditoria'] == 'Ejecutado NO Programado']['MW_Indisponibles_Ejec'].sum()
                
                texto_diagnostico = f"""
                **📌 Resumen Automatizado de Fiscalización:** En la presente ventana, los documentos del COES reportaron un volumen crudo de **{total_prog_raw} Mantenimientos Programados** y **{total_ejec_raw} Mantenimientos Ejecutados**. Tras la conciliación de la base de datos, el sistema ha generado un universo unificado de **{total_registros} registros evaluados**.
                De este universo, se detectaron **{total_forzados} eventos forzados o de emergencia** que ingresaron sin programación, comprometiendo **{mw_perdidos_forzados:.2f} MW**. 
                """
                st.info(texto_diagnostico)
                
                st.markdown("#### 📑 1. Base Documental (Volumen Extraído del COES)")
                col_doc1, col_doc2, col_doc3 = st.columns(3)
                col_doc1.metric("Programados Originales (Raw)", total_prog_raw, help="Cantidad exacta de filas leídas en el Anexo Osinergmin.")
                col_doc2.metric("Ejecutados Originales (Raw)", total_ejec_raw, help="Cantidad exacta de filas leídas en el Anexo A.")
                col_doc3.metric("Universo Total Integrado", total_registros, delta="Sin duplicados", delta_color="normal", help="Total de intervenciones únicas (Match + Forzados + Cancelados).")
                
                st.markdown("#### 🔍 2. Resultados de la Auditoría")
                c1, c2, c3, c4 = st.columns(4)
                c1.metric("Match Exacto", total_match, help="Estaban programados y se ejecutaron.")
                c2.metric("Mantenimientos Forzados", total_forzados, delta=f"{mw_perdidos_forzados:.1f} MW", delta_color="inverse")
                c3.metric("Mantenimientos Cancelados", total_no_ejec)
                nota_descriptiva = "Representa el impacto temporal global en el sistema. Es la suma de las horas de retraso o adelanto de todos los mantenimientos ejecutados frente a su programación original."
                c4.metric(label="Desviación Neta SEIN (Horas)", value=f"{df_filtrado['Desviacion_Horas'].sum():.2f} h", delta="Impacto Global", delta_color="off", help=nota_descriptiva)
                
                st.markdown("---")
                
                st.markdown("#### ⏱️ Análisis de Tiempos de Ejecución (Programado vs. Ejecutado)")
                df_match_kpi = df_filtrado[df_filtrado['Estado_Auditoria'] == 'Match: Programado Y Ejecutado'].copy()
                
                if not df_match_kpi.empty:
                    df_match_kpi['Desempeño_Tiempo'] = np.where(df_match_kpi['Desviacion_Horas'] > 0, 'Mayor al Programado (Retraso)', 
                                                   np.where(df_match_kpi['Desviacion_Horas'] < 0, 'Menor al Programado (Adelanto)', 'Exactamente en Tiempo'))
                    cant_mayor = len(df_match_kpi[df_match_kpi['Desviacion_Horas'] > 0])
                    cant_menor = len(df_match_kpi[df_match_kpi['Desviacion_Horas'] < 0])
                    cant_exacto = len(df_match_kpi[df_match_kpi['Desviacion_Horas'] == 0])
                    horas_retraso = df_match_kpi[df_match_kpi['Desviacion_Horas'] > 0]['Desviacion_Horas'].sum()
                    horas_ahorro = abs(df_match_kpi[df_match_kpi['Desviacion_Horas'] < 0]['Desviacion_Horas'].sum())
                    
                    col_t1, col_t2, col_t3 = st.columns(3)
                    col_t1.metric("🔴 Excedieron su Programación", f"{cant_mayor} Mantenimientos", delta=f"+{horas_retraso:.2f} hrs de exceso", delta_color="inverse")
                    col_t2.metric("🟢 Terminaron antes de Tiempo", f"{cant_menor} Mantenimientos", delta=f"-{horas_ahorro:.2f} hrs de ahorro", delta_color="normal")
                    col_t3.metric("🔵 Exactamente a Tiempo", f"{cant_exacto} Mantenimientos", delta="0.00 hrs de desvío", delta_color="off")
                    
                    col_gt1, col_gt2 = st.columns(2)
                    with col_gt1:
                        tiempos_counts = df_match_kpi['Desempeño_Tiempo'].value_counts().reset_index()
                        tiempos_counts.columns = ['Estado', 'Cantidad']
                        fig_pie_tiempos = px.pie(tiempos_counts, values='Cantidad', names='Estado', hole=0.4, title="Proporción de Cumplimiento de Tiempos (Match)", color='Estado', color_discrete_map={'Mayor al Programado (Retraso)': '#d62728', 'Menor al Programado (Adelanto)': '#2ca02c', 'Exactamente en Tiempo': '#1f77b4'})
                        fig_pie_tiempos.update_traces(textinfo='value+percent', textfont_size=14, hoverinfo='label+percent+value')
                        st.plotly_chart(fig_pie_tiempos, use_container_width=True)
                        
                    with col_gt2:
                        df_sector_tiempos = df_match_kpi.groupby('Sector')[['Horas_Prog', 'Horas_Ejec']].sum().reset_index()
                        df_sector_tiempos_melted = df_sector_tiempos.melt(id_vars='Sector', value_vars=['Horas_Prog', 'Horas_Ejec'], var_name='Tipo', value_name='Horas')
                        fig_bar_tiempos = px.bar(df_sector_tiempos_melted, x='Sector', y='Horas', color='Tipo', barmode='group', title="Comparativa: Horas Programadas vs Ejecutadas por Sector", color_discrete_map={'Horas_Prog': '#1f77b4', 'Horas_Ejec': '#ff7f0e'}, text_auto='.1f')
                        st.plotly_chart(fig_bar_tiempos, use_container_width=True)
                else:
                    st.info("No hay mantenimientos ejecutados bajo programación (Match) para generar métricas de eficiencia.")
                
                st.markdown("---")
                st.markdown("#### 🛠️ Estadísticas Operativas por Tipo de Mantenimiento")
                col_tm1, col_tm2 = st.columns(2)
                with col_tm1:
                    tipo_counts = df_filtrado['Tipo_Mantenimiento'].value_counts().reset_index()
                    tipo_counts.columns = ['Tipo de Mantenimiento', 'Cantidad de Registros']
                    fig_pie_tipo = px.pie(tipo_counts, values='Cantidad de Registros', names='Tipo de Mantenimiento', hole=0.4, title="Registros Acumulados por Tipo", color_discrete_sequence=px.colors.qualitative.Prism)
                    fig_pie_tipo.update_traces(textinfo='value+percent', textfont_size=14, hoverinfo='label+percent+value')
                    st.plotly_chart(fig_pie_tipo, use_container_width=True)

                with col_tm2:
                    df_tipo_horas = df_filtrado.groupby('Tipo_Mantenimiento')[['Horas_Prog', 'Horas_Ejec']].sum().reset_index()
                    df_tipo_horas_melted = df_tipo_horas.melt(id_vars='Tipo_Mantenimiento', value_vars=['Horas_Prog', 'Horas_Ejec'], var_name='Categoría', value_name='Total de Horas')
                    fig_bar_tipo = px.bar(df_tipo_horas_melted, x='Tipo_Mantenimiento', y='Total de Horas', color='Categoría', barmode='group', title="Volumen de Horas (Prog vs Ejec) por Tipo", color_discrete_map={'Horas_Prog': '#1f77b4', 'Horas_Ejec': '#ff7f0e'}, text_auto='.1f')
                    st.plotly_chart(fig_bar_tipo, use_container_width=True)

                st.markdown("---")
                st.markdown("#### Análisis Gráfico de Auditoría Global")
                col_graf1, col_graf2 = st.columns(2)
                with col_graf1:
                    estado_counts = df_filtrado['Estado_Auditoria'].value_counts().reset_index()
                    estado_counts.columns = ['Estado', 'Cantidad']
                    fig_pie = px.pie(estado_counts, values='Cantidad', names='Estado', hole=0.4, title="Distribución de Todos los Eventos Operativos", color='Estado', color_discrete_map={'Match: Programado Y Ejecutado': '#2ca02c', 'Ejecutado NO Programado': '#d62728', 'Programado NO Ejecutado': '#ff7f0e'})
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
                st.subheader("✅ Match: Desviaciones de Potencia y Tiempos de Ejecución")
                df_tab2 = df_filtrado[df_filtrado['Estado_Auditoria'] == 'Match: Programado Y Ejecutado'].copy()
                if not df_tab2.empty:
                    df_tab2['Estado_Tiempo'] = np.where(df_tab2['Desviacion_Horas'] > 0, '🔴 Mayor Tiempo', 
                                               np.where(df_tab2['Desviacion_Horas'] < 0, '🟠 Menor Tiempo', '🟢 Tiempo Exacto'))
                    columnas_tab2 = ['Estado_Tiempo', 'Empresa', 'Central/Ubicacion', 'Equipo', 'Sector', 'Tipo_Mantenimiento', 'Disponibilidad_Equipo', 
                                     'Inicio_Prog', 'Fin_Prog', 'Horas_Prog', 
                                     'Inicio_Ejec', 'Fin_Ejec', 'Horas_Ejec', 'Desviacion_Horas',
                                     'MW_Indisponibles_Prog', 'MW_Indisponibles_Ejec', 'Desviacion_MW', 'Descripcion_Ejec']
                    st.dataframe(df_tab2[columnas_tab2], use_container_width=True)
                else:
                    st.info("No hay coincidencias perfectas bajo los filtros actuales.")

            # --- PESTAÑA 3: FORZADOS ---
            with tab3:
                st.subheader("⚠️ Mantenimientos Ejecutados NO Programados (Forzados/Emergencias)")
                df_tab3 = df_filtrado[df_filtrado['Estado_Auditoria'] == 'Ejecutado NO Programado']
                if not df_tab3.empty:
                    columnas_tab3 = ['Empresa', 'Central/Ubicacion', 'Equipo', 'Sector', 'Tipo_Mantenimiento', 'Disponibilidad_Equipo', 'Inicio_Ejec', 'Fin_Ejec', 'Horas_Ejec', 'MW_Indisponibles_Ejec', 'Descripcion_Ejec']
                    st.dataframe(df_tab3[columnas_tab3], use_container_width=True)
                else:
                    st.info("No hay eventos forzados bajo los filtros actuales.")

            # --- PESTAÑA 4: NO EJECUTADOS ---
            with tab4:
                st.subheader("❌ Mantenimientos Programados NO Ejecutados")
                df_tab4 = df_filtrado[df_filtrado['Estado_Auditoria'] == 'Programado NO Ejecutado']
                if not df_tab4.empty:
                    columnas_tab4 = ['Empresa', 'Central/Ubicacion', 'Equipo', 'Sector', 'Tipo_Mantenimiento', 'Disponibilidad_Equipo', 'Inicio_Prog', 'Fin_Prog', 'Horas_Prog', 'MW_Indisponibles_Prog', 'Descripcion_Prog']
                    st.dataframe(df_tab4[columnas_tab4], use_container_width=True)
                else:
                    st.info("No hay eventos no ejecutados bajo los filtros actuales.")

            # --- PESTAÑA 5: RAW DATA ---
            with tab5:
                st.subheader("🗄️ Trazabilidad: Archivos Crudos (Raw Data)")
                col_raw1, col_raw2 = st.columns(2)
                with col_raw1:
                    st.markdown(f"**📁 Anexo Osinergmin (Programado Diario) - {len(df_prog_raw)} Registros**")
                    st.dataframe(df_prog_raw, use_container_width=True)
                with col_raw2:
                    st.markdown(f"**📁 Anexo A (Mantenimientos Ejecutados) - {len(df_ejec_raw)} Registros**")
                    st.dataframe(df_ejec_raw, use_container_width=True)

        else:
            st.warning("No se pudieron obtener ambos archivos. Verifique que la fecha seleccionada ya tenga los reportes publicados.")