import streamlit as st
import pandas as pd
import requests
import io
import re
import unicodedata
from datetime import datetime

# ---------------------------------------------------------
# 1. CONFIGURACIÓN DE PÁGINA Y ESTILOS CORPORATIVOS
# ---------------------------------------------------------
st.set_page_config(
    page_title="Control de Calidad - Uva 2026",
    page_icon="🍇",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.markdown("""
    <style>
    :root {
        --primary-blue: #002060;
        --secondary-blue: #1A365D;
        --accent-orange: #D97706;
        --accent-green: #059669;
        --accent-red: #DC2626;
    }
    
    .block-container {
        padding-top: 1.2rem;
        padding-bottom: 2rem;
    }
    
    .status-banner {
        padding: 0.6rem 1rem;
        border-radius: 8px;
        font-weight: 600;
        font-size: 0.88rem;
        margin-bottom: 1.2rem;
    }
    .status-online {
        background-color: #D1FAE5;
        color: #065F46;
        border: 1px solid #A7F3D0;
    }
    .status-local {
        background-color: #FEF3C7;
        color: #92400E;
        border: 1px solid #FDE68A;
    }
    
    .kpi-container {
        display: grid;
        grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
        gap: 1rem;
        margin-bottom: 1.5rem;
    }
    .kpi-card {
        background: #FFFFFF;
        border-radius: 10px;
        padding: 1.1rem;
        border-left: 5px solid var(--primary-blue);
        box-shadow: 0 2px 6px rgba(0,0,0,0.06);
    }
    .kpi-card.ok { border-left-color: var(--accent-green); }
    .kpi-card.no-conf { border-left-color: var(--accent-red); }
    .kpi-card.obs { border-left-color: var(--accent-orange); }
    
    .kpi-title {
        font-size: 0.78rem;
        text-transform: uppercase;
        color: #64748B;
        font-weight: 700;
        letter-spacing: 0.5px;
    }
    .kpi-value {
        font-size: 1.8rem;
        font-weight: 800;
        color: #0F172A;
        margin: 0.2rem 0;
    }
    .kpi-sub {
        font-size: 0.8rem;
        color: #64748B;
    }

    @media (max-width: 768px) {
        .kpi-value { font-size: 1.4rem; }
        .block-container { padding-left: 0.5rem; padding-right: 0.5rem; }
    }
    </style>
""", unsafe_allow_html=True)


# ---------------------------------------------------------
# 2. TRANSFORMACIÓN DE ENLACE Y LECTURA EN TIEMPO REAL
# ---------------------------------------------------------
ENLACE_SHAREPOINT_DEFAULT = "https://betacompe-my.sharepoint.com/:x:/g/personal/apizarro_beta_com_pe/IQDiOQlfZUjtQpijpRx5SMP_AQlGgiWqFG4fPK2eYW-e2OM?e=Haqj3X"

def normalizar_url_sharepoint(url: str) -> str:
    url = url.strip()
    if "sharepoint.com" in url or "onedrive.live.com" in url or "1drv.ms" in url:
        url = re.sub(r'[\?&]web=1', '', url)
        if "download=1" not in url:
            url += ("&download=1" if "?" in url else "?download=1")
    return url

@st.cache_data(ttl=15)
def cargar_datos_online(url: str):
    url_directa = normalizar_url_sharepoint(url)
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/115.0.0.0 Safari/537.36"
    }
    
    resp = requests.get(url_directa, headers=headers, timeout=20, allow_redirects=True)
    resp.raise_for_status()
    
    content_type = resp.headers.get("Content-Type", "")
    if "text/html" in content_type or resp.content.startswith(b"<!DOCTYPE") or resp.content.startswith(b"<html"):
        raise PermissionError("SharePoint devolvió una página de login. Verifica los permisos generales del enlace.")
        
    return pd.read_excel(io.BytesIO(resp.content))

def normalizar_dataframe_robusto(df_in: pd.DataFrame) -> tuple[pd.DataFrame, dict]:
    """ Normalización inteligente de columnas y valores para evitar pérdidas de datos. """
    df = df_in.copy()
    df.columns = [str(c).strip() for c in df.columns]

    def norm_txt(t):
        if not isinstance(t, str):
            return ""
        t_norm = unicodedata.normalize('NFD', t)
        return ''.join(c for c in t_norm if unicodedata.category(c) != 'Mn').strip().upper()

    col_map = {}
    mapped_targets = set()
    
    # Mapeo por palabras clave de encabezado
    for c in df.columns:
        c_clean = norm_txt(c)
        if 'RESOLUCION' not in mapped_targets and any(k in c_clean for k in ['RESOLUC', 'ESTADO', 'ESTATUS', 'DICTAMEN', 'CALIFIC', 'RESULTAD', 'EVALUAC', 'CONFORMID', 'CONDICION']):
            col_map[c] = 'RESOLUCION'
            mapped_targets.add('RESOLUCION')
        elif 'DIA DE PROCESO' not in mapped_targets and ('PROCESO' in c_clean or ('DIA' in c_clean and 'PROC' in c_clean)):
            col_map[c] = 'DIA DE PROCESO'
            mapped_targets.add('DIA DE PROCESO')
        elif 'CLIENTE' not in mapped_targets and any(k in c_clean for k in ['CLIENTE', 'DESTINO', 'RECIBIDOR', 'IMPORTADOR']):
            col_map[c] = 'CLIENTE'
            mapped_targets.add('CLIENTE')
        elif 'VARIEDAD' not in mapped_targets and any(k in c_clean for k in ['VARIEDAD', 'CULTIVAR']):
            col_map[c] = 'VARIEDAD'
            mapped_targets.add('VARIEDAD')
        elif 'LOTE' not in mapped_targets and 'LOTE' in c_clean:
            col_map[c] = 'LOTE'
            mapped_targets.add('LOTE')
        elif 'N° PALLET' not in mapped_targets and ('PALLET' in c_clean or 'PALET' in c_clean):
            col_map[c] = 'N° PALLET'
            mapped_targets.add('N° PALLET')
        elif 'ORDEN BETA' not in mapped_targets and ('ORDEN' in c_clean or 'BETA' in c_clean):
            col_map[c] = 'ORDEN BETA'
            mapped_targets.add('ORDEN BETA')
        elif 'OBSERVACION' not in mapped_targets and any(k in c_clean for k in ['OBSERVAC', 'MOTIVO', 'COMENTARIO', 'DETALLE']):
            col_map[c] = 'OBSERVACION'
            mapped_targets.add('OBSERVACION')

    df = df.rename(columns=col_map)

    # Función para extraer series únicas incluso con duplicados
    def obtener_serie(col_name):
        if col_name not in df.columns:
            return pd.Series([None] * len(df), index=df.index)
        data = df[col_name]
        if isinstance(data, pd.DataFrame):
            return data.iloc[:, 0]
        return data

    # Clasificador flexible de estado
    def clasificar_res(val):
        if pd.isna(val):
            return 'SIN ESTATUS'
        s = str(val).replace('\xa0', ' ').strip().upper()
        if not s or s in ['NAN', 'NONE', 'NULL', '']:
            return 'SIN ESTATUS'
        if any(k in s for k in ['NO', 'OBS', 'RECH', 'DESAP', 'FAIL', 'DEFECT', 'INCONF']):
            return 'NO CONFORME'
        if any(k in s for k in ['CONF', 'OK', 'APROB', 'ACEPT', 'PASS', 'BUENO']):
            return 'CONFORME'
        return s

    df['RESOLUCION'] = obtener_serie('RESOLUCION').apply(clasificar_res)
    df['DIA DE PROCESO'] = obtener_serie('DIA DE PROCESO').astype(str).str.replace('.0', '', regex=False).str.strip()
    df['CLIENTE'] = obtener_serie('CLIENTE').fillna('SIN CLIENTE').astype(str).str.strip()
    df['VARIEDAD'] = obtener_serie('VARIEDAD').fillna('DESCONOCIDO').astype(str).str.strip()
    df['LOTE'] = obtener_serie('LOTE').astype(str).str.replace('.0', '', regex=False).str.strip()
    df['OBSERVACION_CLEAN'] = obtener_serie('OBSERVACION').fillna('').astype(str).str.strip()

    return df, col_map


# ---------------------------------------------------------
# 3. BARRA LATERAL
# ---------------------------------------------------------
st.sidebar.title("🍇 Control de Calidad")
st.sidebar.markdown("**Pallets Observados - Uva 2026**")
st.sidebar.markdown("---")

url_input = st.sidebar.text_input(
    "🔗 Enlace SharePoint Web en línea:",
    value=ENLACE_SHAREPOINT_DEFAULT,
    help="El sistema ajusta automáticamente el link para descarga directa en tiempo real."
)

st.sidebar.markdown("---")
st.sidebar.markdown("**🚨 Respaldo Manual**")
archivo_manual = st.sidebar.file_uploader(
    "Subir Excel manualmente:",
    type=["xlsx", "xls"],
    help="Utilizar si la red de SharePoint se interrumpe."
)

if st.sidebar.button("🔄 Refrescar Datos Nube"):
    st.cache_data.clear()

# ---------------------------------------------------------
# 4. CARGA DE DATOS
# ---------------------------------------------------------
df_raw = None
origen_mensaje = ""
tipo_banner = ""

if archivo_manual is not None:
    try:
        df_raw = pd.read_excel(archivo_manual)
        origen_mensaje = f"📁 **MODO MANUAL:** Mostrando datos del archivo subido ({archivo_manual.name})"
        tipo_banner = "status-local"
    except Exception as e:
        st.error(f"Error al leer el archivo Excel: {e}")
else:
    target_url = url_input.strip()
    if target_url:
        try:
            with st.spinner("Descargando desde SharePoint en tiempo real..."):
                df_raw = cargar_datos_online(target_url)
                hora_actual = datetime.now().strftime("%H:%M:%S")
                origen_mensaje = f"🟢 **TIEMPO REAL (SharePoint Web):** Conectado con éxito a las {hora_actual}"
                tipo_banner = "status-online"
        except PermissionError as pe:
            st.sidebar.error(f"⚠️ {pe}")
        except Exception as e:
            st.sidebar.warning(f"⚠️ No se pudo descargar desde el enlace ({e}).")

# ---------------------------------------------------------
# 5. DASHBOARD EJECUTIVO
# ---------------------------------------------------------
st.title("📊 Resumen Ejecutivo de Calidad - Planta Norte")

if df_raw is None:
    st.warning("💡 **Esperando conexión:** Carga el archivo en línea o sube un Excel mediante la barra lateral.")
else:
    st.markdown(f'<div class="status-banner {tipo_banner}">{origen_mensaje}</div>', unsafe_allow_html=True)
    
    # Procesamiento flexible
    df, mapa_cols = normalizar_dataframe_robusto(df_raw)

    # SECCIÓN INTERACTIVA DE DIAGNÓSTICO (Útil para auditar columnas)
    with st.expander("🔍 Diagnóstico de Columnas y Datos Detectados"):
        st.write("**Columnas originales leídas del Excel:**", df_raw.columns.tolist())
        st.write("**Mapeo automático de columnas aplicados:**", mapa_cols)
        st.write("**Valores únicos detectados en Resolución:**", df['RESOLUCION'].value_counts().to_dict())

    # FILTRO POR DÍA DE PROCESO
    dias_disponibles = sorted([d for d in df['DIA DE PROCESO'].unique() if d not in ['SIN DIA', 'nan', 'NaN']])
    
    col_f1, _ = st.columns([1, 2])
    with col_f1:
        dia_sel = st.selectbox(
            "📆 Seleccionar Día de Proceso:",
            options=["Todos los días"] + dias_disponibles,
            index=len(dias_disponibles) if len(dias_disponibles) > 0 else 0
        )

    df_filtrado = df if dia_sel == "Todos los días" else df[df['DIA DE PROCESO'] == dia_sel]

    # KPIs GLOBALES
    total_pallets = len(df_filtrado)
    df_conf = df_filtrado[df_filtrado['RESOLUCION'] == 'CONFORME']
    df_no_conf = df_filtrado[df_filtrado['RESOLUCION'] == 'NO CONFORME']
    
    cant_conf = len(df_conf)
    cant_no_conf = len(df_no_conf)
    
    df_obs_cliente = df_no_conf[df_no_conf['CLIENTE'] != 'SIN CLIENTE']
    cant_obs_cliente = len(df_obs_cliente)

    st.markdown(f"""
        <div class="kpi-container">
            <div class="kpi-card">
                <div class="kpi-title">Total Evaluados</div>
                <div class="kpi-value">{total_pallets}</div>
                <div class="kpi-sub">Día: {dia_sel}</div>
            </div>
            <div class="kpi-card ok">
                <div class="kpi-title">Conformes</div>
                <div class="kpi-value">{cant_conf}</div>
                <div class="kpi-sub">{(cant_conf/total_pallets*100 if total_pallets else 0):.1f}% del total</div>
            </div>
            <div class="kpi-card no-conf">
                <div class="kpi-title">No Conformes</div>
                <div class="kpi-value">{cant_no_conf}</div>
                <div class="kpi-sub">{(cant_no_conf/total_pallets*100 if total_pallets else 0):.1f}% del total</div>
            </div>
            <div class="kpi-card obs">
                <div class="kpi-title">Observados por Cliente</div>
                <div class="kpi-value">{cant_obs_cliente}</div>
                <div class="kpi-sub">De {cant_no_conf} No Conformes</div>
            </div>
        </div>
    """, unsafe_allow_html=True)

    # 1. NO CONFORMES POR CLIENTE
    st.subheader("🏢 1. No Conformes por Cliente")
    col_c1, col_c2 = st.columns([1, 1])

    with col_c1:
        if cant_obs_cliente > 0:
            df_cli_summary = df_obs_cliente['CLIENTE'].value_counts().reset_index()
            df_cli_summary.columns = ['Cliente', 'Pallets Observados']
            st.dataframe(df_cli_summary, use_container_width=True, hide_index=True)
        else:
            st.info("No existen pallets no conformes asignados a clientes en este filtro.")

    with col_c2:
        st.bar_chart(pd.DataFrame({
            'Cantidad': [cant_conf, cant_obs_cliente, cant_no_conf - cant_obs_cliente]
        }, index=['Conformes', 'Obs. Cliente', 'Sin Cliente']))

    st.markdown("---")

    # 2. NO CONFORMES POR VARIEDAD Y LOTE
    st.subheader("🍇 2. No Conformes por Variedad y Lote")
    
    if cant_no_conf > 0:
        agg_dict = {'N° PALLET': 'count'}
        if 'CAJAS EVALUADAS' in df_no_conf.columns:
            agg_dict['CAJAS EVALUADAS'] = 'sum'
        if 'TOTAL BAYAS AFECTADAS' in df_no_conf.columns:
            agg_dict['TOTAL BAYAS AFECTADAS'] = 'sum'

        df_var_lote = df_no_conf.groupby(['VARIEDAD', 'LOTE']).agg(agg_dict).reset_index()
        
        rename_map = {
            'N° PALLET': 'Pallets Observados',
            'CAJAS EVALUADAS': 'Cajas Evaluadas',
            'TOTAL BAYAS AFECTADAS': 'Total Bayas Afectadas'
        }
        df_var_lote = df_var_lote.rename(columns=rename_map).sort_values(by='Pallets Observados', ascending=False)
        
        col_vl1, col_vl2 = st.columns([3, 2])
        with col_vl1:
            st.dataframe(df_var_lote, use_container_width=True, hide_index=True)
        with col_vl2:
            st.bar_chart(df_var_lote.set_index('LOTE')['Pallets Observados'])
    else:
        st.info("No existen lotes ni variedades con observaciones no conformes.")

    st.markdown("---")

    # 3. DETALLE POR ORDEN BETA Y PALLETS
    st.subheader("📋 3. Detalle Registrado por Orden Beta y Pallet")

    cols_deseadas = [
        'ORDEN BETA', 'N° PALLET', 'CLIENTE', 'VARIEDAD', 'LOTE', 
        'DIA DE PROCESO', 'RESOLUCION', 'TOTAL BAYAS AFECTADAS', 
        '% BAYAS AFECTADOS', 'OBSERVACION_CLEAN', 'INSPECTOR', 'SEDE'
    ]
    cols_visibles = [c for c in cols_deseadas if c in df_filtrado.columns or c == 'OBSERVACION_CLEAN']

    tab_nc, tab_conf, tab_todo = st.tabs([
        f"🚨 Observados / No Conformes ({cant_no_conf})", 
        f"✅ Conformes ({cant_conf})", 
        f"📦 Registro Total ({total_pallets})"
    ])

    with tab_nc:
        if cant_no_conf > 0:
            st.dataframe(
                df_no_conf[cols_visibles].rename(columns={'OBSERVACION_CLEAN': 'OBSERVACION'}), 
                use_container_width=True, 
                hide_index=True
            )
        else:
            st.info("Sin registros no conformes.")

    with tab_conf:
        if cant_conf > 0:
            st.dataframe(
                df_conf[cols_visibles].rename(columns={'OBSERVACION_CLEAN': 'OBSERVACION'}), 
                use_container_width=True, 
                hide_index=True
            )
        else:
            st.info("Sin registros conformes.")

    with tab_todo:
        st.dataframe(
            df_filtrado[cols_visibles].rename(columns={'OBSERVACION_CLEAN': 'OBSERVACION'}), 
            use_container_width=True, 
            hide_index=True
        )
