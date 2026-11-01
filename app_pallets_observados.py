import streamlit as st
import pandas as pd
import requests
import io
import re
import unicodedata
from datetime import datetime

# ---------------------------------------------------------
# 1. CONFIGURACIÓN DE PÁGINA Y ESTILOS EN AZUL CORPORATIVO
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
# 2. TRANSFORMACIÓN DE ENLACE Y DESCARGA EN LÍNEA
# ---------------------------------------------------------
ENLACE_SHAREPOINT_PREDETERMINADO = "https://betacompe-my.sharepoint.com/:x:/g/personal/apizarro_beta_com_pe/IQDiOQlfZUjtQpijpRx5SMP_AQlGgiWqFG4fPK2eYW-e2OM?e=Haqj3X"

def normalizar_url_sharepoint(url: str) -> str:
    """ Convierte enlaces compartidos de SharePoint/OneDrive a URLs de descarga directa. """
    url = url.strip()
    if "sharepoint.com" in url or "onedrive.live.com" in url or "1drv.ms" in url:
        url = re.sub(r'[\?&]web=1', '', url)
        if "download=1" not in url:
            url += ("&download=1" if "?" in url else "?download=1")
    return url

@st.cache_data(ttl=15)
def cargar_datos_online(url: str):
    """ Descarga el archivo Excel validando la respuesta HTTP en Streamlit Cloud. """
    url_directa = normalizar_url_sharepoint(url)
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/115.0.0.0 Safari/537.36"
    }
    
    resp = requests.get(url_directa, headers=headers, timeout=20, allow_redirects=True)
    resp.raise_for_status()
    
    # Validar si SharePoint devolvió una página HTML (pantalla de inicio de sesión)
    content_type = resp.headers.get("Content-Type", "")
    if "text/html" in content_type or resp.content.startswith(b"<!DOCTYPE") or resp.content.startswith(b"<html"):
        raise PermissionError("SharePoint bloqueó el acceso anónimo y requiere inicio de sesión corporativo.")
        
    return pd.read_excel(io.BytesIO(resp.content))

def normalizar_dataframe(df_in: pd.DataFrame) -> pd.DataFrame:
    """ Detecta columnas sin importar acentos, mayúsculas o variaciones de nombre. """
    df = df_in.copy()
    df.columns = [str(c).strip() for c in df.columns]
    
    def limpiar_texto_llave(t):
        if not isinstance(t, str):
            return ""
        t_norm = unicodedata.normalize('NFD', t)
        return ''.join(c for c in t_norm if unicodedata.category(c) != 'Mn').strip().upper()

    # Mapeo dinámico de columnas por similitud de contenido
    col_map = {}
    for c in df.columns:
        c_clean = limpiar_texto_llave(c)
        if 'DIA' in c_clean and 'PROCESO' in c_clean:
            col_map[c] = 'DIA DE PROCESO'
        elif 'RESOLUC' in c_clean:
            col_map[c] = 'RESOLUCION'
        elif 'CLIENTE' in c_clean:
            col_map[c] = 'CLIENTE'
        elif 'VARIEDAD' in c_clean:
            col_map[c] = 'VARIEDAD'
        elif 'LOTE' in c_clean:
            col_map[c] = 'LOTE'
        elif 'PALLET' in c_clean:
            col_map[c] = 'N° PALLET'
        elif 'ORDEN' in c_clean and 'BETA' in c_clean:
            col_map[c] = 'ORDEN BETA'
        elif 'OBSERVAC' in c_clean:
            col_map[c] = 'OBSERVACION'

    df = df.rename(columns=col_map)

    # Creación segura de columnas por defecto para evitar KeyError
    if 'DIA DE PROCESO' not in df.columns:
        df['DIA DE PROCESO'] = 'SIN DIA'
    else:
        df['DIA DE PROCESO'] = df['DIA DE PROCESO'].astype(str).str.replace('.0', '', regex=False).str.strip()

    if 'RESOLUCION' not in df.columns:
        df['RESOLUCION'] = 'SIN ESTATUS'
    else:
        df['RESOLUCION'] = df['RESOLUCION'].fillna('SIN ESTATUS').astype(str).str.strip().str.upper()

    if 'CLIENTE' not in df.columns:
        df['CLIENTE'] = 'SIN CLIENTE'
    else:
        df['CLIENTE'] = df['CLIENTE'].fillna('SIN CLIENTE').astype(str).str.strip()

    if 'VARIEDAD' not in df.columns:
        df['VARIEDAD'] = 'DESCONOCIDO'
    else:
        df['VARIEDAD'] = df['VARIEDAD'].fillna('DESCONOCIDO').astype(str).str.strip()

    if 'LOTE' not in df.columns:
        df['LOTE'] = 'SIN LOTE'
    else:
        df['LOTE'] = df['LOTE'].astype(str).str.replace('.0', '', regex=False).str.strip()

    if 'OBSERVACION' not in df.columns:
        df['OBSERVACION_CLEAN'] = ''
    else:
        df['OBSERVACION_CLEAN'] = df['OBSERVACION'].fillna('').astype(str).str.strip()

    return df


# ---------------------------------------------------------
# 3. BARRA LATERAL (ENTORNO Y CONFIGURACIÓN)
# ---------------------------------------------------------
st.sidebar.title("🍇 Control de Calidad")
st.sidebar.markdown("**Pallets Observados - Uva 2026**")
st.sidebar.markdown("---")

url_input = st.sidebar.text_input(
    "🔗 Enlace SharePoint Web en línea:",
    value=ENLACE_SHAREPOINT_PREDETERMINADO,
    help="El sistema convierte automáticamente la URL a formato de descarga directa."
)

st.sidebar.markdown("---")
st.sidebar.markdown("**🚨 Respaldo por Archivo Local**")
archivo_manual = st.sidebar.file_uploader(
    "Subir Excel manualmente:",
    type=["xlsx", "xls"],
    help="Subida directa en caso de interrupción de acceso a SharePoint."
)

if st.sidebar.button("🔄 Refrescar Datos Nube"):
    st.cache_data.clear()

# ---------------------------------------------------------
# 4. LÓGICA DE CARGA EN LÍNEA
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
            with st.spinner("Conectando con SharePoint en tiempo real..."):
                df_raw = cargar_datos_online(target_url)
                hora_actual = datetime.now().strftime("%H:%M:%S")
                origen_mensaje = f"🟢 **TIEMPO REAL (SharePoint Web):** Conectado con éxito a las {hora_actual}"
                tipo_banner = "status-online"
        except PermissionError as pe:
            st.sidebar.error(f"⚠️ {pe}")
        except Exception as e:
            st.sidebar.warning(f"⚠️ No se pudo obtener el archivo desde el enlace web ({e}).")

# ---------------------------------------------------------
# 5. DASHBOARD EJECUTIVO DE CONTROL DE CALIDAD
# ---------------------------------------------------------
st.title("📊 Resumen Ejecutivo de Calidad - Planta Norte")

if df_raw is None:
    st.warning(
        "💡 **Esperando conexión:** Si SharePoint requiere inicio de sesión en Streamlit Cloud, "
        "sube el Excel mediante la opción **'Subir Excel manualmente'** en la barra lateral."
    )
else:
    st.markdown(f'<div class="status-banner {tipo_banner}">{origen_mensaje}</div>', unsafe_allow_html=True)
    df = normalizar_dataframe(df_raw)

    # FILTRO SEGURO POR DÍA DE PROCESO
    dias_disponibles = sorted([d for d in df['DIA DE PROCESO'].unique() if d != 'SIN DIA'])
    
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
            st.success("No existen pallets no conformes asignados a clientes en este filtro.")

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
        st.success("No existen lotes ni variedades con observaciones no conformes.")

    st.markdown("---")

    # 3. TABLAS DETALLADAS POR ORDEN BETA Y PALLETS
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
