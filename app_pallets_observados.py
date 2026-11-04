import streamlit as st
import pandas as pd
import requests
import io
import re
import unicodedata
import matplotlib.pyplot as plt
from datetime import datetime

# ---------------------------------------------------------
# 1. CONFIGURACIÓN DE PÁGINA Y ESTILOS CORPORATIVOS
# ---------------------------------------------------------
st.set_page_config(
    page_title="Control de Pallets Observados - Uva 2026",
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
# 2. GENERADOR DE GRÁFICOS ESTÁTICOS Y LIMPIOS (MATPLOTLIB)
# ---------------------------------------------------------
def crear_grafico_barras_estatico(serie_data: pd.Series, color_barras='#002060'):
    """ Genera una imagen estática de barras sin eje Y, sin cuadrículas y con etiquetas de datos. """
    fig, ax = plt.subplots(figsize=(6, 3.2), dpi=120)
    
    x_labels = [str(k) for k in serie_data.index]
    y_values = list(serie_data.values)
    
    bars = ax.bar(x_labels, y_values, color=color_barras, width=0.55, edgecolor='none')
    
    # Eliminar cuadrícula y eje Y
    ax.grid(False)
    ax.yaxis.set_visible(False)
    
    # Eliminar líneas del marco (spines) excepto la base inferior
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
    ax.spines['left'].set_visible(False)
    ax.spines['bottom'].set_color('#CBD5E1')
    ax.spines['bottom'].set_linewidth(1)
    
    # Configuración de etiquetas en eje X
    plt.xticks(rotation=35, ha='right', fontsize=8.5, fontweight='bold', color='#1E293B')
    
    # Agregar etiquetas numéricas directas en la parte superior de cada barra
    for bar in bars:
        height = bar.get_height()
        if height > 0:
            ax.annotate(
                f'{int(height)}',
                xy=(bar.get_x() + bar.get_width() / 2, height),
                xytext=(0, 4),
                textcoords="offset points",
                ha='center', va='bottom',
                fontsize=9.5, fontweight='bold', color='#002060'
            )
                    
    if y_values and max(y_values) > 0:
        ax.set_ylim(0, max(y_values) * 1.22)
        
    plt.tight_layout()
    return fig


# ---------------------------------------------------------
# 3. TRANSFORMACIÓN DE ENLACE Y PROCESAMIENTO DE DATOS
# ---------------------------------------------------------
ENLACE_SHAREPOINT_PREDETERMINADO = "https://betacompe-my.sharepoint.com/:x:/g/personal/apizarro_beta_com_pe/IQDiOQlfZUjtQpijpRx5SMP_AQlGgiWqFG4fPK2eYW-e2OM?e=Haqj3X"

def normalizar_url_sharepoint(url: str) -> str:
    url = url.strip()
    if "sharepoint.com" in url or "onedrive.live.com" in url or "1drv.ms" in url:
        url = re.sub(r'[\?&]web=1', '', url)
        if "download=1" not in url:
            url += ("&download=1" if "?" in url else "?download=1")
    return url

@st.cache_data(ttl=10)
def cargar_datos_online(url: str):
    url_directa = normalizar_url_sharepoint(url)
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/115.0.0.0 Safari/537.36"
    }
    
    resp = requests.get(url_directa, headers=headers, timeout=20, allow_redirects=True)
    resp.raise_for_status()
    
    content_type = resp.headers.get("Content-Type", "")
    if "text/html" in content_type or resp.content.startswith(b"<!DOCTYPE") or resp.content.startswith(b"<html"):
        raise PermissionError("SharePoint bloqueó la descarga directa y solicita inicio de sesión.")
        
    return pd.read_excel(io.BytesIO(resp.content))

def clasificar_variedad_categoria(row):
    """ Clasifica las variedades simples e identifica pallets MIXTOS leyendo N° PALLET y VARIEDAD. """
    pallet = str(row.get('N° PALLET', '')).upper()
    variedad = str(row.get('VARIEDAD', '')).upper()
    
    # 1. Si contiene barra inclinada '/' o 'MIX' en la variedad o pallet -> MIXTO
    if '/' in pallet or 'MIX' in pallet or '/' in variedad or 'MIX' in variedad:
        return 'MIXTO'
    
    # 2. Variedades individuales
    if 'ALLISON' in variedad or 'SHEEGENE' in variedad or pallet.startswith('SH-'):
        return 'SHEGEENE 20-ALISSON'
    elif 'SWEET GLOBE' in variedad or 'IFG TEN' in variedad or 'IFGTEN' in pallet:
        return 'IFG TEN-SWEET GLOBE'
    elif 'BEBOP' in variedad or 'CELEBRATION' in variedad or 'IFG THREE' in variedad:
        return 'IFG THIRTY SEVEN-BEBOP'
    else:
        return variedad if variedad not in ['NAN', 'DESCONOCIDO', 'NONE', ''] else 'Otras Variedades'

def normalizar_dataframe_robusto(df_in: pd.DataFrame) -> pd.DataFrame:
    df = df_in.copy()
    df.columns = [str(c).strip() for c in df.columns]

    def norm_txt(t):
        if not isinstance(t, str):
            return ""
        t_norm = unicodedata.normalize('NFD', t)
        return ''.join(c for c in t_norm if unicodedata.category(c) != 'Mn').strip().upper()

    col_map = {}
    mapped_targets = set()
    
    for c in df.columns:
        c_clean = norm_txt(c)
        if 'RESOLUCION' not in mapped_targets and any(k in c_clean for k in ['RESOLUC', 'ESTADO', 'ESTATUS', 'DICTAMEN']):
            col_map[c] = 'RESOLUCION'
            mapped_targets.add('RESOLUCION')
        elif 'DIA DE PROCESO' not in mapped_targets and ('PROCESO' in c_clean or ('DIA' in c_clean and 'PROC' in c_clean)):
            col_map[c] = 'DIA DE PROCESO'
            mapped_targets.add('DIA DE PROCESO')
        elif 'CLIENTE' not in mapped_targets and any(k in c_clean for k in ['CLIENTE', 'DESTINO', 'RECIBIDOR']):
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
        elif 'OBSERVACION' not in mapped_targets and c == 'OBSERVACION':
            col_map[c] = 'OBSERVACION'
            mapped_targets.add('OBSERVACION')

    df = df.rename(columns=col_map)

    def obtener_serie(col_name):
        if col_name not in df.columns:
            return pd.Series([None] * len(df), index=df.index)
        data = df[col_name]
        if isinstance(data, pd.DataFrame):
            return data.iloc[:, 0]
        return data

    df['DIA DE PROCESO'] = obtener_serie('DIA DE PROCESO').fillna('SIN DIA').astype(str).str.replace('.0', '', regex=False).str.strip()
    df['CLIENTE'] = obtener_serie('CLIENTE').fillna('SIN CLIENTE').astype(str).str.strip()
    df['CLIENTE'] = df['CLIENTE'].replace({'nan': 'SIN CLIENTE', 'None': 'SIN CLIENTE', '': 'SIN CLIENTE'})
    df['VARIEDAD'] = obtener_serie('VARIEDAD').fillna('DESCONOCIDO').astype(str).str.strip()
    df['LOTE'] = obtener_serie('LOTE').fillna('SIN LOTE').astype(str).str.replace('.0', '', regex=False).str.strip()
    df['OBSERVACION_CLEAN'] = obtener_serie('OBSERVACION').fillna('').astype(str).str.strip()

    # Clasificación de Variedad y Mixtos
    df['VARIEDAD_CATEGORIA'] = df.apply(clasificar_variedad_categoria, axis=1)

    # Clasificación de Resolución
    def clasificar_resolucion(row):
        val = row.get('RESOLUCION')
        if pd.notna(val):
            s = str(val).replace('\xa0', ' ').strip().upper()
            if any(k in s for k in ['NO', 'OBS', 'RECH', 'DESAP', 'FAIL', 'DEFECT', 'INCONF']):
                return 'NO CONFORME'
            if any(k in s for k in ['CONF', 'OK', 'APROB', 'ACEPT', 'PASS', 'BUENO']):
                return 'CONFORME'
        
        bayas = row.get('TOTAL BAYAS AFECTADAS', 0)
        obs = str(row.get('OBSERVACION_CLEAN', '')).strip()
        cli = str(row.get('CLIENTE', '')).strip()
        
        bayas_val = 0
        try:
            bayas_val = float(bayas) if pd.notna(bayas) else 0
        except Exception:
            bayas_val = 0
            
        if bayas_val > 0 or (obs != '' and obs.lower() != 'nan') or (cli != '' and cli != 'SIN CLIENTE' and cli.lower() != 'nan'):
            return 'NO CONFORME'
            
        return 'SIN ESTATUS'

    df['RESOLUCION'] = df.apply(clasificar_resolucion, axis=1)

    return df


# ---------------------------------------------------------
# 4. BARRA LATERAL
# ---------------------------------------------------------
st.sidebar.title("🍇 FILTROS")
st.sidebar.markdown("**Pallets Observados - Uva 2026**")
st.sidebar.markdown("---")

url_input = st.sidebar.text_input(
    "🔗 Enlace SharePoint Web en línea:",
    value=ENLACE_SHAREPOINT_PREDETERMINADO,
    help="El sistema convierte automáticamente el link para descarga directa en tiempo real."
)

st.sidebar.markdown("---")
st.sidebar.markdown("**🚨 Respaldo Manual (Cargar Excel)**")
archivo_manual = st.sidebar.file_uploader(
    "Subir Excel de Space AG:",
    type=["xlsx", "xls"],
    help="Subida directa si SharePoint no está actualizado."
)

st.sidebar.markdown("---")
if st.sidebar.button("🔄 Limpiar Caché y Re-descargar"):
    st.cache_data.clear()
    st.sidebar.success("¡Caché limpiada! Re-descargando datos de la nube...")

# ---------------------------------------------------------
# 5. CARGA DE DATOS Y PRIORIZACIÓN
# ---------------------------------------------------------
df_raw = None
origen_mensaje = ""
tipo_banner = ""
registros_descargados = 0

if archivo_manual is not None:
    try:
        df_raw = pd.read_excel(archivo_manual)
        registros_descargados = len(df_raw)
        origen_mensaje = f"📁 **MODO MANUAL:** Mostrando datos del archivo subido ({archivo_manual.name}) — Total filas: {registros_descargados}"
        tipo_banner = "status-local"
    except Exception as e:
        st.error(f"Error al leer el archivo Excel: {e}")
else:
    target_url = url_input.strip()
    if target_url:
        try:
            with st.spinner("Descargando desde SharePoint en tiempo real..."):
                df_raw = cargar_datos_online(target_url)
                registros_descargados = len(df_raw)
                hora_actual = datetime.now().strftime("%H:%M:%S")
                origen_mensaje = f"🟢 **TIEMPO REAL (SharePoint Web):** Conectado con éxito a las {hora_actual} — Total filas en nube: {registros_descargados}"
                tipo_banner = "status-online"
        except PermissionError as pe:
            st.sidebar.error(f"⚠️ {pe}")
        except Exception as e:
            st.sidebar.warning(f"⚠️ No se pudo descargar desde el enlace web ({e}). Sube el Excel manualmente mientras se sincroniza SharePoint.")

# ---------------------------------------------------------
# 6. DASHBOARD PALLETS OBSERVADOS
# ---------------------------------------------------------
st.title("Resumen Pallets Observados - Uva 2026")

if df_raw is None:
    st.warning("💡 **Esperando conexión:** Carga el archivo en línea o sube el Excel de Space AG mediante la barra lateral.")
else:
    st.markdown(f'<div class="status-banner {tipo_banner}">{origen_mensaje}</div>', unsafe_allow_html=True)
    
    df = normalizar_dataframe_robusto(df_raw)

    # AUDITORÍA DE DATOS
    with st.expander("🔍Conexión y Fuente de Datos"):
        col_aud1, col_aud2 = st.columns(2)
        with col_aud1:
            st.write("**Total de registros leídos:**", len(df))
            st.write("**Días de Proceso detectados:**", sorted([str(d) for d in df['DIA DE PROCESO'].unique() if d != 'SIN DIA']))
        with col_aud2:
            st.write("**Desglose de Resoluciones:**", df['RESOLUCION'].value_counts().to_dict())
            st.write("**Clientes identificados:**", df[df['CLIENTE'] != 'SIN CLIENTE']['CLIENTE'].nunique())

    # FILTRO POR DÍA DE PROCESO
    dias_disponibles = sorted([d for d in df['DIA DE PROCESO'].unique() if d not in ['SIN DIA', 'nan', 'NaN']], key=lambda x: int(x) if str(x).isdigit() else 0)
    
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
    
    df_obs_con_cliente = df_no_conf[df_no_conf['CLIENTE'] != 'SIN CLIENTE']
    df_obs_sin_cliente = df_no_conf[df_no_conf['CLIENTE'] == 'SIN CLIENTE']
    
    cant_obs_con_cliente = len(df_obs_con_cliente)
    cant_obs_sin_cliente = len(df_obs_sin_cliente)

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
                <div class="kpi-title">Sin Cliente (Pendientes)</div>
                <div class="kpi-value">{cant_obs_sin_cliente}</div>
                <div class="kpi-sub">{cant_obs_con_cliente} con cliente asignado</div>
            </div>
        </div>
    """, unsafe_allow_html=True)

    # ---------------------------------------------------------
    # SECCIÓN 1: NO CONFORMES POR CLIENTE
    # ---------------------------------------------------------
    st.subheader("🏢 1. No Conformes por Cliente / Asignación")
    col_c1, col_c2 = st.columns([1, 1])

    with col_c1:
        if cant_no_conf > 0:
            df_cli_summary = df_no_conf['CLIENTE'].value_counts().reset_index()
            df_cli_summary.columns = ['Cliente / Estado', 'Pallets Observados']
            st.dataframe(df_cli_summary, use_container_width=True, hide_index=True)
        else:
            st.info("No existen pallets no conformes en este filtro.")

    with col_c2:
        if cant_no_conf > 0:
            serie_cli = df_no_conf['CLIENTE'].value_counts()
            fig_cli = crear_grafico_barras_estatico(serie_cli, color_barras='#002060')
            st.pyplot(fig_cli, use_container_width=True)

    st.markdown("---")

    # ---------------------------------------------------------
    # SECCIÓN 2: NO CONFORMES POR VARIEDAD (NUEVA TABLA Y LÓGICA)
    # ---------------------------------------------------------
    st.subheader("🍇 2. No Conformes por Categoría de Variedad")
    
    if cant_no_conf > 0:
        agg_var = {'N° PALLET': 'count'}
        if 'CAJAS EVALUADAS' in df_no_conf.columns:
            agg_var['CAJAS EVALUADAS'] = lambda x: int(pd.to_numeric(x, errors='coerce').fillna(0).sum())
        if 'TOTAL BAYAS AFECTADAS' in df_no_conf.columns:
            agg_var['TOTAL BAYAS AFECTADAS'] = lambda x: round(pd.to_numeric(x, errors='coerce').fillna(0).sum(), 1)

        df_var_summary = df_no_conf.groupby('VARIEDAD_CATEGORIA').agg(agg_var).reset_index()
        df_var_summary = df_var_summary.rename(columns={
            'VARIEDAD_CATEGORIA': 'Categoría de Variedad',
            'N° PALLET': 'Pallets Observados',
            'CAJAS EVALUADAS': 'Cajas Evaluadas',
            'TOTAL BAYAS AFECTADAS': 'Total Bayas Afectadas'
        }).sort_values(by='Pallets Observados', ascending=False)

        col_v1, col_v2 = st.columns([1, 1])
        with col_v1:
            st.dataframe(df_var_summary, use_container_width=True, hide_index=True)
        with col_v2:
            serie_var = df_no_conf['VARIEDAD_CATEGORIA'].value_counts()
            fig_var = crear_grafico_barras_estatico(serie_var, color_barras='#002060')
            st.pyplot(fig_var, use_container_width=True)
    else:
        st.info("No existen observaciones registradas por variedad.")

    st.markdown("---")

    # ---------------------------------------------------------
    # SECCIÓN 3: NO CONFORMES POR LOTE
    # ---------------------------------------------------------
    st.subheader("📦 3. No Conformes por Lote")
    
    if cant_no_conf > 0:
        agg_dict = {'N° PALLET': 'count'}
        if 'CAJAS EVALUADAS' in df_no_conf.columns:
            agg_dict['CAJAS EVALUADAS'] = lambda x: int(pd.to_numeric(x, errors='coerce').fillna(0).sum())
        if 'TOTAL BAYAS AFECTADAS' in df_no_conf.columns:
            agg_dict['TOTAL BAYAS AFECTADAS'] = lambda x: round(pd.to_numeric(x, errors='coerce').fillna(0).sum(), 1)

        df_var_lote = df_no_conf.groupby(['VARIEDAD', 'LOTE']).agg(agg_dict).reset_index()
        df_var_lote = df_var_lote.rename(columns={
            'N° PALLET': 'Pallets Observados',
            'CAJAS EVALUADAS': 'Cajas Evaluadas',
            'TOTAL BAYAS AFECTADAS': 'Total Bayas Afectadas'
        }).sort_values(by='Pallets Observados', ascending=False)
        
        col_vl1, col_vl2 = st.columns([1, 1])
        with col_vl1:
            st.dataframe(df_var_lote, use_container_width=True, hide_index=True)
        with col_vl2:
            serie_lote = df_no_conf['LOTE'].value_counts().head(8) # Mostrar los 8 lotes principales
            fig_lote = crear_grafico_barras_estatico(serie_lote, color_barras='#002060')
            st.pyplot(fig_lote, use_container_width=True)
    else:
        st.info("No existen lotes con observaciones no conformes.")

    st.markdown("---")

    # ---------------------------------------------------------
    # SECCIÓN 4: TABLAS DETALLADAS POR ORDEN BETA Y PALLETS
    # ---------------------------------------------------------
    st.subheader("📋 4. Detalle Registrado por Orden Beta y Pallet")

    cols_deseadas = [
        'ORDEN BETA', 'N° PALLET', 'CLIENTE', 'VARIEDAD_CATEGORIA', 'LOTE', 
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
                df_no_conf[cols_visibles].rename(columns={'OBSERVACION_CLEAN': 'OBSERVACION', 'VARIEDAD_CATEGORIA': 'VARIEDAD'}), 
                use_container_width=True, 
                hide_index=True
            )
        else:
            st.info("Sin registros no conformes.")

    with tab_conf:
        if cant_conf > 0:
            st.dataframe(
                df_conf[cols_visibles].rename(columns={'OBSERVACION_CLEAN': 'OBSERVACION', 'VARIEDAD_CATEGORIA': 'VARIEDAD'}), 
                use_container_width=True, 
                hide_index=True
            )
        else:
            st.info("Sin registros conformes.")

    with tab_todo:
        st.dataframe(
            df_filtrado[cols_visibles].rename(columns={'OBSERVACION_CLEAN': 'OBSERVACION', 'VARIEDAD_CATEGORIA': 'VARIEDAD'}), 
            use_container_width=True, 
            hide_index=True
        )
