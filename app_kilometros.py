import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from datetime import date
import os
import io

# ---------------------------------------------------------
# CONFIGURACIÓN DE PÁGINA Y ESTILOS
# ---------------------------------------------------------
st.set_page_config(
    page_title="Citybus Querétaro | Cumplimiento, Cobro Semanal y Flota",
    page_icon="🚌",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.markdown("""
<style>
    .block-container { padding-top: 1.4rem; padding-bottom: 2rem; }
    div[data-testid="metric-container"] {
        background-color: #f8fafc;
        border: 1px solid #e2e8f0;
        padding: 14px 18px;
        border-radius: 10px;
        box-shadow: 0 1px 2px rgba(0,0,0,0.03);
    }
    .stTabs [data-baseweb="tab-list"] { gap: 8px; }
    .stTabs [data-baseweb="tab"] {
        padding: 8px 18px;
        border-radius: 6px 6px 0 0;
        font-weight: 600;
    }
</style>
""", unsafe_allow_html=True)

# ---------------------------------------------------------
# AUTENTICACIÓN Y PROTECCIÓN DE DATOS SENSIBLES
# ---------------------------------------------------------
def verificar_acceso():
    if st.session_state.get("autenticado", False):
        return True

    st.markdown("## 🔒 Acceso Restringido | Citybus Querétaro")
    st.caption("Información operativa y financiera confidencial. Ingresa tu clave para continuar.")

    with st.form("login_form"):
        password_ingresada = st.text_input("Contraseña de acceso:", type="password")
        enviar = st.form_submit_button("Ingresar", type="primary", use_container_width=True)

    if enviar:
        try:
            clave_correcta = st.secrets["APP_PASSWORD"]
        except Exception:
            clave_correcta = "CitybusQro2026*"

        if password_ingresada == clave_correcta:
            st.session_state["autenticado"] = True
            st.rerun()
        else:
            st.error("❌ Contraseña incorrecta.")

    st.stop()

verificar_acceso()

# ---------------------------------------------------------
# CONSTANTES Y REGLAS DE CONTRATO (CITYBUS QUERÉTARO)
# ---------------------------------------------------------
ARCHIVO_DEFAULT = "KM OCTUBRE NOVIEMBRE 18.xlsx"
ARCHIVO_LOCAL_DB = "historico_citybus_consolidado.xlsx"

TARIFA_TRONCAL = 43.00                 # C51 y C52 ($43.00 / km a cobrar)
TARIFA_REGULAR = 34.144094736842106    # Resto de rutas ($34.1441 / km a cobrar; Realizados * 0.95)
META_VERDE = 95.0
META_AMARILLA = 85.0

MAPA_RUTA_PATIO = {
    "C30": "EPQ", "C37": "EPQ", "C38": "EPQ", "C39": "EPQ",
    "C47": "CDM", "C51": "CDM", "C52": "EPQ", "C62": "EPQ",
    "C68": "EPQ", "L100": "CDM", "L154": "EPQ", "L155": "EPQ",
    "L52": "EPQ", "L53": "CDM", "L54": "CDM", "L55": "CDM",
    "L56": "CDM", "L59": "EPQ", "L61": "EPQ", "L62": "EPQ"
}

RUTAS_ORDEN = [
    "C52", "L62", "L52", "C39", "C51", "C30", "L56", "L154", "C38", "L53",
    "L100", "C62", "L59", "L155", "C37", "L61", "C47", "L54", "C68"
]

MOTIVOS_COLS = [
    "Km operaciones", "Km mantenimiento", "Km carga de gas", "Km tráfico",
    "Km conductor", "Km accidentes", "Km validador", "Km recurso bus",
    "Km recurso conductor"
]

TURNOS_FLOTA = ["Matutino", "Vespertino"]

FLOTA_PATIO_COLS = [
    "FECHA", "TURNO", "PATIO", "Buses_despacho", "Buses_NRTEC",
    "Buses_inactivos", "Inactivos_con_tarjeta", "Inactivos_sin_tarjeta"
]

FLOTA_RUTA_COLS = [
    "FECHA", "TURNO", "RUTA", "PATIO", "Buses_despacho", "Buses_NRTEC",
    "Buses_inactivos", "Inactivos_con_tarjeta", "Inactivos_sin_tarjeta"
]

# ---------------------------------------------------------
# FUNCIONES DE CÁLCULO Y PERSISTENCIA
# ---------------------------------------------------------
def calcular_metricas_fila(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df["Programados completos"] = pd.to_numeric(df["Programados completos"], errors="coerce").fillna(0.0)
    df["Realizados"] = pd.to_numeric(df["Realizados"], errors="coerce").fillna(0.0)
    df["Km penalizado"] = pd.to_numeric(df.get("Km penalizado", 0.0), errors="coerce").fillna(0.0)

    if "KM PERDIDOS" not in df.columns:
        df["KM PERDIDOS"] = (df["Programados completos"] - df["Realizados"]).clip(lower=0.0)
    else:
        df["KM PERDIDOS"] = pd.to_numeric(df["KM PERDIDOS"], errors="coerce").fillna(
            (df["Programados completos"] - df["Realizados"]).clip(lower=0.0)
        )

    es_troncal = df["RUTA"].isin(["C51", "C52"])

    if "KM A COBRAR" not in df.columns:
        df["KM A COBRAR"] = None
    df["KM A COBRAR"] = pd.to_numeric(df["KM A COBRAR"], errors="coerce")

    km_cobrar_calc = df["Realizados"] * 0.95
    km_cobrar_troncal = (df["Programados completos"] - df["Km penalizado"]).clip(lower=0.0)
    df["KM A COBRAR"] = df["KM A COBRAR"].fillna(
        km_cobrar_troncal.where(es_troncal, km_cobrar_calc)
    )

    if "INGRESO" not in df.columns:
        df["INGRESO"] = None
    df["INGRESO"] = pd.to_numeric(df["INGRESO"], errors="coerce")

    tarifa = pd.Series(TARIFA_REGULAR, index=df.index).where(~es_troncal, TARIFA_TRONCAL)
    df["INGRESO"] = df["INGRESO"].fillna(df["KM A COBRAR"] * tarifa)

    km_cobrar_potencial = df["Programados completos"].where(es_troncal, df["Programados completos"] * 0.95)
    df["INGRESO_POTENCIAL"] = km_cobrar_potencial * tarifa
    df["INGRESO_PERDIDO"] = (df["INGRESO_POTENCIAL"] - df["INGRESO"]).clip(lower=0.0)

    df["ALERTA_CONCILIACION"] = es_troncal & (df["Km penalizado"] > 0) & (
        (df["KM A COBRAR"] - df["Programados completos"]).abs() < 0.05
    )
    if "ES_ESTIMADO" not in df.columns:
        df["ES_ESTIMADO"] = False
    return df


@st.cache_data
def leer_excel_bytes(file_bytes: bytes):
    xls = pd.ExcelFile(io.BytesIO(file_bytes))
    df_prod = pd.read_excel(xls, sheet_name="KM PRODUCCIÓN")
    df_prod.columns = df_prod.columns.str.strip()
    df_prod["FECHA"] = pd.to_datetime(df_prod["FECHA"]).dt.normalize()
    df_prod["PATIO"] = df_prod["PATIO"].astype(str).str.strip().str.upper()
    df_prod["RUTA"] = df_prod["RUTA"].astype(str).str.strip().str.upper()
    if "ES_ESTIMADO" not in df_prod.columns:
        df_prod["ES_ESTIMADO"] = False
    df_prod = calcular_metricas_fila(df_prod)

    if "DESGLOSE PERDIDOS" in xls.sheet_names:
        df_perd = pd.read_excel(xls, sheet_name="DESGLOSE PERDIDOS")
        df_perd.columns = df_perd.columns.str.strip()
        col_fecha = "Fecha" if "Fecha" in df_perd.columns else "FECHA"
        col_patio = "Patio" if "Patio" in df_perd.columns else "PATIO"
        col_ruta = "Ruta" if "Ruta" in df_perd.columns else "RUTA"
        df_perd["FECHA"] = pd.to_datetime(df_perd[col_fecha]).dt.normalize()
        df_perd["PATIO"] = df_perd[col_patio].astype(str).str.strip().str.upper()
        df_perd["RUTA"] = df_perd[col_ruta].astype(str).str.strip().str.upper()
        for col in MOTIVOS_COLS:
            if col not in df_perd.columns:
                df_perd[col] = 0.0
            df_perd[col] = pd.to_numeric(df_perd[col], errors="coerce").fillna(0.0)
        df_perd = df_perd[["FECHA", "PATIO", "RUTA"] + MOTIVOS_COLS]
    else:
        df_perd = pd.DataFrame(columns=["FECHA", "PATIO", "RUTA"] + MOTIVOS_COLS)

    if "FLOTA MATUTINA" in xls.sheet_names:
        df_flota = pd.read_excel(xls, sheet_name="FLOTA MATUTINA")
        df_flota.columns = df_flota.columns.str.strip()
        df_flota["FECHA"] = pd.to_datetime(df_flota["FECHA"]).dt.normalize()
        if "TURNO" not in df_flota.columns:
            df_flota["TURNO"] = "Matutino"
        for c in ["Buses_despacho", "Buses_NRTEC", "Buses_inactivos", "Inactivos_con_tarjeta", "Inactivos_sin_tarjeta"]:
            if c not in df_flota.columns:
                df_flota[c] = 0
        df_flota = df_flota[FLOTA_PATIO_COLS]
    else:
        df_flota = pd.DataFrame(columns=FLOTA_PATIO_COLS)

    if "FLOTA RUTAS INACTIVAS" in xls.sheet_names:
        df_flota_rutas = pd.read_excel(xls, sheet_name="FLOTA RUTAS INACTIVAS")
        df_flota_rutas.columns = df_flota_rutas.columns.str.strip()
        df_flota_rutas["FECHA"] = pd.to_datetime(df_flota_rutas["FECHA"]).dt.normalize()
        if "TURNO" not in df_flota_rutas.columns:
            df_flota_rutas["TURNO"] = "Matutino"
        for c in ["Buses_despacho", "Buses_NRTEC", "Buses_inactivos", "Inactivos_con_tarjeta", "Inactivos_sin_tarjeta"]:
            if c not in df_flota_rutas.columns:
                df_flota_rutas[c] = 0
        df_flota_rutas = df_flota_rutas[FLOTA_RUTA_COLS]
    else:
        df_flota_rutas = pd.DataFrame(columns=FLOTA_RUTA_COLS)

    return df_prod, df_perd, df_flota, df_flota_rutas


def inicializar_datos():
    if "df_prod" not in st.session_state:
        ruta_cargar = ARCHIVO_LOCAL_DB if os.path.exists(ARCHIVO_LOCAL_DB) else (
            ARCHIVO_DEFAULT if os.path.exists(ARCHIVO_DEFAULT) else None
        )
        if ruta_cargar:
            with open(ruta_cargar, "rb") as f:
                p, d, fl, fl_r = leer_excel_bytes(f.read())
            st.session_state["df_prod"] = p
            st.session_state["df_perd"] = d
            st.session_state["df_flota"] = fl
            st.session_state["df_flota_rutas"] = fl_r
        else:
            st.session_state["df_prod"] = pd.DataFrame()
            st.session_state["df_perd"] = pd.DataFrame()
            st.session_state["df_flota"] = pd.DataFrame(columns=FLOTA_PATIO_COLS)
            st.session_state["df_flota_rutas"] = pd.DataFrame(columns=FLOTA_RUTA_COLS)


def guardar_consolidado_excel(
    df_prod: pd.DataFrame,
    df_perd: pd.DataFrame,
    df_flota: pd.DataFrame,
    df_flota_rutas: pd.DataFrame
) -> bytes:
    output = io.BytesIO()
    cols_prod_export = [
        "FECHA", "PATIO", "RUTA", "Programados completos", "Realizados",
        "Km penalizado", "KM A COBRAR", "INGRESO", "KM PERDIDOS", "ES_ESTIMADO"
    ]
    for c in cols_prod_export:
        if c not in df_prod.columns:
            df_prod[c] = False if c == "ES_ESTIMADO" else 0.0

    df_perd_export = df_perd.rename(columns={"FECHA": "Fecha", "RUTA": "Ruta", "PATIO": "Patio"})
    cols_perd_export = ["Fecha", "Ruta", "Patio"] + MOTIVOS_COLS
    for c in cols_perd_export:
        if c not in df_perd_export.columns:
            df_perd_export[c] = 0.0

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df_prod[cols_prod_export].to_excel(writer, sheet_name="KM PRODUCCIÓN", index=False)
        df_perd_export[cols_perd_export].to_excel(writer, sheet_name="DESGLOSE PERDIDOS", index=False)
        if not df_flota.empty:
            df_flota.to_excel(writer, sheet_name="FLOTA MATUTINA", index=False)
        if not df_flota_rutas.empty:
            df_flota_rutas.to_excel(writer, sheet_name="FLOTA RUTAS INACTIVAS", index=False)
    return output.getvalue()


def badge_semaforo(pct: float) -> str:
    if pct >= META_VERDE:
        return f"🟢 {pct:.1f}%"
    elif pct >= META_AMARILLA:
        return f"🟡 {pct:.1f}%"
    return f"🔴 {pct:.1f}%"


inicializar_datos()

# ---------------------------------------------------------
# BARRA LATERAL: CARGA DE ARCHIVOS Y FILTROS
# ---------------------------------------------------------
with st.sidebar:
    st.title("🚌 Citybus Querétaro")
    st.caption("Control Operativo y Flota")

    st.subheader("1. Cargar / Actualizar Excel")
    archivo_subido = st.file_uploader(
        "Sube el Excel conciliado (.xlsx)",
        type=["xlsx"],
        help="Al subir un Excel oficial, sobrescribe los días estimados con las cifras conciliadas."
    )
    modo_carga = st.radio(
        "Al subir Excel:",
        ["Actualizar fechas", "Reemplazar la base"],
        horizontal=True
    )
    if archivo_subido is not None:
        if st.button("Procesar Excel subido", use_container_width=True, type="primary"):
            nuevo_prod, nuevo_perd, nuevo_flota, nuevo_flota_r = leer_excel_bytes(archivo_subido.getvalue())
            if modo_carga == "Reemplazar la base" or st.session_state["df_prod"].empty:
                st.session_state["df_prod"] = nuevo_prod
                st.session_state["df_perd"] = nuevo_perd
                if not nuevo_flota.empty:
                    st.session_state["df_flota"] = nuevo_flota
                if not nuevo_flota_r.empty:
                    st.session_state["df_flota_rutas"] = nuevo_flota_r
            else:
                fechas_nuevas = nuevo_prod["FECHA"].unique()
                prod_ant = st.session_state["df_prod"][~st.session_state["df_prod"]["FECHA"].isin(fechas_nuevas)]
                perd_ant = st.session_state["df_perd"][~st.session_state["df_perd"]["FECHA"].isin(fechas_nuevas)]
                st.session_state["df_prod"] = pd.concat([prod_ant, nuevo_prod], ignore_index=True)
                st.session_state["df_perd"] = pd.concat([perd_ant, nuevo_perd], ignore_index=True)

            with open(ARCHIVO_LOCAL_DB, "wb") as f:
                f.write(guardar_consolidado_excel(
                    st.session_state["df_prod"],
                    st.session_state["df_perd"],
                    st.session_state["df_flota"],
                    st.session_state["df_flota_rutas"]
                ))
            st.success(f"Base actualizada ({len(nuevo_prod)} registros).")
            st.rerun()

    st.divider()

    df_prod = st.session_state["df_prod"]
    df_perd = st.session_state["df_perd"]
    df_flota = st.session_state["df_flota"]
    df_flota_rutas = st.session_state["df_flota_rutas"]

    if df_prod.empty:
        st.warning("No hay datos cargados. Sube tu Excel arriba para comenzar.")
        if st.button("🔒 Cerrar sesión", use_container_width=True):
            st.session_state["autenticado"] = False
            st.rerun()
        st.stop()

    st.subheader("2. Filtros del Dashboard")
    min_fecha = df_prod["FECHA"].min().date()
    max_fecha = max(
        df_prod["FECHA"].max().date(),
        df_flota["FECHA"].max().date() if not df_flota.empty else df_prod["FECHA"].max().date()
    )

    rango_fechas = st.date_input(
        "Rango de fechas",
        value=(min_fecha, max_fecha),
        min_value=min_fecha,
        max_value=max_fecha
    )

    patios_sel = st.multiselect(
        "Patio",
        options=["EPQ", "CDM"],
        default=["EPQ", "CDM"]
    )

    rutas_disponibles = sorted(df_prod["RUTA"].unique())
    rutas_sel = st.multiselect(
        "Rutas (vacío = todas)",
        options=rutas_disponibles,
        default=[]
    )

    st.divider()
    excel_bytes = guardar_consolidado_excel(df_prod, df_perd, df_flota, df_flota_rutas)
    st.download_button(
        label="Descargar Histórico Consolidado (.xlsx)",
        data=excel_bytes,
        file_name=f"Citybus_Historico_{max_fecha.strftime('%Y%m%d')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )

    if st.button("🔒 Cerrar sesión", use_container_width=True):
        st.session_state["autenticado"] = False
        st.rerun()

# Aplicar filtros generales
if isinstance(rango_fechas, tuple) and len(rango_fechas) == 2:
    f_ini, f_fin = pd.to_datetime(rango_fechas[0]), pd.to_datetime(rango_fechas[1])
else:
    f_ini = f_fin = pd.to_datetime(rango_fechas[0] if isinstance(rango_fechas, tuple) else rango_fechas)

mask_prod = (df_prod["FECHA"] >= f_ini) & (df_prod["FECHA"] <= f_fin) & (df_prod["PATIO"].isin(patios_sel))
if rutas_sel:
    mask_prod = mask_prod & (df_prod["RUTA"].isin(rutas_sel))
df_f = df_prod[mask_prod].copy()

mask_perd = (df_perd["FECHA"] >= f_ini) & (df_perd["FECHA"] <= f_fin) & (df_perd["PATIO"].isin(patios_sel))
if rutas_sel:
    mask_perd = mask_perd & (df_perd["RUTA"].isin(rutas_sel))
df_perd_f = df_perd[mask_perd].copy()

# ---------------------------------------------------------
# ENCABEZADO Y SELECTOR DE UNIDAD
# ---------------------------------------------------------
col_t1, col_t2 = st.columns([3, 1.3])
with col_t1:
    st.title("Cumplimiento y Cobro")
    dias_count = df_f["FECHA"].nunique()
    st.caption(f"Periodo activo: **{f_ini.strftime('%d/%m/%Y')}** al **{f_fin.strftime('%d/%m/%Y')}** ({dias_count} días operativos)")

with col_t2:
    unidad_vista = st.radio(
        "Modo de visualización:",
        ["Ambos (KM + MXN)", "Solo Kilómetros (km)", "Solo Financiero ($ MXN)"],
        index=0
    )

# ---------------------------------------------------------
# PESTAÑAS PRINCIPALES (4 SECCIONES)
# ---------------------------------------------------------
tab_semanal, tab_dash, tab_rutas, tab_captura = st.tabs([
    "Cobro Semanal",
    "Dashboard Ejecutivo & Flota",
    "Matriz por Ruta y Unidades Inactivas",
    "Captura Diaria"
])

# =========================================================
# TAB 1: COBRO SEMANAL (Lunes a Domingo)
# =========================================================
with tab_semanal:
    st.subheader("Corte de Cobro Semanal")

    df_sem_base = df_prod[df_prod["PATIO"].isin(patios_sel)].copy()
    if rutas_sel:
        df_sem_base = df_sem_base[df_sem_base["RUTA"].isin(rutas_sel)]

    df_sem_base["SEM_INI"] = df_sem_base["FECHA"].dt.to_period("W-SUN").apply(lambda r: r.start_time)
    df_sem_base["SEM_FIN"] = df_sem_base["FECHA"].dt.to_period("W-SUN").apply(lambda r: r.end_time)

    resumen_sem = df_sem_base.groupby(["SEM_INI", "SEM_FIN"], as_index=False).agg(
        Días=("FECHA", "nunique"),
        Km_Programados=("Programados completos", "sum"),
        Km_Realizados=("Realizados", "sum"),
        Km_Penalizados=("Km penalizado", "sum"),
        Km_a_Cobrar=("KM A COBRAR", "sum"),
        Monto_a_Cobrar_MXN=("INGRESO", "sum"),
        Ingreso_Perdido_MXN=("INGRESO_PERDIDO", "sum"),
        Tiene_Estimados=("ES_ESTIMADO", "any")
    ).sort_values("SEM_INI", ascending=False)

    resumen_sem["Semana"] = (
        resumen_sem["SEM_INI"].dt.strftime("%d/%b") + " al " + resumen_sem["SEM_FIN"].dt.strftime("%d/%b/%Y")
    )
    resumen_sem["Cumplimiento %"] = (resumen_sem["Km_Realizados"] / resumen_sem["Km_Programados"] * 100).round(2)
    resumen_sem["Estatus"] = resumen_sem.apply(
        lambda r: ("⚠️ Parcial" if r["Días"] < 7 else "✅ Completa") + (" (c/ est.)" if r["Tiene_Estimados"] else ""),
        axis=1
    )

    sem_actual = resumen_sem.iloc[0]
    sem_completa_df = resumen_sem[resumen_sem["Días"] == 7]
    ult_completa = sem_completa_df.iloc[0] if not sem_completa_df.empty else sem_actual

    c_s1, c_s2, c_s3, c_s4 = st.columns(4)
    c_s1.metric(
        f"Semana en Curso ({sem_actual['Semana']})",
        f"${sem_actual['Monto_a_Cobrar_MXN']:,.2f}",
        delta=f"{int(sem_actual['Días'])} de 7 días capturados",
        delta_color="off"
    )
    c_s2.metric(
        f"Última Semana Completa ({ult_completa['Semana']})",
        f"${ult_completa['Monto_a_Cobrar_MXN']:,.2f}",
        delta=f"{ult_completa['Km_a_Cobrar']:,.1f} km a cobrar",
        delta_color="normal"
    )
    c_s3.metric(
        "Promedio Semanal (Semanas 7 días)",
        f"${sem_completa_df['Monto_a_Cobrar_MXN'].mean():,.2f}" if not sem_completa_df.empty else f"${sem_actual['Monto_a_Cobrar_MXN']:,.2f}",
        delta=f"Cumpl. prom: {sem_completa_df['Cumplimiento %'].mean():.2f}%" if not sem_completa_df.empty else "",
        delta_color="off"
    )
    c_s4.metric(
        "Dejado de Cobrar (Últ. Sem. Completa)",
        f"${ult_completa['Ingreso_Perdido_MXN']:,.2f}",
        delta=f"Penalizados: {ult_completa['Km_Penalizados']:,.1f} km",
        delta_color="inverse"
    )

    st.markdown("---")

    cols_sem_tabla = [
        "Semana", "Estatus", "Días", "Cumplimiento %",
        "Km_Programados", "Km_Realizados", "Km_Penalizados",
        "Km_a_Cobrar", "Monto_a_Cobrar_MXN", "Ingreso_Perdido_MXN"
    ]
    st.dataframe(
        resumen_sem[cols_sem_tabla].style.format({
            "Cumplimiento %": "{:.2f}%",
            "Km_Programados": "{:,.2f}",
            "Km_Realizados": "{:,.2f}",
            "Km_Penalizados": "{:,.2f}",
            "Km_a_Cobrar": "{:,.2f}",
            "Monto_a_Cobrar_MXN": "${:,.2f}",
            "Ingreso_Perdido_MXN": "${:,.2f}"
        }),
        use_container_width=True,
        hide_index=True
    )

    st.markdown("### Desglose Detallado por Semana")
    semana_elegida = st.selectbox(
        "Selecciona la semana que quieres desglosar por Patio y Ruta:",
        options=resumen_sem["Semana"].tolist(),
        index=0
    )

    df_sem_sel = df_sem_base[
        (df_sem_base["SEM_INI"].dt.strftime("%d/%b") + " al " + df_sem_base["SEM_FIN"].dt.strftime("%d/%b/%Y")) == semana_elegida
    ].copy()

    col_sd1, col_sd2 = st.columns([1, 1.5])
    with col_sd1:
        st.markdown("**Cobro por Patio en la Semana**")
        patio_sem = df_sem_sel.groupby("PATIO", as_index=False).agg(
            Días=("FECHA", "nunique"),
            Km_a_Cobrar=("KM A COBRAR", "sum"),
            Monto_a_Cobrar_MXN=("INGRESO", "sum"),
            Dejado_de_Cobrar_MXN=("INGRESO_PERDIDO", "sum")
        )
        st.dataframe(
            patio_sem.style.format({
                "Km_a_Cobrar": "{:,.2f}",
                "Monto_a_Cobrar_MXN": "${:,.2f}",
                "Dejado_de_Cobrar_MXN": "${:,.2f}"
            }),
            use_container_width=True,
            hide_index=True
        )

        dia_sem = df_sem_sel.groupby("FECHA", as_index=False).agg(
            Km_a_Cobrar=("KM A COBRAR", "sum"),
            Monto_MXN=("INGRESO", "sum")
        ).sort_values("FECHA")
        dia_sem["Día"] = dia_sem["FECHA"].dt.strftime("%a %d/%b")
        fig_dia_sem = px.bar(
            dia_sem, x="Día", y="Monto_MXN", text_auto="$,.0f",
            title="Cobro Diario dentro de la Semana",
            color_discrete_sequence=["#16a34a"]
        )
        fig_dia_sem.update_layout(height=280, margin=dict(l=10, r=10, t=40, b=10), yaxis_title="MXN ($)")
        st.plotly_chart(fig_dia_sem, use_container_width=True)

    with col_sd2:
        st.markdown("**Desglose por Ruta para Facturación Semanal**")
        ruta_sem = df_sem_sel.groupby(["RUTA", "PATIO"], as_index=False).agg(
            Días=("FECHA", "nunique"),
            Programados=("Programados completos", "sum"),
            Realizados=("Realizados", "sum"),
            Penalizados=("Km penalizado", "sum"),
            Km_a_Cobrar=("KM A COBRAR", "sum"),
            Monto_a_Cobrar_MXN=("INGRESO", "sum")
        ).sort_values("Monto_a_Cobrar_MXN", ascending=False)

        st.dataframe(
            ruta_sem.style.format({
                "Programados": "{:,.2f}",
                "Realizados": "{:,.2f}",
                "Penalizados": "{:,.2f}",
                "Km_a_Cobrar": "{:,.2f}",
                "Monto_a_Cobrar_MXN": "${:,.2f}"
            }),
            use_container_width=True,
            hide_index=True,
            height=410
        )

        out_sem = io.BytesIO()
        with pd.ExcelWriter(out_sem, engine="openpyxl") as writer:
            resumen_sem[cols_sem_tabla].to_excel(writer, sheet_name="Resumen Todas las Semanas", index=False)
            patio_sem.to_excel(writer, sheet_name="Cobro por Patio", index=False)
            ruta_sem.to_excel(writer, sheet_name="Cobro por Ruta", index=False)
        st.download_button(
            label=f"Descargar Excel de Cobro ({semana_elegida})",
            data=out_sem.getvalue(),
            file_name=f"Cobro_Semanal_Citybus_{semana_elegida.replace('/', '-').replace(' ', '_')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )

# =========================================================
# TAB 2: DASHBOARD EJECUTIVO & FLOTA
# =========================================================
with tab_dash:
    if df_f.empty:
        st.warning("No hay registros con los filtros seleccionados.")
    else:
        tot_prog = df_f["Programados completos"].sum()
        tot_real = df_f["Realizados"].sum()
        tot_perd = df_f["KM PERDIDOS"].sum()
        tot_cobrar = df_f["KM A COBRAR"].sum()
        tot_ing = df_f["INGRESO"].sum()
        tot_ing_pot = df_f["INGRESO_POTENCIAL"].sum()
        tot_ing_perd = df_f["INGRESO_PERDIDO"].sum()
        cumpl_global = (tot_real / tot_prog * 100) if tot_prog > 0 else 0.0
        brecha_95_km = max(0.0, (tot_prog * 0.95) - tot_real)

        k1, k2, k3, k4, k5 = st.columns(5)
        k1.metric("Cumplimiento Ponderado", f"{cumpl_global:.2f}%", delta=f"{cumpl_global - META_VERDE:+.2f}% vs Meta 95%")
        k2.metric("Kilómetros Realizados", f"{tot_real:,.0f} km", delta=f"de {tot_prog:,.0f} km prog.", delta_color="off")
        k3.metric("Kilómetros Perdidos", f"{tot_perd:,.0f} km", delta=f"Faltan {brecha_95_km:,.0f} km p/ 95%", delta_color="inverse")
        k4.metric("Ingreso Real Devengado", f"${tot_ing:,.0f}", delta=f"{tot_cobrar:,.0f} km a cobrar", delta_color="off")
        k5.metric("Ingreso Dejado de Percibir", f"${tot_ing_perd:,.0f}", delta=f"-{(tot_ing_perd / tot_ing_pot * 100) if tot_ing_pot else 0:.1f}% del potencial", delta_color="inverse")

        # Bloque de Estado de Flota e Inactivos (Matutino y Vespertino)
        if not df_flota.empty or not df_flota_rutas.empty:
            st.markdown("### Estado de Flota e Inactivos")
            ult_fecha_flota = df_flota["FECHA"].max() if not df_flota.empty else df_flota_rutas["FECHA"].max()

            col_ft1, col_ft2 = st.columns([2, 1])
            with col_ft1:
                st.caption(f"Mostrando último corte registrado: **{ult_fecha_flota.strftime('%d/%m/%Y')}**")
            with col_ft2:
                turnos_disp = df_flota[df_flota["FECHA"] == ult_fecha_flota]["TURNO"].unique().tolist() if not df_flota.empty else TURNOS_FLOTA
                if not turnos_disp:
                    turnos_disp = TURNOS_FLOTA
                turno_ver = st.selectbox(
                    "Corte de flota a visualizar:",
                    options=turnos_disp,
                    index=len(turnos_disp) - 1
                )

            fl_hoy = df_flota[
                (df_flota["FECHA"] == ult_fecha_flota) &
                (df_flota["PATIO"].isin(patios_sel)) &
                (df_flota["TURNO"] == turno_ver)
            ] if not df_flota.empty else pd.DataFrame()

            fl_r_hoy = df_flota_rutas[
                (df_flota_rutas["FECHA"] == ult_fecha_flota) &
                (df_flota_rutas["PATIO"].isin(patios_sel)) &
                (df_flota_rutas["TURNO"] == turno_ver)
            ] if not df_flota_rutas.empty else pd.DataFrame()

            f1, f2, f3, f4, f5 = st.columns(5)
            desp_tot = int(fl_hoy["Buses_despacho"].sum()) if not fl_hoy.empty else 0
            nrtec_tot = int(fl_hoy["Buses_NRTEC"].sum()) if not fl_hoy.empty else 0
            inact_tot = int(fl_hoy["Buses_inactivos"].sum()) if not fl_hoy.empty else 0
            inact_ct = int(fl_hoy["Inactivos_con_tarjeta"].sum()) if not fl_hoy.empty else 0
            inact_st = int(fl_hoy["Inactivos_sin_tarjeta"].sum()) if not fl_hoy.empty else 0

            f1.metric(f"Buses en Despacho ({turno_ver.split()[0]})", f"{desp_tot} buses", delta=ult_fecha_flota.strftime('%d/%b'), delta_color="off")
            f2.metric("Buses en Línea NRTEC", f"{nrtec_tot} buses", delta=f"{nrtec_tot - desp_tot:+d} vs despacho" if desp_tot else "Sin dif.", delta_color="normal")
            f3.metric("Buses INACTIVOS (Patios)", f"{inact_tot} buses", delta="Alerta operativa" if inact_tot > 0 else "Flota 100% activa", delta_color="inverse" if inact_tot > 0 else "normal")
            f4.metric("Inactivos CON Tarjeta", f"{inact_ct} buses", delta="Corte de recaudo", delta_color="off")
            f5.metric("Inactivos SIN Tarjeta", f"{inact_st} buses", delta="Sin tarjeta asignada", delta_color="off")

            fl_dia_ambos = df_flota[(df_flota["FECHA"] == ult_fecha_flota) & (df_flota["PATIO"].isin(patios_sel))] if not df_flota.empty else pd.DataFrame()
            if not fl_dia_ambos.empty and fl_dia_ambos["TURNO"].nunique() > 1:
                comp_turnos = fl_dia_ambos.groupby(["TURNO", "PATIO"], as_index=False)[
                    ["Buses_despacho", "Buses_NRTEC", "Buses_inactivos", "Inactivos_con_tarjeta", "Inactivos_sin_tarjeta"]
                ].sum()
                st.caption("**Comparativa Matutino vs. Vespertino por Patio:**")
                st.dataframe(
                    comp_turnos.rename(columns={
                        "TURNO": "Turno / Corte",
                        "PATIO": "Patio",
                        "Buses_despacho": "Despacho",
                        "Buses_NRTEC": "NRTEC",
                        "Buses_inactivos": "INACTIVOS",
                        "Inactivos_con_tarjeta": "Inact. c/ tarjeta",
                        "Inactivos_sin_tarjeta": "Inact. s/ tarjeta"
                    }),
                    use_container_width=True,
                    hide_index=True
                )

            if not fl_r_hoy.empty:
                st.caption(f"**Detalle por Ruta en {turno_ver} ({ult_fecha_flota.strftime('%d/%m/%Y')}):**")
                st.dataframe(
                    fl_r_hoy[["TURNO", "RUTA", "PATIO", "Buses_despacho", "Buses_NRTEC", "Buses_inactivos", "Inactivos_con_tarjeta", "Inactivos_sin_tarjeta"]].rename(columns={
                        "TURNO": "Corte",
                        "Buses_despacho": "Buses en despacho",
                        "Buses_NRTEC": "Buses en línea NRTEC",
                        "Buses_inactivos": "Buses INACTIVOS",
                        "Inactivos_con_tarjeta": "Inactivos c/ tarjeta",
                        "Inactivos_sin_tarjeta": "Inactivos s/ tarjeta"
                    }),
                    use_container_width=True,
                    hide_index=True
                )

        st.markdown("---")

        diario = df_f.groupby("FECHA", as_index=False).agg({
            "Programados completos": "sum",
            "Realizados": "sum",
            "KM PERDIDOS": "sum",
            "INGRESO": "sum",
            "INGRESO_PERDIDO": "sum"
        }).sort_values("FECHA")
        diario["Cumplimiento %"] = (diario["Realizados"] / diario["Programados completos"] * 100).round(2)
        diario["Promedio Móvil 7D"] = diario["Cumplimiento %"].rolling(7, min_periods=1).mean().round(2)

        col_g1, col_g2 = st.columns([1.6, 1])
        with col_g1:
            st.subheader("Tendencia Histórica Diaria")
            if unidad_vista == "Solo Financiero ($ MXN)":
                fig_hist = go.Figure()
                fig_hist.add_trace(go.Bar(x=diario["FECHA"], y=diario["INGRESO"], name="Ingreso Real ($ MXN)", marker_color="#2563eb"))
                fig_hist.add_trace(go.Bar(x=diario["FECHA"], y=diario["INGRESO_PERDIDO"], name="Ingreso Perdido ($ MXN)", marker_color="#ef4444"))
                fig_hist.update_layout(barmode="stack", height=370, margin=dict(l=10, r=10, t=30, b=10), yaxis_title="MXN ($)")
            else:
                fig_hist = go.Figure()
                fig_hist.add_trace(go.Scatter(x=diario["FECHA"], y=diario["Cumplimiento %"], mode="lines+markers", name="% Cumplimiento Diario", line=dict(color="#1d4ed8", width=2.5), marker=dict(size=6)))
                fig_hist.add_trace(go.Scatter(x=diario["FECHA"], y=diario["Promedio Móvil 7D"], mode="lines", name="Promedio Móvil 7 Días", line=dict(color="#0f172a", width=2, dash="dash")))
                fig_hist.add_hline(y=95, line_dash="dot", line_color="#16a34a", annotation_text="Meta Verde (95%)")
                fig_hist.add_hline(y=85, line_dash="dot", line_color="#d97706", annotation_text="Mínimo Amarillo (85%)")
                fig_hist.update_layout(height=370, margin=dict(l=10, r=10, t=30, b=10), yaxis_title="% Cumplimiento", yaxis=dict(range=[60, 102]))
            st.plotly_chart(fig_hist, use_container_width=True)

        with col_g2:
            st.subheader("Comparativa por Patio (EPQ vs. CDM)")
            patio_df = df_f.groupby("PATIO", as_index=False).agg({
                "Programados completos": "sum",
                "Realizados": "sum",
                "KM PERDIDOS": "sum",
                "INGRESO": "sum",
                "INGRESO_PERDIDO": "sum"
            })
            patio_df["Cumplimiento %"] = (patio_df["Realizados"] / patio_df["Programados completos"] * 100).round(2)

            if unidad_vista == "Solo Financiero ($ MXN)":
                fig_patio = px.bar(
                    patio_df, x="PATIO", y=["INGRESO", "INGRESO_PERDIDO"],
                    barmode="group", text_auto="$,.0f",
                    color_discrete_map={"INGRESO": "#16a34a", "INGRESO_PERDIDO": "#dc2626"},
                    labels={"value": "MXN ($)", "variable": "Concepto"}
                )
            else:
                fig_patio = px.bar(
                    patio_df, x="PATIO", y="Cumplimiento %",
                    text="Cumplimiento %", color="PATIO",
                    color_discrete_map={"CDM": "#2563eb", "EPQ": "#7c3aed"}
                )
                fig_patio.update_traces(texttemplate="%{text:.2f}%", textposition="outside")
                fig_patio.add_hline(y=95, line_dash="dot", line_color="#16a34a")
                fig_patio.update_yaxes(range=[0, 108])
            fig_patio.update_layout(height=370, margin=dict(l=10, r=10, t=30, b=10), showlegend=(unidad_vista == "Solo Financiero ($ MXN)"))
            st.plotly_chart(fig_patio, use_container_width=True)

        st.subheader("Causa Raíz de Kilómetros Perdidos (Desglose por Motivo)")
        if not df_perd_f.empty and df_perd_f[MOTIVOS_COLS].sum().sum() > 0:
            motivos_tot = df_perd_f[MOTIVOS_COLS].sum().reset_index()
            motivos_tot.columns = ["Motivo", "Km Perdidos"]
            motivos_tot["Motivo"] = motivos_tot["Motivo"].str.replace("Km ", "", regex=False).str.capitalize()
            motivos_tot = motivos_tot[motivos_tot["Km Perdidos"] > 0].sort_values("Km Perdidos", ascending=True)
            motivos_tot["% del Total"] = (motivos_tot["Km Perdidos"] / motivos_tot["Km Perdidos"].sum() * 100).round(2)
            costo_prom_km_perd = (tot_ing_perd / tot_perd) if tot_perd > 0 else (TARIFA_REGULAR * 0.95)
            motivos_tot["Impacto Est. ($ MXN)"] = motivos_tot["Km Perdidos"] * costo_prom_km_perd

            col_m1, col_m2 = st.columns([1.6, 1])
            with col_m1:
                x_col = "Impacto Est. ($ MXN)" if unidad_vista == "Solo Financiero ($ MXN)" else "Km Perdidos"
                txt_fmt = "$,.0f" if unidad_vista == "Solo Financiero ($ MXN)" else ",.0f"
                fig_mot = px.bar(motivos_tot, x=x_col, y="Motivo", orientation="h", text=x_col, color=x_col, color_continuous_scale="Reds")
                fig_mot.update_traces(texttemplate="%{text:" + txt_fmt + "}", textposition="outside")
                fig_mot.update_layout(height=360, margin=dict(l=10, r=60, t=20, b=10), coloraxis_showscale=False)
                st.plotly_chart(fig_mot, use_container_width=True)

            with col_m2:
                tabla_mot = motivos_tot.sort_values("Km Perdidos", ascending=False).copy()
                st.dataframe(
                    tabla_mot.style.format({
                        "Km Perdidos": "{:,.2f} km",
                        "% del Total": "{:.2f}%",
                        "Impacto Est. ($ MXN)": "${:,.2f}"
                    }),
                    use_container_width=True,
                    hide_index=True,
                    height=350
                )
        else:
            st.info("No hay desglose detallado de motivos registrado para el filtro seleccionado.")

# =========================================================
# TAB 3: MATRIZ POR RUTA Y UNIDADES INACTIVAS
# =========================================================
with tab_rutas:
    st.subheader("Desempeño Acumulado y Financiero por Ruta")
    res_rutas = df_f.groupby(["RUTA", "PATIO"], as_index=False).agg({
        "FECHA": "nunique",
        "Programados completos": "sum",
        "Realizados": "sum",
        "Km penalizado": "sum",
        "KM A COBRAR": "sum",
        "KM PERDIDOS": "sum",
        "INGRESO": "sum",
        "INGRESO_PERDIDO": "sum"
    }).rename(columns={"FECHA": "Días"})

    res_rutas["Cumplimiento %"] = (res_rutas["Realizados"] / res_rutas["Programados completos"] * 100).round(2)
    res_rutas["Semáforo"] = res_rutas["Cumplimiento %"].apply(badge_semaforo)
    res_rutas = res_rutas.sort_values("Cumplimiento %", ascending=False)

    cols_mostrar = [
        "Semáforo", "RUTA", "PATIO", "Días",
        "Programados completos", "Realizados", "KM PERDIDOS",
        "Km penalizado", "KM A COBRAR", "INGRESO", "INGRESO_PERDIDO"
    ]

    st.dataframe(
        res_rutas[cols_mostrar].style.format({
            "Programados completos": "{:,.2f}",
            "Realizados": "{:,.2f}",
            "KM PERDIDOS": "{:,.2f}",
            "Km penalizado": "{:,.2f}",
            "KM A COBRAR": "{:,.2f}",
            "INGRESO": "${:,.2f}",
            "INGRESO_PERDIDO": "${:,.2f}"
        }),
        use_container_width=True,
        hide_index=True,
        height=540
    )

    st.markdown("---")
    st.subheader("Histórico de Despacho y Unidades Inactivas por Ruta")
    if not df_flota_rutas.empty:
        filtro_turno_hist = st.multiselect(
            "Filtrar por corte:",
            options=TURNOS_FLOTA,
            default=TURNOS_FLOTA
        )
        df_fr_show = df_flota_rutas[df_flota_rutas["TURNO"].isin(filtro_turno_hist)].sort_values(
            ["FECHA", "TURNO", "RUTA"], ascending=[False, True, True]
        )
        st.dataframe(
            df_fr_show.rename(columns={
                "TURNO": "Corte / Turno",
                "Buses_despacho": "Buses en despacho",
                "Buses_NRTEC": "Buses en línea NRTEC",
                "Buses_inactivos": "Buses INACTIVOS",
                "Inactivos_con_tarjeta": "Inactivos c/ tarjeta",
                "Inactivos_sin_tarjeta": "Inactivos s/ tarjeta"
            }),
            use_container_width=True,
            hide_index=True
        )
    else:
        st.info("Aún no hay registros de despacho/inactivos por ruta capturados. Regístralos en la pestaña 'Captura Diaria'.")

# =========================================================
# TAB 4: CAPTURA DIARIA EN LA APP (MATUTINO, VESPERTINO Y CIERRE KM)
# =========================================================
with tab_captura:
    st.subheader("Captura Diaria Directa en la App")

    fecha_nueva = st.date_input("Fecha del reporte a capturar", value=date.today())
    fecha_ts_check = pd.to_datetime(fecha_nueva).normalize()

    if fecha_ts_check in df_prod["FECHA"].values:
        st.warning(f"⚠️ Ya existen datos de kilómetros guardados para el **{fecha_nueva.strftime('%d/%m/%Y')}**.")

    # ---------------------------------------------------------
    # BLOQUE A: REPORTE DE FLOTA / RUTAS INACTIVAS (MATUTINO O VESPERTINO)
    # ---------------------------------------------------------
    st.markdown("### Reporte de Flota, Despacho y Rutas Inactivas")
    turno_captura = st.radio(
        "Corte de flota capturando:",
        options=TURNOS_FLOTA,
        horizontal=True
    )

    col_fl1, col_fl2 = st.columns([1, 1.3])

    with col_fl1:
        st.markdown(f"**Resumen por Patio — {turno_captura}**")
        fl_ex = df_flota[(df_flota["FECHA"] == fecha_ts_check) & (df_flota["TURNO"] == turno_captura)] if not df_flota.empty else pd.DataFrame()
        if not fl_ex.empty:
            flota_plantilla = fl_ex[["PATIO", "Buses_despacho", "Buses_NRTEC", "Buses_inactivos", "Inactivos_con_tarjeta", "Inactivos_sin_tarjeta"]].rename(columns={
                "PATIO": "Patio",
                "Buses_despacho": "Buses en despacho",
                "Buses_NRTEC": "Buses en línea NRTEC",
                "Buses_inactivos": "Buses INACTIVOS",
                "Inactivos_con_tarjeta": "Inactivos con tarjeta",
                "Inactivos_sin_tarjeta": "Inactivos sin tarjeta"
            })
        else:
            flota_plantilla = pd.DataFrame([
                {"Patio": "CDM", "Buses en despacho": 32, "Buses en línea NRTEC": 32, "Buses INACTIVOS": 0, "Inactivos con tarjeta": 0, "Inactivos sin tarjeta": 0},
                {"Patio": "EPQ", "Buses en despacho": 61, "Buses en línea NRTEC": 61, "Buses INACTIVOS": 0, "Inactivos con tarjeta": 0, "Inactivos sin tarjeta": 0},
            ])
        df_flota_edit = st.data_editor(
            flota_plantilla,
            key=f"editor_flota_{turno_captura}",
            use_container_width=True,
            hide_index=True,
            column_config={
                "Patio": st.column_config.TextColumn("Patio", disabled=True),
                "Buses en despacho": st.column_config.NumberColumn("Despacho", min_value=0, step=1),
                "Buses en línea NRTEC": st.column_config.NumberColumn("NRTEC", min_value=0, step=1),
                "Buses INACTIVOS": st.column_config.NumberColumn("INACTIVOS", min_value=0, step=1),
                "Inactivos con tarjeta": st.column_config.NumberColumn("Inact. c/ tarjeta", min_value=0, step=1),
                "Inactivos sin tarjeta": st.column_config.NumberColumn("Inact. s/ tarjeta", min_value=0, step=1),
            }
        )

    with col_fl2:
        st.markdown(f"**Detalle por Ruta / Rutas Inactivas — {turno_captura}**")
        fl_r_ex = df_flota_rutas[(df_flota_rutas["FECHA"] == fecha_ts_check) & (df_flota_rutas["TURNO"] == turno_captura)] if not df_flota_rutas.empty else pd.DataFrame()
        if not fl_r_ex.empty:
            flota_rutas_plantilla = fl_r_ex[["RUTA", "PATIO", "Buses_despacho", "Buses_NRTEC", "Buses_inactivos", "Inactivos_con_tarjeta", "Inactivos_sin_tarjeta"]].rename(columns={
                "RUTA": "Ruta",
                "PATIO": "Patio",
                "Buses_despacho": "Buses en despacho",
                "Buses_NRTEC": "Buses en línea NRTEC",
                "Buses_inactivos": "Buses INACTIVOS",
                "Inactivos_con_tarjeta": "Inactivos con tarjeta",
                "Inactivos_sin_tarjeta": "Inactivos sin tarjeta"
            })
        else:
            flota_rutas_plantilla = pd.DataFrame([
                {"Ruta": "C51", "Patio": "CDM", "Buses en despacho": 25, "Buses en línea NRTEC": 25, "Buses INACTIVOS": 0, "Inactivos con tarjeta": 0, "Inactivos sin tarjeta": 0},
                {"Ruta": "C52", "Patio": "EPQ", "Buses en despacho": 21, "Buses en línea NRTEC": 21, "Buses INACTIVOS": 0, "Inactivos con tarjeta": 0, "Inactivos sin tarjeta": 0},
            ])
        df_flota_rutas_edit = st.data_editor(
            flota_rutas_plantilla,
            key=f"editor_flota_rutas_{turno_captura}",
            use_container_width=True,
            hide_index=True,
            num_rows="dynamic",
            column_config={
                "Ruta": st.column_config.TextColumn("Ruta", required=True),
                "Patio": st.column_config.SelectboxColumn("Patio", options=["EPQ", "CDM"], required=True),
                "Buses en despacho": st.column_config.NumberColumn("Despacho", min_value=0, step=1),
                "Buses en línea NRTEC": st.column_config.NumberColumn("NRTEC", min_value=0, step=1),
                "Buses INACTIVOS": st.column_config.NumberColumn("INACTIVOS", min_value=0, step=1),
                "Inactivos con tarjeta": st.column_config.NumberColumn("Inact. c/ tarjeta", min_value=0, step=1),
                "Inactivos sin tarjeta": st.column_config.NumberColumn("Inact. s/ tarjeta", min_value=0, step=1),
            }
        )

    if st.button(f"Guardar Reporte {turno_captura}", use_container_width=True):
        filas_flota_dia = []
        for _, f_row in df_flota_edit.iterrows():
            filas_flota_dia.append({
                "FECHA": fecha_ts_check,
                "TURNO": turno_captura,
                "PATIO": f_row["Patio"],
                "Buses_despacho": int(f_row["Buses en despacho"] or 0),
                "Buses_NRTEC": int(f_row["Buses en línea NRTEC"] or 0),
                "Buses_inactivos": int(f_row["Buses INACTIVOS"] or 0),
                "Inactivos_con_tarjeta": int(f_row["Inactivos con tarjeta"] or 0),
                "Inactivos_sin_tarjeta": int(f_row["Inactivos sin tarjeta"] or 0),
            })
        df_flota_dia = pd.DataFrame(filas_flota_dia)
        if not st.session_state["df_flota"].empty:
            flota_sin_turno = st.session_state["df_flota"][
                ~((st.session_state["df_flota"]["FECHA"] == fecha_ts_check) & (st.session_state["df_flota"]["TURNO"] == turno_captura))
            ]
            st.session_state["df_flota"] = pd.concat([flota_sin_turno, df_flota_dia], ignore_index=True)
        else:
            st.session_state["df_flota"] = df_flota_dia

        filas_flota_r_dia = []
        for _, fr_row in df_flota_rutas_edit.iterrows():
            if pd.notna(fr_row["Ruta"]) and str(fr_row["Ruta"]).strip() != "":
                filas_flota_r_dia.append({
                    "FECHA": fecha_ts_check,
                    "TURNO": turno_captura,
                    "RUTA": str(fr_row["Ruta"]).strip().upper(),
                    "PATIO": str(fr_row["Patio"]).strip().upper(),
                    "Buses_despacho": int(fr_row["Buses en despacho"] or 0),
                    "Buses_NRTEC": int(fr_row["Buses en línea NRTEC"] or 0),
                    "Buses_inactivos": int(fr_row["Buses INACTIVOS"] or 0),
                    "Inactivos_con_tarjeta": int(fr_row["Inactivos con tarjeta"] or 0),
                    "Inactivos_sin_tarjeta": int(fr_row["Inactivos sin tarjeta"] or 0),
                })
        df_flota_r_dia = pd.DataFrame(filas_flota_r_dia)
        if not st.session_state["df_flota_rutas"].empty:
            flota_r_sin_turno = st.session_state["df_flota_rutas"][
                ~((st.session_state["df_flota_rutas"]["FECHA"] == fecha_ts_check) & (st.session_state["df_flota_rutas"]["TURNO"] == turno_captura))
            ]
            st.session_state["df_flota_rutas"] = pd.concat([flota_r_sin_turno, df_flota_r_dia], ignore_index=True)
        else:
            st.session_state["df_flota_rutas"] = df_flota_r_dia

        with open(ARCHIVO_LOCAL_DB, "wb") as f:
            f.write(guardar_consolidado_excel(
                st.session_state["df_prod"],
                st.session_state["df_perd"],
                st.session_state["df_flota"],
                st.session_state["df_flota_rutas"]
            ))
        st.success(f"✅ Reporte **{turno_captura}** del {fecha_nueva.strftime('%d/%m/%Y')} guardado.")
        st.rerun()

    st.markdown("---")

    # ---------------------------------------------------------
    # BLOQUE B: CIERRE DE KILÓMETROS POR RUTA
    # ---------------------------------------------------------
    if fecha_ts_check in df_prod["FECHA"].values:
        df_existente_dia = df_prod[df_prod["FECHA"] == fecha_ts_check]
        filas_plantilla = []
        for _, r_ex in df_existente_dia.iterrows():
            filas_plantilla.append({
                "RUTA": r_ex["RUTA"],
                "PATIO": r_ex["PATIO"],
                "Km programado": float(r_ex["Programados completos"]),
                "Km realizados": float(r_ex["Realizados"]),
                "Km penalizado (C51/C52)": float(r_ex["Km penalizado"])
            })
    else:
        ultimos_prog = df_prod.sort_values("FECHA").groupby("RUTA")["Programados completos"].last().to_dict()
        filas_plantilla = []
        for r in RUTAS_ORDEN:
            prog_def = float(ultimos_prog.get(r, 0.0))
            filas_plantilla.append({
                "RUTA": r,
                "PATIO": MAPA_RUTA_PATIO.get(r, "EPQ"),
                "Km programado": prog_def,
                "Km realizados": prog_def,
                "Km penalizado (C51/C52)": 0.0
            })
    df_plantilla = pd.DataFrame(filas_plantilla)

    st.markdown("### Cierre de Kilómetros por Ruta")
    df_editado = st.data_editor(
        df_plantilla,
        use_container_width=True,
        hide_index=True,
        num_rows="dynamic",
        column_config={
            "RUTA": st.column_config.TextColumn("Ruta", required=True),
            "PATIO": st.column_config.SelectboxColumn("Patio", options=["EPQ", "CDM"], required=True),
            "Km programado": st.column_config.NumberColumn("Km programado", min_value=0.0, format="%.2f"),
            "Km realizados": st.column_config.NumberColumn("Km realizados", min_value=0.0, format="%.2f"),
            "Km penalizado (C51/C52)": st.column_config.NumberColumn("Km penalizado (opcional)", min_value=0.0, format="%.2f"),
        }
    )

    preview_rows = []
    for _, row in df_editado.iterrows():
        r_cod = str(row["RUTA"]).strip().upper()
        p_cod = str(row["PATIO"]).strip().upper()
        prog_v = float(row["Km programado"] or 0.0)
        real_v = float(row["Km realizados"] or 0.0)
        pen_v = float(row["Km penalizado (C51/C52)"] or 0.0)
        preview_rows.append({
            "FECHA": fecha_ts_check,
            "PATIO": p_cod,
            "RUTA": r_cod,
            "Programados completos": prog_v,
            "Realizados": real_v,
            "Km penalizado": pen_v,
            "KM PERDIDOS": max(0.0, prog_v - real_v),
            "ES_ESTIMADO": (pen_v == 0.0 and r_cod in ["C51", "C52"])
        })
    df_preview = calcular_metricas_fila(pd.DataFrame(preview_rows))

    p_prog = df_preview["Programados completos"].sum()
    p_real = df_preview["Realizados"].sum()
    p_cumpl = (p_real / p_prog * 100) if p_prog > 0 else 0.0
    p_ing = df_preview["INGRESO"].sum()
    p_perd_mxn = df_preview["INGRESO_PERDIDO"].sum()

    st.info(
        f"**Vista Previa del Cierre ({fecha_nueva.strftime('%d/%m/%Y')}):** "
        f"Cumplimiento **{p_cumpl:.2f}%** ({p_real:,.1f} / {p_prog:,.1f} km) | "
        f"Cobro del día: **${p_ing:,.2f} MXN** | "
        f"Dejado de percibir: **${p_perd_mxn:,.2f} MXN**"
    )

    st.markdown("### Kilómetros Perdidos por Motivo y Patio")
    motivos_plantilla = pd.DataFrame({
        "Motivo": [
            "Km operaciones", "Km mantenimiento", "Km carga de gas", "Km tráfico",
            "Km conductor", "Km accidentes", "Km validador", "Km recurso bus",
            "Km recurso conductor"
        ],
        "EPQ": [0.0] * 9,
        "CDM": [0.0] * 9
    })
    df_mot_edit = st.data_editor(
        motivos_plantilla,
        use_container_width=True,
        hide_index=True,
        column_config={
            "Motivo": st.column_config.TextColumn("Motivo", disabled=True),
            "EPQ": st.column_config.NumberColumn("EPQ (km)", min_value=0.0, format="%.1f"),
            "CDM": st.column_config.NumberColumn("CDM (km)", min_value=0.0, format="%.1f"),
        }
    )

    if st.button("Guardar Todo el Día Completo", type="primary", use_container_width=True):
        # 1. Producción y cobro
        prod_sin_fecha = st.session_state["df_prod"][st.session_state["df_prod"]["FECHA"] != fecha_ts_check]
        st.session_state["df_prod"] = pd.concat([prod_sin_fecha, df_preview], ignore_index=True)

        # 2. Desglose de perdidos
        filas_perd_dia = []
        for p_cod in ["EPQ", "CDM"]:
            sub_r = df_preview[df_preview["PATIO"] == p_cod]
            tot_perd_patio = sub_r["KM PERDIDOS"].sum()
            for _, r_row in sub_r.iterrows():
                peso = (r_row["KM PERDIDOS"] / tot_perd_patio) if tot_perd_patio > 0 else 0.0
                reg_perd = {"FECHA": fecha_ts_check, "PATIO": p_cod, "RUTA": r_row["RUTA"]}
                for _, m_row in df_mot_edit.iterrows():
                    reg_perd[m_row["Motivo"]] = round(float(m_row[p_cod] or 0.0) * peso, 2)
                filas_perd_dia.append(reg_perd)
        df_perd_nuevo_dia = pd.DataFrame(filas_perd_dia)

        perd_sin_fecha = st.session_state["df_perd"][st.session_state["df_perd"]["FECHA"] != fecha_ts_check]
        st.session_state["df_perd"] = pd.concat([perd_sin_fecha, df_perd_nuevo_dia], ignore_index=True)

        # 3. Flota por patio (turno activo)
        filas_flota_dia = []
        for _, f_row in df_flota_edit.iterrows():
            filas_flota_dia.append({
                "FECHA": fecha_ts_check,
                "TURNO": turno_captura,
                "PATIO": f_row["Patio"],
                "Buses_despacho": int(f_row["Buses en despacho"] or 0),
                "Buses_NRTEC": int(f_row["Buses en línea NRTEC"] or 0),
                "Buses_inactivos": int(f_row["Buses INACTIVOS"] or 0),
                "Inactivos_con_tarjeta": int(f_row["Inactivos con tarjeta"] or 0),
                "Inactivos_sin_tarjeta": int(f_row["Inactivos sin tarjeta"] or 0),
            })
        df_flota_dia = pd.DataFrame(filas_flota_dia)
        if not st.session_state["df_flota"].empty:
            flota_sin_turno = st.session_state["df_flota"][
                ~((st.session_state["df_flota"]["FECHA"] == fecha_ts_check) & (st.session_state["df_flota"]["TURNO"] == turno_captura))
            ]
            st.session_state["df_flota"] = pd.concat([flota_sin_turno, df_flota_dia], ignore_index=True)
        else:
            st.session_state["df_flota"] = df_flota_dia

        # 4. Flota y rutas inactivas por ruta (turno activo)
        filas_flota_r_dia = []
        for _, fr_row in df_flota_rutas_edit.iterrows():
            if pd.notna(fr_row["Ruta"]) and str(fr_row["Ruta"]).strip() != "":
                filas_flota_r_dia.append({
                    "FECHA": fecha_ts_check,
                    "TURNO": turno_captura,
                    "RUTA": str(fr_row["Ruta"]).strip().upper(),
                    "PATIO": str(fr_row["Patio"]).strip().upper(),
                    "Buses_despacho": int(fr_row["Buses en despacho"] or 0),
                    "Buses_NRTEC": int(fr_row["Buses en línea NRTEC"] or 0),
                    "Buses_inactivos": int(fr_row["Buses INACTIVOS"] or 0),
                    "Inactivos_con_tarjeta": int(fr_row["Inactivos con tarjeta"] or 0),
                    "Inactivos_sin_tarjeta": int(fr_row["Inactivos sin tarjeta"] or 0),
                })
        df_flota_r_dia = pd.DataFrame(filas_flota_r_dia)
        if not st.session_state["df_flota_rutas"].empty:
            flota_r_sin_turno = st.session_state["df_flota_rutas"][
                ~((st.session_state["df_flota_rutas"]["FECHA"] == fecha_ts_check) & (st.session_state["df_flota_rutas"]["TURNO"] == turno_captura))
            ]
            st.session_state["df_flota_rutas"] = pd.concat([flota_r_sin_turno, df_flota_r_dia], ignore_index=True)
        else:
            st.session_state["df_flota_rutas"] = df_flota_r_dia

        with open(ARCHIVO_LOCAL_DB, "wb") as f:
            f.write(guardar_consolidado_excel(
                st.session_state["df_prod"],
                st.session_state["df_perd"],
                st.session_state["df_flota"],
                st.session_state["df_flota_rutas"]
            ))

        st.success(f"✅ ¡Listo! Se guardó el reporte completo del {fecha_nueva.strftime('%d/%m/%Y')} ({turno_captura}, Kilómetros y Cobro Semanal).")
        st.rerun()