import streamlit as st
import pandas as pd
import numpy as np
import re
import math
import unicodedata
from io import BytesIO
from datetime import date
import plotly.graph_objects as go

# ==============================================================================
# CONFIGURACIÓN
# ==============================================================================
st.set_page_config(page_title="Auditoría Master CONTPAQ", layout="wide", page_icon="🛡️")
UMBRAL_TOLERANCIA = 1.0 

st.title("🛡️ Auditoría Master de Saldos (Contpaq)")
st.markdown("""
Esta herramienta combina:
1. **Lectura Blindada:** Soporta Excel/CSV y detecta errores en el archivo.
2. **Lógica de Negocio:** Detecta cruces de cuentas, facturas pendientes y saldos iniciales.
""")

# ==============================================================================
# 1. UTILIDADES DE LIMPIEZA Y NORMALIZACIÓN
# ==============================================================================

@st.cache_data
def parse_spanish_date(s: str):
    if pd.isna(s): return pd.NaT
    s = str(s).strip()
    m = re.match(r"^(\d{1,2})[/\-]([A-Za-z]{3})[/\-](\d{4})$", s)
    if not m: return pd.NaT
    day, mon_abbr, year = m.groups()
    month_map = {"ene":"01","feb":"02","mar":"03","abr":"04","may":"05","jun":"06",
                 "jul":"07","ago":"08","sep":"09","oct":"10","nov":"11","dic":"12"}
    mon_key = mon_abbr[:3].lower()
    if mon_key not in month_map: return pd.NaT
    return pd.to_datetime(f"{day.zfill(2)}/{month_map[mon_key]}/{year}", format="%d/%m/%Y", errors="coerce")

def normalizar_referencia_base(ref):
    """Limpieza agresiva de basura en referencias."""
    if pd.isna(ref): return None
    if isinstance(ref, float) and ref.is_integer(): s = str(int(ref))
    else: s = str(ref).strip()
    
    s = "".join(c for c in unicodedata.normalize("NFKD", s) if not unicodedata.combining(c)).upper()
    s = re.sub(r"\.0+$", "", s)
    s = re.sub(r"^(FACTURA|FAC|FOLIO|REF|REFERENCIA|RECIBO|DEPOSITO|DEP|PAGO|ABONO|NC|NOTA)\s*[:.\-]?\s*", "", s)
    s = re.sub(r"^F\s*[-:]?\s*", "", s)
    s = re.sub(r"[ \-_/]", "", s)
    return s if s else None

def aplicar_mapeo_inteligente(df):
    """Recupera la lógica del código antiguo para unir '123' con 'NCTA123'."""
    df["ref_base"] = df["referencia"].apply(normalizar_referencia_base)
    
    # Detectar números huérfanos y buscarles papá
    uniques = pd.Series(df["ref_base"].dropna().unique(), dtype=object)
    set_uniques = set(uniques)
    numeros = [u for u in uniques if isinstance(u, str) and re.fullmatch(r"\d+", u)]
    prefijos = ["NCTA", "ANCT", "PNCT", "NC", "F"]
    
    mapa = {}
    for num in numeros:
        for pre in prefijos:
            candidate = pre + num
            if candidate in set_uniques:
                mapa[num] = candidate
                break
    
    df["referencia_norm"] = df["ref_base"].apply(lambda x: mapa.get(x, x))
    return df

def cargar_archivo_robusto(uploaded_file):
    try:
        return pd.read_excel(uploaded_file, header=None)
    except:
        uploaded_file.seek(0)
        try: return pd.read_csv(uploaded_file, header=None, encoding='latin-1')
        except: return pd.read_csv(uploaded_file, header=None, encoding='utf-8', errors='replace')

def to_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False)
    return output.getvalue()

# ==============================================================================
# 2. PROCESAMIENTO CENTRAL (ENGINE BLINDADO)
# ==============================================================================

@st.cache_data
def procesar_contpaq_engine(file):
    raw = cargar_archivo_robusto(file)
    raw_str = raw.astype(str)
    
    # Detección de filas
    patron_cuenta = r"^\d{3}-\d{3}-\d{3}-\d{3}"
    mask_cuenta = raw_str[0].str.match(patron_cuenta, na=False)
    # Buscar "Saldo inicial" en cualquier columna de la fila
    mask_saldo = raw_str.apply(lambda r: r.str.contains("Saldo inicial", case=False, na=False).any(), axis=1)
    is_header = mask_cuenta & mask_saldo
    
    # Extracción de Datos Maestros
    df = raw.copy()
    df["meta_codigo"] = np.where(is_header, df[0], np.nan)
    df["meta_nombre"] = np.where(is_header, df[1], np.nan)
    
    # Saldo Inicial Real (Buscamos el número en la misma fila)
    # A veces está en col 7, a veces 6. Tomamos el último valor numérico de la fila header.
    def get_saldo_ini(row):
        if not pd.isna(row[7]): return row[7] # Prioridad col 7
        return 0.0
    
    df["meta_saldo_inicial"] = np.where(is_header, pd.to_numeric(df[7], errors='coerce'), np.nan)
    
    df["meta_codigo"] = df["meta_codigo"].ffill()
    df["meta_nombre"] = df["meta_nombre"].ffill()
    df["meta_saldo_inicial"] = df["meta_saldo_inicial"].ffill()
    
    # Movimientos
    patron_fecha = r"^\d{1,2}/[A-Za-z]{3}/\d{4}"
    is_mov = raw_str[0].str.match(patron_fecha, na=False)
    movs = df[is_mov].copy()
    
    col_map = {0:"fecha_raw", 1:"tipo", 2:"poliza", 3:"concepto", 4:"referencia", 5:"cargos", 6:"abonos", 7:"saldo_acumulado"}
    movs = movs.rename(columns=col_map)
    
    for c in ["cargos", "abonos", "saldo_acumulado"]:
        movs[c] = pd.to_numeric(movs[c], errors='coerce').fillna(0)
        
    movs["fecha"] = movs["fecha_raw"].apply(parse_spanish_date)
    
    # Aplicar Mapeo Inteligente (Lógica Vieja en Motor Nuevo)
    movs = aplicar_mapeo_inteligente(movs)
    
    # Totales Auxiliar (Bottom-Up)
    mask_total = raw_str[4].str.strip() == "Total:"
    totales = df[mask_total].copy()
    totales["saldo_final_aux"] = pd.to_numeric(totales[7], errors='coerce').fillna(0)
    
    resumen = totales.groupby(["meta_codigo", "meta_nombre"])["saldo_final_aux"].sum().reset_index()
    
    # Pegar Saldo Inicial al resumen
    saldos_ini = df[is_header][["meta_codigo", "meta_saldo_inicial"]].drop_duplicates()
    resumen = resumen.merge(saldos_ini, on="meta_codigo", how="left")
    
    return movs, resumen

# ==============================================================================
# 3. LÓGICA DE NEGOCIO (EL VALOR AGREGADO)
# ==============================================================================

@st.cache_data
def detectar_cruces(movs):
    """
    Detecta facturas que tocan múltiples cuentas (Ej: Cargo en Cta A, Abono en Cta B).
    Esta es la función 'estrella' del código viejo.
    """
    validos = movs[movs["referencia_norm"].notna()]
    
    # Agrupar por Referencia y Cuenta
    por_cuenta = validos.groupby(["referencia_norm", "meta_codigo", "meta_nombre"]).agg(
        cargos=("cargos", "sum"),
        abonos=("abonos", "sum")
    ).reset_index()
    
    por_cuenta["tiene_cargo"] = por_cuenta["cargos"] > 0
    por_cuenta["tiene_abono"] = por_cuenta["abonos"] > 0
    
    # Analizar a nivel Referencia
    nivel_ref = por_cuenta.groupby("referencia_norm").agg(
        num_cuentas=("meta_codigo", "nunique"),
        hay_cargo=("tiene_cargo", "max"),
        hay_abono=("tiene_abono", "max")
    ).reset_index()
    
    # Filtro: Más de 1 cuenta Y (tiene cargo Y tiene abono distribuidos)
    cruces = nivel_ref[ (nivel_ref["num_cuentas"] > 1) & nivel_ref["hay_cargo"] & nivel_ref["hay_abono"] ]
    
    if cruces.empty:
        return pd.DataFrame()
    
    # Traer detalle
    detalle = por_cuenta[por_cuenta["referencia_norm"].isin(cruces["referencia_norm"])].copy()
    detalle["saldo_en_cuenta"] = detalle["cargos"] - detalle["abonos"]
    return detalle.sort_values("referencia_norm")

def analizar_saldos(movs, resumen):
    """Construye la tabla maestra de auditoría (Semáforo)."""
    # 1. Saldo Facturas (Top-Down)
    vivas = movs[movs["referencia_norm"].notna()]
    saldo_facturas = vivas.groupby(["meta_codigo"]).apply(lambda x: (x["cargos"] - x["abonos"]).sum()).reset_index(name="saldo_facturas")
    
    # 2. Merge con Auxiliar
    df = resumen.merge(saldo_facturas, on="meta_codigo", how="left").fillna(0)
    df["diferencia"] = df["saldo_final_aux"] - df["saldo_facturas"]
    
    # 3. Clasificación (Semáforo)
    def clasificar(row):
        dif = abs(row["diferencia"])
        fin = abs(row["saldo_final_aux"])
        fact = abs(row["saldo_facturas"])
        
        if dif <= UMBRAL_TOLERANCIA: return "🟢 OK"
        if fact <= UMBRAL_TOLERANCIA and dif > UMBRAL_TOLERANCIA: return "🟡 Solo Saldo Inicial / Ajuste"
        return "🔴 Diferencia No Explicada"
        
    df["estado"] = df.apply(clasificar, axis=1)
    return df

# ==============================================================================
# APP UI
# ==============================================================================

uploaded_file = st.file_uploader("📂 Sube reporte CONTPAQ (Excel o CSV)", type=["xlsx", "csv"])

if uploaded_file:
    with st.spinner("🚀 Procesando archivo híbrido..."):
        try:
            movs, resumen = procesar_contpaq_engine(uploaded_file)
            
            # Ejecutar lógicas avanzadas
            df_audit = analizar_saldos(movs, resumen)
            df_cruces = detectar_cruces(movs)
            
            # Facturas Pendientes (Detalle)
            movs_validos = movs[movs["referencia_norm"].notna()]
            facturas_pend = movs_validos.groupby(["meta_codigo", "meta_nombre", "referencia_norm"]).agg(
                fecha=("fecha", "min"),
                saldo=("cargos", lambda x: x.sum() - movs_validos.loc[x.index, "abonos"].sum())
            ).reset_index()
            facturas_pend = facturas_pend[facturas_pend["saldo"].abs() > 0.01] # Quitar ceros
            
        except Exception as e:
            st.error(f"Error procesando: {e}")
            st.stop()
            
    # KPIs Globales
    st.divider()
    col1, col2, col3, col4 = st.columns(4)
    saldo_total = df_audit["saldo_final_aux"].sum()
    diferencia_total = df_audit["diferencia"].sum()
    
    col1.metric("Saldo Contable Total", f"${saldo_total:,.2f}")
    col2.metric("Diferencia sin Soporte", f"${diferencia_total:,.2f}", delta_color="inverse")
    col3.metric("Facturas con Cruce", len(df_cruces["referencia_norm"].unique()) if not df_cruces.empty else 0)
    col4.metric("Cuentas con Error", len(df_audit[df_audit["estado"].str.contains("🔴")]))

    # Pestañas
    t1, t2, t3, t4 = st.tabs(["🚦 Semáforo", "📑 Facturas Pendientes", "🔀 Cruces de Cuentas", "📉 Gráficos"])
    
    with t1:
        st.subheader("Conciliación por Cuenta")
        ver_todo = st.toggle("Ver solo cuentas con problemas", value=True)
        df_show = df_audit[df_audit["estado"] != "🟢 OK"] if ver_todo else df_audit
        
        st.dataframe(
            df_show[["meta_codigo", "meta_nombre", "estado", "saldo_final_aux", "saldo_facturas", "diferencia"]],
            use_container_width=True,
            column_config={
                "saldo_final_aux": st.column_config.NumberColumn("Saldo Auxiliar", format="$%.2f"),
                "saldo_facturas": st.column_config.NumberColumn("Suma Facturas", format="$%.2f"),
                "diferencia": st.column_config.NumberColumn("Diferencia", format="$%.2f"),
            }
        )
        st.download_button("Descargar Semáforo", to_excel(df_audit), "semaforo.xlsx")
        
    with t2:
        st.subheader("Detalle de Facturas Vivas")
        st.dataframe(
            facturas_pend.sort_values("fecha"),
            use_container_width=True,
            column_config={"saldo": st.column_config.NumberColumn("Saldo Pendiente", format="$%.2f")}
        )
        
    with t3:
        st.subheader("Referencias Cruzadas (Error común de aplicación de pagos)")
        if df_cruces.empty:
            st.success("✅ No se detectaron cruces de referencias entre cuentas.")
        else:
            st.warning("⚠️ Estas facturas tienen cargos en una cuenta y abonos en otra distinta.")
            st.dataframe(df_cruces)
            
    with t4:
        # Gráfico de composición
        fig = go.Figure(data=[
            go.Bar(name='Facturas OK', x=['Total'], y=[saldo_total - diferencia_total], marker_color='#2ecc71'),
            go.Bar(name='Diferencias', x=['Total'], y=[diferencia_total], marker_color='#e74c3c')
        ])
        fig.update_layout(barmode='stack', title="Calidad del Saldo")
        st.plotly_chart(fig, use_container_width=True)

else:
    st.info("Esperando archivo...")
