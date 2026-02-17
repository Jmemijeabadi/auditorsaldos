import streamlit as st
import pandas as pd
import numpy as np
import re
import math
import unicodedata
from io import BytesIO
from datetime import date
import plotly.graph_objects as go

# --------------------------------------------------------------------
# Configuración inicial de la página
# --------------------------------------------------------------------
st.set_page_config(page_title="Auditoría de Saldos", layout="wide", page_icon="🔍")

# Umbral para considerar "casi cero" el saldo neto
UMBRAL_SALDO_INICIAL = 1.0 

st.title("🔍 Tablero de Auditoría de Saldos (CONTPAQ i)")
st.markdown("""
Esta herramienta concilia automáticamente el reporte de **Movimientos Auxiliares** contra el detalle de facturas.
**Objetivo:** Detectar diferencias entre el saldo contable y el saldo vivo de facturas.
""")

# --------------------------------------------------------------------
# Cuadro de ayuda (Simplificado)
# --------------------------------------------------------------------
with st.expander("ℹ️ Ayuda rápida: ¿Qué significan los colores?", expanded=False):
    st.markdown("""
    - 🔴 **Revisar (Diferencia):** El saldo final de la cuenta contable NO coincide con la suma de las facturas pendientes. Requiere ajuste manual.
    - 🟡 **Solo Saldo Inicial:** La cuenta tiene saldo, pero no hay facturas vivas en este periodo que lo expliquen (probablemente saldo arrastrado de años anteriores).
    - 🟢 **Conciliado:** El saldo contable coincide perfectamente con las facturas.
    - ⚪ **Saldada:** Cuenta en ceros.
    """)

# --------------------------------------------------------------------
# FUNCIONES DE UTILIDAD Y PROCESAMIENTO
# --------------------------------------------------------------------

@st.cache_data
def parse_spanish_date(s: str):
    """Convierte fechas tipo '02/Ene/2025' a datetime."""
    if pd.isna(s):
        return pd.NaT
    s = str(s).strip()
    m = re.match(r"^(\d{2})/([A-Za-z]{3})/(\d{4})$", s)
    if not m:
        return pd.NaT
    day, mon_abbr, year = m.groups()
    month_map = {
        "Ene": "01", "Feb": "02", "Mar": "03", "Abr": "04",
        "May": "05", "Jun": "06", "Jul": "07", "Ago": "08",
        "Sep": "09", "Oct": "10", "Nov": "11", "Dic": "12",
    }
    mon_key = mon_abbr[:3].title()
    if mon_key not in month_map:
        return pd.NaT
    date_str = f"{day}/{month_map[mon_key]}/{year}"
    return pd.to_datetime(date_str, format="%d/%m/%Y", errors="coerce")


def _strip_accents(s: str) -> str:
    """Quita acentos para normalizar texto."""
    return "".join(
        c for c in unicodedata.normalize("NFKD", s) if not unicodedata.combining(c)
    )


def normalizar_referencia_base(ref):
    """Normaliza referencias para cruzar datos (quita prefijos, espacios, etc)."""
    if pd.isna(ref):
        return None
    if isinstance(ref, (int, float)) and not math.isnan(ref):
        s = str(int(ref))
    else:
        s = str(ref).strip()

    s = _strip_accents(s).upper()
    if re.fullmatch(r"\d+(\.0+)?", s):
        s = re.sub(r"\.0+$", "", s)

    s = re.sub(r"^(FACTURA|FAC|FOLIO|REF|REFERENCIA|RECIBO|DEPOSITO|DEPOS|PAGO|ABONO)\s*[:\-]?\s*", "", s)
    s = re.sub(r"^F\s*[-:]?\s*", "", s)
    s = " ".join(s.split())
    if not s:
        return None
    s = re.sub(r"[ \-_/]", "", s)
    s = re.sub(r"\.0+$", "", s)
    s = re.sub(r"\.$", "", s)
    return s or None


def aplicar_normalizacion_referencias(movs_valid: pd.DataFrame) -> pd.DataFrame:
    """Aplica normalización y mapeo de números huérfanos a referencias con prefijo."""
    df = movs_valid.copy()
    df["referencia_norm_base"] = df["referencia"].apply(normalizar_referencia_base)
    uniques = pd.Series(df["referencia_norm_base"].dropna().unique(), dtype=object)
    set_uniques = set(uniques)
    numeros = [u for u in uniques if isinstance(u, str) and re.fullmatch(r"\d+", u)]
    pref_list = ["NCTA", "ANCT", "PNCT", "NC"]

    mapa_num_a_pref = {}
    for num in numeros:
        for pref in pref_list:
            cand = pref + num
            if cand in set_uniques:
                mapa_num_a_pref[num] = cand
                break

    def _map_final(x):
        if pd.isna(x):
            return None
        return mapa_num_a_pref.get(x, x)

    df["referencia_norm"] = df["referencia_norm_base"].apply(_map_final)
    return df


@st.cache_data
def procesar_movimientos(file):
    """Lee el Excel y extrae movimientos y saldos del auxiliar."""
    raw = pd.read_excel(file, header=None)
    
    # Detectar filas cabecera de cuenta
    account_pattern = re.compile(r"^\d{3}-\d{3}-\d{3}-\d{3}$")
    is_account_header = raw[0].astype(str).str.match(account_pattern) & raw[6].astype(str).str.contains("Saldo inicial", na=False)

    df = raw.copy()
    df["account_code"] = np.where(is_account_header, df[0], np.nan)
    df["account_name"] = np.where(is_account_header, df[1], np.nan)
    df["account_code"] = df["account_code"].ffill()
    df["account_name"] = df["account_name"].ffill()

    # Filas de movimientos
    date_pattern = re.compile(r"^\d{2}/[A-Za-z]{3}/\d{4}$")
    is_date_row = df[0].astype(str).str.match(date_pattern)

    movs = df.loc[is_date_row].copy()
    movs = movs.rename(columns={0: "fecha_raw", 1: "tipo", 2: "numero_poliza", 3: "concepto", 4: "referencia", 5: "cargos", 6: "abonos", 7: "saldo"})

    for col in ["cargos", "abonos", "saldo"]:
        movs[col] = pd.to_numeric(movs[col], errors="coerce")

    movs["referencia"] = movs["referencia"].astype(str).str.strip().replace({"nan": np.nan, "": np.nan})
    movs["fecha"] = movs["fecha_raw"].apply(parse_spanish_date)
    
    movs_valid = movs[movs["referencia"].notna()].copy()
    movs_valid = aplicar_normalizacion_referencias(movs_valid)

    # Totales globales del auxiliar (Fila Total Clientes)
    total_cargos_aux = np.nan
    total_abonos_aux = np.nan
    saldo_final_aux = np.nan

    mask_total_clientes = raw[0].astype(str).str.strip() == "Total Clientes :"
    total_rows = raw.loc[mask_total_clientes]
    if not total_rows.empty:
        row = total_rows.iloc[0]
        total_cargos_aux = pd.to_numeric(row[1], errors="coerce")
        total_abonos_aux = pd.to_numeric(row[2], errors="coerce")
        saldo_final_aux = pd.to_numeric(row[3], errors="coerce")

    total_cargos_movs = movs_valid["cargos"].sum()
    total_abonos_movs = movs_valid["abonos"].sum()
    saldo_neto_movs = total_cargos_movs - total_abonos_movs

    saldo_inicial_implicito = np.nan
    if not math.isnan(saldo_final_aux) and not math.isnan(saldo_neto_movs):
        saldo_inicial_implicito = saldo_final_aux - saldo_neto_movs

    resumen_auxiliar = {
        "saldo_neto_movs": float(saldo_neto_movs),
        "saldo_final_aux": float(saldo_final_aux) if not math.isnan(saldo_final_aux) else None,
        "saldo_inicial_implicito": float(saldo_inicial_implicito) if not math.isnan(saldo_inicial_implicito) else None,
    }

    # Totales por cuenta (Fila Total:)
    tot_rows = df[df[4].astype(str).str.strip() == "Total:"].copy()
    for c in [5, 6, 7]:
        tot_rows[c] = pd.to_numeric(tot_rows[c], errors="coerce")

    totales_cuentas_aux = (
        tot_rows.rename(columns={5: "cargos_total_cuenta_aux", 6: "abonos_total_cuenta_aux", 7: "saldo_final_cuenta_aux"})
        [["account_code", "account_name", "cargos_total_cuenta_aux", "abonos_total_cuenta_aux", "saldo_final_cuenta_aux"]]
        .groupby(["account_code", "account_name"]).agg(
            cargos_total_cuenta_aux=("cargos_total_cuenta_aux", "sum"),
            abonos_total_cuenta_aux=("abonos_total_cuenta_aux", "sum"),
            saldo_final_cuenta_aux=("saldo_final_cuenta_aux", "sum"),
        ).reset_index()
    )

    return movs_valid, resumen_auxiliar, totales_cuentas_aux


@st.cache_data
def construir_facturas_global(movs_valid: pd.DataFrame) -> pd.DataFrame:
    """Agrupa movimientos por referencia para obtener el saldo vivo de cada factura."""
    facturas = (
        movs_valid.groupby("referencia_norm")
        .agg(fecha_factura=("fecha", "min"), cargos_total=("cargos", "sum"), abonos_total=("abonos", "sum"))
        .reset_index()
    )
    movs_valid2 = movs_valid.copy()
    movs_valid2["es_cargo_pos"] = movs_valid2["cargos"] > 0
    
    # Determinar cuenta principal
    main_from_cargo = (
        movs_valid2[movs_valid2["es_cargo_pos"]]
        .sort_values(["referencia_norm", "cargos"], ascending=[True, False])
        .drop_duplicates("referencia_norm")[["referencia_norm", "account_code", "account_name"]]
    )
    main_any = (
        movs_valid2.sort_values(["referencia_norm", "fecha"])
        .drop_duplicates("referencia_norm")[["referencia_norm", "account_code", "account_name"]]
    )
    main_account = pd.concat([main_from_cargo, main_any], ignore_index=True)
    main_account = main_account.drop_duplicates("referencia_norm", keep="first")
    
    facturas = facturas.merge(main_account, on="referencia_norm", how="left")
    facturas["saldo_factura"] = facturas["cargos_total"] - facturas["abonos_total"]
    facturas["cuenta"] = facturas["account_code"].astype(str) + " - " + facturas["account_name"].astype(str)
    facturas = facturas.rename(columns={"referencia_norm": "referencia"})
    return facturas


@st.cache_data
def construir_facturas_por_cuenta(movs_valid: pd.DataFrame) -> pd.DataFrame:
    """Agrupa movimientos por referencia Y cuenta contable."""
    facturas = (
        movs_valid.groupby(["account_code", "account_name", "referencia_norm"])
        .agg(fecha_factura=("fecha", "min"), cargos_total=("cargos", "sum"), abonos_total=("abonos", "sum"))
        .reset_index()
    )
    facturas["saldo_factura"] = facturas["cargos_total"] - facturas["abonos_total"]
    facturas["cuenta"] = facturas["account_code"].astype(str) + " - " + facturas["account_name"].astype(str)
    facturas = facturas.rename(columns={"referencia_norm": "referencia"})
    return facturas


@st.cache_data
def detectar_cruces_referencias(movs_valid: pd.DataFrame):
    """Detecta si una referencia tiene cargos en una cuenta y abonos en otra."""
    df = movs_valid.copy()
    por_cuenta = (
        df.groupby(["referencia_norm", "account_code", "account_name"])
        .agg(cargos_total=("cargos", "sum"), abonos_total=("abonos", "sum"))
        .reset_index()
    )
    por_cuenta["tiene_cargo"] = por_cuenta["cargos_total"] > 0
    por_cuenta["tiene_abono"] = por_cuenta["abonos_total"] > 0

    ref_level = (
        por_cuenta.groupby("referencia_norm")
        .agg(num_cuentas=("account_code", "nunique"), cuentas_con_cargo=("tiene_cargo", "sum"), cuentas_con_abono=("tiene_abono", "sum"),
             cargos_tot_ref=("cargos_total", "sum"), abonos_tot_ref=("abonos_total", "sum"))
        .reset_index()
    )
    ref_level["saldo_neto_ref"] = ref_level["cargos_tot_ref"] - ref_level["abonos_tot_ref"]
    ref_level["es_cruce"] = (ref_level["num_cuentas"] > 1) & (ref_level["cuentas_con_cargo"] >= 1) & (ref_level["cuentas_con_abono"] >= 1)

    resumen_cruces = ref_level[ref_level["es_cruce"]].copy()
    if resumen_cruces.empty:
        return por_cuenta.head(0), resumen_cruces

    detalle_cruces = por_cuenta.merge(
        resumen_cruces[["referencia_norm", "num_cuentas", "cargos_tot_ref", "abonos_tot_ref", "saldo_neto_ref"]],
        on="referencia_norm", how="inner"
    )
    detalle_cruces["saldo_por_cuenta"] = detalle_cruces["cargos_total"] - detalle_cruces["abonos_total"]
    detalle_cruces = detalle_cruces.rename(columns={"referencia_norm": "referencia"})
    resumen_cruces = resumen_cruces.rename(columns={"referencia_norm": "referencia"})
    return detalle_cruces, resumen_cruces


def filtrar_por_fecha(df: pd.DataFrame, fecha_desde: date, fecha_hasta: date) -> pd.DataFrame:
    if df.empty:
        return df.copy()
    mask = pd.Series(True, index=df.index)
    if fecha_desde:
        mask &= df["fecha_factura"] >= pd.to_datetime(fecha_desde)
    if fecha_hasta:
        mask &= df["fecha_factura"] <= pd.to_datetime(fecha_hasta)
    return df[mask].copy()


def to_excel(df: pd.DataFrame) -> bytes:
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Reporte")
    return output.getvalue()


# --------------------------------------------------------------------
# APLICACIÓN PRINCIPAL (MAIN)
# --------------------------------------------------------------------

uploaded_file = st.file_uploader(
    "📎 Sube el archivo Excel de movimientos (auxiliares del catálogo)",
    type=["xlsx"],
)

if uploaded_file is None:
    st.info("Sube un archivo `.xlsx` exportado desde CONTPAQ (Movimientos, Auxiliares del Catálogo) para comenzar.")

else:
    # 1. Procesamiento de datos
    with st.spinner("Procesando datos y detectando anomalías..."):
        movs_valid, resumen_aux, totales_cuentas_aux = procesar_movimientos(uploaded_file)
        facturas_global = construir_facturas_global(movs_valid)
        facturas_cuenta = construir_facturas_por_cuenta(movs_valid)
        detalle_cruces, resumen_cruces = detectar_cruces_referencias(movs_valid)

        # Cálculo de cuentas sin facturas (saldos muertos)
        cuentas_con_facturas = facturas_cuenta[["account_code", "account_name"]].drop_duplicates()
        aux_sin_facturas = totales_cuentas_aux.merge(cuentas_con_facturas, on=["account_code", "account_name"], how="left", indicator=True)
        saldo_cuentas_sin_facturas = aux_sin_facturas.loc[aux_sin_facturas["_merge"] == "left_only", "saldo_final_cuenta_aux"].sum()
        
        # Diferencia residual global
        conciliado = resumen_aux["saldo_neto_movs"] + saldo_cuentas_sin_facturas
        diferencia_residual = (resumen_aux["saldo_final_aux"] or 0) - conciliado

    if facturas_global.empty and facturas_cuenta.empty:
        st.success("✅ No se encontraron facturas en el archivo.")
    
    else:
        # ----------------------------------------------------------------
        # SECCIÓN DE NARRATIVA Y GRÁFICO (INSIGHTS)
        # ----------------------------------------------------------------
        st.header("📊 Diagnóstico Ejecutivo")
        
        col_narrativa, col_grafico = st.columns([1, 1])

        with col_narrativa:
            # Texto inteligente que explica la situación
            hay_diferencia_grave = abs(diferencia_residual) > 10.0
            emoji_status = "⚠️" if hay_diferencia_grave else "✅"
            
            st.markdown(f"""
            ### {emoji_status} Resumen de Conciliación
            El saldo contable total reportado por el auxiliar es de **${resumen_aux['saldo_final_aux']:,.2f}**.
            
            **¿Cómo se compone este saldo?**
            1. **${resumen_aux['saldo_neto_movs']:,.2f}** están soportados por movimientos de facturas en este reporte.
            2. **${saldo_cuentas_sin_facturas:,.2f}** corresponden a cuentas inactivas o saldos iniciales sin movimientos referenciados.
            
            **Resultado:**
            """)
            
            if hay_diferencia_grave:
                st.error(f"""
                Existe una **diferencia no explicada de ${diferencia_residual:,.2f}**.
                Esto suele deberse a pólizas manuales sin referencia o ajustes contables directos.
                Revisa las cuentas marcadas en ROJO 🔴 en la pestaña de detalle.
                """)
            else:
                st.success(f"""
                **La conciliación es correcta.** La diferencia residual es de solo ${diferencia_residual:,.2f}, 
                probablemente debida a redondeos.
                """)

        with col_grafico:
            # Gráfico de cascada para ver composición del saldo
            fig = go.Figure(data=[go.Bar(
                x=['Facturas Vigentes', 'Ctas Inactivas/Inicial', 'Diferencia (Error)'],
                y=[resumen_aux['saldo_neto_movs'], saldo_cuentas_sin_facturas, diferencia_residual],
                marker_color=['#2ecc71', '#f1c40f', '#e74c3c'], # Verde, Amarillo, Rojo
                texttemplate='$%{y:,.0f}',
                textposition='auto'
            )])
            fig.update_layout(
                title="Composición del Saldo Auxiliar",
                yaxis_title="Importe ($)",
                margin=dict(l=20, r=20, t=40, b=20),
                height=300
            )
            st.plotly_chart(fig, use_container_width=True)

        st.divider()

        # ------------------------- Filtros -------------------------
        all_fechas = pd.concat([facturas_global["fecha_factura"], facturas_cuenta["fecha_factura"]]).dropna()
        if not all_fechas.empty:
            min_date, max_date = all_fechas.min(), all_fechas.max()
        else:
            min_date, max_date = pd.Timestamp.now(), pd.Timestamp.now()

        col_f1, col_f2, col_f3 = st.columns([1, 1, 2])
        with col_f1:
            fecha_desde = st.date_input("Fecha desde", value=min_date.date() if pd.notna(min_date) else None)
        with col_f2:
            fecha_hasta = st.date_input("Fecha hasta", value=max_date.date() if pd.notna(max_date) else None)
        with col_f3:
            todas_cuentas = facturas_cuenta["cuenta"].dropna().sort_values().unique().tolist()
            cuenta_seleccionada = st.selectbox("Filtrar por Cuenta (opcional)", options=["(Todas las cuentas)"] + todas_cuentas)

        codigo_cuenta_seleccionada = None
        if cuenta_seleccionada != "(Todas las cuentas)":
            codigo_cuenta_seleccionada = cuenta_seleccionada.split(" - ")[0].strip()

        # Aplicar filtros de fecha
        facturas_global_f = filtrar_por_fecha(facturas_global, fecha_desde, fecha_hasta)
        facturas_cuenta_f = filtrar_por_fecha(facturas_cuenta, fecha_desde, fecha_hasta)

        # ----------------------------------------------------------------
        # PESTAÑAS
        # ----------------------------------------------------------------
        tab_resumen, tab_pendientes, tab_favor = st.tabs([
            "📂 Semáforo de Cuentas", 
            "📑 Facturas Pendientes", 
            "💳 Saldos a Favor"
        ])

        # ================================================================
        # TAB 1: SEMÁFORO (CON FORMATO CONDICIONAL)
        # ================================================================
        with tab_resumen:
            st.subheader("Estado de Conciliación por Cuenta")
            
            # Preparar datos base
            tot_aux_f = totales_cuentas_aux.copy()
            if codigo_cuenta_seleccionada:
                tot_aux_f = tot_aux_f[tot_aux_f["account_code"] == codigo_cuenta_seleccionada]

            if tot_aux_f.empty:
                st.info("No hay datos para mostrar con los filtros actuales.")
            else:
                # Calcular métricas desde facturas
                if not facturas_cuenta_f.empty:
                    metrics = facturas_cuenta_f.groupby(["account_code", "account_name"]).agg(
                        saldo_neto=("saldo_factura", "sum")
                    ).reset_index()
                else:
                    metrics = pd.DataFrame(columns=["account_code", "account_name", "saldo_neto"])

                # Cruce con auxiliar
                resumen_cuenta = tot_aux_f.merge(metrics, on=["account_code", "account_name"], how="left")
                resumen_cuenta["saldo_neto"] = resumen_cuenta["saldo_neto"].fillna(0)
                resumen_cuenta["diferencia_vs_auxiliar"] = resumen_cuenta["saldo_final_cuenta_aux"] - resumen_cuenta["saldo_neto"]
                resumen_cuenta["solo_saldo_inicial"] = (resumen_cuenta["saldo_neto"].abs() < UMBRAL_SALDO_INICIAL) & \
                                                       (resumen_cuenta["diferencia_vs_auxiliar"].abs() > UMBRAL_SALDO_INICIAL)

                # --- Lógica de Semáforo ---
                def clasificar_estado(row):
                    if abs(row['diferencia_vs_auxiliar']) > UMBRAL_SALDO_INICIAL and not row['solo_saldo_inicial']:
                         return "🔴 Revisar (Diferencia)"
                    elif row['solo_saldo_inicial']:
                         return "🟡 Solo Saldo Inicial"
                    elif row['saldo_neto'] == 0 and row['saldo_final_cuenta_aux'] == 0:
                         return "⚪ Saldada"
                    else:
                         return "🟢 Correcta"

                resumen_cuenta['Estado'] = resumen_cuenta.apply(clasificar_estado, axis=1)

                # --- Filtro de Acción Inmediata ---
                col_t1, col_t2 = st.columns([2,1])
                with col_t1:
                    solo_urgente = st.toggle("🚨 Mostrar solo lo urgente (Diferencias y Saldos Iniciales)", value=True)
                
                if solo_urgente:
                    df_view = resumen_cuenta[
                        (resumen_cuenta['Estado'].str.contains("🔴")) | 
                        (resumen_cuenta['Estado'].str.contains("🟡"))
                    ]
                    if df_view.empty:
                        st.success("¡Excelente! No hay cuentas urgentes que revisar.")
                else:
                    df_view = resumen_cuenta

                # --- Dataframe con Column Config ---
                st.dataframe(
                    df_view[[
                        "account_code", "account_name", "Estado", 
                        "saldo_neto", "saldo_final_cuenta_aux", "diferencia_vs_auxiliar"
                    ]].sort_values("diferencia_vs_auxiliar", ascending=False),
                    column_config={
                        "account_code": "Cuenta",
                        "account_name": "Nombre",
                        "saldo_neto": st.column_config.NumberColumn("Suma Facturas", format="$%.2f"),
                        "saldo_final_cuenta_aux": st.column_config.NumberColumn("Saldo Auxiliar", format="$%.2f"),
                        "diferencia_vs_auxiliar": st.column_config.ProgressColumn(
                            "Diferencia", format="$%.2f", 
                            min_value=0, max_value=float(resumen_cuenta["diferencia_vs_auxiliar"].abs().max() or 100)
                        ),
                        "Estado": st.column_config.TextColumn("Diagnóstico"),
                    },
                    use_container_width=True,
                    hide_index=True
                )
                
                st.download_button(
                    "⬇️ Descargar Resumen en Excel",
                    data=to_excel(resumen_cuenta),
                    file_name="resumen_conciliacion.xlsx"
                )

        # ================================================================
        # TAB 2: PENDIENTES
        # ================================================================
        with tab_pendientes:
            base_df = facturas_cuenta_f[facturas_cuenta_f["account_code"] == codigo_cuenta_seleccionada].copy() \
                      if codigo_cuenta_seleccionada else facturas_global_f.copy()
            
            df_pend = base_df[base_df["saldo_factura"] > 0].copy()
            
            col_k1, col_k2 = st.columns(2)
            col_k1.metric("Facturas Pendientes", value=len(df_pend))
            col_k2.metric("Saldo Total Pendiente", value=f"${df_pend['saldo_factura'].sum():,.2f}")
            
            st.dataframe(
                df_pend[["referencia", "fecha_factura", "saldo_factura", "account_name"]].sort_values("fecha_factura"),
                use_container_width=True,
                column_config={
                    "saldo_factura": st.column_config.NumberColumn("Saldo Pendiente", format="$%.2f")
                }
            )
            if not df_pend.empty:
                st.download_button("⬇️ Descargar Excel", data=to_excel(df_pend), file_name="facturas_pendientes.xlsx")

        # ================================================================
        # TAB 3: SALDOS A FAVOR
        # ================================================================
        with tab_favor:
            base_df_f = facturas_cuenta_f[facturas_cuenta_f["account_code"] == codigo_cuenta_seleccionada].copy() \
                        if codigo_cuenta_seleccionada else facturas_global_f.copy()
            
            df_favor = base_df_f[base_df_f["saldo_factura"] < 0].copy()

            if df_favor.empty:
                st.info("No hay facturas con saldo a favor.")
            else:
                col_k1, col_k2 = st.columns(2)
                col_k1.metric("Notas/Pagos a favor", value=len(df_favor))
                col_k2.metric("Total a favor", value=f"${df_favor['saldo_factura'].sum():,.2f}")
                
                st.dataframe(
                    df_favor[["referencia", "fecha_factura", "saldo_factura", "account_name"]],
                    use_container_width=True,
                    column_config={
                        "saldo_factura": st.column_config.NumberColumn("Saldo a Favor", format="$%.2f")
                    }
                )
                st.download_button("⬇️ Descargar Excel", data=to_excel(df_favor), file_name="saldos_a_favor.xlsx")

        # ================================================================
        # SECCIÓN EXTRA: CRUCES (SI EXISTEN)
        # ================================================================
        if not detalle_cruces.empty:
            st.divider()
            st.warning("🔄 Se detectaron referencias cruzadas (Cargos en una cuenta, abonos en otra)")
            st.dataframe(detalle_cruces)
