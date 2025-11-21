import streamlit as st
import pandas as pd
import numpy as np
import re
from io import BytesIO
from datetime import datetime, date

st.set_page_config(page_title="Facturas no pagadas", layout="wide")

st.title("🔍 Reporte de facturas no pagadas")
st.write(
    "Sube el archivo de **Movimientos, Auxiliares del Catálogo** generado desde CONTPAQ i "
    "y el sistema identificará las facturas no pagadas por cliente."
)

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
        'Ene': '01', 'Feb': '02', 'Mar': '03', 'Abr': '04',
        'May': '05', 'Jun': '06', 'Jul': '07', 'Ago': '08',
        'Sep': '09', 'Oct': '10', 'Nov': '11', 'Dic': '12',
    }
    mon_key = mon_abbr[:3].title()
    if mon_key not in month_map:
        return pd.NaT
    date_str = f"{day}/{month_map[mon_key]}/{year}"
    return pd.to_datetime(date_str, format="%d/%m/%Y", errors="coerce")


@st.cache_data
def procesar_archivo(file) -> pd.DataFrame:
    """
    Lee el Excel de CONTPAQ y regresa un DataFrame de facturas (una fila por factura)
    con su saldo pendiente por cliente.
    """
    # Leer tal cual, sin encabezados
    raw = pd.read_excel(file, header=None)

    # Detectar filas cabecera de cuenta (tienen el código de cuenta y 'Saldo inicial :')
    account_pattern = re.compile(r"^\d{3}-\d{3}-\d{3}-\d{3}$")
    is_account_header = raw[0].astype(str).str.match(account_pattern) & \
                        raw[6].astype(str).str.contains("Saldo inicial", na=False)

    # Rellenar código y nombre de cuenta hacia abajo
    df = raw.copy()
    df["account_code"] = np.where(is_account_header, df[0], np.nan)
    df["account_name"] = np.where(is_account_header, df[1], np.nan)
    df["account_code"] = df["account_code"].ffill()
    df["account_name"] = df["account_name"].ffill()

    # Filas de movimientos: aquellas donde la columna 0 es una fecha dd/Mon/aaaa
    date_pattern = re.compile(r"^\d{2}/[A-Za-z]{3}/\d{4}$")
    is_date_row = df[0].astype(str).str.match(date_pattern)

    movs = df.loc[is_date_row].copy()
    movs = movs.rename(
        columns={
            0: "fecha_raw",
            1: "tipo",
            2: "numero_poliza",
            3: "concepto",
            4: "referencia",
            5: "cargos",
            6: "abonos",
            7: "saldo",
        }
    )

    # Convertir importes a numérico
    for col in ["cargos", "abonos", "saldo"]:
        movs[col] = pd.to_numeric(movs[col], errors="coerce")

    # Limpiar referencia (número de factura)
    movs["referencia"] = movs["referencia"].astype(str).str.strip()
    movs["referencia"] = movs["referencia"].replace({"nan": np.nan, "": np.nan})

    # Convertir fecha
    movs["fecha"] = movs["fecha_raw"].apply(parse_spanish_date)

    # Nos quedamos solo con movimientos que tienen número de factura
    movs_valid = movs[movs["referencia"].notna()].copy()

    # Agrupamos por cliente + referencia (factura)
    group_cols = ["account_code", "account_name", "referencia"]
    facturas = (
        movs_valid.groupby(group_cols)
        .agg(
            fecha_factura=("fecha", "min"),
            cargos_total=("cargos", "sum"),
            abonos_total=("abonos", "sum"),
        )
        .reset_index()
    )

    facturas["saldo_factura"] = facturas["cargos_total"] - facturas["abonos_total"]

    return facturas


def filtrar_facturas(df_facturas: pd.DataFrame, fecha_desde: date, fecha_hasta: date, clientes_sel):
    mask = pd.Series(True, index=df_facturas.index)

    if fecha_desde:
        mask &= df_facturas["fecha_factura"] >= pd.to_datetime(fecha_desde)
    if fecha_hasta:
        mask &= df_facturas["fecha_factura"] <= pd.to_datetime(fecha_hasta)
    if clientes_sel:
        mask &= df_facturas["account_name"].isin(clientes_sel)

    return df_facturas[mask].copy()


def to_excel(df: pd.DataFrame) -> bytes:
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Facturas_pendientes")
    return output.getvalue()


uploaded_file = st.file_uploader(
    "📎 Sube el archivo Excel de movimientos (auxiliares del catálogo)",
    type=["xlsx"]
)

if uploaded_file is None:
    st.info(
        "Sube un archivo `.xlsx` exportado desde CONTPAQ "
        "(Movimientos, Auxiliares del Catálogo) para comenzar."
    )
else:
    with st.spinner("Procesando archivo..."):
        facturas = procesar_archivo(uploaded_file)

    # Solo facturas con saldo pendiente > 0
    facturas_pend = facturas[facturas["saldo_factura"] > 0].copy()

    if facturas_pend.empty:
        st.success("✅ No se encontraron facturas con saldo pendiente en el archivo.")
    else:
        st.subheader("📌 Filtros")

        min_date = facturas_pend["fecha_factura"].min()
        max_date = facturas_pend["fecha_factura"].max()

        col_f1, col_f2 = st.columns(2)
        with col_f1:
            fecha_desde = st.date_input(
                "Fecha desde",
                value=min_date.date() if pd.notna(min_date) else None,
                min_value=min_date.date() if pd.notna(min_date) else None,
                max_value=max_date.date() if pd.notna(max_date) else None,
            )
        with col_f2:
            fecha_hasta = st.date_input(
                "Fecha hasta",
                value=max_date.date() if pd.notna(max_date) else None,
                min_value=min_date.date() if pd.notna(min_date) else None,
                max_value=max_date.date() if pd.notna(max_date) else None,
            )

        clientes = sorted(facturas_pend["account_name"].dropna().unique().tolist())
        clientes_sel = st.multiselect(
            "Clientes (cuentas contables)", options=clientes, default=[]
        )

        facturas_filtradas = filtrar_facturas(
            facturas_pend, fecha_desde, fecha_hasta, clientes_sel
        )

        st.subheader("📊 Resumen por cliente")

        resumen = (
            facturas_filtradas.groupby(["account_code", "account_name"])
            .agg(
                facturas_pendientes=("referencia", "nunique"),
                saldo_pendiente_total=("saldo_factura", "sum"),
            )
            .reset_index()
            .sort_values("saldo_pendiente_total", ascending=False)
        )

        c1, c2 = st.columns(2)
        with c1:
            st.metric(
                "Total de facturas pendientes",
                value=int(resumen["facturas_pendientes"].sum()),
            )
        with c2:
            st.metric(
                "Saldo pendiente total",
                value=f"${resumen['saldo_pendiente_total'].sum():,.2f}",
            )

        st.dataframe(resumen, use_container_width=True)

        st.subheader("📄 Detalle de facturas pendientes")
        st.dataframe(
            facturas_filtradas.sort_values(
                ["account_name", "fecha_factura", "referencia"]
            ),
            use_container_width=True,
        )

        # Descarga
        xls_bytes = to_excel(facturas_filtradas)
        st.download_button(
            label="⬇️ Descargar detalle en Excel",
            data=xls_bytes,
            file_name="facturas_pendientes.xlsx",
            mime=(
                "application/vnd.openxmlformats-officedocument."
                "spreadsheetml.sheet"
            ),
        )
