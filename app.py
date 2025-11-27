import streamlit as st
import pandas as pd
import numpy as np
import re
import math
from io import BytesIO
from datetime import date

st.set_page_config(page_title="Facturas no pagadas", layout="wide")

st.title("🔍 Auditoria Integracion de Saldos")
st.write(
    "Sube el archivo de **Movimientos, Auxiliares del Catálogo** generado desde CONTPAQ i "
    "y el sistema identificará las facturas no pagadas, con cuatro vistas: "
    "**por factura (global)**, **por cuenta contable (sin cruzar cuentas)**, "
    "**facturas cruzadas pendientes** y **facturas cruzadas pagadas**."
)

# --------------------------------------------------------------------
# Utilidades
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
        "Ene": "01",
        "Feb": "02",
        "Mar": "03",
        "Abr": "04",
        "May": "05",
        "Jun": "06",
        "Jul": "07",
        "Ago": "08",
        "Sep": "09",
        "Oct": "10",
        "Nov": "11",
        "Dic": "12",
    }
    mon_key = mon_abbr[:3].title()
    if mon_key not in month_map:
        return pd.NaT
    date_str = f"{day}/{month_map[mon_key]}/{year}"
    return pd.to_datetime(date_str, format="%d/%m/%Y", errors="coerce")


@st.cache_data
def procesar_movimientos(file):
    """
    Lee el Excel de CONTPAQ y regresa:
      - movs_valid: DataFrame de movimientos con referencia (y cuenta asignada).
      - resumen_auxiliar: dict con totales globales (cargos, abonos, saldo final auxiliar, etc.)
      - totales_cuentas_aux: DataFrame con totales por cuenta (Total: del auxiliar).
    """
    # Leer tal cual, sin encabezados
    raw = pd.read_excel(file, header=None)

    # Detectar filas cabecera de cuenta (tienen el código de cuenta y 'Saldo inicial :')
    account_pattern = re.compile(r"^\d{3}-\d{3}-\d{3}-\d{3}$")
    is_account_header = raw[0].astype(str).str.match(account_pattern) & raw[
        6
    ].astype(str).str.contains("Saldo inicial", na=False)

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

    # Movimientos con número de factura
    movs_valid = movs[movs["referencia"].notna()].copy()

    # ----------------------------------------------------------------
    # Totales globales del auxiliar (fila 'Total Clientes :')
    # ----------------------------------------------------------------
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

    # Totales de movimientos (solo filas con referencia)
    total_cargos_movs = movs_valid["cargos"].sum()
    total_abonos_movs = movs_valid["abonos"].sum()
    saldo_neto_movs = total_cargos_movs - total_abonos_movs

    saldo_inicial_implicito = np.nan
    if not math.isnan(saldo_final_aux) and not math.isnan(saldo_neto_movs):
        saldo_inicial_implicito = saldo_final_aux - saldo_neto_movs

    resumen_auxiliar = {
        "total_cargos_movs": float(total_cargos_movs),
        "total_abonos_movs": float(total_abonos_movs),
        "saldo_neto_movs": float(saldo_neto_movs),
        "total_cargos_aux": float(total_cargos_aux) if not math.isnan(total_cargos_aux) else None,
        "total_abonos_aux": float(total_abonos_aux) if not math.isnan(total_abonos_aux) else None,
        "saldo_final_aux": float(saldo_final_aux) if not math.isnan(saldo_final_aux) else None,
        "saldo_inicial_implicito": float(saldo_inicial_implicito)
        if not math.isnan(saldo_inicial_implicito)
        else None,
    }

    # ----------------------------------------------------------------
    # Totales por cuenta (fila 'Total:' en la columna 4)
    # ----------------------------------------------------------------
    tot_rows = df[df[4].astype(str).str.strip() == "Total:"].copy()
    for c in [5, 6, 7]:
        tot_rows[c] = pd.to_numeric(tot_rows[c], errors="coerce")

    totales_cuentas_aux = (
        tot_rows.rename(
            columns={
                5: "cargos_total_cuenta_aux",
                6: "abonos_total_cuenta_aux",
                7: "saldo_final_cuenta_aux",
            }
        )[
            [
                "account_code",
                "account_name",
                "cargos_total_cuenta_aux",
                "abonos_total_cuenta_aux",
                "saldo_final_cuenta_aux",
            ]
        ]
        .groupby(["account_code", "account_name"])
        .agg(
            cargos_total_cuenta_aux=("cargos_total_cuenta_aux", "sum"),
            abonos_total_cuenta_aux=("abonos_total_cuenta_aux", "sum"),
            saldo_final_cuenta_aux=("saldo_final_cuenta_aux", "sum"),
        )
        .reset_index()
    )

    return movs_valid, resumen_auxiliar, totales_cuentas_aux


@st.cache_data
def construir_facturas_global(movs_valid: pd.DataFrame) -> pd.DataFrame:
    """
    Construye facturas a nivel global (por referencia), cruzando todas las cuentas.
    Asigna una 'cuenta principal' por factura (normalmente donde está el cargo) y
    calcula en cuántas cuentas aparece cada referencia.
    """
    # 1) Agregados globales por referencia
    facturas = (
        movs_valid.groupby("referencia")
        .agg(
            fecha_factura=("fecha", "min"),
            cargos_total=("cargos", "sum"),
            abonos_total=("abonos", "sum"),
        )
        .reset_index()
    )

    # 2) Determinar cuenta principal (preferimos donde haya cargos positivos)
    movs_valid = movs_valid.copy()
    movs_valid["es_cargo_pos"] = movs_valid["cargos"] > 0

    # a) Principal desde cargos positivos (factura original)
    main_from_cargo = (
        movs_valid[movs_valid["es_cargo_pos"]]
        .sort_values(["referencia", "cargos"], ascending=[True, False])
        .drop_duplicates("referencia")
        [["referencia", "account_code", "account_name"]]
    )

    # b) Para referencias sin cargos positivos, tomar el primer movimiento que aparezca
    main_any = (
        movs_valid.sort_values(["referencia", "fecha"])
        .drop_duplicates("referencia")
        [["referencia", "account_code", "account_name"]]
    )

    main_account = pd.concat([main_from_cargo, main_any], ignore_index=True)
    main_account = main_account.drop_duplicates("referencia", keep="first")

    facturas = facturas.merge(main_account, on="referencia", how="left")

    # 3) Número de cuentas involucradas por referencia + lista de cuentas
    cuentas_por_factura = (
        movs_valid.groupby("referencia")["account_code"]
        .nunique()
        .reset_index(name="num_cuentas")
    )
    facturas = facturas.merge(cuentas_por_factura, on="referencia", how="left")
    facturas["num_cuentas"] = facturas["num_cuentas"].fillna(0).astype(int)
    facturas["cruza_cuentas"] = facturas["num_cuentas"] > 1

    # Lista de cuentas involucradas (código + nombre) como texto
    cuentas_involucradas = (
        movs_valid.assign(
            cuenta=lambda d: d["account_code"].astype(str)
            + " - "
            + d["account_name"].astype(str)
        )
        .groupby("referencia")["cuenta"]
        .apply(lambda x: " | ".join(sorted(set(x))))
        .reset_index(name="cuentas_involucradas")
    )
    facturas = facturas.merge(cuentas_involucradas, on="referencia", how="left")

    # 4) Saldo pendiente por referencia (neto C-A)
    facturas["saldo_factura"] = facturas["cargos_total"] - facturas["abonos_total"]

    return facturas


@st.cache_data
def construir_facturas_por_cuenta(movs_valid: pd.DataFrame) -> pd.DataFrame:
    """
    Construye facturas por cuenta contable (sin cruzar cuentas).
    Una misma referencia puede aparecer en varias cuentas.
    """
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


def filtrar_por_fecha(df: pd.DataFrame, fecha_desde: date, fecha_hasta: date) -> pd.DataFrame:
    """Filtra un DataFrame de facturas por rango de fechas."""
    mask = pd.Series(True, index=df.index)
    if fecha_desde:
        mask &= df["fecha_factura"] >= pd.to_datetime(fecha_desde)
    if fecha_hasta:
        mask &= df["fecha_factura"] <= pd.to_datetime(fecha_hasta)
    return df[mask].copy()


def to_excel(df: pd.DataFrame) -> bytes:
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Facturas")
    return output.getvalue()


# --------------------------------------------------------------------
# App
# --------------------------------------------------------------------

uploaded_file = st.file_uploader(
    "📎 Sube el archivo Excel de movimientos (auxiliares del catálogo)",
    type=["xlsx"],
)

if uploaded_file is None:
    st.info(
        "Sube un archivo `.xlsx` exportado desde CONTPAQ "
        "(Movimientos, Auxiliares del Catálogo) para comenzar."
    )
else:
    with st.spinner("Procesando archivo..."):
        movs_valid, resumen_aux, totales_cuentas_aux = procesar_movimientos(uploaded_file)
        facturas_global = construir_facturas_global(movs_valid)
        facturas_cuenta = construir_facturas_por_cuenta(movs_valid)

    # Solo facturas con saldo pendiente > 0 (para algunas vistas)
    facturas_global_pend = facturas_global[facturas_global["saldo_factura"] > 0].copy()
    facturas_cuenta_pend = facturas_cuenta[facturas_cuenta["saldo_factura"] > 0].copy()

    # Si de plano no hay facturas, salimos
    if facturas_global.empty and facturas_cuenta.empty:
        st.success("✅ No se encontraron facturas en el archivo.")
    else:
        # ------------------------------------------------------------
        # RESUMEN GLOBAL VS AUXILIAR (solo netos)
        # ------------------------------------------------------------
        st.subheader("📊 Resumen global vs auxiliar")

        colg1, colg2, colg3 = st.columns(3)
        with colg1:
            st.metric(
                "Saldo neto movimientos (C-A)",
                value=f"${resumen_aux['saldo_neto_movs']:,.2f}",
            )
        with colg2:
            if resumen_aux.get("saldo_final_aux") is not None:
                st.metric(
                    "Saldo final cartera (auxiliar - Total Clientes)",
                    value=f"${resumen_aux['saldo_final_aux']:,.2f}",
                )
            else:
                st.metric(
                    "Saldo final cartera (auxiliar)",
                    value="N/D",
                )
        with colg3:
            if (
                resumen_aux.get("saldo_final_aux") is not None
                and resumen_aux.get("saldo_neto_movs") is not None
            ):
                diferencia_global = (
                    resumen_aux["saldo_final_aux"] - resumen_aux["saldo_neto_movs"]
                )
                st.metric(
                    "Diferencia (auxiliar - neto movimientos)",
                    value=f"${diferencia_global:,.2f}",
                )
            else:
                st.metric(
                    "Diferencia (auxiliar - neto movimientos)",
                    value="N/D",
                )

        if resumen_aux.get("saldo_inicial_implicito") is not None:
            st.caption(
                f"Saldo inicial implícito según auxiliar: "
                f"${resumen_aux['saldo_inicial_implicito']:,.2f} "
                f"(Saldo final auxiliar - saldo neto de movimientos con referencia)."
            )

        # Rango de fechas global para filtros (usamos todas las facturas, no solo pendientes)
        all_fechas = pd.concat(
            [
                facturas_global["fecha_factura"],
                facturas_cuenta["fecha_factura"],
            ]
        ).dropna()

        min_date = all_fechas.min()
        max_date = all_fechas.max()

        st.subheader("📌 Filtros de periodo")

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

        # ----------------------------------------------------------------
        # Vistas en pestañas
        # ----------------------------------------------------------------
        (
            tab_global,
            tab_cuenta,
            tab_cruzadas,
            tab_cruzadas_pagadas,
        ) = st.tabs(
            [
                "📑 Por factura (global)",
                "📂 Por cuenta contable (sin cruzar cuentas)",
                "🧩 Facturas cruzadas pendientes",
                "✅ Facturas cruzadas pagadas",
            ]
        )

        # ================================================================
        # TAB 1: Por factura (global)
        # ================================================================
        with tab_global:
            st.markdown("### Vista por factura (global)")
            st.caption(
                "Agrupa por **referencia de factura**, cruzando todas las cuentas de clientes. "
                "Muestra el **saldo neto** por factura a nivel global y permite conciliar por cuenta."
            )

            # Filtramos el universo completo de referencias en el rango de fechas
            facturas_global_filtradas = filtrar_por_fecha(
                facturas_global, fecha_desde=fecha_desde, fecha_hasta=fecha_hasta
            ).copy()

            # Agregamos columna de texto de cuenta principal
            facturas_global_filtradas["cuenta"] = (
                facturas_global_filtradas["account_code"].astype(str)
                + " - "
                + facturas_global_filtradas["account_name"].astype(str)
            )

            # Selector de cuenta principal
            cuentas_global = (
                facturas_global_filtradas["cuenta"]
                .dropna()
                .sort_values()
                .unique()
                .tolist()
            )
            cuentas_sel_global = st.multiselect(
                "Cuenta principal (opcional)",
                options=cuentas_global,
                default=[],
                key="cuentas_global",
            )

            # Aplicamos filtro por cuenta (si se selecciona)
            if cuentas_sel_global:
                subset_all = facturas_global_filtradas[
                    facturas_global_filtradas["cuenta"].isin(cuentas_sel_global)
                ].copy()
            else:
                subset_all = facturas_global_filtradas

            if subset_all.empty:
                st.info("No hay facturas en este rango de fechas / filtros.")
            else:
                # Saldos netos por referencia dentro del filtro actual
                saldo_neto_referencias = subset_all["saldo_factura"].sum()

                st.subheader("📊 Resumen por referencia (según filtros actuales)")

                col_resumen = st.columns(1)[0]
                with col_resumen:
                    st.metric(
                        "Saldo neto por referencia (C-A)",
                        value=f"${saldo_neto_referencias:,.2f}",
                    )

                # Conciliación contra auxiliar para las cuentas seleccionadas (si aplica)
                if cuentas_sel_global:
                    aux_tmp = totales_cuentas_aux.copy()
                    aux_tmp["cuenta"] = (
                        aux_tmp["account_code"].astype(str)
                        + " - "
                        + aux_tmp["account_name"].astype(str)
                    )
                    aux_subset = aux_tmp[aux_tmp["cuenta"].isin(cuentas_sel_global)]

                    saldo_final_aux_sel = aux_subset["saldo_final_cuenta_aux"].sum()

                    c4, c5 = st.columns(2)
                    with c4:
                        st.metric(
                            "Saldo final cuenta(s) (auxiliar)",
                            value=f"${saldo_final_aux_sel:,.2f}",
                        )
                    with c5:
                        diferencia_vs_aux = saldo_final_aux_sel - saldo_neto_referencias
                        st.metric(
                            "Diferencia vs auxiliar",
                            value=f"${diferencia_vs_aux:,.2f}",
                        )

                    st.caption(
                        "La diferencia vs auxiliar incluye saldo inicial de la(s) cuenta(s) "
                        "y movimientos sin referencia o fuera del rango de fechas."
                    )
                else:
                    # Conciliación global contra auxiliar (si tenemos dato)
                    if resumen_aux.get("saldo_final_aux") is not None:
                        diferencia_vs_aux_global = (
                            resumen_aux["saldo_final_aux"] - saldo_neto_referencias
                        )
                        st.caption(
                            f"Saldo final auxiliar global (Total Clientes): "
                            f"${resumen_aux['saldo_final_aux']:,.2f}. "
                            f"Diferencia vs saldo neto por referencia: "
                            f"${diferencia_vs_aux_global:,.2f} "
                            f"(incluye saldo inicial global y movimientos sin referencia)."
                        )

                # Ahora solo pendientes (saldo > 0) para detalle
                df_tab1 = subset_all[subset_all["saldo_factura"] > 0].copy()

                if df_tab1.empty:
                    st.info(
                        "No hay facturas pendientes (saldo > 0) con estos filtros."
                    )
                else:
                    # Resumen por cuenta principal (solo facturas pendientes)
                    st.subheader("📊 Resumen por cuenta principal (solo pendientes)")

                    resumen_global = (
                        df_tab1.groupby(["account_code", "account_name"])
                        .agg(
                            facturas_pendientes=("referencia", "nunique"),
                            saldo_pendiente_total=("saldo_factura", "sum"),
                        )
                        .reset_index()
                        .sort_values("saldo_pendiente_total", ascending=False)
                    )

                    c4, c5 = st.columns(2)
                    with c4:
                        st.metric(
                            "Total de facturas pendientes (global)",
                            value=int(df_tab1["referencia"].nunique()),
                        )
                    with c5:
                        st.metric(
                            "Saldo pendiente total (solo facturas positivas)",
                            value=f"${df_tab1['saldo_factura'].sum():,.2f}",
                        )

                    st.dataframe(resumen_global, use_container_width=True)

                    # Detalle por factura (solo netos)
                    st.subheader("📄 Detalle de facturas (global)")

                    cols_detalle_global = [
                        "referencia",
                        "fecha_factura",
                        "saldo_factura",
                        "account_code",
                        "account_name",
                        "num_cuentas",
                        "cruza_cuentas",
                        "cuentas_involucradas",
                    ]

                    df_detalle_global = df_tab1[cols_detalle_global].sort_values(
                        ["fecha_factura", "referencia"]
                    )

                    st.dataframe(df_detalle_global, use_container_width=True)

                    # Descarga Excel
                    xls_global = to_excel(df_detalle_global)
                    st.download_button(
                        label="⬇️ Descargar detalle global en Excel",
                        data=xls_global,
                        file_name="facturas_pendientes_global.xlsx",
                        mime=(
                            "application/vnd.openxmlformats-officedocument."
                            "spreadsheetml.sheet"
                        ),
                    )

        # ================================================================
        # TAB 2: Por cuenta contable (sin cruzar cuentas)
        # ================================================================
        with tab_cuenta:
            st.markdown("### Vista por cuenta contable (sin cruzar cuentas)")
            st.caption(
                "Agrupa por **número de cuenta + nombre de cuenta**. "
                "La misma referencia puede aparecer en varias cuentas; aquí NO se cruzan. "
                "Se muestran solo **saldos netos** por referencia y el saldo final del auxiliar."
            )

            # Usamos TODO el universo de facturas por cuenta (positivas y negativas), filtrado por fecha
            df_cuenta_all = filtrar_por_fecha(
                facturas_cuenta,
                fecha_desde=fecha_desde,
                fecha_hasta=fecha_hasta,
            )

            # Columna de texto de cuenta
            df_cuenta_all["cuenta"] = (
                df_cuenta_all["account_code"].astype(str)
                + " - "
                + df_cuenta_all["account_name"].astype(str)
            )

            cuentas_cuenta = (
                df_cuenta_all["cuenta"]
                .dropna()
                .sort_values()
                .unique()
                .tolist()
            )
            cuentas_sel_cuenta = st.multiselect(
                "Cuenta contable",
                options=cuentas_cuenta,
                default=[],
                key="cuentas_cuenta",
            )
            if cuentas_sel_cuenta:
                df_cuenta_all = df_cuenta_all[
                    df_cuenta_all["cuenta"].isin(cuentas_sel_cuenta)
                ]

            if df_cuenta_all.empty:
                st.info("No hay facturas en este rango de fechas / filtros.")
            else:
                # Resumen por cuenta: usamos saldo neto; positivas/negativas solo internas
                resumen_cuenta = (
                    df_cuenta_all.groupby(["account_code", "account_name"])
                    .agg(
                        facturas_positivas=("saldo_factura", lambda s: int((s > 0).sum())),
                        saldo_positivas=("saldo_factura", lambda s: float(s[s > 0].sum())),
                        referencias_negativas=("saldo_factura", lambda s: int((s < 0).sum())),
                        saldo_negativas=("saldo_factura", lambda s: float(s[s < 0].sum())),
                        saldo_neto=("saldo_factura", "sum"),
                    )
                    .reset_index()
                )

                # Unimos totales del auxiliar por cuenta (cargos, abonos, saldo final)
                resumen_cuenta = resumen_cuenta.merge(
                    totales_cuentas_aux,
                    on=["account_code", "account_name"],
                    how="left",
                )

                resumen_cuenta["diferencia_vs_auxiliar"] = (
                    resumen_cuenta["saldo_final_cuenta_aux"] - resumen_cuenta["saldo_neto"]
                )

                # Métrica global (en el rango de fechas actual): solo neto
                saldo_neto_total = resumen_cuenta["saldo_neto"].sum()

                col_net = st.columns(1)[0]
                with col_net:
                    st.metric(
                        "Saldo neto por referencia (en rango)",
                        value=f"${saldo_neto_total:,.2f}",
                    )

                st.caption(
                    "Los saldos del auxiliar (cargos, abonos y saldo final por cuenta) "
                    "corresponden al periodo completo del reporte exportado. "
                    "El saldo neto por referencia corresponde solo al rango de fechas filtrado."
                )

                # Tabla de resumen por cuenta (solo netos + auxiliar)
                st.subheader("📊 Resumen por cuenta contable")

                cols_resumen = [
                    "account_code",
                    "account_name",
                    "facturas_positivas",
                    "referencias_negativas",
                    "saldo_neto",
                    "cargos_total_cuenta_aux",
                    "abonos_total_cuenta_aux",
                    "saldo_final_cuenta_aux",
                    "diferencia_vs_auxiliar",
                ]

                st.dataframe(
                    resumen_cuenta[cols_resumen].sort_values(
                        "saldo_neto", ascending=False
                    ),
                    use_container_width=True,
                )

                # Detalle de facturas pendientes (solo saldo > 0) por cuenta
                st.subheader("📄 Detalle de facturas pendientes por cuenta contable")

                df_pend_cuenta = df_cuenta_all[df_cuenta_all["saldo_factura"] > 0].copy()
                df_pend_cuenta = df_pend_cuenta.sort_values(
                    ["account_code", "account_name", "fecha_factura", "referencia"]
                )

                for (code, name), grp in df_pend_cuenta.groupby(
                    ["account_code", "account_name"], sort=False
                ):
                    total_cuenta = grp["saldo_factura"].sum()
                    num_facturas = grp["referencia"].nunique()

                    titulo = (
                        f"{code} - {name}  |  {num_facturas} facturas con saldo positivo  |  "
                        f"saldo pendiente ${total_cuenta:,.2f}"
                    )

                    with st.expander(titulo, expanded=False):
                        st.dataframe(
                            grp[
                                [
                                    "referencia",
                                    "fecha_factura",
                                    "saldo_factura",
                                ]
                            ].sort_values(["fecha_factura", "referencia"]),
                            use_container_width=True,
                        )

                # Descarga Excel (detalle por cuenta, solo pendientes, solo netos)
                xls_cuenta = to_excel(
                    df_pend_cuenta[
                        [
                            "account_code",
                            "account_name",
                            "referencia",
                            "fecha_factura",
                            "saldo_factura",
                        ]
                    ]
                )
                st.download_button(
                    label="⬇️ Descargar detalle por cuenta en Excel (solo pendientes, netos)",
                    data=xls_cuenta,
                    file_name="facturas_pendientes_por_cuenta.xlsx",
                    mime=(
                        "application/vnd.openxmlformats-officedocument."
                        "spreadsheetml.sheet"
                    ),
                )

        # ================================================================
        # TAB 3: Facturas cruzadas entre cuentas pendientes
        # ================================================================
        with tab_cruzadas:
            st.markdown("### Facturas cruzadas entre cuentas (pendientes)")
            st.caption(
                "Muestra solo facturas (referencias) que aparecen en **más de una cuenta contable** "
                "y que todavía tienen saldo pendiente (saldo neto > 0)."
            )

            df_tab3_base = filtrar_por_fecha(
                facturas_global_pend,
                fecha_desde=fecha_desde,
                fecha_hasta=fecha_hasta,
            )

            df_tab3 = df_tab3_base[df_tab3_base["cruza_cuentas"]].copy()

            if df_tab3.empty:
                st.info(
                    "No se encontraron facturas cruzadas pendientes "
                    "en este rango de fechas."
                )
            else:
                # Construir columna de cuenta principal (texto)
                df_tab3["cuenta_principal"] = (
                    df_tab3["account_code"].astype(str)
                    + " - "
                    + df_tab3["account_name"].astype(str)
                )

                # Calcular otras cuentas (todas menos la principal)
                def get_otras_cuentas(row):
                    if pd.isna(row["cuentas_involucradas"]):
                        return ""
                    cuentas = [c.strip() for c in str(row["cuentas_involucradas"]).split("|")]
                    principal = str(row["cuenta_principal"]).strip()
                    otras = [c for c in cuentas if c != principal]
                    return " | ".join(otras)

                df_tab3["otras_cuentas"] = df_tab3.apply(get_otras_cuentas, axis=1)

                # Filtro opcional por cuenta principal
                cuentas_principales = (
                    df_tab3["cuenta_principal"].dropna().sort_values().unique().tolist()
                )
                cuentas_sel_princ = st.multiselect(
                    "Filtrar por cuenta principal (opcional)",
                    options=cuentas_principales,
                    default=[],
                    key="cuentas_principales_cruzadas_pend",
                )
                if cuentas_sel_princ:
                    df_tab3 = df_tab3[
                        df_tab3["cuenta_principal"].isin(cuentas_sel_princ)
                    ]

                if df_tab3.empty:
                    st.info(
                        "No hay facturas cruzadas pendientes que cumplan con los filtros seleccionados."
                    )
                else:
                    st.subheader("📊 Resumen de facturas cruzadas pendientes")

                    c1, c2 = st.columns(2)
                    with c1:
                        st.metric(
                            "Facturas cruzadas pendientes",
                            value=int(df_tab3["referencia"].nunique()),
                        )
                    with c2:
                        st.metric(
                            "Saldo pendiente total (facturas cruzadas)",
                            value=f"${df_tab3['saldo_factura'].sum():,.2f}",
                        )

                    # Detalle de facturas cruzadas (solo netos)
                    st.subheader("📄 Detalle de facturas cruzadas pendientes")

                    cols_cruzadas = [
                        "referencia",
                        "fecha_factura",
                        "saldo_factura",
                        "cuenta_principal",
                        "otras_cuentas",
                        "num_cuentas",
                        "cuentas_involucradas",
                    ]

                    df_detalle_cruzadas = df_tab3[cols_cruzadas].sort_values(
                        ["fecha_factura", "referencia"]
                    )

                    st.dataframe(df_detalle_cruzadas, use_container_width=True)

                    # Descarga Excel
                    xls_cruzadas = to_excel(df_detalle_cruzadas)
                    st.download_button(
                        label="⬇️ Descargar detalle de facturas cruzadas pendientes en Excel",
                        data=xls_cruzadas,
                        file_name="facturas_cruzadas_pendientes.xlsx",
                        mime=(
                            "application/vnd.openxmlformats-officedocument."
                            "spreadsheetml.sheet"
                        ),
                    )

        # ================================================================
        # TAB 4: Facturas cruzadas entre cuentas pagadas
        # ================================================================
        with tab_cruzadas_pagadas:
            st.markdown("### Facturas cruzadas entre cuentas (pagadas)")
            st.caption(
                "Muestra facturas cruzadas entre cuentas cuya integración de cargos y abonos "
                "da como resultado **saldo neto 0** (consideradas pagadas)."
            )

            df_tab4_base = filtrar_por_fecha(
                facturas_global,
                fecha_desde=fecha_desde,
                fecha_hasta=fecha_hasta,
            )

            # Definimos 'pagadas' como saldo_factura == 0; si quieres incluir saldos a favor, cambia a <= 0
            df_tab4 = df_tab4_base[
                (df_tab4_base["cruza_cuentas"]) & (df_tab4_base["saldo_factura"] == 0)
            ].copy()

            if df_tab4.empty:
                st.info(
                    "No se encontraron facturas cruzadas pagadas "
                    "en este rango de fechas."
                )
            else:
                # Cuenta principal (texto)
                df_tab4["cuenta_principal"] = (
                    df_tab4["account_code"].astype(str)
                    + " - "
                    + df_tab4["account_name"].astype(str)
                )

                # Otras cuentas
                def get_otras_cuentas_pag(row):
                    if pd.isna(row["cuentas_involucradas"]):
                        return ""
                    cuentas = [c.strip() for c in str(row["cuentas_involucradas"]).split("|")]
                    principal = str(row["cuenta_principal"]).strip()
                    otras = [c for c in cuentas if c != principal]
                    return " | ".join(otras)

                df_tab4["otras_cuentas"] = df_tab4.apply(get_otras_cuentas_pag, axis=1)

                # Filtro opcional por cuenta principal
                cuentas_principales_pag = (
                    df_tab4["cuenta_principal"].dropna().sort_values().unique().tolist()
                )
                cuentas_sel_princ_pag = st.multiselect(
                    "Filtrar por cuenta principal (opcional)",
                    options=cuentas_principales_pag,
                    default=[],
                    key="cuentas_principales_cruzadas_pag",
                )
                if cuentas_sel_princ_pag:
                    df_tab4 = df_tab4[
                        df_tab4["cuenta_principal"].isin(cuentas_sel_princ_pag)
                    ]

                if df_tab4.empty:
                    st.info(
                        "No hay facturas cruzadas pagadas que cumplan con los filtros seleccionados."
                    )
                else:
                    st.subheader("📊 Resumen de facturas cruzadas pagadas")

                    c1, c2 = st.columns(2)
                    with c1:
                        st.metric(
                            "Facturas cruzadas pagadas",
                            value=int(df_tab4["referencia"].nunique()),
                        )
                    with c2:
                        st.metric(
                            "Saldo total (debería ser 0)",
                            value=f"${df_tab4['saldo_factura'].sum():,.2f}",
                        )

                    # Detalle (solo netos)
                    st.subheader("📄 Detalle de facturas cruzadas pagadas")

                    cols_cruzadas_pag = [
                        "referencia",
                        "fecha_factura",
                        "saldo_factura",
                        "cuenta_principal",
                        "otras_cuentas",
                        "num_cuentas",
                        "cuentas_involucradas",
                    ]

                    df_detalle_cruzadas_pag = df_tab4[cols_cruzadas_pag].sort_values(
                        ["fecha_factura", "referencia"]
                    )

                    st.dataframe(df_detalle_cruzadas_pag, use_container_width=True)

                    # Descarga
                    xls_cruzadas_pag = to_excel(df_detalle_cruzadas_pag)
                    st.download_button(
                        label="⬇️ Descargar detalle de facturas cruzadas pagadas en Excel",
                        data=xls_cruzadas_pag,
                        file_name="facturas_cruzadas_pagadas.xlsx",
                        mime=(
                            "application/vnd.openxmlformats-officedocument."
                            "spreadsheetml.sheet"
                        ),
                    )
