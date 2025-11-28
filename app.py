import streamlit as st
import pandas as pd
import numpy as np
import re
import math
from io import BytesIO
from datetime import date

# Umbral para considerar "casi cero" el saldo neto y marcar solo_saldo_inicial
UMBRAL_SALDO_INICIAL = 1.0  # en pesos; ajusta si quieres mayor precisión

st.set_page_config(page_title="Facturas no pagadas", layout="wide")

st.title("🔍 Auditoría Integración de Saldos")
st.write(
    "Sube el archivo de **Movimientos, Auxiliares del Catálogo** generado desde CONTPAQ i. "
    "La app calcula salos **netos** por factura y por cuenta, y los compara contra el auxiliar."
)

# --------------------------------------------------------------------
# Cuadro de ayuda para auditoría (versión simplificada)
# --------------------------------------------------------------------
with st.expander("ℹ️ Cómo leer este reporte (ayuda rápida para auditoría)", expanded=False):
    st.markdown(
        """
**1. Qué hace la app**

- Lee el reporte de **Movimientos, Auxiliares del Catálogo**.
- Toma solo las filas que:
  - Tienen **fecha** válida (columna 0, formato `dd/Mon/aaaa`), y  
  - Tienen valor en **Referencia** (número de factura).
- Para cada referencia calcula el **saldo neto**:

> `saldo_factura = cargos_total – abonos_total`

Interpretación:
- `saldo_factura > 0`  → factura con saldo pendiente.  
- `saldo_factura = 0`  → factura saldada.  
- `saldo_factura < 0`  → saldo a favor / crédito.

Los movimientos **sin referencia** no se ven factura por factura, pero sí están absorbidos en el **saldo final del auxiliar**, que usamos para conciliar.

---

**2. Nivel global (cartera completa)**

- **Saldo neto facturas (C-A, solo con referencia)**: suma de todos los saldos netos de facturas (positivos y negativos).
- **Saldo final cartera (auxiliar – Total Clientes)**: lo que dice el reporte en “Total Clientes :”.
- La diferencia entre ambos = **saldo de cuentas sin facturas + movimientos sin referencia** (y/o fuera de rango de fechas).

---

**3. Pestaña “Resumen por cuenta vs auxiliar”**

Por cada cuenta (cliente) muestra:

- `saldo_neto`: suma neta de saldos de sus facturas (solo las del rango de fechas).
- `saldo_final_cuenta_aux`: saldo final de la fila **“Total:”** del auxiliar para esa cuenta.
- `diferencia_vs_auxiliar` = saldo_final_cuenta_aux – saldo_neto.
- `saldo_no_explicado_por_facturas` = **igual que diferencia_vs_auxiliar**, solo con nombre más obvio para auditor.
- `solo_saldo_inicial` = `True` cuando:
  - El saldo neto por facturas es casi 0 (`abs(saldo_neto) < UMBRAL`), y
  - La diferencia con el auxiliar es significativa (`abs(diferencia_vs_auxiliar) > UMBRAL`).

Si `solo_saldo_inicial = True`, la lectura es:

> “Esta cuenta tiene saldo en el auxiliar que **no proviene de facturas vigentes**, sino de saldo inicial u otros movimientos sin referencia.”

---

**4. Pestaña “Facturas pendientes”**

- Si no filtras por cuenta: muestra facturas a nivel **global** (una fila por referencia).
- Si filtras por una cuenta en el combo: muestra facturas a nivel **de esa cuenta**, alineadas con el auxiliar.
- Lista solo facturas con `saldo_factura > 0` (detalle).
- La métrica de **saldo** es **neta**: suma de todos los saldos por referencia en el rango (positivos y negativos).

---

**5. Pestaña “Facturas con saldo a favor”**

- Lista facturas con `saldo_factura < 0` (notas de crédito, saldos a favor, etc.).
- También respeta el filtro de cuenta: global o por cuenta contable.

---

**6. Pestaña “Referencias cruzadas entre cuentas”**

- Muestra referencias (números de factura) que aparecen en **más de una cuenta contable** en el periodo.
- Indica:
  - Cuántas cuentas están involucradas.
  - El **saldo global** de la referencia (sumando todas las cuentas).
  - Si el saldo global está prácticamente en cero (`es_saldo_global_cero = True`), pero sigue generando saldos distintos por cuenta.
- Útil para detectar **pagos aplicados en cuentas equivocadas** o necesidad de reclasificación contable.

---

**7. Importante**

- Los **saldos del auxiliar** siempre mandan como referencia final.
- La app es una herramienta de **análisis**:
  - Para ver facturas pendientes.
  - Para conciliar saldo por cuenta vs auxiliar.
  - Para detectar cuentas con “solo saldo inicial / ajustes” (`solo_saldo_inicial = True`).
  - Para identificar referencias cruzadas entre cuentas y armar listas de trabajo para contabilidad.
        """
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
      - movs_valid: movimientos con referencia (y cuenta asignada).
      - resumen_auxiliar: totales globales (netos + saldo final auxiliar).
      - totales_cuentas_aux: totales por cuenta (Total: del auxiliar).
    """
    raw = pd.read_excel(file, header=None)

    # Detectar filas cabecera de cuenta (código de cuenta + 'Saldo inicial :')
    account_pattern = re.compile(r"^\d{3}-\d{3}-\d{3}-\d{3}$")
    is_account_header = raw[0].astype(str).str.match(account_pattern) & raw[
        6
    ].astype(str).str.contains("Saldo inicial", na=False)

    df = raw.copy()
    df["account_code"] = np.where(is_account_header, df[0], np.nan)
    df["account_name"] = np.where(is_account_header, df[1], np.nan)
    df["account_code"] = df["account_code"].ffill()
    df["account_name"] = df["account_name"].ffill()

    # Filas de movimientos (columna 0 es fecha dd/Mon/aaaa)
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

    # Limpiar referencia
    movs["referencia"] = movs["referencia"].astype(str).str.strip()
    movs["referencia"] = movs["referencia"].replace({"nan": np.nan, "": np.nan})

    # Convertir fecha
    movs["fecha"] = movs["fecha_raw"].apply(parse_spanish_date)

    # Solo movimientos con referencia
    movs_valid = movs[movs["referencia"].notna()].copy()

    # Totales globales del auxiliar ("Total Clientes :")
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
        "saldo_inicial_implicito": float(saldo_inicial_implicito)
        if not math.isnan(saldo_inicial_implicito)
        else None,
    }

    # Totales por cuenta (filas "Total:" en col 4)
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
        )[[
            "account_code",
            "account_name",
            "cargos_total_cuenta_aux",
            "abonos_total_cuenta_aux",
            "saldo_final_cuenta_aux",
        ]]
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
    Facturas a nivel global (por referencia), cruzando todas las cuentas.
    Asigna una cuenta principal (normalmente donde está el cargo).
    """
    facturas = (
        movs_valid.groupby("referencia")
        .agg(
            fecha_factura=("fecha", "min"),
            cargos_total=("cargos", "sum"),
            abonos_total=("abonos", "sum"),
        )
        .reset_index()
    )

    movs_valid2 = movs_valid.copy()
    movs_valid2["es_cargo_pos"] = movs_valid2["cargos"] > 0

    main_from_cargo = (
        movs_valid2[movs_valid2["es_cargo_pos"]]
        .sort_values(["referencia", "cargos"], ascending=[True, False])
        .drop_duplicates("referencia")[["referencia", "account_code", "account_name"]]
    )

    main_any = (
        movs_valid2.sort_values(["referencia", "fecha"])
        .drop_duplicates("referencia")[["referencia", "account_code", "account_name"]]
    )

    main_account = pd.concat([main_from_cargo, main_any], ignore_index=True)
    main_account = main_account.drop_duplicates("referencia", keep="first")

    facturas = facturas.merge(main_account, on="referencia", how="left")

    # Suma neta por referencia (global)
    facturas["saldo_factura"] = facturas["cargos_total"] - facturas["abonos_total"]

    # Texto de cuenta principal
    facturas["cuenta"] = (
        facturas["account_code"].astype(str)
        + " - "
        + facturas["account_name"].astype(str)
    )

    return facturas


@st.cache_data
def construir_facturas_por_cuenta(movs_valid: pd.DataFrame) -> pd.DataFrame:
    """
    Facturas por cuenta contable (sin cruzar cuentas).
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

    facturas["cuenta"] = (
        facturas["account_code"].astype(str)
        + " - "
        + facturas["account_name"].astypestr()
    )

    return facturas


def filtrar_por_fecha(df: pd.DataFrame, fecha_desde: date, fecha_hasta: date) -> pd.DataFrame:
    """Filtra un DataFrame por columna fecha_factura."""
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

        # ---- Cálculo de cuentas del auxiliar que NO tienen ninguna factura ----
        cuentas_con_facturas = (
            facturas_cuenta[["account_code", "account_name"]]
            .drop_duplicates()
        )

        aux_sin_facturas = totales_cuentas_aux.merge(
            cuentas_con_facturas,
            on=["account_code", "account_name"],
            how="left",
            indicator=True,
        )

        saldo_cuentas_sin_facturas = aux_sin_facturas.loc[
            aux_sin_facturas["_merge"] == "left_only", "saldo_final_cuenta_aux"
        ].sum()

    if facturas_global.empty and facturas_cuenta.empty:
        st.success("✅ No se encontraron facturas en el archivo.")
    else:
        # ------------------------- Resumen global (netos + cuentas sin facturas) -------------------------
        st.subheader("📊 Resumen global vs auxiliar (netos)")

        colg1, colg2, colg3, colg4 = st.columns(4)
        with colg1:
            st.metric(
                "Saldo neto facturas (C-A, solo con referencia)",
                value=f"${resumen_aux['saldo_neto_movs']:,.2f}",
            )
        with colg2:
            st.metric(
                "Saldo cuentas sin facturas (según auxiliar)",
                value=f"${saldo_cuentas_sin_facturas:,.2f}",
            )
        with colg3:
            if resumen_aux.get("saldo_final_aux") is not None:
                st.metric(
                    "Saldo final cartera (auxiliar – Total Clientes)",
                    value=f"${resumen_aux['saldo_final_aux']:,.2f}",
                )
            else:
                st.metric("Saldo final cartera (auxiliar)", value="N/D")
        with colg4:
            if resumen_aux.get("saldo_final_aux") is not None:
                conciliado = resumen_aux["saldo_neto_movs"] + saldo_cuentas_sin_facturas
                diferencia_residual = resumen_aux["saldo_final_aux"] - conciliado
                st.metric(
                    "Diferencia residual",
                    value=f"${diferencia_residual:,.2f}",
                )
            else:
                st.metric("Diferencia residual", value="N/D")

        if resumen_aux.get("saldo_inicial_implicito") is not None:
            st.caption(
                f"Saldo inicial implícito según auxiliar: "
                f"${resumen_aux['saldo_inicial_implicito']:,.2f} "
                f"(Saldo final auxiliar – saldo neto de movimientos con referencia)."
            )

        st.caption(
            "- **Saldo neto facturas**: solo movimientos con referencia (lo que se analiza factura por factura).\n"
            "- **Saldo cuentas sin facturas**: saldos de cuentas que no tienen ninguna factura en el periodo "
            "(por ejemplo, cuentas con solo saldo inicial).\n"
            "- La suma de ambos debe aproximarse al **Saldo final cartera (Total Clientes)** del auxiliar; "
            "cualquier diferencia residual corresponde a redondeos o particularidades del archivo."
        )

        # ------------------------- Filtros globales -------------------------
        all_fechas = pd.concat(
            [
                facturas_global["fecha_factura"],
                facturas_cuenta["fecha_factura"],
            ]
        ).dropna()

        min_date = all_fechas.min()
        max_date = all_fechas.max()

        st.subheader("📌 Filtros")

        col_f1, col_f2, col_f3 = st.columns([1, 1, 2])
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
        with col_f3:
            # Lista de cuentas para filtrar (opcional)
            todas_cuentas = (
                facturas_cuenta["cuenta"].dropna().sort_values().unique().tolist()
            )
            opciones_cuentas = ["(Todas las cuentas)"] + todas_cuentas
            cuenta_seleccionada = st.selectbox(
                "Cuenta contable (opcional)",
                options=opciones_cuentas,
                index=0,
            )

        # Código de cuenta seleccionada (si aplica)
        codigo_cuenta_seleccionada = None
        if cuenta_seleccionada != "(Todas las cuentas)":
            codigo_cuenta_seleccionada = cuenta_seleccionada.split(" - ")[0].strip()

        # Aplica filtro de fechas a las dos vistas base
        facturas_global_f = filtrar_por_fecha(facturas_global, fecha_desde, fecha_hasta)
        facturas_cuenta_f = filtrar_por_fecha(facturas_cuenta, fecha_desde, fecha_hasta)

        # ----------------------------------------------------------------
        # Pestañas
        # ----------------------------------------------------------------
        tab_resumen, tab_pendientes, tab_favor, tab_cruzadas = st.tabs(
            [
                "📂 Resumen por cuenta vs auxiliar",
                "📑 Facturas pendientes",
                "💳 Facturas con saldo a favor",
                "🔀 Referencias cruzadas entre cuentas",
            ]
        )

        # ================================================================
        # TAB 1: Resumen por cuenta vs auxiliar
        # ================================================================
        with tab_resumen:
            st.markdown("### Resumen por cuenta contable vs auxiliar")

            # Filtramos los totales del auxiliar por cuenta (si se seleccionó una)
            tot_aux_f = totales_cuentas_aux.copy()
            if codigo_cuenta_seleccionada is not None:
                tot_aux_f = tot_aux_f[
                    tot_aux_f["account_code"] == codigo_cuenta_seleccionada
                ]

            if tot_aux_f.empty:
                st.info("No hay cuentas en el auxiliar para mostrar con los filtros actuales.")
            else:
                # Métricas por cuenta a partir de las facturas (puede estar vacío si no hay facturas en rango)
                if not facturas_cuenta_f.empty:
                    metrics = (
                        facturas_cuenta_f.groupby(["account_code", "account_name"])
                        .agg(
                            saldo_neto=("saldo_factura", "sum"),
                            facturas_positivas=("saldo_factura", lambda s: int((s > 0).sum())),
                            referencias_negativas=("saldo_factura", lambda s: int((s < 0).sum())),
                        )
                        .reset_index()
                    )
                else:
                    metrics = pd.DataFrame(
                        columns=[
                            "account_code",
                            "account_name",
                            "saldo_neto",
                            "facturas_positivas",
                            "referencias_negativas",
                        ]
                    )

                # Unimos: partimos SIEMPRE del auxiliar, y pegamos los saldos por facturas
                resumen_cuenta = tot_aux_f.merge(
                    metrics,
                    on=["account_code", "account_name"],
                    how="left",
                )

                # Para cuentas sin ninguna factura en el rango, llenamos con ceros
                for col in ["saldo_neto", "facturas_positivas", "referencias_negativas"]:
                    resumen_cuenta[col] = resumen_cuenta[col].fillna(0)

                resumen_cuenta["diferencia_vs_auxiliar"] = (
                    resumen_cuenta["saldo_final_cuenta_aux"] - resumen_cuenta["saldo_neto"]
                )

                resumen_cuenta["saldo_no_explicado_por_facturas"] = resumen_cuenta[
                    "diferencia_vs_auxiliar"
                ]

                resumen_cuenta["solo_saldo_inicial"] = (
                    resumen_cuenta["saldo_neto"].abs() < UMBRAL_SALDO_INICIAL
                ) & (
                    resumen_cuenta["diferencia_vs_auxiliar"].abs()
                    > UMBRAL_SALDO_INICIAL
                )

                # Métrica global de saldo neto en rango (solo de lo que viene de facturas)
                saldo_neto_total = resumen_cuenta["saldo_neto"].sum()
                st.metric(
                    "Saldo neto total por referencia (en rango y filtro, solo facturas)",
                    value=f"${saldo_neto_total:,.2f}",
                )

                st.caption(
                    "La columna **saldo_no_explicado_por_facturas** muestra la diferencia entre "
                    "el saldo final del auxiliar y el saldo neto de facturas en el rango. "
                    "Si **solo_saldo_inicial** es True, la cuenta tiene saldo en auxiliar que no proviene "
                    "de facturas vigentes (saldo inicial / otros movimientos sin referencia)."
                )

                cols_resumen = [
                    "account_code",
                    "account_name",
                    "facturas_positivas",
                    "referencias_negativas",
                    "saldo_neto",
                    "saldo_final_cuenta_aux",
                    "saldo_no_explicado_por_facturas",
                    "solo_saldo_inicial",
                ]

                st.dataframe(
                    resumen_cuenta[cols_resumen].sort_values(
                        "saldo_neto", ascending=False
                    ),
                    use_container_width=True,
                )

        # ================================================================
        # TAB 2: Facturas pendientes (NETO, no bruto)
        # ================================================================
        with tab_pendientes:
            st.markdown("### Facturas pendientes (detalle) y saldo neto por referencia")

            # Base: global o por cuenta, según el filtro
            if codigo_cuenta_seleccionada is not None:
                base_df = facturas_cuenta_f[
                    facturas_cuenta_f["account_code"] == codigo_cuenta_seleccionada
                ].copy()
            else:
                base_df = facturas_global_f.copy()

            # Facturas con saldo neto > 0 (detalle)
            df_pend = base_df[base_df["saldo_factura"] > 0].copy()

            # Saldo neto de TODAS las referencias (positivas y negativas) en el rango/filtro
            saldo_neto_referencias = base_df["saldo_factura"].sum()

            if df_pend.empty:
                st.info("No hay facturas pendientes (saldo neto > 0) con estos filtros.")
            else:
                total_facturas = df_pend["referencia"].nunique()

                c1, c2 = st.columns(2)
                with c1:
                    st.metric(
                        "Número de facturas pendientes (detalle)",
                        value=int(total_facturas),
                    )
                with c2:
                    st.metric(
                        "Saldo neto por referencia (en este rango y filtros)",
                        value=f"${saldo_neto_referencias:,.2f}",
                    )

                st.caption(
                    "El saldo mostrado es **neto**: suma de todos los saldos por referencia "
                    "(facturas con saldo positivo y facturas con saldo a favor) "
                    "en el rango de fechas y cuenta(s) seleccionados. "
                    "La tabla de abajo lista solo las facturas con saldo > 0 (pendientes)."
                )

                cols_detalle = [
                    "referencia",
                    "fecha_factura",
                    "saldo_factura",
                    "account_code",
                    "account_name",
                ]

                df_detalle = df_pend[cols_detalle].sort_values(
                    ["account_code", "fecha_factura", "referencia"]
                )

                st.dataframe(df_detalle, use_container_width=True)

                xls_pend = to_excel(df_detalle)
                st.download_button(
                    label="⬇️ Descargar facturas pendientes en Excel",
                    data=xls_pend,
                    file_name="facturas_pendientes.xlsx",
                    mime=(
                        "application/vnd.openxmlformats-officedocument."
                        "spreadsheetml.sheet"
                    ),
                )

        # ================================================================
        # TAB 3: Facturas con saldo a favor
        # ================================================================
        with tab_favor:
            st.markdown("### Facturas con saldo a favor (saldo neto < 0)")

            # Base: global o por cuenta, según el filtro
            if codigo_cuenta_seleccionada is not None:
                base_df_f = facturas_cuenta_f[
                    facturas_cuenta_f["account_code"] == codigo_cuenta_seleccionada
                ].copy()
            else:
                base_df_f = facturas_global_f.copy()

            df_favor = base_df_f[base_df_f["saldo_factura"] < 0].copy()

            if df_favor.empty:
                st.info("No hay facturas con saldo a favor (saldo neto < 0) con estos filtros.")
            else:
                total_refs = df_favor["referencia"].nunique()
                saldo_total_favor = df_favor["saldo_factura"].sum()

                c1, c2 = st.columns(2)
                with c1:
                    st.metric(
                        "Número de referencias con saldo a favor",
                        value=int(total_refs),
                    )
                with c2:
                    st.metric(
                        "Saldo total a favor (neto, suele ser negativo)",
                        value=f"${saldo_total_favor:,.2f}",
                    )

                cols_detalle_f = [
                    "referencia",
                    "fecha_factura",
                    "saldo_factura",
                    "account_code",
                    "account_name",
                ]

                df_detalle_f = df_favor[cols_detalle_f].sort_values(
                    ["account_code", "fecha_factura", "referencia"]
                )

                st.dataframe(df_detalle_f, use_container_width=True)

                xls_favor = to_excel(df_detalle_f)
                st.download_button(
                    label="⬇️ Descargar facturas con saldo a favor en Excel",
                    data=xls_favor,
                    file_name="facturas_saldo_a_favor.xlsx",
                    mime=(
                        "application/vnd.openxmlformats-officedocument."
                        "spreadsheetml.sheet"
                    ),
                )

        # ================================================================
        # TAB 4: Referencias cruzadas entre cuentas
        # ================================================================
        with tab_cruzadas:
            st.markdown("### Referencias cruzadas entre cuentas")

            if facturas_cuenta_f.empty:
                st.info("No hay movimientos con referencia en este rango de fechas.")
            else:
                # Resumen por referencia (en TODO el universo de cuentas del periodo)
                ref_summary = (
                    facturas_cuenta_f.groupby("referencia")
                    .agg(
                        fecha_min_referencia=("fecha_factura", "min"),
                        cuentas_involucradas=("account_code", "nunique"),
                        saldo_global=("saldo_factura", "sum"),
                    )
                    .reset_index()
                )

                # Bandera: referencia globalmente saldada (≈0)
                ref_summary["es_saldo_global_cero"] = (
                    ref_summary["saldo_global"].abs() < UMBRAL_SALDO_INICIAL
                )

                # Solo referencias que aparecen en más de una cuenta
                cross_refs = ref_summary[ref_summary["cuentas_involucradas"] > 1].copy()

                if cross_refs.empty:
                    st.info(
                        "No se encontraron referencias cruzadas entre cuentas con los filtros de fechas actuales."
                    )
                else:
                    # Unimos contra el detalle por cuenta
                    df_cross = facturas_cuenta_f.merge(
                        cross_refs[
                            [
                                "referencia",
                                "fecha_min_referencia",
                                "cuentas_involucradas",
                                "saldo_global",
                                "es_saldo_global_cero",
                            ]
                        ],
                        on="referencia",
                        how="inner",
                    )

                    # Si el usuario filtró una cuenta, nos quedamos solo con esa cuenta,
                    # pero conservando la información global de la referencia.
                    if codigo_cuenta_seleccionada is not None:
                        df_cross = df_cross[
                            df_cross["account_code"] == codigo_cuenta_seleccionada
                        ]

                    if df_cross.empty:
                        st.info(
                            "No hay referencias cruzadas que involucren la cuenta seleccionada con estos filtros."
                        )
                    else:
                        # Clasificamos el tipo de saldo por cuenta
                        df_cross["tipo_saldo_cuenta"] = np.where(
                            df_cross["saldo_factura"] > 0,
                            "Saldo pendiente",
                            np.where(
                                df_cross["saldo_factura"] < 0,
                                "Saldo a favor",
                                "Saldo cero",
                            ),
                        )

                        # Métricas a nivel referencia (sin duplicar por cuenta)
                        total_refs_cruzadas = cross_refs["referencia"].nunique()
                        total_refs_saldo_cero = cross_refs[
                            cross_refs["es_saldo_global_cero"]
                        ]["referencia"].nunique()

                        c1, c2 = st.columns(2)
                        with c1:
                            st.metric(
                                "Número de referencias cruzadas (más de una cuenta)",
                                value=int(total_refs_cruzadas),
                            )
                        with c2:
                            st.metric(
                                "Referencias cruzadas con saldo global ≈ 0",
                                value=int(total_refs_saldo_cero),
                            )

                        st.caption(
                            "- **cuentas_involucradas**: cuántas cuentas contables tienen movimientos con esa referencia.\n"
                            "- **saldo_global**: suma del saldo de la referencia en todas las cuentas (si ≈0, globalmente está saldada).\n"
                            "- **es_saldo_global_cero = True** suele indicar un pago aplicado en otra cuenta distinta al cliente de origen."
                        )

                        cols_cross = [
                            "referencia",
                            "fecha_min_referencia",
                            "cuentas_involucradas",
                            "saldo_global",
                            "es_saldo_global_cero",
                            "account_code",
                            "account_name",
                            "saldo_factura",
                            "tipo_saldo_cuenta",
                        ]

                        df_cross_show = df_cross[cols_cross].sort_values(
                            ["referencia", "account_code", "fecha_min_referencia"]
                        )

                        st.dataframe(df_cross_show, use_container_width=True)

                        xls_cross = to_excel(df_cross_show)
                        st.download_button(
                            label="⬇️ Descargar referencias cruzadas en Excel",
                            data=xls_cross,
                            file_name="referencias_cruzadas_entre_cuentas.xlsx",
                            mime=(
                                "application/vnd.openxmlformats-officedocument."
                                "spreadsheetml.sheet"
                            ),
                        )
