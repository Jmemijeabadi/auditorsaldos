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

# --------------------------------------------------------------------
# Encabezado principal
# --------------------------------------------------------------------
st.title("🔍 Auditoría Integración de Saldos")
st.write(
    "Sube el archivo de **Movimientos, Auxiliares del Catálogo** generado desde CONTPAQ i. "
    "La app calcula saldos **netos** por factura y por cuenta, y los compara contra el auxiliar."
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
- La métrica de **saldo** es:
  - **“Saldo pendiente total”** = suma de saldos de facturas con saldo > 0.
  - **“Saldo neto total (según auxiliar)”** = saldo final real de la cuenta o cartera completa.

---

**5. Pestaña “Facturas con saldo a favor”**

- Lista facturas con `saldo_factura < 0` (notas de crédito, saldos a favor, etc.).
- También respeta el filtro de cuenta: global o por cuenta contable.

---

**6. Importante**

- Los **saldos del auxiliar** siempre mandan como referencia final.
- La app es una herramienta de **análisis**:
  - Para ver facturas pendientes.
  - Para conciliar saldo por cuenta vs auxiliar.
  - Para detectar cuentas con “solo saldo inicial / ajustes” (`solo_saldo_inicial = True`).
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

    # Suma neta por referencia
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
        + facturas["account_name"].astype(str)
    )

    return facturas


# --------------------------------------------------------------------
# NUEVA UTILIDAD: detectar cruces de referencias entre cuentas
# --------------------------------------------------------------------
@st.cache_data
def detectar_cruces_referencias(movs_valid: pd.DataFrame):
    """
    Detecta referencias (facturas) que aparecen en más de una cuenta contable
    y que tienen cargos en alguna cuenta y abonos en otra.

    Regresa:
      - detalle_cruces: nivel cuenta + referencia.
      - resumen_cruces: nivel referencia (global).
    """
    df = movs_valid.copy()

    # Agrupamos por referencia y cuenta
    por_cuenta = (
        df.groupby(["referencia", "account_code", "account_name"])
        .agg(
            cargos_total=("cargos", "sum"),
            abonos_total=("abonos", "sum"),
        )
        .reset_index()
    )

    por_cuenta["tiene_cargo"] = por_cuenta["cargos_total"] > 0
    por_cuenta["tiene_abono"] = por_cuenta["abonos_total"] > 0

    # Nivel referencia (global)
    ref_level = (
        por_cuenta.groupby("referencia")
        .agg(
            num_cuentas=("account_code", "nunique"),
            cuentas_con_cargo=("tiene_cargo", "sum"),
            cuentas_con_abono=("tiene_abono", "sum"),
            cargos_tot_ref=("cargos_total", "sum"),
            abonos_tot_ref=("abonos_total", "sum"),
        )
        .reset_index()
    )

    ref_level["saldo_neto_ref"] = ref_level["cargos_tot_ref"] - ref_level["abonos_tot_ref"]

    # Definición de "cruce":
    # - la referencia aparece en más de una cuenta, y
    # - hay al menos una cuenta con cargo y otra con abono
    ref_level["es_cruce"] = (
        (ref_level["num_cuentas"] > 1)
        & (ref_level["cuentas_con_cargo"] >= 1)
        & (ref_level["cuentas_con_abono"] >= 1)
    )

    resumen_cruces = ref_level[ref_level["es_cruce"]].copy()

    # Traemos el detalle por cuenta solo para esas referencias
    detalle_cruces = por_cuenta.merge(
        resumen_cruces[
            [
                "referencia",
                "num_cuentas",
                "cargos_tot_ref",
                "abonos_tot_ref",
                "saldo_neto_ref",
            ]
        ],
        on="referencia",
        how="inner",
    )

    detalle_cruces["saldo_por_cuenta"] = (
        detalle_cruces["cargos_total"] - detalle_cruces["abonos_total"]
    )

    return detalle_cruces, resumen_cruces


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
# App (layout tipo dashboard: sidebar = opciones, main = resultados)
# --------------------------------------------------------------------

# LADO IZQUIERDO: SIDEBAR (archivo + filtros)
st.sidebar.markdown("## ⚙️ Configuración")

uploaded_file = st.sidebar.file_uploader(
    "📎 Sube el archivo Excel de movimientos (auxiliares del catálogo)",
    type=["xlsx"],
)

if uploaded_file is None:
    st.sidebar.info("Sube un archivo `.xlsx` para comenzar.")
    st.info(
        "Sube un archivo `.xlsx` exportado desde CONTPAQ "
        "(Movimientos, Auxiliares del Catálogo) para comenzar."
    )
else:
    with st.spinner("Procesando archivo..."):
        movs_valid, resumen_aux, totales_cuentas_aux = procesar_movimientos(uploaded_file)
        facturas_global = construir_facturas_global(movs_valid)
        facturas_cuenta = construir_facturas_por_cuenta(movs_valid)

        # ---- Detección de referencias cruzadas entre cuentas ----
        detalle_cruces, resumen_cruces = detectar_cruces_referencias(movs_valid)

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

        # ------------------------- Filtros (en sidebar) -------------------------
        # Determinar rango de fechas disponible
        all_fechas = pd.concat(
            [
                facturas_global["fecha_factura"],
                facturas_cuenta["fecha_factura"],
            ]
        ).dropna()

        if all_fechas.empty:
            min_date = pd.to_datetime(date.today())
            max_date = pd.to_datetime(date.today())
        else:
            min_date = all_fechas.min()
            max_date = all_fechas.max()

        st.sidebar.markdown("---")
        st.sidebar.markdown("## 📌 Filtros")

        fecha_desde = st.sidebar.date_input(
            "Fecha desde",
            value=min_date.date() if pd.notna(min_date) else date.today(),
            min_value=min_date.date() if pd.notna(min_date) else date(2000, 1, 1),
            max_value=max_date.date() if pd.notna(max_date) else date.today(),
            key="fecha_desde",
        )

        fecha_hasta = st.sidebar.date_input(
            "Fecha hasta",
            value=max_date.date() if pd.notna(max_date) else date.today(),
            min_value=min_date.date() if pd.notna(min_date) else date(2000, 1, 1),
            max_value=max_date.date() if pd.notna(max_date) else date.today(),
            key="fecha_hasta",
        )

        # Lista de cuentas para filtrar (opcional)
        todas_cuentas = (
            facturas_cuenta["cuenta"].dropna().sort_values().unique().tolist()
        )
        opciones_cuentas = ["(Todas las cuentas)"] + todas_cuentas
        cuenta_seleccionada = st.sidebar.selectbox(
            "Cuenta contable (opcional)",
            options=opciones_cuentas,
            index=0,
            key="cuenta_seleccionada",
        )

        # Código de cuenta seleccionada (si aplica)
        codigo_cuenta_seleccionada = None
        if cuenta_seleccionada != "(Todas las cuentas)":
            codigo_cuenta_seleccionada = cuenta_seleccionada.split(" - ")[0].strip()

        # ------------------------- Aplicar filtros de fecha -------------------------
        facturas_global_f = filtrar_por_fecha(facturas_global, fecha_desde, fecha_hasta)
        facturas_cuenta_f = filtrar_por_fecha(facturas_cuenta, fecha_desde, fecha_hasta)

    # LADO DERECHO: MAIN (KPIs + tabs + cruces)
    if facturas_global.empty and facturas_cuenta.empty:
        st.success("✅ No se encontraron facturas en el archivo.")
    else:
        # ------------------------- Resumen global (netos + inicial + cuentas sin facturas) -------------------------
        st.subheader("📊 Resumen global vs auxiliar (netos e inicial)")

        colg1, colg2, colg3, colg4, colg5 = st.columns(5)

        # 1) Saldo neto de movimientos del periodo (solo con referencia)
        with colg1:
            st.metric(
                "Saldo neto de movimientos del periodo (C-A, solo con referencia)",
                value=f"${resumen_aux['saldo_neto_movs']:,.2f}",
            )

        # 2) Saldo inicial global de cartera (implícito según auxiliar)
        with colg2:
            if resumen_aux.get("saldo_inicial_implicito") is not None:
                st.metric(
                    "Saldo inicial global de cartera (implícito según auxiliar)",
                    value=f"${resumen_aux['saldo_inicial_implicito']:,.2f}",
                )
            else:
                st.metric(
                    "Saldo inicial global de cartera (implícito según auxiliar)",
                    value="N/D",
                )

        # 3) Saldos de cuentas sin facturas
        with colg3:
            st.metric(
                "Saldo cuentas sin facturas (según auxiliar)",
                value=f"${saldo_cuentas_sin_facturas:,.2f}",
            )

        # 4) Saldo final total de cartera (Total Clientes)
        with colg4:
            if resumen_aux.get("saldo_final_aux") is not None:
                st.metric(
                    "Saldo final cartera (auxiliar – 'Total Clientes')",
                    value=f"${resumen_aux['saldo_final_aux']:,.2f}",
                )
            else:
                st.metric("Saldo final cartera (auxiliar)", value="N/D")

        # 5) Diferencia residual vs auxiliar
        with colg5:
            if resumen_aux.get("saldo_final_aux") is not None:
                conciliado = resumen_aux["saldo_neto_movs"] + saldo_cuentas_sin_facturas
                diferencia_residual = resumen_aux["saldo_final_aux"] - conciliado
                st.metric(
                    "Diferencia residual vs auxiliar",
                    value=f"${diferencia_residual:,.2f}",
                )
            else:
                st.metric("Diferencia residual vs auxiliar", value="N/D")

        st.caption(
            "- **Saldo neto de movimientos del periodo**: suma de cargos menos abonos de todos los movimientos "
            "con referencia (facturas y notas) que vienen en el archivo.\n"
            "- **Saldo inicial global de cartera (implícito)**: diferencia entre el saldo final de cartera del auxiliar "
            "y el saldo neto de movimientos con referencia. En muchos catálogos coincide con el saldo inicial de la "
            "cuenta madre de clientes (por ejemplo, 104-000-000-000) más cualquier movimiento sin referencia.\n"
            "- **Identidad esperada**: `Saldo final ≈ saldo inicial global + saldo neto de movimientos del periodo`.\n"
            "- **Saldo cuentas sin facturas**: saldos finales de cuentas que no tienen ninguna factura con referencia "
            "en el periodo (por ejemplo, cuentas con solo saldo inicial o ajustes sin referencia).\n"
            "- **Saldo final cartera**: saldo 'Total Clientes' reportado por el auxiliar.\n"
            "- **Diferencia residual vs auxiliar**: parte del saldo final que no se explica solo con el saldo neto de "
            "movimientos del periodo y el saldo de cuentas sin facturas; puede deberse a movimientos sin referencia, "
            "reclasificaciones o redondeos propios del archivo."
        )

        # ----------------------------------------------------------------
        # Pestañas (detalle)
        # ----------------------------------------------------------------
        tab_resumen, tab_pendientes, tab_favor = st.tabs(
            [
                "📂 Resumen por cuenta vs auxiliar",
                "📑 Facturas pendientes",
                "💳 Facturas con saldo a favor",
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
            st.markdown("### Facturas pendientes (detalle) y saldos")

            # Base: global o por cuenta, según el filtro
            if codigo_cuenta_seleccionada is not None:
                base_df = facturas_cuenta_f[
                    facturas_cuenta_f["account_code"] == codigo_cuenta_seleccionada
                ].copy()
            else:
                base_df = facturas_global_f.copy()

            # Facturas con saldo neto > 0 (detalle)
            df_pend = base_df[base_df["saldo_factura"] > 0].copy()

            # Saldo pendiente total (solo facturas con saldo > 0)
            saldo_pendiente_total = df_pend["saldo_factura"].sum() if not df_pend.empty else 0.0

            # Saldo neto total "real" (según auxiliar)
            if codigo_cuenta_seleccionada is not None:
                saldo_neto_total_real = totales_cuentas_aux.loc[
                    totales_cuentas_aux["account_code"] == codigo_cuenta_seleccionada,
                    "saldo_final_cuenta_aux",
                ].sum()
            else:
                saldo_neto_total_real = resumen_aux.get("saldo_final_aux")

            if df_pend.empty:
                st.info("No hay facturas pendientes (saldo neto > 0) con estos filtros.")
            else:
                total_facturas = df_pend["referencia"].nunique()

                c1, c2, c3 = st.columns(3)
                with c1:
                    st.metric(
                        "Número de facturas pendientes (detalle)",
                        value=int(total_facturas),
                    )
                with c2:
                    st.metric(
                        "Saldo pendiente total (solo facturas con saldo > 0)",
                        value=f"${saldo_pendiente_total:,.2f}",
                    )
                with c3:
                    if saldo_neto_total_real is not None:
                        st.metric(
                            "Saldo neto total (según auxiliar)",
                            value=f"${saldo_neto_total_real:,.2f}",
                        )
                    else:
                        st.metric("Saldo neto total (según auxiliar)", value="N/D")

                st.caption(
                    "- **Saldo pendiente total**: suma de los saldos de todas las facturas con saldo > 0 "
                    "en el rango de fechas y cuenta(s) seleccionados.\n"
                    "- **Saldo neto total (según auxiliar)**: saldo final contable real de la cuenta (o de toda la cartera), "
                    "incluyendo saldo inicial y todos los movimientos."
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
        # SECCIÓN: Referencias en varias cuentas (cruces)
        # ================================================================
        st.subheader("🔁 Referencias en varias cuentas (cargos en una cuenta, abonos en otra)")

        if detalle_cruces.empty:
            st.info(
                "No se encontraron referencias que tengan cargos en una cuenta y abonos en otra."
            )
        else:
            num_refs_cruce = resumen_cruces["referencia"].nunique()
            total_cuentas_afectadas = detalle_cruces["account_code"].nunique()

            c1, c2, c3 = st.columns(3)
            with c1:
                st.metric(
                    "Referencias con cruce entre cuentas",
                    value=int(num_refs_cruce),
                )
            with c2:
                st.metric(
                    "Cuentas contables involucradas",
                    value=int(total_cuentas_afectadas),
                )
            with c3:
                saldo_global_cruces = resumen_cruces["saldo_neto_ref"].sum()
                st.metric(
                    "Saldo neto global de estas referencias",
                    value=f"${saldo_global_cruces:,.2f}",
                )

            st.caption(
                "- Se listan referencias que aparecen en **más de una cuenta contable**, "
                "y que tienen **cargos en alguna cuenta y abonos en otra**.\n"
                "- Útil para revisar pagos aplicados en cuentas distintas a donde se originó la factura."
            )

            cols_det = [
                "referencia",
                "account_code",
                "account_name",
                "cargos_total",
                "abonos_total",
                "saldo_por_cuenta",
                "num_cuentas",
                "cargos_tot_ref",
                "abonos_tot_ref",
                "saldo_neto_ref",
            ]

            df_cruces_view = detalle_cruces[cols_det].sort_values(
                ["referencia", "account_code"]
            )

            st.dataframe(df_cruces_view, use_container_width=True)

            xls_cruces = to_excel(df_cruces_view)
            st.download_button(
                label="⬇️ Descargar referencias cruzadas en Excel",
                data=xls_cruces,
                file_name="referencias_cruzadas_por_cuenta.xlsx",
                mime=(
                    "application/vnd.openxmlformats-officedocument."
                    "spreadsheetml.sheet"
                ),
            )
