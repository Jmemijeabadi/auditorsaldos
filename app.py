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
# CONFIGURACIÓN E INTERFAZ
# ==============================================================================
st.set_page_config(page_title="Auditoría de Saldos CONTPAQ", layout="wide", page_icon="🛡️")

UMBRAL_TOLERANCIA = 1.0  # Tolerancia para diferencias por redondeo

st.title("🛡️ Auditoría Blindada de Saldos (CONTPAQ i)")
st.markdown("""
Esta herramienta está calibrada para leer reportes de **Movimientos Auxiliares del Catálogo**.
Soporta archivos originales **.xlsx** o conversiones a **.csv**.

**Proceso:**
1. Lee el **Saldo Inicial** real de cada cuenta.
2. Suma los movimientos (facturas, pagos).
3. Compara contra el **Saldo Final** del reporte.
4. Cruza la información contra el detalle de facturas vivas.
""")

# ==============================================================================
# LÓGICA DE NEGOCIO Y LIMPIEZA
# ==============================================================================

@st.cache_data
def parse_spanish_date(s: str):
    """Convierte fechas '02/Ene/2025' a datetime de forma robusta."""
    if pd.isna(s):
        return pd.NaT
    s = str(s).strip()
    # Regex para dd/Mmm/aaaa
    m = re.match(r"^(\d{1,2})[/\-]([A-Za-z]{3})[/\-](\d{4})$", s)
    if not m:
        return pd.NaT
    day, mon_abbr, year = m.groups()
    month_map = {
        "ene": "01", "feb": "02", "mar": "03", "abr": "04",
        "may": "05", "jun": "06", "jul": "07", "ago": "08",
        "sep": "09", "oct": "10", "nov": "11", "dic": "12"
    }
    mon_key = mon_abbr[:3].lower()
    if mon_key not in month_map:
        return pd.NaT
    
    # Relleno de ceros para el día (ej: 2 -> 02)
    day = day.zfill(2)
    date_str = f"{day}/{month_map[mon_key]}/{year}"
    return pd.to_datetime(date_str, format="%d/%m/%Y", errors="coerce")

def normalizar_referencia(ref):
    """Limpia referencias (quita 'FACTURA', espacios extra, ceros decimales)."""
    if pd.isna(ref):
        return None
    # Si viene como flotante (ej: 123.0)
    if isinstance(ref, float):
        if ref.is_integer():
            s = str(int(ref))
        else:
            s = str(ref)
    else:
        s = str(ref).strip()

    # Normalización Unicode
    s = "".join(c for c in unicodedata.normalize("NFKD", s) if not unicodedata.combining(c)).upper()
    
    # Quitar .0 al final
    s = re.sub(r"\.0+$", "", s)
    
    # Quitar prefijos comunes ruidosos
    patron_ruido = r"^(FACTURA|FAC|FOLIO|REF|REFERENCIA|RECIBO|DEPOSITO|DEP|PAGO|ABONO|NC|NOTA)\s*[:.\-]?\s*"
    s = re.sub(patron_ruido, "", s)
    
    # Quitar prefijo "F" suelto (ej: F-123 -> 123)
    s = re.sub(r"^F\s*[-:]?\s*", "", s)
    
    # Quitar espacios internos y caracteres especiales
    s = re.sub(r"[ \-_/]", "", s)
    
    return s if s else None

def cargar_archivo_robusto(uploaded_file):
    """Intenta cargar Excel, si falla intenta CSV (codificación CONTPAQ)."""
    try:
        # Intento 1: Excel estándar
        df = pd.read_excel(uploaded_file, header=None)
        return df
    except Exception:
        # Intento 2: CSV (puede venir del usuario)
        uploaded_file.seek(0)
        try:
            # CONTPAQ suele usar CP1252 o Latin-1
            df = pd.read_csv(uploaded_file, header=None, encoding='latin-1')
            return df
        except Exception:
            try:
                uploaded_file.seek(0)
                df = pd.read_csv(uploaded_file, header=None, encoding='utf-8')
                return df
            except Exception as e:
                st.error(f"No se pudo leer el archivo. Asegúrate de subir el Excel (.xlsx) o CSV original. Error: {e}")
                return pd.DataFrame()

@st.cache_data
def procesar_datos_contpaq(file):
    raw = cargar_archivo_robusto(file)
    if raw.empty:
        return None, None, None

    # 1. Identificar filas de Encabezado de Cuenta
    # Formato: Col 0 tiene código 000-000... Y Col 6 contiene "Saldo inicial"
    # Ajuste: A veces "Saldo inicial" está en col 5 o 6 dependiendo del reporte. Buscamos en fila.
    
    # Regex código cuenta: 3 dígitos - 3 dígitos...
    patron_cuenta = r"^\d{3}-\d{3}-\d{3}-\d{3}"
    
    raw_str = raw.astype(str)
    
    # Máscara: Columna 0 parece cuenta
    mask_cuenta_code = raw_str[0].str.match(patron_cuenta, na=False)
    
    # Máscara: Alguna columna tiene "Saldo inicial"
    mask_saldo_ini_txt = raw_str.apply(lambda row: row.str.contains("Saldo inicial", case=False, na=False).any(), axis=1)
    
    is_account_row = mask_cuenta_code & mask_saldo_ini_txt
    
    # Propagar datos de cuenta hacia abajo
    df = raw.copy()
    df["meta_codigo"] = np.where(is_account_row, df[0], np.nan)
    df["meta_nombre"] = np.where(is_account_row, df[1], np.nan)
    
    # EXTRAER SALDO INICIAL REAL (Columna 7 usualmente)
    # Buscamos dónde está el valor numérico del saldo inicial. Normalmente col 7.
    df["meta_saldo_inicial"] = np.where(is_account_row, pd.to_numeric(df[7], errors='coerce'), np.nan)
    
    df["meta_codigo"] = df["meta_codigo"].ffill()
    df["meta_nombre"] = df["meta_nombre"].ffill()
    df["meta_saldo_inicial"] = df["meta_saldo_inicial"].ffill()

    # 2. Identificar filas de Movimientos (fechas)
    patron_fecha = r"^\d{1,2}/[A-Za-z]{3}/\d{4}"
    is_mov_row = raw_str[0].str.match(patron_fecha, na=False)
    
    movs = df[is_mov_row].copy()
    
    # Mapeo estricto de columnas según estructura CONTPAQ
    # 0:Fecha, 1:Tipo, 2:Num, 3:Concepto, 4:Ref, 5:Cargo, 6:Abono, 7:Saldo
    cols_map = {
        0: "fecha_raw", 1: "tipo", 2: "poliza", 3: "concepto", 
        4: "referencia", 5: "cargos", 6: "abonos", 7: "saldo_acumulado"
    }
    movs = movs.rename(columns=cols_map)
    
    # Limpieza de numéricos
    for c in ["cargos", "abonos", "saldo_acumulado"]:
        if c in movs.columns:
            movs[c] = pd.to_numeric(movs[c], errors='coerce').fillna(0)
            
    # Parsear fechas y referencias
    movs["fecha"] = movs["fecha_raw"].apply(parse_spanish_date)
    movs["referencia_norm"] = movs["referencia"].apply(normalizar_referencia)
    
    # Filtrar movimientos válidos
    movs_validos = movs.dropna(subset=["fecha"]).copy()
    
    # 3. Procesar Totales por Cuenta (Fila "Total:")
    # Buscamos filas donde la columna 4 (usualmente) dice "Total:"
    mask_total = raw_str[4].str.strip() == "Total:"
    totales_raw = df[mask_total].copy()
    
    # Extraer totales del auxiliar
    totales_raw["cargos_aux"] = pd.to_numeric(totales_raw[5], errors='coerce').fillna(0)
    totales_raw["abonos_aux"] = pd.to_numeric(totales_raw[6], errors='coerce').fillna(0)
    totales_raw["saldo_final_aux"] = pd.to_numeric(totales_raw[7], errors='coerce').fillna(0)
    
    resumen_cuentas = totales_raw[[
        "meta_codigo", "meta_nombre", "meta_saldo_inicial", 
        "cargos_aux", "abonos_aux", "saldo_final_aux"
    ]].copy()
    
    # Agrupar por si acaso hay duplicados (raro en este reporte)
    resumen_cuentas = resumen_cuentas.groupby(["meta_codigo", "meta_nombre"]).sum().reset_index()

    # VALIDACIÓN MATEMÁTICA INTERNA DEL REPORTE
    # Saldo Inicial + Cargos - Abonos DEBE SER IGUAL a Saldo Final
    resumen_cuentas["saldo_calculado"] = resumen_cuentas["meta_saldo_inicial"] + resumen_cuentas["cargos_aux"] - resumen_cuentas["abonos_aux"]
    resumen_cuentas["discrepancia_reporte"] = resumen_cuentas["saldo_final_aux"] - resumen_cuentas["saldo_calculado"]
    resumen_cuentas["error_integridad"] = resumen_cuentas["discrepancia_reporte"].abs() > UMBRAL_TOLERANCIA

    return movs_validos, resumen_cuentas

def cruzar_informacion(movs, resumen_cuentas):
    # 1. Agrupar movimientos por Cuenta para ver la suma de lo que leímos
    movs_agrupados = movs.groupby(["meta_codigo"]).agg(
        cargos_leidos=("cargos", "sum"),
        abonos_leidos=("abonos", "sum")
    ).reset_index()
    
    # 2. Unir con el resumen del auxiliar
    final = resumen_cuentas.merge(movs_agrupados, on="meta_codigo", how="left").fillna(0)
    
    # 3. Calcular saldo de facturas vivas (solo lo que tiene referencia válida)
    movs_con_ref = movs[movs["referencia_norm"].notna()]
    saldo_vivas = movs_con_ref.groupby("meta_codigo").apply(
        lambda x: (x["cargos"] - x["abonos"]).sum()
    ).reset_index(name="saldo_facturas_vivas")
    
    final = final.merge(saldo_vivas, on="meta_codigo", how="left").fillna(0)
    
    # 4. Determinación de Diferencias
    # Diferencia Real = Saldo Final Auxiliar - Saldo Facturas Vivas
    final["diferencia_conciliacion"] = final["saldo_final_aux"] - final["saldo_facturas_vivas"]
    
    # Lógica de Semáforo
    def diagnostico(row):
        if row["error_integridad"]:
            return "❌ Error en Archivo (Integridad)"
        
        dif = row["diferencia_conciliacion"]
        saldo_fin = row["saldo_final_aux"]
        
        if abs(dif) <= UMBRAL_TOLERANCIA:
            return "✅ Conciliado"
        
        # Caso: Solo saldo inicial (diferencia es casi igual al saldo final)
        if abs(dif - saldo_fin) <= UMBRAL_TOLERANCIA and abs(saldo_fin) > UMBRAL_TOLERANCIA:
            return "⚠️ Posible Saldo Inicial Arrastrado"
            
        return "🔴 Diferencia a Investigar"

    final["estado"] = final.apply(diagnostico, axis=1)
    
    return final

# ==============================================================================
# APP PRINCIPAL
# ==============================================================================

uploaded_file = st.file_uploader("📂 Sube tu reporte (Excel o CSV)", type=["xlsx", "csv"])

if uploaded_file:
    with st.spinner("🔍 Analizando estructura del archivo..."):
        movs, resumen, error = None, None, None
        try:
            movs, resumen = procesar_datos_contpaq(uploaded_file)
        except Exception as e:
            st.error(f"Error crítico procesando el archivo: {e}")
            st.stop()
            
        if movs is None or resumen is None or resumen.empty:
            st.warning("No se encontraron datos válidos. Verifica que sea el reporte 'Movimientos Auxiliares del Catálogo'.")
            st.stop()
            
        # Cruzar datos
        df_audit = cruzar_informacion(movs, resumen)
        
        # Métricas Globales
        saldo_contable_total = df_audit["saldo_final_aux"].sum()
        saldo_facturas_total = df_audit["saldo_facturas_vivas"].sum()
        diferencia_global = saldo_contable_total - saldo_facturas_total
        
        # Alerta de Integridad
        cuentas_error_integridad = df_audit[df_audit["error_integridad"]]
        if not cuentas_error_integridad.empty:
            st.error(f"🚨 ALERTA CRÍTICA: Hay {len(cuentas_error_integridad)} cuentas donde el reporte no cuadra matemáticamente (Saldo Inicial + Movs != Final). Revisa si hay renglones ocultos o errores en CONTPAQ.")
            st.dataframe(cuentas_error_integridad)

        # Dashboard
        c1, c2, c3 = st.columns(3)
        c1.metric("Saldo Contable Total", f"${saldo_contable_total:,.2f}")
        c2.metric("Saldo Soportado por Facturas", f"${saldo_facturas_total:,.2f}")
        c3.metric("Diferencia Global", f"${diferencia_global:,.2f}", delta_color="inverse")
        
        st.divider()
        
        # Pestañas
        tab1, tab2, tab3 = st.tabs(["🚦 Semáforo de Cuentas", "📄 Detalle de Movimientos", "📉 Gráfico de Composición"])
        
        with tab1:
            st.subheader("Estado de Conciliación por Cuenta")
            
            # Filtro rápido
            modo_ver = st.radio("Mostrar:", ["Todo", "Solo Problemas"], horizontal=True)
            
            df_view = df_audit.copy()
            if modo_ver == "Solo Problemas":
                df_view = df_view[df_view["estado"].str.contains("🔴|⚠️|❌")]
            
            st.dataframe(
                df_view[[
                    "meta_codigo", "meta_nombre", "estado",
                    "meta_saldo_inicial", "saldo_final_aux", 
                    "saldo_facturas_vivas", "diferencia_conciliacion"
                ]].sort_values("diferencia_conciliacion", ascending=False),
                column_config={
                    "meta_saldo_inicial": st.column_config.NumberColumn("Saldo Inicial (Real)", format="$%.2f"),
                    "saldo_final_aux": st.column_config.NumberColumn("Saldo Final (Aux)", format="$%.2f"),
                    "saldo_facturas_vivas": st.column_config.NumberColumn("Facturas Vivas", format="$%.2f"),
                    "diferencia_conciliacion": st.column_config.NumberColumn("Diferencia", format="$%.2f"),
                    "estado": "Diagnóstico"
                },
                use_container_width=True,
                hide_index=True
            )

        with tab2:
            st.subheader("Explorador de Movimientos")
            cuenta_sel = st.selectbox("Selecciona Cuenta", df_audit["meta_nombre"].unique())
            codigo_sel = df_audit[df_audit["meta_nombre"] == cuenta_sel]["meta_codigo"].iloc[0]
            
            movs_cuenta = movs[movs["meta_codigo"] == codigo_sel].copy()
            
            st.dataframe(
                movs_cuenta[["fecha", "tipo", "poliza", "concepto", "referencia", "cargos", "abonos", "saldo_acumulado"]],
                use_container_width=True
            )
        
        with tab3:
            # Gráfico Waterfall
            fig = go.Figure(go.Waterfall(
                name = "Conciliación", orientation = "v",
                measure = ["relative", "relative", "total"],
                x = ["Facturas Vivas", "Diferencia/Ajustes", "Saldo Contable"],
                textposition = "outside",
                text = [f"${saldo_facturas_total:,.0f}", f"${diferencia_global:,.0f}", f"${saldo_contable_total:,.0f}"],
                y = [saldo_facturas_total, diferencia_global, 0], # El total se calcula solo, pero el conector necesita ajuste
                connector = {"line":{"color":"rgb(63, 63, 63)"}},
            ))
            
            # Ajuste simple de barras para Plotly
            fig = go.Figure(data=[
                go.Bar(name='Facturas Vivas', x=['Totales'], y=[saldo_facturas_total], marker_color='green'),
                go.Bar(name='Diferencia', x=['Totales'], y=[diferencia_global], marker_color='red'),
                go.Bar(name='Saldo Contable', x=['Totales'], y=[saldo_contable_total], marker_color='blue')
            ])
            fig.update_layout(title="Comparativa Global", barmode='group')
            
            st.plotly_chart(fig, use_container_width=True)

else:
    st.info("👆 Sube tu archivo 'Movimientos Auxiliares' para comenzar.")
