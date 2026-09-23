"""
Control de Empaque QR — María Almenara
========================================
"""

from datetime import datetime
import io
import math
import uuid

from google.oauth2.service_account import Credentials
import gspread
from PIL import Image, ImageDraw, ImageFont
import pandas as pd
import plotly.graph_objects as go
import streamlit as st
from streamlit_qrcode_scanner import qrcode_scanner

st.set_page_config(
    page_title="Control de Empaque QR — María Almenara",
    page_icon="📦",
    layout="wide",
    initial_sidebar_state="expanded",
)

C_PRIMARY = "#0F5C55"
C_PRIMARY_DARK = "#0A3F3A"
C_LIGHT = "#EAF5F3"
C_ACCENT = "#D98E73"
C_BG = "#FFFBF7"
C_OK = "#1F9D55"
C_WARN = "#F5A623"
C_BAD = "#E24C4C"

st.markdown(
    f"""
    <style>
    .stApp {{ background-color: {C_BG}; }}
    section[data-testid="stSidebar"] {{ background-color: {C_PRIMARY_DARK}; }}
    section[data-testid="stSidebar"] * {{ color: #F2F2F2 !important; }}
    div[data-testid="stMetric"] {{
        background-color: white; border: 1px solid #E2E2E2;
        border-left: 6px solid {C_PRIMARY}; border-radius: 10px; padding: 12px 16px;
    }}
    </style>
    """,
    unsafe_allow_html=True,
)


def fmt_num(x) -> str:
  try:
    x = float(x)
  except (TypeError, ValueError):
    return str(x)
  if abs(x - round(x)) < 0.01:
    return f"{x:,.0f}"
  return f"{x:,.2f}"


GOOGLE_SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]
PEDIDO_HEADERS = [
    "codigo",
    "descripcion",
    "umi",
    "factor",
    "requerimiento",
    "fecha_pedido",
]
CATALOGO_HEADERS = ["codigo", "producto", "factor", "umr"]
REGISTROS_HEADERS = [
    "id",
    "codigo",
    "producto",
    "cantidad_cajas",
    "factor",
    "unidades",
    "umi",
    "operario",
    "timestamp",
]


@st.cache_resource
def get_spreadsheet():
  creds = Credentials.from_service_account_info(
      dict(st.secrets["gcp_service_account"]), scopes=GOOGLE_SCOPES
  )
  client = gspread.authorize(creds)
  return client.open_by_key(st.secrets["SHEET_ID"])


@st.cache_resource
def get_ws(nombre: str, headers: tuple):
  sh = get_spreadsheet()
  try:
    ws = sh.worksheet(nombre)
  except gspread.WorksheetNotFound:
    ws = sh.add_worksheet(title=nombre, rows=2000, cols=max(len(headers) + 2, 8))
    ws.append_row(list(headers), value_input_option="RAW")
    return ws
  if not ws.row_values(1):
    ws.append_row(list(headers), value_input_option="RAW")
  return ws


def _rows_seguras(df: pd.DataFrame, columnas: list, numericas: set) -> list:
  filas = []
  for _, r in df.iterrows():
    fila = []
    for col in columnas:
      val = r.get(col)
      if col in numericas:
        fila.append(float(val) if pd.notna(val) else 0.0)
      else:
        fila.append("" if pd.isna(val) else str(val))
    filas.append(fila)
  return filas


def _sheet_a_df(ws, columnas_esperadas: list, numericas: set) -> pd.DataFrame:
  registros = ws.get_all_records(value_render_option="UNFORMATTED_VALUE")
  df = pd.DataFrame(registros)
  if df.empty:
    return pd.DataFrame(columns=columnas_esperadas)
  for col in numericas:
    if col in df.columns:
      df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)
  return df


def guardar_pedido(
    data_df: pd.DataFrame, receta_df: pd.DataFrame, factor_lima_df: pd.DataFrame
):
  d = data_df[
      ["CÓDIGO", "DESCRIPCIÓN", "UMI", "FACTOR", "REQUERIMIENTO", "FECHA PEDIDO"]
  ].copy()
  d.columns = PEDIDO_HEADERS
  d["codigo"] = d["codigo"].astype(str).str.strip()
  d["fecha_pedido"] = pd.to_datetime(
      d["fecha_pedido"], errors="coerce"
  ).dt.strftime("%Y-%m-%d")

  c = receta_df.drop_duplicates(subset="CÓDIGO")[
      ["CÓDIGO", "PRODUCTO", "FACTOR", "UMR"]
  ].copy()
  c.columns = CATALOGO_HEADERS
  c["codigo"] = c["codigo"].astype(str).str.strip()

  ws_pedido = get_ws("Pedido", tuple(PEDIDO_HEADERS))
  ws_pedido.clear()
  ws_pedido.append_row(PEDIDO_HEADERS, value_input_option="RAW")
  ws_pedido.append_rows(
      _rows_seguras(d, PEDIDO_HEADERS, {"factor", "requerimiento"}),
      value_input_option="RAW",
  )

  ws_catalogo = get_ws("Catalogo", tuple(CATALOGO_HEADERS))
  ws_catalogo.clear()
  ws_catalogo.append_row(CATALOGO_HEADERS, value_input_option="RAW")
  ws_catalogo.append_rows(
      _rows_seguras(c, CATALOGO_HEADERS, {"factor"}), value_input_option="RAW"
  )

  fl_headers = ["codigo", "descripcion", "factor"]
  ws_factor = get_ws("FactorLima", tuple(fl_headers))
  ws_factor.clear()
  ws_factor.append_row(fl_headers, value_input_option="RAW")
  fl_data = factor_lima_df[["CÓDIGO", "DESCRIPCIÓN", "FACTOR"]].copy()
  fl_data.columns = fl_headers
  fl_data["codigo"] = fl_data["codigo"].astype(str).str.strip()
  ws_factor.append_rows(
      _rows_seguras(fl_data, fl_headers, {"factor"}), value_input_option="RAW"
  )
  st.cache_data.clear()


@st.cache_data(ttl=5)
def cargar_pedido():
  try:
    ws_pedido = get_ws("Pedido", tuple(PEDIDO_HEADERS))
    ws_catalogo = get_ws("Catalogo", tuple(CATALOGO_HEADERS))
    ws_factor = get_ws("FactorLima", tuple(["codigo", "descripcion", "factor"]))
    pedido = _sheet_a_df(ws_pedido, PEDIDO_HEADERS, {"factor", "requerimiento"})
    catalogo = _sheet_a_df(ws_catalogo, CATALOGO_HEADERS, {"factor"})
    factor_lima = _sheet_a_df(
        ws_factor, ["codigo", "descripcion", "factor"], {"factor"}
    )
    return pedido, catalogo, factor_lima
  except Exception:
    return (
        pd.DataFrame(columns=PEDIDO_HEADERS),
        pd.DataFrame(columns=CATALOGO_HEADERS),
        pd.DataFrame(),
    )


@st.cache_data(ttl=3)
def cargar_registros():
  try:
    ws = get_ws("Registros", tuple(REGISTROS_HEADERS))
    df = _sheet_a_df(
        ws, REGISTROS_HEADERS, {"cantidad_cajas", "factor", "unidades"}
    )
    return df.iloc[::-1].reset_index(drop=True) if not df.empty else df
  except Exception:
    return pd.DataFrame(columns=REGISTROS_HEADERS)


def registrar_escaneo(codigo, producto, cajas, factor, unidades, umi, operario):
  ws = get_ws("Registros", tuple(REGISTROS_HEADERS))
  fila = [
      uuid.uuid4().hex[:8],
      str(codigo),
      str(producto),
      float(cajas),
      float(factor),
      float(unidades),
      str(umi),
      str(operario),
      datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
  ]
  ws.append_row(fila, value_input_option="RAW")
  st.cache_data.clear()


def eliminar_registro(reg_id: str):
  ws = get_ws("Registros", tuple(REGISTROS_HEADERS))
  celda = ws.find(str(reg_id), in_column=1)
  if celda:
    ws.delete_rows(celda.row)
  st.cache_data.clear()


def reiniciar_registros():
  ws = get_ws("Registros", tuple(REGISTROS_HEADERS))
  ws.clear()
  ws.append_row(REGISTROS_HEADERS, value_input_option="RAW")
  st.cache_data.clear()


def calcular_avance(
    pedido_df: pd.DataFrame, registros_df: pd.DataFrame
) -> pd.DataFrame:
  if pedido_df.empty:
    return pd.DataFrame()
  if registros_df.empty:
    acumulado = pd.DataFrame(
        columns=["codigo", "unidades_entregadas", "cajas_entregadas"]
    )
  else:
    acumulado = (
        registros_df.groupby("codigo")
        .agg(
            unidades_entregadas=("unidades", "sum"),
            cajas_entregadas=("cantidad_cajas", "sum"),
        )
        .reset_index()
    )
  df = pedido_df.merge(acumulado, on="codigo", how="left")
  df["unidades_entregadas"] = df["unidades_entregadas"].fillna(0)
  df["cajas_entregadas"] = df["cajas_entregadas"].fillna(0)
  req_seguro = df["requerimiento"].replace(0, pd.NA)
  df["pct_avance"] = (
      (df["unidades_entregadas"] / req_seguro * 100).fillna(0)
  ).astype(float)
  df["estado"] = df["pct_avance"].apply(
      lambda p: (
          "⛔ Sin iniciar"
          if p <= 0
          else ("🟡 Parcial" if p < 100 else "✅ Completo")
      )
  )
  return df.sort_values("pct_avance")


def calcular_componentes_factor_m(pedido_df: pd.DataFrame, archivo_subido=None):
  try:
    if archivo_subido is not None:
      xl = pd.ExcelFile(archivo_subido)
      df_data = xl.parse("data")
      df_receta = xl.parse("Receta")
      df_factor = xl.parse("factor-lima")
    else:
      return pd.DataFrame()
  except Exception:
    return pd.DataFrame()

  df_data.columns = [str(c).strip().upper() for c in df_data.columns]
  df_receta.columns = [str(c).strip().upper() for c in df_receta.columns]
  df_factor.columns = [str(c).strip().upper() for c in df_factor.columns]

  # Extraer el requerimiento real directamente de la hoja data
  col_req_data = next(
      (c for c in df_data.columns if "REQUERIMIENTO" in c), "REQUERIMIENTO"
  )
  col_cod_data = next(
      (c for c in df_data.columns if "CÓDIGO" in c or "CODIGO" in c), "CÓDIGO"
  )

  df_data_limpio = df_data[[col_cod_data, col_req_data]].copy()
  df_data_limpio.columns = ["CÓDIGO_PROD", "REQ_REAL"]
  df_data_limpio["REQ_REAL"] = (
      pd.to_numeric(df_data_limpio["REQ_REAL"], errors="coerce").fillna(0)
  )

  # Cruzar la receta con el requerimiento real de data
  col_cod_receta = next(
      (c for c in df_receta.columns if "CÓDIGO" in c or "CODIGO" in c), "CÓDIGO"
  )
  col_comp_receta = next(
      (c for c in df_receta.columns if "COMPONENTE" in c), "COD COMPONENTE"
  )

  df_merged = pd.merge(
      df_receta,
      df_data_limpio,
      left_on=col_cod_receta,
      right_on="CÓDIGO_PROD",
      how="inner",
  )

  # Asignar directamente el requerimiento de data sin multiplicar por cantidades unitarias de receta
  df_merged["REQ_COMPONENTE"] = df_merged["REQ_REAL"]

  # Filtrar componentes que empiezan con M
  df_m = df_merged[
      df_merged[col_comp_receta].astype(str).str.upper().str.startswith("M")
  ].copy()

  # Cruzar con factor-lima
  col_cod_f = next(
      (c for c in df_factor.columns if "COD" in c), df_factor.columns[0]
  )
  col_fac_f = next(
      (c for c in df_factor.columns if "FACTOR" in c), df_factor.columns[-1]
  )
  df_factor = df_factor.rename(
      columns={col_cod_f: "CÓDIGO_LIMA", col_fac_f: "FACTOR_VALOR"}
  )

  df_final = pd.merge(
      df_m,
      df_factor,
      left_on=col_comp_receta,
      right_on="CÓDIGO_LIMA",
      how="left",
  )
  df_final["FACTOR_LIMA"] = (
      pd.to_numeric(df_final["FACTOR_VALOR"], errors="coerce").fillna(1.0)
  )
  df_final["FACTOR_LIMA"] = df_final["FACTOR_LIMA"].apply(
      lambda x: x if x > 0 else 1.0
  )

  col_desc_comp = next(
      (
          c
          for c in df_final.columns
          if "DESCRIPCIÓN" in c or "PRODUCTO" in c or c == "COMPONENTE"
      ),
      col_comp_receta,
  )

  df_grouped = (
      df_final.groupby(
          [col_comp_receta, col_desc_comp, "FACTOR_LIMA"], as_index=False
      )["REQ_COMPONENTE"]
      .sum()
      .rename(columns={"REQ_COMPONENTE": "REQ_TOTAL"})
  )

  df_grouped["REQ_REDONDEADO"] = df_grouped.apply(
      lambda row: math.ceil(row["REQ_TOTAL"] / row["FACTOR_LIMA"])
      * row["FACTOR_LIMA"],
      axis=1,
  )

  return (
      df_grouped[[
          col_comp_receta,
          col_desc_comp,
          "REQ_TOTAL",
          "FACTOR_LIMA",
          "REQ_REDONDEADO",
      ]]
      .rename(
          columns={
              col_comp_receta: "CÓDIGO",
              col_desc_comp: "DESCRIPCIÓN",
              "REQ_TOTAL": "REQUERIMIENTO NETO",
              "FACTOR_LIMA": "FACTOR LIMA",
              "REQ_REDONDEADO": "REQUERIMIENTO REDONDEADO",
          }
      )
      .sort_values("CÓDIGO")
      .reset_index(drop=True)
  )


# --- INTERFAZ ---
try:
  with st.sidebar:
    st.markdown("## 📦 Control de Empaque QR")
    operario = st.text_input(
        "👤 Operario", value=st.session_state.get("operario", "")
    )
    st.session_state["operario"] = operario
    st.markdown("---")

    archivo = st.file_uploader(
        "Excel con hojas 'data', 'Receta' y 'factor-lima'", type=["xlsx"]
    )
    if archivo is not None:
      firma = f"{archivo.name}-{archivo.size}"
      if firma != st.session_state.get("ultimo_archivo_cargado"):
        xl_temp = pd.ExcelFile(archivo)
        if {"data", "Receta", "factor-lima"}.issubset(
            set(xl_temp.sheet_names)
        ):
          guardar_pedido(
              xl_temp.parse("data"),
              xl_temp.parse("Receta"),
              xl_temp.parse("factor-lima"),
          )
          st.session_state["ultimo_archivo_cargado"] = firma
          st.session_state["archivo_bytes_actual"] = archivo.getvalue()
          st.success("¡Pedido cargado con éxito!")
          st.rerun()
        else:
          st.error("El Excel debe tener las hojas: data, Receta, factor-lima")

    pedido_df, catalogo_df, _ = cargar_pedido()
    if not pedido_df.empty:
      st.caption(f"✅ {len(pedido_df)} productos cargados en sistema.")

    if st.button("🔄 Limpiar Caché / Forzar Recarga"):
      st.cache_data.clear()
      st.cache_resource.clear()
      st.success("Caché limpia.")
      st.rerun()

  if pedido_df.empty or catalogo_df.empty:
    st.title("📦 Control de Empaque QR")
    st.info("👈 Sube el archivo Excel en la barra lateral para iniciar.")
    st.stop()

  registros_df = cargar_registros()
  avance_df = calcular_avance(pedido_df, registros_df)

  st.title("📦 Control de Empaque QR")
  tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs([
      "📷 Escanear",
      "📊 Avance",
      "🕒 Historial",
      "🏷️ Etiquetas",
      "⚙️ Factor M",
      "📤 Exportar",
  ])

  with tab1:
    codigo_leido = qrcode_scanner(key="scanner")
    c_man = st.text_input("O ingresa código manual", key="manual_in")
    if c_man:
      codigo_leido = c_man

    if codigo_leido:
      codigo_actual = codigo_leido.strip().upper()
      cat_row = catalogo_df[catalogo_df["codigo"] == codigo_actual]
      if cat_row.empty:
        st.error(f"El código {codigo_actual} no está en el catálogo.")
      else:
        prod = cat_row.iloc[0]["producto"]
        ped_row = pedido_df[pedido_df["codigo"] == codigo_actual]
        factor = (
            ped_row.iloc[0]["factor"]
            if not ped_row.empty
            else cat_row.iloc[0]["factor"]
        )
        umi = (
            ped_row.iloc[0]["umi"]
            if not ped_row.empty
            else cat_row.iloc[0]["umr"]
        )

        st.markdown(f"**Producto:** {prod} (`{codigo_actual}`)")
        cajas = st.number_input("Cajas", min_value=0.0, value=1.0, step=1.0)
        unidades = cajas * float(factor)
        st.caption(f"= {fmt_num(unidades)} {umi}")

        if st.button("✅ Registrar entrega", type="primary"):
          registrar_escaneo(
              codigo_actual,
              prod,
              cajas,
              float(factor),
              unidades,
              umi,
              operario or "Operario",
          )
          st.success("¡Registrado!")
          st.rerun()

  with tab2:
    st.dataframe(avance_df, use_container_width=True, hide_index=True)

  with tab3:
    if registros_df.empty:
      st.info("No hay registros aún.")
    else:
      st.dataframe(registros_df, use_container_width=True, hide_index=True)

  with tab4:
    st.info("Generación de etiquetas disponible.")

  with tab5:
    st.markdown("### ⚙️ Requerimiento de Componentes (Factor M)")
    if "archivo_bytes_actual" in st.session_state:
      df_fm = calcular_componentes_factor_m(
          pedido_df,
          archivo_subido=io.BytesIO(st.session_state["archivo_bytes_actual"]),
      )
      if not df_fm.empty:
        st.dataframe(df_fm, use_container_width=True, hide_index=True)
      else:
        st.info("No hay componentes M o faltan columnas.")
    else:
      st.info("Carga el Excel del pedido en la barra lateral.")

  with tab6:
    st.dataframe(avance_df, use_container_width=True, hide_index=True)

except Exception as e:
  st.error(f"Se ha producido un error al ejecutar la aplicación: {e}")
  if st.button("🗑️ Limpiar Todo y Reiniciar"):
    st.cache_data.clear()
    st.cache_resource.clear()
    st.rerun()
