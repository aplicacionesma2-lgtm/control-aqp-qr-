"""
Control de Empaque QR — María Almenara
========================================
App Streamlit para escanear (QR/código) las cajas de producto terminado en
línea de empaque, hacer match contra el pedido del día (hoja "data") usando
el catálogo de productos (hoja "Receta") y calcular el % de avance de
entrega por producto en tiempo real.
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

# ======================================================================
# CONFIG Y ESTILOS
# ======================================================================
st.set_page_config(
    page_title="Control de Empaque QR — María Almenara",
    page_icon="📦",
    layout="wide",
    initial_sidebar_state="expanded",
)

C_PRIMARY = "#0F5C55"  # verde azulado corporativo
C_PRIMARY_DARK = "#0A3F3A"
C_LIGHT = "#EAF5F3"
C_ACCENT = "#D98E73"  # terracota cálido
C_BG = "#FFFBF7"
C_OK = "#1F9D55"
C_WARN = "#F5A623"
C_BAD = "#E24C4C"

st.markdown(
    f"""
    <style>
    .stApp {{ background-color: {C_BG}; }}
    section[data-testid="stSidebar"] {{
        background-color: {C_PRIMARY_DARK};
    }}
    section[data-testid="stSidebar"] * {{ color: #F2F2F2 !important; }}
    section[data-testid="stSidebar"] input {{ color: #111 !important; }}

    div[data-testid="stMetric"] {{
        background-color: white;
        border: 1px solid #E2E2E2;
        border-left: 6px solid {C_PRIMARY};
        border-radius: 10px;
        padding: 12px 16px;
        box-shadow: 0 1px 3px rgba(0,0,0,0.06);
    }}
    .producto-card {{
        background-color: white;
        border: 1px solid #E2E2E2;
        border-left: 8px solid {C_PRIMARY};
        border-radius: 12px;
        padding: 18px 22px;
        margin-bottom: 14px;
        box-shadow: 0 1px 4px rgba(0,0,0,0.08);
    }}
    .producto-card h3 {{ margin: 0 0 6px 0; color: {C_PRIMARY_DARK}; }}
    .badge {{
        display: inline-block; padding: 2px 10px; border-radius: 999px;
        font-size: 0.8rem; font-weight: 600; color: white;
    }}
    .stTabs [data-baseweb="tab"] {{ font-weight: 600; }}
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


# ======================================================================
# CAPA DE DATOS (Google Sheets)
# ======================================================================
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


def config_faltante() -> list:
  faltan = []
  if "gcp_service_account" not in st.secrets:
    faltan.append(
        "`[gcp_service_account]` (credenciales de la cuenta de servicio)"
    )
  if "SHEET_ID" not in st.secrets:
    faltan.append("`SHEET_ID`")
  return faltan


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


def procesar_excel(file):
  """Lee y valida el excel subido. Devuelve (data_df, receta_df, factor_lima_df, error)."""
  try:
    xl = pd.ExcelFile(file)
  except Exception as e:
    return None, None, None, f"No se pudo leer el archivo: {e}"

  if (
      "data" not in xl.sheet_names
      or "Receta" not in xl.sheet_names
      or "factor-lima" not in xl.sheet_names
  ):
    return (
        None,
        None,
        None,
        "El Excel debe tener las hojas 'data', 'Receta' y 'factor-lima'.",
    )

  data_df = xl.parse("data")
  receta_df = xl.parse("Receta")
  factor_lima_df = xl.parse("factor-lima")

  req_data = {
      "CÓDIGO",
      "DESCRIPCIÓN",
      "UMI",
      "FACTOR",
      "REQUERIMIENTO",
      "FECHA PEDIDO",
  }
  req_receta = {"CÓDIGO", "PRODUCTO", "FACTOR", "UMR"}
  req_factor = {"CÓDIGO", "DESCRIPCIÓN", "FACTOR"}

  if not req_data.issubset(set(data_df.columns)):
    faltan = req_data - set(data_df.columns)
    return None, None, None, f"A la hoja 'data' le faltan columnas: {faltan}"
  if not req_receta.issubset(set(receta_df.columns)):
    faltan = req_receta - set(receta_df.columns)
    return (
        None,
        None,
        None,
        f"A la hoja 'Receta' le faltan columnas: {faltan}",
    )
  if not req_factor.issubset(set(factor_lima_df.columns)):
    faltan = req_factor - set(factor_lima_df.columns)
    return (
        None,
        None,
        None,
        f"A la hoja 'factor-lima' le faltan columnas: {faltan}",
    )

  return data_df, receta_df, factor_lima_df, None


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
  ws_pedido = get_ws("Pedido", tuple(PEDIDO_HEADERS))
  ws_catalogo = get_ws("Catalogo", tuple(CATALOGO_HEADERS))
  ws_factor = get_ws("FactorLima", tuple(["codigo", "descripcion", "factor"]))
  pedido = _sheet_a_df(ws_pedido, PEDIDO_HEADERS, {"factor", "requerimiento"})
  catalogo = _sheet_a_df(ws_catalogo, CATALOGO_HEADERS, {"factor"})
  factor_lima = _sheet_a_df(
      ws_factor, ["codigo", "descripcion", "factor"], {"factor"}
  )
  return pedido, catalogo, factor_lima


@st.cache_data(ttl=3)
def cargar_registros():
  ws = get_ws("Registros", tuple(REGISTROS_HEADERS))
  df = _sheet_a_df(ws, REGISTROS_HEADERS, {"cantidad_cajas", "factor", "unidades"})
  return df.iloc[::-1].reset_index(drop=True) if not df.empty else df


def registrar_escaneo(
    codigo, producto, cajas, factor, unidades, umi, operario
):
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


# ======================================================================
# LÓGICA DE NEGOCIO
# ======================================================================
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
      (df["unidades_entregadas"] / req_seguro * 100).fillna(0).astype(float)
  )

  def estado(pct):
    if pct <= 0:
      return "⛔ Sin iniciar"
    elif pct < 100:
      return "🟡 Parcial"
    return "✅ Completo"

  df["estado"] = df["pct_avance"].apply(estado)
  return df.sort_values("pct_avance")


def calcular_componentes_factor_m(
    pedido_df: pd.DataFrame,
    archivo_subido=None,
    factor_lima_cargado=None,
):
  """Calcula los componentes que empiezan con M cruzando data -> Receta -> factor-lima con redondeo a múltiplo superior."""
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

  col_cod_f = next(
      (
          c
          for c in df_factor.columns
          if "COD" in c or c == "CÓDIGO" or c == "CODIGO"
      ),
      df_factor.columns[0],
  )
  col_fac_f = next(
      (c for c in df_factor.columns if "FACTOR" in c), df_factor.columns[-1]
  )

  df_factor = df_factor.rename(
      columns={col_cod_f: "CÓDIGO_LIMA", col_fac_f: "FACTOR_VALOR"}
  )

  df_merged = pd.merge(
      df_data,
      df_receta,
      on="CÓDIGO",
      how="inner",
      suffixes=("_PEDIDO", "_RECETA"),
  )

  df_merged["REQ_COMPONENTE"] = (
      df_merged["REQUERIMIENTO"] * df_merged["CANTIDAD"]
  )

  col_comp_receta = next(
      (c for c in df_merged.columns if "COMPONENTE" in c), "COD COMPONENTE"
  )
  df_m = df_merged[
      df_merged[col_comp_receta].astype(str).str.upper().str.startswith("M")
  ].copy()

  df_final = pd.merge(
      df_m,
      df_factor,
      left_on=col_comp_receta,
      right_on="CÓDIGO_LIMA",
      how="left",
  )

  if "FACTOR_VALOR" in df_final.columns:
    df_final["FACTOR_LIMA"] = (
        pd.to_numeric(df_final["FACTOR_VALOR"], errors="coerce").fillna(1)
    )
  else:
    df_final["FACTOR_LIMA"] = 1.0

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

  def redondear_factor(row):
    req = row["REQ_TOTAL"]
    factor = row["FACTOR_LIMA"]
    if factor <= 0:
      return req
    return math.ceil(req / factor) * factor

  df_grouped["REQ_REDONDEADO"] = df_grouped.apply(redondear_factor, axis=1)

  resultado = df_grouped[[
      col_comp_receta,
      col_desc_comp,
      "REQ_TOTAL",
      "FACTOR_LIMA",
      "REQ_REDONDEADO",
  ]].rename(
      columns={
          col_comp_receta: "CÓDIGO",
          col_desc_comp: "DESCRIPCIÓN",
          "REQ_TOTAL": "REQUERIMIENTO NETO",
          "FACTOR_LIMA": "FACTOR LIMA",
          "REQ_REDONDEADO": "REQUERIMIENTO REDONDEADO",
      }
  )
  return resultado.sort_values("CÓDIGO").reset_index(drop=True)


def generar_qr(codigo: str, size: int = 300) -> Image.Image:
  qr = qrcode.QRCode(version=1, box_size=10, border=2)
  qr.add_data(codigo)
  qr.make(fit=True)
  img = qr.make_image(fill_color="black", back_color="white").convert("RGB")
  return img.resize((size, size))


def generar_hoja_etiquetas(productos, columnas: int = 3) -> Image.Image:
  label_w, label_h = 380, 460
  filas = -(-len(productos) // columnas)
  sheet = Image.new("RGB", (label_w * columnas, label_h * filas), "white")
  draw = ImageDraw.Draw(sheet)
  try:
    font = ImageFont.truetype(
        "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf", 16
    )
    font_small = ImageFont.truetype(
        "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf", 13
    )
  except Exception:
    font = ImageFont.load_default()
    font_small = ImageFont.load_default()

  for idx, (codigo, nombre) in enumerate(productos):
    col, fila = idx % columnas, idx // columnas
    x0, y0 = col * label_w, fila * label_h
    draw.rectangle(
        [x0 + 5, y0 + 5, x0 + label_w - 5, y0 + label_h - 5],
        outline=C_PRIMARY,
        width=2,
    )
    qr_img = generar_qr(codigo, size=280)
    sheet.paste(qr_img, (x0 + (label_w - 280) // 2, y0 + 15))
    draw.text(
        (x0 + label_w // 2, y0 + 300),
        codigo,
        font=font,
        fill="black",
        anchor="mm",
    )
    nombre_corto = nombre if len(nombre) <= 40 else nombre[:37] + "..."
    draw.text(
        (x0 + label_w // 2, y0 + 330),
        nombre_corto,
        font=font_small,
        fill="black",
        anchor="mm",
    )

  return sheet


# ======================================================================
# APP
# ======================================================================
faltantes = config_faltante()
if faltantes:
  st.title("📦 Control de Empaque QR")
  st.error(
      "Faltan credenciales de Google Sheets en `secrets`: "
      + ", ".join(faltantes)
      + ". Revisa el README para configurarlas."
  )
  st.stop()

with st.sidebar:
  st.markdown("## 📦 Control de Empaque QR")
  st.caption("María Almenara")
  st.markdown("---")

  operario = st.text_input(
      "👤 Nombre del operario", value=st.session_state.get("operario", "")
  )
  st.session_state["operario"] = operario

  st.markdown("---")
  st.markdown("### 📁 Pedido del día")
  archivo = st.file_uploader(
      "Excel con hojas 'data', 'Receta' y 'factor-lima'", type=["xlsx"]
  )
  if archivo is not None:
    firma_archivo = f"{archivo.name}-{archivo.size}"
    if firma_archivo != st.session_state.get("ultimo_archivo_cargado"):
      data_df, receta_df, factor_lima_df, error = procesar_excel(archivo)
      if error:
        st.error(error)
      else:
        guardar_pedido(data_df, receta_df, factor_lima_df)
        st.cache_data.clear()
        st.session_state["ultimo_archivo_cargado"] = firma_archivo
        st.session_state["archivo_bytes_actual"] = archivo.getvalue()
        st.success(f"Pedido cargado: {len(data_df)} productos.")
        st.rerun()

  pedido_df, catalogo_df, factor_lima_df = cargar_pedido()
  if not pedido_df.empty:
    fechas = ", ".join(sorted(pedido_df["fecha_pedido"].dropna().unique()))
    st.caption(
        f"📅 {fechas}  \n{len(pedido_df)} productos en pedido ·"
        f" {len(catalogo_df)} en catálogo"
    )

  st.markdown("---")
  with st.expander("⚠️ Reiniciar historial del día"):
    st.caption(
        "Borra todos los escaneos registrados. No afecta el pedido ni el"
        " catálogo."
    )
    if st.button("🗑️ Borrar historial de escaneos"):
      reiniciar_registros()
      st.success("Historial reiniciado.")
      st.rerun()

if pedido_df.empty or catalogo_df.empty:
  st.title("📦 Control de Empaque QR")
  st.info(
      "👈 Sube el Excel del pedido del día (hojas **data**, **Receta** y"
      " **factor-lima**) en la barra lateral para comenzar."
  )
  st.stop()

registros_df = cargar_registros()
avance_df = calcular_avance(pedido_df, registros_df)

st.title("📦 Control de Empaque QR")

total_productos = len(avance_df)
completos = int((avance_df["pct_avance"] >= 100).sum())
total_req = avance_df["requerimiento"].sum()
total_entregado = avance_df["unidades_entregadas"].sum()
pct_general = (total_entregado / total_req * 100) if total_req else 0

k1, k2, k3, k4 = st.columns(4)
k1.metric("Productos en pedido", total_productos)
k2.metric("Completados", f"{completos} / {total_productos}")
k3.metric("Avance general", f"{pct_general:.1f}%")
k4.metric("Escaneos registrados", len(registros_df))

tab_scan, tab_avance, tab_hist, tab_labels, tab_factor_m, tab_export = (
    st.tabs([
        "📷 Escanear",
        "📊 Avance del pedido",
        "🕒 Historial",
        "🏷️ Etiquetas",
        "⚙️ Factor M",
        "📤 Exportar",
    ])
)

# ----------------------------------------------------------------
# TAB: ESCANEAR
# ----------------------------------------------------------------
with tab_scan:
  if not operario:
    st.warning(
        "Ingresa tu nombre en la barra lateral antes de registrar escaneos."
    )

  st.markdown("#### Escanea el QR de la caja")
  codigo_leido = qrcode_scanner(key="scanner")

  col_m1, col_m2 = st.columns([3, 1])
  codigo_manual = col_m1.text_input(
      "O ingresa el código manualmente", key="manual_input"
  )
  buscar_manual = col_m2.button("🔍 Buscar", use_container_width=True)

  if (
      codigo_leido
      and codigo_leido.strip()
      != st.session_state.get("ultimo_escaneo_procesado")
  ):
    st.session_state["codigo_actual"] = codigo_leido.strip().upper()
  if buscar_manual and codigo_manual.strip():
    st.session_state["codigo_actual"] = codigo_manual.strip().upper()

  codigo_actual = st.session_state.get("codigo_actual")

  if codigo_actual:
    cat_row = catalogo_df[catalogo_df["codigo"] == codigo_actual]

    if cat_row.empty:
      st.error(
          f"❌ Código **{codigo_actual}** no está en el catálogo (hoja Receta)."
      )
    else:
      producto = cat_row.iloc[0]["producto"]
      ped_row = pedido_df[pedido_df["codigo"] == codigo_actual]

      if ped_row.empty:
        st.warning(
            f"⚠️ **{producto}** no tiene pedido registrado hoy. Se puede"
            " registrar solo para trazabilidad."
        )
        factor, umi, requerimiento = (
            cat_row.iloc[0]["factor"],
            cat_row.iloc[0]["umr"],
            None,
        )
      else:
        factor = ped_row.iloc[0]["factor"]
        umi = ped_row.iloc[0]["umi"]
        requerimiento = ped_row.iloc[0]["requerimiento"]

      st.markdown(
          f"""<div class="producto-card">
                  <h3>{producto}</h3>
                  <p>Código: <b>{codigo_actual}</b> &nbsp;|&nbsp; Factor: <b>{fmt_num(factor)}</b>
                  &nbsp;|&nbsp; U. Medida: <b>{umi}</b></p>
              </div>""",
          unsafe_allow_html=True,
      )

      if requerimiento:
        fila_av = avance_df[avance_df["codigo"] == codigo_actual]
        ya_entregado = (
            float(fila_av["unidades_entregadas"].iloc[0])
            if not fila_av.empty
            else 0.0
        )
        pct_actual = (
            (ya_entregado / requerimiento * 100) if requerimiento else 0
        )
        st.progress(
            min(pct_actual / 100, 1.0),
            text=(
                f"{fmt_num(ya_entregado)} / {fmt_num(requerimiento)} {umi}"
                f" entregadas ({pct_actual:.1f}%)"
            ),
        )

      cajas = st.number_input(
          "Cantidad de cajas escaneadas / empacadas",
          min_value=0.0,
          step=1.0,
          value=1.0,
          key=f"cajas_{codigo_actual}",
      )
      unidades_calc = cajas * float(factor)
      st.caption(f"= {fmt_num(unidades_calc)} {umi}")

      cb1, cb2 = st.columns(2)
      if cb1.button(
          "✅ Registrar entrega", type="primary", use_container_width=True
      ):
        registrar_escaneo(
            codigo_actual,
            producto,
            cajas,
            float(factor),
            unidades_calc,
            umi,
            operario or "Sin nombre",
        )
        st.session_state["ultimo_escaneo_procesado"] = codigo_actual
        st.session_state["codigo_actual"] = None
        st.success(
            f"Registrado: {fmt_num(cajas)} cajas de {producto}"
            f" ({fmt_num(unidades_calc)} {umi})."
        )
        st.rerun()
      if cb2.button("✖️ Cancelar", use_container_width=True):
        st.session_state["codigo_actual"] = None
        st.rerun()

# ----------------------------------------------------------------
# TAB: AVANCE DEL PEDIDO
# ----------------------------------------------------------------
with tab_avance:
  col_f1, col_f2 = st.columns([2, 1])
  busqueda = col_f1.text_input("🔎 Buscar por código o descripción", "")
  filtro_estado = col_f2.multiselect(
      "Estado",
      ["⛔ Sin iniciar", "🟡 Parcial", "✅ Completo"],
      default=["⛔ Sin iniciar", "🟡 Parcial", "✅ Completo"],
  )

  df_show = avance_df.copy()
  if busqueda:
    mask = df_show["codigo"].str.contains(
        busqueda, case=False, na=False
    ) | df_show["descripcion"].str.contains(busqueda, case=False, na=False)
    df_show = df_show[mask]
  df_show = df_show[df_show["estado"].isin(filtro_estado)]

  colores_estado = {
      "⛔ Sin iniciar": C_BAD,
      "🟡 Parcial": C_WARN,
      "✅ Completo": C_OK,
  }

  for _, row in df_show.iterrows():
    c1, c2, c3 = st.columns([3, 3, 1])
    c1.markdown(f"**{row['descripcion']}**  \n`{row['codigo']}`")
    c2.progress(
        min(row["pct_avance"] / 100, 1.0),
        text=(
            f"{fmt_num(row['unidades_entregadas'])}/{fmt_num(row['requerimiento'])}"
            f" {row['umi']} — {row['pct_avance']:.1f}%"
        ),
    )
    c3.markdown(
        f"<span class='badge'"
        f" style='background-color:{colores_estado[row['estado']]}'>{row['estado'].split()[0]}</span>",
        unsafe_allow_html=True,
    )

  if df_show.empty:
    st.info("No hay productos que coincidan con el filtro.")

  st.markdown("---")
  fig = go.Figure(
      data=[
          go.Pie(
              labels=["✅ Completo", "🟡 Parcial", "⛔ Sin iniciar"],
              values=[
                  int((avance_df["pct_avance"] >= 100).sum()),
                  int(
                      (
                          (avance_df["pct_avance"] > 0)
                          & (avance_df["pct_avance"] < 100)
                      ).sum()
                  ),
                  int((avance_df["pct_avance"] <= 0).sum()),
              ],
              hole=0.55,
              marker_colors=[C_OK, C_WARN, C_BAD],
          )
      ]
  )
  fig.update_layout(
      title="Distribución de productos por estado", height=380, showlegend=True
  )
  st.plotly_chart(fig, use_container_width=True)

# ----------------------------------------------------------------
# TAB: HISTORIAL
# ----------------------------------------------------------------
with tab_hist:
  if registros_df.empty:
    st.info("Aún no hay escaneos registrados.")
  else:
    operarios = ["Todos"] + sorted(
        registros_df["operario"].dropna().unique().tolist()
    )
    filtro_op = st.selectbox("Filtrar por operario", operarios)
    df_hist = (
        registros_df
        if filtro_op == "Todos"
        else registros_df[registros_df["operario"] == filtro_op]
    )

    st.dataframe(
        df_hist[[
            "id",
            "timestamp",
            "operario",
            "codigo",
            "producto",
            "cantidad_cajas",
            "unidades",
            "umi",
        ]],
        use_container_width=True,
        hide_index=True,
    )

    st.markdown("#### Eliminar / corregir un registro")
    opciones_borrado = {
        f"#{r.id} — {r.timestamp} — {r.producto}"
        f" ({fmt_num(r.cantidad_cajas)} cajas, {r.operario})": r.id
        for r in df_hist.itertuples()
    }
    if opciones_borrado:
      sel = st.selectbox(
          "Selecciona el registro a eliminar", list(opciones_borrado.keys())
      )
      if st.button("🗑️ Eliminar registro seleccionado"):
        eliminar_registro(opciones_borrado[sel])
        st.success("Registro eliminado.")
        st.rerun()

# ----------------------------------------------------------------
# TAB: ETIQUETAS
# ----------------------------------------------------------------
with tab_labels:
  st.markdown(
      "Genera etiquetas QR para pegar en las cajas. El valor del QR es el"
      " mismo **código SAP** del producto."
  )
  opciones = (catalogo_df["codigo"] + " — " + catalogo_df["producto"]).tolist()
  seleccion = st.multiselect(
      "Selecciona productos (vacío = todos los del pedido de hoy)", opciones
  )
  columnas_hoja = st.slider(
      "Columnas por hoja", min_value=2, max_value=5, value=3
  )

  if st.button("🏷️ Generar hoja de etiquetas"):
    if seleccion:
      codigos_sel = [s.split(" — ")[0] for s in seleccion]
    else:
      codigos_sel = pedido_df["codigo"].tolist()

    productos = catalogo_df[catalogo_df["codigo"].isin(codigos_sel)][
        ["codigo", "producto"]
    ].values.tolist()
    if not productos:
      st.warning("No se encontraron productos para generar etiquetas.")
    else:
      hoja = generar_hoja_etiquetas(productos, columnas=columnas_hoja)
      buf = io.BytesIO()
      hoja.save(buf, format="PNG")
      st.image(
          hoja,
          caption=f"{len(productos)} etiqueta(s) generada(s)",
          use_container_width=True,
      )
      st.download_button(
          "⬇️ Descargar hoja de etiquetas (PNG, para imprimir)",
          data=buf.getvalue(),
          file_name="etiquetas_qr.png",
          mime="image/png",
      )

# ----------------------------------------------------------------
# TAB: FACTOR M
# ----------------------------------------------------------------
with tab_factor_m:
  st.markdown("### ⚙️ Requerimiento de Componentes (Factor M)")
  st.markdown(
      "Análisis de componentes cuyos códigos empiezan con **M**, cruzando el"
      " pedido con las recetas y aplicando el factor de redondeo de la hoja"
      " **factor-lima**."
  )

  if "archivo_bytes_actual" in st.session_state:
    df_factor_m = calcular_componentes_factor_m(
        pedido_df,
        archivo_subido=io.BytesIO(st.session_state["archivo_bytes_actual"]),
    )
    if not df_factor_m.empty:
      st.dataframe(df_factor_m, use_container_width=True, hide_index=True)

      buf_fm = io.BytesIO()
      with pd.ExcelWriter(buf_fm, engine="openpyxl") as writer:
        df_factor_m.to_excel(
            writer, sheet_name="Requerimiento_Factor_M", index=False
        )
      st.download_button(
          "📥 Descargar reporte de Factor M en Excel",
          data=buf_fm.getvalue(),
          file_name=f"factor_m_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
          mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
          type="primary",
      )
    else:
      st.info(
          "No se encontraron componentes con código 'M' o faltan datos para"
          " calcular."
      )
  else:
    st.info(
        "👈 Por favor, carga el archivo Excel del pedido en la barra lateral"
        " para procesar este reporte."
    )

# ----------------------------------------------------------------
# TAB: EXPORTAR
# ----------------------------------------------------------------
with tab_export:
  st.markdown(
      "Exporta el detalle de avance y el historial completo de escaneos a un"
      " Excel."
  )
  resumen_export = avance_df.rename(
      columns={
          "codigo": "Código",
          "descripcion": "Descripción",
          "umi": "U. Medida",
          "factor": "Factor",
          "requerimiento": "Requerimiento",
          "unidades_entregadas": "Entregado",
          "cajas_entregadas": "Cajas entregadas",
          "pct_avance": "% Avance",
          "estado": "Estado",
          "fecha_pedido": "Fecha pedido",
      }
  )
  st.dataframe(resumen_export, use_container_width=True, hide_index=True)

  buf = io.BytesIO()
  with pd.ExcelWriter(buf, engine="openpyxl") as writer:
    resumen_export.to_excel(writer, sheet_name="Resumen Avance", index=False)
    registros_df.to_excel(writer, sheet_name="Historial Escaneos", index=False)
  st.download_button(
      "📥 Descargar Excel de entrega",
      data=buf.getvalue(),
      file_name=f"entrega_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
      mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      type="primary",
  )
