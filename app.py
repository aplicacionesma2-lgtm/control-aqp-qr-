"""
Control de Empaque QR — María Almenara
========================================
"""

from datetime import datetime, timezone, timedelta
import io
import math
import uuid

from google.oauth2.service_account import Credentials
import gspread
from PIL import Image, ImageDraw, ImageFont
import pandas as pd
import streamlit as st
from streamlit_qrcode_scanner import qrcode_scanner
import qrcode

st.set_page_config(
    page_title="Control de Empaque QR — María Almenara",
    page_icon="📦",
    layout="wide",
    initial_sidebar_state="expanded",
)

# Zona horaria oficial para Lima, Perú (UTC-5)
TZ_LIMA = timezone(timedelta(hours=-5))

def ahora_lima():
  return datetime.now(TZ_LIMA)

# --- GESTIÓN DE TEMA (CLARO / OSCURO) ---
if "tema" not in st.session_state:
  st.session_state["tema"] = "Claro"

with st.sidebar:
  st.markdown("### 🎨 Apariencia")
  st.session_state["tema"] = st.radio(
      "Seleccionar Modo", ["Claro", "Oscuro"], index=0 if st.session_state["tema"] == "Claro" else 1, horizontal=True
  )
  st.markdown("---")

# Paletas de colores adaptativas con acento corporativo de María Almenara (Rosa/Fucsia: #D4145A)
if st.session_state["tema"] == "Oscuro":
  C_BG = "#1A252F"
  C_SIDEBAR = "#11181E"
  C_CARD_BG = "#212F3D"
  C_TEXT = "#F2F4F4"
  C_PRIMARY = "#D4145A"
  C_PRIMARY_DARK = "#B5104C"
  C_ACCENT = "#E64A81"
  CSS_THEME_EXTRA = "color: #F2F4F4 !important;"
else:
  C_BG = "#F8F9FA"
  C_SIDEBAR = "#154360"
  C_CARD_BG = "#FFFFFF"
  C_TEXT = "#2C3E50"
  C_PRIMARY = "#D4145A"
  C_PRIMARY_DARK = "#B5104C"
  C_ACCENT = "#E64A81"
  CSS_THEME_EXTRA = ""

st.markdown(
    f"""
    <style>
    .stApp {{ background-color: {C_BG}; {CSS_THEME_EXTRA} }}
    section[data-testid="stSidebar"] {{ background-color: {C_SIDEBAR}; }}
    section[data-testid="stSidebar"] * {{ color: #F2F2F2 !important; }}
    div[data-testid="stMetric"] {{
        background-color: {C_CARD_BG}; border: 1px solid #FADBD8;
        border-left: 6px solid {C_PRIMARY}; border-radius: 12px; padding: 14px 18px;
        box-shadow: 0 4px 6px rgba(0,0,0,0.08);
    }}
    .stButton>button {{
        background-color: {C_PRIMARY};
        color: white;
        border-radius: 10px;
        border: none;
        padding: 10px 20px;
        font-weight: bold;
        min-height: 48px;
        width: 100%;
    }}
    .stButton>button:hover {{
        background-color: {C_ACCENT};
        color: white;
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
HISTORIAL_CIERRES_HEADERS = [
    "id_cierre",
    "timestamp_cierre",
    "operario_cierre",
    "total_registros",
    "total_unidades",
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
  for col in df.columns:
    if col in numericas:
      df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)
    else:
      df[col] = df[col].astype(str).str.strip()
      df[col] = df[col].mask(
          df[col].str.lower().isin(["nan", "none", "nat"]), ""
      )
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
      ahora_lima().strftime("%Y-%m-%d %H:%M:%S"),
  ]
  ws.append_row(fila, value_input_option="RAW")
  st.cache_data.clear()


def eliminar_registro(reg_id: str):
  ws = get_ws("Registros", tuple(REGISTROS_HEADERS))
  celda = ws.find(str(reg_id), in_column=1)
  if celda:
    ws.delete_rows(celda.row)
  st.cache_data.clear()


def actualizar_registro(reg_id: str, nuevas_cajas: float, nuevo_factor: float, operario):
  ws = get_ws("Registros", tuple(REGISTROS_HEADERS))
  celda = ws.find(str(reg_id), in_column=1)
  if celda:
    fila_num = celda.row
    nuevas_unidades = nuevas_cajas * nuevo_factor
    ws.update_cell(fila_num, 4, float(nuevas_cajas))
    ws.update_cell(fila_num, 6, float(nuevas_unidades))
    ws.update_cell(fila_num, 8, str(operario))
    ws.update_cell(fila_num, 9, ahora_lima().strftime("%Y-%m-%d %H:%M:%S"))
  st.cache_data.clear()


def respaldo_diario_turno(operario_actual):
  registros_actuales = cargar_registros()
  if registros_actuales.empty:
    return False, "No hay registros activos para respaldar."

  ws_cierres = get_ws("HistorialCierres", tuple(HISTORIAL_CIERRES_HEADERS))
  id_cierre = uuid.uuid4().hex[:8]
  timestamp_cierre = ahora_lima().strftime("%Y-%m-%d %H:%M:%S")
  total_regs = len(registros_actuales)
  total_unids = float(registros_actuales["unidades"].sum())

  ws_cierres.append_row(
      [
          f"RESPALDO-{id_cierre}",
          timestamp_cierre,
          str(operario_actual or "Supervisor"),
          total_regs,
          total_unids,
      ],
      value_input_option="RAW",
  )
  return True, id_cierre


def cerrar_pedido_semanal(operario_actual):
  registros_actuales = cargar_registros()
  if registros_actuales.empty:
    return False, "No hay registros para cerrar."

  ws_cierres = get_ws("HistorialCierres", tuple(HISTORIAL_CIERRES_HEADERS))
  id_cierre = uuid.uuid4().hex[:8]
  timestamp_cierre = ahora_lima().strftime("%Y-%m-%d %H:%M:%S")
  total_regs = len(registros_actuales)
  total_unids = float(registros_actuales["unidades"].sum())

  ws_cierres.append_row(
      [
          f"CIERRE_SEMANAL-{id_cierre}",
          timestamp_cierre,
          str(operario_actual or "Supervisor"),
          total_regs,
          total_unids,
      ],
      value_input_option="RAW",
  )

  reiniciar_registros()
  return True, id_cierre


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


def obtener_componentes_producto(codigo_prod: str, cajas: float, archivo_subido=None) -> pd.DataFrame:
  """Extrae los componentes de la hoja 'Receta' haciendo match exacto con el código del producto."""
  try:
    if archivo_subido is not None:
      xl = pd.ExcelFile(archivo_subido)
      if "Receta" not in xl.sheet_names:
        return pd.DataFrame()
      df_receta = xl.parse("Receta")
    else:
      return pd.DataFrame()
  except Exception:
    return pd.DataFrame()

  if df_receta.empty:
    return pd.DataFrame()

  # Limpiar nombres de columnas y datos
  df_receta.columns = [str(c).strip().upper() for c in df_receta.columns]
  codigo_buscado = str(codigo_prod).strip().upper()

  # Buscar la columna que contiene el código del producto principal en la hoja Receta
  # (comúnmente llamada CÓDIGO, COD_PROD, ARTICULO, etc.)
  col_prod = next(
      (c for c in df_receta.columns if any(k in c for k in ["CÓDIGO", "CODIGO", "PROD", "ART"])),
      df_receta.columns[0]
  )

  # Normalizar la columna de productos para el match
  df_receta[col_prod] = df_receta[col_prod].fillna("").astype(str).str.strip().str.upper()

  # Filtrar estrictamente por el producto buscado
  matches = df_receta[df_receta[col_prod] == codigo_buscado]
  if matches.empty:
    return pd.DataFrame()

  # Identificar la columna del código del componente / insumo
  col_cod_comp = next(
      (c for c in df_receta.columns if any(k in c for k in ["COMPONENTE", "INSUMO", "MATERIA", "HIJO"])),
      df_receta.columns[1] if len(df_receta.columns) > 1 else df_receta.columns[0]
  )

  # Identificar la columna de descripción del componente
  col_desc_comp = next(
      (c for c in df_receta.columns if any(k in c for k in ["DESCRIPCIÓN", "DESCRIPCION", "NOMBRE", "DETALLE"]) and c != col_prod),
      col_cod_comp
  )

  # Identificar la columna de cantidad base
  col_cant = next(
      (c for c in df_receta.columns if any(k in c for k in ["CANT", "QTY", "UNID", "CONSUMO"])),
      None
  )

  resultados = []
  for _, row in matches.iterrows():
    c_comp = str(row.get(col_cod_comp, "")).strip()
    d_comp = str(row.get(col_desc_comp, c_comp)).strip()
    
    # Omitir filas vacías
    if not c_comp or c_comp.upper() == "NAN":
      continue

    cant_base = 1.0
    if col_cant:
      try:
        cant_base = float(row.get(col_cant, 1.0))
      except (TypeError, ValueError):
        cant_base = 1.0

    cant_total = cant_base * float(cajas)
    resultados.append({
        "Código Componente": c_comp,
        "Descripción Componente": d_comp,
        "Cantidad": cant_total
    })

  return pd.DataFrame(resultados).drop_duplicates().reset_index(drop=True)


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

  col_req_data = next((c for c in df_data.columns if "REQUERIMIENTO" in c), "REQUERIMIENTO")
  col_cod_data = next((c for c in df_data.columns if "CÓDIGO" in c or "CODIGO" in c), "CÓDIGO")

  df_data_limpio = df_data[[col_cod_data, col_req_data]].copy()
  df_data_limpio.columns = ["CÓDIGO_PROD", "REQ_REAL"]
  df_data_limpio["REQ_REAL"] = pd.to_numeric(df_data_limpio["REQ_REAL"], errors="coerce").fillna(0)
  df_data_limpio["CÓDIGO_PROD"] = df_data_limpio["CÓDIGO_PROD"].fillna("").astype(str).str.strip()

  col_cod_receta = next((c for c in df_receta.columns if "CÓDIGO" in c or "CODIGO" in c), "CÓDIGO")
  col_cod_comp = next((c for c in df_receta.columns if "COD" in c and ("COMPONENTE" in c or "COMP" in c)), None)
  col_desc_comp = next((c for c in df_receta.columns if c != col_cod_comp and "COMPONENTE" in c), None)

  if not col_cod_comp:
    col_cod_comp = next((c for c in df_receta.columns if "COMPONENTE" in c), df_receta.columns[1])
  if not col_desc_comp:
    col_desc_comp = col_cod_comp

  df_receta[col_cod_receta] = df_receta[col_cod_receta].fillna("").astype(str).str.strip()
  df_receta[col_cod_comp] = df_receta[col_cod_comp].fillna("").astype(str).str.strip()
  df_receta[col_desc_comp] = df_receta[col_desc_comp].fillna("").astype(str).str.strip()

  df_merged = pd.merge(
      df_receta,
      df_data_limpio,
      left_on=col_cod_receta,
      right_on="CÓDIGO_PROD",
      how="inner",
  )

  df_merged["REQ_COMPONENTE"] = df_merged["REQ_REAL"]

  s_cod = df_merged[col_cod_comp]
  df_m = df_merged[
      s_cod.str.upper().str.startswith("M")
      & (s_cod.str.upper() != "NAN")
      & (s_cod != "")
  ].copy()

  df_m["CÓDIGO_EXTRACTO"] = df_m[col_cod_comp]
  df_m["DESCRIPCIÓN_COMP"] = df_m[col_desc_comp].mask(
      df_m[col_desc_comp] == "", df_m["CÓDIGO_EXTRACTO"]
  )

  col_cod_f = next((c for c in df_factor.columns if "COD" in c), df_factor.columns[0])
  col_fac_f = next((c for c in df_factor.columns if "FACTOR" in c), df_factor.columns[-1])
  df_factor = df_factor.rename(columns={col_cod_f: "CÓDIGO_LIMA", col_fac_f: "FACTOR_VALOR"})
  df_factor["CÓDIGO_LIMA"] = df_factor["CÓDIGO_LIMA"].fillna("").astype(str).str.strip()

  df_final = pd.merge(
      df_m,
      df_factor,
      left_on="CÓDIGO_EXTRACTO",
      right_on="CÓDIGO_LIMA",
      how="left",
  )
  df_final["FACTOR_LIMA"] = pd.to_numeric(df_final["FACTOR_VALOR"], errors="coerce").fillna(1.0)
  df_final["FACTOR_LIMA"] = df_final["FACTOR_LIMA"].apply(lambda x: x if x > 0 else 1.0)

  df_grouped = (
      df_final.groupby(["CÓDIGO_EXTRACTO", "DESCRIPCIÓN_COMP", "FACTOR_LIMA"], as_index=False)["REQ_COMPONENTE"]
      .sum()
      .rename(columns={"REQ_COMPONENTE": "REQ_TOTAL"})
  )

  df_grouped["REQ_REDONDEADO"] = df_grouped.apply(
      lambda row: math.ceil(row["REQ_TOTAL"] / row["FACTOR_LIMA"]) * row["FACTOR_LIMA"],
      axis=1,
  )

  return (
      df_grouped[["CÓDIGO_EXTRACTO", "DESCRIPCIÓN_COMP", "REQ_TOTAL", "FACTOR_LIMA", "REQ_REDONDEADO"]]
      .rename(
          columns={
              "CÓDIGO_EXTRACTO": "CÓDIGO",
              "DESCRIPCIÓN_COMP": "DESCRIPCIÓN",
              "REQ_TOTAL": "REQUERIMIENTO NETO",
              "FACTOR_LIMA": "FACTOR LIMA",
              "REQ_REDONDEADO": "REQUERIMIENTO REDONDEADO",
          }
      )
      .sort_values("CÓDIGO")
      .reset_index(drop=True)
  )


def generar_imagen_etiqueta(codigo, producto, cajas, unidades, umi, operario):
  img = Image.new("RGB", (600, 400), color="white")
  d = ImageDraw.Draw(img)

  d.rectangle([0, 0, 600, 70], fill=C_PRIMARY_DARK)
  d.text((20, 20), "MARÍA ALMENARA - CONTROL QR", fill="white")

  d.text((20, 90), f"CÓDIGO: {codigo}", fill="black")
  prod_cortado = producto if len(producto) <= 40 else producto[:37] + "..."
  d.text((20, 125), f"PRODUCTO: {prod_cortado}", fill="black")

  d.text(
      (20, 165),
      f"CANTIDAD: {fmt_num(cajas)} Cajas ({fmt_num(unidades)} {umi})",
      fill="black",
  )
  d.text((20, 200), f"OPERARIO: {operario}", fill="gray")
  d.text((20, 230), f"FECHA: {ahora_lima().strftime('%Y-%m-%d %H:%M')}", fill="gray")

  qr = qrcode.QRCode(box_size=4, border=1)
  qr.add_data(str(codigo))
  qr.make(fit=True)
  qr_img = qr.make_image(fill_color="black", back_color="white").convert("RGB")
  img.paste(qr_img, (420, 210))

  buf = io.BytesIO()
  img.save(buf, format="PNG")
  buf.seek(0)
  return buf


# --- INTERFAZ ---
try:
  with st.sidebar:
    st.markdown("## 📦 Control de Empaque QR")
    operario = st.text_input("👤 Operario", value=st.session_state.get("operario", ""))
    st.session_state["operario"] = operario
    st.markdown("---")

    archivo = st.file_uploader(
        "Excel con hojas 'data', 'Receta' y 'factor-lima'", type=["xlsx"]
    )
    if archivo is not None:
      firma = f"{archivo.name}-{archivo.size}"
      if firma != st.session_state.get("ultimo_archivo_cargado"):
        xl_temp = pd.ExcelFile(archivo)
        if {"data", "Receta", "factor-lima"}.issubset(set(xl_temp.sheet_names)):
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

    st.markdown("---")
    st.markdown("### 🗂️ Control de Turnos y Pedido")
    
    if st.button("📑 Registrar Respaldo Diario (Turno)"):
      exito, msg = respaldo_diario_turno(operario)
      if exito:
        st.success(f"Respaldo diario guardado (ID: {msg}).")
      else:
        st.warning(msg)

    if st.button("🔒 Cerrar Pedido Semanal y Reiniciar"):
      exito, msg = cerrar_pedido_semanal(operario)
      if exito:
        st.success(f"Pedido semanal cerrado. ID: {msg}")
        st.rerun()
      else:
        st.warning(msg)

    if st.button("🔄 Limpiar Caché / Forzar Recarga"):
      st.cache_data.clear()
      st.cache_resource.clear()
      st.success("Caché limpia.")
      st.rerun()

  if pedido_df.empty or catalogo_df.empty:
    st.title("📦 Control de Empaque QR — María Almenara")
    st.info("👈 Sube el archivo Excel en la barra lateral para iniciar.")
    st.stop()

  registros_df = cargar_registros()
  avance_df = calcular_avance(pedido_df, registros_df)

  st.title("📦 Control de Empaque QR — María Almenara")
  
  tab1, tab2, tab3, tab4, tab5 = st.tabs([
      "📷 Escanear",
      "📊 Avance",
      "🕒 Historial",
      "🏷️ Etiquetas",
      "⚙️ Factor M",
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
            else cat_row["umr"]
        )

        avance_prod = avance_df[avance_df["codigo"] == codigo_actual]
        if not avance_prod.empty:
          p_av = float(avance_prod.iloc[0]["pct_avance"])
          u_ent = float(avance_prod.iloc[0]["unidades_entregadas"])
          req_tot = float(avance_prod.iloc[0]["requerimiento"])
          if p_av >= 100:
            st.warning(
                f"⚠️ **¡ALERTA! Este producto ya está cubierto al {p_av:.1f}%"
                f" ({fmt_num(u_ent)} / {fmt_num(req_tot)} {umi}).**"
            )

        st.markdown(f"**Producto:** {prod} (`{codigo_actual}`)")
        st.markdown(f"**Factor:** {fmt_num(factor)}")

        cajas = st.number_input("Cajas", min_value=0.0, value=1.0, step=1.0)
        unidades = cajas * float(factor)
        st.caption(f"= {fmt_num(unidades)} {umi}")

        # --- DESGLOSE DETALLADO DE COMPONENTES DE LA RECETA ---
        if "archivo_bytes_actual" in st.session_state:
          df_comp_prod = obtener_componentes_producto(
              codigo_actual, 
              cajas,
              io.BytesIO(st.session_state["archivo_bytes_actual"])
          )
          if not df_comp_prod.empty:
            st.markdown("### 📋 Componentes / Insumos:")
            for _, r in df_comp_prod.iterrows():
              st.markdown(f"- **{r['Código Componente']}** {r['Descripción Componente']} **({fmt_num(r['Cantidad'])})**")

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
    st.markdown("### 📊 Avance General del Pedido")
    
    total_productos = len(avance_df)
    if total_productos > 0:
      productos_completados = len(avance_df[avance_df["pct_avance"] >= 100])
      pct_global_productos = (productos_completados / total_productos) * 100
    else:
      pct_global_productos = 0.0
      productos_completados = 0

    color_barra = "#27AE60" if pct_global_productos >= 100 else C_PRIMARY

    st.markdown(
        f"""
        <div style="background-color: {C_CARD_BG}; border: 1px solid #FADBD8; border-left: 8px solid {color_barra}; border-radius: 14px; padding: 22px 26px; box-shadow: 0 6px 12px rgba(0,0,0,0.08); margin-bottom: 25px;">
            <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 12px;">
                <div>
                    <span style="font-size: 13px; font-weight: 700; color: #7F8C8D; letter-spacing: 1px;">🚀 AVANCE GLOBAL DE PRODUCTOS</span>
                    <h1 style="font-size: 42px; font-weight: 800; color: {C_TEXT}; margin: 0; line-height: 1.1;">{pct_global_productos:.1f}%</h1>
                </div>
                <div style="text-align: right;">
                    <span style="background-color: #FDEDEC; color: #D4145A; padding: 6px 14px; border-radius: 20px; font-weight: bold; font-size: 14px; border: 1px solid #F5B7B1;">
                        ✓ {productos_completados} de {total_productos} productos al 100%
                    </span>
                </div>
            </div>
            <div style="background-color: #EAEDED; border-radius: 10px; width: 100%; height: 24px; overflow: hidden; box-shadow: inset 0 1px 3px rgba(0,0,0,0.1);">
                <div style="background-color: {color_barra}; width: {min(pct_global_productos, 100.0)}%; height: 100%; border-radius: 10px; transition: width 0.5s ease-in-out;"></div>
            </div>
        </div>
        """,
        unsafe_allow_html=True
    )

    col_dl1, col_dl2 = [st.columns([4, 1])[0], st.columns([4, 1])[1]]
    with col_dl2:
      buffer_reporte = io.BytesIO()
      with pd.ExcelWriter(buffer_reporte, engine="openpyxl") as writer:
        avance_df.to_excel(writer, index=False, sheet_name="Avance de Produccion")
        if not registros_df.empty:
          registros_df.to_excel(writer, index=False, sheet_name="Detalle Registros")
      buffer_reporte.seek(0)

      st.download_button(
          label="📥 Descargar Reporte",
          data=buffer_reporte,
          file_name=(
              "reporte_consolidado_empaque_"
              f"{ahora_lima().strftime('%Y%m%d_%H%M%S')}.xlsx"
          ),
          mime=(
              "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
          ),
      )

    st.markdown("---")
    st.dataframe(avance_df, use_container_width=True, hide_index=True)

  with tab3:
    st.markdown("### 🕒 Historial de Registros, Filtros y Edición")
    
    if registros_df.empty:
      st.info("No hay registros aún.")
    else:
      f_col1, f_col2, f_col3 = st.columns(3)
      with f_col1:
        filtro_codigo = st.text_input("🔍 Filtrar por Código", "").strip().upper()
      with f_col2:
        filtro_producto = st.text_input("🔍 Filtrar por Producto", "").strip().upper()
      with f_col3:
        filtro_fecha = st.text_input("📅 Filtrar por Fecha (YYYY-MM-DD)", "").strip()

      df_filtrado = registros_df.copy()
      if filtro_codigo:
        df_filtrado = df_filtrado[df_filtrado["codigo"].str.upper().str.contains(filtro_codigo)]
      if filtro_producto:
        df_filtrado = df_filtrado[df_filtrado["producto"].str.upper().str.contains(filtro_producto)]
      if filtro_fecha:
        df_filtrado = df_filtrado[df_filtrado["timestamp"].str.contains(filtro_fecha)]

      st.markdown(f"Mostrando **{len(df_filtrado)}** de **{len(registros_df)}** registros totales.")

      df_editable = df_filtrado.copy()
      df_editable.insert(0, "Seleccionar", False)

      edited_df = st.data_editor(
          df_editable,
          use_container_width=True,
          hide_index=True,
          disabled=[col for col in df_editable.columns if col != "Seleccionar"]
      )

      seleccionados = edited_df[edited_df["Seleccionar"] == True]

      if not seleccionados.empty:
        st.markdown("---")
        st.markdown(f"🛠️ **Acciones para los {len(seleccionados)} registros seleccionados:**")
        
        reg_id_activo = seleccionados.iloc[0]["id"]
        prod_activo = seleccionados.iloc[0]["producto"]
        cajas_actuales = float(seleccionados.iloc[0]["cantidad_cajas"])
        factor_activo = float(seleccionados.iloc[0]["factor"])

        col_act1, col_act2, col_act3 = st.columns([2, 1, 1])
        with col_act1:
          st.info(f"Editando: **{prod_activo}**")
          nueva_caja_val = st.number_input("Nueva cantidad de Cajas", min_value=0.0, value=cajas_actuales, step=1.0, key="edit_cajas_val")
        with col_act2:
          st.markdown("<br>", unsafe_allow_html=True)
          if st.button("💾 Guardar Edición"):
            actualizar_registro(reg_id_activo, nueva_caja_val, factor_activo, operario or "Operario")
            st.success("¡Registro actualizado!")
            st.rerun()
        with col_act3:
          st.markdown("<br>", unsafe_allow_html=True)
          if st.button("🗑️ Eliminar Seleccionados", type="primary"):
            for _, r in seleccionados.iterrows():
              eliminar_registro(r["id"])
            st.success("¡Registros eliminados con éxito!")
            st.rerun()

  with tab4:
    st.markdown("### 🏷️ Generador de Etiquetas QR Múltiples")
    st.markdown("Selecciona uno o varios productos para generar y descargar sus etiquetas en lote.")

    opciones_prod = [
        f"{row['codigo']} - {row['producto']}"
        for _, row in catalogo_df.iterrows()
    ]
    productos_seleccionados = st.multiselect("Elegir Productos", opciones_prod)

    if productos_seleccionados:
      num_cajas_lote = st.number_input(
          "Cantidad de Cajas por Defecto para el Lote seleccionado",
          min_value=1.0,
          value=1.0,
          step=1.0,
      )

      if st.button("🖨️ Generar Lote de Etiquetas (PNG)", type="primary"):
        for item_str in productos_seleccionados:
          codigo_sel = item_str.split(" - ")[0].strip()
          cat_row = catalogo_df[catalogo_df["codigo"] == codigo_sel].iloc[0]
          ped_row = pedido_df[pedido_df["codigo"] == codigo_sel]

          prod_nombre = cat_row["producto"]
          factor_val = (
              ped_row.iloc[0]["factor"]
              if not ped_row.empty
              else cat_row["factor"]
          )
          umi_val = (
              ped_row.iloc[0]["umi"] if not ped_row.empty else cat_row["umr"]
          )
          total_unidades = num_cajas_lote * float(factor_val)

          img_buffer = generar_imagen_etiqueta(
              codigo_sel,
              prod_nombre,
              num_cajas_lote,
              total_unidades,
              umi_val,
              operario or "Operario",
          )
          st.image(
              img_buffer, caption=f"Etiqueta - {codigo_sel}: {prod_nombre}", width=400
          )
          st.download_button(
              label=f"⬇️ Descargar Etiqueta [{codigo_sel}]",
              data=img_buffer,
              file_name=(
                  f"etiqueta_{codigo_sel}_{ahora_lima().strftime('%Y%m%d_%H%M%S')}.png"
              ),
              mime="image/png",
              key=f"dl_{codigo_sel}_{uuid.uuid4().hex[:4]}"
          )
          st.markdown("---")

  with tab5:
    st.markdown("### ⚙️ Requerimiento de Componentes (Factor M)")
    if "archivo_bytes_actual" in st.session_state:
      df_fm = calcular_componentes_factor_m(
          pedido_df,
          archivo_subido=io.BytesIO(st.session_state["archivo_bytes_actual"]),
      )
      if not df_fm.empty:
        st.dataframe(df_fm, use_container_width=True, hide_index=True)

        buffer_excel = io.BytesIO()
        with pd.ExcelWriter(buffer_excel, engine="openpyxl") as writer:
          df_fm.to_excel(writer, index=False, sheet_name="Factor M")
        buffer_excel.seek(0)

        st.download_button(
            label="📥 Descargar Factor M a Excel",
            data=buffer_excel,
            file_name=(
                "requerimiento_factor_m_"
                f"{ahora_lima().strftime('%Y%m%d_%H%M%S')}.xlsx"
            ),
            mime=(
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            ),
        )
      else:
        st.info("No hay componentes M o faltan columnas.")
    else:
      st.info("Carga el Excel del pedido en la barra lateral.")

except Exception as e:
  st.error(f"Se ha producido un error al ejecutar la aplicación: {e}")
  if st.button("🗑️ Limpiar Todo y Reiniciar"):
    st.cache_data.clear()
    st.cache_resource.clear()
    st.rerun()
