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

  # Cruzamos la hoja 'data' (que tiene el REQUERIMIENTO de la orden) con la 'Receta'
  df_merged = pd.merge(
      df_data,
      df_receta,
      on="CÓDIGO",
      how="inner",
      suffixes=("_PEDIDO", "_RECETA"),
  )

  # Si la receta tiene una proporción por unidad, la multiplicamos por el REQUERIMIENTO del pedido (ej. 2160)
  # Verificamos si existe la columna REQUERIMIENTO y CANTIDAD en la receta
  req_col = next((c for c in df_merged.columns if "REQUERIMIENTO" in c), None)
  cant_col = next((c for c in df_merged.columns if "CANTIDAD" in c), None)

  if req_col and cant_col:
    # REQ_COMPONENTE = (Cantidad unitaria en receta) * (Requerimiento total del pedido)
    # O si la receta ya trae la proporción unitaria, se escala con el requerimiento de data.
    # Ajustamos para que tome directamente el requerimiento de la orden cuando corresponda:
    df_merged["REQ_COMPONENTE"] = pd.to_numeric(
        df_merged[cant_col], errors="coerce"
    ).fillna(0) * pd.to_numeric(df_merged[req_col], errors="coerce").fillna(1)
  else:
    df_merged["REQ_COMPONENTE"] = pd.to_numeric(
        df_merged.get("CANTIDAD", 0), errors="coerce"
    ).fillna(0)

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
