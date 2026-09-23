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

  # Normalizar columnas clave
  # data debe tener CÓDIGO y REQUERIMIENTO
  # Receta debe tener CÓDIGO, CANTIDAD (o proporción) y COD COMPONENTE
  
  df_merged = pd.merge(
      df_data,
      df_receta,
      on="CÓDIGO",
      how="inner",
      suffixes=("_PEDIDO", "_RECETA"),
  )

  # Identificar columnas de cantidad y requerimiento de manera segura
  col_req = next((c for c in df_merged.columns if "REQUERIMIENTO" in c), None)
  col_cant = next((c for c in df_merged.columns if "CANTIDAD" in c), None)

  val_req = pd.to_numeric(df_merged[col_req], errors="coerce").fillna(0) if col_req else 2160
  val_cant = pd.to_numeric(df_merged[col_cant], errors="coerce").fillna(0) if col_cant else 1

  # Cálculo neto: si la receta ya viene por unidad, se multiplica por el requerimiento total del pedido
  df_merged["REQ_COMPONENTE"] = val_cant * val_req

  col_comp_receta = next(
      (c for c in df_merged.columns if "COMPONENTE" in c), "COD COMPONENTE"
  )
  df_m = df_merged[
      df_merged[col_comp_receta].astype(str).str.upper().str.startswith("M")
  ].copy()

  # Normalizar factor-lima
  col_cod_f = next((c for c in df_factor.columns if "COD" in c), df_factor.columns[0])
  col_fac_f = next((c for c in df_factor.columns if "FACTOR" in c), df_factor.columns[-1])
  df_factor = df_factor.rename(columns={col_cod_f: "CÓDIGO_LIMA", col_fac_f: "FACTOR_VALOR"})

  df_final = pd.merge(
      df_m,
      df_factor,
      left_on=col_comp_receta,
      right_on="CÓDIGO_LIMA",
      how="left",
  )

  df_final["FACTOR_LIMA"] = pd.to_numeric(df_final["FACTOR_VALOR"], errors="coerce").fillna(1.0)
  df_final["FACTOR_LIMA"] = df_final["FACTOR_LIMA"].apply(lambda x: x if x > 0 else 1.0)

  col_desc_comp = next(
      (c for c in df_final.columns if "DESCRIPCIÓN" in c or "PRODUCTO" in c),
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
      lambda row: math.ceil(row["REQ_TOTAL"] / row["FACTOR_LIMA"]) * row["FACTOR_LIMA"],
      axis=1
  )

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
