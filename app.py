import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Simulador de Costos y Recetas - Maria Almenara", page_icon="📦", layout="wide")

st.title("📦 Sistema Integral de Producción y Recetas - Maria Almenara")
st.markdown("---")

# ==========================================
# 1. CARGA DE ARCHIVOS Y MENÚ PRINCIPAL
# ==========================================
st.sidebar.header("📁 Carga de Archivos")
archivo_excel = st.sidebar.file_uploader("Sube tu archivo Excel principal (costos, rep_cal, receta)", type=["xlsx", "xls"])

if archivo_excel is not None:
    try:
        archivo_excel.seek(0)
        xl_temp = pd.ExcelFile(archivo_excel)
        hojas = xl_temp.sheet_names
        st.sidebar.success(f"Archivo cargado. Hojas: {', '.join(hojas)}")
    except Exception as e:
        st.error(f"Error al leer el archivo Excel: {e}")
        hojas = []
        xl_temp = None

    if hojas:
        st.sidebar.markdown("---")
        st.sidebar.header("🧭 Menú de Módulos")
        menu_opcion = st.sidebar.radio(
            "Selecciona la sección:",
            [
                "🔍 Consulta de Recetas",
                "📊 Reportes (Merma / Prod)",
                "🏷️ Impresión de Etiquetas",
                "💰 Simulador de Costos"
            ]
        )

        # Función robusta de extracción de componentes
        def obtener_componentes_producto(codigo_prod: str, cajas: float) -> pd.DataFrame:
            """Extrae los componentes de la hoja 'Receta' de manera robusta y exacta."""
            try:
                archivo_excel.seek(0)
                xl = pd.ExcelFile(archivo_excel)
                
                nombre_hoja = next((s for s in xl.sheet_names if "receta" in s.strip().lower()), None)
                if not nombre_hoja:
                    return pd.DataFrame()
                
                df_receta = xl.parse(nombre_hoja)
            except Exception:
                return pd.DataFrame()

            if df_receta.empty:
                return pd.DataFrame()

            df_receta.columns = [str(c).strip().upper() for c in df_receta.columns]
            codigo_buscado = str(codigo_prod).strip().upper()

            col_prod = df_receta.columns[0]
            df_receta[col_prod] = df_receta[col_prod].fillna("").astype(str).str.strip().str.upper()

            matches = df_receta[df_receta[col_prod] == codigo_buscado]
            if matches.empty:
                matches = df_receta[df_receta[col_prod].str.contains(codigo_buscado, na=False)]

            if matches.empty:
                return pd.DataFrame()

            col_cod_comp = next((c for c in df_receta.columns if "COD" in c and "COMP" in c), df_receta.columns[4] if len(df_receta.columns) > 4 else df_receta.columns[1])
            col_desc_comp = next((c for c in df_receta.columns if c == "COMPONENTE" or ("COMP" in c and c != col_cod_comp)), df_receta.columns[5] if len(df_receta.columns) > 5 else col_cod_comp)
            col_cant = next((c for c in df_receta.columns if "CANTIDAD" in c or "CANT" in c), df_receta.columns[6] if len(df_receta.columns) > 6 else None)

            resultados = []
            for _, row in matches.iterrows():
                c_comp = str(row.get(col_cod_comp, "")).strip()
                d_comp = str(row.get(col_desc_comp, c_comp)).strip()

                if not c_comp or c_comp.upper() in ["NAN", "NONE", "", "NAT"]:
                    continue

                cant_base = 1.0
                if col_cant is not None:
                    try:
                        val_cant = row.get(col_cant, 1.0)
                        cant_base = float(val_cant) if pd.notna(val_cant) else 1.0
                    except (TypeError, ValueError):
                        cant_base = 1.0

                cant_total = cant_base * float(cajas)
                resultados.append({
                    "Código Componente": c_comp,
                    "Descripción Componente": d_comp if d_comp and d_comp.upper() != "NAN" else c_comp,
                    "Cantidad Requerida": cant_total
                })

            return pd.DataFrame(resultados).drop_duplicates().reset_index(drop=True)

        # ==========================================
        # MÓDULO 1: CONSULTA DE RECETAS
        # ==========================================
        if menu_opcion == "🔍 Consulta de Recetas":
            st.subheader("🔍 Consulta de Recetas por Producto")
            
            try:
                archivo_excel.seek(0)
                xl_master = pd.ExcelFile(archivo_excel)
                hoja_receta_nombre = next((s for s in xl_master.sheet_names if "receta" in s.strip().lower()), xl_master.sheet_names[0])
                df_master = xl_master.parse(hoja_receta_nombre)
                df_master.columns = [str(c).strip().upper() for c in df_master.columns]
                
                c_cod = df_master.columns[0]
                c_prod = df_master.columns[1] if len(df_master.columns) > 1 else c_cod
                
                df_master['OPCION'] = df_master[c_cod].astype(str) + " - " + df_master[c_prod].astype(str)
                lista_productos = sorted(df_master['OPCION'].dropna().unique())
            except Exception:
                lista_productos = []
                df_master = pd.DataFrame()

            col1, col2 = st.columns([2, 1])
            
            with col1:
                seleccion_combo = st.selectbox("Selecciona o busca un producto:", options=["-- Seleccione un producto --"] + lista_productos)
                codigo_manual = st.text_input("O ingresa el código manual (ej. M1020108):").strip()

            codigo_a_buscar = ""
            if codigo_manual:
                codigo_a_buscar = codigo_manual
            elif seleccion_combo != "-- Seleccione un producto --":
                codigo_a_buscar = seleccion_combo.split(" - ")[0]

            with col2:
                cajas_input = st.number_input("Cantidad de Cajas / Lotes:", min_value=0.01, value=1.0, step=1.0)

            if codigo_a_buscar:
                st.markdown("---")
                desc_prod = "Producto Seleccionado"
                factor_val = 3 
                try:
                    if not df_master.empty:
                        match_row = df_master[df_master[c_cod].astype(str).str.strip().str.upper() == codigo_a_buscar.upper()]
                        if not match_row.empty:
                            desc_prod = match_row.iloc[0][c_prod]
                            if "FACTOR" in df_master.columns:
                                factor_val = match_row.iloc[0]["FACTOR"]
                except Exception:
                    pass

                st.markdown(f"### 📋 **Producto:** {desc_prod} (`{codigo_a_buscar.upper()}`)")
                st.markdown(f"**Factor base:** {factor_val}")
                
                total_unidades = factor_val * cajas_input
                st.markdown(f"= **{total_unidades} UNIDAD (BIENES)**")
                
                st.markdown("---")
                st.subheader("📋 Componentes / Insumos de la Receta:")
                
                df_componentes = obtener_componentes_producto(codigo_a_buscar, cajas_input)
                
                if not df_componentes.empty:
                    st.dataframe(df_componentes, use_container_width=True)
                    
                    csv_data = df_componentes.to_csv(index=False).encode('utf-8')
                    st.download_button(
                        label="📥 Descargar Insumos en CSV",
                        data=csv_data,
                        file_name=f"insumos_{codigo_a_buscar}.csv",
                        mime="text/csv"
                    )
                else:
                    st.warning(f"No se encontraron componentes en la hoja 'Receta' para el código `{codigo_a_buscar}`.")

        # ==========================================
        # MÓDULO 2: REPORTES (MERMA / PROD)
        # ==========================================
        elif menu_opcion == "📊 Reportes (Merma / Prod)":
            st.subheader("📊 Módulo de Reportes y Control de Calidad")
            st.markdown("Visualización directa de las hojas operativas (Merma, Prod, etc.) de tu libro de Excel:")
            
            try:
                archivo_excel.seek(0)
                xl_rep = pd.ExcelFile(archivo_excel)
                hojas_disp = xl_rep.sheet_names
                
                hoja_sel = st.selectbox("Selecciona la hoja de reporte:", options=hojas_disp)
                if hoja_sel:
                    df_rep = xl_rep.parse(hoja_sel)
                    st.info(f"Mostrando hoja: **{hoja_sel}** ({len(df_rep)} registros totales)")
                    st.dataframe(df_rep, use_container_width=True)
                    
                    csv_rep = df_rep.to_csv(index=False).encode('utf-8')
                    st.download_button(
                        label=f"📥 Descargar {hoja_sel} en CSV",
                        data=csv_rep,
                        file_name=f"reporte_{hoja_sel}.csv",
                        mime="text/csv"
                    )
            except Exception as e:
                st.error(f"Error al cargar reportes: {e}")

        # ==========================================
        # MÓDULO 3: IMPRESIÓN DE ETIQUETAS
        # ==========================================
        elif menu_opcion == "🏷️ Impresión de Etiquetas":
            st.subheader("🏷️ Generación e Impresión de Etiquetas de Empaque")
            st.markdown("Selecciona el producto para visualizar y emitir las etiquetas de sus insumos:")
            
            try:
                archivo_excel.seek(0)
                xl_etq = pd.ExcelFile(archivo_excel)
                hoja_receta_nombre = next((s for s in xl_etq.sheet_names if "receta" in s.strip().lower()), xl_etq.sheet_names[0])
                df_etq_master = xl_etq.parse(hoja_receta_nombre)
                df_etq_master.columns = [str(c).strip().upper() for c in df_etq_master.columns]
                c_cod = df_etq_master.columns[0]
                c_prod = df_etq_master.columns[1] if len(df_etq_master.columns) > 1 else c_cod
                df_etq_master['OPCION'] = df_etq_master[c_cod].astype(str) + " - " + df_etq_master[c_prod].astype(str)
                lista_prod_etq = sorted(df_etq_master['OPCION'].dropna().unique())
            except Exception:
                lista_prod_etq = []

            prod_elegido = st.selectbox("Selecciona producto para etiquetas:", options=["-- Seleccione un producto --"] + lista_prod_etq)
            cajas_etq = st.number_input("Cantidad de Cajas / Lotes:", min_value=1.0, value=1.0, step=1.0)

            if prod_elegido != "-- Seleccione un producto --":
                cod_etq = prod_elegido.split(" - ")[0]
                desc_etq = prod_elegido.split(" - ")[1]
                
                df_comp_etq = obtener_componentes_producto(cod_etq, cajas_etq)
                if not df_comp_etq.empty:
                    st.markdown("---")
                    for _, row in df_comp_etq.iterrows():
                        st.markdown(
                            f"""
                            <div style="border: 2px dashed #4CAF50; padding: 14px; border-radius: 8px; margin-bottom: 12px; background-color: #fafafa;">
                                <h4 style="margin: 0; color: #2E7D32;">🏷️ ETIQUETA DE COMPONENTE / INSUMO</h4>
                                <hr style="margin: 6px 0;">
                                <p style="margin: 3px 0;"><b>Producto Padre:</b> {desc_etq} (<code>{cod_etq}</code>)</p>
                                <p style="margin: 3px 0;"><b>Insumo:</b> {row['Descripción Componente']} (<code>{row['Código Componente']}</code>)</p>
                                <p style="margin: 3px 0; font-size: 1.05em;"><b>Cantidad Requerida:</b> <span style="color: #c62828; font-weight: bold;">{row['Cantidad Requerida']}</span></p>
                            </div>
                            """,
                            unsafe_allow_html=True
                        )
                else:
                    st.warning("No se encontraron componentes para generar etiquetas en este producto.")

        # ==========================================
        # MÓDULO 4: SIMULADOR DE COSTOS
        # ==========================================
        elif menu_opcion == "💰 Simulador de Costos":
            st.subheader("💰 Simulador de Costos y Sustitución de Ingredientes")
            st.markdown("Analiza y evalúa el impacto de costos por cada orden de fabricación:")
            
            try:
                archivo_excel.seek(0)
                xl_cost = pd.ExcelFile(archivo_excel)
                hoja_cost = next((s for s in xl_cost.sheet_names if "costo" in s.strip().lower()), xl_cost.sheet_names[0])
                df_cost = xl_cost.parse(hoja_cost)
                st.info(f"Visualizando hoja de costos: **{hoja_cost}**")
                st.dataframe(df_cost.head(100), use_container_width=True)
            except Exception as e:
                st.warning(f"No se encontró una hoja dedicada a costos: {e}")

    else:
        st.info("El archivo cargado no contiene hojas válidas.")
else:
    st.info("👈 Sube tu archivo Excel usando el panel de la izquierda para desplegar todas las secciones del sistema.")
