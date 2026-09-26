import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Simulador de Costos y Recetas - Maria Almenara", page_icon="📦", layout="wide")

st.title("📦 Simulador de Costos y Recetas - Producción")
st.markdown("---")

# ==========================================
# 1. CARGA DE ARCHIVOS
# ==========================================
st.sidebar.header("📁 Carga de Archivos")
archivo_excel = st.sidebar.file_uploader("Sube tu archivo Excel principal (ej. costos / rep_cal / receta)", type=["xlsx", "xls"])

@st.cache_data
def cargar_datos(uploaded_file):
    if uploaded_file is None:
        return None, []
    try:
        xl = pd.ExcelFile(uploaded_file)
        return xl, xl.sheet_names
    except Exception as e:
        st.error(f"Error al leer el archivo Excel: {e}")
        return None, []

if archivo_excel is not None:
    xl_file, hojas = cargar_datos(archivo_excel)
    
    if xl_file:
        st.sidebar.success((
            f"Archivo cargado con éxito. Hojas disponibles: {', '.join(hojas)}"
        ))
        
        # ==========================================
        # 2. FUNCIÓN DE EXTRACCIÓN DE COMPONENTES
        # ==========================================
        def obtener_componentes_producto(codigo_prod: str, cajas: float) -> pd.DataFrame:
            """Extrae los componentes de la hoja 'Receta' de manera robusta y exacta."""
            try:
                archivo_excel.seek(0)
                xl = pd.ExcelFile(archivo_excel)
                
                # Buscar hoja 'Receta' sin distinción de mayúsculas/minúsculas o tildes
                nombre_hoja = next((s for s in xl.sheet_names if "receta" in s.strip().lower()), None)
                if not nombre_hoja:
                    return pd.DataFrame()
                
                df_receta = xl.parse(nombre_hoja)
            except Exception:
                return pd.DataFrame()

            if df_receta.empty:
                return pd.DataFrame()

            # Normalizar nombres de columnas a mayúsculas
            df_receta.columns = [str(c).strip().upper() for c in df_receta.columns]
            codigo_buscado = str(codigo_prod).strip().upper()

            # La primera columna es el código del producto principal
            col_prod = df_receta.columns[0]
            df_receta[col_prod] = df_receta[col_prod].fillna("").astype(str).str.strip().str.upper()

            # Filtrar coincidencias exactas con el código buscado
            matches = df_receta[df_receta[col_prod] == codigo_buscado]
            if matches.empty:
                matches = df_receta[df_receta[col_prod].str.contains(codigo_buscado, na=False)]

            if matches.empty:
                return pd.DataFrame()

            # Identificar columnas correspondientes de acuerdo a tu estructura de Excel
            # Columna E (aprox index 4): COD COMP, Columna F (aprox index 5): COMPONENTE, Columna G (aprox index 6): CANTIDAD
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
        # 3. INTERFAZ DE BÚSQUEDA Y VISUALIZACIÓN
        # ==========================================
        st.subheader("🔍 Consulta de Recetas por Producto")
        
        # Intentar cargar la lista de productos únicos desde la hoja 'Receta'
        try:
            archivo_excel.seek(0)
            xl_temp = pd.ExcelFile(archivo_excel)
            hoja_receta_nombre = next((s for s in xl_temp.sheet_names if "receta" in s.strip().lower()), xl_temp.sheet_names[0])
            df_master = xl_temp.parse(hoja_receta_nombre)
            df_master.columns = [str(c).strip().upper() for c in df_master.columns]
            
            c_cod = df_master.columns[0]
            c_prod = df_master.columns[1] if len(df_master.columns) > 1 else c_cod
            
            # Crear diccionario de opciones: "CODIGO - DESCRIPCION"
            df_master['OPCION'] = df_master[c_cod].astype(str) + " - " + df_master[c_prod].astype(str)
            lista_productos = sorted(df_master['OPCION'].dropna().unique())
        except Exception:
            lista_productos = []

        col1, col2 = st.columns([2, 1])
        
        with col1:
            seleccion_combo = st.selectbox("Selecciona o busca un producto:", options=["-- Seleccione un producto --"] + lista_productos)
            codigo_manual = st.text_input("O ingresa el código manual (ej. M1020108):").strip()

        # Determinar el código final a buscar
        codigo_a_buscar = ""
        if codigo_manual:
            codigo_a_buscar = codigo_manual
        elif seleccion_combo != "-- Seleccione un producto --":
            codigo_a_buscar = seleccion_combo.split(" - ")[0]

        # Configurar cantidad de cajas / factor
        with col2:
            cajas_input = st.number_input("Cantidad de Cajas / Lotes:", min_value=0.01, value=1.0, step=1.0)

        if codigo_a_buscar:
            st.markdown("---")
            # Extraer descripción del producto si existe
            desc_prod = "Producto Seleccionado"
            factor_val = 3 # Valor por defecto según tus capturas
            try:
                match_row = df_master[df_master[c_cod].astype(str).str.strip().str.upper() == codigo_a_buscar.upper()]
                if not match_row.empty:
                    desc_prod = match_row.iloc[0][c_prod]
                    if "FACTOR" in df_master.columns:
                        factor_val = match_row.iloc[0]["FACTOR"]
            except Exception:
                pass

            st.markdown(f"### 📋 **Producto:** {desc_prod} (`{codigo_a_buscar.upper()}`)")
            st.markdown(f"**Factor base:** {factor_val}")
            
            # Calcular total unidades
            total_unidades = factor_val * cajas_input
            st.markdown(f"= **{total_unidades} UNIDAD (BIENES)**")
            
            st.markdown("---")
            st.subheader("📋 Componentes / Insumos de la Receta:")
            
            df_componentes = obtener_componentes_producto(codigo_a_buscar, cajas_input)
            
            if not df_componentes.empty:
                st.dataframe(df_componentes, use_container_width=True)
                
                # Botón de descarga de resultados
                csv_data = df_componentes.to_csv(index=False).encode('utf-8')
                st.download_button(
                    label="📥 Descargar Insumos en CSV",
                    data=csv_data,
                    file_name=f"insumos_{codigo_a_buscar}.csv",
                    mime="text/csv"
                )
            else:
                st.warning(f"No se encontraron componentes en la hoja 'Receta' para el código `{codigo_a_buscar}`.")
    else:
        st.info("Por favor, sube tu archivo Excel en la barra lateral para comenzar.")
else:
    st.info("👈 Sube un archivo Excel usando el panel de la izquierda para desplegar la aplicación.")
