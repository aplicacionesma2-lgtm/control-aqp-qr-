# 📦 Control de Empaque QR — María Almenara

App Streamlit para que el operario escanee (con la cámara del celular) el
código QR de cada caja empacada, la haga *match* con el pedido del día y
vea en tiempo real el % de avance de entrega por producto.

## Cómo funciona

1. **Sube el Excel del pedido** (barra lateral). Debe tener dos hojas:
   - `data`: el pedido del día — columnas `CÓDIGO`, `DESCRIPCIÓN`, `UMI`,
     `FACTOR`, `REQUERIMIENTO`, `FECHA PEDIDO`.
   - `Receta`: catálogo de productos — columnas `CÓDIGO`, `PRODUCTO`,
     `FACTOR`, `UMR` (las demás columnas de componentes se ignoran).
2. **Pestaña 🏷️ Etiquetas**: genera una hoja imprimible con un QR por
   producto (el valor del QR es el mismo código SAP, p. ej. `M3020044`).
   Imprime y pega la etiqueta en las cajas.
3. **Pestaña 📷 Escanear**: el operario escanea la caja con la cámara (o
   digita el código a mano si el QR no se puede leer), ingresa la
   cantidad de cajas y confirma. La app calcula
   `unidades = cajas × FACTOR` y las acumula contra el `REQUERIMIENTO`.
4. **📊 Avance del pedido**: detalle por producto con barra de progreso,
   estado (⛔ sin iniciar / 🟡 parcial / ✅ completo) y gráfico de dona.
5. **🕒 Historial**: lista de todos los escaneos, filtrable por operario,
   con opción de eliminar un registro si hubo un error.
6. **📤 Exportar**: descarga un Excel con el resumen de avance y el
   historial completo de escaneos.

`sample_data.xlsx` incluido es el archivo que compartiste, útil para
probar la app localmente antes de subir el pedido real de cada día.

## Persistencia: Google Sheets

El pedido, el catálogo y el historial de escaneos se guardan en **un
Google Sheet** con 3 pestañas (`Pedido`, `Catalogo`, `Registros`), que la
app crea sola la primera vez que corre. Ventajas frente a un archivo
local: sobrevive a redeploys de Streamlit Cloud y varios celulares
pueden escanear a la vez contra el mismo pedido.

### Configurar Google Sheets

1. **Crea (o reutiliza) una cuenta de servicio de Google Cloud.**
   Si quieres ahorrarte este paso, puedes reutilizar el mismo proyecto
   y cuenta de servicio que ya usas en tu app `entregas` — solo
   compártele acceso a la hoja nueva (paso 3). Si prefieres una nueva:
   - Ve a [Google Cloud Console](https://console.cloud.google.com/) →
     crea un proyecto.
   - Habilita **Google Sheets API** y **Google Drive API**.
   - Ve a "Cuentas de servicio" → crea una → genera una clave **JSON** y
     descárgala.
2. **Crea un Google Sheet vacío** para esta app (puede estar vacío, la
   app crea las pestañas solas). Copia su ID desde la URL:
   `https://docs.google.com/spreadsheets/d/ESTE_ES_EL_ID/edit`.
3. **Comparte ese Sheet** con el `client_email` que aparece en el JSON
   de la cuenta de servicio (dale permiso de **Editor**), igual que ya
   haces con tu app `entregas`.
4. **Configura los secrets.** Copia `.streamlit/secrets.toml.example` a
   `.streamlit/secrets.toml` (local) y completa los valores con los del
   JSON descargado. En Streamlit Cloud, pega el mismo contenido en
   *tu app → Settings → Secrets*.

```toml
SHEET_ID = "el-id-de-tu-google-sheet"

[gcp_service_account]
type = "service_account"
project_id = "..."
private_key_id = "..."
private_key = "-----BEGIN PRIVATE KEY-----\n...\n-----END PRIVATE KEY-----\n"
client_email = "...@....iam.gserviceaccount.com"
client_id = "..."
auth_uri = "https://accounts.google.com/o/oauth2/auth"
token_uri = "https://oauth2.googleapis.com/token"
auth_provider_x509_cert_url = "https://www.googleapis.com/oauth2/v1/certs"
client_x509_cert_url = "..."
```

Si faltan estos secrets, la app te lo dice apenas abre (no falla en
silencio).

## Instalación local

```bash
pip install -r requirements.txt
streamlit run app.py
```

Abre la URL que te muestre la consola (normalmente `localhost:8501`).
Para escanear con el celular necesitas HTTPS — en local puro
(`http://localhost`, en la misma compu) sí funciona, pero desde otro
dispositivo en la red local normalmente el navegador bloquea la cámara
salvo que uses un túnel HTTPS. Lo más simple es probar el escaneo ya
desplegado en Streamlit Cloud.

## Despliegue en Streamlit Cloud (igual que tus otras apps)

1. Sube esta carpeta a un repo de GitHub (`ruizleyva1982-lab/...`).
2. Conéctalo en Streamlit Cloud y configura los *Secrets* (sección
   anterior). Al ser HTTPS, la cámara funciona sin configuración extra.
3. Sube el Excel del pedido cada día desde el celular o la compu.

## Flujo recomendado día a día

1. Al iniciar el turno: sube el Excel del pedido del día.
2. Si es un pedido nuevo (no continuación del día anterior): usa
   "⚠️ Reiniciar historial del día" en la barra lateral para partir en
   cero.
3. El o los operarios escanean normalmente durante el turno.
4. Al cerrar: exporta el Excel de entrega desde la pestaña 📤 Exportar.

## Notas técnicas

- El escáner usa la librería `html5-qrcode` (vía `streamlit-qrcode-scanner`),
  que lee tanto QR como varios formatos de código de barras — no es
  necesario que el código sea estrictamente QR.
- Si ya generas etiquetas con BarTender/Zebra en tu otro flujo, puedes
  seguir usándolo: solo asegúrate de que el valor codificado sea el mismo
  `CÓDIGO` de SAP y esta app lo reconocerá igual.
- Las escrituras a Google Sheets usan `value_input_option="RAW"` (no
  `USER_ENTERED`) a propósito: es lo que evita que Sheets reinterprete
  números como fechas — el mismo problema que viste en `entregas`.
- Las lecturas usan una caché corta (`st.cache_data`, TTL de 3–5 s) para
  no saturar la cuota de la API de Sheets si varios celulares escanean
  seguido.
- Si la app llega a crashear en Streamlit Cloud al interactuar con
  widgets (como te pasó antes en `entregas`), prueba fijando versiones
  más antiguas y estables en `requirements.txt`
  (`streamlit==1.38.0`, `pandas==2.1.4`, `numpy==1.26.4`) — esta app no
  usa `pandas.Styler` (la causa sospechada la otra vez), pero por si
  acaso.
