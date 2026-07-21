import streamlit as st
import pandas as pd
import os
import re
import openpyxl
from datetime import datetime, timedelta

# Configuración de la página con inyección de CSS para comprimir toda la interfaz
st.set_page_config(page_title="Sistema Imprenta", page_icon="🖨️", layout="wide")

st.markdown("""
    <style>
    /* Compactar padding general de Streamlit para ganar pantalla */
    .block-container { padding-top: 1rem !important; padding-bottom: 1rem !important; }
    h1 { margin-bottom: 0rem !important; padding-bottom: 0rem !important; font-size: 24px !important; }
    h3 { margin-top: 0.5rem !important; margin-bottom: 0.5rem !important; font-size: 18px !important; }
    .stTabs [data-baseweb="tab-list"] { gap: 4px !important; }
    .stTabs [data-baseweb="tab"] { padding-top: 4px !important; padding-bottom: 4px !important; font-size: 13px !important; }
    </style>
""", unsafe_allow_html=True)

st.title("🖨️ Sistema de Gestión - Imprenta")

# ==============================================================================
# 👇 CONFIGURACIÓN DE RUTA COMPARTIDA (YA CONFIGURADA CON TU DRIVE) 👇
# ==============================================================================
RUTA_COMPARTIDA = r"G:\Mi unidad\sistema_imprenta_datos" 

def obtener_ruta(nombre_archivo):
    if RUTA_COMPARTIDA:
        return os.path.join(RUTA_COMPARTIDA, nombre_archivo)
    return nombre_archivo
# ==============================================================================

# Creación de pestañas
tab_inicio, tab_pedidos, tab_materiales, tab_cheques, tab_semanal, tab_mensual = st.tabs([
    "📊 Dashboard", 
    "🛍️ Pedidos y Clientes", 
    "📦 Materiales y Gastos", 
    "💳 Cartera de Cheques",
    "📆 Cierre Semanal",
    "📅 Cierre Mensual"
])

# --- 1. ARCHIVO COMPARTIDO DE PEDIDOS DE CLIENTES ---
archivo_pedidos = obtener_ruta("PEDIDOS_CLIENTES.csv")
if os.path.exists(archivo_pedidos):
    # Aseguramos que todas estas columnas se lean estrictamente como texto (str)
    df_pedidos = pd.read_csv(archivo_pedidos, dtype={
        "Estado": str, 
        "Cobró": str, 
        "Observaciones": str, 
        "Impresión": str, 
        "Fecha Cobro": str,
        "Fecha": str
    })
    if "Cobró" not in df_pedidos.columns:
        df_pedidos["Cobró"] = "N/A"
    if "Observaciones" not in df_pedidos.columns:
        df_pedidos["Observaciones"] = ""
    if "Impresión" not in df_pedidos.columns:
        df_pedidos["Impresión"] = "🔴 Falta Imprimir"
        
    df_pedidos["Cobró"] = df_pedidos["Cobró"].fillna("N/A").astype(str).replace({"ROBERTO": "ROB", "GONZALO": "GON"})
    df_pedidos["Estado"] = df_pedidos["Estado"].fillna("Pendiente").astype(str)
    df_pedidos["Observaciones"] = df_pedidos["Observaciones"].fillna("").astype(str)
    df_pedidos["Impresión"] = df_pedidos["Impresión"].fillna("🔴 Falta Imprimir").astype(str)
else:
    df_pedidos = pd.DataFrame(columns=["Fecha", "Cliente", "Material", "Medida", "Copias", "Monto Total", "Estado", "Cobró", "Observaciones", "Impresión"])

# --- 2. PROCESAMIENTO DE MATERIALES (GASTOS PROVEEDORES) ---
archivo_materiales = obtener_ruta("MATERIALES.xlsx")
if os.path.exists(archivo_materiales):
    df_mat = pd.read_excel(archivo_materiales)
    df_mat.columns = df_mat.columns.str.strip().str.upper()
    
    columnas_requeridas = ["FECHA", "DESCRIPCION", "MONTO", "FORMA DE PAGO", "PROVEEDOR", "ESTADO", "SALDO"]
    for col in columnas_requeridas:
        if col not in df_mat.columns:
            df_mat[col] = ""
            
    df_mat["ESTADO"] = df_mat["ESTADO"].astype(str).str.upper().str.strip()
    df_mat["ESTADO"] = df_mat["ESTADO"].replace({"P": "PENDIENTE", "F": "PAGADO", "NAN": "PENDIENTE", "": "PENDIENTE"})
    df_mat["MONTO"] = pd.to_numeric(df_mat["MONTO"], errors="coerce").fillna(0.0)
    df_mat["SALDO"] = df_mat.apply(lambda r: 0.0 if r["ESTADO"] == "PAGADO" else r["MONTO"], axis=1)
    df_mat = df_mat.dropna(subset=["DESCRIPCION"], how="all")
else:
    df_mat = pd.DataFrame(columns=["FECHA", "DESCRIPCION", "MONTO", "FORMA DE PAGO", "PROVEEDOR", "ESTADO", "SALDO"])

# --- 3. PROCESAMIENTO DE CHEQUES ---
archivo_cheques = obtener_ruta("CHEQUES NUEVO.xlsx")
lista_cheques_procesados = []

if os.path.exists(archivo_cheques):
    try:
        wb = openpyxl.load_workbook(archivo_cheques, data_only=True)
        sheet = wb.active
        proveedores_cols = {}
        for col_idx in range(1, sheet.max_column + 1):
            prov_name = sheet.cell(row=1, column=col_idx).value
            if prov_name and not str(prov_name).startswith("Unnamed:"):
                proveedores_cols[col_idx] = str(prov_name).strip()
        
        for row_idx in range(2, sheet.max_row + 1):
            for col_idx, proveedor in proveedores_cols.items():
                cell = sheet.cell(row=row_idx, column=col_idx)
                val = cell.value
                if val is not None:
                    val_str = str(val).strip()
                    monto_match = re.search(r'\$?(\d+)', val_str)
                    fecha_match = re.search(r'(\d{2}/\d{2})', val_str)
                    if monto_match:
                        monto = float(monto_match.group(1))
                        fecha_str = fecha_match.group(1) if fecha_match else "Sin fecha"
                        fill_color = cell.fill.start_color.rgb if (cell.fill and cell.fill.start_color) else None
                        is_yellow = False
                        if fill_color:
                            is_yellow = "FFFF00" in str(fill_color) or str(fill_color) == "00FFFF00" or str(fill_color) == "FFFFFF00"
                        estado = "Pagado" if is_yellow else "Pendiente"
                        
                        lista_cheques_procesados.append({
                            "Proveedor": proveedor, "Monto": monto, "Fecha Vencimiento": fecha_str, "Estado": estado, "Detalle en Excel": val_str
                        })
        df_cheques = pd.DataFrame(lista_cheques_procesados)
    except Exception as e:
        df_cheques = pd.DataFrame()
else:
    df_cheques = pd.DataFrame()

# --- 4. ARCHIVO DE GASTOS DE RENDICIÓN SEMANAL ---
archivo_gastos_semanales = obtener_ruta("GASTOS_SEMANALES_CAJA.csv")
if os.path.exists(archivo_gastos_semanales):
    df_gastos_sem = pd.read_csv(archivo_gastos_semanales)
else:
    df_gastos_sem = pd.DataFrame(columns=["Semana", "Socio", "Concepto", "Monto"])


# --- PESTAÑA DASHBOARD ---
with tab_inicio:
    st.subheader("Resumen General")
    col1, col2, col3, col4 = st.columns(4)
    
    if not df_mat.empty:
        total_deuda_prov = df_mat["SALDO"].sum()
        col1.metric(label="Deuda a Proveedores 🟥", value=f"${total_deuda_prov:,.2f}")
    else:
        col1.metric(label="Deuda a Proveedores", value="$0.00")
        
    if not df_cheques.empty:
        df_pendientes = df_cheques[df_cheques["Estado"] == "Pendiente"]
        col2.metric(label="Cheques PENDIENTES ⚠️", value=f"${df_pendientes['Monto'].sum():,.2f}", delta=f"{len(df_pendientes)} activos")
    else:
        col2.metric(label="Cheques Pendientes", value="$0.00")

    total_ventas = df_pedidos["Monto Total"].sum() if not df_pedidos.empty else 0.0
    col3.metric(label="Trabajos Registrados", value=f"{len(df_pedidos)}")
    col4.metric(label="Facturación Clientes 💰", value=f"${total_ventas:,.2f}")


# --- PESTAÑA PEDIDOS Y CLIENTES (EDITABLE INTERACTIVA) ---
with tab_pedidos:
    st.subheader("Gestión de Clientes y Trabajos de Impresión / Cartelería")
    
    with st.expander("➕ Cargar Nuevo Trabajo (Impresión o Cartelería)", expanded=False):
        with st.form("form_imprenta_pedido"):
            col_a, col_b = st.columns(2)
            with col_a:
                pedido_fecha = st.date_input("Fecha del Trabajo:", value=datetime.today())
                cliente_name = st.text_input("Nombre del Cliente:")
                material_tipo = st.selectbox("Tipo de Material / Rubro:", [
                    "LONA FRONTLIGHT", "LONA BACKLIGHT", "VINILO BASE BLANCA", 
                    "VINILO BASE GRIS", "VINILO MATE", "VINILO CRISTAL", 
                    "VINILO MICROPERFORADO", "TELA BANDERA", "PAPEL BLUEBACK",
                    "TRABAJOS DE CARTELERIA"
                ])
                medida_trabajo = st.text_input("Medida / Tamaño:")
            with col_b:
                monto_trabajo = st.number_input("Monto total ($):", min_value=0.0, step=100.0)
                estado_pedido = st.selectbox("Estado Pago / Entrega:", ["Pendiente", "Entregado Cobrado", "Entregado Sin Cobrar"])
                usuario_cobro = st.selectbox("¿Quién cobró?:", ["N/A", "ROBERTO", "GONZALO"])
                estado_impresion = st.selectbox("Estado Producción:", ["🔴 Falta Imprimir", "🟢 Finalizado"])
            
            obs_trabajo = st.text_area("🗒️ Observaciones:")
            bot_guardar_p = st.form_submit_button("💾 Registrar Trabajo")
            
            if bot_guardar_p:
                if cliente_name:
                    fecha_formateada = pedido_fecha.strftime("%d/%m/%Y")
                    quien_guarda = "ROB" if "ROBERTO" in usuario_cobro.upper() else ("GON" if "GONZALO" in usuario_cobro.upper() else "N/A")
                    
                    nuevo_registro = pd.DataFrame([{
                        "Fecha": fecha_formateada, "Cliente": cliente_name.upper().strip(), "Material": material_tipo,
                        "Medida": medida_trabajo.upper().strip() if medida_trabajo else "N/A", "Copias": "1 UNIDAD",
                        "Monto Total": monto_trabajo, "Estado": estado_pedido, "Cobró": quien_guarda,
                        "Observaciones": obs_trabajo.strip() if obs_trabajo else "", "Impresión": estado_impresion
                    }])
                    df_pedidos = pd.concat([df_pedidos, nuevo_registro], ignore_index=True)
                    df_pedidos.to_csv(archivo_pedidos, index=False)
                    st.success("¡Trabajo registrado!")
                    st.rerun()

    st.write("---")
    st.subheader("📋 Historial General de Trabajos (Doble clic para editar celdas)")
    
    df_visualizacion = df_pedidos.copy()
    if not df_visualizacion.empty:
        buscar_cli = st.text_input("🔍 Filtrar historial por Cliente:")
        if buscar_cli:
            df_visualizacion = df_visualizacion[df_visualizacion["Cliente"].astype(str).str.contains(buscar_cli, case=False, na=False)]
        
        df_agrupado = df_visualizacion.groupby(["Fecha", "Cliente"], as_index=False).agg({
            "Material": lambda x: " / ".join(x.astype(str)), 
            "Medida": lambda x: " / ".join(x.astype(str)),
            "Monto Total": "sum", 
            "Estado": "first", 
            "Cobró": "first",
            "Observaciones": lambda x: " | ".join([v for v in x.astype(str) if v.strip() != ""]),
            "Impresión": "first"
        })
        
        df_agrupado["Cobró"] = df_agrupado["Cobró"].replace({"ROBERTO": "ROB", "GONZALO": "GON"})
        
        df_agrupado["Fecha_temp"] = pd.to_datetime(df_agrupado["Fecha"], format="%d/%m/%Y", errors="coerce")
        df_agrupado = df_agrupado.sort_values(by="Fecha_temp", ascending=False).drop(columns=["Fecha_temp"])
        df_agrupado = df_agrupado[["Fecha", "Cliente", "Material", "Medida", "Monto Total", "Estado", "Cobró", "Impresión", "Observaciones"]]
        
        # --- GENERAR RENGLÓN EN BLANCO / SEPARADOR AL CAMBIAR DE DÍA ---
        lista_con_separadores = []
        for i in range(len(df_agrupado)):
            fila_actual = df_agrupado.iloc[i].copy()
            
            if fila_actual["Estado"] == "Entregado Cobrado":
                fila_actual["Fecha_Visual"] = f"🟢 {fila_actual['Fecha']}"
            else:
                fila_actual["Fecha_Visual"] = fila_actual["Fecha"]
                
            lista_con_separadores.append(fila_actual)
            
            if i < len(df_agrupado) - 1:
                fecha_siguiente = df_agrupado.iloc[i+1]["Fecha"]
                if fila_actual["Fecha"] != fecha_siguiente:
                    fila_separadora = pd.Series({
                        "Fecha": "", "Cliente": "--- CAMBIO DE DÍA ---", "Material": "---", "Medida": "---", 
                        "Monto Total": 0.0, "Estado": "---", "Cobró": "---", "Impresión": "---", 
                        "Observaciones": "---", "Fecha_Visual": "────────────────"
                    })
                    lista_con_separadores.append(fila_separadora)
                    
        df_final_visual = pd.DataFrame(lista_con_separadores).reset_index(drop=True)
        
        columnas_ordenadas = ["Fecha_Visual", "Cliente", "Material", "Medida", "Monto Total", "Estado", "Cobró", "Impresión", "Observaciones"]
        df_editor_vista = df_final_visual[columnas_ordenadas]
        
        config_columnas = {
            "Fecha_Visual": st.column_config.TextColumn("Fecha", disabled=True),
            "Cliente": st.column_config.TextColumn(disabled=True),
            "Material": st.column_config.TextColumn(disabled=True),
            "Medida": st.column_config.TextColumn(disabled=True),
            "Monto Total": st.column_config.NumberColumn(format="$%.2f", disabled=True),
            "Estado": st.column_config.SelectboxColumn(options=["Pendiente", "Entregado Cobrado", "Entregado Sin Cobrar"]),
            "Cobró": st.column_config.SelectboxColumn(options=["N/A", "ROB", "GON"]),
            "Impresión": st.column_config.SelectboxColumn(options=["🔴 Falta Imprimir", "🟢 Finalizado"]),
            "Observaciones": st.column_config.TextColumn()
        }
        
        df_cambiado = st.data_editor(
            df_editor_vista,
            column_config=config_columnas,
            use_container_width=True,
            hide_index=True,
            key="editor_historial_pedidos"
        )
        
        # --- Lógica de guardado y fecha automática ---
        if not df_cambiado.equals(df_editor_vista):
            for idx, fila in df_cambiado.iterrows():
                if fila["Cliente"] == "--- CAMBIO DE DÍA ---": continue
                
                fecha_real = df_final_visual.iloc[idx]["Fecha"]
                f_mask = (df_pedidos["Fecha"] == fecha_real) & (df_pedidos["Cliente"] == fila["Cliente"])
                
                if f_mask.any():
                    # Si el estado cambia a Entregado Cobrado y antes NO lo estaba
                    if fila["Estado"] == "Entregado Cobrado" and df_pedidos.loc[f_mask, "Estado"].values[0] != "Entregado Cobrado":
                        df_pedidos.loc[f_mask, "Fecha Cobro"] = datetime.now().strftime("%d/%m/%Y")
                    
                    df_pedidos.loc[f_mask, "Estado"] = fila["Estado"]
                    df_pedidos.loc[f_mask, "Cobró"] = fila["Cobró"]
                    df_pedidos.loc[f_mask, "Impresión"] = fila["Impresión"]
                    df_pedidos.loc[f_mask, "Observaciones"] = fila["Observaciones"]
            
            df_pedidos.to_csv(archivo_pedidos, index=False)
            st.rerun()
            
        fecha_hoy_str = datetime.now().strftime("%d/%m/%Y")
        st.write("---")
        st.markdown(f"### 💵 Rendición de Caja ({fecha_hoy_str})")
        
        # Suma de TODO lo pendiente en el historial general (sin filtrar por fecha de hoy)
        monto_pendiente = df_pedidos[df_pedidos["Estado"].isin(["Pendiente", "Entregado Sin Cobrar"])]["Monto Total"].sum()
        
        # Cobrados del día actual
        df_hoy = df_visualizacion[df_visualizacion["Fecha"] == fecha_hoy_str]
        monto_roberto = df_hoy[(df_hoy["Estado"] == "Entregado Cobrado") & (df_hoy["Cobró"].isin(["ROB", "ROBERTO"]))]["Monto Total"].sum()
        monto_gonzalo = df_hoy[(df_hoy["Estado"] == "Entregado Cobrado") & (df_hoy["Cobró"].isin(["GON", "GONZALO"]))]["Monto Total"].sum()
        
        col_m1, col_m2, col_m3 = st.columns(3)
        col_m1.metric(label="🔴 MONTO PENDIENTE TOTAL (GENERAL)", value=f"${monto_pendiente:,.2f}")
        col_m2.metric(label="🟢 COBRADO POR ROB (HOY)", value=f"${monto_roberto:,.2f}")
        col_m3.metric(label="🟢 COBRADO POR GON (HOY)", value=f"${monto_gonzalo:,.2f}")


# --- PESTAÑA MATERIALES Y GASTOS PROVEEDORES ---
with tab_materiales:
    st.subheader("Gastos en Materiales e Insumos")
    
    with st.expander("➕ Cargar Nueva Compra / Gasto de Proveedor", expanded=False):
        with st.form("form_nuevo_gasto"):
            col_g1, col_g2 = st.columns(2)
            with col_g1:
                gasto_fecha = st.date_input("Fecha de Compra:", value=datetime.today())
                gasto_desc = st.text_input("Descripción del material:")
                gasto_monto = st.number_input("Monto total ($):", min_value=0.0, step=100.0)
            with col_g2:
                gasto_prov = st.text_input("Proveedor:").upper().strip()
                gasto_pago = st.text_input("Forma de Pago (Opcional):")
                gasto_estado = st.selectbox("Estado de la Factura:", ["PENDIENTE", "PAGADO"])
                
            bot_guardar_gasto = st.form_submit_button("📦 Registrar Compra")
            
            if bot_guardar_gasto:
                if gasto_desc and gasto_monto > 0:
                    gasto_saldo = 0.0 if gasto_estado == "PAGADO" else gasto_monto
                    nuevo_gasto = pd.DataFrame([{
                        "FECHA": gasto_fecha.strftime("%Y-%m-%d"), "DESCRIPCION": gasto_desc.upper().strip(),
                        "MONTO": gasto_monto, "FORMA DE PAGO": gasto_pago.upper().strip() if gasto_pago else "",
                        "PROVEEDOR": gasto_prov if gasto_prov else "VARIOS", "ESTADO": gasto_estado, "SALDO": gasto_saldo
                    }])
                    df_mat = pd.concat([df_mat, nuevo_gasto], ignore_index=True)
                    df_mat.to_excel(archivo_materiales, index=False)
                    st.success("¡Compra guardada!")
                    st.rerun()

    st.write("---")
    st.subheader("📋 Historial de Compras (Doble clic para editar celdas)")
    
    if not df_mat.empty:
        buscador_m = st.text_input("🔍 Buscar por descripción o proveedor (Materiales):")
        df_mat_filt = df_mat.copy()
        if buscador_m:
            df_mat_filt = df_mat_filt[df_mat_filt["DESCRIPCION"].astype(str).str.contains(buscador_m, case=False, na=False) | df_mat_filt["PROVEEDOR"].astype(str).str.contains(buscador_m, case=False, na=False)]
            
        df_mat_filt["FECHA"] = df_mat_filt["FECHA"].astype(str).apply(lambda x: x.split(" ")[0])
        df_mat_filt = df_mat_filt.sort_values(by="FECHA", ascending=False)
        
        df_mat_filt["FECHA_ORIGINAL"] = df_mat_filt["FECHA"]
        df_mat_filt["FECHA"] = df_mat_filt.apply(
            lambda r: f"🟢 {r['FECHA']}" if r["ESTADO"] == "PAGADO" else r["FECHA"], 
            axis=1
        )
        
        config_mat = {
            "FECHA": st.column_config.TextColumn(disabled=True),
            "DESCRIPCION": st.column_config.TextColumn(),
            "MONTO": st.column_config.NumberColumn(format="$%.2f"),
            "FORMA DE PAGO": st.column_config.TextColumn(),
            "PROVEEDOR": st.column_config.TextColumn(),
            "ESTADO": st.column_config.SelectboxColumn(options=["PENDIENTE", "PAGADO"]),
            "SALDO": st.column_config.NumberColumn(format="$%.2f", disabled=True)
        }
        
        df_mat_cambiado = st.data_editor(
            df_mat_filt.drop(columns=["FECHA_ORIGINAL"]),
            column_config=config_mat,
            use_container_width=True,
            hide_index=True,
            key="editor_materiales"
        )
        
        if not df_mat_cambiado.equals(df_mat_filt.drop(columns=["FECHA_ORIGINAL"])):
            for idx, fila in df_mat_cambiado.iterrows():
                fecha_real = df_mat_filt.iloc[idx]["FECHA_ORIGINAL"]
                desc_real = df_mat_filt.iloc[idx]["DESCRIPCION"]
                
                f_mask = (df_mat["FECHA"].astype(str).str.contains(fecha_real)) & (df_mat["DESCRIPCION"] == desc_real)
                if f_mask.any():
                    df_mat.loc[f_mask, "DESCRIPCION"] = fila["DESCRIPCION"]
                    df_mat.loc[f_mask, "MONTO"] = fila["MONTO"]
                    df_mat.loc[f_mask, "FORMA DE PAGO"] = fila["FORMA DE PAGO"]
                    df_mat.loc[f_mask, "PROVEEDOR"] = fila["PROVEEDOR"]
                    df_mat.loc[f_mask, "ESTADO"] = fila["ESTADO"]
                    df_mat.loc[f_mask, "SALDO"] = 0.0 if fila["ESTADO"] == "PAGADO" else fila["MONTO"]
            
            df_mat.to_excel(archivo_materiales, index=False)
            st.rerun()
        
        total_deuda_filtrada = df_mat_filt[df_mat_filt["ESTADO"] != "PAGADO"]["MONTO"].sum()
        st.error(f"⚠️ Total adeudado en esta selección de materiales: ${total_deuda_filtrada:,.2f}")


# --- PESTAÑA CHEQUES ---
with tab_cheques:
    st.subheader("💳 Cartera de Cheques")
    
    # --- FORMULARIO DE CARGA ---
    with st.expander("➕ Cargar Nuevo Cheque", expanded=True):
        with st.form("form_nuevo_cheque_v3", clear_on_submit=True):
            col_c1, col_c2 = st.columns(2)
            with col_c1:
                nuevo_prov = st.text_input("Proveedor:")
                nuevo_monto = st.number_input("Monto ($):", min_value=0.0, step=100.0)
            with col_c2:
                nueva_fecha = st.text_input("Fecha Vencimiento (ej: 20/12):")
                nuevo_estado = st.selectbox("Estado:", ["Pendiente", "Pagado"])
            
            if st.form_submit_button("💾 Registrar Cheque"):
                if nuevo_prov and nuevo_monto > 0:
                    wb = openpyxl.load_workbook(archivo_cheques)
                    ws = wb.active
                    col_idx = None
                    for c in range(1, ws.max_column + 1):
                        if str(ws.cell(1, c).value).strip() == nuevo_prov:
                            col_idx = c
                            break
                    if not col_idx:
                        col_idx = ws.max_column + 1
                        ws.cell(1, col_idx).value = nuevo_prov
                    
                    next_row = ws.max_row + 1
                    ws.cell(row=next_row, column=col_idx).value = f"${nuevo_monto} {nueva_fecha}"
                    wb.save(archivo_cheques)
                    st.success("¡Registrado!")
                    st.rerun()
                else:
                    st.error("Completá proveedor y monto.")

    st.write("---")
    
    # --- TABLA DE CHEQUES (SIN DETALLE EN EXCEL) ---
    if not df_cheques.empty:
        col_filtro1, col_filtro2 = st.columns(2)
        with col_filtro1:
            proveedor_cheque = st.selectbox("Filtrar por Proveedor:", ["Todos"] + sorted(df_cheques["Proveedor"].unique().tolist()))
        with col_filtro2:
            estado_cheque = st.radio("Filtrar por Estado:", ["Todos", "Pendientes", "Pagados"], horizontal=True)
        
        df_cheques_filt = df_cheques.copy()
        if proveedor_cheque != "Todos":
            df_cheques_filt = df_cheques_filt[df_cheques_filt["Proveedor"] == proveedor_cheque]
        if estado_cheque == "Pendientes":
            df_cheques_filt = df_cheques_filt[df_cheques_filt["Estado"] == "Pendiente"]
        elif estado_cheque == "Pagados":
            df_cheques_filt = df_cheques_filt[df_cheques_filt["Estado"] == "Pagado"]
            
        df_cheques_filt["Fecha Vencimiento"] = df_cheques_filt.apply(
            lambda r: f"🟢 {r['Fecha Vencimiento']}" if r["Estado"] == "Pagado" else r["Fecha Vencimiento"], 
            axis=1
        )
        
        # Ocultar la columna 'Detalle en Excel'
        st.data_editor(
            df_cheques_filt[["Proveedor", "Monto", "Fecha Vencimiento", "Estado"]],
            column_config={
                "Monto": st.column_config.NumberColumn(format="$%.2f"),
                "Estado": st.column_config.SelectboxColumn(options=["Pendiente", "Pagado"])
            },
            use_container_width=True, hide_index=True
        )


# --- PESTAÑA: CIERRE SEMANAL ---
with tab_semanal:
    st.subheader("Control, Rendición y Gastos de Caja Semanal")
    if not df_pedidos.empty:
        df_sem = df_pedidos.copy()
        
        # Ajuste para leer la columna que acabás de crear
        # Usamos .get() o verificamos existencia para evitar errores
        if "Fecha Cobro" in df_sem.columns:
            df_sem["Fecha_Uso"] = df_sem["Fecha Cobro"].fillna(df_sem["Fecha"])
        else:
            df_sem["Fecha_Uso"] = df_sem["Fecha"]
            
        df_sem["Fecha_dt"] = pd.to_datetime(df_sem["Fecha_Uso"], format="%d/%m/%Y", errors='coerce')
        df_sem = df_sem.dropna(subset=["Fecha_dt"])
        df_sem["Semana_Inicio"] = df_sem["Fecha_dt"].dt.to_period('W').dt.start_time
        
        semanas_disponibles = sorted(df_sem["Semana_Inicio"].unique(), reverse=True)
        
        if semanas_disponibles:
            semana_sel = st.selectbox("Seleccioná la semana a auditar:", semanas_disponibles, format_func=lambda x: f"Semana del Lunes {x.strftime('%d/%m/%Y')} al Domingo {(x + timedelta(days=6)).strftime('%d/%m/%Y')}")
            semana_sel_str = semana_sel.strftime("%Y-%m-%d")
            
            df_semana_filtrada = df_sem[df_sem["Semana_Inicio"] == semana_sel].copy()
            df_cobrado_sem = df_semana_filtrada[df_semana_filtrada["Estado"] == "Entregado Cobrado"]
            
            total_sem_bruto_rob = df_cobrado_sem[df_cobrado_sem["Cobró"].isin(["ROB", "ROBERTO"])]["Monto Total"].sum()
            total_sem_bruto_gon = df_cobrado_sem[df_cobrado_sem["Cobró"].isin(["GON", "GONZALO"])]["Monto Total"].sum()
            
            st.write("---")
            
            # --- DISEÑO EN DOS COLUMNAS INDEPENDIENTES ---
            col_roberto, col_gonzalo = st.columns(2)
            
            # COLUMNA ROBERTO
            with col_roberto:
                st.markdown("### 💼 RENDICIÓN ROBERTO")
                
                df_gastos_rob = df_gastos_sem[(df_gastos_sem["Semana"] == semana_sel_str) & (df_gastos_sem["Socio"] == "ROB")][["Concepto", "Monto"]].reset_index(drop=True)
                
                st.markdown(f"**Cobrado Bruto Clientes:** `${total_sem_bruto_rob:,.2f}`")
                st.caption("👇 Agregar o editar gastos de Roberto acá abajo:")
                
                gastos_rob_editado = st.data_editor(
                    df_gastos_rob,
                    column_config={
                        "Concepto": st.column_config.TextColumn(),
                        "Monto": st.column_config.NumberColumn(format="$%.2f", min_value=0.0)
                    },
                    num_rows="dynamic",
                    use_container_width=True,
                    key="editor_gastos_roberto"
                )
                
                total_gastos_rob = gastos_rob_editado["Monto"].sum()
                neto_roberto = total_sem_bruto_rob - total_gastos_rob
                
                st.metric(label="💰 TOTAL NETO LIMPIO (ROB)", value=f"${neto_roberto:,.2f}", delta=f"-${total_gastos_rob:,.2f} Gastos")
                
            # COLUMNA GONZALO
            with col_gonzalo:
                st.markdown("### 💼 RENDICIÓN GONZALO")
                
                df_gastos_gon = df_gastos_sem[(df_gastos_sem["Semana"] == semana_sel_str) & (df_gastos_sem["Socio"] == "GON")][["Concepto", "Monto"]].reset_index(drop=True)
                
                st.markdown(f"**Cobrado Bruto Clientes:** `${total_sem_bruto_gon:,.2f}`")
                st.caption("👇 Agregar o editar gastos de Gonzalo acá abajo:")
                
                gastos_gon_editado = st.data_editor(
                    df_gastos_gon,
                    column_config={
                        "Concepto": st.column_config.TextColumn(),
                        "Monto": st.column_config.NumberColumn(format="$%.2f", min_value=0.0)
                    },
                    num_rows="dynamic",
                    use_container_width=True,
                    key="editor_gastos_gonzalo"
                )
                
                total_gastos_gon = gastos_gon_editado["Monto"].sum()
                neto_gonzalo = total_sem_bruto_gon - total_gastos_gon
                
                st.metric(label="💰 TOTAL NETO LIMPIO (GON)", value=f"${neto_gonzalo:,.2f}", delta=f"-${total_gastos_gon:,.2f} Gastos")

            # --- BALANCE Y DIVISIÓN DE CAJA ---
            st.write("---")
            st.markdown("### ⚖️ BALANCE Y DIVISIÓN DE CAJA (50% / 50%)")
            
            mitad_roberto = neto_roberto / 2
            mitad_gonzalo = neto_gonzalo / 2
            diferencia_netos = mitad_roberto - mitad_gonzalo
            
            col_div1, col_div2, col_div3 = st.columns(3)
            col_div1.metric(label="👤 Mitad Roberto (Neto / 2)", value=f"${mitad_roberto:,.2f}")
            col_div2.metric(label="👤 Mitad Gonzalo (Neto / 2)", value=f"${mitad_gonzalo:,.2f}")
            
            if diferencia_netos > 0:
                col_div3.metric(label="🔄 Ajuste de Cuentas", value=f"${abs(diferencia_netos):,.2f}", delta="A favor de GON", delta_color="inverse")
                st.info(f"💡 **Explicación del Balance:** Para que queden iguales, **Roberto debe darle ${abs(diferencia_netos):,.2f} a Gonzalo**.")
            elif diferencia_netos < 0:
                col_div3.metric(label="🔄 Ajuste de Cuentas", value=f"${abs(diferencia_netos):,.2f}", delta="A favor de ROB")
                st.info(f"💡 **Explicación del Balance:** Para que queden iguales, **Gonzalo debe darle ${abs(diferencia_netos):,.2f} a Roberto**.")
            else:
                col_div3.metric(label="🔄 Ajuste de Cuentas", value="$0.00", delta="Caja Perfecta")
                st.success("¡Están perfectamente empatados! No se deben nada.")

            # --- BOTÓN DE GUARDADO DE GASTOS ---
            st.write("")
            if st.button("💾 Guardar Cambios de Gastos Semanales"):
                df_gastos_sem = df_gastos_sem[df_gastos_sem["Semana"] != semana_sel_str]
                
                if not gastos_rob_editado.empty:
                    gastos_rob_save = gastos_rob_editado.dropna(subset=["Concepto"])
                    gastos_rob_save.insert(0, "Semana", semana_sel_str)
                    gastos_rob_save.insert(1, "Socio", "ROB")
                    df_gastos_sem = pd.concat([df_gastos_sem, gastos_rob_save], ignore_index=True)
                    
                if not gastos_gon_editado.empty:
                    gastos_gon_save = gastos_gon_editado.dropna(subset=["Concepto"])
                    gastos_gon_save.insert(0, "Semana", semana_sel_str)
                    gastos_gon_save.insert(1, "Socio", "GON")
                    df_gastos_sem = pd.concat([df_gastos_sem, gastos_gon_save], ignore_index=True)
                
                df_gastos_sem.to_csv(archivo_gastos_semanales, index=False)
                st.success("¡Gastos de la semana guardados correctamente!")
                st.rerun()

            st.write("---")
            total_sem_pendiente = df_semana_filtrada[df_semana_filtrada["Estado"].isin(["Pendiente", "Entregado Sin Cobrar"])]["Monto Total"].sum()
            st.error(f"📉 TOTAL TRABAJOS PENDIENTES DE COBRO EN ESTA SEMANA: ${total_sem_pendiente:,.2f}")
            
            st.markdown("#### 📋 Detalle de operaciones de la semana seleccionada")
            df_semana_filtrada = df_semana_filtrada.sort_values(by="Fecha_dt", ascending=False)
            df_semana_filtrada = df_semana_filtrada[["Fecha", "Cliente", "Material", "Medida", "Monto Total", "Estado", "Cobró", "Impresión", "Observaciones"]]
            st.dataframe(df_semana_filtrada, use_container_width=True, hide_index=True)
            # --- NUEVA PESTAÑA: CIERRE MENSUAL ---
with tab_mensual:
    st.subheader("📅 Cierre Mensual: Resumen de Ingresos y Gastos")
    if not df_pedidos.empty:
        df_pedidos["Fecha_dt"] = pd.to_datetime(df_pedidos["Fecha"], format="%d/%m/%Y", errors="coerce")
        meses = sorted(df_pedidos["Fecha_dt"].dt.to_period("M").unique(), reverse=True)
        mes_sel = st.selectbox("Seleccionar Mes a Auditar:", meses, format_func=lambda x: x.strftime("%B %Y"))
        
        # Ingresos: Pedidos entregados y cobrados en ese mes
        pedidos_mes = df_pedidos[df_pedidos["Fecha_dt"].dt.to_period("M") == mes_sel]
        ingresos = pedidos_mes[pedidos_mes["Estado"] == "Entregado Cobrado"]["Monto Total"].sum()
        
        # Gastos: Suma de los gastos semanales registrados en ese mismo mes
        gastos_mes = 0
        if not df_gastos_sem.empty:
            df_gastos_sem["Fecha_dt"] = pd.to_datetime(df_gastos_sem["Semana"], errors="coerce")
            gastos_mes = df_gastos_sem[df_gastos_sem["Fecha_dt"].dt.to_period("M") == mes_sel]["Monto"].sum()
            
        col1, col2, col3 = st.columns(3)
        col1.metric("Ingresos Cobrados", f"${ingresos:,.2f}")
        col2.metric("Gastos Totales", f"${gastos_mes:,.2f}")
        col3.metric("Resultado Neto Mensual", f"${(ingresos - gastos_mes):,.2f}")
        
        st.write("---")
        st.dataframe(pedidos_mes[["Fecha", "Cliente", "Monto Total", "Estado", "Cobró"]], use_container_width=True)