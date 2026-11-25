import io
import os
import re
from datetime import datetime, timedelta
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="Imprenta Celular",
    page_icon="📱",
    layout="wide",
)

st.markdown(
    """
    <style>
        .stApp { background-color: #0b0f19; color: #e2e8f0; }
        [data-testid="stHeader"], [data-testid="stToolbar"], .stDeployButton, #MainMenu {
            display: none !important;
        }
        .block-container { padding-top: 1rem !important; padding-bottom: 2rem !important; }
        .kpi-card {
            background-color: #111827;
            border: 1px solid #243044;
            border-radius: 10px;
            padding: 12px 14px;
            margin-bottom: 8px;
        }
        .kpi-label { font-size: 12px; color: #94a3b8; }
        .kpi-value { font-size: 20px; font-weight: 800; color: #f8fafc; }
    </style>
    """,
    unsafe_allow_html=True,
)


def _extraer_id_drive(valor: str) -> str:
    """Acepta un ID de Google Drive o un enlace completo de Drive/Sheets y devuelve el ID."""
    txt = (valor or "").strip()
    if not txt:
        return ""
    m = re.search(r"/d/([a-zA-Z0-9_-]+)", txt)
    if m:
        return m.group(1)
    m2 = re.search(r"[?&]id=([a-zA-Z0-9_-]+)", txt)
    if m2:
        return m2.group(1)
    return txt


def _url_descarga_drive(file_id: str, fmt: str | None = None) -> str:
    if fmt:
        return f"https://docs.google.com/spreadsheets/d/{file_id}/export?format={fmt}"
    return f"https://drive.google.com/uc?export=download&id={file_id}"


def _parsear_cheques_df(df_raw: pd.DataFrame) -> pd.DataFrame:
    """Extrae Proveedor, Vencimiento y Monto desde CHEQUES NUEVO.xlsx."""
    if df_raw is None or df_raw.empty:
        return pd.DataFrame(columns=["Proveedor", "Vencimiento", "Monto", "Estado"])

    registros = []
    proveedores_cols = {
        idx: str(col).strip()
        for idx, col in enumerate(df_raw.columns)
        if str(col).strip() and not str(col).strip().lower().startswith("unnamed")
    }

    patron_item = re.compile(
        r"\$?\s*([\d.,]+)\s*(?:\(([^)]+)\)|(\d{1,2}/\d{1,2}(?:/\d{2,4})?))?"
    )

    for _, fila in df_raw.iterrows():
        for col_idx, prov in proveedores_cols.items():
            val = fila.iloc[col_idx]
            if pd.isna(val):
                continue
            txt = str(val).strip()
            if not txt or txt.upper() in ("NAN", "NONE"):
                continue

            for m in patron_item.finditer(txt):
                raw_monto = (m.group(1) or "").strip()
                raw_fecha = (m.group(2) or m.group(3) or "").strip()
                if not raw_monto:
                    continue

                limpio = raw_monto.replace(" ", "")
                if "," in limpio and "." in limpio:
                    limpio = limpio.replace(".", "").replace(",", ".")
                elif "." in limpio and len(limpio.split(".")[-1]) == 3:
                    limpio = limpio.replace(".", "")
                elif "," in limpio:
                    limpio = limpio.replace(",", ".")

                try:
                    monto_num = float(limpio)
                except ValueError:
                    continue

                if monto_num > 0:
                    registros.append({
                        "Proveedor": prov.upper(),
                        "Vencimiento": raw_fecha if raw_fecha else "Sin fecha",
                        "Monto": monto_num,
                        "Estado": "Pendiente",
                    })

    if not registros:
        return pd.DataFrame(columns=["Proveedor", "Vencimiento", "Monto", "Estado"])
    return pd.DataFrame(registros)


@st.cache_data(ttl=120, show_spinner="Cargando datos desde Google Drive...")
def cargar_datos_drive(id_pedidos: str, id_cheques: str):
    df_pedidos = pd.DataFrame(
        columns=[
            "ID", "Fecha", "Fecha Cobro", "Cliente", "Material",
            "Medida", "Monto Total", "Estado", "Cobró", "Impresión", "Observaciones"
        ]
    )
    df_cheques = pd.DataFrame(columns=["Proveedor", "Vencimiento", "Monto", "Estado"])
    df_gastos_sem = pd.DataFrame(columns=["Semana", "Socio", "Concepto", "Monto"])
    errores = []

    pid = _extraer_id_drive(id_pedidos)
    cid = _extraer_id_drive(id_cheques)

    if pid:
        try:
            url_p = _url_descarga_drive(pid)
            try:
                df_p = pd.read_csv(url_p, encoding="utf-8", on_bad_lines="skip")
            except Exception:
                df_p = pd.read_csv(
                    url_p,
                    encoding="latin1",
                    engine="python",
                    on_bad_lines="skip",
                )

            df_p.columns = [str(c).strip() for c in df_p.columns]
            cols_base = [
                c for c in [
                    "Fecha", "Cliente", "Material", "Medida", "Copias",
                    "Monto Total", "Estado", "Cobró", "Observaciones",
                    "Impresión", "Fecha Cobro", "ID"
                ] if c in df_p.columns
            ]
            if cols_base:
                df_pedidos = df_p[cols_base].dropna(how="all").copy()
            else:
                errores.append(
                    f"El CSV no tiene las columnas esperadas. Columnas encontradas: {list(df_p.columns[:10])}"
                )
        except Exception as e:
            errores.append(f"Pedidos (Drive): {e}")

    if cid:
        try:
            df_c_raw = pd.read_excel(_url_descarga_drive(cid, fmt="xlsx"))
            df_cheques = _parsear_cheques_df(df_c_raw)
        except Exception as e:
            errores.append(f"Cheques (Drive): {e}")

    return df_pedidos, df_cheques, df_gastos_sem, errores


# ==============================================================================
# CONFIGURACIÓN DE FUENTE (GOOGLE DRIVE)
# ==============================================================================
with st.expander("☁️ Fuente Google Drive (solo consulta)", expanded=False):
    st.caption(
        "Pegá el enlace o ID de Google Drive de cada archivo (deben tener acceso de lectura para el enlace)."
    )
    c_d1, c_d2 = st.columns(2)
    with c_d1:
        inp_id_pedidos = st.text_input(
            "Enlace o ID de PEDIDOS_CLIENTES.csv:",
            value=st.session_state.get("drive_id_pedidos_cfg", ""),
            key="inp_drive_pedidos",
        )
    with c_d2:
        inp_id_cheques = st.text_input(
            "Enlace o ID de CHEQUES NUEVO.xlsx:",
            value=st.session_state.get("drive_id_cheques_cfg", ""),
            key="inp_drive_cheques",
        )

    b_g1, b_g2 = st.columns(2)
    with b_g1:
        if st.button("🔄 Actualizar desde Drive", use_container_width=True):
            st.session_state["drive_id_pedidos_cfg"] = inp_id_pedidos.strip()
            st.session_state["drive_id_cheques_cfg"] = inp_id_cheques.strip()
            st.cache_data.clear()
            st.rerun()
    with b_g2:
        # Fallback local si existe el CSV en la misma carpeta del script
        if os.path.exists("PEDIDOS_CLIENTES.csv") and st.button("📂 Leer CSV local (prueba)", use_container_width=True):
            st.session_state["usar_csv_local_prueba"] = True
            st.rerun()

if st.session_state.get("usar_csv_local_prueba") and os.path.exists("PEDIDOS_CLIENTES.csv"):
    df_pedidos = pd.read_csv("PEDIDOS_CLIENTES.csv")
    df_cheques = pd.DataFrame(columns=["Proveedor", "Vencimiento", "Monto", "Estado"])
    df_gastos_sem = pd.DataFrame(columns=["Semana", "Socio", "Concepto", "Monto"])
    errores_drive = []
else:
    df_pedidos, df_cheques, df_gastos_sem, errores_drive = cargar_datos_drive(
        st.session_state.get("drive_id_pedidos_cfg", inp_id_pedidos),
        st.session_state.get("drive_id_cheques_cfg", inp_id_cheques),
    )

for err in errores_drive:
    st.warning(f"⚠️ {err}")


# ==============================================================================
# SOLO 3 PESTAÑAS PARA EL CELULAR
# ==============================================================================
tab_ped, tab_chq, tab_sem = st.tabs([
    "📋 Pedidos",
    "💳 Cheques",
    "📅 Cierre Semanal",
])


# ------------------------------------------------------------------------------
# PESTAÑA 1: HISTORIAL GENERAL DE TRABAJOS
# ------------------------------------------------------------------------------
with tab_ped:
    st.markdown("### 📋 Historial General de Trabajos")

    if df_pedidos.empty:
        st.info("Sin pedidos cargados aún. Pegá el enlace o ID de `PEDIDOS_CLIENTES.csv` en '☁️ Fuente Google Drive'.")
    else:
        df_v = df_pedidos.copy()
        df_v["Monto_Num"] = pd.to_numeric(df_v["Monto Total"], errors="coerce").fillna(0.0)

        c_f1, c_f2 = st.columns(2)
        with c_f1:
            solo_deben = st.toggle("🔴 Solo faltan cobrar", key="cel_solo_deben")
        with c_f2:
            solo_imp = st.toggle("🖨️ Solo faltan imprimir", key="cel_solo_imp")

        buscar = st.text_input("🔍 Buscar cliente:", key="cel_buscar_cli")

        if solo_deben and "Estado" in df_v.columns:
            df_v = df_v[df_v["Estado"].isin(["Pendiente", "Entregado Sin Cobrar"])]
        if solo_imp and "Impresión" in df_v.columns:
            df_v = df_v[df_v["Impresión"] == "🔴 Falta Imprimir"]
        if buscar and "Cliente" in df_v.columns:
            df_v = df_v[df_v["Cliente"].astype(str).str.contains(buscar, case=False, na=False)]

        st.markdown(
            f'<div class="kpi-card"><div class="kpi-label">Total en vista</div>'
            f'<div class="kpi-value">${df_v["Monto_Num"].sum():,.0f}</div></div>',
            unsafe_allow_html=True,
        )

        if "Fecha" in df_v.columns:
            df_v["_f"] = pd.to_datetime(df_v["Fecha"], format="%d/%m/%Y", errors="coerce")
            df_v = df_v.sort_values("_f", ascending=False).reset_index(drop=True)

        cols_ver = [
            c for c in ["Fecha", "Cliente", "Material", "Medida", "Monto_Num", "Estado", "Cobró", "Impresión"]
            if c in df_v.columns
        ]
        st.dataframe(
            df_v[cols_ver].rename(columns={"Monto_Num": "Monto ($)"}),
            column_config={"Monto ($)": st.column_config.NumberColumn(format="$%.0f")},
            use_container_width=True,
            hide_index=True,
        )


# ------------------------------------------------------------------------------
# PESTAÑA 2: CHEQUES
# ------------------------------------------------------------------------------
with tab_chq:
    st.markdown("### 💳 Cartera de Cheques")

    if df_cheques.empty:
        st.info("Sin cheques cargados aún. Pegá el enlace o ID de `CHEQUES NUEVO.xlsx` en '☁️ Fuente Google Drive'.")
    else:
        df_c = df_cheques.copy()
        df_c["Monto_Num"] = pd.to_numeric(df_c["Monto"], errors="coerce").fillna(0.0)

        filtro_est_chq = st.radio(
            "Estado:",
            ["Todos", "Pendiente", "Pagado"],
            horizontal=True,
            key="cel_filtro_chq",
        )
        if filtro_est_chq != "Todos" and "Estado" in df_c.columns:
            df_c = df_c[df_c["Estado"] == filtro_est_chq]

        pend_total = (
            float(pd.to_numeric(df_cheques[df_cheques["Estado"] == "Pendiente"]["Monto"], errors="coerce").fillna(0).sum())
            if "Estado" in df_cheques.columns
            else 0.0
        )
        st.markdown(
            f'<div class="kpi-card"><div class="kpi-label">Cheques Pendientes</div>'
            f'<div class="kpi-value">${pend_total:,.0f}</div></div>',
            unsafe_allow_html=True,
        )

        st.dataframe(
            df_c[["Proveedor", "Vencimiento", "Monto_Num", "Estado"]].rename(columns={"Monto_Num": "Monto ($)"}),
            column_config={"Monto ($)": st.column_config.NumberColumn(format="$%.0f")},
            use_container_width=True,
            hide_index=True,
        )


# ------------------------------------------------------------------------------
# PESTAÑA 3: CIERRE SEMANAL
# ------------------------------------------------------------------------------
with tab_sem:
    st.markdown("### 📅 Cierre Semanal")

    if df_pedidos.empty or "Fecha" not in df_pedidos.columns:
        st.info("Sin datos de pedidos para calcular el cierre semanal.")
    else:
        df_s = df_pedidos.copy()
        fecha_col = "Fecha Cobro" if "Fecha Cobro" in df_s.columns else "Fecha"
        df_s["Fecha_Uso"] = df_s[fecha_col].fillna(df_s["Fecha"])
        df_s["Fecha_dt"] = pd.to_datetime(df_s["Fecha_Uso"], format="%d/%m/%Y", errors="coerce")
        df_s = df_s.dropna(subset=["Fecha_dt"])
        df_s["Semana_Inicio"] = df_s["Fecha_dt"].dt.to_period("W").dt.start_time

        semanas = sorted(df_s["Semana_Inicio"].unique(), reverse=True)
        if not semanas:
            st.info("No hay fechas válidas para agrupar por semana.")
        else:
            sem_sel = st.selectbox(
                "Semana:",
                semanas,
                format_func=lambda x: f"{x.strftime('%d/%m/%Y')} al {(x + timedelta(days=6)).strftime('%d/%m/%Y')}",
                key="cel_sem_sel",
            )
            sem_str = sem_sel.strftime("%Y-%m-%d")

            cob_sem = df_s[(df_s["Semana_Inicio"] == sem_sel) & (df_s["Estado"] == "Entregado Cobrado")].copy()
            cob_sem["Monto_Num"] = pd.to_numeric(cob_sem["Monto Total"], errors="coerce").fillna(0.0)

            bruto_rob = float(cob_sem[cob_sem["Cobró"].isin(["ROB", "ROBERTO"])]["Monto_Num"].sum()) if "Cobró" in cob_sem.columns else 0.0
            bruto_gon = float(cob_sem[cob_sem["Cobró"].isin(["GON", "GONZALO"])]["Monto_Num"].sum()) if "Cobró" in cob_sem.columns else 0.0

            g_sem = df_gastos_sem[df_gastos_sem["Semana"] == sem_str] if not df_gastos_sem.empty else pd.DataFrame()
            gas_rob = float(pd.to_numeric(g_sem[g_sem["Socio"] == "ROB"]["Monto"], errors="coerce").fillna(0).sum()) if not g_sem.empty else 0.0
            gas_gon = float(pd.to_numeric(g_sem[g_sem["Socio"] == "GON"]["Monto"], errors="coerce").fillna(0).sum()) if not g_sem.empty else 0.0

            neto_rob = bruto_rob - gas_rob
            neto_gon = bruto_gon - gas_gon
            ajuste = (neto_rob / 2) - (neto_gon / 2)

            c1, c2 = st.columns(2)
            with c1:
                st.markdown(
                    f'<div class="kpi-card"><div class="kpi-label">Neto ROB</div>'
                    f'<div class="kpi-value">${neto_rob:,.0f}</div></div>',
                    unsafe_allow_html=True,
                )
            with c2:
                st.markdown(
                    f'<div class="kpi-card"><div class="kpi-label">Neto GON</div>'
                    f'<div class="kpi-value">${neto_gon:,.0f}</div></div>',
                    unsafe_allow_html=True,
                )

            if ajuste > 0:
                st.warning(f"ROB debe transferir **${abs(ajuste):,.0f}** a GON.")
            elif ajuste < 0:
                st.warning(f"GON debe transferir **${abs(ajuste):,.0f}** a ROB.")
            else:
                st.success("Cuentas empatadas.")