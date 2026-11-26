import re
from datetime import datetime, timedelta
import pandas as pd
import streamlit as st

# PONÉ ACÁ TUS ENLACES O IDs FIJOS DE GOOGLE DRIVE
URL_PEDIDOS_DEFAULT = "https://drive.google.com/file/d/1N2ilCbRRKguBYFwOnISyuDdYuLEVzFfT/view?usp=drive_link"
URL_GASTOS_DEFAULT = "https://drive.google.com/file/d/13kC1dwUT9y4JjTtwt5f-RiunMjdiKiR3/view?usp=drive_link"  


st.set_page_config(
    page_title="Informe Imprenta",
    page_icon="🧾",
    layout="centered",
)

st.markdown(
    """
    <style>
        .stApp { background-color: #0b0f19; color: #e2e8f0; }
        [data-testid="stHeader"], [data-testid="stToolbar"], .stDeployButton, #MainMenu {
            display: none !important;
        }
        .block-container { padding-top: 1rem !important; padding-bottom: 2rem !important; }
    </style>
    """,
    unsafe_allow_html=True,
)


def _extraer_id_drive(valor: str) -> str:
    txt = (valor or "").strip()
    if not txt or txt.startswith("PEGA_AQUI_"):
        return ""
    m = re.search(r"/d/([a-zA-Z0-9_-]+)", txt)
    if m:
        return m.group(1)
    m2 = re.search(r"[?&]id=([a-zA-Z0-9_-]+)", txt)
    if m2:
        return m2.group(1)
    return txt


def _url_descarga_drive(file_id: str) -> str:
    return f"https://drive.google.com/uc?export=download&id={file_id}"


@st.cache_data(ttl=15, show_spinner="Cargando informe...")
def cargar_datos_drive(id_pedidos: str, id_gastos: str):
    df_pedidos = pd.DataFrame()
    df_gastos = pd.DataFrame()

    pid = _extraer_id_drive(id_pedidos)
    gid = _extraer_id_drive(id_gastos)

    if pid:
        try:
            url_p = _url_descarga_drive(pid)
            try:
                df_pedidos = pd.read_csv(url_p, encoding="utf-8", on_bad_lines="skip")
            except Exception:
                df_pedidos = pd.read_csv(url_p, encoding="latin1", engine="python", on_bad_lines="skip")
            df_pedidos.columns = [str(c).strip() for c in df_pedidos.columns]
        except Exception:
            pass

    if gid:
        try:
            url_g = _url_descarga_drive(gid)
            try:
                df_gastos = pd.read_csv(url_g, encoding="utf-8", on_bad_lines="skip")
            except Exception:
                df_gastos = pd.read_csv(url_g, encoding="latin1", engine="python", on_bad_lines="skip")
            df_gastos.columns = [str(c).strip() for c in df_gastos.columns]
        except Exception:
            pass

    return df_pedidos, df_gastos


id_ped_cfg = st.secrets.get("DRIVE_ID_PEDIDOS", URL_PEDIDOS_DEFAULT) if hasattr(st, "secrets") else URL_PEDIDOS_DEFAULT
id_gas_cfg = st.secrets.get("DRIVE_ID_GASTOS", URL_GASTOS_DEFAULT) if hasattr(st, "secrets") else URL_GASTOS_DEFAULT

df_pedidos, df_gastos = cargar_datos_drive(id_ped_cfg, id_gas_cfg)

if df_pedidos.empty:
    st.warning("Falta configurar `URL_PEDIDOS_DEFAULT` en `app_celular.py` (o verificar que el archivo en Drive tenga acceso de lectura para cualquiera con el enlace).")
else:
    df_i = df_pedidos.copy()
    df_i["Monto_Num"] = pd.to_numeric(df_i.get("Monto Total", 0), errors="coerce").fillna(0.0)

    fecha_col_i = "Fecha Cobro" if "Fecha Cobro" in df_i.columns else "Fecha"
    df_i["Fecha_Cobro_Uso"] = df_i[fecha_col_i].fillna(df_i["Fecha"])
    df_i["Fecha_Cobro_dt"] = pd.to_datetime(df_i["Fecha_Cobro_Uso"], format="%d/%m/%Y", errors="coerce")
    df_i["Fecha_Ped_dt"] = pd.to_datetime(df_i["Fecha"], format="%d/%m/%Y", errors="coerce")

    semanas_i = sorted(df_i["Fecha_Cobro_dt"].dropna().dt.to_period("W").dt.start_time.unique(), reverse=True)
    sem_i = semanas_i[0] if semanas_i else None

    bruto_rob_i = bruto_gon_i = gas_rob_i = gas_gon_i = 0.0
    rango_txt_i = "Semana actual"

    if sem_i is not None:
        rango_txt_i = f"{sem_i.strftime('%d/%m')} al {(sem_i + timedelta(days=6)).strftime('%d/%m')}"
        sem_str_i = sem_i.strftime("%Y-%m-%d")

        cob_i = df_i[
            (df_i["Fecha_Cobro_dt"].dt.to_period("W").dt.start_time == sem_i)
            & (df_i["Estado"] == "Entregado Cobrado")
        ]
        bruto_rob_i = float(cob_i[cob_i["Cobró"].isin(["ROB", "ROBERTO"])]["Monto_Num"].sum()) if "Cobró" in cob_i.columns else 0.0
        bruto_gon_i = float(cob_i[cob_i["Cobró"].isin(["GON", "GONZALO"])]["Monto_Num"].sum()) if "Cobró" in cob_i.columns else 0.0

        if not df_gastos.empty and "Semana" in df_gastos.columns:
            g_i = df_gastos[df_gastos["Semana"] == sem_str_i]
            if not g_i.empty:
                gas_rob_i = float(pd.to_numeric(g_i[g_i["Socio"] == "ROB"]["Monto"], errors="coerce").fillna(0).sum())
                gas_gon_i = float(pd.to_numeric(g_i[g_i["Socio"] == "GON"]["Monto"], errors="coerce").fillna(0).sum())

    neto_rob_i = bruto_rob_i - gas_rob_i
    neto_gon_i = bruto_gon_i - gas_gon_i
    ajuste_i = (neto_rob_i / 2) - (neto_gon_i / 2)

    pend_gral_i = float(
        df_i[df_i["Estado"].isin(["Pendiente", "Entregado Sin Cobrar"])]["Monto_Num"].sum()
    ) if "Estado" in df_i.columns else 0.0

    if ajuste_i > 0:
        cierre_linea = f"⚖️ ROB debe transferir ${abs(ajuste_i):,.0f} a GON."
    elif ajuste_i < 0:
        cierre_linea = f"⚖️ GON debe transferir ${abs(ajuste_i):,.0f} a ROB."
    else:
        cierre_linea = "⚖️ Cuentas empatadas."

    df_hist_corto = df_i.sort_values("Fecha_Ped_dt", ascending=False).head(30).copy()

    lineas_trabajos = []
    fecha_anterior = None
    for _, r in df_hist_corto.iterrows():
        f_str = str(r.get("Fecha", "")).strip()
        if fecha_anterior is not None and f_str != fecha_anterior:
            lineas_trabajos.append("")
        fecha_anterior = f_str

        cli_str = str(r.get("Cliente", "")).strip()
        mat_str = str(r.get("Material", "")).strip()
        mto_str = f"${float(r.get('Monto_Num', 0.0)):,.0f}"
        imp_raw = str(r.get("Impresión", "")).strip()
        punto_imp = "🟢" if "FINALIZADO" in imp_raw.upper() else "🔴"
        lineas_trabajos.append(f"• {f_str} | {cli_str} | {mat_str} | {mto_str} | {punto_imp}")

    texto_informe = (
        f"🧾 INFORME IMPRENTA — {datetime.now().strftime('%H:%M hs (%d/%m)')}\n"
        f"Pendiente general: ${pend_gral_i:,.0f}\n"
        f"────────────────────────────────────\n"
        f"📅 CIERRE SEMANAL ({rango_txt_i})\n"
        f"• ROB: Cobró ${bruto_rob_i:,.0f} | Gastos ${gas_rob_i:,.0f}\n"
        f"  👉 Neto Limpio ROB: ${neto_rob_i:,.0f}\n"
        f"• GON: Cobró ${bruto_gon_i:,.0f} | Gastos ${gas_gon_i:,.0f}\n"
        f"  👉 Neto Limpio GON: ${neto_gon_i:,.0f}\n"
        f"{cierre_linea}\n"
        f"────────────────────────────────────\n"
        f"📋 HISTORIAL\n"
        + ("\n".join(lineas_trabajos) if lineas_trabajos else "• Sin trabajos recientes")
    )

    st.code(texto_informe, language="text")