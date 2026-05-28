import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import openpyxl, io, os, warnings
warnings.filterwarnings("ignore")

# ── CONFIGURACIÓN DE PÁGINA ───────────────────────────────────────────────────
st.set_page_config(page_title="Tax Impact Dashboard", page_icon="📊", layout="wide",
                   initial_sidebar_state="collapsed")

# ── ESTILOS CSS ───────────────────────────────────────────────────────────────
st.markdown("""
<style>
.main .block-container{max-width:1400px;padding:1.5rem 2rem;overflow-x:auto;}
section[data-testid="stSidebar"]{display:none;}
.header-box{background:#4D93D9 !important;padding:5px 32px;
    border-radius:12px;margin:0 auto 24px;width:80%;}
.header-title{color:#FFFFFF !important;font-size:26px;font-weight:900;letter-spacing:2px;}
.header-sub{color:#D0E8FF !important;font-size:13px;margin-top:4px;}
.zona-title{font-size:15px;font-weight:900;color:#0A2463;background:#D6EAFF;
    padding:8px 20px;border-radius:8px;margin:20px 0 14px;letter-spacing:1px;
    display:inline-block;min-width:460px;text-align:center;}
.info-box{background:#D6EAFF;border-radius:8px;
    padding:12px 20px;margin:16px auto 8px auto;max-width:860px;
    font-size:13px;color:#0A2463;line-height:1.8;}
.info-box b{color:#0A2463;}
.scenario-label{font-size:13px;font-weight:700;color:#0A2463;background:#D6EAFF;
    padding:6px 24px;border-radius:6px;display:inline-block;letter-spacing:.5px;}
.kpi-box{background:#ffffff;border-radius:10px;padding:14px 18px;
    box-shadow:0 2px 8px rgba(0,0,0,.08);text-align:center;
    width:100%;box-sizing:border-box;}
.kpi-lbl{font-size:10px;font-weight:700;color:#888;text-transform:uppercase;
    letter-spacing:.5px;margin-bottom:6px;}
.kpi-val{font-size:22px;font-weight:900;color:#0A2463;}
/* Responsive móvil */
@media(max-width:768px){
    .main .block-container{padding:0.5rem !important;}
    .header-box{padding:14px 16px !important;}
    .header-title{font-size:16px !important;}
    .header-sub{font-size:11px !important;}
    .kpi-box{width:100% !important;max-width:100% !important;
        margin-right:0 !important;margin-bottom:8px;}
    .kpi-val{font-size:16px !important;}
    .zona-title{font-size:12px !important;padding:8px 12px !important;}
}
/* Móvil + dark mode: refuerzo del header */
@media(max-width:768px) and (prefers-color-scheme:dark){
    .header-box   { background:#4D93D9 !important; }
    .header-title { color:#FFFFFF !important; }
    .header-sub   { color:#D0E8FF !important; }
}
/* Forzar light mode en móvil con dark mode del sistema */
@media(prefers-color-scheme:dark){
    html, body,
    [data-testid="stAppViewContainer"],
    [data-testid="stHeader"],
    [data-testid="stMainBlockContainer"],
    section, div {
        background-color:#FFFFFF !important;
        color:#111827 !important;
    }
    /* Preservar header principal */
    .header-box   { background:#4D93D9 !important; }
    .header-title { color:#FFFFFF !important; }
    .header-sub   { color:#D0E8FF !important; }
    /* Preservar elementos claros */
    .zona-title, .info-box, .scenario-label { background:#D6EAFF !important; }
    .zona-title, .info-box, .info-box b, .info-box *, .scenario-label { color:#0A2463 !important; }
}
</style>""", unsafe_allow_html=True)

# ── CARGA DE DATOS DESDE GOOGLE SHEETS ───────────────────────────────────────
# Cache de 5 minutos para evitar recargas innecesarias
@st.cache_data(ttl=300)
def cargar():
    import requests
    FILE_ID = st.secrets["FILE_ID"]
    r = requests.get(f"https://docs.google.com/spreadsheets/d/{FILE_ID}/export?format=xlsx")
    wb = openpyxl.load_workbook(io.BytesIO(r.content), data_only=True)
    ws = wb["DATA"]
    rows = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if row[0]:
            rows.append({"id": row[0], "zona": row[1], "proyecto": row[2],
                         "metrica": row[3], "actual": row[4], "sin_tx": row[5]})
    return pd.DataFrame(rows)

df_all = cargar()

# ── HELPERS DE FORMATO ────────────────────────────────────────────────────────
def fmt_v(v, pct=False):
    # Retorna "—" si el valor es nulo, porcentaje o dólares según el tipo
    if v is None or (isinstance(v, float) and pd.isna(v)): return "—"
    return f"{v*100:.2f}%" if pct else f"${v:,.0f}"

def es_pct_metrica(m):
    # Métricas que se expresan como porcentaje
    return m in ["CASH ON CASH", "IRR"]

# ── ENCABEZADO PRINCIPAL ──────────────────────────────────────────────────────
st.markdown("""
<div class="header-box">
  <div class="header-title">PORTFOLIO — TAX IMPACT DASHBOARD</div>
  <div class="header-sub">Current Value vs Tax-Free Value · CASH ON CASH · IRR · FCF FROM FINANCING</div>
</div>""", unsafe_allow_html=True)

# ── SECCIÓN: RESUMEN FCF PORTFOLIO ───────────────────────────────────────────
# Muestra FCF With Tax, Tax-Free y Variance para Zona 1, Zona 2 y Total
def resumen_fcf(df):
    def get_fcf(zona):
        sub = df[(df["zona"] == zona) & (df["metrica"] == "FCF FROM FINANCING")]
        return sub["actual"].sum(), sub["sin_tx"].sum()

    a_z1, s_z1 = get_fcf("ZONA 1")
    a_z2, s_z2 = get_fcf("ZONA 2")
    a_tot = a_z1 + a_z2
    s_tot = s_z1 + s_z2

    st.markdown('<div style="text-align:center;"><div class="zona-title">PORTFOLIO SUMMARY — FCF FROM FINANCING</div></div>', unsafe_allow_html=True)
    st.markdown("<div style='margin-top:28px;'></div>", unsafe_allow_html=True)

    def pct_var(base, diff):
        # Variación porcentual entre escenarios
        if not base or base == 0: return ""
        return f"{diff / base * 100:+.1f}%"

    def pct_of(part, total):
        # Participación de cada zona sobre el total
        if not total or total == 0: return ""
        return f"{part / total * 100:.1f}%"

    # Tres filas: With Tax / Tax-Free / Variance
    filas = [
        [("FCF With Tax — Zone 1", f"${a_z1:,.0f}",  pct_of(a_z1,  a_tot)),
         ("FCF With Tax — Zone 2", f"${a_z2:,.0f}",  pct_of(a_z2,  a_tot)),
         ("FCF With Tax — Total",  f"${a_tot:,.0f}", pct_of(a_tot, a_tot))],
        [("FCF Tax-Free — Zone 1", f"${s_z1:,.0f}",  pct_of(s_z1,  s_tot)),
         ("FCF Tax-Free — Zone 2", f"${s_z2:,.0f}",  pct_of(s_z2,  s_tot)),
         ("FCF Tax-Free — Total",  f"${s_tot:,.0f}", pct_of(s_tot, s_tot))],
        [("FCF Variance — Zone 1", f"${s_z1-a_z1:+,.0f}",   pct_var(a_z1,  s_z1-a_z1)),
         ("FCF Variance — Zone 2", f"${s_z2-a_z2:+,.0f}",   pct_var(a_z2,  s_z2-a_z2)),
         ("FCF Variance — Total",  f"${s_tot-a_tot:+,.0f}", pct_var(a_tot, s_tot-a_tot))],
    ]
    for fila in filas:
        html = '<div style="display:flex;gap:10px;justify-content:center;margin-bottom:10px;">'
        for lbl, val, pct in fila:
            if pct:
                # Verde si positivo, rojo si negativo, azul si neutro
                if pct.startswith("+"):
                    color = "#2ECC71"
                elif pct.startswith("-"):
                    color = "#E74C3C"
                else:
                    color = "#0070C0"
                pct_html = f'<div style="font-size:13px;font-weight:700;color:{color};margin-top:4px;">{pct}</div>'
            else:
                pct_html = ""
            html += (
                '<div class="kpi-box">'
                f'<div class="kpi-lbl">{lbl}</div>'
                f'<div class="kpi-val">{val}</div>'
                f'{pct_html}'
                '</div>'
            )
        html += '</div>'
        st.markdown(html, unsafe_allow_html=True)

    # Nota explicativa de los tres escenarios
    st.markdown("""
<div class="info-box">
  <b>FCF With Tax</b> — Free cash flow from financing applying the current tax burden (base scenario).<br>
  <b>FCF Tax-Free</b> — Free cash flow under the tax exemption benefit.<br>
  <b>FCF Variance</b> — Absolute and relative difference between both scenarios; reflects the direct economic impact of the tax benefit.
</div>""", unsafe_allow_html=True)
    st.markdown("<hr style='border:none;border-top:2px solid #e0e0e0;width:80%;margin:24px auto;'>",
                unsafe_allow_html=True)

# ── SECCIÓN: KPIs POR ZONA Y PROYECTO ────────────────────────────────────────
# Permite filtrar por proyecto y muestra métricas en escenario Base vs Tax-Free
def render_zona_etiquetas(zona, metricas, default_proyecto):
    df_z = df_all[df_all["zona"] == zona]
    opciones = ["All"] + sorted(df_z["proyecto"].unique().tolist())
    default_idx = opciones.index(default_proyecto) if default_proyecto in opciones else 0

    # Selector de proyecto
    _, col_sel, _ = st.columns([0.2, 0.8, 4])
    with col_sel:
        proyecto_sel = st.selectbox("Project:", opciones, index=default_idx,
                                    key=f"sel_etq_{zona}")

    st.markdown(f'<div style="text-align:center;"><div class="zona-title">{zona}</div></div>',
                unsafe_allow_html=True)

    # Filtrar por proyecto seleccionado
    df_p = df_z.copy() if proyecto_sel == "All" else df_z[df_z["proyecto"] == proyecto_sel].copy()

    def kpi_card(lbl, val, pct_str, pct_color):
        # Tarjeta KPI con label, valor principal y variación opcional
        pct_html = (f'<div style="font-size:13px;font-weight:700;color:{pct_color};margin-top:4px;">{pct_str}</div>'
                    if pct_str else
                    '<div style="font-size:13px;margin-top:4px;visibility:hidden;">—</div>')
        return (
            '<div class="kpi-box">'
            f'<div class="kpi-lbl">{lbl}</div>'
            f'<div class="kpi-val">{val}</div>'
            f'{pct_html}'
            '</div>'
        )

    def render_fila(titulo, campo, show_pct=True):
        # Renderiza una fila de KPIs para un campo dado (actual o sin_tx)
        st.markdown(
            f'<div style="text-align:center;margin:18px 0 8px;">'
            f'<span class="scenario-label">{titulo}</span>'
            f'</div>', unsafe_allow_html=True)
        html = '<div style="display:grid;grid-template-columns:repeat(3,1fr);gap:10px;max-width:960px;margin:0 auto;">'
        for m in metricas:
            row_m = df_p[df_p["metrica"] == m]
            if row_m.empty:
                html += kpi_card(m, "—", "", "")
                continue
            actual = row_m["actual"].sum()
            sin_tx = row_m["sin_tx"].sum()
            val    = row_m[campo].sum()
            base   = actual if campo == "sin_tx" else sin_tx
            diff   = val - base
            val_fmt = fmt_v(val, pct=es_pct_metrica(m))
            if show_pct and base and base != 0:
                if es_pct_metrica(m):
                    # Para IRR y COC: diferencia en puntos porcentuales
                    pp = diff * 100
                    pct_str   = f"{pp:+.2f} pp"
                    pct_color = "#2ECC71" if pp >= 0 else "#E74C3C"
                else:
                    # Para valores monetarios: variación porcentual
                    pct_change = diff / abs(base) * 100
                    pct_str    = f"{pct_change:+.1f}%"
                    pct_color  = "#2ECC71" if pct_change >= 0 else "#E74C3C"
            else:
                pct_str, pct_color = "", ""
            html += kpi_card(m, val_fmt, pct_str, pct_color)
        html += '</div>'
        st.markdown(html, unsafe_allow_html=True)

    st.markdown("<div style='margin-top:12px;'></div>", unsafe_allow_html=True)
    render_fila("BASE SCENARIO — WITH TAX", "actual", show_pct=False)
    st.markdown("<div style='margin-top:28px;'></div>", unsafe_allow_html=True)
    render_fila("TAX-FREE", "sin_tx", show_pct=True)

    # Nota explicativa de escenarios
    st.markdown("""
<div class="info-box">
  <b>Base Scenario — With Tax</b> — Project metrics under the current tax burden without applying exemption benefits.<br>
  <b>Tax-Free</b> — Project metrics applying the tax exemption. The percentage indicates the variance from the base scenario.
</div>""", unsafe_allow_html=True)
    st.markdown("<div style='margin-top:20px;'></div>", unsafe_allow_html=True)

# ── RENDERIZADO PRINCIPAL ─────────────────────────────────────────────────────
st.markdown("<div style='margin-top:48px;'></div>", unsafe_allow_html=True)

# Resumen consolidado del portfolio
resumen_fcf(df_all)

# Zona 1
render_zona_etiquetas(
    zona             = "ZONA 1",
    metricas         = ["CASH ON CASH", "FCF FROM FINANCING", "IRR"],
    default_proyecto = "PROYECTO 1"
)

st.markdown("<hr style='border:none;border-top:2px solid #e0e0e0;width:80%;margin:32px auto;'>",
            unsafe_allow_html=True)

# Zona 2
render_zona_etiquetas(
    zona             = "ZONA 2",
    metricas         = ["CASH ON CASH", "FCF FROM FINANCING", "IRR"],
    default_proyecto = "PROYECTO 8"
)
