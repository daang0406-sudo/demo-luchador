# -*- coding: utf-8 -*-
"""
PLATAFORMA DE GESTIÓN E INTELIGENCIA EMPRESARIAL — EL LUCHADOR
MÓDULO DE RECURSOS HUMANOS (ETAPA 1)
Desarrollado bajo Streamlit conforme a ESPECIFICACION_TECNICA_ARQUITECTURA_RRHH.md
"""

import calendar
import hashlib
import html
import io
import re
import random
import unicodedata
from datetime import date, datetime, timedelta
from pathlib import Path
from zoneinfo import ZoneInfo

import holidays
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

# ============================================================
# 1. CONFIGURACIÓN GENERAL Y CONSTANTES
# ============================================================

st.set_page_config(
    page_title="El Luchador | Gestión Integral de Recursos Humanos",
    page_icon="👥",
    layout="wide",
    initial_sidebar_state="expanded",
)

BASE = Path(__file__).resolve().parent
RUTA_EXCEL = BASE / "data" / "sueldos.xlsx"

# Paleta corporativa El Luchador
NEGRO = "#090A10"
AZUL_PRIMARIO = "#2F6BFF"
AZUL_OSCURO = "#1D346F"
NARANJA = "#FF7A1A"
BLANCO = "#F5F7FF"
GRIS_TEXTO = "#CBD5E1"
GRIS_BORDE = "#2B354C"
SUPERFICIE = "#141825"
SUPERFICIE_CARD = "#191E2E"
VERDE_EXITO = "#22C55E"
ROJO_PELIGRO = "#EF4444"
AMARILLO_ALERTA = "#F59E0B"

MESES = [
    "Enero", "Febrero", "Marzo", "Abril",
    "Mayo", "Junio", "Julio", "Agosto",
    "Septiembre", "Octubre", "Noviembre", "Diciembre",
]

DIAS_SEMANA = [
    "Lunes", "Martes", "Miércoles",
    "Jueves", "Viernes", "Sábado", "Domingo",
]

# Estructura Multiempresa y Multisucursal
RAZONES_SOCIALES = ["Misión", "Il Comendatori"]

SUCURSALES_INFO = {
    "Entrada": {
        "numero": 1,
        "nombre_completo": "Sucursal 1 Entrada",
        "razon_social": "Misión",
        "localidad": "Pilar",
    },
    "Ruta 26": {
        "numero": 2,
        "nombre_completo": "Sucursal 2 Ruta 26",
        "razon_social": "Il Comendatori",
        "localidad": "Del Viso",
    },
    "Colectora": {
        "numero": 3,
        "nombre_completo": "Sucursal 3 Colectora",
        "razon_social": "Il Comendatori",
        "localidad": "Escobar",
    },
    "Savio": {
        "numero": 4,
        "nombre_completo": "Sucursal 4 Savio",
        "razon_social": "Misión",
        "localidad": "Maquinista Savio",
    },
}

ROLES_RBAC = [
    "Administrador General",
    "Gerente de Recursos Humanos",
    "Responsable de Sucursal",
    "Supervisor de Área",
    "Asesor Laboral / Auditor",
]

UNIDADES_NEGOCIO = [
    "Carnicería",
    "Verdulería",
    "Cajas / Facturación",
    "Fiambrería",
    "Producción / Avícola",
    "Almacén",
    "Administración y Dirección",
]

DOTACION_MINIMA_DEFECTO = {
    "Carnicería": 2,
    "Verdulería": 2,
    "Cajas / Facturación": 1,
    "Fiambrería": 1,
    "Producción / Avícola": 1,
    "Almacén": 1,
    "Administración y Dirección": 1,
}

# ============================================================
# 2. ESTILOS VISUALES Y CORRECCIÓN DE CONTRASTE (CSS)
# ============================================================

st.markdown(
    f"""
<style>
/* 1. ESTRUCTURA GLOBAL Y FONDOS */
.stApp {{
    background: {NEGRO} !important;
    color: {BLANCO} !important;
    font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif;
}}

/* ELIMINACIÓN TOTAL DE LA BARRA BLANCA SUPERIOR */
header, [data-testid="stHeader"], .stAppHeader {{
    background-color: {NEGRO} !important;
    background: {NEGRO} !important;
    color: {BLANCO} !important;
    border-bottom: 1px solid #1C2436 !important;
}}

[data-testid="stDecoration"] {{
    display: none !important;
}}

[data-testid="stToolbar"] {{
    background-color: transparent !important;
    background: transparent !important;
    color: {BLANCO} !important;
}}

[data-testid="stToolbar"] button, [data-testid="stToolbar"] svg {{
    color: {BLANCO} !important;
    fill: {BLANCO} !important;
}}

/* 2. SIDEBAR CON CONTRASTE ÓPTIMO */
[data-testid="stSidebar"] {{
    background: #0E121D !important;
    border-right: 1px solid #232A3B !important;
}}

[data-testid="stSidebar"] label,
[data-testid="stSidebar"] [data-testid="stWidgetLabel"] p,
[data-testid="stSidebar"] [data-testid="stWidgetLabel"] span {{
    color: #E2E8F0 !important;
    font-size: 13.5px !important;
    font-weight: 600 !important;
}}

[data-testid="stSidebar"] p, [data-testid="stSidebar"] span {{
    color: #CBD5E1 !important;
}}

[data-testid="stSidebar"] h1, [data-testid="stSidebar"] h2, [data-testid="stSidebar"] h3, [data-testid="stSidebar"] h4 {{
    color: #FFFFFF !important;
    font-weight: 700 !important;
}}

/* NAVEGACIÓN RADIO BUTTONS */
[data-testid="stSidebar"] [data-testid="stRadio"] label {{
    padding: 8px 12px !important;
    border-radius: 9px !important;
    margin-bottom: 3px !important;
    color: #CBD5E1 !important;
    font-size: 13.5px !important;
    font-weight: 500 !important;
    transition: background 0.15s ease, color 0.15s ease !important;
}}

[data-testid="stSidebar"] [data-testid="stRadio"] label:hover {{
    background: #1A2338 !important;
    color: #FFFFFF !important;
}}

/* 3. CONTENEDOR PRINCIPAL Y ENCABEZADOS */
.block-container {{
    padding-top: 1.5rem !important;
    padding-bottom: 2.5rem !important;
    max-width: 1580px !important;
}}

h1, h2, h3, h4 {{
    color: #FFFFFF !important;
    font-weight: 700 !important;
}}

/* 4. TARJETAS DE MÉTRICAS */
[data-testid="stMetric"] {{
    border-radius: 14px !important;
    background: {SUPERFICIE} !important;
    border: 1px solid {GRIS_BORDE} !important;
    padding: 14px 18px !important;
    box-shadow: 0 4px 14px rgba(0,0,0,0.3) !important;
    transition: transform 0.15s ease, border-color 0.15s ease !important;
}}
[data-testid="stMetric"]:hover {{
    border-color: {AZUL_PRIMARIO} !important;
    transform: translateY(-2px) !important;
}}

[data-testid="stMetricLabel"] p {{
    color: #94A3B8 !important;
    font-size: 12px !important;
    font-weight: 700 !important;
    text-transform: uppercase !important;
    letter-spacing: 0.6px !important;
    white-space: normal !important;
    word-break: break-word !important;
}}

[data-testid="stMetricValue"] {{
    color: #FFFFFF !important;
    font-size: 25px !important;
    font-weight: 800 !important;
}}

/* 5. SELECTBOXES, INPUTS Y CONTROLES BASEWEB */
[data-testid="stTextInput"] input,
[data-testid="stTextArea"] textarea,
[data-testid="stNumberInput"] input,
[data-testid="stDateInput"] input,
div[data-baseweb="input"] > div,
div[data-baseweb="base-input"] > div,
div[data-baseweb="select"] > div {{
    background: {NARANJA} !important;
    border: 2px solid {NARANJA} !important;
    color: {NEGRO} !important;
    border-radius: 12px !important;
    font-weight: 700 !important;
    min-height: 44px !important;
    box-shadow: none !important;
}}

[data-testid="stTextInput"] input::placeholder,
[data-testid="stTextArea"] textarea::placeholder,
[data-testid="stNumberInput"] input::placeholder,
[data-testid="stDateInput"] input::placeholder {{
    color: rgba(9, 10, 16, 0.75) !important;
}}

[data-testid="stTextInput"] input,
[data-testid="stTextArea"] textarea,
[data-testid="stNumberInput"] input,
[data-testid="stDateInput"] input,
div[data-baseweb="select"] span,
div[data-baseweb="input"] input {{
    color: {NEGRO} !important;
    font-weight: 700 !important;
    -webkit-text-fill-color: {NEGRO} !important;
}}

[data-testid="stTextInput"] svg,
[data-testid="stNumberInput"] svg,
[data-testid="stDateInput"] svg,
div[data-baseweb="select"] svg {{
    fill: {NEGRO} !important;
    color: {NEGRO} !important;
}}

[data-testid="stTextInput"] input:focus,
[data-testid="stTextArea"] textarea:focus,
[data-testid="stNumberInput"] input:focus,
[data-testid="stDateInput"] input:focus,
div[data-baseweb="input"] > div:focus-within,
div[data-baseweb="select"] > div:focus-within {{
    border: 2px solid {AZUL_PRIMARIO} !important;
    box-shadow: 0 0 0 3px rgba(47,107,255,0.18) !important;
}}

/* Menús desplegables */
div[data-baseweb="popover"], div[data-baseweb="menu"], ul[role="listbox"] {{
    background-color: #111827 !important;
    border: 1px solid #334155 !important;
    border-radius: 12px !important;
}}

li[role="option"] {{
    color: #E2E8F0 !important;
    font-size: 13.5px !important;
    font-weight: 600 !important;
}}

li[role="option"]:hover, li[aria-selected="true"] {{
    background-color: {AZUL_PRIMARIO} !important;
    color: #FFFFFF !important;
}}

label, [data-testid="stWidgetLabel"] p, [data-testid="stWidgetLabel"] span {{
    color: #E2E8F0 !important;
    font-size: 13.5px !important;
    font-weight: 700 !important;
}}

/* 6. BOTONES */
.stButton > button[kind="primary"] {{
    background: {NARANJA} !important;
    color: {NEGRO} !important;
    border: 0 !important;
    border-radius: 12px !important;
    font-weight: 800 !important;
    padding: 8px 18px !important;
    transition: all 0.2s ease !important;
}}
.stButton > button[kind="primary"]:hover {{
    background: #FF8F3C !important;
    box-shadow: 0 4px 16px rgba(255, 122, 26, 0.35) !important;
}}

.stButton > button {{
    border-radius: 12px !important;
    border: 1px solid {AZUL_PRIMARIO} !important;
    background: {AZUL_PRIMARIO} !important;
    color: {BLANCO} !important;
    font-weight: 700 !important;
}}
.stButton > button:hover {{
    background: #1F57E6 !important;
    border-color: #1F57E6 !important;
    box-shadow: 0 4px 16px rgba(47, 107, 255, 0.35) !important;
}}

/* 7. HERO BANNER Y TARJETAS */
.hero {{
    border: 1px solid #344C8C;
    border-radius: 18px;
    padding: 24px 30px;
    background: linear-gradient(115deg, #162447 0%, #111524 100%);
    margin-bottom: 20px;
    box-shadow: 0 8px 24px rgba(0,0,0,0.35);
}}
.hero .eyebrow {{
    font-size: 11px;
    font-weight: 800;
    letter-spacing: 2px;
    color: #9BB8FF;
    text-transform: uppercase;
}}
.hero h1 {{
    font-size: 28px !important;
    margin: 8px 0;
}}
.hero p {{
    margin: 0;
    color: #CBD5EF;
    font-size: 14.5px;
}}

.panel-card {{
    background: {SUPERFICIE};
    border: 1px solid {GRIS_BORDE};
    border-radius: 16px;
    padding: 20px;
    margin-bottom: 18px;
}}

/* 8. PESTAÑAS (TABS) */
[data-baseweb="tab-list"] {{
    background: #101420 !important;
    border-radius: 10px !important;
    padding: 4px !important;
    border: 1px solid #283147 !important;
    gap: 4px !important;
}}

[data-baseweb="tab"] {{
    color: #94A3B8 !important;
    font-weight: 600 !important;
    font-size: 13.5px !important;
    border-radius: 8px !important;
    padding: 8px 16px !important;
    border: none !important;
}}

[aria-selected="true"][data-baseweb="tab"] {{
    background-color: #1F2B47 !important;
    color: #60A5FA !important;
    border: 1px solid #3B82F6 !important;
}}

/* 9. DATAFRAMES Y TABLAS */
[data-testid="stDataFrame"] {{
    border: 1px solid {GRIS_BORDE} !important;
    border-radius: 10px !important;
    background: {SUPERFICIE} !important;
    overflow: hidden !important;
}}

/* BADGES */
.badge {{
    display: inline-block;
    padding: 3px 10px;
    border-radius: 12px;
    font-size: 11.5px;
    font-weight: 700;
    letter-spacing: 0.3px;
}}
.badge-blue {{ background: #1F3668; color: #94BAFF; border: 1px solid #355BA8; }}
.badge-green {{ background: #133924; color: #72EAA5; border: 1px solid #236E46; }}
.badge-orange {{ background: #3E2413; color: #FFAF75; border: 1px solid #824921; }}
.badge-red {{ background: #3C1618; color: #FFA3A6; border: 1px solid #78272C; }}

/* CALENDARIO */
.calbox {{
    border: 1px solid #2D3349;
    background: {SUPERFICIE};
    border-radius: 10px;
    padding: 8px;
    min-height: 165px;
    margin: 3px 0;
    overflow: hidden;
}}
.calbox.holiday {{ border-color: {NARANJA}; background: #221B1A; }}
.calbox.today {{ border: 2px solid {AZUL_PRIMARIO}; }}
.calbox.other {{ opacity: 0.4; }}
.caldate {{ font-size: 14px; font-weight: 700; margin-bottom: 4px; }}
.calbadge-feriado {{
    font-size: 10.5px;
    font-weight: 700;
    color: {NARANJA};
    margin-bottom: 5px;
    white-space: nowrap;
    overflow: hidden;
    text-overflow: ellipsis;
}}
.calshift {{
    font-size: 10.5px;
    background: #1B2135;
    color: #CCD8F5;
    padding: 2px 5px;
    border-radius: 4px;
    margin-bottom: 3px;
    border-left: 3px solid {AZUL_PRIMARIO};
    white-space: nowrap;
    overflow: hidden;
    text-overflow: ellipsis;
}}
.calshift-warning {{
    font-size: 10.5px;
    background: #362214;
    color: #FFB380;
    padding: 2px 5px;
    border-radius: 4px;
    margin-bottom: 3px;
    border-left: 3px solid {NARANJA};
}}

/* MATRIZ DE TURNOS */
.schedule-wrap {{
    margin: 12px 0 24px;
    border: 1px solid #24314C;
    border-radius: 16px;
    overflow: hidden;
    background: #0D1321;
}}
.schedule-header {{
    display:flex;
    justify-content:space-between;
    align-items:center;
    gap:12px;
    padding: 14px 18px;
    background: linear-gradient(90deg, #111A2D 0%, #17284F 100%);
    border-bottom: 1px solid #24314C;
}}
.schedule-header-title {{
    color:#FFFFFF;
    font-size:15px;
    font-weight:800;
}}
.schedule-header-subtitle {{
    color:#AFC4FF;
    font-size:12px;
}}
.schedule-legend {{
    display:flex;
    gap:8px;
    flex-wrap:wrap;
}}
.legend-chip {{
    display:inline-flex;
    align-items:center;
    gap:6px;
    padding:4px 10px;
    border-radius:999px;
    font-size:11px;
    font-weight:700;
    border:1px solid #2D3D5C;
    background:#111827;
    color:#E2E8F0;
}}
.schedule-scroll {{
    overflow-x:auto;
}}
.table-schedule {{
    width:100%;
    min-width: 1080px;
    border-collapse: collapse;
    background:#0D1321;
}}
.table-schedule th, .table-schedule td {{
    border: 1px solid #24314C;
    vertical-align: top;
}}
.table-schedule th {{
    background: #111C34;
    color:#FFFFFF;
    padding:10px 8px;
    text-align:center;
}}
.table-schedule th:first-child {{
    min-width:250px;
    text-align:left;
    position: sticky;
    left: 0;
    z-index: 3;
    background:#111C34;
}}
.table-schedule td:first-child {{
    position: sticky;
    left: 0;
    z-index: 2;
    background:#0F182B;
    min-width:250px;
}}
.schedule-name {{
    padding:10px 12px;
}}
.schedule-name strong {{
    display:block;
    color:#FFFFFF;
    font-size:13.5px;
}}
.schedule-name span {{
    display:block;
    color:#94A3B8;
    font-size:11.5px;
    margin-top:3px;
}}
.schedule-day-title {{
    font-size:12px;
    font-weight:800;
    color:#FFFFFF;
}}
.schedule-day-date {{
    font-size:11px;
    color:#AFC4FF;
    margin-top:2px;
}}
.schedule-day-cell {{
    padding:8px;
    min-width:120px;
    background:#0D1321;
}}
.schedule-day-cell.out-month {{
    background:#0A0F1B;
    opacity:0.55;
}}
.schedule-day-cell.today {{
    box-shadow: inset 0 0 0 2px {AZUL_PRIMARIO};
}}
.schedule-feriado {{
    display:inline-block;
    margin-bottom:6px;
    padding:3px 8px;
    border-radius:999px;
    font-size:10px;
    font-weight:800;
    background:rgba(255,122,26,0.16);
    border:1px solid rgba(255,122,26,0.45);
    color:#FFD1B0;
}}
.schedule-badge {{
    display:inline-block;
    padding:4px 8px;
    border-radius:10px;
    font-size:11px;
    font-weight:800;
    line-height:1.2;
}}
.schedule-detail {{
    display:block;
    margin-top:5px;
    color:#D6E0FF;
    font-size:10.5px;
    line-height:1.35;
}}
.status-real {{ background: rgba(34,197,94,0.18); border:1px solid rgba(34,197,94,0.45); color:#9EF3BA; }}
.status-sim {{ background: rgba(47,107,255,0.18); border:1px solid rgba(47,107,255,0.45); color:#AFC4FF; }}
.status-descanso {{ background: rgba(148,163,184,0.18); border:1px solid rgba(148,163,184,0.42); color:#E2E8F0; }}
.status-vacaciones {{ background: rgba(255,122,26,0.18); border:1px solid rgba(255,122,26,0.45); color:#FFD1B0; }}
.status-fuera {{ background: rgba(71,85,105,0.16); border:1px dashed rgba(148,163,184,0.35); color:#94A3B8; }}
.status-sinregistro {{ background: rgba(239,68,68,0.16); border:1px solid rgba(239,68,68,0.35); color:#FFB3B3; }}

/* MATRIZ APROBADOR VACACIONES */
.vac-wrap {{
    margin: 12px 0 24px;
    border: 1px solid #24314C;
    border-radius: 16px;
    overflow: hidden;
    background: #0D1321;
}}
.vac-table {{
    width:100%;
    min-width:1400px;
    border-collapse: collapse;
    background:#0D1321;
}}
.vac-table th, .vac-table td {{
    border: 1px solid #24314C;
}}
.vac-table th:first-child, .vac-table td:first-child {{
    position: sticky;
    left: 0;
    z-index: 2;
    background:#0F182B;
    min-width:260px;
}}
.vac-table th:first-child {{
    z-index: 3;
    background:#111C34;
}}
.vac-table th {{
    background:#111C34;
    color:#FFFFFF;
    padding:6px 4px;
    text-align:center;
}}
.vac-day-num {{ font-size:12px; font-weight:800; color:#FFFFFF; }}
.vac-day-dow {{ font-size:9px; color:#AFC4FF; margin-top:2px; }}
.vac-cell {{
    min-width:32px;
    height:30px;
    text-align:center;
    font-size:10px;
    font-weight:800;
    color:#FFFFFF;
    background:#0D1321;
}}
.vac-empty {{ background:#0D1321; color:#0D1321; }}
.vac-aprobada {{ background:rgba(255,122,26,0.80); color:#090A10; }}
.vac-pendiente {{ background:rgba(255,122,26,0.18); color:#FFD1B0; box-shadow: inset 0 0 0 1px rgba(255,122,26,0.45); }}
.vac-art {{ background:rgba(47,107,255,0.80); color:#FFFFFF; }}
.vac-art-pendiente {{ background:rgba(47,107,255,0.18); color:#AFC4FF; box-shadow: inset 0 0 0 1px rgba(47,107,255,0.45); }}
.vac-otro {{ background:rgba(255,255,255,0.12); color:#FFFFFF; }}
.vac-otro-pendiente {{ background:rgba(148,163,184,0.18); color:#E2E8F0; }}
.vac-rechazada {{ background:rgba(71,85,105,0.22); color:#94A3B8; text-decoration:line-through; }}

/* MOCKUP WHATSAPP */
.wa-container {{
    max-width: 540px;
    margin: 0 auto;
    background: #111B21;
    border: 1px solid #2A3942;
    border-radius: 20px;
    overflow: hidden;
    box-shadow: 0 10px 30px rgba(0,0,0,0.6);
}}
.wa-header {{
    background: #202C33;
    padding: 12px 18px;
    display: flex;
    align-items: center;
    border-bottom: 1px solid #2A3942;
}}
.wa-header-avatar {{
    width: 38px;
    height: 38px;
    border-radius: 50%;
    background: #00A884;
    display: flex;
    align-items: center;
    justify-content: center;
    font-weight: 800;
    color: #FFF;
    margin-right: 12px;
}}
.wa-body {{
    background: #0B141A;
    background-image: radial-gradient(#182229 1px, transparent 1px);
    background-size: 16px 16px;
    padding: 18px;
    min-height: 280px;
}}
.wa-bubble {{
    background: #202C33;
    color: #E9EDEF;
    border-radius: 12px;
    border-top-left-radius: 2px;
    padding: 12px 14px;
    max-width: 90%;
    margin-bottom: 14px;
    box-shadow: 0 2px 6px rgba(0,0,0,0.3);
    font-size: 13.5px;
    line-height: 1.45;
}}
.wa-bubble-title {{
    font-weight: 700;
    color: #53BDEB;
    margin-bottom: 6px;
}}
.wa-bubble-time {{
    text-align: right;
    font-size: 10.5px;
    color: #8696A0;
    margin-top: 4px;
}}
</style>
    """,
    unsafe_allow_html=True,
)

# ============================================================
# 3. UTILIDADES GENERALES Y REGLAS LABORALES ARGENTINAS (LCT)
# ============================================================

try:
    HOY_AR = datetime.now(ZoneInfo("America/Argentina/Buenos_Aires")).date()
except Exception:
    HOY_AR = date(2026, 11, 12)


def norm(txt):
    if txt is None:
        return ""
    s = str(txt).strip().lower()
    return unicodedata.normalize("NFKD", s).encode("ASCII", "ignore").decode("utf-8")


def entera(val):
    try:
        return f"{int(round(val)):,}".replace(",", ".")
    except Exception:
        return "0"


def ars(val):
    if val is None or pd.isna(val):
        return "—"
    try:
        return f"$ {float(val):,.2f}".replace(",", "@").replace(".", ",").replace("@", ".")
    except Exception:
        return str(val)


def millones(val):
    if val is None or pd.isna(val):
        return "—"
    try:
        return f"$ {float(val) / 1_000_000:.2f} M".replace(".", ",")
    except Exception:
        return str(val)


def semilla(valor):
    return int(hashlib.md5(str(valor).encode("utf-8")).hexdigest()[:8], 16)


def cabecera(titulo, descripcion, etiqueta_modulo="Módulo de Recursos Humanos"):
    st.markdown(
        f"""
        <div class="hero">
            <div class="eyebrow">{etiqueta_modulo} · El Luchador</div>
            <h1>{html.escape(titulo)}</h1>
            <p>{html.escape(descripcion)}</p>
        </div>
        """,
        unsafe_allow_html=True,
    )


def generar_cuil(dni, sexo=0):
    prefijo = 20 if sexo == 0 else 27
    dni_str = str(dni).zfill(8)
    base_str = f"{prefijo}{dni_str}"
    factores = [5, 4, 3, 2, 7, 6, 5, 4, 3, 2]
    suma = sum(int(base_str[i]) * factores[i] for i in range(10))
    resto = suma % 11
    if resto == 0:
        verif = 0
    elif resto == 1:
        prefijo = 23
        verif = 9 if sexo == 0 else 4
    else:
        verif = 11 - resto
    return f"{prefijo}-{dni_str}-{verif}"


def calcular_edad(fecha_nac, fecha_ref=None):
    if pd.isna(fecha_nac):
        return None
    ref = fecha_ref or HOY_AR
    return ref.year - fecha_nac.year - ((ref.month, ref.day) < (fecha_nac.month, fecha_nac.day))


def calcular_antiguedad_exacta(fecha_ingreso, fecha_ref=None):
    if pd.isna(fecha_ingreso):
        return None, None, None, "Sin registrar"
    ref = fecha_ref or HOY_AR
    anios = ref.year - fecha_ingreso.year
    meses = ref.month - fecha_ingreso.month
    dias = ref.day - fecha_ingreso.day
    if dias < 0:
        meses -= 1
        dias += 30
    if meses < 0:
        anios -= 1
        meses += 12
    anios = max(0, anios)
    meses = max(0, meses)
    dias = max(0, dias)
    return anios, meses, dias, f"{anios}a {meses}m {dias}d"


def escala_vacaciones_lct(antiguedad_anios):
    if pd.isna(antiguedad_anios):
        return None
    if antiguedad_anios <= 5:
        return 14
    elif antiguedad_anios <= 10:
        return 21
    elif antiguedad_anios <= 20:
        return 28
    else:
        return 35


def calcular_plus_vacacional(sueldo_mensual, dias_vacaciones):
    if pd.isna(sueldo_mensual):
        return None, None, None
    if sueldo_mensual <= 0 or dias_vacaciones <= 0:
        return 0.0, 0.0, 0.0
    jornal_vacacional = sueldo_mensual / 25.0
    jornal_habitual = sueldo_mensual / 30.0
    monto_total_vacaciones = jornal_vacacional * dias_vacaciones
    plus_neto = (jornal_vacacional - jornal_habitual) * dias_vacaciones
    return jornal_vacacional, monto_total_vacaciones, plus_neto


def costo_laboral_total(sueldo_bruto):
    if pd.isna(sueldo_bruto):
        sueldo_bruto = float("nan")
    cargas_patronales = sueldo_bruto * 0.240
    art = sueldo_bruto * 0.030
    provision_sac = sueldo_bruto * (1.0 / 12.0)
    cargas_sac = provision_sac * 0.240
    costo_mensual = sueldo_bruto + cargas_patronales + art + provision_sac + cargas_sac
    return {
        "bruto": sueldo_bruto,
        "cargas_patronales": cargas_patronales,
        "art": art,
        "provision_sac": provision_sac,
        "cargas_sac": cargas_sac,
        "costo_total_mensual": costo_mensual,
        "costo_total_anual": costo_mensual * 12.0,
    }


def calcular_pago_art(sueldo_mensual, dias_licencia):
    if pd.isna(sueldo_mensual) or sueldo_mensual <= 0 or dias_licencia <= 0:
        return 0.0, 0, 0, 0.0, 0.0
    jornal = sueldo_mensual / 30.0
    dias_empleador = min(int(dias_licencia), 10)
    dias_art = max(0, int(dias_licencia) - 10)
    costo_empleador = jornal * dias_empleador
    costo_art = jornal * dias_art
    return jornal, dias_empleador, dias_art, costo_empleador, costo_art


def aplicar_estilo_plotly(fig):
    fig.update_layout(
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        font=dict(color=BLANCO),
        title_font=dict(color=BLANCO, size=18),
        legend=dict(font=dict(color=BLANCO), title=dict(font=dict(color=BLANCO))),
        margin=dict(t=48, b=24, l=24, r=24),
        hoverlabel=dict(bgcolor="#111827", font_color=BLANCO, bordercolor=AZUL_PRIMARIO),
    )
    fig.update_xaxes(title_font=dict(color=BLANCO), tickfont=dict(color="#D7E3FF"), gridcolor="rgba(255,255,255,0.10)", zerolinecolor="rgba(255,255,255,0.10)")
    fig.update_yaxes(title_font=dict(color=BLANCO), tickfont=dict(color="#D7E3FF"), gridcolor="rgba(255,255,255,0.10)", zerolinecolor="rgba(255,255,255,0.10)")
    return fig


def area_de_rol(rol):
    r = norm(rol)
    if any(k in r for k in ["cortador", "carnic"]):
        return "Carnicería"
    if any(k in r for k in ["verdulero", "verdul"]):
        return "Verdulería"
    if any(k in r for k in ["cajer", "factur"]):
        return "Cajas / Facturación"
    if any(k in r for k in ["fiambr"]):
        return "Fiambrería"
    if any(k in r for k in ["produccion", "avicol"]):
        return "Producción / Avícola"
    if any(k in r for k in ["almacen"]):
        return "Almacén"
    if any(k in r for k in ["gerente", "admin", "direccion"]):
        return "Administración y Dirección"
    return "Operaciones Generales"


def registrar_auditoria(usuario, rol, accion, entidad, detalle):
    if "rrhh_audit_log" not in st.session_state:
        st.session_state.rrhh_audit_log = []
    st.session_state.rrhh_audit_log.append({
        "timestamp": datetime.now(ZoneInfo("America/Argentina/Buenos_Aires")).strftime("%Y-%m-%d %H:%M:%S"),
        "usuario": usuario,
        "rol": rol,
        "accion": accion,
        "entidad": entidad,
        "detalle": detalle,
    })


# ============================================================
# 4. CARGA Y ENRIQUECIMIENTO DE DATOS
# ============================================================

@st.cache_data(show_spinner="Procesando registros de personal...")
def cargar_personal(contenido):
    # Leer las dos filas de encabezado del nuevo Excel sin convertirlas en empleados.
    raw = pd.read_excel(io.BytesIO(contenido), sheet_name="Hoja1", header=None, engine="openpyxl", dtype=object)
    encabezados = ["Razon social", "Sucursal", "Local", "Empleados", "Atiende al publico?", "Rol",
                   "Nombre Excel", "Apellido", "Fecha de nacimiento", "DNI", "Fecha de ingreso",
                   "Fecha de Reingreso 1", "Fecha de Reingreso 2", "Calle", "Nro", "Depto", "Localidad",
                   "Celular", "Telefono fijo", "Horas diarias", "Horas semanales", "Jornada horaria",
                   "Días laborables", "Salario en recibo de sueldo", "Semanal", "Salario mensual Calculo",
                   "Aguinaldo junio", "Aguinaldo diciembre", "Semanas 2026", "Total (sin aguinaldo)",
                   "semanal real", "Total real (sin aguinaldo)", "Diferencia"]
    if raw.shape[1] < len(encabezados) or norm(raw.iloc[0, 0]) != "razon social" or norm(raw.iloc[1, 25]) != "salario mensual calculo":
        raise ValueError("El formato esperado es Hoja1, con dos encabezados y Salario mensual Calculo en Z2.")
    df = raw.iloc[2:, :len(encabezados)].copy()
    df.columns = encabezados
    df["Fila Excel"] = df.index + 1
    df = df.loc[df["Empleados"].notna() & df["Sucursal"].notna() & df["Local"].notna()].copy()
    # Mantener el alcance del piloto: las partidas compartidas y Natalia no integran esta nómina.
    df = df.loc[~df["Empleados"].map(norm).isin(["compras verduleria", "compras", "verduleria", "encargado", "natalia"])].copy()
    def texto(v):
        if pd.isna(v):
            return ""
        if isinstance(v, (int, float)) and float(v).is_integer():
            return str(int(v))
        return str(v).strip()
    for col in ["Nombre Excel", "Apellido", "Rol", "Sucursal", "Empleados", "DNI", "Calle", "Nro", "Depto", "Localidad", "Celular", "Telefono fijo", "Jornada horaria", "Días laborables", "Atiende al publico?"]:
        df[col] = df[col].map(texto)
    def nombre_visible(r):
        partes = " ".join(x for x in (r["Nombre Excel"], r["Apellido"]) if x)
        return partes or f'{r["Rol"] or r["Empleados"]} (sin nombre registrado)'
    df["Nombre"] = df.apply(nombre_visible, axis=1)
    df["Identidad completa"] = df["Nombre Excel"].ne("") & df["Apellido"].ne("")
    df["Razón social"] = df["Razon social"].map(lambda v: "Misión" if "mision" in norm(v) else ("Il Comendatori" if "comendatori" in norm(v) else texto(v)))
    df["Local"] = df["Local"].map(lambda v: next((k for k in SUCURSALES_INFO if norm(k) == norm(v)), texto(v)))
    num = df["Sucursal"].str.extract(r"(\d+)")[0].fillna("X")
    numero_empleado = df["Empleados"].str.extract(r"(?i)^empleado\s+(\d+)$")[0]
    df["ID"] = [f"S{n}-E{int(e):02d}" if pd.notna(e) else f"S{n}-{norm(rol).replace(' ', '-')}" for n, e, rol in zip(num, numero_empleado, df["Rol"])]
    if df["ID"].duplicated().any():
        raise ValueError("Se detectaron IDs de personal duplicados; revise los códigos de empleado por sucursal.")
    df["Unidad de Negocio"] = df["Rol"].map(area_de_rol)
    df["Área"] = df["Unidad de Negocio"]
    for col in ["Salario en recibo de sueldo", "Semanal", "Salario mensual Calculo", "semanal real", "Aguinaldo junio", "Aguinaldo diciembre", "Semanas 2026", "Total (sin aguinaldo)", "Total real (sin aguinaldo)", "Diferencia", "Horas diarias", "Horas semanales"]:
        df[col] = pd.to_numeric(df[col], errors="coerce")
    df["Mensual ARS"] = df["Salario mensual Calculo"].where(df["Salario mensual Calculo"].ge(0))
    df["Mensual real ARS"] = (df["semanal real"] * 52 / 12).where(df["semanal real"].ge(0))
    df["Semanal equivalente ARS"] = df["Mensual ARS"] * 12 / 52
    df["Anual base ARS"] = df["Mensual ARS"] * 12
    for origen, destino in [("Fecha de nacimiento", "Fecha Nacimiento"), ("Fecha de ingreso", "Fecha Ingreso"), ("Fecha de Reingreso 1", "Fecha Reingreso 1"), ("Fecha de Reingreso 2", "Fecha Reingreso 2")]:
        df[destino] = df[origen].map(lambda v: pd.to_datetime(v, errors="coerce", dayfirst=True)).map(lambda v: v.date() if pd.notna(v) else None)
    df["Domicilio"] = df.apply(lambda r: ", ".join(x for x in (" ".join(x for x in [r["Calle"], r["Nro"]] if x), r["Depto"], r["Localidad"]) if x), axis=1)
    # El Excel no contiene estos antecedentes: no generar datos personales ficticios.
    for col in [
        "CUIL", "ART Aseguradora", "ART Poliza", "ART Estado Caso", "ART Fecha Denuncia",
        "ART Días a cargo empleador", "ART Días a cargo aseguradora", "Obra Social",
        "Prepaga", "Nro Afiliado Salud", "Examen preocupacional", "Categoria Convenio",
        "Familiares", "Contacto Emergencia", "Estado Laboral"
    ]:
        df[col] = "Sin registrar"
    df["Ultima Confirmacion Domicilio"] = None
    df["ART Vencimiento"] = None
    return df.reset_index(drop=True)


def mostrar_dato(v):
    return "Sin registrar" if pd.isna(v) or str(v).strip() == "" else str(v)


def mostrar_fecha(v):
    return v.strftime("%d/%m/%Y") if pd.notna(v) else "Sin registrar"


def promedio_conocidos(valores):
    validos = [v for v in valores if pd.notna(v)]
    return sum(validos) / len(validos) if validos else None

# ============================================================
# 5. CALENDARIO OFICIAL ARGENTINO
# ============================================================

@st.cache_data
def feriados_argentina(anio):
    try:
        feriados_ar = holidays.Argentina(years=[anio])
        return {fecha: nombre for fecha, nombre in feriados_ar.items()}
    except Exception:
        return {}


# ============================================================
# 6. GESTIÓN DE TURNOS, COBERTURA Y REEMPLAZOS
# ============================================================

def turno_base(uid, dia):
    # Las jornadas explícitas tienen prioridad. El resto se identifica como simulación.
    coincidencias = personal.loc[personal["ID"] == uid]
    if coincidencias.empty:
        return "Sin registro"
    r = coincidencias.iloc[0]
    dias = norm(r["Días laborables"])
    if dias:
        if norm(DIAS_SEMANA[dia.weekday()]) not in dias:
            return "Descanso"
        return r["Jornada horaria"] or "Laborable (horario sin registrar)"
    descanso = semilla("descanso-" + uid) % 7
    if dia.weekday() == descanso:
        return "Descanso simulado"
    semana = dia.isocalendar().week
    tipo = (semilla("turno-" + uid) + semana) % 3
    return ["Sim. Apertura (07:00-15:00)", "Sim. Cierre (14:00-22:00)", "Sim. Cortado (08:30-12:30 / 16:30-20:30)"][tipo]

def licencia_activa(uid, dia):
    for s in st.session_state.rrhh_vacaciones:
        if s["ID"] == uid and s["Estado"] == "Aprobada" and s["Inicio"] <= dia <= s["Fin"]:
            return s.get("Tipo", "Vacaciones LCT")
    return None


def esta_en_vacaciones(uid, dia):
    tipo = licencia_activa(uid, dia)
    return tipo == "Vacaciones LCT"


def turno_efectivo(uid, dia):
    clave_reemplazo = (dia, uid)
    if clave_reemplazo in st.session_state.rrhh_reemplazos:
        return st.session_state.rrhh_reemplazos[clave_reemplazo]["turno"]
    tipo = licencia_activa(uid, dia)
    if tipo:
        return "Vacaciones" if tipo == "Vacaciones LCT" else tipo
    return turno_base(uid, dia)


def evaluar_cobertura_sucursal(plantilla, local, fecha):
    sub = plantilla.loc[plantilla["Local"] == local]
    resultados = []
    for area in UNIDADES_NEGOCIO:
        empleados_area = sub.loc[sub["Unidad de Negocio"] == area]
        total_asignados = len(empleados_area)
        if total_asignados == 0:
            continue
        minimo_requerido = DOTACION_MINIMA_DEFECTO.get(area, 1)
        presentes = 0
        ausentes = 0
        for _, emp in empleados_area.iterrows():
            t = turno_efectivo(emp["ID"], fecha)
            if not t.startswith("Descanso") and t not in ["Vacaciones", "Enfermedad / Certificado", "ART / Accidente laboral", "Otra licencia", "Sin registro"]:
                presentes += 1
            else:
                ausentes += 1
        cubierto = presentes >= minimo_requerido
        resultados.append({
            "Área": area,
            "Total Asignados": total_asignados,
            "Presentes": presentes,
            "Mínimo Requerido": minimo_requerido,
            "Déficit": max(0, minimo_requerido - presentes),
            "Estado": "Óptimo" if presentes > minimo_requerido else ("Al Límite" if cubierto else "Déficit Crítico"),
        })
    return pd.DataFrame(resultados)


def etiqueta_turno_resumida(turno):
    t = norm(turno)
    if not t or t == "nan":
        return "Sin registro"
    if "art" in t:
        return "ART"
    if "enfermedad" in t or "certificado" in t:
        return "Enfermedad"
    if "otra licencia" in t:
        return "Licencia"
    if "vacacion" in t:
        return "Vacaciones"
    if "descanso" in t:
        return "Descanso"
    if "apertura" in t:
        return "Apertura"
    if "cierre" in t:
        return "Cierre"
    if "cortado" in t:
        return "Cortado"
    if "laborable" in t:
        return "Laborable"
    if "sin registro" in t:
        return "Sin registro"
    return str(turno)


def clase_turno(turno, es_mes=True):
    t = norm(turno)
    if not es_mes:
        return "status-fuera"
    if "art" in t:
        return "status-sim"
    if "enfermedad" in t or "otra licencia" in t:
        return "status-vacaciones"
    if "vacacion" in t:
        return "status-vacaciones"
    if "descanso" in t:
        return "status-descanso"
    if "sin registro" in t or not t:
        return "status-sinregistro"
    if t.startswith("sim."):
        return "status-sim"
    return "status-real"


def detalle_turno(turno, es_mes=True):
    if not es_mes:
        return "Fuera del mes"
    return html.escape(str(turno), quote=True)


def render_matriz_turnos(sub_personal, fechas, feriados_anio, titulo, subtitulo=""):
    if sub_personal.empty:
        return f'<div class="panel-card"><b>No hay trabajadores para mostrar.</b><br><span style="color:{GRIS_TEXTO};">Ajuste los filtros para visualizar la matriz.</span></div>'

    df = sub_personal.sort_values(["Unidad de Negocio", "Rol", "Nombre", "ID"]).copy()
    legend = (
        '<div class="schedule-legend">'
        '<span class="legend-chip">🟢 Jornada registrada</span>'
        '<span class="legend-chip">🔵 Turno simulado</span>'
        '<span class="legend-chip">🟠 Vacaciones / feriado</span>'
        '<span class="legend-chip">⚪ Descanso</span>'
        '</div>'
    )

    html_out = [
        '<div class="schedule-wrap">',
        '<div class="schedule-header">',
        f'<div><div class="schedule-header-title">{html.escape(titulo)}</div><div class="schedule-header-subtitle">{html.escape(subtitulo)}</div></div>',
        legend,
        '</div>',
        '<div class="schedule-scroll">',
        '<table class="table-schedule">',
        '<thead><tr><th>Trabajador</th>'
    ]

    for d in fechas:
        html_out.append(
            f'<th><div class="schedule-day-title">{html.escape(DIAS_SEMANA[d.weekday()])}</div><div class="schedule-day-date">{d.strftime("%d/%m")}</div></th>'
        )
    html_out.append('</tr></thead><tbody>')

    mes_referencia = fechas[3].month if len(fechas) >= 4 else fechas[0].month
    for _, emp in df.iterrows():
        html_out.append('<tr>')
        nombre = html.escape(str(emp["Nombre"]))
        rol = html.escape(str(emp["Rol"]))
        local = html.escape(str(emp["Local"]))
        html_out.append(
            f'<td><div class="schedule-name"><strong>{nombre}</strong><span>{rol} · {local}</span></div></td>'
        )

        for d in fechas:
            es_mes = d.month == mes_referencia
            turno = turno_efectivo(emp["ID"], d)
            clase_td = 'schedule-day-cell'
            if not es_mes:
                clase_td += ' out-month'
            if d == HOY_AR:
                clase_td += ' today'
            badge_class = clase_turno(turno, es_mes=es_mes)
            badge_text = html.escape(etiqueta_turno_resumida(turno))
            detalle = detalle_turno(turno, es_mes=es_mes)
            feriado_html = ''
            if d in feriados_anio:
                feriado_html = f'<div class="schedule-feriado">🇦🇷 {html.escape(str(feriados_anio[d]))}</div>'
            html_out.append(
                f'<td class="{clase_td}">{feriado_html}<span class="schedule-badge {badge_class}" title="{detalle}">{badge_text}</span><span class="schedule-detail">{detalle}</span></td>'
            )

        html_out.append('</tr>')

    html_out.append('</tbody></table></div></div>')
    return ''.join(html_out)


def licencia_en_fecha(uid, dia):
    activos = []
    for s in st.session_state.rrhh_vacaciones:
        if s["ID"] == uid and s["Inicio"] <= dia <= s["Fin"]:
            activos.append(s)
    if not activos:
        return None
    prioridad = {"Aprobada": 0, "Pendiente": 1, "Rechazada": 2}
    activos.sort(key=lambda s: prioridad.get(s.get("Estado", "Pendiente"), 9))
    return activos[0]


def clase_licencia_celda(sol):
    if sol is None:
        return "vac-empty"
    tipo = sol.get("Tipo", "Vacaciones LCT")
    estado = sol.get("Estado", "Pendiente")
    if estado == "Rechazada":
        return "vac-rechazada"
    if tipo == "ART / Accidente laboral":
        return "vac-art" if estado == "Aprobada" else "vac-art-pendiente"
    if tipo == "Vacaciones LCT":
        return "vac-aprobada" if estado == "Aprobada" else "vac-pendiente"
    return "vac-otro" if estado == "Aprobada" else "vac-otro-pendiente"


def texto_licencia_celda(sol):
    if sol is None:
        return ""
    tipo = sol.get("Tipo", "Vacaciones LCT")
    if tipo == "Vacaciones LCT":
        return "V"
    if tipo == "ART / Accidente laboral":
        return "ART"
    if tipo == "Enfermedad / Certificado":
        return "E"
    return "L"


def render_matriz_aprobador_vacaciones(sub_personal, anio, mes):
    if sub_personal.empty:
        return '<div class="panel-card">No hay personal para la vista de aprobador.</div>'
    dias_mes = calendar.monthrange(anio, mes)[1]
    fechas = [date(anio, mes, d) for d in range(1, dias_mes + 1)]
    leyenda = ('<div class="schedule-legend">'
               '<span class="legend-chip">🟧 Vacación aprobada</span>'
               '<span class="legend-chip">🟠 Vacación pendiente</span>'
               '<span class="legend-chip">🟦 ART / accidente</span>'
               '<span class="legend-chip">⬜ Otras licencias</span>'
               '</div>')
    out = ['<div class="vac-wrap">',
           f'<div class="schedule-header"><div><div class="schedule-header-title">Vista aprobador · {MESES[mes-1]} {anio}</div><div class="schedule-header-subtitle">Matriz diaria para detectar solapes, backups y cobertura por persona.</div></div>{leyenda}</div>',
           '<div class="schedule-scroll"><table class="vac-table"><thead><tr><th>Trabajador</th>']
    for d in fechas:
        out.append(f'<th><div class="vac-day-num">{d.day}</div><div class="vac-day-dow">{DIAS_SEMANA[d.weekday()][:3]}</div></th>')
    out.append('</tr></thead><tbody>')
    df = sub_personal.sort_values(["Local", "Unidad de Negocio", "Nombre"]).copy()
    for _, emp in df.iterrows():
        out.append('<tr>')
        out.append(f'<td><div class="schedule-name"><strong>{html.escape(str(emp["Nombre"]))}</strong><span>{html.escape(str(emp["Rol"]))} · {html.escape(str(emp["Local"]))}</span></div></td>')
        for d in fechas:
            sol = licencia_en_fecha(emp["ID"], d)
            cls = clase_licencia_celda(sol)
            txt = texto_licencia_celda(sol)
            title = ''
            if sol is not None:
                title = html.escape(f'{sol.get("Tipo", "Vacaciones LCT")} · {sol.get("Estado", "Pendiente")} · {sol["Inicio"].strftime("%d/%m/%Y")} a {sol["Fin"].strftime("%d/%m/%Y")}', quote=True)
            out.append(f'<td class="vac-cell {cls}" title="{title}">{txt}</td>')
        out.append('</tr>')
    out.append('</tbody></table></div></div>')
    return ''.join(out)


# ============================================================
# 7. ESTADO DE SESIÓN (SESSION STATE)
# ============================================================

def inicializar_session_state(personal, version_fuente):
    if st.session_state.get("rrhh_version_fuente") != version_fuente:
        for clave in list(st.session_state.keys()):
            if clave.startswith("rrhh_"):
                del st.session_state[clave]
        st.session_state.rrhh_version_fuente = version_fuente
    defaults = {"rrhh_vacaciones": [], "rrhh_anticipos": [], "rrhh_reemplazos": {},
                "rrhh_audit_log": [], "rrhh_alertas_vistas": set(), "rrhh_docs_modificados": {},
                "rrhh_domicilios_ratificados": set(), "rrhh_whatsapp_mensajes": []}
    for clave, valor in defaults.items():
        if clave not in st.session_state:
            st.session_state[clave] = valor

def antiguedad_a_cierre(ingreso, anio):
    if pd.isna(ingreso):
        return None
    corte = date(anio, 12, 31)
    return max(0, corte.year - ingreso.year - ((corte.month, corte.day) < (ingreso.month, ingreso.day)))


def dias_aprobados_empleado(uid, anio):
    return sum(
        s["Dias"]
        for s in st.session_state.rrhh_vacaciones
        if s["ID"] == uid
        and s["Estado"] == "Aprobada"
        and s["Inicio"].year == anio
        and s.get("Tipo", "Vacaciones LCT") == "Vacaciones LCT"
    )


# ============================================================
# 8. CARGA DE ARCHIVO Y PREPARACIÓN DE DATOS
# ============================================================

with st.sidebar:
    st.markdown(
        """
        <div style="background:linear-gradient(135deg, #162447, #0E1322); padding:18px; border:1px solid #355BA8; border-radius:14px; margin-bottom:15px;">
            <div style="font-size:11px; font-weight:800; letter-spacing:1.5px; color:#8EB2FF;">GRUPO EMPRESARIAL</div>
            <h2 style="margin:4px 0 2px; font-size:20px; color:#FFF;">EL LUCHADOR</h2>
            <div style="font-size:12px; color:#A7BEE8;">Módulo de Recursos Humanos · Etapa 1</div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    if RUTA_EXCEL.exists():
        fuente = RUTA_EXCEL
    else:
        fuente = st.file_uploader("Cargar archivo sueldos.xlsx", type=["xlsx"])
        if fuente is None:
            st.warning("Coloque sueldos.xlsx en la carpeta data/ o cárguelo para iniciar.")
            st.stop()

try:
    contenido = fuente.read_bytes() if isinstance(fuente, Path) else fuente.getvalue()
    personal = cargar_personal(contenido)
except Exception as ex:
    st.error(f"Error procesando sueldos.xlsx: {ex}")
    st.stop()

if personal.empty:
    st.error("No se encontraron registros válidos de trabajadores.")
    st.stop()

inicializar_session_state(personal, hashlib.sha256(contenido).hexdigest())

# ============================================================
# 9. BARRA LATERAL, PERFIL RBAC Y FILTROS GLOBALES
# ============================================================

with st.sidebar:
    st.caption("Datos: sueldos.xlsx · Las operaciones de la demo se guardan solo en la sesión. El selector de perfil no es autenticación real.")
    st.markdown("### Perfil de Acceso (RBAC)")
    rol_activo = st.selectbox(
        "Rol Activo",
        ROLES_RBAC,
        index=0,
        help="Simula permisos y vistas según el rol de seguridad asignado (Sección 2.2).",
    )

    badge_color = "badge-blue" if "Administrador" in rol_activo else ("badge-green" if "Gerente" in rol_activo else "badge-orange")
    st.markdown(f'<span class="badge {badge_color}">PERFIL: {rol_activo.upper()}</span>', unsafe_allow_html=True)

    st.divider()

    st.markdown("### Navegación")
    pagina = st.radio(
        "Sección del Módulo",
        [
            "📊 Resumen Ejecutivo (KPIs)",
            "🗂️ Legajos Digitales",
            "📅 Calendario & Turnos",
            "🏖️ Vacaciones & Régimen LCT",
            "💰 Remuneraciones & Anticipos",
            "📈 Productividad Operativa (POS)",
            "🔔 Centro de Alertas & WhatsApp",
            "🛡️ Registro de Auditoría",
        ],
        label_visibility="collapsed",
    )

    st.divider()

    st.markdown("### Filtros Globales")
    rs_filtro = st.multiselect(
        "Razón Social",
        RAZONES_SOCIALES,
        default=RAZONES_SOCIALES,
    )

    sucursales_disponibles = sorted(personal["Local"].unique().tolist())
    sucursales_filtro = st.multiselect(
        "Sucursales",
        sucursales_disponibles,
        default=sucursales_disponibles,
    )

    anio_filtro = st.selectbox("Año de Gestión", [2026, 2027, 2028], index=0)

base = personal.loc[
    (personal["Razón social"].isin(rs_filtro)) &
    (personal["Local"].isin(sucursales_filtro))
].copy()

if base.empty:
    st.warning("No hay empleados que coincidan con los filtros seleccionados.")
    st.stop()


# ============================================================
# 10. MÓDULO 1: RESUMEN EJECUTIVO (DASHBOARD DE KPIS)
# ============================================================

if pagina == "📊 Resumen Ejecutivo (KPIs)":
    cabecera(
        "Dashboard Ejecutivo de Recursos Humanos",
        "Indicadores de personal, presupuesto salarial, calidad de datos y cobertura simulada.",
        "Etapa 1 · Panel Gerencial",
    )

    total_empleados = len(base)
    salario_mensual_total = base["Mensual ARS"].sum(min_count=1)
    costos_totales = [costo_laboral_total(s) for s in base["Mensual ARS"]]
    costo_mensual_empresa = sum(c["costo_total_mensual"] for c in costos_totales)
    
    edades = [calcular_edad(nac) for nac in base["Fecha Nacimiento"]]
    edad_promedio = promedio_conocidos(edades)
    
    antiguedades = [antiguedad_a_cierre(ing, anio_filtro) for ing in base["Fecha Ingreso"]]
    antig_promedio = promedio_conocidos(antiguedades)

    anticipos_activos = sum(a["Saldo Pendiente"] for a in st.session_state.rrhh_anticipos if a["Estado"] == "Aprobado" and a["ID"] in set(base["ID"]))

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Registros de Personal", f"{total_empleados} pers.")
    c2.metric("Masa Salarial Base", millones(salario_mensual_total))
    c3.metric("Costo total estimado", millones(costo_mensual_empresa), help="Incluye Salarios + Cargas Patronales (~24%) + ART (~3%) + Provisiones SAC.")
    c4.metric("Anticipos Activos", ars(anticipos_activos))

    c5, c6, c7, c8 = st.columns(4)
    c5.metric("Edad Promedio", f"{edad_promedio:.1f} años" if edad_promedio is not None else "Sin datos", help=f"Fechas registradas: {sum(pd.notna(v) for v in edades)}/{len(base)}")
    c6.metric("Antigüedad Media", f"{antig_promedio:.1f} años" if antig_promedio is not None else "Sin datos")
    
    coberturas_locales = [evaluar_cobertura_sucursal(personal, s, HOY_AR) for s in base["Local"].unique()]
    if coberturas_locales:
        df_cob_global = pd.concat(coberturas_locales)
        pct_cob = (1.0 - (df_cob_global["Déficit"].sum() / max(1, df_cob_global["Mínimo Requerido"].sum()))) * 100
    else:
        pct_cob = 100.0
    c7.metric("Cobertura simulada", f"{pct_cob:.1f} %", delta="Operativo" if pct_cob >= 90 else "-Alerta")

    solic_pendientes = sum(1 for s in st.session_state.rrhh_vacaciones if s["Estado"] == "Pendiente" and s["ID"] in set(base["ID"]))
    c8.metric("Vacaciones x Resolver", f"{solic_pendientes} sol.", delta=f"{solic_pendientes} pendientes", delta_color="inverse")

    st.divider()

    st.markdown("### Calidad de los datos del personal")
    st.caption("Se excluyen partidas compartidas de Compras Verdulería y Encargado, y Natalia, según el alcance original de la demo. Producción se conserva por puesto aunque falte identificación.")
    campos = {"Nombre y apellido": base["Identidad completa"], "DNI": base["DNI"].ne(""),
              "Nacimiento": base["Fecha Nacimiento"].notna(), "Ingreso": base["Fecha Ingreso"].notna(),
              "Domicilio": base["Domicilio"].ne(""), "Celular": base["Celular"].ne(""),
              "Jornada": base["Jornada horaria"].ne(""), "Días laborables": base["Días laborables"].ne(""),
              "Salario mensual": base["Mensual ARS"].notna(), "Salario en recibo": base["Salario en recibo de sueldo"].notna()}
    calidad = pd.DataFrame([{"Campo": k, "Registrados": int(v.sum()), "Pendientes": int((~v).sum())} for k,v in campos.items()])
    st.dataframe(calidad, hide_index=True, use_container_width=True)
    st.caption("Cobertura ART, antigüedad y cumplimiento documental no verificables con el Excel. Costos patronales y turnos no registrados son estimaciones de la demo.")

    st.markdown("### Desglose y Distribución Operativa")
    g1, g2 = st.columns(2)

    with g1:
        fig_suc = px.bar(
            base.groupby(["Local", "Razón social"]).size().reset_index(name="Dotación"),
            x="Local",
            y="Dotación",
            color="Razón social",
            barmode="group",
            title="Registros de Personal por Sucursal y Razón Social",
            color_discrete_map={"Misión": AZUL_PRIMARIO, "Il Comendatori": NARANJA},
        )
        aplicar_estilo_plotly(fig_suc)
        st.plotly_chart(fig_suc, use_container_width=True)

    with g2:
        cargas_tot = sum(c["cargas_patronales"] for c in costos_totales)
        art_tot = sum(c["art"] for c in costos_totales)
        sac_tot = sum(c["provision_sac"] + c["cargas_sac"] for c in costos_totales)
        fig_costos = px.pie(
            names=["Salario Mensual Base", "Cargas Patronales (24%)", "ART (3%)", "Provisión SAC + Cargas"],
            values=[salario_mensual_total, cargas_tot, art_tot, sac_tot],
            title="Composición estimada del costo laboral",
            color_discrete_sequence=[AZUL_PRIMARIO, "#4B7DFF", NARANJA, "#FCD34D"],
            hole=0.45,
        )
        aplicar_estilo_plotly(fig_costos)
        st.plotly_chart(fig_costos, use_container_width=True)


# ============================================================
# 11. MÓDULO 2: LEGAJOS DIGITALES Y FICHA DEL TRABAJADOR
# ============================================================

elif pagina == "🗂️ Legajos Digitales":
    cabecera("Legajos Digitales y Ficha Integral del Trabajador", "Datos del Excel, información pendiente y documentos de la sesión.", "Sección 4.1 a 4.3 · Legajo Digital")
    busqueda = st.text_input("Buscar trabajador por nombre, puesto, ID o DNI")
    listado = base.loc[base.apply(lambda r: norm(busqueda) in norm(" ".join(str(r[c]) for c in ["Nombre", "Rol", "ID", "DNI"])), axis=1)]
    st.dataframe(listado[["ID", "Nombre", "Local", "Rol", "Identidad completa"]], hide_index=True, use_container_width=True)
    if listado.empty:
        st.info("No se encontraron coincidencias.")
        st.stop()
    seleccion_id = st.selectbox("Seleccionar trabajador", listado["ID"].tolist(), format_func=lambda x: f"{x} — {listado.loc[listado['ID']==x, 'Nombre'].iloc[0]}")
    emp = listado.loc[listado["ID"] == seleccion_id].iloc[0]
    st.subheader(emp["Nombre"])
    a,b,c,d = st.columns(4)
    a.metric("Sucursal", emp["Local"])
    b.metric("Salario mensual calculado (ARS)", ars(emp["Mensual ARS"]))
    c.metric("Semanal registrado (ARS)", ars(emp["Semanal"]))
    d.metric("Semanal real (ARS)", ars(emp["semanal real"]))
    if not emp["Identidad completa"]:
        st.warning("Nombre o apellido pendiente de completar.")

    t_pers,t_lab,t_sal,t_salud,t_art,t_docs,t_hist = st.tabs(["Datos personales", "Datos laborales", "Remuneración", "Cobertura médica", "ART / riesgo laboral", "Documentos", "Historial"])
    def ficha(pares):
        st.dataframe(pd.DataFrame(pares, columns=["Campo", "Información"]).assign(Información=lambda x: x["Información"].map(mostrar_dato)), hide_index=True, use_container_width=True)

    with t_pers:
        edad = calcular_edad(emp["Fecha Nacimiento"])
        ficha([("Nombre",emp["Nombre Excel"]),("Apellido",emp["Apellido"]),("DNI",emp["DNI"]),("CUIL",emp["CUIL"]),("Nacimiento",mostrar_fecha(emp["Fecha Nacimiento"])),("Edad",edad),("Calle",emp["Calle"]),("Número",emp["Nro"]),("Departamento",emp["Depto"]),("Localidad",emp["Localidad"]),("Celular",emp["Celular"]),("Teléfono fijo",emp["Telefono fijo"]),("Familiares",emp["Familiares"]),("Contacto emergencia",emp["Contacto Emergencia"])])

    with t_lab:
        _,_,_,ant = calcular_antiguedad_exacta(emp["Fecha Ingreso"])
        ficha([("Código original",emp["Empleados"]),("Razón social",emp["Razón social"]),("Sucursal",emp["Sucursal"]),("Local",emp["Local"]),("Rol",emp["Rol"]),("Unidad de negocio",emp["Unidad de Negocio"]),("Atiende al público",emp["Atiende al publico?"]),("Fecha ingreso",mostrar_fecha(emp["Fecha Ingreso"])),("Reingreso 1",mostrar_fecha(emp["Fecha Reingreso 1"])),("Reingreso 2",mostrar_fecha(emp["Fecha Reingreso 2"])),("Antigüedad",ant),("Horas diarias",emp["Horas diarias"]),("Horas semanales",emp["Horas semanales"]),("Jornada horaria",emp["Jornada horaria"]),("Días laborables",emp["Días laborables"]),("Categoría convenio",emp["Categoria Convenio"]),("Estado laboral",emp["Estado Laboral"])])

    with t_sal:
        cols = ["Salario en recibo de sueldo", "Semanal", "Salario mensual Calculo", "Aguinaldo junio", "Aguinaldo diciembre", "Total (sin aguinaldo)", "semanal real", "Mensual real ARS", "Total real (sin aguinaldo)", "Diferencia"]
        ficha([(col,ars(emp[col]) if pd.notna(emp[col]) else "Sin registrar") for col in cols]+[("Semanas 2026",emp["Semanas 2026"]),("Semanal equivalente calculado",ars(emp["Semanal equivalente ARS"])),("Anual base calculada",ars(emp["Anual base ARS"]))])
        st.caption("Cecilia indicó que el salario liquidado puede diferir del realmente percibido. Por eso se exponen por separado salario base, semanal registrado y semanal real.")

    with t_salud:
        st.info("La cobertura médica no es lo mismo que la ART. Aquí se concentra obra social / prepaga y datos de afiliación del trabajador.")
        ficha([("Obra social", emp["Obra Social"]), ("Prepaga", emp["Prepaga"]), ("N° de afiliado", emp["Nro Afiliado Salud"]), ("Examen preocupacional", emp["Examen preocupacional"])])

    with t_art:
        st.info("La ART se trata como bloque separado, tal como pidió Cecilia. Sirve para accidentes laborales o in itinere, con seguimiento diferenciado respecto de la cobertura médica.")
        ficha([
            ("ART aseguradora", emp["ART Aseguradora"]),
            ("Póliza ART", emp["ART Poliza"]),
            ("Vencimiento póliza", mostrar_fecha(emp["ART Vencimiento"])),
            ("Estado de caso ART", emp["ART Estado Caso"]),
            ("Fecha de denuncia", emp["ART Fecha Denuncia"]),
            ("Días a cargo empleador", emp["ART Días a cargo empleador"]),
            ("Días a cargo aseguradora", emp["ART Días a cargo aseguradora"]),
        ])
        st.caption("Regla operativa sugerida en la demo: en licencias ART prolongadas, los primeros 10 días quedan a cargo del empleador y el excedente debe tramitarse con la aseguradora.")

    with t_docs:
        documentos = ["Contrato de trabajo", "DNI frente y dorso", "Constancia CUIL", "Alta laboral", "Certificado médico", "Constancia de EPP", "Recibo de haberes"]
        repo = st.session_state.rrhh_docs_modificados

        st.markdown("#### Semáforo de control legal y vencimientos")
        controles = [
            {"Control": "Identidad completa", "Estado": "Cumple" if emp["Identidad completa"] else "Pendiente", "Comentario": "Nombre y apellido completos" if emp["Identidad completa"] else "Falta completar nombre o apellido"},
            {"Control": "Fecha de ingreso", "Estado": "Cumple" if pd.notna(emp["Fecha Ingreso"]) else "Pendiente", "Comentario": mostrar_fecha(emp["Fecha Ingreso"])},
            {"Control": "CUIL", "Estado": "Pendiente" if mostrar_dato(emp["CUIL"]) == "Sin registrar" else "Cumple", "Comentario": mostrar_dato(emp["CUIL"])},
            {"Control": "Cobertura médica", "Estado": "Pendiente" if mostrar_dato(emp["Obra Social"]) == "Sin registrar" and mostrar_dato(emp["Prepaga"]) == "Sin registrar" else "Cumple", "Comentario": f"OS: {mostrar_dato(emp['Obra Social'])} · Prepaga: {mostrar_dato(emp['Prepaga'])}"},
            {"Control": "ART", "Estado": "Pendiente" if mostrar_dato(emp["ART Aseguradora"]) == "Sin registrar" else "Cumple", "Comentario": f"Aseguradora: {mostrar_dato(emp['ART Aseguradora'])}"},
            {"Control": "Documentación esencial adjunta", "Estado": "Cumple" if all((seleccion_id, d) in repo for d in ["Contrato de trabajo", "DNI frente y dorso", "Constancia CUIL"]) else "Parcial", "Comentario": f"Adjuntos en sesión: {sum((seleccion_id, d) in repo for d in documentos)}/{len(documentos)}"},
        ]
        st.dataframe(pd.DataFrame(controles), hide_index=True, use_container_width=True)

        st.dataframe(pd.DataFrame([{"Documento":d,"Estado":"Adjunto en sesión" if (seleccion_id,d) in repo else "No aportado","Archivo":repo.get((seleccion_id,d),{}).get("nombre", "—")} for d in documentos]), hide_index=True, use_container_width=True)
        tipo_doc = st.selectbox("Tipo de documento", documentos)
        archivo = st.file_uploader("Adjuntar documento (solo sesión)",type=["pdf","png","jpg"],key=f"doc_{seleccion_id}")
        if st.button("Guardar documento en sesión", disabled=archivo is None):
            repo[(seleccion_id,tipo_doc)] = {"nombre":archivo.name,"datos":archivo.getvalue()}
            registrar_auditoria(rol_activo,rol_activo,"CARGA_DOCUMENTO",seleccion_id,tipo_doc)
            st.rerun()

    with t_hist:
        ficha([(k,mostrar_fecha(emp[k])) for k in ["Fecha Ingreso", "Fecha Reingreso 1", "Fecha Reingreso 2"]])
        st.caption("Aquí puede consolidarse más adelante el historial cronológico de ausencias, accidentes, enfermedad, sanciones o incidencias del legajo.")


# ============================================================
# 12. MÓDULO 3: CALENDARIO & PLANIFICADOR DE TURNOS
# ============================================================

elif pagina == "📅 Calendario & Turnos":
    cabecera(
        "Calendario Laboral Oficial & Planificador de Turnos",
        "Vista operativa por trabajador con matriz semanal o mensual, feriados nacionales, control de cobertura y sustitución de personal.",
        "Sección 4.4 · Turnos y Cobertura",
    )

    st.info("Jornadas y días del Excel tienen prioridad. Si no existen datos explícitos, la demo marca el turno como simulado para no dejar huecos visuales.")
    c1, c2, c3, c4 = st.columns([1, 1, 1, 1.1])
    mes_sel = c1.selectbox("Mes", range(1, 13), index=HOY_AR.month - 1, format_func=lambda m: MESES[m - 1])
    sucursal_sel = c2.selectbox("Sucursal", sorted(base["Local"].unique().tolist()))
    area_sel = c3.selectbox("Unidad de Negocio", ["Todas"] + UNIDADES_NEGOCIO)
    vista_modo = c4.selectbox("Vista del Planificador", ["Semanal", "Mensual", "Control de Cobertura Diaria"])

    feriados_anio = feriados_argentina(anio_filtro)
    sub_personal = base.loc[base["Local"] == sucursal_sel].copy()
    if area_sel != "Todas":
        sub_personal = sub_personal.loc[sub_personal["Unidad de Negocio"] == area_sel]

    if sub_personal.empty:
        st.warning("No hay personal cargado para esa combinación de sucursal y unidad de negocio.")
        st.stop()

    if vista_modo == "Control de Cobertura Diaria":
        st.markdown("### Verificación de Dotación Mínima Operativa por Turno")
        dia_inspeccion = st.date_input("Fecha a evaluar", value=HOY_AR)
        df_cob = evaluar_cobertura_sucursal(personal, sucursal_sel, dia_inspeccion)
        st.dataframe(df_cob, use_container_width=True, hide_index=True)

        hay_deficit = not df_cob.empty and df_cob["Déficit"].sum() > 0
        if hay_deficit:
            st.error(f"⚠️ **Alerta Operativa:** Se detectaron puestos descubiertos en la sucursal {sucursal_sel} para el día {dia_inspeccion.strftime('%d/%m/%Y')}.")
        else:
            st.success(f"✓ Cobertura suficiente en la simulación de las áreas registradas de {sucursal_sel}.")

        with st.expander("🔄 Consultar candidatos para contingencias"):
            emp_reemplazado = st.selectbox("Colaborador a cubrir", sub_personal["ID"].tolist(), format_func=lambda x: f"{x} - {sub_personal.loc[sub_personal['ID']==x, 'Nombre'].iloc[0]}")
            ausente = sub_personal.loc[sub_personal["ID"] == emp_reemplazado].iloc[0]
            candidatos = personal.loc[(personal["ID"] != emp_reemplazado) & (personal["Unidad de Negocio"] == ausente["Unidad de Negocio"])].copy()
            candidatos["Turno previsto"] = candidatos["ID"].map(lambda uid: turno_efectivo(uid, dia_inspeccion))
            st.dataframe(candidatos[["Nombre", "Local", "Rol", "Turno previsto"]], hide_index=True, use_container_width=True)
            st.caption("Validar disponibilidad, descansos y cobertura de origen antes de autorizar un reemplazo real.")

    else:
        cal = calendar.Calendar(firstweekday=0)
        semanas_mes = cal.monthdatescalendar(anio_filtro, mes_sel)

        st.markdown(f"### Matriz de turnos del personal — {sucursal_sel}")
        k1, k2, k3, k4 = st.columns(4)
        k1.metric("Trabajadores en la vista", entera(len(sub_personal)))
        k2.metric("Con jornada registrada", entera(sub_personal['Jornada horaria'].astype(str).str.strip().replace({'nan': ''}).ne('').sum()))
        k3.metric("Con días laborables declarados", entera(sub_personal['Días laborables'].astype(str).str.strip().replace({'nan': ''}).ne('').sum()))
        k4.metric("Mes seleccionado", MESES[mes_sel - 1])

        st.caption("La vista semanal muestra una sola semana con columnas de lunes a domingo. La vista mensual apila todas las semanas del mes con el mismo formato para facilitar la lectura por aprobador.")

        if vista_modo == "Semanal":
            semana_sel = st.selectbox(
                "Semana del mes",
                range(len(semanas_mes)),
                format_func=lambda i: f"Semana {i + 1} · {semanas_mes[i][0]:%d/%m} – {semanas_mes[i][-1]:%d/%m}",
            )
            semana = semanas_mes[semana_sel]
            html_matriz = render_matriz_turnos(
                sub_personal,
                semana,
                feriados_anio,
                titulo=f"Semana {semana_sel + 1}: {semana[0]:%d/%m} – {semana[-1]:%d/%m}",
                subtitulo=f"Sucursal {sucursal_sel} · {MESES[mes_sel - 1]} {anio_filtro} · {len(sub_personal)} trabajadores",
            )
            st.markdown(html_matriz, unsafe_allow_html=True)
        else:
            for idx, semana in enumerate(semanas_mes, start=1):
                html_matriz = render_matriz_turnos(
                    sub_personal,
                    semana,
                    feriados_anio,
                    titulo=f"Semana {idx}: {semana[0]:%d/%m} – {semana[-1]:%d/%m}",
                    subtitulo=f"Vista mensual consolidada · {MESES[mes_sel - 1]} {anio_filtro}",
                )
                st.markdown(html_matriz, unsafe_allow_html=True)


# ============================================================
# 13. MÓDULO 4: VACACIONES & RÉGIMEN LCT (LEY 20.744)
# ============================================================

elif pagina == "🏖️ Vacaciones & Régimen LCT":
    cabecera(
        "Gestión de Vacaciones, Licencias y Plus Vacacional (Ley 20.744)",
        "Escala de antigüedad legal, seguimiento de vacaciones y otras licencias, verificación de dotación mínima, cálculo de plus vacacional y flujo de aprobaciones.",
        "Sección 4.5 · Vacaciones LCT y Licencias",
    )

    with st.expander("📖 Marco Legal Argentino: Régimen de Vacaciones LCT N° 20.744"):
        st.markdown(
            """
            - **Escala de Licencia Anual Ordinaria (Art. 150 LCT):**
              - Antigüedad <= 5 años: 14 días corridos
              - 5 < Antigüedad <= 10 años: 21 días corridos
              - 10 < Antigüedad <= 20 años: 28 días corridos
              - Antigüedad > 20 años: 35 días corridos
            - **Cálculo del Plus Vacacional (Art. 155 LCT):**
              El día de vacación se remunera dividiendo el sueldo mensual por 25 (y no por 30).
            - **Observación funcional pedida por Cecilia:** la vista aprobador debe permitir detectar solapes, backups y cobertura por persona. Además, la **ART** debe tratarse como licencia separada respecto de la cobertura médica.
            """
        )

    tab_solic, tab_matriz, tab_aprob = st.tabs(["📝 Registro & Simulación de Solicitud", "📊 Matriz de Saldos por Empleado", "🗓️ Vista Aprobador"])

    with tab_solic:
        st.markdown("### Nueva Solicitud de Licencia")
        col_s1, col_s2 = st.columns(2)

        with col_s1:
            tipo_licencia = st.selectbox("Tipo de solicitud", ["Vacaciones LCT", "Enfermedad / Certificado", "ART / Accidente laboral", "Otra licencia"])
            emp_vac_id = st.selectbox(
                "Colaborador",
                base["ID"].tolist(),
                format_func=lambda x: f"{x} - {base.loc[base['ID']==x, 'Nombre'].iloc[0]} ({base.loc[base['ID']==x, 'Local'].iloc[0]})",
            )
            emp_sel = base.loc[base["ID"] == emp_vac_id].iloc[0]
            antig_anios = antiguedad_a_cierre(emp_sel["Fecha Ingreso"], anio_filtro)
            cupo_legal = escala_vacaciones_lct(antig_anios)
            dias_usados = dias_aprobados_empleado(emp_vac_id, anio_filtro)
            saldo_disp = max(0, cupo_legal - dias_usados) if cupo_legal is not None else None

            if tipo_licencia == "Vacaciones LCT":
                if saldo_disp is None:
                    st.warning("Fecha de ingreso sin registrar: cupo legal y saldo no verificables. Puede simular una solicitud, pero no aprobarla como licencia validada.")
                else:
                    st.info(f"Antigüedad: {antig_anios} años · Cupo estimado: {cupo_legal} días · Saldo: {saldo_disp} días")
            elif tipo_licencia == "ART / Accidente laboral":
                st.info("Seguimiento diferencial solicitado por Cecilia: la licencia ART se registra aparte y puede distinguir días a cargo del empleador y días recuperables vía aseguradora.")
            else:
                st.info("Esta licencia se registra en la bandeja y en la vista aprobador, sin consumir cupo legal de vacaciones.")

            inicio_vac = st.date_input("Fecha Inicio", value=HOY_AR + timedelta(days=10))
            max_dias = max(1, saldo_disp) if tipo_licencia == "Vacaciones LCT" and saldo_disp is not None else 35
            valor_default = min(14, max_dias) if tipo_licencia == "Vacaciones LCT" else 3
            dias_pedidos = st.number_input("Días Corridos Solicitados", min_value=1, max_value=max_dias, value=valor_default)
            fin_vac = inicio_vac + timedelta(days=int(dias_pedidos) - 1)

        with col_s2:
            if tipo_licencia == "Vacaciones LCT":
                st.markdown("#### Impacto Salarial Estimado (LCT Art. 155)")
                jornal_vac, monto_total_vac, plus_neto = calcular_plus_vacacional(emp_sel["Mensual ARS"], int(dias_pedidos))
                fecha_pago_prev = inicio_vac - timedelta(days=3)

                st.write(f"**Período Solicitado:** {inicio_vac.strftime('%d/%m/%Y')} al {fin_vac.strftime('%d/%m/%Y')} ({dias_pedidos} días)")
                st.write(f"**Jornal Vacacional (Base 25):** {ars(jornal_vac)} / día")
                st.write(f"**Monto Total Vacaciones Estimado:** {ars(monto_total_vac)}")
                st.metric("Diferencia vacacional estimada", ars(plus_neto), help="Diferencia adicional a favor del trabajador por base divisor 25.")
                st.write(f"**Fecha Límite Prevista de Pago:** `{fecha_pago_prev.strftime('%d/%m/%Y')}`")
            elif tipo_licencia == "ART / Accidente laboral":
                st.markdown("#### Estimación de impacto ART")
                jornal_art, dias_emp, dias_art, costo_emp, costo_art = calcular_pago_art(emp_sel["Mensual ARS"], int(dias_pedidos))
                st.write(f"**Período Solicitado:** {inicio_vac.strftime('%d/%m/%Y')} al {fin_vac.strftime('%d/%m/%Y')} ({dias_pedidos} días)")
                st.write(f"**Jornal de referencia (base 30):** {ars(jornal_art)} / día")
                st.metric("Días a cargo del empleador", f"{dias_emp} días")
                st.metric("Días a tramitar con ART", f"{dias_art} días")
                st.write(f"**Costo estimado a cargo del empleador:** {ars(costo_emp)}")
                st.write(f"**Monto referencial recuperable vía ART:** {ars(costo_art)}")
                fecha_pago_prev = inicio_vac
            else:
                st.markdown("#### Seguimiento operativo")
                st.write(f"**Período Solicitado:** {inicio_vac.strftime('%d/%m/%Y')} al {fin_vac.strftime('%d/%m/%Y')} ({dias_pedidos} días)")
                st.metric("Impacto salarial directo", "Revisión manual")
                st.caption("Para enfermedad, certificados u otras licencias, la plataforma registra el evento y permite evaluarlo en la cobertura y la vista aprobador.")
                jornal_vac = monto_total_vac = plus_neto = None
                fecha_pago_prev = inicio_vac

            disabled_btn = (tipo_licencia == "Vacaciones LCT" and (saldo_disp == 0 or pd.isna(emp_sel["Mensual ARS"])))
            if st.button("Registrar solicitud de prueba", type="primary", disabled=disabled_btn):
                nueva_solicitud = {
                    "N": len(st.session_state.rrhh_vacaciones) + 1,
                    "Tipo": tipo_licencia,
                    "ID": emp_sel["ID"],
                    "Nombre": emp_sel["Nombre"],
                    "Local": emp_sel["Local"],
                    "Área": emp_sel["Unidad de Negocio"],
                    "Inicio": inicio_vac,
                    "Fin": fin_vac,
                    "Dias": int(dias_pedidos),
                    "Estado": "Pendiente",
                    "Plus Vacacional ARS": plus_neto,
                    "Total Vacaciones ARS": monto_total_vac,
                    "Fecha Pago Prevista": fecha_pago_prev,
                    "Aprobado Por": "—",
                }
                if tipo_licencia == "ART / Accidente laboral":
                    _, dias_emp, dias_art, costo_emp, costo_art = calcular_pago_art(emp_sel["Mensual ARS"], int(dias_pedidos))
                    nueva_solicitud["Días cargo empleador"] = dias_emp
                    nueva_solicitud["Días cargo ART"] = dias_art
                    nueva_solicitud["Costo estimado empleador"] = costo_emp
                    nueva_solicitud["Monto recuperable ART"] = costo_art
                st.session_state.rrhh_vacaciones.append(nueva_solicitud)
                registrar_auditoria(rol_activo, rol_activo, "SOLICITUD_LICENCIA", emp_sel["ID"], f"{tipo_licencia}: {dias_pedidos} días del {inicio_vac} al {fin_vac}.")
                st.success("Solicitud registrada con éxito y remitida a la bandeja de autorizaciones.")
                st.rerun()

        st.markdown("### Simulación de cobertura durante la licencia")
        st.caption("Disponibilidad orientativa según jornadas registradas y turnos simulados; no representa asistencia confirmada ni autorización de traslado.")
        pares = personal.loc[(personal["ID"] != emp_vac_id) & (personal["Unidad de Negocio"] == emp_sel["Unidad de Negocio"])]
        filas = []
        for offset in range(int(dias_pedidos)):
            dia = inicio_vac + timedelta(days=offset)
            turno_actual = turno_efectivo(emp_vac_id, dia)
            candidatos = []
            for _, r in pares.iterrows():
                t = turno_efectivo(r["ID"], dia)
                if r["Local"] == emp_sel["Local"] and t.startswith("Descanso"):
                    candidatos.append(r["Nombre"] + " (descanso: requiere autorización)")
            cobertura = evaluar_cobertura_sucursal(personal, emp_sel["Local"], dia)
            area_c = cobertura.loc[cobertura["Área"] == emp_sel["Unidad de Negocio"]]
            disponibles = int(area_c["Presentes"].iloc[0]) if not area_c.empty else 0
            minimo = int(area_c["Mínimo Requerido"].iloc[0]) if not area_c.empty else 1
            presente = not turno_actual.startswith("Descanso") and turno_actual not in ["Vacaciones", "Enfermedad / Certificado", "ART / Accidente laboral", "Otra licencia", "Sin registro"]
            despues = max(0, disponibles - int(presente))
            filas.append({"Fecha":dia, "Estado previo":turno_actual,"Presentes después":despues,"Mínimo simulado":minimo,"Déficit":max(0,minimo-despues),"Posibles reemplazos": "; ".join(candidatos) or "Sin candidato en descanso del área"})
        st.dataframe(pd.DataFrame(filas), hide_index=True, use_container_width=True)
        st.divider()
        st.markdown("### Bandeja de Solicitudes y Licencias")
        vac_filtradas = [v for v in st.session_state.rrhh_vacaciones if v["ID"] in set(base["ID"])]
        if vac_filtradas:
            st.dataframe(pd.DataFrame(vac_filtradas), use_container_width=True, hide_index=True)

            solic_pendientes = [s for s in vac_filtradas if s["Estado"] == "Pendiente"]
            if solic_pendientes and "Auditor" not in rol_activo:
                st.markdown("#### Resolver Solicitudes Pendientes")
                idx_res = st.selectbox("Seleccione Solicitud a Resolver", range(len(solic_pendientes)), format_func=lambda i: f"#{solic_pendientes[i]['N']} - {solic_pendientes[i]['Nombre']} · {solic_pendientes[i].get('Tipo', 'Vacaciones LCT')} ({solic_pendientes[i]['Dias']} días desde {solic_pendientes[i]['Inicio'].strftime('%d/%m')})")
                sol_elegida = solic_pendientes[idx_res]

                b_ap, b_rec = st.columns(2)
                if sol_elegida.get("Tipo", "Vacaciones LCT") == "Vacaciones LCT":
                    ingreso_sol = personal.loc[personal["ID"] == sol_elegida["ID"], "Fecha Ingreso"].iloc[0]
                    cupo_sol = escala_vacaciones_lct(antiguedad_a_cierre(ingreso_sol, sol_elegida["Inicio"].year))
                    puede_aprobar = cupo_sol is not None and sol_elegida["Dias"] <= cupo_sol - dias_aprobados_empleado(sol_elegida["ID"], sol_elegida["Inicio"].year)
                    if not puede_aprobar:
                        st.warning("Aprobación bloqueada: falta validar ingreso y saldo suficiente.")
                else:
                    puede_aprobar = True
                    st.info("Licencias no vacacionales pueden aprobarse sin consumir cupo LCT; mantener respaldo documental según corresponda.")
                if b_ap.button("✅ Aprobar Solicitud", type="primary", disabled=not puede_aprobar):
                    for s in st.session_state.rrhh_vacaciones:
                        if s["N"] == sol_elegida["N"]:
                            s["Estado"] = "Aprobada"
                            s["Aprobado Por"] = rol_activo
                    registrar_auditoria(rol_activo, rol_activo, "APROBACION_LICENCIA", sol_elegida["ID"], f"Aprobada licencia #{sol_elegida['N']} ({sol_elegida['Tipo']}).")
                    st.success("Solicitud APROBADA exitosamente.")
                    st.rerun()
                if b_rec.button("❌ Rechazar Solicitud"):
                    for s in st.session_state.rrhh_vacaciones:
                        if s["N"] == sol_elegida["N"]:
                            s["Estado"] = "Rechazada"
                            s["Aprobado Por"] = rol_activo
                    registrar_auditoria(rol_activo, rol_activo, "RECHAZO_LICENCIA", sol_elegida["ID"], f"Rechazada licencia #{sol_elegida['N']}." )
                    st.warning("Solicitud RECHAZADA.")
                    st.rerun()

    with tab_matriz:
        st.markdown(f"### Matriz Consolidada de Vacaciones — Período {anio_filtro}")
        filas_matriz = []
        for _, r in base.iterrows():
            ant = antiguedad_a_cierre(r["Fecha Ingreso"], anio_filtro)
            cupo = escala_vacaciones_lct(ant)
            aprob = dias_aprobados_empleado(r["ID"], anio_filtro)
            filas_matriz.append({
                "ID": r["ID"],
                "Nombre": r["Nombre"],
                "Sucursal": r["Local"],
                "Área": r["Unidad de Negocio"],
                "Antigüedad (Años)": ant,
                "Días por Ley LCT": cupo,
                "Días Gozados / Aprobados": aprob,
                "Saldo Disponible": max(0, cupo - aprob) if cupo is not None else None,
            })
        st.dataframe(pd.DataFrame(filas_matriz), use_container_width=True, hide_index=True)
        st.caption("Esta matriz conserva el enfoque legal de cupo y saldo. Las otras licencias no descuentan vacaciones, pero sí aparecen en la vista aprobador.")

    with tab_aprob:
        st.markdown("### Vista aprobador de vacaciones y licencias")
        cva1, cva2 = st.columns(2)
        local_vista = cva1.selectbox("Sucursal a visualizar", sorted(base["Local"].unique().tolist()), key="vac_local_view")
        mes_vista = cva2.selectbox("Mes a visualizar", range(1, 13), index=HOY_AR.month - 1, format_func=lambda m: MESES[m - 1], key="vac_month_view")
        sub_vista = base.loc[base["Local"] == local_vista].copy()
        st.markdown(render_matriz_aprobador_vacaciones(sub_vista, anio_filtro, mes_vista), unsafe_allow_html=True)
        st.caption("Objetivo funcional: que el aprobador vea rápidamente quién ya está licenciado, quién tiene solicitudes pendientes y dónde podrían generarse solapes o faltantes de backup.")


# ============================================================
# 14. MÓDULO 5: REMUNERACIONES, ANTICIPOS & COSTO LABORAL
# ============================================================

elif pagina == "💰 Remuneraciones & Anticipos":
    cabecera(
        "Administración Salarial, Anticipos y Costos Laborales",
        "Presupuesto sobre salario base o semanal real, gestión auditable de anticipos, estimación de cargas patronales y provisiones de aguinaldo (SAC).",
        "Sección 4.6 · Gestión Salarial",
    )

    t_nom, t_ant, t_costo = st.tabs([
        "💵 Nómina Salarial",
        "💳 Gestión de Anticipos",
        "🏢 Costo Total Empresa & Cargas Patronales",
    ])

    with t_nom:
        base_nomina = st.radio("Base visual de remuneración", ["Salario mensual calculado", "Semanal real anualizado"], horizontal=True)
        col_base = "Mensual ARS" if base_nomina == "Salario mensual calculado" else "Mensual real ARS"
        etiqueta_base = "Masa Mensual Base" if col_base == "Mensual ARS" else "Masa Mensual Realizada"
        total_sal_base = base[col_base].sum(min_count=1)
        total_sac_jun = base["Aguinaldo junio"].sum(min_count=1)
        total_sac_dic = base["Aguinaldo diciembre"].sum(min_count=1)

        c1, c2, c3, c4 = st.columns(4)
        c1.metric(etiqueta_base, millones(total_sal_base))
        c2.metric("Proyección Anual", millones(total_sal_base * 12) if pd.notna(total_sal_base) else "Sin datos")
        c3.metric("SAC Junio (Aguinaldo)", millones(total_sac_jun))
        c4.metric("SAC Diciembre (Aguinaldo)", millones(total_sac_dic))

        st.markdown("#### Detalle de Planilla y Comparación")
        st.caption(f"SAC informado: junio {base['Aguinaldo junio'].notna().sum()}/{len(base)}, diciembre {base['Aguinaldo diciembre'].notna().sum()}/{len(base)}. Cecilia indicó que conviene poder alternar entre salario base y semanal real; por eso la vista permite ambas lecturas. Todos los montos en ARS.")
        cols_mostrar = ["ID", "Nombre", "Local", "Rol", "Salario en recibo de sueldo", "Mensual ARS", "Semanal", "semanal real", "Mensual real ARS", "Semanal equivalente ARS", "Aguinaldo junio", "Aguinaldo diciembre", "Semanas 2026", "Total (sin aguinaldo)", "Total real (sin aguinaldo)", "Diferencia"]
        st.dataframe(base[cols_mostrar], use_container_width=True, hide_index=True)

    with t_ant:
        st.markdown("### Solicitud y Autorización de Anticipos de Sueldo")
        with st.form("form_anticipo"):
            col_a1, col_a2 = st.columns(2)
            with col_a1:
                emp_ant_id = st.selectbox("Colaborador Solicitante", base["ID"].tolist(), format_func=lambda x: f"{x} - {base.loc[base['ID']==x, 'Nombre'].iloc[0]}")
                monto_ant = st.number_input("Monto del Anticipo (ARS)", min_value=5000.0, value=60000.0, step=5000.0)
            with col_a2:
                fecha_ant = st.date_input("Fecha de Solicitud", value=HOY_AR)
                tope_porc = st.slider("Tope Máximo Permitido sobre Salario", 10, 50, 30, 5, help="Política de empresa: habitualmente tope del 30% o 50% de la remuneración mensual.")

            btn_ant = st.form_submit_button("Registrar Solicitud de Anticipo", type="primary")

        if btn_ant:
            sueldo_emp = base.loc[base["ID"] == emp_ant_id, "Mensual ARS"].iloc[0]
            nombre_emp = base.loc[base["ID"] == emp_ant_id, "Nombre"].iloc[0]
            local_emp = base.loc[base["ID"] == emp_ant_id, "Local"].iloc[0]
            limite_dinero = sueldo_emp * (tope_porc / 100.0)

            if pd.isna(sueldo_emp):
                st.error("Salario mensual no registrado; no se puede estimar el límite.")
            elif monto_ant > limite_dinero:
                st.error(f"El anticipo solicitado ({ars(monto_ant)}) supera el límite del {tope_porc}% del sueldo ({ars(limite_dinero)}).")
            else:
                nuevo_ant = {
                    "N": len(st.session_state.rrhh_anticipos) + 1,
                    "ID": emp_ant_id,
                    "Nombre": nombre_emp,
                    "Local": local_emp,
                    "Monto": monto_ant,
                    "Fecha Solicitud": fecha_ant,
                    "Fecha Pago": fecha_ant + timedelta(days=2),
                    "Estado": "Solicitado",
                    "Saldo Pendiente": monto_ant,
                    "Aprobado Por": "—",
                }
                st.session_state.rrhh_anticipos.append(nuevo_ant)
                registrar_auditoria(rol_activo, rol_activo, "SOLICITUD_ANTICIPO", emp_ant_id, f"Anticipo solicitado por {ars(monto_ant)}.")
                st.success("Anticipo registrado en el sistema.")
                st.rerun()

        st.markdown("#### Historial y Control de Anticipos")
        ant_filtrados = [a for a in st.session_state.rrhh_anticipos if a["ID"] in set(base["ID"])]
        if ant_filtrados:
            st.dataframe(pd.DataFrame(ant_filtrados), use_container_width=True, hide_index=True)

            pendientes_ant = [a for a in ant_filtrados if a["Estado"] == "Solicitado"]
            if pendientes_ant and "Auditor" not in rol_activo:
                st.markdown("#### Autorizar Anticipos Pendientes")
                idx_ant = st.selectbox("Seleccione Anticipo a Autorizar", range(len(pendientes_ant)), format_func=lambda i: f"#{pendientes_ant[i]['N']} - {pendientes_ant[i]['Nombre']} ({ars(pendientes_ant[i]['Monto'])})")
                ant_sel = pendientes_ant[idx_ant]
                if st.button("✅ Autorizar y Programar Descuento", type="primary"):
                    for a in st.session_state.rrhh_anticipos:
                        if a["N"] == ant_sel["N"]:
                            a["Estado"] = "Aprobado"
                            a["Aprobado Por"] = rol_activo
                    registrar_auditoria(rol_activo, rol_activo, "APROBACION_ANTICIPO", ant_sel["ID"], f"Anticipo #{ant_sel['N']} aprobado por {rol_activo}.")
                    st.success("Anticipo APROBADO.")
                    st.rerun()

    with t_costo:
        base_costos = st.radio("Base para costo empresa", ["Salario mensual calculado", "Semanal real anualizado"], horizontal=True, key="base_costos_radio")
        col_base_cost = "Mensual ARS" if base_costos == "Salario mensual calculado" else "Mensual real ARS"
        st.markdown("### Estructura de Costo Laboral Total por Empleado y Sucursal")
        st.info("Escenario de presupuesto: la base puede alternarse entre salario calculado y semanal real anualizado. Las tasas de cargas y ART son ilustrativas para la demo.")
        costo_df_list = []
        for _, r in base.iterrows():
            c = costo_laboral_total(r[col_base_cost])
            costo_df_list.append({
                "ID": r["ID"],
                "Nombre": r["Nombre"],
                "Sucursal": r["Local"],
                "Base utilizada": base_costos,
                "Salario base considerado": r[col_base_cost],
                "Cargas Patronales (24%)": c["cargas_patronales"],
                "Alícuota ART (3%)": c["art"],
                "Provisión SAC Mensual": c["provision_sac"],
                "Costo Total Mensual": c["costo_total_mensual"],
                "Costo Total Anual": c["costo_total_anual"],
            })
        df_costos = pd.DataFrame(costo_df_list)
        st.dataframe(df_costos, use_container_width=True, hide_index=True)


# ============================================================
# 15. MÓDULO 6: PRODUCTIVIDAD OPERATIVA (POS)
# ============================================================

elif pagina == "📈 Productividad Operativa (POS)":
    cabecera(
        "Conexión con Productividad Operativa (Punto de Venta / Mostrador)",
        "Seguimiento de actividad por puesto sin emitir juicios causales automáticos: operaciones atendidas y volúmenes facturados por turno.",
        "Sección 4.7 · Productividad",
    )

    st.markdown(
        """
        > [!NOTE]
        > Conforme a los principios de diseño de la especificación técnica (Sección 4.7), el sistema presenta
        > indicadores cuantitativos de actividad operativa y atención de caja sin inferir conclusiones disciplinarias automáticas.
        """
    )

    st.info("sueldos.xlsx no contiene operaciones POS ni ventas por trabajador. No se muestran métricas de productividad inventadas vinculadas a personas reales.")
    st.markdown("### Puestos con atención al público registrada")
    st.dataframe(base[["ID", "Nombre", "Local", "Rol", "Atiende al publico?"]].replace("", "Sin registrar"), hide_index=True, use_container_width=True)

# ============================================================
# 16. MÓDULO 7: CENTRO DE ALERTAS & WHATSAPP BUSINESS
# ============================================================

elif pagina == "🔔 Centro de Alertas & WhatsApp":
    cabecera(
        "Centro de Notificaciones & Simulador WhatsApp Business",
        "Alertas administrativas categorizadas por prioridad e integración simulada con WhatsApp Business Platform (Cloud API).",
        "Sección 5.1 · Notificaciones y Alertas",
    )

    alertas = []

    for _, r in base.iterrows():
        faltantes = []
        for campo, valor in [("nombre/apellido", r["Identidad completa"]), ("DNI", bool(r["DNI"])), ("nacimiento", pd.notna(r["Fecha Nacimiento"])), ("ingreso", pd.notna(r["Fecha Ingreso"])), ("domicilio", bool(r["Domicilio"])), ("jornada", bool(r["Jornada horaria"]))]:
            if not valor:
                faltantes.append(campo)
        if faltantes:
            alertas.append({"Clave": f"DAT-{r['ID']}", "Tipo": "Legajo incompleto", "Prioridad": "Media", "Referencia": f"{r['ID']} - {r['Nombre']}", "Descripción": "Pendiente: " + ", ".join(faltantes), "Accion": "Completar datos de origen"})

    for s in st.session_state.rrhh_vacaciones:
        if s["Estado"] == "Pendiente" and s["ID"] in set(base["ID"]):
            alertas.append({
                "Clave": f"VAC-{s['N']}",
                "Tipo": "Vacaciones",
                "Prioridad": "Media",
                "Referencia": f"{s['ID']} - {s['Nombre']}",
                "Descripción": f"Solicitud de {s['Dias']} días ({s['Inicio'].strftime('%d/%m')} al {s['Fin'].strftime('%d/%m')}) pendiente de resolución.",
                "Accion": "Aprobar Vacación",
            })

    for a in st.session_state.rrhh_anticipos:
        if a["Estado"] == "Solicitado" and a["ID"] in set(base["ID"]):
            alertas.append({
                "Clave": f"ANT-{a['N']}",
                "Tipo": "Anticipo de Sueldo",
                "Prioridad": "Media",
                "Referencia": f"{a['ID']} - {a['Nombre']}",
                "Descripción": f"Solicitud de anticipo por {ars(a['Monto'])} pendiente de aprobación.",
                "Accion": "Autorizar Anticipo",
            })

    for suc in base["Local"].unique():
        df_c = evaluar_cobertura_sucursal(personal, suc, HOY_AR)
        def_crit = df_c.loc[df_c["Déficit"] > 0]
        for _, dc in def_crit.iterrows():
            alertas.append({
                "Clave": f"COB-{suc}-{dc['Área']}",
                "Tipo": "Dotación Mínima (simulación)",
                "Prioridad": "Alta",
                "Referencia": f"{suc} · {dc['Área']}",
                "Descripción": f"Déficit simulado: faltan {dc['Déficit']} personas para alcanzar la dotación mínima.",
                "Accion": "Asignar Reemplazo",
            })

    visibles = [a for a in alertas if a["Clave"] not in st.session_state.rrhh_alertas_vistas]

    col_izq, col_der = st.columns([1.2, 1])

    with col_izq:
        st.markdown("### Alertas Abiertas en el Sistema")
        m1, m2 = st.columns(2)
        m1.metric("Total Pendientes", len(visibles))
        m2.metric("Prioridad Alta", sum(1 for a in visibles if a["Prioridad"] == "Alta"))

        if visibles:
            st.dataframe(pd.DataFrame(visibles).drop(columns=["Clave"]), use_container_width=True, hide_index=True)
            sel_alerta_idx = st.selectbox("Seleccionar Alerta para gestionar", range(len(visibles)), format_func=lambda i: f"[{visibles[i]['Prioridad']}] {visibles[i]['Tipo']} — {visibles[i]['Referencia']}")
            al_activa = visibles[sel_alerta_idx]

            if st.button("Marcar como Revisada / Descartar Alerta"):
                st.session_state.rrhh_alertas_vistas.add(al_activa["Clave"])
                registrar_auditoria(rol_activo, rol_activo, "REVISION_ALERTA", al_activa["Referencia"], f"Alerta {al_activa['Clave']} descartada.")
                st.success("Alerta marcada como revisada.")
                st.rerun()
        else:
            st.success("✓ No hay alertas abiertas en esta vista de la demo.")

    with col_der:
        st.markdown("### 📱 Simulador WhatsApp Business Platform (API)")
        st.markdown(
            """
            <div class="wa-container">
                <div class="wa-header">
                    <div class="wa-header-avatar">EL</div>
                    <div>
                        <div style="font-weight:700; color:#E9EDEF; font-size:14px;">El Luchador · RR. HH.</div>
                        <div style="font-size:11px; color:#00A884;">● Vista simulada · sin conexión API</div>
                    </div>
                </div>
                <div class="wa-body">
            """,
            unsafe_allow_html=True,
        )

        if visibles:
            al_ejemplo = visibles[0]
            txt_wa = f"⚠️ *NOTIFICACIÓN DE RECURSOS HUMANOS*\n\n*Tipo:* {al_ejemplo['Tipo']}\n*Afectado:* {al_ejemplo['Referencia']}\n*Detalle:* {al_ejemplo['Descripción']}\n\nAcción directa requerida por la jefatura."
        else:
            txt_wa = "✓ *SISTEMA DE GESTIÓN EL LUCHADOR*\nTodas las alertas operativas han sido atendidas. Dotaciones cubiertas y nómina en orden."

        txt_wa_html = html.escape(txt_wa).replace("\n", "<br>")

        st.markdown(
            f"""
            <div class="wa-bubble">
                <div class="wa-bubble-title">EL LUCHADOR IN-APP BOT</div>
                <div>{txt_wa_html}</div>
                <div class="wa-bubble-time">{HOY_AR.strftime('%H:%M')} ✓✓</div>
            </div>
            </div>
            </div>
            """,
            unsafe_allow_html=True,
        )

        st.markdown("#### Botones de Acción Rápida (Interactive Quick Actions)")
        if visibles:
            b_wa1, b_wa2 = st.columns(2)
            if b_wa1.button("🟢 Resolver desde WhatsApp", type="primary"):
                st.session_state.rrhh_alertas_vistas.add(visibles[0]["Clave"])
                registrar_auditoria(rol_activo, rol_activo, "RESOLUCION_WHATSAPP", visibles[0]["Referencia"], "Resuelto vía botón rápido de WhatsApp API.")
                st.success("Simulación: alerta marcada como revisada en esta sesión.")
                st.rerun()
            if b_wa2.button("📲 Reenviar Plantilla"):
                st.info("Simulación: no se ha enviado ningún mensaje real.")


# ============================================================
# 17. MÓDULO 8: REGISTRO DE AUDITORÍA (AUDIT LOG)
# ============================================================

elif pagina == "🛡️ Registro de Auditoría":
    cabecera(
        "Pistas de Auditoría y Trazabilidad (Event Sourcing)",
        "Registro de las acciones de la sesión, no persistente ni inmutable.",
        "Sección 2.2 · Seguridad y Auditoría",
    )

    if st.session_state.rrhh_audit_log:
        df_audit = pd.DataFrame(st.session_state.rrhh_audit_log)
        st.dataframe(df_audit.iloc[::-1], use_container_width=True, hide_index=True)
    else:
        st.info("No hay eventos registrados en la sesión actual.")

    if st.button("Descargar Registro de Auditoría (.CSV)"):
        st.download_button(
            label="Descargar CSV",
            data=pd.DataFrame(st.session_state.rrhh_audit_log).to_csv(index=False),
            file_name="auditoria_rrhh_luchador.csv",
            mime="text/csv",
        )


# ============================================================
# 18. PIE DE PÁGINA INSTITUCIONAL
# ============================================================

st.divider()
st.markdown(
    f"""
    <div style="display:flex; justify-content:space-between; align-items:center; font-size:12px; color:{GRIS_TEXTO};">
        <div><b>El Luchador</b> · Plataforma de Gestión e Inteligencia Empresarial (Etapa 1: Recursos Humanos)</div>
        <div>Demo no certifica cumplimiento legal · Multiempresa: Misión & Il Comendatori</div>
    </div>
    """,
    unsafe_allow_html=True,
)
